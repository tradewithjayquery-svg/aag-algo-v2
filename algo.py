"""
AAG LIVE ALGO - Paper Trade (cloud version)

Logic:
  - Entry time par F&O ka top gainer -> uska CALL kharido (paper)
  - Entry time par F&O ka top loser  -> uska PUT kharido (paper)
  - Strike sirf ATM / 1st ITM / 2nd ITM mein se, jo sabse liquid ho (OTM kabhi nahi)
  - Sirf wahi stocks jinka lot size max_lot_size (default 2500) se kam ho
  - Har leg ka alag SL / Target (rupees mein), ya time exit
  - SL/Target entry par hi price mein badal jaate hain (SL price = entry - SL_rupees / qty)
  - Upstox ke live websocket feed se HAR TICK par check hota hai; SL/Target price touch hote hi
    exact us price par exit. Feed na chale toh REST polling (poll_seconds) par apne aap chala jaata hai
  - Koi real order nahi jaata. Sirf Upstox ka live price padhte hain.

Secrets (GitHub Secrets se aate hain, code mein nahi likhne):
  UPSTOX_ACCESS_TOKEN
  TELEGRAM_TOKEN
  TELEGRAM_CHAT_ID   -> "apni_personal_id,group_ya_channel_id"  (comma se alag; pehli ID = aapki)
Settings: config.json (phone se edit kar sakte ho)
"""
import csv
import gzip
import io
import json
import math
import os
import sys
import threading
import time
from datetime import date, datetime, timedelta, timezone

import openpyxl
import pandas as pd
import requests
from openpyxl.styles import Alignment, Font, PatternFill

IST = timezone(timedelta(hours=5, minutes=30))
HERE = os.path.dirname(os.path.abspath(__file__))

UPSTOX_ACCESS_TOKEN = os.environ.get("UPSTOX_ACCESS_TOKEN", "")
TG_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
TG_CHAT = os.environ.get("TELEGRAM_CHAT_ID", "")
TEST_MODE = os.environ.get("TEST_MODE", "0") == "1"

UPSTOX_HEADERS = {
    "Accept": "application/json",
    "Authorization": f"Bearer {UPSTOX_ACCESS_TOKEN}",
}

DEFAULTS = {
    "enabled": True,
    "entry_time": "09:20",
    "exit_time": "12:30",
    "sl_amount": 2500,
    "target_amount": 4000,
    "num_lots": 1,
    "heartbeat_minutes": 10,
    "late_start_grace_minutes": 3,
    "list_size": 10,
    # --- liquidity settings ---
    "itm_levels": 2,          # ATM + kitne ITM strike dekhne hain (2 = ATM, ITM1, ITM2)
    "max_spread_pct": 3.0,    # bid-ask spread is % se zyada ho toh strike illiquid maano
    "min_oi": 0,              # kam se kam Open Interest (0 = check band)
    "max_stock_tries": 3,     # top stock ka option illiquid ho toh list mein agla stock (max itne)
    # --- naye settings ---
    "max_lot_size": 2500,     # sirf wahi stock jinka lot size isse KAM ho (0 = check band)
    "poll_seconds": 5,        # sirf FALLBACK: feed band ho toh REST se kitne second mein price dekhe
    "use_websocket": True,    # True = har tick par check (live feed); False = sirf REST polling
    "feed_stale_seconds": 20, # itne second koi tick na aaye toh feed band maano, REST par jao
}

TICK = 0.05                   # option ka minimum price step

BASE_URL = "https://www.nseindia.com"
GAINERS_URL = f"{BASE_URL}/api/live-analysis-variations?index=gainers"
LOSERS_URL = f"{BASE_URL}/api/live-analysis-variations?index=loosers"

_token_alert_sent = False


def now_ist():
    return datetime.now(IST)


def load_config():
    cfg = dict(DEFAULTS)
    try:
        with open(os.path.join(HERE, "config.json"), encoding="utf-8") as f:
            cfg.update(json.load(f))
    except Exception as e:
        print(f"config.json padh nahi paaye, default use ho rahe hain: {e}")
    return cfg


# =========================
# TELEGRAM
# Pehli chat ID = owner. Baaki (group / channel) ko sirf trade messages jaate hain.
# Technical errors sirf owner ko jaate hain. Test mode mein bhi sirf owner ko.
# =========================

def chat_ids():
    return [c.strip() for c in TG_CHAT.split(",") if c.strip()]


def _send(ids, text):
    if not TG_TOKEN:
        return
    for c in ids:
        try:
            r = requests.post(
                f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage",
                data={"chat_id": c, "text": text},
                timeout=20,
            )
            if r is not None and getattr(r, "status_code", 200) != 200:
                print(f"Telegram ({c}) ne mana kiya: {getattr(r, 'text', '')[:200]}")
        except Exception as e:
            print(f"Telegram error ({c}): {e}")


def tg(text):
    """Sabko bhejo (test mode mein sirf owner ko)."""
    print(text)
    ids = chat_ids()
    _send(ids[:1] if TEST_MODE else ids, text)


def tg_admin(text):
    """Sirf owner ko (technical baatein)."""
    print("[ADMIN]", text)
    _send(chat_ids()[:1], text)


def tg_file(path, caption=""):
    if not TG_TOKEN or not os.path.exists(path):
        return
    ids = chat_ids()
    for c in (ids[:1] if TEST_MODE else ids):
        try:
            with open(path, "rb") as f:
                requests.post(
                    f"https://api.telegram.org/bot{TG_TOKEN}/sendDocument",
                    data={"chat_id": c, "caption": caption},
                    files={"document": f},
                    timeout=60,
                )
        except Exception as e:
            print(f"Telegram file error ({c}): {e}")


def skip(reason, notify_all=False):
    """
    Aaj trade nahi hua. Wajah sirf owner ko jaati hai.
    notify_all=True sirf tab, jab data ki problem se trade chhoota (group ko bhi pata chale).
    """
    tg_admin(f"⚠️ Aaj trade skip: {reason}")
    if notify_all and not TEST_MODE:
        tg("ℹ️ Aaj ka paper trade nahi hua (technical wajah).")


# =========================
# DUPLICATE RUN GUARD
# GitHub ek hi subah mein kai baar (3 schedule + manual) run start kar sakta hai.
# Sabse purana run chalta hai, baaki chup-chaap band ho jaate hain.
# =========================

def another_run_active():
    repo = os.environ.get("GITHUB_REPOSITORY")
    token = os.environ.get("GITHUB_TOKEN")
    my_id = os.environ.get("GITHUB_RUN_ID")
    if not (repo and token and my_id):
        return False
    try:
        r = requests.get(
            f"https://api.github.com/repos/{repo}/actions/runs",
            headers={"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"},
            params={"status": "in_progress", "per_page": 50},
            timeout=20,
        )
        for run in r.json().get("workflow_runs", []):
            title = str(run.get("display_title") or run.get("name") or "")
            if title.upper().startswith("TEST"):
                continue
            if not str(run.get("path", "")).endswith("algo.yml"):
                continue
            if int(run["id"]) < int(my_id):
                print(f"Purana run #{run['id']} pehle se chal raha hai, ye run band.")
                return True
    except Exception as e:
        print(f"Duplicate check fail (ignore): {e}")
    return False


# =========================
# UPSTOX
# =========================

def upstox_get(url, params=None, tries=3):
    global _token_alert_sent
    for i in range(tries):
        try:
            r = requests.get(url, headers=UPSTOX_HEADERS, params=params, timeout=20)
            if r.status_code == 200:
                return r.json()
            print(f"Upstox {r.status_code}: {r.text[:200]}")
            if r.status_code in (401, 403):
                if not _token_alert_sent:
                    _token_alert_sent = True
                    tg_admin("❌ Upstox token mein problem (401/403). Naya token banakar GitHub Secret update karo.")
                return None
        except Exception as e:
            print(f"Upstox request error: {e}")
        time.sleep(2)
    return None


def get_live_ltp(instrument_key):
    if not instrument_key:
        return None
    data = upstox_get(
        "https://api.upstox.com/v2/market-quote/ltp",
        params={"instrument_key": instrument_key},
        tries=2,
    )
    try:
        vals = (data or {}).get("data", {})
        if vals:
            return float(list(vals.values())[0]["last_price"])
    except Exception as e:
        print(f"LTP parse error: {e}")
    return None


def get_nifty_line():
    """Nifty 50 ka spot, % aur points change. Data na mile toh None (line chhod di jaati hai)."""
    data = upstox_get(
        "https://api.upstox.com/v2/market-quote/quotes",
        params={"instrument_key": "NSE_INDEX|Nifty 50"},
        tries=2,
    )
    try:
        v = list(((data or {}).get("data", {}) or {}).values())[0]
        ltp = float(v["last_price"])
        chg = float(v["net_change"])
        prev = ltp - chg
        if prev <= 0:
            return None
        return f"{'🟢' if chg >= 0 else '🔴'} NIFTY 50: {ltp:,.2f} | {chg / prev * 100:+.2f}% ({chg:+.2f} pts)"
    except Exception as e:
        print(f"Nifty data nahi mila: {e}")
        return None


def is_market_holiday(d):
    data = upstox_get(f"https://api.upstox.com/v2/market/holidays/{d.isoformat()}", tries=1)
    try:
        for h in (data or {}).get("data", []):
            if h.get("holiday_type") == "TRADING_HOLIDAY" and "NSE" in (h.get("closed_exchanges") or []):
                return True, h.get("description", "")
    except Exception:
        pass
    return False, ""


def load_fo_instruments():
    """Lot sizes + stock instrument keys + F&O stock list."""
    print("\nF&O Instruments CSV loading...")
    try:
        r = requests.get(
            "https://assets.upstox.com/market-quote/instruments/exchange/complete.csv.gz",
            timeout=60,
        )
        df = pd.read_csv(io.BytesIO(gzip.decompress(r.content)))
        fo = df[(df["exchange"] == "NSE_FO") & (df["instrument_type"] == "OPTSTK")].copy()
        fo["base_symbol"] = fo["tradingsymbol"].str.extract(r"^([A-Z&\-]+)")
        lot_map = fo.groupby("base_symbol")["lot_size"].first().astype(int).to_dict()
        fo_syms = sorted(lot_map.keys())
        eq = df[(df["exchange"] == "NSE_EQ") & (df["instrument_type"] == "EQUITY")]
        eq_map = dict(zip(eq["tradingsymbol"], eq["instrument_key"]))
        print(f"  {len(lot_map)} F&O symbols loaded.")
        return lot_map, eq_map, fo_syms
    except Exception as e:
        print(f"  CSV load failed: {e}")
        return {}, {}, []


def get_eq_key(symbol, eq_map):
    if symbol in eq_map:
        return eq_map[symbol]
    data = upstox_get(
        "https://api.upstox.com/v2/instruments/search",
        params={"query": symbol, "exchange": "NSE_EQ"},
    )
    items = (data or {}).get("data", [])
    for inst in items:
        if (inst.get("trading_symbol", "").upper() == symbol.upper()
                and inst.get("exchange", "") == "NSE"):
            return inst["instrument_key"]
    for inst in items:
        if inst.get("instrument_key", "").startswith("NSE_EQ"):
            return inst["instrument_key"]
    return None


# =========================
# EXPIRY (monthly = last Tuesday)
# =========================

def last_tuesday(yr, mo):
    if mo == 12:
        last_day = date(yr + 1, 1, 1) - timedelta(days=1)
    else:
        last_day = date(yr, mo + 1, 1) - timedelta(days=1)
    while last_day.weekday() != 1:
        last_day -= timedelta(days=1)
    return last_day


def get_monthly_expiry(today):
    exp = last_tuesday(today.year, today.month)
    if exp < today:
        y, m = (today.year + 1, 1) if today.month == 12 else (today.year, today.month + 1)
        exp = last_tuesday(y, m)
    return exp


# =========================
# GAINERS / LOSERS
# =========================

def row_ltp(s):
    try:
        return float(s.get("lastPrice") or s.get("ltp") or s.get("closePrice") or 0)
    except Exception:
        return 0.0


def row_pct(s):
    try:
        return float(s.get("pChange") or s.get("perChange") or 0)
    except Exception:
        return 0.0


def fmt_list(title, rows, n):
    lines = [title]
    for i, s in enumerate(rows[:n], 1):
        lines.append(f"{i}. {s.get('symbol', '-')}  ₹{row_ltp(s):.2f}  ({row_pct(s):+.2f}%)")
    return "\n".join(lines)


def fetch_primary_gainers_losers():
    try:
        from curl_cffi import requests as cffi_requests

        s = cffi_requests.Session(impersonate="chrome124")
        s.get(BASE_URL, timeout=20)
        time.sleep(1)
        headers = {"Accept": "application/json", "Referer": BASE_URL}
        g = s.get(GAINERS_URL, headers=headers, timeout=20).json()
        l = s.get(LOSERS_URL, headers=headers, timeout=20).json()
        gainers = g.get("FOSec", {}).get("data", [])
        losers = l.get("FOSec", {}).get("data", [])
        if gainers and losers:
            return gainers, losers
    except Exception as e:
        print(f"primary fetch fail: {e}")
    return None


def upstox_gainers_losers(eq_map, fo_syms):
    """Pehla source fail ho toh Upstox ke quotes se top gainers/losers nikalo."""
    keys = [eq_map[s] for s in fo_syms if s in eq_map]
    rows = []
    for i in range(0, len(keys), 400):
        data = upstox_get(
            "https://api.upstox.com/v2/market-quote/quotes",
            params={"instrument_key": ",".join(keys[i:i + 400])},
        )
        for k, v in ((data or {}).get("data", {}) or {}).items():
            try:
                ltp = float(v.get("last_price"))
                chg = float(v.get("net_change"))
                prev = ltp - chg
                if prev > 0:
                    rows.append({
                        "symbol": v.get("symbol") or k.split(":")[-1],
                        "lastPrice": ltp,
                        "pChange": round(chg / prev * 100, 2),
                    })
            except Exception:
                continue
    if not rows:
        return None
    return (sorted(rows, key=lambda x: -x["pChange"])[:20],
            sorted(rows, key=lambda x: x["pChange"])[:20])


def get_gainers_losers(eq_map, fo_syms):
    res = fetch_primary_gainers_losers()
    if not res:
        print("Primary data nahi mila, Upstox se nikal rahe hain...")
        res = upstox_gainers_losers(eq_map, fo_syms)
    if not res:
        return None
    gainers, losers = res
    return (sorted(gainers, key=lambda x: -row_pct(x)),
            sorted(losers, key=lambda x: row_pct(x)))


# =========================
# LIQUID STRIKE SELECTION
# =========================

def _num(x):
    try:
        return float(x)
    except Exception:
        return 0.0


def get_liquid_leg(symbol, ltp_fallback, expiry_date, lot_map, eq_map, side, cfg):
    """
    side = "CE" ya "PE".
    ATM, 1st ITM, 2nd ITM mein se sabse liquid strike chunta hai. OTM kabhi nahi.
    Liquid = bid aur ask dono maujood, spread max_spread_pct se kam, OI min_oi se zyada.
    Jo pass ho unme se spread sabse kam wala; barabari par zyada OI wala.
    Koi bhi pass na ho toh None.
    """
    eq_key = get_eq_key(symbol, eq_map)
    if not eq_key:
        return None
    chain, used_expiry = [], None
    # Agar expiry holiday par shift ho gayi ho toh pichhle din try karo
    for back in (0, 1, 2):
        exp = (expiry_date - timedelta(days=back)).strftime("%Y-%m-%d")
        data = upstox_get(
            "https://api.upstox.com/v2/option/chain",
            params={"instrument_key": eq_key, "expiry_date": exp},
        )
        chain = (data or {}).get("data", [])
        if chain:
            used_expiry = exp
            break
    if not chain:
        return None

    spot = float(chain[0].get("underlying_spot_price", ltp_fallback))
    items = sorted(chain, key=lambda it: it["strike_price"])
    strikes = [it["strike_price"] for it in items]
    atm_i = min(range(len(strikes)), key=lambda i: abs(strikes[i] - spot))
    step = -1 if side == "CE" else 1          # CE ka ITM neeche ke strikes, PE ka upar ke
    field = "call_options" if side == "CE" else "put_options"

    cands = []
    for lvl in range(0, int(cfg["itm_levels"]) + 1):
        i = atm_i + step * lvl
        if i < 0 or i >= len(items):
            break
        opt = items[i].get(field, {}) or {}
        md = opt.get("market_data", {}) or {}
        key = opt.get("instrument_key")
        ltp = _num(md.get("ltp"))
        bid, ask = _num(md.get("bid_price")), _num(md.get("ask_price"))
        oi, vol = _num(md.get("oi")), _num(md.get("volume"))
        if not key or ltp <= 0:
            continue
        spread = ((ask - bid) / ((ask + bid) / 2) * 100) if (bid > 0 and ask >= bid) else None
        cands.append({"label": "ATM" if lvl == 0 else f"ITM{lvl}", "strike": strikes[i], "key": key,
                      "ltp": ltp, "bid": bid, "ask": ask, "oi": oi, "volume": vol, "spread": spread})

    for c in cands:
        sp = "no-quote" if c["spread"] is None else f"{c['spread']:.2f}%"
        print(f"  {symbol} {side} {c['label']} {c['strike']:g}: ltp {c['ltp']} bid {c['bid']} ask {c['ask']} "
              f"spread {sp} oi {c['oi']:.0f}")

    ok = [c for c in cands
          if c["spread"] is not None
          and c["spread"] <= float(cfg["max_spread_pct"])
          and c["oi"] >= float(cfg["min_oi"])]
    if not ok:
        print(f"  {symbol} {side}: koi liquid strike nahi mila")
        return None
    best = min(ok, key=lambda c: (round(c["spread"], 1), -c["oi"]))
    best.update({"symbol": symbol, "side": side, "spot": spot, "expiry": used_expiry,
                 "eq_key": eq_key, "lot": lot_map.get(symbol.upper())})
    return best


def pick_leg(rows, side, expiry, lot_map, eq_map, cfg):
    """
    List ke upar se stock lo. Lot size max_lot_size ya usse zyada ho toh stock chhod do
    (ye try mein nahi gina jaata). Option illiquid ho toh bhi agla stock (max_stock_tries tak).
    """
    max_lot = int(cfg["max_lot_size"])
    tries = 0
    for rank, s in enumerate(rows, 1):
        if tries >= int(cfg["max_stock_tries"]):
            break
        sym = str(s["symbol"])
        lot = lot_map.get(sym.upper())
        if max_lot and lot and lot >= max_lot:
            print(f"  #{rank} {sym}: lot size {lot} >= {max_lot}, skip")
            continue
        tries += 1
        leg = get_liquid_leg(sym, row_ltp(s), expiry, lot_map, eq_map, side, cfg)
        if leg and leg["lot"]:
            leg["rank"] = rank
            return leg
        print(f"  #{rank} {sym}: liquid {side} nahi mila, agla stock dekhte hain")
    return None


# =========================
# EXCEL + CSV LOG
# =========================

XL_COLS = ["Stock", "Type", "Strike", "Lot Size", "Lots", "Qty", "Entry Time", "Exit Time",
           "Spot Entry", "Spot Exit", "Entry (₹)", "Exit (₹)", "PnL (₹)", "Reason"]


def save_to_excel(trade_summary, mtm_logs, filename):
    wb = openpyxl.Workbook()
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    green_fill = PatternFill("solid", fgColor="C6EFCE")
    red_fill = PatternFill("solid", fgColor="FFC7CE")

    ws1 = wb.active
    ws1.title = "Trade Summary"
    ws1.append(XL_COLS)
    for cell in ws1[1]:
        cell.fill, cell.font = header_fill, header_font
        cell.alignment = Alignment(horizontal="center")
    pnl_col = XL_COLS.index("PnL (₹)") + 1
    for t in trade_summary:
        ws1.append([t[c] for c in XL_COLS])
        ws1.cell(ws1.max_row, pnl_col).fill = green_fill if t["PnL (₹)"] >= 0 else red_fill
    total = round(sum(t["PnL (₹)"] for t in trade_summary), 2)
    row = [""] * len(XL_COLS)
    row[pnl_col - 2] = "TOTAL PNL →"
    row[pnl_col - 1] = total
    ws1.append(row)
    ws1.cell(ws1.max_row, pnl_col).font = Font(bold=True)
    ws1.cell(ws1.max_row, pnl_col).fill = green_fill if total >= 0 else red_fill
    for col in ws1.columns:
        ws1.column_dimensions[col[0].column_letter].width = 15

    for title, log in mtm_logs:
        ws = wb.create_sheet(title)
        ws.append(["Time", "Option LTP (₹)", "MTM (₹)"])
        for cell in ws[1]:
            cell.fill, cell.font = header_fill, header_font
            cell.alignment = Alignment(horizontal="center")
        for r in log:
            ws.append([r["time"], r["ltp"], r["mtm"]])
            ws.cell(ws.max_row, 3).fill = green_fill if r["mtm"] >= 0 else red_fill
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 18

    wb.save(filename)
    print(f"Excel saved: {filename}")


def append_history(trade_summary, day):
    path = os.path.join(HERE, "trade_history.csv")
    new = not os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["date"] + XL_COLS)
        for t in trade_summary:
            w.writerow([day] + [t[c] for c in XL_COLS])


# =========================
# TRADE ENGINE (paper)
# =========================

def _tick_up(x):
    """Price ko upar ke nazdeeki 0.05 tick par laao."""
    return round(math.ceil(round(x / TICK, 6)) * TICK, 2)


def calc_levels(entry, qty, sl_amount, target_amount):
    """
    SL price  = entry - (SL rupees / qty)     [tick par upar round, taaki loss SL se zyada na ho]
    Target    = entry + (Target rupees / qty)
    Misal: lot 25, entry 1000, SL 2500 -> 1000 - 2500/25 = 900.
    Agar premium itna gir hi nahi sakta ki SL rupees ka loss ho, toh SL = None.
    """
    sl_raw = entry - sl_amount / qty
    sl_price = _tick_up(sl_raw) if sl_raw >= TICK else None
    return sl_price, _tick_up(entry + target_amount / qty)


def make_leg(direction, info, entry, lots, entry_time, sl_amount, target_amount):
    qty = int(info["lot"]) * int(lots)
    sl_price, target_price = calc_levels(float(entry), qty, sl_amount, target_amount)
    return {"direction": direction, "opt": "CE" if direction == "CALL" else "PE",
            "sl_price": sl_price, "target_price": target_price,
            "symbol": info["symbol"], "key": info["key"], "eq_key": info["eq_key"],
            "strike": info["strike"], "info": info,
            "entry": float(entry), "spot_entry": float(info["spot"]), "spot": float(info["spot"]),
            "spot_exit": None, "lot": int(info["lot"]), "lots": int(lots),
            "qty": int(info["lot"]) * int(lots),
            "expiry": info["expiry"], "entry_time": entry_time,
            "status": "OPEN", "ltp": float(entry),
            "exit_price": None, "exit_time": None, "reason": None, "log": []}


def icon(lg):
    return "🟢" if lg["opt"] == "CE" else "🔴"


def leg_name(lg):
    return f"{lg['symbol']} {lg['strike']:g} {lg['opt']}"


def leg_mtm(lg):
    price = lg["ltp"] if lg["status"] == "OPEN" else lg["exit_price"]
    return round((price - lg["entry"]) * lg["qty"], 2)


def leg_entry_block(lg):
    i = lg["info"]
    lines = [f"{icon(lg)} BUY {lg['direction']}",
             f"Stock: {lg['symbol']}"]
    if i.get("rank", 1) > 1:
        lines.append(f"(List mein #{i['rank']} stock; upar ke stocks lot size ya liquidity ki wajah se skip hue)")
    lines += [f"Spot (CMP): ₹{lg['spot_entry']:.2f}",
              f"Strike: {lg['strike']:g} ({i['label']})",
              f"Option type: {lg['direction']} ({lg['opt']})",
              f"Entry price: ₹{lg['entry']:.2f}",
              f"Bid / Ask: ₹{i['bid']:.2f} / ₹{i['ask']:.2f} (spread {i['spread']:.1f}%)",
              f"Open Interest: {int(i['oi']):,}",
              f"Lot size: {lg['lot']} × {lg['lots']} lot = Qty {lg['qty']}",
              f"Expiry: {lg['expiry']}"]
    if lg["sl_price"] is not None:
        loss = (lg["entry"] - lg["sl_price"]) * lg["qty"]
        lines.append(f"🛑 SL price: ₹{lg['sl_price']:.2f} (loss ≈ ₹{loss:,.0f})")
    else:
        lines.append("🛑 SL: nahi (is premium par poora loss bhi SL se kam hai)")
    gain = (lg["target_price"] - lg["entry"]) * lg["qty"]
    lines.append(f"🎯 Target price: ₹{lg['target_price']:.2f} (profit ≈ ₹{gain:,.0f})")
    return "\n".join(lines)


# ---- State (network ke bina; tick thread mein bhi safe) ----
STATE_LOCK = threading.Lock()


def close_state(lg, price, now_sec, reason):
    lg["exit_price"], lg["exit_time"], lg["reason"], lg["status"] = price, now_sec, reason, "CLOSED"
    lg["spot_exit"] = lg["spot"]


def check_exit(lg, price, now_sec):
    """SL ya Target price touch hua toh EXACT us price par band karo. True agar band hua."""
    if lg["status"] != "OPEN":
        return False
    if lg["sl_price"] is not None and price <= lg["sl_price"]:
        close_state(lg, lg["sl_price"], now_sec, "🛑 SL HIT")
        return True
    if price >= lg["target_price"]:
        close_state(lg, lg["target_price"], now_sec, "🎯 TARGET HIT")
        return True
    return False


def refresh_spot(lg):
    """REST se spot lo (feed na ho tab); band leg ka spot_exit bhi yahi hota hai."""
    sp = get_live_ltp(lg["eq_key"])
    if sp:
        lg["spot"] = sp
        if lg["status"] == "CLOSED":
            lg["spot_exit"] = sp


def announce_close(lg):
    tg(f"{lg['reason']}\n"
       f"{icon(lg)} {lg['direction']} — {leg_name(lg)}\n"
       f"Exit time: {lg['exit_time']}\n"
       f"Option price: ₹{lg['entry']:.2f} → ₹{lg['exit_price']:.2f}\n"
       f"Spot (CMP): ₹{lg['spot_entry']:.2f} → ₹{lg['spot_exit']:.2f}\n"
       f"Lot size: {lg['lot']} × {lg['lots']} lot = Qty {lg['qty']}\n"
       f"PnL: ₹{leg_mtm(lg):+.2f}")


def live_pnl_msg(legs, now_str, use_rest=True):
    lines = [f"💓 LIVE PNL — {now_str}"]
    total = 0.0
    for lg in legs:
        if lg["status"] == "OPEN" and use_rest:
            sp = get_live_ltp(lg["eq_key"])
            if sp:
                lg["spot"] = sp
        with STATE_LOCK:
            if lg["status"] == "OPEN":
                price, spot, st = lg["ltp"], lg["spot"], "OPEN"
            else:
                price, spot, st = lg["exit_price"], lg["spot_exit"] or lg["spot"], "CLOSED"
            m = leg_mtm(lg)
        total += m
        lines.append(f"\n{icon(lg)} {lg['direction']} {leg_name(lg)}  [{st}]\n"
                     f"Spot ₹{spot:.2f} | Option ₹{price:.2f}\n"
                     f"PnL ₹{m:+.2f}")
    lines.append(f"\n💰 Total PnL: ₹{total:+.2f}")
    return "\n".join(lines)


# =========================
# LIVE TICK FEED (Upstox websocket, official upstox-python-sdk)
# =========================

def extract_ticks(msg):
    """Feed message (dict) se {instrument_key: ltp}. 'ltpc' aur 'full' dono format chalte hain."""
    out = {}
    feeds = msg.get("feeds") if isinstance(msg, dict) else None
    for key, f in (feeds or {}).items():
        try:
            ltpc = (f or {}).get("ltpc")
            if not ltpc:
                ff = (f or {}).get("fullFeed", {}) or {}
                ltpc = (ff.get("marketFF") or ff.get("indexFF") or {}).get("ltpc")
            ltp = float((ltpc or {}).get("ltp"))
            if ltp > 0:
                out[key] = ltp
        except Exception:
            continue
    return out


class TickFeed:
    def __init__(self, token, keys, on_tick):
        self.token, self.keys, self.on_tick = token, list(keys), on_tick
        self.streamer = None
        self.last_tick = 0.0     # kisi bhi instrument ka aakhri tick (feed zinda hai ya nahi)
        self.error = None

    def start(self):
        try:
            import upstox_client

            cfg = upstox_client.Configuration()
            cfg.access_token = self.token
            self.streamer = upstox_client.MarketDataStreamerV3(
                upstox_client.ApiClient(cfg), self.keys, "ltpc")
            self.streamer.on("message", self._on_message)
            self.streamer.on("error", lambda *a: print(f"Feed error: {a}"))
            try:
                self.streamer.auto_reconnect(True, 5, 20)
            except Exception as e:
                print(f"auto_reconnect set nahi hua: {e}")
            threading.Thread(target=self.streamer.connect, daemon=True).start()
            return True
        except Exception as e:
            self.error = f"{type(e).__name__}: {e}"
            return False

    def _on_message(self, msg):
        ticks = extract_ticks(msg)
        if not ticks:
            return
        self.last_tick = time.time()
        for k, p in ticks.items():
            try:
                self.on_tick(k, p)
            except Exception as e:
                print(f"tick handler error: {e}")

    def healthy(self, stale):
        return self.last_tick > 0 and (time.time() - self.last_tick) <= stale

    def stop(self):
        try:
            if self.streamer:
                self.streamer.disconnect()
        except Exception:
            pass


def run_trade(cfg, legs, exit_time):
    """
    Feed chalu: har tick par SL/Target check (tick thread mein), exact price par exit.
    Feed band/khamosh (feed_stale_seconds): REST se har poll_seconds par check - purani tarah.
    Main loop har 1 sec: Telegram bhejna, minute log, heartbeat, time exit.
    """
    hb_secs = (1 if TEST_MODE else int(cfg["heartbeat_minutes"])) * 60
    poll = max(1, int(cfg["poll_seconds"]))
    stale = int(cfg["feed_stale_seconds"])
    last_hb = time.time()
    last_min = None
    last_rest = 0.0
    pending = []                      # tick thread mein band hue legs; announce main thread karta hai
    ce, pe = legs

    by_key = {}
    for lg in legs:
        by_key.setdefault(lg["key"], []).append(("opt", lg))
        by_key.setdefault(lg["eq_key"], []).append(("spot", lg))

    def on_tick(key, price):
        now_sec = now_ist().strftime("%H:%M:%S")
        with STATE_LOCK:
            for kind, lg in by_key.get(key, []):
                if lg["status"] != "OPEN":
                    continue
                if kind == "spot":
                    lg["spot"] = price
                else:
                    lg["ltp"] = price
                    if check_exit(lg, price, now_sec):
                        pending.append(lg)

    def flush_pending():
        while True:
            with STATE_LOCK:
                if not pending:
                    return
                lg = pending.pop(0)
            announce_close(lg)

    feed, started_at = None, time.time()
    if cfg["use_websocket"]:
        feed = TickFeed(UPSTOX_ACCESS_TOKEN, list(by_key.keys()), on_tick)
        if not feed.start():
            tg_admin(f"⚠️ Live tick feed shuru nahi hua ({feed.error}). REST polling par chal raha hoon.")
            feed = None
    confirmed, fell_back = False, False

    print(f"\n{'Time':<6} | {'CE LTP':>8} | {'CE MTM':>10} | {'PE LTP':>8} | {'PE MTM':>10} | {'Combined':>10}")
    print("-" * 70)

    try:
        while True:
            now = now_ist()
            now_str, now_sec = now.strftime("%H:%M"), now.strftime("%H:%M:%S")
            feed_ok = bool(feed and feed.healthy(stale))
            past_exit = now_str >= exit_time or (not TEST_MODE and now_str > "15:30")

            if feed and feed_ok and not confirmed:
                confirmed = True
                tg_admin("📡 Live tick feed chalu: SL/Target ab har tick par check ho raha hai.")
            if feed and not feed_ok and not fell_back and time.time() - started_at > stale + 10:
                fell_back = True
                tg_admin("⚠️ Live feed se tick nahi aa rahe. REST polling "
                         f"(har {poll} sec) par chal raha hoon.")
            if feed and feed_ok and fell_back:
                fell_back = False
                tg_admin("✅ Live tick feed wapas chalu.")

            if past_exit:
                for lg in legs:
                    if lg["status"] != "OPEN":
                        continue
                    ltp = None if feed_ok else get_live_ltp(lg["key"])
                    with STATE_LOCK:
                        if lg["status"] != "OPEN":
                            continue
                        if ltp:
                            lg["ltp"] = ltp
                        close_state(lg, lg["ltp"], now_sec, "⏰ TIME EXIT")
                        pending.append(lg)
                    if not feed_ok:
                        refresh_spot(lg)
                flush_pending()
                break

            flush_pending()
            if all(lg["status"] == "CLOSED" for lg in legs):
                break

            if not feed_ok and time.time() - last_rest >= poll:      # REST fallback
                last_rest = time.time()
                for lg in legs:
                    if lg["status"] != "OPEN":
                        continue
                    ltp = get_live_ltp(lg["key"])
                    if not ltp:
                        continue
                    with STATE_LOCK:
                        lg["ltp"] = ltp
                        closed = check_exit(lg, ltp, now_sec)
                        if closed:
                            pending.append(lg)
                    if closed:
                        refresh_spot(lg)
                flush_pending()

            if now_str != last_min:
                last_min = now_str
                with STATE_LOCK:
                    for lg in legs:
                        if lg["status"] == "OPEN":
                            lg["log"].append({"time": now_str, "ltp": lg["ltp"], "mtm": leg_mtm(lg)})
                    row = (ce["ltp"], leg_mtm(ce), pe["ltp"], leg_mtm(pe))
                print(f"{now_str:<6} | ₹{row[0]:>7.2f} | ₹{row[1]:>9.2f} | "
                      f"₹{row[2]:>7.2f} | ₹{row[3]:>9.2f} | ₹{row[1] + row[3]:>9.2f}")

            if time.time() - last_hb >= hb_secs and any(lg["status"] == "OPEN" for lg in legs):
                last_hb = time.time()
                tg(live_pnl_msg(legs, now_str, use_rest=not feed_ok))

            time.sleep(1)
    finally:
        if feed:
            feed.stop()


def wait_until_entry(entry_time, grace_min):
    h, m = map(int, entry_time.split(":"))
    entry_dt = now_ist().replace(hour=h, minute=m, second=0, microsecond=0)
    if now_ist() > entry_dt + timedelta(minutes=grace_min):
        return False
    last_print = 0
    while now_ist() < entry_dt:
        if time.time() - last_print > 300:
            print(f"Entry {entry_time} ka wait... abhi {now_ist().strftime('%H:%M:%S')} IST")
            last_print = time.time()
        time.sleep(5)
    return True


# =========================
# MAIN
# =========================

def main():
    cfg = load_config()
    today = now_ist().date()
    tag = "🧪 TEST " if TEST_MODE else ""

    if not UPSTOX_ACCESS_TOKEN:
        tg_admin("❌ UPSTOX_ACCESS_TOKEN secret nahi mila. GitHub Secrets check karo.")
        sys.exit(1)

    if not TEST_MODE:
        # Market ka time nikal chuka (late/delayed run): chup-chaap band, kisi ko message nahi.
        if now_ist().strftime("%H:%M") >= cfg["exit_time"]:
            print("Exit time nikal chuka, kuch nahi karna.")
            return
        if today.weekday() >= 5:
            print("Weekend, exit.")
            return
        if another_run_active():
            return

    if not cfg["enabled"] and not TEST_MODE:
        tg_admin("⏸️ Algo config mein OFF hai (enabled=false). Aaj koi trade nahi.")
        return

    if not TEST_MODE:
        holiday, desc = is_market_holiday(today)
        if holiday:
            tg(f"🏖️ Aaj market holiday hai ({desc}). Paper trade nahi hoga.")
            return

    entry_time, exit_time = cfg["entry_time"], cfg["exit_time"]
    if TEST_MODE:
        n = now_ist()
        entry_time = n.strftime("%H:%M")
        exit_time = (n + timedelta(minutes=3)).strftime("%H:%M")

    # Entry ka time (+grace) nikal chuka toh friends ko kuch dikhaye bina chup-chaap skip
    h, m = map(int, entry_time.split(":"))
    if now_ist() > now_ist().replace(hour=h, minute=m, second=0, microsecond=0) + timedelta(
            minutes=int(cfg["late_start_grace_minutes"])):
        skip(f"algo late start hua (entry time {entry_time} nikal chuka), galat price se bachne ke liye skip")
        return

    tg(f"{tag}🟢 AAG Algo chalu (Paper trade)\n"
       f"Entry {entry_time} | Exit {exit_time}\n"
       f"Har leg: SL ₹{cfg['sl_amount']} | Target ₹{cfg['target_amount']} | Lots {cfg['num_lots']}"
       + (f"\nSirf lot size {cfg['max_lot_size']} se kam wale stocks" if cfg["max_lot_size"] else ""))

    lot_map, eq_map, fo_syms = load_fo_instruments()
    if not lot_map:
        skip("F&O instruments list load nahi hui", notify_all=True)
        return

    if not wait_until_entry(entry_time, int(cfg["late_start_grace_minutes"])):
        skip(f"algo late start hua (entry time {entry_time} nikal chuka), galat price se bachne ke liye skip")
        return

    res = get_gainers_losers(eq_map, fo_syms)
    if not res:
        skip("gainers/losers data nahi mila", notify_all=True)
        return
    gainers, losers = res
    n = int(cfg["list_size"])
    nifty = get_nifty_line()
    tg(f"{tag}📊 {entry_time} — F&O TOP GAINERS & LOSERS\n"
       + (f"{nifty}\n" if nifty else "") + "\n"
       + fmt_list("🟢 TOP GAINERS", gainers, n) + "\n\n"
       + fmt_list("🔴 TOP LOSERS", losers, n))

    expiry = get_monthly_expiry(today)
    ce = pick_leg(gainers, "CE", expiry, lot_map, eq_map, cfg)
    pe = pick_leg(losers, "PE", expiry, lot_map, eq_map, cfg)
    if not ce or not pe:
        skip(f"liquid strike nahi mila (CE: {'mila' if ce else 'nahi'}, PE: {'mila' if pe else 'nahi'}, "
             f"expiry {expiry}). max_spread_pct ya max_stock_tries badhakar dekho (lot size >= max_lot_size wale stocks bhi chhode jaate hain).", notify_all=True)
        return

    ce_entry = get_live_ltp(ce["key"]) or ce["ltp"]
    pe_entry = get_live_ltp(pe["key"]) or pe["ltp"]
    if not ce_entry or not pe_entry:
        skip("entry price (LTP) nahi mila", notify_all=True)
        return

    actual_entry = now_ist().strftime("%H:%M")
    lots = int(cfg["num_lots"])
    ce_leg = make_leg("CALL", ce, ce_entry, lots, actual_entry, cfg["sl_amount"], cfg["target_amount"])
    pe_leg = make_leg("PUT", pe, pe_entry, lots, actual_entry, cfg["sl_amount"], cfg["target_amount"])

    tg(f"{tag}🚀 TRADE ENTRY (Paper) — {actual_entry}\n\n"
       f"{leg_entry_block(ce_leg)}\n\n{leg_entry_block(pe_leg)}\n\n"
       f"SL ₹{cfg['sl_amount']} | Target ₹{cfg['target_amount']} (har leg)")

    legs = [ce_leg, pe_leg]
    run_trade(cfg, legs, exit_time)

    summary = []
    for lg in legs:
        summary.append({
            "Stock": lg["symbol"], "Type": lg["direction"], "Strike": f"{lg['strike']:g} {lg['opt']}",
            "Lot Size": lg["lot"], "Lots": lg["lots"], "Qty": lg["qty"],
            "Entry Time": lg["entry_time"], "Exit Time": lg["exit_time"],
            "Spot Entry": round(lg["spot_entry"], 2), "Spot Exit": round(lg["spot_exit"] or lg["spot"], 2),
            "Entry (₹)": lg["entry"], "Exit (₹)": lg["exit_price"], "PnL (₹)": leg_mtm(lg),
            "Reason": lg["reason"],
        })
    total = round(sum(t["PnL (₹)"] for t in summary), 2)

    lines = [f"{tag}📊 DAY SUMMARY (Paper) — {today}"]
    for t, lg in zip(summary, legs):
        lines.append(f"\n{icon(lg)} {t['Type']} {leg_name(lg)}\n"
                     f"Option: ₹{t['Entry (₹)']:.2f} → ₹{t['Exit (₹)']:.2f}\n"
                     f"Spot: ₹{t['Spot Entry']:.2f} → ₹{t['Spot Exit']:.2f}\n"
                     f"Lot {t['Lot Size']} × {t['Lots']} = Qty {t['Qty']}\n"
                     f"PnL ₹{t['PnL (₹)']:+.2f}  [{t['Reason']}]")
    lines.append(f"\n💰 TOTAL PnL: ₹{total:+.2f}")
    tg("\n".join(lines))

    xlsx = os.path.join(HERE, f"live_trade_log_{today}.xlsx")
    save_to_excel(summary, [("CE MTM Log", ce_leg["log"]), ("PE MTM Log", pe_leg["log"])], xlsx)
    tg_file(xlsx, f"Trade log {today}")
    if not TEST_MODE:
        append_history(summary, str(today))


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        tg_admin(f"❌ Algo crash hua: {type(e).__name__}: {e}")
        raise
    sys.stdout.flush()
    os._exit(0)        # websocket ka thread band na ho paaye toh bhi process khatam ho
