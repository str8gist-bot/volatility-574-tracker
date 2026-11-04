import os
import math
import time
import datetime
import threading
import requests
import logging
import warnings
import json
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from zoneinfo import ZoneInfo

from flask import Flask, jsonify, render_template_string
from flask_cors import CORS

try:
    import docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import yfinance as yf
    HAS_YF = True
except ImportError:
    HAS_YF = False

app = Flask(__name__)
CORS(app)
logging.getLogger("yfinance").setLevel(logging.CRITICAL)
warnings.filterwarnings("ignore")

PRICE_HISTORY_1M = {}
PRICE_HISTORY_3M = {}
PRICE_HISTORY_15M = {}
LAST_CMP_STORE = {}
PEAK_PRICE_STORE = {}
PREV_1M_STORE = {}
NSE_LIVE_CACHE = {}
AVG_VOL_CACHE = {}
NSE_LAST_FETCH = "Starting..."
MARKET_STATUS = "CHECKING"
NSE_FAIL_COUNT = 0
USE_YF_FALLBACK = False
LAST_TRADING_DATE = ""
LAST_PRIME_TIME = 0

SENTIMENT_CACHE = {}
SENTIMENT_RAW = ""
LAST_SENTIMENT_FETCH = 0
SENTIMENT_STATUS = "Not fetched"

IST = ZoneInfo("Asia/Kolkata")

CONFIG = {
    "SHEET_ID": "15Cfcs2QLL9kxapoutu8iaRwizGIwxHBukEV2C--gbVY",
    "GEMINI_API_KEY": "AQ.Ab8RN6LXlxYU47X5-CSfI_OMxxuzvpK71Mw8JHhlrAhSMYGGhQ",
    "GEMINI_MODELS": ["gemini-2.5-flash", "gemini-1.5-flash-8b", "gemini-flash-latest"],
}

NSE_SESSION = requests.Session()
NSE_SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/market-data/live-equity-market",
    "X-Requested-With": "XMLHttpRequest",
})

def now_ist():
    return datetime.datetime.now(IST)

def normalize_sym(s: str) -> str:
    return s.upper().replace(" ", "").replace("&", "").replace("-", "").replace("_", "").strip()

def safe_float(v, default=0.0):
    try:
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return default
        return float(v)
    except Exception:
        return default

def is_market_open() -> bool:
    n = now_ist()
    if n.weekday() >= 5:
        return False
    minutes = n.hour * 60 + n.minute
    return 555 <= minutes <= 930

def prime_nse(forced: bool = False) -> bool:
    global LAST_PRIME_TIME
    now = time.time()
    if not forced and (now - LAST_PRIME_TIME < 90):
        return True
    try:
        NSE_SESSION.get("https://www.nseindia.com", timeout=8)
        time.sleep(0.3)
        NSE_SESSION.get("https://www.nseindia.com/market-data/live-equity-market", timeout=8)
        time.sleep(0.2)
        NSE_SESSION.get("https://www.nseindia.com/api/allIndices", timeout=8)
        LAST_PRIME_TIME = time.time()
        print("[NSE] Session primed")
        return True
    except Exception as e:
        print(f"[NSE] Prime failed: {e}")
        return False

prime_nse(forced=True)

def parse_gviz_json(text):
    try:
        m = re.search(r'setResponse\((.*)\)', text, re.DOTALL)
        if not m:
            return None
        return json.loads(m.group(1))
    except Exception as e:
        print(f"[SENTIMENT] gviz parse error: {e}")
        return None

def fetch_sheet_tab_gviz(sheet_id, tab_name):
    try:
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:json&sheet={tab_name}"
        r = requests.get(url, timeout=10)
        if r.status_code != 200:
            return None
        j = parse_gviz_json(r.text)
        if not j:
            return None
        table = j.get("table", {})
        cols = [c.get("label","") for c in table.get("cols",[])]
        rows = []
        for row in table.get("rows",[]):
            vals = []
            for c in row.get("c",[]):
                vals.append(c.get("v","") if c and c.get("v") is not None else (c.get("f","") if c else ""))
            rows.append(vals)
        return {"cols": cols, "rows": rows}
    except Exception as e:
        print(f"[SENTIMENT] gviz fetch {tab_name} failed: {e}")
        return None

def fetch_sentiment_from_sheet():
    global SENTIMENT_CACHE, SENTIMENT_RAW, LAST_SENTIMENT_FETCH, SENTIMENT_STATUS
    now = time.time()
    if now - LAST_SENTIMENT_FETCH < 65:
        return
    LAST_SENTIMENT_FETCH = now
    try:
        sheet_id = CONFIG["SHEET_ID"]
        synced = fetch_sheet_tab_gviz(sheet_id, "Synced_View")
        stock_master = fetch_sheet_tab_gviz(sheet_id, "Stock_Master_DB_Stocks_Keywords")
        new_cache = {}
        if synced and synced["rows"]:
            raw_text = "\n".join([" ".join([str(x) for x in row]) for row in synced["rows"]])
            SENTIMENT_RAW = raw_text[:5000]
            for line in raw_text.split("\n"):
                m = re.search(r'([A-Z]{2,20})\s*\|\s*(STRONG BUY|BUY|HOLD|SELL|STRONG SELL)\s*\|\s*(.*)', line, re.I)
                if m:
                    ticker = m.group(1).upper()
                    call = m.group(2).upper()
                    reason = m.group(3)[:200]
                    norm = normalize_sym(ticker)
                    new_cache[norm] = {"call": call, "reason": reason, "keyword_match": "", "score": 3 if "STRONG" in call else (2 if "BUY" in call else (-2 if "SELL" in call else 0)), "source": "Synced_View", "orig": ticker}
        if stock_master and stock_master["rows"]:
            cols = [c.lower() for c in stock_master["cols"]]
            try:
                t_idx = next((i for i,c in enumerate(cols) if "ticker" in c), 1)
                type_idx = next((i for i,c in enumerate(cols) if "type" in c), 2)
                subj_idx = next((i for i,c in enumerate(cols) if "subject" in c), 4)
            except:
                t_idx, type_idx, subj_idx = 1,2,4
            last_rows = stock_master["rows"][-80:]
            stocks = [r for r in last_rows if len(r) > type_idx and "STOCK_TICKER" in str(r[type_idx]).upper()]
            keywords = [r for r in last_rows if len(r) > type_idx and ("MARKET_KEYWORD" in str(r[type_idx]).upper() or "KEYWORD" in str(r[type_idx]).upper())]
            positive_kw = ["EXCEEDS","POSITIVE","STRONG PERFORMANCE","TARGETED","PROGRESS","LEADERSHIP","NET PROFIT","NEW ORDER","BREAKOUT","UPGRADE","MERGER","BUYBACK","DIVIDEND"]
            negative_kw = ["LOSS","DOWNGRADE","NEGATIVE","NON-COMPLIANCE","LAWSUIT","CONCERN","BORDER","VOLATILITY","DECREASE","SETTLEMENT","ISSUE","CHALLENGES","WRITE-DOWN","SLOWDOWN"]
            regulatory_kw = ["SEBI","REGULATORY","BREAKING NEWS"]
            keyword_blob = " ".join([str(r[t_idx]) + " " + str(r[subj_idx]) for r in keywords[-30:]]).upper()
            for r in stocks[-40:]:
                if len(r) <= t_idx: continue
                ticker = str(r[t_idx]).strip().upper()
                if not ticker or ticker == "TICKER" or len(ticker) < 2 or ticker == "GENERAL": continue
                norm = normalize_sym(ticker)
                if norm in new_cache: continue
                score = 0
                k_match = ""
                subj_upper = str(r[subj_idx]).upper() if len(r) > subj_idx else ""
                for kw in positive_kw:
                    if kw in keyword_blob or kw in subj_upper:
                        score += 1
                        k_match += kw + ","
                for kw in negative_kw:
                    if kw in keyword_blob or kw in subj_upper:
                        score -= 1
                        k_match += kw + ","
                for kw in regulatory_kw:
                    if kw in keyword_blob:
                        k_match += kw + ","
                call = "BUY" if score >=2 else ("SELL" if score <= -2 else "HOLD")
                new_cache[norm] = {"call": call, "reason": f"Keyword sync: {k_match[:100]} | Subj: {subj_upper[:80]}", "keyword_match": k_match[:120], "score": score, "source": "Stock_Master", "orig": ticker}
        if new_cache:
            SENTIMENT_CACHE = new_cache
            SENTIMENT_STATUS = f"OK - {len(new_cache)} tickers at {now_ist().strftime('%H:%M:%S')}"
            print(f"[SENTIMENT] {SENTIMENT_STATUS}")
        else:
            SENTIMENT_STATUS = f"No data - Sheet empty? at {now_ist().strftime('%H:%M:%S')}"
    except Exception as e:
        SENTIMENT_STATUS = f"Error: {e}"
        print(f"[SENTIMENT] {SENTIMENT_STATUS}")

def fetch_one_nse(orig: str):
    try:
        symbol = orig.replace("&", "%26")
        url = f"https://www.nseindia.com/api/quote-equity?symbol={symbol}"
        r = NSE_SESSION.get(url, timeout=4)
        if r.status_code in (401, 403, 406, 429):
            return None
        if r.status_code != 200:
            return None
        j = r.json()
        pi = j.get("priceInfo", {})
        last = safe_float(pi.get("lastPrice"))
        if last <= 0:
            return None
        open_px = safe_float(pi.get("open"))
        high_px = safe_float(pi.get("intraDayHighLow", {}).get("max"), last)
        low_px = safe_float(pi.get("intraDayHighLow", {}).get("min"), last)
        prev_close = safe_float(pi.get("previousClose"), open_px or last)
        if open_px <= 0:
            open_px = prev_close
        change_open = ((last - open_px) / open_px * 100) if open_px > 0 else 0.0
        vol = safe_float(j.get("tradeInfo", {}).get("totalTradedVolume") or j.get("securityWiseDP", {}).get("quantityTraded") or pi.get("totalTradedVolume"))
        return {"orig": orig, "norm": normalize_sym(orig), "close": last, "open": open_px, "high": high_px, "low": low_px, "prev_close": prev_close, "change_from_open": change_open, "volume": vol, "source": "NSE"}
    except Exception:
        return None

def fetch_one_yf(orig: str):
    if not HAS_YF:
        return None
    try:
        t = yf.Ticker(orig + ".NS")
        hist = t.history(period="1d", interval="1m")
        if hist.empty:
            hist = t.history(period="1d")
        if hist.empty:
            return None
        last = safe_float(hist["Close"].iloc[-1])
        open_px = safe_float(hist["Open"].iloc[0])
        high = safe_float(hist["High"].max())
        low = safe_float(hist["Low"].min())
        vol = safe_float(hist["Volume"].sum()) if "Volume" in hist else 0.0
        change_open = ((last - open_px) / open_px * 100) if open_px > 0 else 0.0
        return {"orig": orig, "norm": normalize_sym(orig), "close": last, "open": open_px, "high": high, "low": low, "prev_close": open_px, "change_from_open": change_open, "volume": vol, "source": "YF"}
    except Exception:
        return None

def fetch_avg_vol_safe(orig: str) -> float:
    norm = normalize_sym(orig)
    if norm in AVG_VOL_CACHE:
        return AVG_VOL_CACHE[norm]
    if not HAS_YF:
        AVG_VOL_CACHE[norm] = 0.0
        return 0.0
    try:
        t = yf.Ticker(orig + ".NS")
        hist = t.history(period="20d")
        if not hist.empty and "Volume" in hist:
            avg = float(hist["Volume"].mean())
            if avg > 0:
                AVG_VOL_CACHE[norm] = avg
                return avg
    except Exception:
        pass
    AVG_VOL_CACHE[norm] = 0.0
    return 0.0

def load_watchlist(doc_path: str = "watchlist.docx"):
    orig = []
    # Try multiple paths for Render
    for p in [doc_path, os.path.join(os.path.dirname(__file__), doc_path), "/opt/render/project/src/watchlist.docx"]:
        if HAS_DOCX and os.path.exists(p):
            try:
                doc = docx.Document(p)
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for line in cell.text.replace("\r", "\n").split("\n"):
                                clean = line.replace("NSE:", "").strip().upper().split(".")[0].strip()
                                if clean and clean not in ("STOCK", "SYMBOL") and clean not in orig and len(clean) >= 2:
                                    orig.append(clean)
                if orig:
                    print(f"[WATCHLIST] Loaded {len(orig)} from {p}")
                    break
            except Exception as e:
                print(f"[WATCHLIST] Failed {p}: {e}")
    if not orig:
        # Fallback list - will be replaced by bundled watchlist.docx
        orig = ["RELIANCE","TCS","INFY","HDFCBANK","ICICIBANK","FSL","BSE","NCC","SWIGGY","TATAMOTORS","RECLTD","YESBANK","ADANIENT","ADANIPORTS","HDFCBANK"]
        print("[WATCHLIST] Using fallback default")
    return orig

WATCHLIST_ORIG = load_watchlist()
SYMBOL_ORDER = {normalize_sym(s): i for i, s in enumerate(WATCHLIST_ORIG)}
print(f"[BOOT] Watchlist size = {len(WATCHLIST_ORIG)}")

if HAS_YF:
    def _preload_avg():
        for s in WATCHLIST_ORIG[:100]:
            fetch_avg_vol_safe(s)
            time.sleep(0.1)
    threading.Thread(target=_preload_avg, daemon=True).start()

def fetch_all_realtime() -> float:
    global NSE_LIVE_CACHE, NSE_LAST_FETCH, MARKET_STATUS, LAST_TRADING_DATE, NSE_FAIL_COUNT, USE_YF_FALLBACK
    now = now_ist()
    is_open = is_market_open()
    MARKET_STATUS = "LIVE MARKET OPEN" if is_open else "MARKET CLOSED"
    today_str = now.strftime("%Y-%m-%d")
    if is_open and today_str != LAST_TRADING_DATE and now.hour == 9 and now.minute < 25:
        PRICE_HISTORY_1M.clear(); PRICE_HISTORY_3M.clear(); PRICE_HISTORY_15M.clear()
        PEAK_PRICE_STORE.clear(); PREV_1M_STORE.clear(); LAST_CMP_STORE.clear(); NSE_LIVE_CACHE.clear()
        LAST_TRADING_DATE = today_str
    batch_size = 90
    n_batches = max(1, (len(WATCHLIST_ORIG) + batch_size - 1) // batch_size)
    start_idx = (int(time.time() / 3) % n_batches) * batch_size
    batch = WATCHLIST_ORIG[start_idx : start_idx + batch_size]
    combined = {}
    if not USE_YF_FALLBACK:
        with ThreadPoolExecutor(max_workers=15) as ex:
            futures = {ex.submit(fetch_one_nse, s): s for s in batch}
            for fut in as_completed(futures):
                try:
                    r = fut.result()
                    if r:
                        combined[r["norm"]] = r
                except: pass
        got = len(combined)
        if got < max(5, len(batch) // 4):
            NSE_FAIL_COUNT += 1
            print(f"[NSE] Only {got}/{len(batch)} – fail {NSE_FAIL_COUNT}/3")
            if NSE_FAIL_COUNT >= 3:
                USE_YF_FALLBACK = True
                def _reset():
                    global USE_YF_FALLBACK, NSE_FAIL_COUNT
                    USE_YF_FALLBACK = False; NSE_FAIL_COUNT = 0; prime_nse(forced=True); print("[NSE] Back to NSE")
                threading.Timer(40, _reset).start()
        else:
            NSE_FAIL_COUNT = 0
    if USE_YF_FALLBACK or len(combined) < max(5, len(batch) // 4):
        if HAS_YF:
            print(f"[YF] Fetching {len(batch)}")
            with ThreadPoolExecutor(max_workers=12) as ex:
                futures = {ex.submit(fetch_one_yf, s): s for s in batch}
                for fut in as_completed(futures):
                    try:
                        r = fut.result()
                        if r and r["norm"] not in combined:
                            combined[r["norm"]] = r
                    except: pass
    if combined:
        NSE_LIVE_CACHE.update(combined)
        NSE_LAST_FETCH = now.strftime("%H:%M:%S") + f" ({len(NSE_LIVE_CACHE)} total | batch {start_idx}-{start_idx + len(batch) - 1})"
        print(f"[TICK] {NSE_LAST_FETCH} | {SENTIMENT_STATUS}")
    return 0.8 if is_open else 6.0

def background_loop():
    while True:
        try:
            sleep_for = fetch_all_realtime()
        except Exception as e:
            print(f"[LOOP] {e}"); sleep_for = 5.0
        time.sleep(sleep_for)

def sentiment_loop():
    while True:
        try:
            fetch_sentiment_from_sheet()
        except Exception as e:
            print(f"[SENTIMENT LOOP] {e}")
        time.sleep(65)

threading.Thread(target=sentiment_loop, daemon=True).start()
threading.Thread(target=background_loop, daemon=True).start()

@app.route("/api/data")
def get_data():
    try:
        results = []; now_ts = time.time()
        for norm_name, data in list(NSE_LIVE_CACHE.items()):
            cmp = safe_float(data.get("close"))
            if cmp <= 0: continue
            open_px = safe_float(data.get("open")); high = safe_float(data.get("high"))
            vol_today = safe_float(data.get("volume")); change_open = safe_float(data.get("change_from_open"))
            if change_open == 0 and open_px > 0:
                change_open = (cmp - open_px) / open_px * 100
            volatility = ((high - open_px) / open_px * 100) if open_px > 0 else 0.0
            avg_vol = AVG_VOL_CACHE.get(norm_name, 0.0)
            if avg_vol == 0 and HAS_YF:
                avg_vol = fetch_avg_vol_safe(data.get("orig", norm_name))
            rvol = (vol_today / avg_vol) if avg_vol > 0 else 1.0
            if norm_name not in PRICE_HISTORY_1M: PRICE_HISTORY_1M[norm_name] = []
            PRICE_HISTORY_1M[norm_name].append((now_ts, cmp))
            PRICE_HISTORY_1M[norm_name] = [(t,p) for t,p in PRICE_HISTORY_1M[norm_name] if now_ts - t <= 65]
            price_1m_ago = PRICE_HISTORY_1M[norm_name][0][1] if PRICE_HISTORY_1M[norm_name] else cmp
            one_min_chg = (cmp - price_1m_ago) / price_1m_ago * 100 if price_1m_ago > 0 else 0.0
            if norm_name not in PRICE_HISTORY_3M: PRICE_HISTORY_3M[norm_name] = []
            PRICE_HISTORY_3M[norm_name].append((now_ts, cmp))
            PRICE_HISTORY_3M[norm_name] = [(t,p) for t,p in PRICE_HISTORY_3M[norm_name] if now_ts - t <= 185]
            price_3m_ago = PRICE_HISTORY_3M[norm_name][0][1] if PRICE_HISTORY_3M[norm_name] else cmp
            three_min_chg = (cmp - price_3m_ago) / price_3m_ago * 100 if price_3m_ago > 0 else 0.0
            if norm_name not in PRICE_HISTORY_15M: PRICE_HISTORY_15M[norm_name] = []
            PRICE_HISTORY_15M[norm_name].append((now_ts, cmp))
            PRICE_HISTORY_15M[norm_name] = [(t,p) for t,p in PRICE_HISTORY_15M[norm_name] if now_ts - t <= 910]
            price_15m_ago = PRICE_HISTORY_15M[norm_name][0][1] if PRICE_HISTORY_15M[norm_name] else cmp
            fifteen_min_chg = (cmp - price_15m_ago) / price_15m_ago * 100 if price_15m_ago > 0 else 0.0
            peak = PEAK_PRICE_STORE.get(norm_name, cmp)
            if cmp > peak:
                PEAK_PRICE_STORE[norm_name] = cmp; peak = cmp
            pullback = ((peak - cmp) / peak * 100) if peak > 0 else 0.0
            prev_1m = PREV_1M_STORE.get(norm_name, 0.0)
            is_momentum_drop = (prev_1m >= 0.30) and (one_min_chg < 0)
            PREV_1M_STORE[norm_name] = one_min_chg
            is_15m_breakout = (fifteen_min_chg >= 1.00) and (change_open > 0)
            is_false_breakout = (fifteen_min_chg >= 1.00) and (pullback >= 0.50 or change_open < 0)
            is_healthy_pullback = (one_min_chg < 0) and (rvol >= 1.5) and (change_open > 0)
            is_continuous_surger = (one_min_chg >= 0.30) and (cmp > LAST_CMP_STORE.get(norm_name, cmp))
            LAST_CMP_STORE[norm_name] = cmp
            vol_chg_pct = ((vol_today - avg_vol) / avg_vol * 100) if avg_vol > 0 and vol_today > 0 else 0.0
            is_1m_gt_1pct = one_min_chg >= 1.0
            sent = SENTIMENT_CACHE.get(norm_name, {})
            sent_call = sent.get("call", ""); sent_reason = sent.get("reason", ""); sent_score = sent.get("score", 0); sent_kw = sent.get("keyword_match", "")
            final_call = "HOLD"; final_color = "#6b7280"; final_reason = ""
            if is_false_breakout:
                final_call = "SELL"; final_color = "#ef4444"; final_reason = f"False breakout {fifteen_min_chg:.1f}% + pullback {pullback:.1f}%"
                if sent_call in ("SELL","STRONG SELL"):
                    final_call = "STRONG SELL"; final_reason += f" | Confirmed: {sent_kw}"
            elif is_15m_breakout and is_1m_gt_1pct:
                if sent_call in ("BUY","STRONG BUY"):
                    final_call = "STRONG BUY"; final_color = "#059669"; final_reason = f"15M breakout + {sent_call} + {sent_kw}"
                elif sent_call == "SELL":
                    final_call = "HOLD"; final_color = "#f59e0b"; final_reason = f"Breakout but Sentiment SELL ({sent_kw}) - avoid"
                else:
                    final_call = "BUY"; final_color = "#10b981"; final_reason = f"15M breakout {fifteen_min_chg:.1f}% + RVOL {rvol:.1f}x"
            elif is_healthy_pullback and sent_call in ("BUY","STRONG BUY"):
                final_call = "BUY DIP"; final_color = "#0ea5e9"; final_reason = f"Healthy pullback + {sent_call}: {sent_kw}"
            elif is_momentum_drop and sent_call in ("SELL","STRONG SELL"):
                final_call = "SELL"; final_color = "#ef4444"; final_reason = f"Momentum drop + SELL {sent_kw}"
            elif sent_call:
                final_call = sent_call
                if "BUY" in sent_call: final_color = "#10b981" if "STRONG" in sent_call else "#34d399"
                elif "SELL" in sent_call: final_color = "#ef4444"
                final_reason = sent_reason[:120]
            else:
                if one_min_chg >= 0.5 and change_open > 0 and rvol >=1.2:
                    final_call = "WATCH BUY"; final_color = "#8b5cf6"; final_reason = f"Momentum {one_min_chg:.1f}% + RVOL {rvol:.1f}x"
            results.append({"stock": data.get("orig", norm_name), "cmp": round(cmp,2), "change_open": round(change_open,2), "volatility": round(volatility,2), "rvol": round(rvol,2), "one_min_chg": round(one_min_chg,2), "three_min_chg": round(three_min_chg,2), "fifteen_min_chg": round(fifteen_min_chg,2), "vol_chg_pct": round(vol_chg_pct,2), "is_1m_gt_1pct": is_1m_gt_1pct, "is_continuous_surger": is_continuous_surger, "is_15m_breakout": is_15m_breakout, "is_false_breakout": is_false_breakout, "is_momentum_drop": is_momentum_drop, "is_healthy_pullback": is_healthy_pullback, "word_order": SYMBOL_ORDER.get(norm_name, 9999), "source": data.get("source","NSE"), "sentiment_call": sent_call, "sentiment_reason": sent_reason, "sentiment_kw": sent_kw, "sentiment_score": sent_score, "final_call": final_call, "final_color": final_color, "final_reason": final_reason})
        results.sort(key=lambda x: (-x["one_min_chg"], x["word_order"]))
        off_results = []
        watch_norms = set(SYMBOL_ORDER.keys())
        for norm, sent in SENTIMENT_CACHE.items():
            if norm in watch_norms: continue
            live = NSE_LIVE_CACHE.get(norm)
            if live:
                cmp = safe_float(live.get("close")); chg_open = safe_float(live.get("change_from_open")); vol = safe_float(live.get("volume")); avg_vol = AVG_VOL_CACHE.get(norm, 0.0); rvol = (vol / avg_vol) if avg_vol>0 else 0.0
            else:
                cmp = 0.0; chg_open = 0.0; rvol = 0.0
            if sent.get("call","") in ("BUY","STRONG BUY","BUY DIP","STRONG SELL","SELL"):
                off_results.append({"stock": sent.get("orig", norm), "norm": norm, "cmp": round(cmp,2) if cmp else 0, "change_open": round(chg_open,2), "rvol": round(rvol,2), "sentiment_call": sent.get("call",""), "sentiment_kw": sent.get("keyword_match","") or sent.get("reason","")[:80], "sentiment_reason": sent.get("reason",""), "source": sent.get("source","")})
        off_results.sort(key=lambda x: (0 if "BUY" in x["sentiment_call"] else 1, x["stock"]))
        return jsonify({"results": results, "off_watchlist": off_results, "nse_time": NSE_LAST_FETCH, "market_status": MARKET_STATUS, "is_open": is_market_open(), "cache_size": len(NSE_LIVE_CACHE), "watchlist_size": len(WATCHLIST_ORIG), "using_fallback": USE_YF_FALLBACK, "sentiment_status": SENTIMENT_STATUS, "sentiment_count": len(SENTIMENT_CACHE), "sentiment_raw": SENTIMENT_RAW[:1000]})
    except Exception as e:
        return jsonify({"results": [], "error": str(e), "watchlist_size": len(WATCHLIST_ORIG)})

@app.route("/api/sentiment")
def sentiment_api():
    return jsonify({"cache": SENTIMENT_CACHE, "status": SENTIMENT_STATUS, "raw": SENTIMENT_RAW[:2000], "count": len(SENTIMENT_CACHE)})

@app.route("/")
def index():
    return render_template_string("""
<!DOCTYPE html><html><head><title>Live Tracker – Stocks × Keywords Sync</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>body{font-family:Arial,sans-serif;margin:12px;background:#f4f4f9}table{width:100%;border-collapse:collapse;background:white;font-size:.85em}th,td{padding:7px 8px;border:1px solid #ddd}th{background:#111827;color:white;cursor:pointer;position:sticky;top:0}th.sortable{color:#60a5fa;text-decoration:underline}.rising{color:green;font-weight:bold}.falling{color:red;font-weight:bold}.continuous-surge{background:#e2f0d9}.standard-highlight{background:rgba(16,185,129,.08)}.surge-1pct-row{background:linear-gradient(90deg,rgba(255,215,0,.30),rgba(255,215,0,.06))!important;border-left:5px solid #FFD700!important}.flicker-1pct{font-weight:900;font-size:1.05em;padding:2px 7px;border-radius:6px;animation:flicker3 .7s infinite alternate;display:inline-block}@keyframes flicker3{0%{background:#FF1493;color:#fff;box-shadow:0 0 8px #FF1493}50%{background:#00FFFF;color:#000;box-shadow:0 0 12px #00FFFF}100%{background:#39FF14;color:#000;box-shadow:0 0 14px #39FF14}}.badge-1m-surge{background:#FF1493;color:#fff;padding:2px 8px;border-radius:10px;font-weight:900;font-size:.75em;animation:pulse .7s infinite}@keyframes pulse{0%{transform:scale(1)}50%{transform:scale(1.12)}100%{transform:scale(1)}}.badge-15m{background:#8b5cf6;color:white;padding:2px 6px;border-radius:4px;font-size:.7em}.badge-danger{background:#ef4444;color:white;padding:2px 6px;border-radius:4px;font-size:.7em}.badge-warning{background:#f59e0b;color:white;padding:2px 6px;border-radius:4px;font-size:.7em}.badge-info{background:#3b82f6;color:white;padding:2px 6px;border-radius:4px;font-size:.7em}.final-badge{padding:4px 10px;border-radius:12px;font-weight:900;color:white;display:inline-block;min-width:90px;text-align:center}.status{font-size:.9em;color:#555;margin-bottom:8px;line-height:1.4}.tooltip{position:relative;cursor:help}.tooltip:hover::after{content:attr(data-tip);position:absolute;left:0;top:100%;background:#111;color:#fff;padding:6px 8px;border-radius:4px;white-space:normal;width:260px;z-index:99;font-size:.85em}</style>
</head><body>
<h2>Live Tracker – Tech (1m/3m/15m) × Sentiment (Stocks×Keywords) Sync – Render Cloud</h2>
<div class="status">@ <span id="time">–</span> | <span id="market">–</span> | Watchlist: <span id="wsize">–</span> | Cache: <span id="cache">0</span> | Source: <span id="src">–</span> | Sentiment: <span id="sentiment">–</span> (<span id="sentCount">0</span>)<div style="margin-top:4px;font-size:.85em;color:#666;" id="sentRaw"></div></div>
<h3>Main Watchlist – Live Tech × Sentiment Sync (574 rotating)</h3>
<table><thead><tr><th>Stock</th><th>CMP</th><th class="sortable" onclick="toggleSort('one_min_chg')">1m% ↕</th><th class="sortable" onclick="toggleSort('three_min_chg')">3m% ↕</th><th class="sortable" onclick="toggleSort('fifteen_min_chg')">15m% ↕</th><th class="sortable" onclick="toggleSort('change_open')">Chg Open% ↕</th><th class="sortable" onclick="toggleSort('volatility')">Volatility% ↕</th><th>RVOL</th><th class="sortable" onclick="toggleSort('vol_chg_pct')">Vol Chg% ↕</th><th>Tech Signal</th><th>Sentiment</th><th class="sortable" onclick="toggleSort('sentiment_score')">Final Call ↕</th><th>Reason</th></tr></thead><tbody id="body"><tr><td colspan="13">Loading live feed + sentiment… first batch 5 sec</td></tr></tbody></table>
<h3 style="margin-top:30px;">Off-Watchlist BUYs from Sheet (RECLTD, YESBANK, FSL, HDFCBANK etc)</h3><div style="font-size:.85em;color:#666;margin-bottom:8px;">Tickers from Sheet 15Cfcs2... that have BUY/SELL but NOT in watchlist.docx batch. Add to docx to get live 1m%/15m%.</div>
<table><thead><tr><th>Stock (Off-list)</th><th>CMP</th><th>Chg Open%</th><th>RVOL</th><th>Sentiment Call</th><th>Keyword Match</th><th>Reason</th><th>Source</th></tr></thead><tbody id="offBody"><tr><td colspan="8">Loading...</td></tr></tbody></table>
<script>
let sortCol='one_min_chg', asc=false;
function toggleSort(c){ if(sortCol===c) asc=!asc; else{sortCol=c; asc=false;} refresh(); }
async function refresh(){
 try{
  const r = await fetch('/api/data'); const d = await r.json();
  document.getElementById('time').innerText = d.nse_time || '–';
  document.getElementById('market').innerText = d.market_status || '–';
  document.getElementById('wsize').innerText = d.watchlist_size || 0;
  document.getElementById('cache').innerText = d.cache_size || 0;
  document.getElementById('src').innerText = d.using_fallback ? 'Yahoo (fallback)' : 'NSE';
  document.getElementById('sentiment').innerText = d.sentiment_status || '–';
  document.getElementById('sentCount').innerText = d.sentiment_count || 0;
  document.getElementById('sentRaw').innerText = (d.sentiment_raw || '').substring(0,180);
  const tbody = document.getElementById('body');
  if(!d.results || !d.results.length){ tbody.innerHTML = '<tr><td colspan="13">Fetching… Cache '+d.cache_size+' | Watchlist '+d.watchlist_size+' | '+(d.nse_time||'')+'</td></tr>'; return; }
  let sorted=[...d.results]; if(sortCol){ sorted.sort((a,b)=> asc ? (a[sortCol]||0)-(b[sortCol]||0) : (b[sortCol]||0)-(a[sortCol]||0)); }
  tbody.innerHTML = sorted.map(i=>{
   let techBadge=''; if(i.is_false_breakout) techBadge='<span class="badge-danger">🛑 FALSE BREAKOUT</span>'; else if(i.is_momentum_drop) techBadge='<span class="badge-warning">⚠ MOMENTUM DROP</span>'; else if(i.is_healthy_pullback) techBadge='<span class="badge-info">🔍 PULLBACK</span>'; else if(i.is_15m_breakout) techBadge=`<span class="badge-15m">⚡ 15M +${i.fifteen_min_chg}%</span>`; if(i.is_1m_gt_1pct) techBadge+=` <span class="badge-1m-surge">🔥 1m >1%</span>`;
   const rowClass = i.is_1m_gt_1pct ? 'surge-1pct-row' : (i.is_continuous_surger ? 'continuous-surge' : (i.one_min_chg>=0.30?'standard-highlight':''));
   const stockDisplay = i.is_1m_gt_1pct ? `<span class="flicker-1pct">${i.stock}</span>` : `<b>${i.stock}</b>`;
   const sentDisplay = i.sentiment_call ? `<span style="font-weight:700;color:${i.final_color}">${i.sentiment_call}</span><br><small style="color:#666">${(i.sentiment_kw||'').substring(0,40)}</small>` : '<small style="color:#999">No keyword</small>';
   const finalBadge = `<span class="final-badge" style="background:${i.final_color}">${i.final_call}</span>`;
   return `<tr class="${rowClass}"><td>${stockDisplay}</td><td>₹${i.cmp}</td><td class="${i.one_min_chg>=0?'rising':'falling'}">${i.one_min_chg>0?'+':''}${i.one_min_chg}%</td><td class="${i.three_min_chg>=0?'rising':'falling'}">${i.three_min_chg>0?'+':''}${i.three_min_chg}%</td><td class="${i.fifteen_min_chg>=0?'rising':'falling'}">${i.fifteen_min_chg>0?'+':''}${i.fifteen_min_chg}%</td><td class="${i.change_open>=0?'rising':'falling'}">${i.change_open>0?'+':''}${i.change_open}%</td><td>${i.volatility>0?'+':''}${i.volatility}%</td><td>${i.rvol.toFixed(2)}x</td><td class="${i.vol_chg_pct>=0?'rising':'falling'}">${i.vol_chg_pct>0?'+':''}${i.vol_chg_pct.toLocaleString()}%</td><td>${techBadge}</td><td class="tooltip" data-tip="${(i.sentiment_reason||'').replace(/"/g,'&quot;')}">${sentDisplay}</td><td>${finalBadge}</td><td style="font-size:.82em;max-width:220px;word-wrap:break-word;">${i.final_reason||''}</td></tr>`;
  }).join('');
  const offBody = document.getElementById('offBody');
  if(d.off_watchlist && d.off_watchlist.length){
   offBody.innerHTML = d.off_watchlist.map(o=>{let color = o.sentiment_call.includes('BUY') ? '#10b981' : (o.sentiment_call.includes('SELL') ? '#ef4444' : '#6b7280'); return `<tr><td><b>${o.stock}</b></td><td>${o.cmp? '₹'+o.cmp : '<small style=color:#999>Not cached</small>'}</td><td class="${o.change_open>=0?'rising':'falling'}">${o.change_open>0?'+':''}${o.change_open}%</td><td>${o.rvol? o.rvol+'x' : '-'}</td><td><span class="final-badge" style="background:${color}">${o.sentiment_call}</span></td><td style="font-size:.82em;">${(o.sentiment_kw||'').substring(0,80)}</td><td style="font-size:.82em;max-width:200px;">${(o.sentiment_reason||'').substring(0,120)}</td><td><small>${o.source||''}</small></td></tr>`;}).join('');
  } else { offBody.innerHTML = '<tr><td colspan="8">No off-watchlist BUYs - all '+ (d.sentiment_count||0) +' synced tickers already in watchlist, or sentiment empty.</td></tr>'; }
 }catch(e){ console.error(e); }
}
refresh(); setInterval(refresh, 2000);
</script></body></html>
""")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    print(f"[START] Render Cloud – Watchlist {len(WATCHLIST_ORIG)} | http://0.0.0.0:{port}")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
