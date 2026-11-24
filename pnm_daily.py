"""
P&M demand monitor (Google Ads search vs Google Trends) - local dashboard
=========================================================================
Start:  double-click start_dashboard.bat   (or: python pnm_daily.py)
        -> opens http://localhost:8765 . Keep the black window open.

Data is stored in TiDB Cloud (tidb.yaml) and read fresh each time the dashboard opens.
First start loads full history.
While running it refreshes once a day after refresh_time:
  - new day: Enabled "packers" campaigns, full row
  - last 7 days: conversions re-pulled until they mature (paused campaigns keep their history)
  - Google Trends: last 90 days, rescaled onto stored history
Missed days (PC off) are caught up on the next refresh.

Other commands:
  python pnm_daily.py --test-api    fetch 3 days and print them
  python pnm_daily.py --once        refresh + Excel without the dashboard
"""

import argparse
import json
import os
import sys
import time
import warnings
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
import requests

warnings.filterwarnings("ignore")

# =====================================================================
# CONFIG — edit this block
# =====================================================================
CONFIG = {
    # ---------- Google Ads API ----------
    # Credentials live in google-ads.yaml (developer_token, client_id, client_secret,
    # refresh_token, login_customer_id = your MCC id without dashes, use_proto_plus: True)
    "google_ads_yaml": "google-ads.yaml",
    "customer_id": "9680382253",          # Porter India (968-038-2253), no dashes
    # Which campaigns to include: case-insensitive regex on campaign name ("" = all Search campaigns)
    "campaign_name_regex": "packers",   # campaign name contains "packers" (any case)
    "only_search_campaigns": True,
    
    # How to get city:
    #   "campaign_name" -> city = first city_to_geo key found in the campaign name
    #   "user_location" -> city = where users actually were (Google Ads location report)
    "city_source": "campaign_name",
    "city_aliases": {"NCR": "Delhi", "BLR": "Bangalore", "MUM": "Mumbai", "HYD": "Hyderabad",
                     "DEL": "Delhi", "CHN": "Chennai", "KOL": "Kolkata", "AMD": "Ahmedabad"},
    "api_version": "",                    # "" = auto-detect newest working version (e.g. "v22")
    "chunk_days": 31,                     # backfill pulled in chunks of this many days
    "gst_multiplier": 1.18,               # cost in every output = Google Ads cost x this

    # ---------- Refresh rules ----------
    "maturity_days": 7,            # conversions re-fetched for this many days back
    "refresh_cols": ["conversions"],  # columns overwritten on already-stored days
    "backfill_start": "2024-11-01",

    # ---------- Google Trends ----------
    "trends_keywords": ["packers and movers", "house shifting", "home shifting"],
    "trends_daily_window": 90,
    "trends_new_states_per_run": 3,   # full Trends history is loaded for at most this many states per run     # days per pull (Trends only gives daily data up to ~90 days)
    "trends_recent_overwrite": 7,
    # ---- demand planning (forecast, all-India, locations, brand) ----
    "forecast_keyword": "packers and movers",   # the keyword the 12-month forecast is built on
    "forecast_refresh_days": 7,                  # 5-year Trends history per state refreshed weekly
    "forecast_states_per_run": 5,                # at most this many 5-year pulls per run (rate limits)
    "forecast_threshold_pct": 12,                # weeks this far above/below normal get a budget note
    # Monthly search volumes from Google Ads Keyword Planner (exact numbers, incl. brand terms)
    "kp_refresh_days": 30,
    "kp_keywords": {
        "Generic": ["packers and movers", "packers and movers near me", "house shifting",
                    "home shifting services", "movers and packers"],
        "Porter": ["porter packers and movers", "porter packers", "porter house shifting"],
        "Competitors": ["nobroker packers and movers", "agarwal packers and movers",
                        "leo packers and movers"],
    },
    # Google's own city names where they differ from ours (Keyword Planner + Trends city view)
    "google_city_names": {"Bangalore": "Bengaluru", "Gurgaon": "Gurugram", "Delhi": "New Delhi"},  # recent days overwritten each run (Trends can revise them)
    # State-level only for India; cities in one state share a line. Verify codes on trends.google.com
    "city_to_geo": {
        "All India": "IN",
        "Mumbai": "IN-MH", "Pune": "IN-MH", "Nagpur": "IN-MH",
        "Delhi": "IN-DL", "New Delhi": "IN-DL", "Navi Mumbai": "IN-MH", "Thane": "IN-MH", "Bangalore": "IN-KA", "Bengaluru": "IN-KA",
        "Hyderabad": "IN-TG", "Chennai": "IN-TN", "Coimbatore": "IN-TN",
        "Kolkata": "IN-WB", "Ahmedabad": "IN-GJ", "Surat": "IN-GJ",
        "Jaipur": "IN-RJ", "Lucknow": "IN-UP", "Noida": "IN-UP",
        "Gurgaon": "IN-HR", "Gurugram": "IN-HR", "Chandigarh": "IN-CH",
        "Indore": "IN-MP", "Kochi": "IN-KL", "Ludhiana": "IN-PB",
    },

    # ---------- Analysis ----------
    "events_csv": "events.csv",
    "add_india_public_holidays": True,
    "event_window_days": 3,
    "spike_z": 2.5,
    "spend_change_pct": 30,

    "refresh_time": "09:00",       # daily auto-refresh time while the dashboard is running
    "port": 8765,                  # dashboard at http://localhost:8765
    "db_config": "tidb.yaml",      # TiDB Cloud connection (written by setup_credentials.py)
    "output_xlsx": "pnm_seasonality_output.xlsx",
}

BASE_COLS = ["impressions", "clicks", "cost", "conversions", "impr_share"]


# =====================================================================
# STORAGE (SQLite) — swap this section to change storage
# =====================================================================
def _read_kv(path):
    """Minimal reader for 'key: value' files (google-ads.yaml, tidb.yaml)."""
    out = {}
    if not os.path.exists(path):
        raise RuntimeError(f"{path} not found. Run: python setup_credentials.py")
    for line in open(path, encoding="utf-8-sig"):
        line = line.strip()
        if not line or line.startswith("#") or ":" not in line:
            continue
        k, v = line.split(":", 1)
        out[k.strip()] = v.strip().strip('"').strip("'")
    return out


def _db_settings(cfg):
    d = _read_kv(cfg["db_config"])
    for k in ("host", "user", "password"):
        if not d.get(k):
            raise RuntimeError(f"{k} missing in {cfg['db_config']}. Run: python setup_credentials.py")
    return d


_SCHEMA_DONE = {"ok": False}


def db_connect(cfg):
    """Connection to TiDB Cloud (MySQL protocol, TLS). Tables are created on first use."""
    import pymysql
    d = _db_settings(cfg)
    ssl = None
    if d.get("ssl", "true").lower() != "false":
        import certifi
        ssl = {"ca": d.get("ca") or certifi.where()}
    dbname = d.get("database") or "pnm_monitor"
    con = pymysql.connect(host=d["host"], port=int(d.get("port") or 4000), user=d["user"],
                          password=d["password"], ssl=ssl, autocommit=True, charset="utf8mb4",
                          connect_timeout=20, read_timeout=300, write_timeout=300)
    cur = con.cursor()
    if not _SCHEMA_DONE["ok"]:
        cur.execute(f"CREATE DATABASE IF NOT EXISTS `{dbname}`")
    cur.execute(f"USE `{dbname}`")
    if not _SCHEMA_DONE["ok"]:
        cur.execute("""CREATE TABLE IF NOT EXISTS ads_campaign_daily (
            date VARCHAR(10) NOT NULL, campaign_id VARCHAR(32) NOT NULL, campaign VARCHAR(255),
            city VARCHAR(64) NOT NULL, impressions DOUBLE, clicks DOUBLE, cost DOUBLE,
            conversions DOUBLE, impr_share DOUBLE, first_fetched_at VARCHAR(19), conv_updated_at VARCHAR(19),
            PRIMARY KEY (date, campaign_id, city))""")
        cur.execute("""CREATE TABLE IF NOT EXISTS trends_daily (
            date VARCHAR(10) NOT NULL, geo VARCHAR(16) NOT NULL, idx DOUBLE, fetched_at VARCHAR(19),
            PRIMARY KEY (date, geo))""")
        cur.execute("""CREATE TABLE IF NOT EXISTS trends_weekly_5y (
            geo VARCHAR(16) NOT NULL, week VARCHAR(10) NOT NULL, idx DOUBLE, fetched_at VARCHAR(19),
            PRIMARY KEY (geo, week))""")
        cur.execute("""CREATE TABLE IF NOT EXISTS trends_city_interest (
            city VARCHAR(64) NOT NULL PRIMARY KEY, idx DOUBLE, fetched_at VARCHAR(19))""")
        cur.execute("""CREATE TABLE IF NOT EXISTS kp_monthly (
            city VARCHAR(64) NOT NULL, keyword VARCHAR(128) NOT NULL, grp VARCHAR(32),
            ym VARCHAR(7) NOT NULL, searches DOUBLE, low_bid DOUBLE, high_bid DOUBLE, fetched_at VARCHAR(19),
            PRIMARY KEY (city, keyword, ym))""")
        cur.execute("""CREATE TABLE IF NOT EXISTS job_runs (
            job VARCHAR(64) NOT NULL PRIMARY KEY, last_ok VARCHAR(19), note TEXT)""")
        cur.execute("""CREATE TABLE IF NOT EXISTS dashboard_cache (
            id VARCHAR(16) NOT NULL PRIMARY KEY, payload LONGTEXT, xlsx LONGBLOB, updated_at VARCHAR(19))""")
        cur.execute("""CREATE TABLE IF NOT EXISTS run_log (
            id BIGINT AUTO_INCREMENT PRIMARY KEY, run_at VARCHAR(19), step VARCHAR(32), detail TEXT)""")
        _SCHEMA_DONE["ok"] = True
    return con


def q(con, sql, params=None):
    """Run a SELECT and return a DataFrame."""
    cur = con.cursor()
    cur.execute(sql, params)
    cols = [c[0] for c in cur.description]
    return pd.DataFrame(list(cur.fetchall()), columns=cols)


def log(con, step, detail):
    print(f"[{step}] {detail}")
    con.cursor().execute("INSERT INTO run_log (run_at, step, detail) VALUES (%s,%s,%s)",
                         (datetime.now().isoformat(timespec="seconds"), step, detail))


def stored_campaign_ids(con):
    return set(q(con, "SELECT DISTINCT campaign_id FROM ads_campaign_daily")["campaign_id"].astype(str))


def upsert_ads(con, df, cfg):
    """New (date, campaign, city) rows -> full insert. Existing rows -> only conversions updated.
    Rows not returned (e.g. a paused campaign) stay untouched."""
    if df.empty:
        return 0, 0
    now = datetime.now().isoformat(timespec="seconds")
    keys = set(map(tuple, q(con, "SELECT date, campaign_id, city FROM ads_campaign_daily WHERE date >= %s",
                            (str(df.date.min()),)).astype(str).values.tolist()))
    upd_cols = ", ".join(f"{c} = VALUES({c})" for c in cfg["refresh_cols"])
    sql = ("INSERT INTO ads_campaign_daily (date, campaign_id, campaign, city, impressions, clicks, cost, "
           "conversions, impr_share, first_fetched_at, conv_updated_at) "
           "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) "
           f"ON DUPLICATE KEY UPDATE {upd_cols}, conv_updated_at = VALUES(conv_updated_at)")
    rows, upd = [], 0
    for r in df.itertuples(index=False):
        k = (str(r.date), str(r.campaign_id), str(r.city))
        upd += k in keys
        rows.append((*k[:2], r.campaign, k[2], float(r.impressions), float(r.clicks), float(r.cost),
                     float(r.conversions), None if pd.isna(r.impr_share) else float(r.impr_share), now, now))
    cur = con.cursor()
    for i in range(0, len(rows), 500):
        cur.executemany(sql, rows[i:i + 500])
    return len(rows) - upd, upd


def to_city_daily(df):
    """Campaign rows -> date x city. Impression share = impression-weighted average."""
    df = df.copy()
    for c in ["impressions", "clicks", "cost", "conversions", "impr_share"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["is_w"] = df["impr_share"] * df["impressions"]
    df["is_imp"] = np.where(df["impr_share"].notna(), df["impressions"], 0)
    g = df.groupby(["date", "city"], as_index=False).agg(
        impressions=("impressions", "sum"), clicks=("clicks", "sum"), cost=("cost", "sum"),
        conversions=("conversions", "sum"), is_w=("is_w", "sum"), is_imp=("is_imp", "sum"))
    g["impr_share"] = g.is_w / g.is_imp.replace(0, np.nan)
    return g.drop(columns=["is_w", "is_imp"])


def read_ads(con, cfg):
    df = q(con, "SELECT * FROM ads_campaign_daily")
    if df.empty:
        return pd.DataFrame(columns=["date", "city"] + BASE_COLS)
    df["cost"] = pd.to_numeric(df["cost"]) * cfg["gst_multiplier"]     # all cost from here on includes GST
    df = df[df.city != "Unmapped"]
    allin = df.copy()
    allin["city"] = "All India"                                         # every city added together
    df = to_city_daily(pd.concat([df, allin], ignore_index=True))
    df["date"] = pd.to_datetime(df["date"])
    return df


def read_trends(con):
    df = q(con, "SELECT date, geo, idx FROM trends_daily")
    df["date"] = pd.to_datetime(df["date"])
    df["idx"] = pd.to_numeric(df["idx"])
    return df


def write_trends(con, geo, s, overwrite_from=None):
    """Insert days not stored; overwrite days >= overwrite_from."""
    now = datetime.now().isoformat(timespec="seconds")
    new, over = [], []
    for d, v in s.items():
        row = (d.strftime("%Y-%m-%d"), geo, float(v), now)
        (over if overwrite_from is not None and d >= overwrite_from else new).append(row)
    cur = con.cursor()
    if new:
        cur.executemany("INSERT IGNORE INTO trends_daily (date, geo, idx, fetched_at) VALUES (%s,%s,%s,%s)", new)
    if over:
        cur.executemany("INSERT INTO trends_daily (date, geo, idx, fetched_at) VALUES (%s,%s,%s,%s) "
                        "ON DUPLICATE KEY UPDATE idx = VALUES(idx), fetched_at = VALUES(fetched_at)", over)
    return len(new) + len(over)


# =====================================================================
# GOOGLE ADS API
# =====================================================================
import re


def _read_yaml(path):
    out = _read_kv(path)
    for k in ("developer_token", "client_id", "client_secret", "refresh_token"):
        if not out.get(k) or out[k].startswith("PASTE_"):
            raise RuntimeError(f"{k} missing in {path}. Run: python setup_credentials.py")
    return out


def _ads_err(r):
    """Short, readable Google Ads error: code + Google's own explanation."""
    try:
        j = r.json()
        j = j[0] if isinstance(j, list) else j
        e = j.get("error", {})
        parts = [e.get("message", "")]
        for d in e.get("details", []):
            for x in d.get("errors", []):
                parts.append(f"{list(x.get('errorCode', {}).values())} {x.get('message', '')}")
        return " | ".join(p for p in parts if p)
    except Exception:
        return r.text[:2000]


class AdsREST:
    """Google Ads API over plain HTTPS (REST). Works on any Python version."""

    def __init__(self, cfg):
        self.cfg = cfg
        self.creds = _read_yaml(cfg["google_ads_yaml"])
        self.token, self.token_exp, self.version = None, 0, None

    def _access_token(self):
        if self.token and time.time() < self.token_exp - 60:
            return self.token
        r = requests.post("https://oauth2.googleapis.com/token", data={
            "client_id": self.creds["client_id"], "client_secret": self.creds["client_secret"],
            "refresh_token": self.creds["refresh_token"], "grant_type": "refresh_token"}, timeout=60)
        if r.status_code != 200:
            raise RuntimeError(f"Google sign-in token refresh failed ({r.status_code}): {r.text[:300]}. "
                               "Run setup_credentials.py again.")
        j = r.json()
        self.token, self.token_exp = j["access_token"], time.time() + int(j.get("expires_in", 3600))
        return self.token

    def _headers(self):
        h = {"Authorization": f"Bearer {self._access_token()}",
             "developer-token": self.creds["developer_token"], "Content-Type": "application/json"}
        if self.creds.get("login_customer_id"):
            h["login-customer-id"] = self.creds["login_customer_id"].replace("-", "")
        return h

    def _post(self, version, query):
        url = (f"https://googleads.googleapis.com/{version}/customers/"
               f"{self.cfg['customer_id']}/googleAds:searchStream")
        return requests.post(url, headers=self._headers(), json={"query": query}, timeout=300)

    def _pick_version(self):
        """Find the newest API version Google currently accepts (versions change a few times a year)."""
        if self.version:
            return self.version
        if self.cfg.get("api_version"):
            self.version = self.cfg["api_version"]
            return self.version
        errors = []
        for v in range(40, 15, -1):
            r = self._post(f"v{v}", "SELECT customer.id FROM customer LIMIT 1")
            if r.status_code == 200:
                self.version = f"v{v}"
                print(f"   Google Ads API version: {self.version}")
                return self.version
            if r.status_code in (401, 403):
                raise RuntimeError(f"Google Ads access denied ({r.status_code}): {_ads_err(r)}")
            errors.append(r.status_code)
        raise RuntimeError(f"No working Google Ads API version found (status codes {set(errors)}).")

    def stream(self, query):
        v = self._pick_version()
        r = self._post(v, query)
        if r.status_code != 200:
            raise RuntimeError(f"Google Ads query failed ({r.status_code}): {_ads_err(r)}")
        for batch in r.json():
            for row in batch.get("results", []):
                yield row


def _city_from_name(name, cfg):
    low = name.lower()
    # longest names first so "Navi Mumbai" would win over "Mumbai" if both are listed
    for city in sorted(cfg["city_to_geo"], key=len, reverse=True):
        if city.lower() in low:
            return city
    for alias, city in cfg["city_aliases"].items():
        if re.search(rf"(^|[^a-z]){re.escape(alias.lower())}([^a-z]|$)", low):
            return city
    return "Unmapped"


def _campaign_where(cfg, start, end, enabled_only):
    w = [f"segments.date BETWEEN '{start:%Y-%m-%d}' AND '{end:%Y-%m-%d}'"]
    if cfg["only_search_campaigns"]:
        w.append("campaign.advertising_channel_type = 'SEARCH'")
    if enabled_only:
        w.append("campaign.status = 'ENABLED'")
    return " AND ".join(w)


def _query_campaign(cfg, start, end, enabled_only):
    return f"""
        SELECT segments.date, campaign.id, campaign.name, campaign.status,
               metrics.impressions, metrics.clicks, metrics.cost_micros,
               metrics.conversions, metrics.search_impression_share
        FROM campaign
        WHERE {_campaign_where(cfg, start, end, enabled_only)}"""


def _query_user_location(cfg, start, end, enabled_only):
    return f"""
        SELECT segments.date, campaign.id, campaign.name, campaign.status, segments.geo_target_city,
               metrics.impressions, metrics.clicks, metrics.cost_micros, metrics.conversions
        FROM user_location_view
        WHERE {_campaign_where(cfg, start, end, enabled_only)}"""


def _geo_names(api, resource_names):
    """Turn geoTargetConstants/1007768 into city names."""
    ids = sorted({r.split("/")[-1] for r in resource_names if r})
    names = {}
    for i in range(0, len(ids), 500):
        q = ("SELECT geo_target_constant.id, geo_target_constant.name FROM geo_target_constant "
             f"WHERE geo_target_constant.id IN ({','.join(ids[i:i + 500])})")
        for row in api.stream(q):
            g = row.get("geoTargetConstant", {})
            names[f"geoTargetConstants/{g.get('id')}"] = g.get("name")
    return names


_API = {}


def fetch_api(cfg, start, end, enabled_only=True):
    """Google Ads rows for [start, end], one row per date x campaign x city (raw cost, no GST)."""
    api = _API.get("api") or AdsREST(cfg)
    _API["api"] = api
    rx = re.compile(cfg["campaign_name_regex"], re.I) if cfg["campaign_name_regex"] else None
    use_geo = cfg["city_source"] == "user_location"
    rows, cur = [], start
    num = lambda x: float(x) if x not in (None, "") else 0.0
    while cur <= end:
        ce = min(cur + timedelta(days=cfg["chunk_days"] - 1), end)
        q = (_query_user_location if use_geo else _query_campaign)(cfg, cur, ce, enabled_only)
        for attempt in range(3):
            try:
                chunk = []
                for r in api.stream(q):
                    c, m, sg = r.get("campaign", {}), r.get("metrics", {}), r.get("segments", {})
                    name = c.get("name", "")
                    if rx and not rx.search(name):
                        continue
                    sis = m.get("searchImpressionShare")
                    chunk.append({
                        "date": sg.get("date"),
                        "campaign_id": str(c.get("id")),
                        "campaign": name,
                        "status": c.get("status", ""),
                        "geo_res": sg.get("geoTargetCity", "") if use_geo else "",
                        "impressions": num(m.get("impressions")),
                        "clicks": num(m.get("clicks")),
                        "cost": num(m.get("costMicros")) / 1e6,
                        "conversions": num(m.get("conversions")),
                        "impr_share": np.nan if (use_geo or sis in (None, "")) else float(sis),
                    })
                rows.extend(chunk)
                break
            except RuntimeError as e:
                if "access denied" in str(e) or "token refresh failed" in str(e) or attempt == 2:
                    raise RuntimeError(f"Google Ads API failed for {cur.date()}..{ce.date()}: {e}")
                print(f"   Ads API retry ({e})")
                time.sleep(15 * (attempt + 1))
            except requests.RequestException as e:
                if attempt == 2:
                    raise RuntimeError(f"Network error for {cur.date()}..{ce.date()}: {e}")
                time.sleep(15 * (attempt + 1))
        cur = ce + timedelta(days=1)
    cols = ["date", "campaign_id", "campaign", "status", "city"] + BASE_COLS
    if not rows:
        return pd.DataFrame(columns=cols)
    df = pd.DataFrame(rows)
    if use_geo:
        names = _geo_names(api, df.geo_res.unique())
        df["city"] = df.geo_res.map(names).fillna("Unknown location")
    else:
        df["city"] = df.campaign.map(lambda n: _city_from_name(n, cfg))
    keys = ["date", "campaign_id", "campaign", "status", "city"]
    df = df.groupby(keys, as_index=False).agg(
        impressions=("impressions", "sum"), clicks=("clicks", "sum"), cost=("cost", "sum"),
        conversions=("conversions", "sum"), impr_share=("impr_share", "mean"))
    return df[cols]


# =====================================================================
# GOOGLE TRENDS
# =====================================================================
def _trends_window(geo, cfg, start, end):
    from pytrends.request import TrendReq
    py = TrendReq(hl="en-IN", tz=-330)
    tf = f"{start:%Y-%m-%d} {end:%Y-%m-%d}"
    for attempt in range(4):
        try:
            py.build_payload(cfg["trends_keywords"], timeframe=tf, geo=geo)
            d = py.interest_over_time()
            break
        except Exception as e:
            wait = 30 * (attempt + 1)
            print(f"   Trends {geo} {tf} failed ({e}); retry in {wait}s")
            time.sleep(wait)
    else:
        raise RuntimeError(f"Trends failed for {geo} {tf}")
    if d.empty:
        return pd.Series(dtype=float)
    if "isPartial" in d.columns:
        partial = d["isPartial"].astype(bool)
        d = d[~partial].drop(columns="isPartial")   # drop incomplete latest day
    s = d.sum(axis=1).astype(float)
    s.index = pd.to_datetime(s.index)
    return s


def _scale_onto(base, new, exclude_recent_days=7):
    """Scale `new` onto `base` using the ratio of their overlap, ignoring the newest days."""
    ov = base.index.intersection(new.index)
    if exclude_recent_days > 0 and len(ov) > exclude_recent_days + 14:
        ov = ov[:-exclude_recent_days]
    if len(ov) == 0 or new.loc[ov].sum() == 0:
        return None
    return new * (base.loc[ov].sum() / new.loc[ov].sum())


def trends_backfill(con, geo, cfg, start, end):
    w = cfg["trends_daily_window"]
    stitched, cur = None, start
    while cur <= end:
        we = min(cur + timedelta(days=w - 1), end)
        s = _trends_window(geo, cfg, cur, we)
        if len(s):
            if stitched is None:
                stitched = s
            else:
                scaled = _scale_onto(stitched, s, exclude_recent_days=0)
                if scaled is not None:
                    stitched = pd.concat([stitched, scaled[~scaled.index.isin(stitched.index)]]).sort_index()
        if we >= end:
            break
        cur = we - timedelta(days=14)
        time.sleep(12)        # gentle pacing: Google Trends rate-limits fast callers
    if stitched is None:
        return 0
    stitched = stitched / stitched.max() * 100   # one-time anchor scale, fixed from now on
    c = db_connect(cfg)                          # fresh connection: Trends waits can be long
    try:
        return write_trends(c, geo, stitched)
    finally:
        c.close()


def trends_daily_update(con, geo, cfg, end):
    c = db_connect(cfg)
    stored = read_trends(c)
    c.close()
    base = stored[stored.geo == geo].set_index("date")["idx"].sort_index()
    if base.empty:
        return trends_backfill(con, geo, cfg, pd.Timestamp(cfg["backfill_start"]), end)
    start = end - timedelta(days=cfg["trends_daily_window"] - 1)
    new = _trends_window(geo, cfg, start, end)
    scaled = _scale_onto(base, new, cfg["trends_recent_overwrite"])
    if scaled is None:
        return 0
    overwrite_from = end - timedelta(days=cfg["trends_recent_overwrite"] - 1)
    c = db_connect(cfg)
    try:
        return write_trends(c, geo, scaled, overwrite_from=overwrite_from)
    finally:
        c.close()


# =====================================================================
# DAILY RUN / BACKFILL
# =====================================================================
def run_refresh(cfg, backfill=False):
    con = db_connect(cfg)
    today = pd.Timestamp(date.today())
    yday = today - timedelta(days=1)

    last = q(con, "SELECT MAX(date) AS d FROM ads_campaign_daily")["d"].iloc[0]
    if last is None:
        backfill = True
    if backfill:
        start = pd.Timestamp(cfg["backfill_start"])
        df = fetch_api(cfg, start, yday, enabled_only=True)
    else:
        # yesterday + previous days whose conversions are still maturing.
        # New data: Enabled campaigns only. Conversion refresh: also campaigns already stored,
        # so a campaign paused mid-window still gets its last days' conversions matured.
        start = min(yday - timedelta(days=cfg["maturity_days"] - 1),
                    pd.Timestamp(last) + timedelta(days=1))   # also fills any missed days
        df = fetch_api(cfg, start, yday, enabled_only=False)
        known = stored_campaign_ids(con)
        df = df[(df.status == "ENABLED") | (df.campaign_id.isin(known))]
    remap_unmapped(con, cfg)
    ins, upd = upsert_ads(con, df, cfg)
    unmapped = sorted(df.loc[df.city == "Unmapped", "campaign"].unique())
    log(con, "ads", f"{start.date()}..{yday.date()}: {ins} new rows, {upd} rows conversions refreshed, "
                    f"{df.campaign_id.nunique()} campaigns")
    if unmapped:
        log(con, "warn", f"campaigns with no city found: {unmapped[:10]}")

    cities = q(con, "SELECT DISTINCT city FROM ads_campaign_daily")["city"].tolist()
    nogeo = [c for c in cities if c not in cfg["city_to_geo"] and c != "Unmapped"]
    if nogeo:
        log(con, "warn", f"no Trends geo for: {nogeo}")
    have = set(q(con, "SELECT DISTINCT geo FROM trends_daily")["geo"])
    log(con, "refresh-ok", f"data up to {yday.date()}")   # Google Ads part is done
    con.close()

    hist_start = pd.Timestamp(cfg["backfill_start"])
    new_budget = cfg["trends_new_states_per_run"]
    geos = {cfg["city_to_geo"][c] for c in cities if c in cfg["city_to_geo"]} | {"IN"}
    for geo in sorted(geos, key=lambda g: (g != "IN", g)):
        if geo not in have:
            if new_budget <= 0:
                print(f"[trends] {geo}: history queued for a later run")
                continue
            new_budget -= 1
        try:
            n = (trends_backfill(None, geo, cfg, hist_start, yday) if geo not in have
                 else trends_daily_update(None, geo, cfg, yday))
            c2 = db_connect(cfg); log(c2, "trends", f"{geo}: {n} days written"); c2.close()
        except Exception as e:
            try:
                c2 = db_connect(cfg); log(c2, "trends-error", f"{geo}: {str(e)[:300]}"); c2.close()
            except Exception:
                print("trends-error", geo, e)
        time.sleep(10)


def remap_unmapped(con, cfg):
    """Campaigns stored as 'Unmapped' get their city once a new alias matches them."""
    if cfg["city_source"] != "campaign_name":
        return
    rows = q(con, "SELECT DISTINCT campaign_id, campaign FROM ads_campaign_daily WHERE city = 'Unmapped'")
    cur = con.cursor()
    for r in rows.itertuples(index=False):
        city = _city_from_name(r.campaign, cfg)
        if city != "Unmapped":
            cur.execute("UPDATE IGNORE ads_campaign_daily SET city = %s WHERE campaign_id = %s AND city = 'Unmapped'",
                        (city, r.campaign_id))
            # any leftover rows would double count: the same day already exists under the right city
            cur.execute("DELETE FROM ads_campaign_daily WHERE campaign_id = %s AND city = 'Unmapped'", (r.campaign_id,))
            log(con, "remap", f"{r.campaign} -> {city}")


# =====================================================================
# ANALYSIS
# =====================================================================
def build_calendar(cfg, start, end):
    idx = pd.date_range(start, end, freq="D")
    cal = pd.DataFrame(index=idx)
    cal["dow"] = idx.day_name()
    cal["dom_bucket"] = np.select(
        [idx.day <= 3, idx.day >= idx.days_in_month - 2, idx.day <= 15],
        ["Start (1-3)", "End (last 3)", "Early-mid (4-15)"], default="Late-mid (16-end-3)")
    cal["month"] = idx.strftime("%m-%b")
    cal["event"] = ""

    def add(d, tag):
        d = pd.Timestamp(d)
        if d in cal.index:
            cal.at[d, "event"] = (cal.at[d, "event"] + "; " if cal.at[d, "event"] else "") + tag

    if cfg["add_india_public_holidays"]:
        try:
            import holidays
            for d, name in holidays.India(years=range(idx.min().year, idx.max().year + 1)).items():
                add(d, f"[Holiday] {name}")
        except ImportError:
            pass
    if os.path.exists(cfg["events_csv"]):
        ev = pd.read_csv(cfg["events_csv"])
        for _, r in ev.iterrows():
            if pd.notna(pd.to_datetime(r["date"], errors="coerce")) and "EXAMPLE" not in str(r["event"]):
                add(r["date"], f"[{r.get('type', 'Event')}] {r['event']}")
    return cal


def stl_decompose(series, period=7):
    from statsmodels.tsa.seasonal import STL
    s = series.astype(float).interpolate(limit_direction="both")
    if s.notna().sum() < period * 4 or s.std() == 0:
        return None
    r = STL(s, period=period, robust=True).fit()
    return pd.DataFrame({"trend": r.trend, "seasonal_weekly": r.seasonal, "residual": r.resid})


def robust_z(x):
    med = np.nanmedian(x)
    mad = np.nanmedian(np.abs(x - med)) * 1.4826
    return (x - med) / mad if mad > 0 else x * 0


def city_daily(city, ads, trends, cal, cfg, matured_cutoff):
    a = ads[ads.city == city].set_index("date")[BASE_COLS]
    df = cal[[]].join(a)
    for c in ["impressions", "clicks", "cost"]:
        df[c] = df[c].fillna(0)
    geo = cfg["city_to_geo"].get(city)
    t = trends[trends.geo == geo].set_index("date")["idx"] if geo else pd.Series(dtype=float)
    df["trends_index"] = t.reindex(df.index)
    df["matured"] = df.index <= matured_cutoff
    # conversions only trusted once matured
    df["conv_matured"] = np.where(df.matured, df.conversions, np.nan)
    df["ctr"] = df.clicks / df.impressions.replace(0, np.nan)
    df["cpc"] = df.cost / df.clicks.replace(0, np.nan)
    df["cvr"] = df.conv_matured / df.clicks.replace(0, np.nan)
    df["cpa"] = df.cost.where(df.matured) / df.conv_matured.replace(0, np.nan)
    df["impr_per_trend_pt"] = df.impressions / df.trends_index.replace(0, np.nan)
    df["city"] = city
    return df


def profile(df, by):
    m = df[df.matured]
    g = df.groupby(by).agg(days=("impressions", "size"), impressions=("impressions", "sum"),
                           clicks=("clicks", "sum"), cost=("cost", "sum"),
                           trends_avg=("trends_index", "mean"), impr_share=("impr_share", "mean"))
    gm = m.groupby(by).agg(matured_days=("impressions", "size"), conversions=("conv_matured", "sum"),
                           clicks_m=("clicks", "sum"), cost_m=("cost", "sum"))
    g = g.join(gm)
    g["impr_index"] = (g.impressions / g.days) / (g.impressions / g.days).mean() * 100
    g["trends_index_rel"] = g.trends_avg / g.trends_avg.mean() * 100
    g["conv_per_day_index"] = (g.conversions / g.matured_days) / (g.conversions / g.matured_days).mean() * 100
    g["ctr"] = g.clicks / g.impressions.replace(0, np.nan)
    g["cpc"] = g.cost / g.clicks.replace(0, np.nan)
    g["cvr"] = g.conversions / g.clicks_m.replace(0, np.nan)
    g["cpa"] = g.cost_m / g.conversions.replace(0, np.nan)
    return g.drop(columns=["clicks_m", "cost_m"]).reset_index()


def find_spikes(df, dec, col, cal, cfg):
    if dec is None:
        return pd.DataFrame()
    z = robust_z(dec["residual"].values)
    spend7 = df["cost"].rolling(7, min_periods=3).mean()
    prev7 = spend7.shift(7)
    w = cfg["event_window_days"]
    rows = []
    for i in np.where(np.abs(z) >= cfg["spike_z"])[0]:
        d = dec.index[i]
        win = cal.loc[d - timedelta(days=w): d + timedelta(days=w)]
        ev = "; ".join(f"{k:%d-%b}: {e}" for k, e in win["event"].items() if e)
        sp = (spend7[d] / prev7[d] - 1) * 100 if pd.notna(prev7.get(d)) and prev7[d] > 0 else np.nan
        rows.append({"date": d.date(), "series": col, "direction": "Spike" if z[i] > 0 else "Dip",
                     "robust_z": round(float(z[i]), 2), "actual": df.at[d, col],
                     "dow": cal.at[d, "dow"], "dom_bucket": cal.at[d, "dom_bucket"],
                     "events_within_window": ev,
                     "spend_7d_change_pct": round(sp, 1) if pd.notna(sp) else None,
                     "possible_internal_change": bool(pd.notna(sp) and abs(sp) > cfg["spend_change_pct"])})
    return pd.DataFrame(rows)


def weekly(df):
    w = df.resample("W-MON", label="left", closed="left").agg(
        impressions=("impressions", "sum"), clicks=("clicks", "sum"), cost=("cost", "sum"),
        conversions=("conv_matured", "sum"), matured_days=("matured", "sum"),
        trends_index=("trends_index", "mean"), impr_share=("impr_share", "mean"),
        days=("impressions", "size"))
    w = w[w.days == 7].drop(columns="days")   # drop part-weeks at the start/end so % changes are fair
    w["trends_wow_pct"] = w.trends_index.pct_change() * 100
    w["impr_wow_pct"] = w.impressions.pct_change() * 100
    w["gap_flag"] = np.select(
        [(w.trends_wow_pct > 10) & (w.impr_wow_pct < 0), (w.trends_wow_pct < -10) & (w.impr_wow_pct > 10)],
        ["Demand up, impressions down (missed demand?)", "Impressions up without demand (internal change?)"],
        default="")
    w["conv_status"] = np.where(w.matured_days == 7, "Matured", "Maturing - ignore conv")
    w["cpa"] = np.where(w.matured_days == 7, w.cost / w.conversions.replace(0, np.nan), np.nan)
    w["cvr"] = np.where(w.matured_days == 7, w.conversions / w.clicks.replace(0, np.nan), np.nan)
    return w.reset_index().rename(columns={"date": "week_start", "index": "week_start"})


def build_sheets(cfg):
    con = db_connect(cfg)
    ads, trends = read_ads(con, cfg), read_trends(con)
    if ads.empty:
        con.close()
        return None
    start, end = ads.date.min(), ads.date.max()
    matured_cutoff = pd.Timestamp(date.today()) - timedelta(days=cfg["maturity_days"] + 1)
    cal = build_calendar(cfg, start, end)

    S = {k: [] for k in ["summary", "weekly", "dow", "dom", "month", "spikes", "lag", "stl", "daily"]}
    for city in sorted(ads.city.unique(), key=lambda c: (c != "All India", c)):
        df = city_daily(city, ads, trends, cal, cfg, matured_cutoff)
        S["daily"].append(df.reset_index().rename(columns={"index": "date"}))

        d_imp = stl_decompose(df["impressions"])
        d_tr = stl_decompose(df["trends_index"]) if df.trends_index.notna().sum() > 28 else None
        d_cv = stl_decompose(df.loc[df.matured, "conv_matured"])
        if d_imp is not None:
            x = d_imp.add_prefix("impr_")
            if d_tr is not None:
                x = x.join(d_tr.add_prefix("trends_"))
            x.insert(0, "city", city)
            S["stl"].append(x.reset_index().rename(columns={"index": "date"}))

        for dec, col in ((d_imp, "impressions"), (d_tr, "trends_index"), (d_cv, "conv_matured")):
            sp = find_spikes(df, dec, col, cal, cfg)
            if len(sp):
                sp.insert(0, "city", city)
                S["spikes"].append(sp)

        best_lag = best_corr = None
        if df.trends_index.notna().sum() > 28:
            lc = pd.DataFrame([{"lag_days": L,
                                "corr_impr_vs_trends": df.impressions.corr(df.trends_index.shift(L)),
                                "corr_conv_vs_trends": df.conv_matured.corr(df.trends_index.shift(L))}
                               for L in range(-14, 15)])
            if lc.corr_impr_vs_trends.notna().any():
                b = lc.loc[lc.corr_impr_vs_trends.idxmax()]
                best_lag, best_corr = int(b.lag_days), round(float(b.corr_impr_vs_trends), 3)
            lc.insert(0, "city", city)
            S["lag"].append(lc)

        for key, by in (("dow", "dow"), ("dom", "dom_bucket"), ("month", "month")):
            p = profile(df.join(cal[[by]]), by)
            p.insert(0, "city", city)
            S[key].append(p)

        wk = weekly(df)
        wk.insert(0, "city", city)
        S["weekly"].append(wk)

        dw = S["dow"][-1].sort_values("impr_index")
        S["summary"].append({
            "city": city, "trends_geo": cfg["city_to_geo"].get(city),
            "data_from": start.date(), "data_to": end.date(), "conv_matured_until": matured_cutoff.date(),
            "best_lag_days (impr follow trends)": best_lag, "corr_at_best_lag": best_corr,
            "weekly_pattern_strength_impr": round(float(1 - d_imp.residual.var() /
                                                        (d_imp.residual + d_imp.seasonal_weekly).var()), 3)
            if d_imp is not None else None,
            "weakest_dow_impr": dw.iloc[0]["dow"], "strongest_dow_impr": dw.iloc[-1]["dow"],
            "missed_demand_weeks": int(wk.gap_flag.str.startswith("Demand up").sum()),
        })

    cat = lambda L: pd.concat(L, ignore_index=True) if L else pd.DataFrame()
    notes = pd.DataFrame({"Notes": [
        f"Google Ads API, account {cfg['customer_id']}, campaigns matching '{cfg['campaign_name_regex']}'. Data {start.date()} to {end.date()}.",
        f"All cost, CPC and CPA include GST (Google Ads cost x {cfg['gst_multiplier']}).",
        "New days: Enabled campaigns only. Campaigns paused later keep their stored history; their last days' conversions keep maturing.",
        f"Conversions are refreshed daily for the last {cfg['maturity_days']} days. "
        f"Only days up to {matured_cutoff.date()} count as matured; conv metrics (CVR, CPA, conv spikes) use matured days only.",
        "Impressions are the main demand-response metric (available immediately, no lag).",
        "Trends: state-level 0-100 index, sampled. History stitched once, then each daily pull is rescaled onto it "
        f"via overlap ratio; last {cfg['trends_recent_overwrite']} days overwritten each run.",
        "STL weekly decomposition (period 7). Yearly patterns need ~2 years of data to trust.",
        f"Spikes: |robust z| of STL residual >= {cfg['spike_z']}; events tagged within ±{cfg['event_window_days']} days.",
        f"possible_internal_change = 7-day avg spend moved > {cfg['spend_change_pct']}%: check before calling it seasonal.",
        "Correlation is not proof of cause; act on patterns that repeat across weeks, years or cities.",
    ]})
    sheets = {"Summary": pd.DataFrame(S["summary"]), "Notes": notes, "Weekly": cat(S["weekly"]),
              "DOW_Profile": cat(S["dow"]), "DayOfMonth_Profile": cat(S["dom"]),
              "Month_Profile": cat(S["month"]), "Spikes_Tagged": cat(S["spikes"]),
              "Lag_Correlation": cat(S["lag"]), "STL_Components": cat(S["stl"]),
              "Daily_Aligned": cat(S["daily"]),
              "Run_Log": q(con, "SELECT run_at, step, detail FROM run_log ORDER BY id DESC LIMIT 200")}
    try:
        sheets.update(build_planning(cfg, con, ads, trends))
    except Exception as e:
        print("planning analysis:", e)
    con.close()
    return sheets


def write_xlsx(cfg, sheets):
    with pd.ExcelWriter(cfg["output_xlsx"], engine="xlsxwriter") as xw:
        wb = xw.book
        hdr = wb.add_format({"bold": True, "font_name": "Arial", "font_size": 10,
                             "bg_color": "#DCE6F1", "border": 1})
        for name, df in sheets.items():
            if df.empty:
                continue
            df = df.copy()
            for c in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[c]):
                    df[c] = df[c].dt.date
            df = df.rename(columns={"cost": "cost_incl_gst", "cpc": "cpc_incl_gst", "cpa": "cpa_incl_gst"})
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            for i, col in enumerate(df.columns):
                ws.write(0, i, col, hdr)
                ws.set_column(i, i, max(12, min(45, len(str(col)) + 2)))
            ws.freeze_panes(1, 0)
        wk = sheets["Weekly"]
        for city in wk.city.unique():
            cw = wk[wk.city == city][["week_start", "impressions", "trends_index", "conversions",
                                      "conv_status", "cpa", "gap_flag"]].reset_index(drop=True)
            cw = cw.rename(columns={"cpa": "cpa_incl_gst"})
            cw["week_start"] = pd.to_datetime(cw.week_start).dt.date
            name = f"Chart_{city}"[:31]
            cw.to_excel(xw, sheet_name=name, index=False)
            n = len(cw)
            ch = wb.add_chart({"type": "line"})
            ch.add_series({"name": "Impressions", "categories": [name, 1, 0, n, 0], "values": [name, 1, 1, n, 1]})
            ch.add_series({"name": "Google Trends", "categories": [name, 1, 0, n, 0],
                           "values": [name, 1, 2, n, 2], "y2_axis": True})
            ch.set_title({"name": f"{city}: weekly impressions vs Google Trends"})
            ch.set_size({"width": 960, "height": 380})
            xw.sheets[name].insert_chart(1, 8, ch)
    print(f"Analysis written -> {cfg['output_xlsx']}")


def run_analysis(cfg):
    sheets = build_sheets(cfg)
    if sheets is None:
        print("No data stored yet.")
        return None
    write_xlsx(cfg, sheets)
    return sheets


# =====================================================================
# LOCAL DASHBOARD SERVER  (http://localhost:<port>)
# =====================================================================
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

STATE = {"busy": False, "job": "", "error": "", "sheets": None, "lock": threading.Lock()}


def _records(df):
    if df is None or len(df) == 0:
        return []
    return json.loads(df.to_json(orient="records", date_format="iso"))


def _last_ok(cfg):
    con = db_connect(cfg)
    r = q(con, "SELECT run_at, detail FROM run_log WHERE step = 'refresh-ok' ORDER BY id DESC LIMIT 1")
    errs = q(con, "SELECT run_at, step, detail FROM run_log WHERE step LIKE '%error%' OR step = 'warn' "
                  "ORDER BY id DESC LIMIT 5")
    con.close()
    return (r.iloc[0].to_dict() if len(r) else None), _records(errs)


def recompute(cfg):
    STATE["sheets"] = build_sheets(cfg)
    STATE["at"] = time.time()


def do_refresh(cfg, backfill=False):
    if not STATE["lock"].acquire(blocking=False):
        return False
    STATE.update(busy=True, job="Loading full history (first run, can take a while)" if backfill
                 else "Refreshing data", error="")

    def work():
        try:
            run_refresh(cfg, backfill=backfill)
            STATE["job"] = "Updating analysis"
            recompute(cfg)
        except Exception as e:
            STATE["error"] = str(e)
            try:
                con = db_connect(cfg); log(con, "refresh-error", str(e)[:500]); con.close()
            except Exception:
                pass
        finally:
            STATE.update(busy=False, job="")
            STATE["lock"].release()

    threading.Thread(target=work, daemon=True).start()
    return True


def scheduler(cfg):
    """While the dashboard runs: refresh once a day after refresh_time (and catch up if missed)."""
    hh, mm = map(int, cfg["refresh_time"].split(":"))
    while True:
        try:
            now = datetime.now()
            last, _ = _last_ok(cfg)
            done_today = last and last["run_at"][:10] == now.strftime("%Y-%m-%d")
            if not done_today and (now.hour, now.minute) >= (hh, mm) and not STATE["busy"]:
                do_refresh(cfg)
        except Exception as e:
            print("scheduler:", e)
        time.sleep(60)


def api_payload(cfg):
    # fetch fresh from TiDB every time the page is opened (at most once a minute)
    try:
        if not STATE["busy"] and time.time() - STATE.get("at", 0) > 60:
            recompute(cfg)
        last, issues = _last_ok(cfg)
    except Exception as e:
        return {"status": {"busy": STATE["busy"], "job": STATE["job"],
                           "error": f"Can't reach the TiDB database: {e}", "last_refresh": None, "issues": [],
                           "refresh_time": cfg["refresh_time"], "account": cfg["customer_id"],
                           "filter": cfg["campaign_name_regex"], "gst": cfg["gst_multiplier"]}, "empty": True}
    sh = STATE["sheets"]
    out = {"status": {"busy": STATE["busy"], "job": STATE["job"], "error": STATE["error"],
                      "last_refresh": last, "issues": issues, "refresh_time": cfg["refresh_time"],
                      "account": cfg["customer_id"], "filter": cfg["campaign_name_regex"],
                      "gst": cfg["gst_multiplier"]}}
    if sh is None:
        out["empty"] = True
        return out
    keep = {"Summary": None, "Weekly": None, "DOW_Profile": None, "DayOfMonth_Profile": None,
            "Month_Profile": None, "Spikes_Tagged": None, "Notes": None}
    for k in keep:
        out[k] = _records(sh.get(k))
    return out


class Handler(BaseHTTPRequestHandler):
    cfg = None

    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype="application/json"):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, default=str).encode()
        elif isinstance(body, str):
            body = body.encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            here = os.path.dirname(os.path.abspath(__file__))
            with open(os.path.join(here, "dashboard.html"), encoding="utf-8") as f:
                return self._send(200, f.read(), "text/html; charset=utf-8")
        if self.path == "/api/data":
            return self._send(200, api_payload(self.cfg))
        if self.path == "/api/status":
            try:
                last, issues = _last_ok(self.cfg)
            except Exception:
                last = None
            return self._send(200, {"busy": STATE["busy"], "job": STATE["job"], "error": STATE["error"],
                                    "last_refresh": last})
        if self.path == "/download":
            if STATE["sheets"] is None:
                return self._send(404, {"error": "no data yet"})
            write_xlsx(self.cfg, STATE["sheets"])
            with open(self.cfg["output_xlsx"], "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f'attachment; filename="{self.cfg["output_xlsx"]}"')
            self.end_headers()
            self.wfile.write(data)
            return
        self._send(404, {"error": "not found"})

    def do_POST(self):
        if self.path == "/api/refresh":
            started = do_refresh(self.cfg)
            return self._send(200, {"started": started})
        self._send(404, {"error": "not found"})


def serve(cfg):
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    Handler.cfg = cfg
    try:
        recompute(cfg)
    except Exception as e:
        print("analysis:", e)
    try:
        con = db_connect(cfg)
        empty = int(q(con, "SELECT COUNT(*) AS n FROM ads_campaign_daily")["n"].iloc[0]) == 0
        con.close()
        if empty:
            do_refresh(cfg, backfill=True)
    except Exception as e:
        STATE["error"] = f"Can't reach the TiDB database: {e}"
        print(STATE["error"])
    threading.Thread(target=scheduler, args=(cfg,), daemon=True).start()
    url = f"http://localhost:{cfg['port']}"
    srv = ThreadingHTTPServer(("127.0.0.1", cfg["port"]), Handler)
    print(f"\nDashboard running at {url}\nKeep this window open. Auto-refresh daily after {cfg['refresh_time']}. "
          "Press Ctrl+C to stop.\n")
    try:
        webbrowser.open(url)
    except Exception:
        pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("Stopped.")


# =====================================================================
# DEMAND PLANNING: 12-month forecast, locations, brand vs generic volumes
# =====================================================================
def _job_due(con, job, days):
    r = q(con, "SELECT last_ok FROM job_runs WHERE job = %s", (job,))
    if r.empty or not r.last_ok.iloc[0]:
        return True
    return datetime.now() - datetime.fromisoformat(r.last_ok.iloc[0]) > timedelta(days=days)


def _job_done(con, job, note=""):
    con.cursor().execute(
        "INSERT INTO job_runs (job, last_ok, note) VALUES (%s,%s,%s) "
        "ON DUPLICATE KEY UPDATE last_ok = VALUES(last_ok), note = VALUES(note)",
        (job, datetime.now().isoformat(timespec="seconds"), note))


def _trends_client():
    from pytrends.request import TrendReq
    return TrendReq(hl="en-IN", tz=-330)


def _with_retry(fn, label, tries=3):
    for attempt in range(tries):
        try:
            return fn()
        except Exception as e:
            if attempt == tries - 1:
                raise
            wait = 40 * (attempt + 1)
            print(f"   {label} failed ({e}); retry in {wait}s")
            time.sleep(wait)


def pull_trends_5y(geo, cfg):
    """5 years of weekly Google Trends for the forecast keyword (one request, one consistent scale)."""
    def go():
        py = _trends_client()
        py.build_payload([cfg["forecast_keyword"]], timeframe="today 5-y", geo=geo)
        d = py.interest_over_time()
        if "isPartial" in d.columns:
            d = d[~d["isPartial"].astype(bool)].drop(columns="isPartial")
        return d.iloc[:, 0].astype(float)
    return _with_retry(go, f"Trends 5y {geo}")


def pull_trends_cities(cfg):
    """Relative search interest by Indian city over the last 12 months (100 = highest city)."""
    def go():
        py = _trends_client()
        py.build_payload([cfg["forecast_keyword"]], timeframe="today 12-m", geo="IN")
        d = py.interest_by_region(resolution="CITY", inc_low_vol=True)
        return d.iloc[:, 0].astype(float)
    return _with_retry(go, "Trends cities")


def _kp_geo_ids(api, cfg, cities):
    """Our city names -> Google geo target resource names (cities in India; All India = country)."""
    gname = {c: cfg["google_city_names"].get(c, c) for c in cities if c != "All India"}
    names = sorted(set(gname.values()))
    out = {"All India": "geoTargetConstants/2356"}
    if names:
        quoted = ",".join("'" + n.replace("'", "") + "'" for n in names)
        qy = ("SELECT geo_target_constant.resource_name, geo_target_constant.name, geo_target_constant.target_type "
              "FROM geo_target_constant WHERE geo_target_constant.country_code = 'IN' "
              f"AND geo_target_constant.status = 'ENABLED' AND geo_target_constant.name IN ({quoted})")
        found = {}
        for row in api.stream(qy):
            g = row.get("geoTargetConstant", {})
            if g.get("targetType") == "City" or g.get("name") not in found:
                found[g.get("name")] = g.get("resourceName")
        for c, n in gname.items():
            if found.get(n):
                out[c] = found[n]
    return out


MONTHS = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST",
          "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]


def pull_kp(api, cfg, geo_res):
    """Keyword Planner monthly searches + top-of-page bid range for every configured keyword."""
    kws = [(k, g) for g, lst in cfg["kp_keywords"].items() for k in lst]
    v = api._pick_version()
    url = f"https://googleads.googleapis.com/{v}/customers/{cfg['customer_id']}:generateKeywordHistoricalMetrics"
    today = date.today()
    end_m = (today.replace(day=1) - timedelta(days=1))
    start_m = end_m.replace(year=end_m.year - 3, day=1)
    body = {"keywords": [k for k, _ in kws], "geoTargetConstants": [geo_res],
            "language": "languageConstants/1000", "keywordPlanNetwork": "GOOGLE_SEARCH",
            "historicalMetricsOptions": {"yearMonthRange": {
                "start": {"year": start_m.year, "month": MONTHS[start_m.month - 1]},
                "end": {"year": end_m.year, "month": MONTHS[end_m.month - 1]}}}}
    r = requests.post(url, headers=api._headers(), json=body, timeout=120)
    if r.status_code == 400 and "yearMonthRange" in r.text:      # older range not allowed: use default 12 months
        body.pop("historicalMetricsOptions")
        r = requests.post(url, headers=api._headers(), json=body, timeout=120)
    if r.status_code != 200:
        raise RuntimeError(f"Keyword Planner ({r.status_code}): {_ads_err(r)}")
    grp = {k.lower(): g for k, g in kws}
    rows = []
    for res in r.json().get("results", []):
        kw = res.get("text", "")
        m = res.get("keywordMetrics", {})
        lo = float(m.get("lowTopOfPageBidMicros", 0) or 0) / 1e6
        hi = float(m.get("highTopOfPageBidMicros", 0) or 0) / 1e6
        for mv in m.get("monthlySearchVolumes", []):
            mi = MONTHS.index(mv["month"]) + 1 if mv.get("month") in MONTHS else int(mv.get("month", 0))
            rows.append((kw, grp.get(kw.lower(), "Other"), f"{int(mv['year'])}-{mi:02d}",
                         float(mv.get("monthlySearches", 0) or 0), lo, hi))
    return rows


def run_planning_jobs(cfg):
    """Slow-changing pulls, spread over days so Google Trends doesn't rate-limit us."""
    con = db_connect(cfg)
    cities = [c for c in q(con, "SELECT DISTINCT city FROM ads_campaign_daily")["city"] if c != "Unmapped"]
    cities = ["All India"] + sorted(cities)
    geos = ["IN"] + sorted({cfg["city_to_geo"][c] for c in cities if c in cfg["city_to_geo"] and c != "All India"})
    now = datetime.now().isoformat(timespec="seconds")

    # 1) 5-year weekly Trends per state -> 12-month forecast
    budget = cfg["forecast_states_per_run"]
    for geo in geos:
        if budget <= 0:
            break
        if not _job_due(con, f"t5y:{geo}", cfg["forecast_refresh_days"]):
            continue
        budget -= 1
        try:
            s5 = pull_trends_5y(geo, cfg)
            con = db_connect(cfg)
            cur = con.cursor()
            cur.execute("DELETE FROM trends_weekly_5y WHERE geo = %s", (geo,))
            cur.executemany("INSERT INTO trends_weekly_5y (geo, week, idx, fetched_at) VALUES (%s,%s,%s,%s)",
                            [(geo, d.strftime("%Y-%m-%d"), float(v), now) for d, v in s5.items()])
            _job_done(con, f"t5y:{geo}", f"{len(s5)} weeks")
            log(con, "forecast-data", f"{geo}: {len(s5)} weeks")
        except Exception as e:
            con = db_connect(cfg); log(con, "trends-error", f"5y {geo}: {str(e)[:300]}")
        time.sleep(15)

    # 2) Trends interest by city across India
    if _job_due(con, "trends-cities", cfg["forecast_refresh_days"]):
        try:
            sc = pull_trends_cities(cfg)
            con = db_connect(cfg)
            cur = con.cursor()
            cur.execute("DELETE FROM trends_city_interest")
            cur.executemany("INSERT INTO trends_city_interest (city, idx, fetched_at) VALUES (%s,%s,%s)",
                            [(str(k)[:64], float(v), now) for k, v in sc.items()])
            _job_done(con, "trends-cities", f"{len(sc)} cities")
            log(con, "locations", f"{len(sc)} cities")
        except Exception as e:
            con = db_connect(cfg); log(con, "trends-error", f"cities: {str(e)[:300]}")

    # 3) Keyword Planner monthly volumes per city (Google Ads API, not rate-limited like Trends)
    try:
        api = _API.get("api") or AdsREST(cfg)
        _API["api"] = api
        due = [c for c in cities if _job_due(con, f"kp:{c}", cfg["kp_refresh_days"])]
        if due:
            ids = _kp_geo_ids(api, cfg, due)
            for c in due:
                if c not in ids:
                    log(con, "warn", f"Keyword Planner: no Google location found for {c}")
                    continue
                try:
                    rows = pull_kp(api, cfg, ids[c])
                    con = db_connect(cfg)
                    con.cursor().executemany(
                        "INSERT INTO kp_monthly (city, keyword, grp, ym, searches, low_bid, high_bid, fetched_at) "
                        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s) ON DUPLICATE KEY UPDATE grp = VALUES(grp), "
                        "searches = VALUES(searches), low_bid = VALUES(low_bid), high_bid = VALUES(high_bid), "
                        "fetched_at = VALUES(fetched_at)",
                        [(c, *r, now) for r in rows])
                    _job_done(con, f"kp:{c}", f"{len(rows)} rows")
                except Exception as e:
                    log(con, "kp-error", f"{c}: {str(e)[:300]}")
                time.sleep(2)
            log(con, "volumes", f"Keyword Planner updated for {len(due)} locations")
    except Exception as e:
        con = db_connect(cfg); log(con, "kp-error", str(e)[:300])
    con.close()


def _suggest(pct, cfg):
    t = cfg["forecast_threshold_pct"]
    if pct >= t:
        return f"Demand high: raise budget ~{min(40, round(pct / 5) * 5)}%, allow higher bids/tCPA"
    if pct <= -t:
        return f"Demand low: trim budget ~{min(40, round(-pct / 5) * 5)}%, hold bids"
    return "Normal: keep steady"


def forecast_weekly(s5, cfg, horizon=52):
    """Seasonal forecast from 5 years of weekly Trends.
    Season = median of (week / its 52-week average) for each week of the year across years.
    Level = last 52 weeks' average, with a damped year-on-year trend.
    Also back-tests itself on the most recent 52 weeks to report a typical error."""
    s = s5.sort_index().astype(float)
    s = s.replace(0, np.nan).interpolate(limit_direction="both")
    if s.notna().sum() < 104:
        return None, None

    def profile_of(x):
        base = x.rolling(52, center=True, min_periods=26).mean()
        ratio = (x / base).dropna()
        wk = ratio.index.isocalendar().week.clip(upper=52).astype(int)
        g = ratio.groupby(wk.values)
        med, lo, hi = g.median(), g.quantile(0.25), g.quantile(0.75)
        idx = pd.Index(range(1, 53))
        sm = lambda v: pd.concat([v.reindex(idx).iloc[-2:], v.reindex(idx), v.reindex(idx).iloc[:2]]) \
            .interpolate().rolling(3, center=True).mean().iloc[2:-2]
        return sm(med), sm(lo), sm(hi), g.size().reindex(idx).fillna(0)

    def project(x, start, n):
        med, lo, hi, cnt = profile_of(x)
        last = x.iloc[-52:].mean()
        prev = x.iloc[-104:-52].mean() if len(x) >= 104 else last
        yoy = float(np.clip((last / prev) ** 0.5 if prev else 1.0, 0.8, 1.25))
        rows = []
        for h in range(n):
            d = start + timedelta(days=7 * h)
            w = min(int(d.isocalendar()[1]), 52)
            lvl = last * yoy ** ((h + 1) / 52)
            rows.append({"week_start": d, "week_of_year": w, "expected_index": lvl * med[w],
                         "vs_normal_pct": (med[w] - 1) * 100, "low_pct": (lo[w] - 1) * 100,
                         "high_pct": (hi[w] - 1) * 100, "years_of_history": int(cnt[w])})
        return pd.DataFrame(rows)

    # back-test: pretend it's 52 weeks ago
    err = None
    if s.notna().sum() >= 156:
        train, test = s.iloc[:-52], s.iloc[-52:]
        bt = project(train, test.index[0], 52)
        err = float(np.nanmedian(np.abs(bt.expected_index.values - test.values) / test.values) * 100)
    fc = project(s, s.index[-1] + timedelta(days=7), horizon)
    fc["suggestion"] = fc.vs_normal_pct.map(lambda p: _suggest(p, cfg))
    return fc, err


def daily_factors(tr_geo):
    """Day-of-week and time-of-month multipliers from daily Trends history."""
    if tr_geo is None or len(tr_geo) < 60:
        return None, None
    t = tr_geo.dropna()
    dow = t.groupby(t.index.dayofweek).mean() / t.mean()
    b = np.select([t.index.day <= 3, t.index.day >= t.index.days_in_month - 2, t.index.day <= 15],
                  ["Start", "End", "Early-mid"], default="Late-mid")
    dom = t.groupby(b).mean() / t.mean()
    return dow, dom


def build_planning(cfg, con, ads, trends):
    """Sheets: Forecast_Weekly, Forecast_Daily, Upcoming_Changes, Locations, KP_Monthly, KP_Keywords."""
    out = {}
    w5 = q(con, "SELECT geo, week, idx FROM trends_weekly_5y")
    today = pd.Timestamp(date.today())
    future_cal = build_calendar(cfg, today, today + timedelta(days=400))
    cities = ["All India"] + sorted(c for c in ads.city.unique() if c != "All India")

    # month-level CPC pressure from our own history
    a = ads.copy()
    a["m"] = a.date.dt.month
    cpc_m = {}
    for c, g in a.groupby("city"):
        mm = g.groupby("m").agg(cost=("cost", "sum"), clicks=("clicks", "sum"))
        cpc = mm.cost / mm.clicks.replace(0, np.nan)
        cpc_m[c] = ((cpc / cpc.mean() - 1) * 100).round(1).to_dict()

    fw, fd, up, acc = [], [], [], []
    for city in cities:
        geo = cfg["city_to_geo"].get(city)
        g = w5[w5.geo == geo] if geo else w5.iloc[0:0]
        if g.empty:
            continue
        s5 = pd.Series(pd.to_numeric(g.idx).values, index=pd.to_datetime(g.week))
        fc, err = forecast_weekly(s5, cfg)
        if fc is None:
            continue
        fc.insert(0, "city", city)
        fc["week_end"] = fc.week_start + timedelta(days=6)
        fc["events"] = [", ".join(e for e in future_cal.loc[r.week_start:r.week_end, "event"] if e)
                        for r in fc.itertuples()]
        fc["past_cpc_this_month_pct"] = [cpc_m.get(city, {}).get(d.month) for d in fc.week_start]
        fw.append(fc)
        acc.append({"city": city, "trends_geo": geo, "backtest_typical_error_pct": None if err is None else round(err, 1),
                    "history_weeks": int(len(s5))})

        # notable upcoming weeks: runs of high/low weeks, merged
        t = cfg["forecast_threshold_pct"]
        flag = np.where(fc.vs_normal_pct >= t, "High", np.where(fc.vs_normal_pct <= -t, "Low", ""))
        run = None
        for i, f in enumerate(flag):
            if f and run and run["type"] == f and i == run["end_i"] + 1:
                run["end_i"] = i
            else:
                if run:
                    up.append(run)
                run = {"city": city, "type": f, "start_i": i, "end_i": i} if f else None
        if run:
            up.append(run)
        for u in [u for u in up if u["city"] == city and "from" not in u]:
            seg = fc.iloc[u["start_i"]:u["end_i"] + 1]
            peak = seg.vs_normal_pct.abs().idxmax()
            u.update({"from": seg.week_start.iloc[0], "to": seg.week_end.iloc[-1],
                      "demand_vs_normal_pct": round(float(seg.vs_normal_pct.loc[peak]), 1),
                      "suggestion": fc.suggestion.loc[peak],
                      "past_cpc_this_month_pct": fc.past_cpc_this_month_pct.loc[peak],
                      "events": ", ".join(sorted({e for e in seg.events if e}))[:200]})

        # next 90 days, day by day
        trg = trends[trends.geo == geo].set_index("date")["idx"] if geo else None
        dow, dom = daily_factors(trg)
        wk = fc.set_index("week_start")
        for d in pd.date_range(today, today + timedelta(days=89)):
            ws = wk.index[wk.index <= d]
            wsel = ws[-1] if len(ws) else wk.index[0]        # days before the first forecast week use it too
            base = 1 + wk.loc[wsel, "vs_normal_pct"] / 100
            f_dow = dow.get(d.dayofweek, 1.0) if dow is not None else 1.0
            bucket = ("Start" if d.day <= 3 else "End" if d.day >= d.days_in_month - 2
                      else "Early-mid" if d.day <= 15 else "Late-mid")
            f_dom = dom.get(bucket, 1.0) if dom is not None else 1.0
            fd.append({"city": city, "date": d, "dow": d.day_name()[:3],
                       "day_vs_normal_pct": round((base * f_dow * f_dom - 1) * 100, 1),
                       "event": future_cal.loc[d, "event"] if d in future_cal.index else ""})

    cat = lambda L: pd.concat(L, ignore_index=True) if L else pd.DataFrame()
    out["Forecast_Weekly"] = cat(fw)
    out["Forecast_Daily"] = pd.DataFrame(fd)
    out["Upcoming_Changes"] = pd.DataFrame([{k: v for k, v in u.items() if k not in ("start_i", "end_i")}
                                            for u in up if "from" in u])
    out["Forecast_Accuracy"] = pd.DataFrame(acc)

    # Locations: Trends city interest + Keyword Planner volume + our coverage
    ci = q(con, "SELECT city, idx FROM trends_city_interest")
    kp = q(con, "SELECT city, keyword, grp, ym, searches, low_bid, high_bid FROM kp_monthly")
    if not kp.empty:
        kp["searches"] = pd.to_numeric(kp.searches)
        kp["low_bid"] = pd.to_numeric(kp.low_bid) * cfg["gst_multiplier"]
        kp["high_bid"] = pd.to_numeric(kp.high_bid) * cfg["gst_multiplier"]
    last30 = ads[ads.date > ads.date.max() - timedelta(days=30)].groupby("city").agg(
        impressions_30d=("impressions", "sum"), cost_30d=("cost", "sum"), clicks_30d=("clicks", "sum"))
    rev = {v.lower(): k for k, v in cfg["google_city_names"].items()}
    rows = []
    names = set(cities) | set(ci.city if not ci.empty else [])
    names = {rev.get(str(c).lower(), c) for c in names}          # one row per city, in our naming
    for c in sorted(names):
        ours = c
        tr = ci[ci.city.str.lower().isin([str(c).lower(), cfg["google_city_names"].get(ours, ours).lower()])] \
            if not ci.empty else ci
        k = kp[(kp.city == ours) & (kp.keyword.str.lower() == cfg["forecast_keyword"])] if not kp.empty else kp
        gen = kp[(kp.city == ours) & (kp.grp == "Generic")] if not kp.empty else kp
        vol = float(k.sort_values("ym").searches.tail(12).mean()) if len(k) else None
        gvol = float(gen.groupby("ym").searches.sum().sort_index().tail(12).mean()) if len(gen) else None
        rows.append({"city": ours, "running_campaigns": ours in last30.index,
                     "trends_interest_0_100": float(tr.idx.iloc[0]) if len(tr) else None,
                     "monthly_searches_main_kw": vol, "monthly_searches_generic_total": gvol,
                     "top_of_page_bid_low": float(k.low_bid.iloc[-1]) if len(k) else None,
                     "top_of_page_bid_high": float(k.high_bid.iloc[-1]) if len(k) else None,
                     "impressions_30d": float(last30.impressions_30d.get(ours, 0)) if ours in last30.index else 0,
                     "cost_30d_incl_gst": float(last30.cost_30d.get(ours, 0)) if ours in last30.index else 0})
    loc = pd.DataFrame(rows)
    if len(loc):
        loc = loc[(loc.city != "All India") | loc.monthly_searches_main_kw.notna()]
        loc["impr_per_generic_search"] = loc.impressions_30d / loc.monthly_searches_generic_total
        loc = loc.sort_values(["trends_interest_0_100", "monthly_searches_main_kw"],
                              ascending=False, na_position="last")
    out["Locations"] = loc
    if not kp.empty:
        out["KP_Monthly"] = kp.groupby(["city", "grp", "ym"], as_index=False).searches.sum()
        last12 = kp[kp.ym >= sorted(kp.ym.unique())[-12]] if kp.ym.nunique() >= 12 else kp
        kk = last12.groupby(["city", "grp", "keyword"], as_index=False).agg(
            avg_monthly_searches=("searches", "mean"), top_of_page_bid_low=("low_bid", "last"),
            top_of_page_bid_high=("high_bid", "last"))
        out["KP_Keywords"] = kk.sort_values(["city", "avg_monthly_searches"], ascending=[True, False])
    return out


# =====================================================================
# CLOUD MODE (GitHub Actions): refresh, then save dashboard data + Excel into TiDB
# =====================================================================
def save_cache(cfg, sheets):
    payload = {"meta": {"account": cfg["customer_id"], "filter": cfg["campaign_name_regex"],
                        "gst": cfg["gst_multiplier"], "refresh_time": "09:00 IST"}}
    if sheets is None:
        payload["empty"] = True
        xbytes = None
    else:
        for k in ["Summary", "Weekly", "DOW_Profile", "DayOfMonth_Profile", "Month_Profile",
                  "Spikes_Tagged", "Notes", "Forecast_Weekly", "Forecast_Daily", "Upcoming_Changes",
                  "Forecast_Accuracy", "Locations", "KP_Monthly", "KP_Keywords"]:
            payload[k] = _records(sheets.get(k))
        write_xlsx(cfg, sheets)
        xbytes = open(cfg["output_xlsx"], "rb").read()
        if len(xbytes) > 5_000_000:
            print("Excel too large to store in TiDB; skipping the Excel copy.")
            xbytes = None
    con = db_connect(cfg)
    con.cursor().execute(
        "INSERT INTO dashboard_cache (id, payload, xlsx, updated_at) VALUES ('main', %s, %s, %s) "
        "ON DUPLICATE KEY UPDATE payload = VALUES(payload), xlsx = VALUES(xlsx), updated_at = VALUES(updated_at)",
        (json.dumps(payload, default=str), xbytes, datetime.now().isoformat(timespec="seconds")))
    con.close()
    print("Dashboard data saved to TiDB.")


def cloud_run(cfg, backfill=False):
    failed = None
    try:
        run_refresh(cfg, backfill=backfill)
    except Exception as e:
        failed = e
        try:
            con = db_connect(cfg); log(con, "refresh-error", str(e)[:500]); con.close()
        except Exception:
            pass
    try:
        run_planning_jobs(cfg)
    except Exception as e:
        print("planning jobs:", e)
    save_cache(cfg, build_sheets(cfg))     # always rebuild from whatever is stored
    if failed:
        raise SystemExit(f"Refresh failed: {failed}")


# =====================================================================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--test-api", action="store_true", help="fetch 3 days and print them")
    ap.add_argument("--once", action="store_true", help="refresh + Excel, no dashboard")
    ap.add_argument("--backfill", action="store_true", help="reload full history")
    a = ap.parse_args()
    if a.test_api:
        y = pd.Timestamp(date.today()) - timedelta(days=1)
        df = fetch_api(CONFIG, y - timedelta(days=2), y, enabled_only=True)
        out = to_city_daily(df)
        out["cost_incl_gst"] = out.pop("cost") * CONFIG["gst_multiplier"]
        print(out.to_string())
        cm = df.groupby(["city", "campaign"], as_index=False).impressions.sum()
        print("\nCampaign -> city mapping (check every line):")
        print(cm.to_string(index=False))
        return
    if a.once or a.backfill:
        cloud_run(CONFIG, backfill=a.backfill)
        return
    serve(CONFIG)


if __name__ == "__main__":
    main()
