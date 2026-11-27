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
    "refresh_cols": ["impressions", "clicks", "cost", "conversions", "impr_share",
                     "budget_lost_is", "rank_lost_is"],  # re-pulled for 7 days (late Google Ads corrections)
    "backfill_start": "2024-11-01",

    # ---------- Google Trends ----------
    # Google Trends keywords: fetched together (one request, same scale) and kept SEPARATE (max 5)
    "trends_keywords": ["packers and movers", "house shifting", "home shifting", "bike shifting"],
    "decision_keyword": "packers and movers",   # keyword the actions / overview / insights are based on
    "trends_daily_window": 90,
    "trends_new_states_per_run": 2,
    "trends_max_429_per_run": 4,      # after this many "too many requests" replies, stop Trends for this run   # full Trends history is loaded for at most this many states per run     # days per pull (Trends only gives daily data up to ~90 days)
    "trends_recent_overwrite": 7,
    # ---- demand planning (forecast, all-India, locations, brand) ----
    "forecast_keyword": "packers and movers",   # used for Keyword Planner's main-keyword column
    "forecast_refresh_days": 7,                  # 5-year Trends history per state refreshed weekly
    "forecast_states_per_run": 4,                # at most this many 5-year pulls per run (rate limits)
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
    # ---- external factors ----
    # Social trend leading indicators: searches people make BEFORE they move
    "indicator_keywords": ["house for rent", "rent agreement", "flat for rent"],
    "heavy_rain_mm": 20,                 # a day with at least this much rain counts as a heavy-rain day
    "weather_refresh_days": 7,
    "macro_refresh_days": 30,
    # World Bank indicators for India (free, open licence)
    "macro_indicators": {"GDP growth %": "NY.GDP.MKTP.KD.ZG", "CPI inflation %": "FP.CPI.TOTL.ZG",
                         "Urban population growth %": "SP.URB.GROW"},
    # city centre coordinates for daily rainfall (NASA POWER)
    "city_coords": {"Mumbai": (19.076, 72.878), "Pune": (18.520, 73.857), "Nagpur": (21.146, 79.088),
                    "Thane": (19.218, 72.978), "Navi Mumbai": (19.033, 73.030), "Delhi": (28.614, 77.209),
                    "New Delhi": (28.614, 77.209), "Bangalore": (12.972, 77.595), "Bengaluru": (12.972, 77.595),
                    "Hyderabad": (17.385, 78.487), "Chennai": (13.083, 80.271), "Coimbatore": (11.017, 76.956),
                    "Kolkata": (22.573, 88.364), "Ahmedabad": (23.023, 72.571), "Surat": (21.170, 72.831),
                    "Jaipur": (26.912, 75.787), "Lucknow": (26.847, 80.946), "Noida": (28.536, 77.391),
                    "Gurgaon": (28.460, 77.027), "Gurugram": (28.460, 77.027), "Chandigarh": (30.733, 76.779),
                    "Indore": (22.720, 75.858), "Kochi": (9.931, 76.267), "Ludhiana": (30.901, 75.857)},
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

    # ---------- Segments within a city (ad group + campaign name rules; first match wins) ----------
    # Anything that matches no rule is "Generic / SKAG". Edit the words to match your naming.
    "segment_rules": [
        ("Brand", ["porter", "brand"]),
        ("Competitor", ["comp", "competitor", "nobroker", "agarwal", "leo packers", "leo ", "gati", "om packers"]),
        ("Intercity", ["intercity", "inter city", "interstate", "inter state", "outstation", "bike", "long distance"]),
        ("Generic / SKAG", ["skag", "generic", "packers", "shifting", "movers"]),
    ],
    "segment_backfill_days": 120,
    "segment_states_per_run": 3,
    # Google Trends terms per segment, pulled together per state (max 5 = one shared scale)
    "segment_trend_terms": [
        ("Brand", "porter packers and movers"),
        ("Generic / SKAG", "packers and movers"),
        ("Generic / SKAG", "house shifting"),
        ("Intercity", "bike shifting"),
        ("Intercity", "interstate shifting"),
    ],
    # Top-5 competitors tracked in Google Trends (display name, search term). Who leads where is measured, not assumed.
    "competitors": [
        ("NoBroker", "nobroker packers and movers"),
        ("Agarwal", "agarwal packers and movers"),
        ("Leo", "leo packers and movers"),
        ("Gati", "gati packers and movers"),
        ("Om Packers", "om packers and movers"),
    ],
    "state_names": {"IN-MH": "Maharashtra", "IN-DL": "Delhi", "IN-KA": "Karnataka", "IN-TG": "Telangana",
                    "IN-TN": "Tamil Nadu", "IN-WB": "West Bengal", "IN-GJ": "Gujarat", "IN-RJ": "Rajasthan",
                    "IN-UP": "Uttar Pradesh", "IN-HR": "Haryana", "IN-CH": "Chandigarh", "IN-MP": "Madhya Pradesh",
                    "IN-KL": "Kerala", "IN-PB": "Punjab"},
    # Keyword Planner CSVs downloaded from the Google Ads UI, used to check Google Trends (kp_IN.csv, kp_IN-KA.csv ...)
    "kp_manual_dir": "keyword_planner",
    # Monthly targets per city (columns: month, city, target_conversions, target_spend_incl_gst). Optional.
    "targets_csv": "targets.csv",
    "tcpa_relax_pct": 12.5,          # option B assumes CPA rises by about this much when tCPA is loosened 10–15%

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

BASE_COLS = ["impressions", "clicks", "cost", "conversions", "impr_share", "budget_lost_is", "rank_lost_is"]


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
        have_cols = set(q(con, "SELECT column_name AS c FROM information_schema.columns WHERE table_schema = %s "
                              "AND table_name = 'ads_campaign_daily'", (dbname,))["c"].str.lower())
        for col in ("budget_lost_is", "rank_lost_is"):
            if col not in have_cols:
                cur.execute(f"ALTER TABLE ads_campaign_daily ADD COLUMN {col} DOUBLE")
        cur.execute("""CREATE TABLE IF NOT EXISTS trends_daily (
            date VARCHAR(10) NOT NULL, geo VARCHAR(16) NOT NULL, idx DOUBLE, fetched_at VARCHAR(19),
            PRIMARY KEY (date, geo))""")
        cur.execute("""CREATE TABLE IF NOT EXISTS trends_daily_kw (
            date VARCHAR(10) NOT NULL, geo VARCHAR(16) NOT NULL, kw VARCHAR(64) NOT NULL, idx DOUBLE,
            fetched_at VARCHAR(19), PRIMARY KEY (date, geo, kw))""")
        cur.execute("""CREATE TABLE IF NOT EXISTS trends_weekly_5y_kw (
            geo VARCHAR(16) NOT NULL, kw VARCHAR(64) NOT NULL, week VARCHAR(10) NOT NULL, idx DOUBLE,
            fetched_at VARCHAR(19), PRIMARY KEY (geo, kw, week))""")
        cur.execute("""CREATE TABLE IF NOT EXISTS trends_city_kw (
            city VARCHAR(64) NOT NULL, kw VARCHAR(64) NOT NULL, idx DOUBLE, fetched_at VARCHAR(19),
            PRIMARY KEY (city, kw))""")
        cur.execute("""CREATE TABLE IF NOT EXISTS trends_indicator_5y (
            geo VARCHAR(16) NOT NULL, kw VARCHAR(64) NOT NULL, week VARCHAR(10) NOT NULL, idx DOUBLE,
            fetched_at VARCHAR(19), PRIMARY KEY (geo, kw, week))""")
        cur.execute("""CREATE TABLE IF NOT EXISTS weather_daily (
            city VARCHAR(64) NOT NULL, date VARCHAR(10) NOT NULL, rain_mm DOUBLE, fetched_at VARCHAR(19),
            PRIMARY KEY (city, date))""")
        cur.execute("""CREATE TABLE IF NOT EXISTS macro_annual (
            indicator VARCHAR(64) NOT NULL, year INT NOT NULL, value DOUBLE, fetched_at VARCHAR(19),
            PRIMARY KEY (indicator, year))""")
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
        cur.execute("""CREATE TABLE IF NOT EXISTS dashboard_files (
            name VARCHAR(32) NOT NULL, part INT NOT NULL, data LONGBLOB, updated_at VARCHAR(19),
            PRIMARY KEY (name, part))""")
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
           "conversions, impr_share, budget_lost_is, rank_lost_is, first_fetched_at, conv_updated_at) "
           "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) "
           f"ON DUPLICATE KEY UPDATE {upd_cols}, conv_updated_at = VALUES(conv_updated_at)")
    rows, upd = [], 0
    for r in df.itertuples(index=False):
        k = (str(r.date), str(r.campaign_id), str(r.city))
        upd += k in keys
        rows.append((*k[:2], r.campaign, k[2], float(r.impressions), float(r.clicks), float(r.cost),
                     float(r.conversions), *[None if pd.isna(getattr(r, c)) else float(getattr(r, c))
                                             for c in ("impr_share", "budget_lost_is", "rank_lost_is")], now, now))
    cur = con.cursor()
    for i in range(0, len(rows), 500):
        cur.executemany(sql, rows[i:i + 500])
    return len(rows) - upd, upd


def to_city_daily(df):
    """Campaign rows -> date x city. Shares are combined over eligible impressions (impressions / share)."""
    df = df.copy()
    for c in ["impressions", "clicks", "cost", "conversions", "impr_share", "budget_lost_is", "rank_lost_is"]:
        if c not in df.columns:
            df[c] = np.nan
        df[c] = pd.to_numeric(df[c], errors="coerce")
    ok = df["impr_share"] > 0
    df["elig"] = np.where(ok, df["impressions"] / df["impr_share"].where(ok), 0)
    df["imp_ok"] = np.where(ok, df["impressions"], 0)
    df["b_w"] = np.where(ok, df["budget_lost_is"].fillna(0) * df["elig"], 0)
    df["r_w"] = np.where(ok, df["rank_lost_is"].fillna(0) * df["elig"], 0)
    df["has_lost"] = np.where(ok & df["budget_lost_is"].notna(), df["elig"], 0)
    g = df.groupby(["date", "city"], as_index=False).agg(
        impressions=("impressions", "sum"), clicks=("clicks", "sum"), cost=("cost", "sum"),
        conversions=("conversions", "sum"), elig=("elig", "sum"), imp_ok=("imp_ok", "sum"),
        b_w=("b_w", "sum"), r_w=("r_w", "sum"), has_lost=("has_lost", "sum"))
    g["impr_share"] = g.imp_ok / g.elig.replace(0, np.nan)
    g["budget_lost_is"] = g.b_w / g.has_lost.replace(0, np.nan)
    g["rank_lost_is"] = g.r_w / g.has_lost.replace(0, np.nan)
    return g.drop(columns=["elig", "imp_ok", "b_w", "r_w", "has_lost"])


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


def read_trends_kw(con, geo=None):
    """All daily Trends: date, geo, kw, idx (every keyword on one shared scale per state)."""
    df = q(con, "SELECT date, geo, kw, idx FROM trends_daily_kw" + (" WHERE geo = %s" if geo else ""),
           (geo,) if geo else None)
    df["date"] = pd.to_datetime(df["date"])
    df["idx"] = pd.to_numeric(df["idx"])
    return df


def read_trends(con, kw=None):
    """Daily Trends of one keyword (the decision keyword by default): date, geo, idx."""
    kw = kw or CONFIG["decision_keyword"]
    df = q(con, "SELECT date, geo, idx FROM trends_daily_kw WHERE kw = %s", (kw,))
    df["date"] = pd.to_datetime(df["date"])
    df["idx"] = pd.to_numeric(df["idx"])
    return df


def write_trends(con, geo, frame, overwrite_from=None):
    """frame: index = date, one column per keyword. Insert days not stored; overwrite days >= overwrite_from."""
    now = datetime.now().isoformat(timespec="seconds")
    new, over = [], []
    for d, row in frame.iterrows():
        for kw, v in row.items():
            if pd.isna(v):
                continue
            r = (d.strftime("%Y-%m-%d"), geo, kw, float(v), now)
            (over if overwrite_from is not None and d >= overwrite_from else new).append(r)
    cur = con.cursor()
    if new:
        cur.executemany("INSERT IGNORE INTO trends_daily_kw (date, geo, kw, idx, fetched_at) VALUES (%s,%s,%s,%s,%s)", new)
    if over:
        cur.executemany("INSERT INTO trends_daily_kw (date, geo, kw, idx, fetched_at) VALUES (%s,%s,%s,%s,%s) "
                        "ON DUPLICATE KEY UPDATE idx = VALUES(idx), fetched_at = VALUES(fetched_at)", over)
    return len(frame)


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
               metrics.conversions, metrics.search_impression_share,
               metrics.search_budget_lost_impression_share, metrics.search_rank_lost_impression_share
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
                    sbl = m.get("searchBudgetLostImpressionShare")
                    srl = m.get("searchRankLostImpressionShare")
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
                        "budget_lost_is": np.nan if (use_geo or sbl in (None, "")) else float(sbl),
                        "rank_lost_is": np.nan if (use_geo or srl in (None, "")) else float(srl),
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
        conversions=("conversions", "sum"), impr_share=("impr_share", "mean"),
        budget_lost_is=("budget_lost_is", "mean"), rank_lost_is=("rank_lost_is", "mean"))
    return df[cols]


# =====================================================================
# GOOGLE TRENDS
# =====================================================================
class TrendsPaused(RuntimeError):
    pass


def _trends_guard():
    if _TRENDS["blocked"]:
        raise TrendsPaused("Google Trends is rate-limiting this run; the rest continues tomorrow")


def _trends_fail(e):
    if "429" in str(e):
        _TRENDS["fails"] += 1
        if _TRENDS["fails"] >= cfg_trends_limit():
            _TRENDS["blocked"] = True


def cfg_trends_limit():
    return CONFIG.get("trends_max_429_per_run", 4)


_TRENDS = {"fails": 0, "blocked": False}


def _trends_window(geo, cfg, start, end):
    from pytrends.request import TrendReq
    _trends_guard()
    py = TrendReq(hl="en-IN", tz=-330)
    tf = f"{start:%Y-%m-%d} {end:%Y-%m-%d}"
    for attempt in range(3):
        try:
            _trends_guard()
            py.build_payload(cfg["trends_keywords"], timeframe=tf, geo=geo)
            d = py.interest_over_time()
            break
        except TrendsPaused:
            raise
        except Exception as e:
            _trends_fail(e)
            wait = 45 * (attempt + 1)
            print(f"   Trends {geo} {tf} failed ({e}); retry in {wait}s")
            time.sleep(wait)
    else:
        raise RuntimeError(f"Trends failed for {geo} {tf}")
    if d.empty:
        return pd.DataFrame(columns=cfg["trends_keywords"], dtype=float)
    if "isPartial" in d.columns:
        partial = d["isPartial"].astype(bool)
        d = d[~partial].drop(columns="isPartial")   # drop incomplete latest day
    d = d.astype(float)
    d.index = pd.to_datetime(d.index)
    return d                                         # one column per keyword, shared scale


def _scale_onto(base, new, exclude_recent_days=7):
    """Scale `new` onto `base` using the ratio of their overlap (all keywords added together),
    ignoring the newest days. One ratio for every keyword keeps them comparable."""
    bs = base.sum(axis=1) if isinstance(base, pd.DataFrame) else base
    ns = new.sum(axis=1) if isinstance(new, pd.DataFrame) else new
    ov = bs.index.intersection(ns.index)
    if exclude_recent_days > 0 and len(ov) > exclude_recent_days + 14:
        ov = ov[:-exclude_recent_days]
    if len(ov) == 0 or ns.loc[ov].sum() == 0:
        return None
    return new * (bs.loc[ov].sum() / ns.loc[ov].sum())


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
    stitched = stitched / stitched.sum(axis=1).max() * 100   # one-time anchor scale, fixed from now on
    c = db_connect(cfg)                          # fresh connection: Trends waits can be long
    try:
        return write_trends(c, geo, stitched)
    finally:
        c.close()


def trends_daily_update(con, geo, cfg, end):
    c = db_connect(cfg)
    stored = read_trends_kw(c, geo)
    c.close()
    base = stored.pivot_table(index="date", columns="kw", values="idx").sort_index()
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
_STATE = {}


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
        print(f"[info] no Trends region mapped for: {nogeo}")
    # Trends keyword changed? Old daily Trends were a different search -> rebuild them
    kw_note = "|".join(cfg["trends_keywords"])
    kr = q(con, "SELECT note FROM job_runs WHERE job = 'trends-keyword'")
    if kr.empty or kr.note.iloc[0] != kw_note:
        con.cursor().execute("DELETE FROM trends_daily_kw")
        con.cursor().execute("DELETE FROM trends_weekly_5y_kw")
        con.cursor().execute("DELETE FROM job_runs WHERE job LIKE 't5y:%%' OR job = 'trends-cities'")
        _job_done(con, "trends-keyword", kw_note)
        log(con, "trends", f"Daily Trends reset to keyword(s): {kw_note}; history reloads over the next runs")
    # one-time: fill lost-impression-share for the last 120 days of already-stored rows
    if _job_due(con, "lost-is-backfill", 100000):
        try:
            hist = fetch_api(cfg, yday - timedelta(days=119), yday, enabled_only=False)
            hist = hist[hist.campaign_id.isin(stored_campaign_ids(con))]
            c2 = dict(cfg, refresh_cols=["impr_share", "budget_lost_is", "rank_lost_is"])
            upsert_ads(con, hist, c2)
            _job_done(con, "lost-is-backfill", f"{len(hist)} rows")
            log(con, "ads", f"Impression-share history filled for {len(hist)} rows")
        except Exception as e:
            log(con, "warn", f"Impression-share history not filled yet: {str(e)[:200]}")
    have = set(q(con, "SELECT DISTINCT geo FROM trends_daily_kw")["geo"])
    log(con, "refresh-ok", f"data up to {yday.date()}")   # Google Ads part is done
    con.close()

    _STATE["trends_plan"] = (cities, have, yday)


def run_trends_daily(cfg):
    """Daily Google Trends per state. Runs after the 5-year pulls, which matter more."""
    if "trends_plan" not in _STATE:
        return
    cities, have, yday = _STATE.pop("trends_plan")
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
        except TrendsPaused:
            print(f"[trends] {geo}: paused (rate limit); continues next run")
            break
        except Exception as e:
            try:
                c2 = db_connect(cfg); log(c2, "warn", f"Google Trends daily {geo}: {str(e)[:200]}"); c2.close()
            except Exception:
                print("trends warn", geo, e)
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
    try:
        sheets.update(build_external(cfg, con, ads, trends, sheets))
    except Exception as e:
        print("external factors:", e)
    try:
        import pnm_segments
        sheets.update(pnm_segments.build_segment_sheets(cfg, con, ads, trends, sheets))
    except Exception as e:
        print("segments analysis:", e)
    try:
        ov = build_overview(cfg, ads, trends, sheets)
        sheets = {"Budget_Board": ov["Budget_Board"], "Overview": ov["Overview"], "Actions": ov["Actions"],
                  "Insights": ov["Insights"], **sheets}
    except Exception as e:
        print("overview:", e)
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
        _trends_guard()
        try:
            return fn()
        except TrendsPaused:
            raise
        except Exception as e:
            _trends_fail(e)
            if attempt == tries - 1 or _TRENDS["blocked"]:
                raise
            wait = 45 * (attempt + 1)
            print(f"   {label} failed ({e}); retry in {wait}s")
            time.sleep(wait)


def pull_trends_5y(geo, cfg):
    """5 years of weekly Google Trends for the forecast keyword (one request, one consistent scale)."""
    def go():
        py = _trends_client()
        py.build_payload(cfg["trends_keywords"], timeframe="today 5-y", geo=geo)
        d = py.interest_over_time()
        if "isPartial" in d.columns:
            d = d[~d["isPartial"].astype(bool)].drop(columns="isPartial")
        return d.astype(float)                   # one column per keyword
    return _with_retry(go, f"Trends 5y {geo}")


def pull_trends_cities(cfg):
    """Relative search interest by Indian city over the last 12 months (100 = highest city)."""
    def go():
        py = _trends_client()
        py.build_payload(cfg["trends_keywords"], timeframe="today 12-m", geo="IN")
        d = py.interest_by_region(resolution="CITY", inc_low_vol=True)
        return d.astype(float)                   # rows = cities, one column per keyword
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
            cur.execute("DELETE FROM trends_weekly_5y_kw WHERE geo = %s", (geo,))
            cur.executemany("INSERT INTO trends_weekly_5y_kw (geo, kw, week, idx, fetched_at) VALUES (%s,%s,%s,%s,%s)",
                            [(geo, kw, d.strftime("%Y-%m-%d"), float(v), now)
                             for d, row in s5.iterrows() for kw, v in row.items() if pd.notna(v)])
            _job_done(con, f"t5y:{geo}", f"{len(s5)} weeks")
            log(con, "forecast-data", f"{geo}: {len(s5)} weeks")
            time.sleep(15)
            try:
                ind = pull_indicators_5y(geo, cfg)
                con = db_connect(cfg)
                cur = con.cursor()
                cur.execute("DELETE FROM trends_indicator_5y WHERE geo = %s", (geo,))
                cur.executemany("INSERT INTO trends_indicator_5y (geo, kw, week, idx, fetched_at) VALUES (%s,%s,%s,%s,%s)",
                                [(geo, kw, d.strftime("%Y-%m-%d"), float(v), now)
                                 for d, row in ind.iterrows() for kw, v in row.items() if pd.notna(v)])
                log(con, "signals", f"{geo}: rental-search history {len(ind)} weeks")
            except Exception as e:
                con = db_connect(cfg); log(con, "warn", f"Google Trends rental searches {geo}: {str(e)[:200]}")
        except TrendsPaused:
            print(f"[trends] {geo}: paused (rate limit); continues next run")
            break
        except Exception as e:
            con = db_connect(cfg); log(con, "warn", f"Google Trends 5-year {geo}: {str(e)[:200]}")
        time.sleep(15)

    # 1b) daily rainfall (NASA POWER) and macro context (World Bank)
    try:
        pull_weather_and_macro(cfg, con, cities)
    except Exception as e:
        con = db_connect(cfg); log(con, "warn", f"weather/macro: {str(e)[:300]}")

    # 2) Trends interest by city across India
    if _job_due(con, "trends-cities", cfg["forecast_refresh_days"]):
        try:
            sc = pull_trends_cities(cfg)
            con = db_connect(cfg)
            cur = con.cursor()
            cur.execute("DELETE FROM trends_city_kw")
            cur.executemany("INSERT INTO trends_city_kw (city, kw, idx, fetched_at) VALUES (%s,%s,%s,%s)",
                            [(str(c)[:64], kw, float(v), now) for c, row in sc.iterrows() for kw, v in row.items()])
            _job_done(con, "trends-cities", f"{len(sc)} cities")
            log(con, "locations", f"{len(sc)} cities")
        except Exception as e:
            con = db_connect(cfg); log(con, "warn", f"Google Trends cities: {str(e)[:200]}")

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
                if _STATE.get("kp_blocked"):
                    break
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
                    if "DEVELOPER_TOKEN_NOT_APPROVED" in str(e) or "explorer access" in str(e):
                        _STATE["kp_blocked"] = True
                        for c2 in due:
                            _job_done(con, f"kp:{c2}", "needs Basic access")
                        log(con, "warn", "Keyword Planner needs Basic or Standard Google Ads API access "
                                         "(you have Explorer). Search volumes and brand data are paused; "
                                         "the script retries in 30 days.")
                        break
                    log(con, "warn", f"Keyword Planner {c}: {str(e)[:300]}")
                time.sleep(2)
            if not _STATE.get("kp_blocked"):
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
    w5 = q(con, "SELECT geo, kw, week, idx FROM trends_weekly_5y_kw")
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

    kws = list(cfg["trends_keywords"])
    ALL = "All keywords"
    tkw = read_trends_kw(con)
    w5["idx"] = pd.to_numeric(w5.idx)
    fw, fd, up, acc = [], [], [], []
    for city in cities:
        geo = cfg["city_to_geo"].get(city)
        if not geo:
            continue
        gw = w5[w5.geo == geo]
        if gw.empty:
            continue
        piv5 = gw.pivot_table(index=pd.to_datetime(gw.week), columns="kw", values="idx").sort_index()
        gd = tkw[tkw.geo == geo]
        pivd = gd.pivot_table(index="date", columns="kw", values="idx").sort_index() if len(gd) else None
        for kw in kws + [ALL]:
            if kw == ALL:
                s5 = piv5.sum(axis=1, min_count=1)
                trg = pivd.sum(axis=1, min_count=1) if pivd is not None else None
            else:
                if kw not in piv5:
                    continue
                s5 = piv5[kw]
                trg = pivd[kw] if pivd is not None and kw in pivd else None
            fc, err = forecast_weekly(s5.dropna(), cfg)
            if fc is None:
                continue
            fc.insert(0, "city", city)
            fc.insert(1, "kw", kw)
            fc["week_end"] = fc.week_start + timedelta(days=6)
            fc["events"] = [", ".join(e for e in future_cal.loc[r.week_start:r.week_end, "event"] if e)
                            for r in fc.itertuples()]
            fc["past_cpc_this_month_pct"] = [cpc_m.get(city, {}).get(d.month) for d in fc.week_start]
            fw.append(fc)
            acc.append({"city": city, "kw": kw, "trends_geo": geo,
                        "backtest_typical_error_pct": None if err is None else round(err, 1),
                        "history_weeks": int(s5.notna().sum())})

            # notable upcoming weeks: runs of high/low weeks, merged
            t = cfg["forecast_threshold_pct"]
            flag = np.where(fc.vs_normal_pct >= t, "High", np.where(fc.vs_normal_pct <= -t, "Low", ""))
            runs, run = [], None
            for i, f in enumerate(flag):
                if f and run and run["type"] == f and i == run["end_i"] + 1:
                    run["end_i"] = i
                else:
                    if run:
                        runs.append(run)
                    run = {"city": city, "kw": kw, "type": f, "start_i": i, "end_i": i} if f else None
            if run:
                runs.append(run)
            for u in runs:
                seg = fc.iloc[u["start_i"]:u["end_i"] + 1]
                peak = seg.vs_normal_pct.abs().idxmax()
                u.update({"from": seg.week_start.iloc[0], "to": seg.week_end.iloc[-1],
                          "demand_vs_normal_pct": round(float(seg.vs_normal_pct.loc[peak]), 1),
                          "suggestion": fc.suggestion.loc[peak],
                          "past_cpc_this_month_pct": fc.past_cpc_this_month_pct.loc[peak],
                          "events": ", ".join(sorted({e for e in seg.events if e}))[:200]})
                up.append(u)

            # next 90 days, day by day
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
                fd.append({"city": city, "kw": kw, "date": d, "dow": d.day_name()[:3],
                           "day_vs_normal_pct": round((base * f_dow * f_dom - 1) * 100, 1),
                           "event": future_cal.loc[d, "event"] if d in future_cal.index else ""})

    # weekly search trend per keyword (last 2 years), for the keyword chart
    tw = []
    since = today - timedelta(days=730)
    for city in cities:
        geo = cfg["city_to_geo"].get(city)
        gd = tkw[(tkw.geo == geo) & (tkw.date >= since)] if geo else tkw.iloc[0:0]
        if gd.empty:
            continue
        piv = gd.pivot_table(index="date", columns="kw", values="idx").sort_index()
        piv[ALL] = piv.sum(axis=1, min_count=1)
        wkly = piv.resample("W-MON", label="left", closed="left").mean()
        wkly = wkly[piv.resample("W-MON", label="left", closed="left").size() == 7]
        for kw in wkly.columns:
            for d, v in wkly[kw].items():
                if pd.notna(v):
                    tw.append({"city": city, "kw": kw, "week_start": d, "idx": round(float(v), 2)})
    out["Trends_Weekly_KW"] = pd.DataFrame(tw)

    cat = lambda L: pd.concat(L, ignore_index=True) if L else pd.DataFrame()
    out["Forecast_Weekly"] = cat(fw)
    out["Forecast_Daily"] = pd.DataFrame(fd)
    out["Upcoming_Changes"] = pd.DataFrame([{k: v for k, v in u.items() if k not in ("start_i", "end_i")}
                                            for u in up if "from" in u])
    out["Forecast_Accuracy"] = pd.DataFrame(acc)

    # Locations: Trends city interest + Keyword Planner volume + our coverage
    ci = q(con, "SELECT city, kw, idx FROM trends_city_kw")
    if not ci.empty:
        ci["idx"] = pd.to_numeric(ci.idx)
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
        tri = tr.groupby("kw").idx.max().to_dict() if len(tr) else {}
        k = kp[(kp.city == ours) & (kp.keyword.str.lower() == cfg["forecast_keyword"])] if not kp.empty else kp
        gen = kp[(kp.city == ours) & (kp.grp == "Generic")] if not kp.empty else kp
        vol = float(k.sort_values("ym").searches.tail(12).mean()) if len(k) else None
        gvol = float(gen.groupby("ym").searches.sum().sort_index().tail(12).mean()) if len(gen) else None
        rows.append({"city": ours, "running_campaigns": ours in last30.index,
                     "trends_interest_0_100": tri.get(cfg["decision_keyword"]),
                     **{f"interest: {k}": tri.get(k) for k in cfg["trends_keywords"]},
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
# EXTERNAL FACTORS: social leading signals, festivals, weather, macro context
# =====================================================================
def pull_indicators_5y(geo, cfg):
    """5 years of weekly Trends for searches people make before moving (rent listings, rent agreements)."""
    def go():
        py = _trends_client()
        py.build_payload(cfg["indicator_keywords"][:5], timeframe="today 5-y", geo=geo)
        d = py.interest_over_time()
        if "isPartial" in d.columns:
            d = d[~d["isPartial"].astype(bool)].drop(columns="isPartial")
        return d.astype(float)
    return _with_retry(go, f"Trends indicators {geo}")


def pull_weather_and_macro(cfg, con, cities):
    now = datetime.now().isoformat(timespec="seconds")
    # daily rainfall per city from NASA POWER (public, no key)
    for c in cities:
        if c == "All India" or c not in cfg["city_coords"]:
            continue
        if not _job_due(con, f"wx:{c}", cfg["weather_refresh_days"]):
            continue
        lat, lon = cfg["city_coords"][c]
        start = pd.Timestamp(cfg["backfill_start"]).strftime("%Y%m%d")
        end = (date.today() - timedelta(days=1)).strftime("%Y%m%d")
        url = ("https://power.larc.nasa.gov/api/temporal/daily/point?parameters=PRECTOTCORR&community=AG"
               f"&longitude={lon}&latitude={lat}&start={start}&end={end}&format=JSON")
        try:
            r = requests.get(url, timeout=120)
            r.raise_for_status()
            vals = r.json()["properties"]["parameter"]["PRECTOTCORR"]
            rows = [(c, f"{k[:4]}-{k[4:6]}-{k[6:]}", float(v), now) for k, v in vals.items() if v is not None and v > -900]
            con = db_connect(cfg)
            con.cursor().executemany(
                "INSERT INTO weather_daily (city, date, rain_mm, fetched_at) VALUES (%s,%s,%s,%s) "
                "ON DUPLICATE KEY UPDATE rain_mm = VALUES(rain_mm), fetched_at = VALUES(fetched_at)", rows)
            _job_done(con, f"wx:{c}", f"{len(rows)} days")
        except Exception as e:
            log(con, "warn", f"rainfall {c}: {str(e)[:200]}")
        time.sleep(1)
    # annual macro indicators for India from the World Bank (open data)
    if _job_due(con, "macro", cfg["macro_refresh_days"]):
        rows = []
        for name, code in cfg["macro_indicators"].items():
            try:
                for attempt in range(4):
                    r = requests.get(f"https://api.worldbank.org/v2/country/IND/indicator/{code}?format=json&per_page=20",
                                     timeout=60, headers={"User-Agent": "pnm-demand-monitor/1.0"})
                    if r.status_code != 429:
                        break
                    time.sleep(10 * (attempt + 1))        # World Bank asks callers to slow down
                r.raise_for_status()
                js = r.json()
                time.sleep(3)
                for it in (js[1] if isinstance(js, list) and len(js) > 1 and js[1] else []):
                    if it.get("value") is not None:
                        rows.append((name, int(it["date"]), float(it["value"]), now))
            except Exception as e:
                print(f"   macro {name} not fetched this run ({str(e)[:120]}); retries next run")
        got = {r[0] for r in rows}
        if rows:
            con.cursor().executemany(
                "INSERT INTO macro_annual (indicator, year, value, fetched_at) VALUES (%s,%s,%s,%s) "
                "ON DUPLICATE KEY UPDATE value = VALUES(value), fetched_at = VALUES(fetched_at)", rows)
            if got >= set(cfg["macro_indicators"]):
                _job_done(con, "macro", f"{len(rows)} values")
            log(con, "macro", f"{len(rows)} values ({', '.join(sorted(got))})")


def _clean_event(e):
    e = re.sub(r"^\[[^\]]*\]\s*", "", e.strip())
    return re.sub(r"\s*\(estimated\)", "", e).strip()


def lead_signal(ind, dem, max_lag=8):
    """Does the indicator move before demand? Uses 4-week % changes (seasonality-neutral),
    tests lags 1..max_lag weeks and returns the best positive relationship."""
    both = pd.concat([ind, dem], axis=1, keys=["ind", "dem"]).dropna()
    if len(both) < 104:
        return None
    ci = both.ind.pct_change(4)
    cd = both.dem.pct_change(4)
    best = None
    for L in range(1, max_lag + 1):
        x, y = ci.shift(L), cd
        ok = x.notna() & y.notna() & np.isfinite(x) & np.isfinite(y)
        if ok.sum() < 60:
            continue
        r = float(np.corrcoef(x[ok], y[ok])[0, 1])
        if best is None or r > best["r"]:
            slope = float(np.polyfit(x[ok], y[ok], 1)[0])
            best = {"lag_weeks": L, "r": r, "slope": slope}
    if best is None:
        return None
    latest = ci.dropna()
    best["indicator_4w_change_pct"] = float(latest.iloc[-1] * 100) if len(latest) else None
    best["expected_demand_change_pct"] = (best["slope"] * latest.iloc[-1] * 100) if len(latest) else None
    best["as_of_week"] = both.index[-1]
    best["strength"] = "strong" if best["r"] >= 0.5 else "moderate" if best["r"] >= 0.3 else "weak"
    return best


def build_external(cfg, con, ads, trends, sh):
    """Sheets: Leading_Signal, Leading_Weekly, Event_Impact, Weather_Impact, Macro_Context."""
    out = {}
    dk = cfg["decision_keyword"]
    today = pd.Timestamp(date.today())
    cities = ["All India"] + sorted(c for c in ads.city.unique() if c != "All India")
    ind = q(con, "SELECT geo, kw, week, idx FROM trends_indicator_5y")
    w5 = q(con, "SELECT geo, kw, week, idx FROM trends_weekly_5y_kw WHERE kw = %s", (dk,))
    for d in (ind, w5):
        if len(d):
            d["idx"] = pd.to_numeric(d.idx)
            d["week"] = pd.to_datetime(d.week)

    # 1) social leading signal: rental searches -> packers demand
    sig, lw = [], []
    for city in cities:
        geo = cfg["city_to_geo"].get(city)
        gi, gd = ind[ind.geo == geo], w5[w5.geo == geo]
        if gi.empty or gd.empty:
            continue
        si = gi.groupby("week").idx.sum().sort_index()
        sd = gd.set_index("week").idx.sort_index()
        res = lead_signal(si, sd)
        if res:
            sig.append({"city": city, "trends_geo": geo, **res})
        both = pd.concat([si, sd], axis=1, keys=["rental_searches", "demand"]).dropna()
        both = both[both.index >= today - timedelta(days=730)]
        for wk, r in both.iterrows():
            lw.append({"city": city, "week_start": wk, "rental_searches": r.rental_searches, "demand": r.demand})
    out["Leading_Signal"] = pd.DataFrame(sig)
    out["Leading_Weekly"] = pd.DataFrame(lw)

    # 2) festivals / holidays / your events: typical effect on demand, and when they come next
    ev = []
    hist_cal = build_calendar(cfg, pd.Timestamp(cfg["backfill_start"]), today)
    fut_cal = build_calendar(cfg, today, today + timedelta(days=365))
    nxt = {}
    for d, e in fut_cal.event.items():
        for part in [x for x in e.split("; ") if x]:
            n = _clean_event(part)
            nxt.setdefault(n, d)
    for city in cities:
        geo = cfg["city_to_geo"].get(city)
        t = trends[trends.geo == geo].set_index("date").idx.sort_index() if geo else pd.Series(dtype=float)
        if len(t) < 120:
            continue
        ratio = t / t.rolling(29, center=True, min_periods=15).median()   # removes season + trend
        dow_adj = ratio.groupby(ratio.index.dayofweek).transform("mean")
        ratio = ratio / dow_adj                                           # removes the weekday effect
        effects = {}
        for d, e in hist_cal.event.items():
            for part in [x for x in e.split("; ") if x]:
                n = _clean_event(part)
                win = ratio.loc[d - timedelta(days=1): d + timedelta(days=1)].dropna()
                if len(win):
                    effects.setdefault(n, []).append(float(win.mean() - 1) * 100)
        for n, vals in effects.items():
            ev.append({"city": city, "event": n, "times_seen": len(vals),
                       "typical_effect_pct": round(float(np.mean(vals)), 1),
                       "next_date": nxt.get(n)})
    evdf = pd.DataFrame(ev)
    if len(evdf):
        evdf = evdf.sort_values(["city", "next_date"], na_position="last")
    out["Event_Impact"] = evdf

    # 3) weather: heavy-rain days vs normal days
    wx = q(con, "SELECT city, date, rain_mm FROM weather_daily")
    wrows = []
    if len(wx):
        wx["date"] = pd.to_datetime(wx.date)
        wx["rain_mm"] = pd.to_numeric(wx.rain_mm)
        for city in cities:
            r = wx[wx.city == city].set_index("date").rain_mm.sort_index()
            geo = cfg["city_to_geo"].get(city)
            t = trends[trends.geo == geo].set_index("date").idx.sort_index() if geo else pd.Series(dtype=float)
            imp = ads[ads.city == city].set_index("date").impressions.sort_index()
            if r.empty:
                continue
            heavy = r[r >= cfg["heavy_rain_mm"]].index

            def eff(sr):
                if len(sr) < 120:
                    return None
                rr = sr / sr.rolling(29, center=True, min_periods=15).median()
                rr = rr / rr.groupby(rr.index.dayofweek).transform("mean")
                h = rr.reindex(heavy).dropna()
                return round(float(h.mean() - 1) * 100, 1) if len(h) >= 5 else None
            wrows.append({"city": city, "heavy_rain_days": int(len(heavy)),
                          "demand_effect_pct": eff(t), "impressions_effect_pct": eff(imp),
                          "rain_last_7d_mm": round(float(r[r.index > r.index.max() - timedelta(days=7)].sum()), 1),
                          "heavy_rain_days_last_30d": int((r[r.index > r.index.max() - timedelta(days=30)]
                                                           >= cfg["heavy_rain_mm"]).sum()),
                          "data_until": r.index.max()})
    out["Weather_Impact"] = pd.DataFrame(wrows)

    # 4) macro context: India yearly indicators next to yearly demand change
    mc = q(con, "SELECT indicator, year, value FROM macro_annual")
    if len(mc):
        mc["value"] = pd.to_numeric(mc.value)
        m = mc.pivot_table(index="year", columns="indicator", values="value")
        gi = w5[w5.geo == "IN"].set_index("week").idx if len(w5) else pd.Series(dtype=float)
        if len(gi):
            yr = gi.groupby(gi.index.year).mean()
            cnt = gi.groupby(gi.index.year).size()
            yr = yr[cnt >= 40]
            m["Demand change % (packers and movers, India)"] = (yr / yr.shift(1) - 1) * 100
        m = m[m.index >= today.year - 8].reset_index().sort_values("year", ascending=False)
        out["Macro_Context"] = m.round(2)
    return out


# =====================================================================
# AT-A-GLANCE: one row per city, plain-language actions and insights
# =====================================================================
def _chg(a, b):
    return None if (a is None or b is None or not b or pd.isna(a) or pd.isna(b)) else (a / b - 1) * 100


def _read_targets(cfg):
    path = cfg.get("targets_csv")
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    try:
        t = pd.read_csv(path, dtype={"month": str})
        t.columns = [c.strip().lower() for c in t.columns]
        t["city"] = t["city"].astype(str).str.strip()
        t["month"] = t["month"].astype(str).str.strip().str[:7]
        for c in ("target_conversions", "target_spend_incl_gst"):
            t[c] = pd.to_numeric(t.get(c), errors="coerce")
        return t
    except Exception as e:
        print("targets.csv not read:", e)
        return pd.DataFrame()


def _pacing(tg, city, a, end, cpa30, a_daily, b_daily, a_bid, b_bid, cfg):
    """Month-end projection under option A and B against this month's target (if one is set)."""
    month = end.strftime("%Y-%m")
    row = tg[(tg.month == month) & (tg.city == city)] if len(tg) else tg
    if len(tg) and not len(row) and city == "All India":
        row = tg[(tg.month == month) & (tg.city != "All India")]
        t_conv = row.target_conversions.sum(min_count=1) if len(row) else None
        t_spend = row.target_spend_incl_gst.sum(min_count=1) if len(row) else None
    elif len(row):
        t_conv, t_spend = row.target_conversions.iloc[0], row.target_spend_incl_gst.iloc[0]
    else:
        return {}
    t_conv = None if t_conv is None or pd.isna(t_conv) else float(t_conv)
    t_spend = None if t_spend is None or pd.isna(t_spend) else float(t_spend)
    if t_conv is None and t_spend is None:
        return {}
    mtd = a[(a.index >= end.replace(day=1)) & (a.index <= end)]
    dim = pd.Period(month).days_in_month
    left = dim - end.day
    relax = 1 + cfg["tcpa_relax_pct"] / 100
    cpa_a = cpa30
    cpa_b = cpa30 * (relax if ("Relax" in b_bid or ("Raise" in b_bid and "Raise" not in a_bid)) else 1) if cpa30 else None
    mc, ms = float(mtd.conversions.sum()), float(mtd.cost.sum())
    pa = mc + left * a_daily / cpa_a if cpa_a else None
    pb = mc + left * b_daily / cpa_b if cpa_b else None
    sa, sb = ms + left * a_daily, ms + left * b_daily
    pc = lambda x, t: None if x is None or not t else round((x / t - 1) * 100, 1)
    hint = None
    if t_conv:
        ga, gb = pc(pa, t_conv), pc(pb, t_conv)
        if ga is not None and ga >= -5:
            hint = "On track under A: efficiency is safe."
        elif gb is not None and gb >= -5:
            hint = "Behind target under A but on track under B: B suits this month."
        elif ga is not None:
            hint = "Behind target under both: discuss target or extra budget with the client."
    return {"target_month": month, "target_conversions": t_conv, "target_spend": t_spend,
            "mtd_conversions": round(mc, 1), "mtd_spend": round(ms), "days_left": left,
            "proj_conv_a": round(pa) if pa is not None else None, "proj_conv_b": round(pb) if pb is not None else None,
            "proj_spend_a": round(sa), "proj_spend_b": round(sb),
            "pace_conv_a_pct": pc(pa, t_conv), "pace_conv_b_pct": pc(pb, t_conv),
            "pace_spend_a_pct": pc(sa, t_spend), "pace_spend_b_pct": pc(sb, t_spend), "target_hint": hint}


def build_overview(cfg, ads, trends, sh):
    today = pd.Timestamp(date.today())
    _targets = _read_targets(cfg)
    mat = today - timedelta(days=cfg["maturity_days"] + 1)
    dk = cfg["decision_keyword"]
    fw = sh.get("Forecast_Weekly", pd.DataFrame())
    up = sh.get("Upcoming_Changes", pd.DataFrame())
    fw = fw[fw.kw == dk] if len(fw) and "kw" in fw else fw
    up = up[up.kw == dk] if len(up) and "kw" in up else up
    dowp, domp = sh.get("DOW_Profile", pd.DataFrame()), sh.get("DayOfMonth_Profile", pd.DataFrame())
    summ = sh.get("Summary", pd.DataFrame())
    sigs = sh.get("Leading_Signal", pd.DataFrame())
    evi = sh.get("Event_Impact", pd.DataFrame())
    wxi = sh.get("Weather_Impact", pd.DataFrame())
    rows, acts, ins, budget_rows = [], [], [], []
    cities = ["All India"] + sorted(c for c in ads.city.unique() if c != "All India")
    for city in cities:
        a = ads[ads.city == city].set_index("date").sort_index()
        if a.empty:
            continue
        end = a.index.max()
        geo = cfg["city_to_geo"].get(city)
        t = trends[trends.geo == geo].set_index("date")["idx"].sort_index() if geo else pd.Series(dtype=float)

        # demand now: last 7 days of Trends vs its last-12-months average
        d7 = None
        if len(t) > 60:
            d7 = _chg(t.iloc[-7:].mean(), t.iloc[-365:].mean())
        # demand next 4 weeks from the forecast
        n4 = None
        f = fw[fw.city == city] if len(fw) else fw
        if len(f):
            n4 = float(f.vs_normal_pct.iloc[:4].mean())
        # our side
        w = lambda d0, d1: a[(a.index > d0) & (a.index <= d1)]
        last7, prev28 = w(end - timedelta(days=7), end), w(end - timedelta(days=35), end - timedelta(days=7))
        imp_chg = _chg(last7.impressions.sum() / 7, prev28.impressions.sum() / 28 if len(prev28) else None)
        l14 = w(end - timedelta(days=14), end)
        base = w(end - timedelta(days=74), end - timedelta(days=14))      # the 60 days before: "normal"

        def wavg(df, col):
            e = (df.impressions / df.impr_share.where(df.impr_share > 0)).fillna(0)
            m = df[col].notna() & (e > 0)
            return (df[col][m] * e[m]).sum() / e[m].sum() if e[m].sum() else None
        r_base, b_base = wavg(base, "rank_lost_is"), wavg(base, "budget_lost_is")
        el = (l14.impressions / l14.impr_share.where(l14.impr_share > 0)).fillna(0)
        is14 = l14.impressions[el > 0].sum() / el.sum() if el.sum() else None
        b14 = (l14.budget_lost_is.fillna(0) * el).sum() / el[l14.budget_lost_is.notna()].sum() \
            if el[l14.budget_lost_is.notna()].sum() else None
        r14 = (l14.rank_lost_is.fillna(0) * el).sum() / el[l14.rank_lost_is.notna()].sum() \
            if el[l14.rank_lost_is.notna()].sum() else None
        c30, cp30 = w(mat - timedelta(days=30), mat), w(mat - timedelta(days=60), mat - timedelta(days=30))
        cpa30 = c30.cost.sum() / c30.conversions.sum() if c30.conversions.sum() else None
        cpap = cp30.cost.sum() / cp30.conversions.sum() if cp30.conversions.sum() else None
        cpa_chg = _chg(cpa30, cpap)
        cost30 = w(end - timedelta(days=30), end).cost.sum()

        # ---------- actions (plain language, max 3) ----------
        A = []
        pctf = lambda v: f"{v:+.0f}%"
        if n4 is not None and n4 >= 10 and b14 is not None and b14 >= 0.10:
            A.append(("Raise budget", f"Raise daily budget ~{int(min(40, max(10, round((n4 + b14 * 100) / 5) * 5)))}%",
                      f"Demand next 4 weeks {pctf(n4)} vs normal, with {b14:.0%} of impressions already lost to budget."))
        if n4 is not None and n4 >= 10 and r14 is not None and r14 >= 0.30:
            A.append(("Raise bids", "Raise bids / tCPA ~10–15%",
                      f"Demand is rising ({pctf(n4)}) and {r14:.0%} of impressions are lost to ad rank."))
        if n4 is not None and n4 >= 10 and not A:
            A.append(("Prepare", "Keep budget headroom",
                      f"Demand next 4 weeks {pctf(n4)} vs normal; make sure budgets aren’t capped."))
        if n4 is not None and n4 <= -10:
            A.append(("Low demand", "Low demand ahead: choose efficiency or share",
                      f"Demand next 4 weeks {pctf(n4)} vs normal. A) Efficiency: trim budget ~{int(min(30, round(-n4 / 5) * 5))}%, "
                      f"hold bids, protect CPA. B) Targets / share: keep budget, relax tCPA ~10–15% to keep spending "
                      f"(CPA likely higher). The AM decides based on this month’s target."))
        # rank / budget: flag a clear worsening vs this city's own normal, or extreme levels
        if r14 is not None and not any(x[0] == "Raise bids" for x in A):
            if r_base is not None and r14 - r_base >= 0.08:
                A.append(("Weak ad rank", "Losing more auctions to ad rank",
                          f"Lost to rank {r14:.0%} vs a 60-day norm of {r_base:.0%}: competitors may be bidding up. "
                          f"Check bids/tCPA and ad strength."))
            elif r14 >= 0.65:
                A.append(("Weak ad rank", "Ad rank is very weak",
                          f"{r14:.0%} of eligible impressions lost to rank: review bids/tCPA, ads and landing page."))
        if b14 is not None and not any(x[0] == "Raise budget" for x in A):
            if b14 >= 0.15 and (b_base is None or b14 - b_base >= -0.02):
                A.append(("Budget-capped", "Budget is limiting reach",
                          f"{b14:.0%} of eligible impressions lost to budget in 14 days"
                          + (f" (normal {b_base:.0%})." if b_base is not None else ".")))
        if cpa_chg is not None and cpa_chg >= 25 and (d7 is None or abs(d7) < 10):
            A.append(("Check efficiency", "CPA rising without a demand change",
                      f"CPA {pctf(cpa_chg)} vs the previous 30 days while demand is normal: review search terms, bids, landing page."))
        if imp_chg is not None and imp_chg <= -15 and d7 is not None and d7 >= -5:
            A.append(("Missed demand", "Impressions down while demand is steady",
                      f"Impressions {pctf(imp_chg)} vs the previous 4 weeks, demand {pctf(d7)}: check budgets, disapprovals, paused campaigns."))
        # external factors: social leading signal, festivals in the next 14 days
        sg = sigs[sigs.city == city] if len(sigs) else sigs
        sig = sg.iloc[0] if len(sg) else None
        if sig is not None and sig["r"] >= 0.3 and pd.notna(sig["expected_demand_change_pct"]):
            x = float(sig["expected_demand_change_pct"])
            if x >= 8:
                A.append(("Early signal", f"Demand likely up ~{x:.0f}% in ~{int(sig['lag_weeks'])} weeks",
                          f"Rental searches moved {sig['indicator_4w_change_pct']:+.0f}% in 4 weeks; they usually lead "
                          f"moving demand by {int(sig['lag_weeks'])} weeks ({sig['strength']} link). Line up budget now."))
            elif x <= -8:
                A.append(("Early signal", f"Demand likely down ~{-x:.0f}% in ~{int(sig['lag_weeks'])} weeks",
                          f"Rental searches moved {sig['indicator_4w_change_pct']:+.0f}% in 4 weeks; they usually lead "
                          f"moving demand by {int(sig['lag_weeks'])} weeks ({sig['strength']} link). Avoid over-spending."))
        ec = evi[(evi.city == city) & evi.next_date.notna()] if len(evi) else evi
        if len(ec):
            soon = ec[(pd.to_datetime(ec.next_date) <= today + timedelta(days=14)) & (ec.times_seen >= 1)]
            for _, e in soon.sort_values("next_date").iterrows():
                if e.typical_effect_pct <= -10:
                    A.append(("Festival ahead", f"{e.event} on {pd.Timestamp(e.next_date).strftime('%d %b')}: lower bids that day",
                              f"Demand typically {e.typical_effect_pct:+.0f}% around it (seen {int(e.times_seen)}×)."))
                elif e.typical_effect_pct >= 10:
                    A.append(("Festival ahead", f"{e.event} on {pd.Timestamp(e.next_date).strftime('%d %b')}: keep budget open",
                              f"Demand typically {e.typical_effect_pct:+.0f}% around it (seen {int(e.times_seen)}×)."))
        if not A:
            A.append(("On track", "No change needed", "Demand and delivery look normal."))
        for i, (kind, title, why) in enumerate(A[:4]):
            acts.append({"city": city, "rank": i + 1, "kind": kind, "action": title, "why": why})

        # ---------- quick insights ----------
        I = []
        if d7 is not None:
            I.append(f"Demand last 7 days: {pctf(d7)} vs its 12-month average.")
        p = dowp[dowp.city == city] if len(dowp) else dowp
        if len(p) and p.trends_index_rel.notna().any():
            hi, lo = p.loc[p.trends_index_rel.idxmax()], p.loc[p.trends_index_rel.idxmin()]
            I.append(f"Busiest search day: {hi.dow} ({hi.trends_index_rel - 100:+.0f}%); quietest: {lo.dow} ({lo.trends_index_rel - 100:+.0f}%).")
        p = domp[domp.city == city] if len(domp) else domp
        if len(p) and p.trends_index_rel.notna().any():
            e = p.set_index("dom_bucket").trends_index_rel
            ends = e.reindex(["Start (1-3)", "End (last 3)"]).mean()
            mid = e.reindex(["Early-mid (4-15)", "Late-mid (16-end-3)"]).mean()
            if pd.notna(ends) and pd.notna(mid):
                I.append(f"Month start/end vs mid-month: {ends / mid * 100 - 100:+.0f}% demand (rent-cycle effect).")
        if len(f):
            m = f.assign(mn=pd.to_datetime(f.week_start).dt.strftime("%b")).groupby("mn", sort=False).vs_normal_pct.mean()
            peak = [k for k, v in m.items() if v >= 8]
            quiet = [k for k, v in m.items() if v <= -8]
            if peak:
                I.append(f"Peak months ahead: {', '.join(peak)}.")
            if quiet:
                I.append(f"Quiet months ahead: {', '.join(quiet)}.")
        u = up[up.city == city] if len(up) else up
        if len(u):
            nx = u.iloc[0]
            I.append(f"Next big shift: {'busier' if nx['type'] == 'High' else 'quieter'} from "
                     f"{pd.Timestamp(nx['from']).strftime('%d %b')} to {pd.Timestamp(nx['to']).strftime('%d %b')} "
                     f"({nx['demand_vs_normal_pct']:+.0f}%).")
        if sig is not None and sig["r"] >= 0.3:
            I.append(f"Rental searches lead moving demand by ~{int(sig['lag_weeks'])} weeks here ({sig['strength']} link); "
                     f"they changed {sig['indicator_4w_change_pct']:+.0f}% in the last 4 weeks.")
        if len(ec):
            big = ec[ec.typical_effect_pct.abs() >= 10].sort_values("next_date").head(2)
            for _, e in big.iterrows():
                I.append(f"{e.event} ({pd.Timestamp(e.next_date).strftime('%d %b')}): demand usually {e.typical_effect_pct:+.0f}%.")
        wr = wxi[wxi.city == city] if len(wxi) else wxi
        if len(wr) and pd.notna(wr.demand_effect_pct.iloc[0]) and abs(wr.demand_effect_pct.iloc[0]) >= 5:
            I.append(f"Heavy-rain days (≥{cfg['heavy_rain_mm']} mm): demand {wr.demand_effect_pct.iloc[0]:+.0f}% vs a normal day.")
        sm = summ[summ.city == city] if len(summ) else summ
        if len(sm) and pd.notna(sm["best_lag_days (impr follow trends)"].iloc[0]):
            lag = int(sm["best_lag_days (impr follow trends)"].iloc[0])
            if lag > 0:
                I.append(f"Your impressions react about {lag} day{'s' if lag != 1 else ''} after demand moves.")
        for i, text in enumerate(I):
            ins.append({"city": city, "rank": i + 1, "insight": text})

        # ---------- budget board ----------
        spend7 = w(end - timedelta(days=7), end).cost.sum() / 7
        conv30 = c30.conversions.sum()
        fa = sh.get("Forecast_Accuracy", pd.DataFrame())
        fa = fa[(fa.city == city) & (fa.kw == dk)] if len(fa) and "kw" in fa else fa.iloc[0:0] if len(fa) else fa
        err = float(fa.backtest_typical_error_pct.iloc[0]) if len(fa) and pd.notna(fa.backtest_typical_error_pct.iloc[0]) else None
        why, nxt_steps = [], []
        if cpa_chg is not None and cpa_chg >= 25:
            chg = 0
            why.append(f"CPA is {cpa_chg:+.0f}% vs the previous 30 days, so extra spend would be inefficient right now.")
            nxt_steps += ["Review the search terms report and add negatives.", "Check ad strength and landing page.",
                          "Revisit budget once CPA is back near normal."]
        else:
            base_d = n4 if n4 is not None else (d7 * 0.5 if d7 is not None else 0)
            chg = float(np.clip(base_d, -30, 30))
            if n4 is not None:
                why.append(f"Demand over the next 4 weeks is expected {n4:+.0f}% vs a normal week.")
            elif d7 is not None:
                why.append(f"No 12-month outlook yet; demand right now is {d7:+.0f}% vs normal (used at half weight).")
            if b14 is not None and b14 >= 0.10 and chg >= -5:
                add = min(b14 * 100, 30) * 0.5
                chg += add
                why.append(f"{b14:.0%} of eligible impressions are lost to budget, so part of that is added back (+{add:.0f}%).")
            chg = float(np.clip(round(chg / 5) * 5, -30, 40))
            if chg > 0 and n4 is not None and n4 >= 10:
                nxt_steps.append(f"Raise the daily budget by about {chg:.0f}% a few days before the busier period.")
            elif chg > 0:
                nxt_steps.append(f"Raise the daily budget by about {chg:.0f}% to recover impressions lost to budget.")
            elif chg < 0:
                nxt_steps.append("Demand is lower: pick option A (trim budget, protect CPA) or option B "
                                 "(keep budget, relax tCPA to hold share and targets) based on this month’s target.")
            else:
                nxt_steps.append("Keep the daily budget as it is.")
        if r14 is not None and r_base is not None and r14 - r_base >= 0.08:
            bid = "Raise tCPA / bids ~10%"
            why.append(f"Lost to rank rose to {r14:.0%} from a normal {r_base:.0%}: competitors may be bidding up.")
            nxt_steps.append("Raise tCPA or max CPC by about 10%, and check Auction Insights for new competitors.")
        elif n4 is not None and n4 >= 10 and r14 is not None and r14 >= 0.30:
            bid = "Raise tCPA / bids ~10%"
            why.append(f"Demand is rising and {r14:.0%} of impressions are lost to ad rank, so budget alone won’t capture it.")
            nxt_steps.append("Raise tCPA or max CPC by about 10% ahead of the rise.")
        elif n4 is not None and n4 <= -10:
            bid = "Hold bids"
            nxt_steps.append("Hold bids to protect ad rank for when demand returns.")
        else:
            bid = "Hold bids"
        nxt_steps.append("Check the result after 3–7 days: impression share up, CPA stable.")
        # two options: A = protect efficiency, B = protect targets / market share
        a_chg, a_bid = chg, bid
        b_chg = max(chg, 0.0)
        if chg < 0:
            b_bid = "Relax tCPA / raise bids ~10–15%"
            b_note = "Keep the budget and loosen tCPA so the budget still spends; holds share and monthly target, CPA likely rises."
            a_note = "Spend less while searches are fewer; CPA stays protected, but monthly volume may fall short of target."
        elif chg > 0:
            b_bid = "Raise tCPA / bids ~10–15%" if (r14 is not None and r14 >= 0.30) else bid
            if n4 is not None and n4 >= 10:
                b_note = "Raise budget and push bids to take more share while demand is up."
            else:
                b_note = "Raise budget to stop losing impressions to budget, and push bids for extra share."
            a_note = "Raise budget only as far as demand and budget-lost share justify; keep bids for efficiency."
        else:
            b_bid = bid
            a_note = b_note = "Demand is normal; both approaches lead to the same action."
        conf = ("High" if (n4 is not None and err is not None and err <= 10 and b14 is not None) else
                "Medium" if (n4 is not None or d7 is not None) else "Low")
        budget_rows.append({"city": city, "avg_daily_spend_7d": round(spend7), "budget_change_pct": chg,
                            "suggested_daily_budget": round(spend7 * (1 + chg / 100)),
                            "bid_action": bid, "demand_now_pct": d7, "demand_next4w_pct": n4,
                            "impr_share_14d": is14, "lost_to_budget_14d": b14, "lost_to_rank_14d": r14,
                            "cpa_30d_incl_gst": cpa30, "cpa_change_pct": cpa_chg, "conversions_30d": conv30,
                            "confidence": conf, "forecast_error_pct": err,
                            "option_a_budget_pct": a_chg, "option_a_daily_budget": round(spend7 * (1 + a_chg / 100)),
                            "option_a_bids": a_bid, "option_a_note": a_note,
                            "option_b_budget_pct": b_chg, "option_b_daily_budget": round(spend7 * (1 + b_chg / 100)),
                            "option_b_bids": b_bid, "option_b_note": b_note,
                            "needs_am_decision": bool(a_chg != b_chg or a_bid != b_bid),
                            **_pacing(_targets, city, a, end, cpa30, spend7 * (1 + a_chg / 100), spend7 * (1 + b_chg / 100),
                                      a_bid, b_bid, cfg),
                            "why": " ".join(why) if why else "Demand and delivery are normal.",
                            "way_forward": " ".join(f"{i + 1}) {t}" for i, t in enumerate(nxt_steps))})

        rows.append({"city": city, "demand_now_pct": d7, "demand_next4w_pct": n4,
                     "outlook": None if n4 is None else ("Rising" if n4 >= 10 else "Falling" if n4 <= -10 else "Steady"),
                     "impressions_7d_vs_prev_pct": imp_chg, "impr_share_14d": is14,
                     "lost_to_budget_14d": b14, "lost_to_rank_14d": r14,
                     "lost_to_budget_normal": b_base, "lost_to_rank_normal": r_base,
                     "cpa_30d_incl_gst": cpa30, "cpa_change_pct": cpa_chg, "cost_30d_incl_gst": cost30,
                     "conversions_30d": c30.conversions.sum(), "top_action": A[0][1], "action_kind": A[0][0]})
    return {"Overview": pd.DataFrame(rows), "Actions": pd.DataFrame(acts), "Insights": pd.DataFrame(ins),
            "Budget_Board": pd.DataFrame(budget_rows)}


# =====================================================================
# CLOUD MODE (GitHub Actions): refresh, then save dashboard data + Excel into TiDB
# =====================================================================
def save_cache(cfg, sheets):
    payload = {"meta": {"keywords": cfg["trends_keywords"], "decision_keyword": cfg["decision_keyword"],
                        "account": cfg["customer_id"], "filter": cfg["campaign_name_regex"],
                        "gst": cfg["gst_multiplier"], "refresh_time": "05:00 IST",
                        "competitors": [n for n, _ in cfg["competitors"]],
                        "city_geo": cfg["city_to_geo"], "state_names": cfg["state_names"],
                        "tcpa_relax_pct": cfg["tcpa_relax_pct"],
                        "segment_terms": [[sg, t] for sg, t in cfg["segment_trend_terms"]]}}
    if sheets is None:
        payload["empty"] = True
        xbytes = None
    else:
        for k in ["Budget_Board", "Overview", "Actions", "Insights", "Summary", "Weekly", "Daily_Aligned", "DOW_Profile", "DayOfMonth_Profile", "Month_Profile",
                  "Spikes_Tagged", "Notes", "Forecast_Weekly", "Forecast_Daily", "Upcoming_Changes",
                  "Forecast_Accuracy", "Locations", "KP_Monthly", "KP_Keywords", "Trends_Weekly_KW",
                  "Leading_Signal", "Leading_Weekly", "Event_Impact", "Weather_Impact", "Macro_Context",
                  "Segments_City", "Segment_Demand_Weekly", "Competitor_State", "Competitor_Weekly",
                  "Validation_Ads", "Validation_Ads_Weekly", "Validation_KP", "Data_Quality"]:
            payload[k] = _records(sheets.get(k))
        write_xlsx(cfg, sheets)
        xbytes = open(cfg["output_xlsx"], "rb").read()
        if len(xbytes) > 40_000_000:
            print("Excel too large to store; skipping the Excel copy.")
            xbytes = None
    con = db_connect(cfg)
    try:   # recent notes (not errors) the dashboard should show, newest first, one per topic
        since = _STATE.get("run_start", (datetime.now() - timedelta(hours=6)).isoformat(timespec="seconds"))
        wn = q(con, "SELECT run_at, detail FROM run_log WHERE step = 'warn' AND run_at >= %s ORDER BY id DESC LIMIT 30",
               (since,))
        seen, notes = set(), []
        for _, r in wn.iterrows():
            if "no Trends geo" in r.detail:
                continue
            key = r.detail.split(":")[0][:40]
            if key not in seen:
                seen.add(key); notes.append(r.detail)
        payload["meta"]["notices"] = notes[:4]
        payload["meta"]["kp_blocked"] = any("Keyword Planner needs Basic" in n for n in notes) or \
            not q(con, "SELECT job FROM job_runs WHERE note = 'needs Basic access' LIMIT 1").empty
    except Exception:
        pass
    import gzip
    now = datetime.now().isoformat(timespec="seconds")
    blob = gzip.compress(json.dumps(payload, default=str).encode("utf-8"), 6)
    cur = con.cursor()
    # TiDB rows are limited to ~6 MB, so big blobs are stored in 4 MB parts
    for name, data in (("payload_gz", blob), ("xlsx", xbytes)):
        cur.execute("DELETE FROM dashboard_files WHERE name = %s", (name,))
        if data:
            parts = [data[i:i + 4_000_000] for i in range(0, len(data), 4_000_000)]
            for i, part in enumerate(parts):
                cur.execute("INSERT INTO dashboard_files (name, part, data, updated_at) VALUES (%s,%s,%s,%s)",
                            (name, i, part, now))
    # small marker row (older dashboards read this table)
    cur.execute("INSERT INTO dashboard_cache (id, payload, xlsx, updated_at) VALUES ('main', %s, NULL, %s) "
                "ON DUPLICATE KEY UPDATE payload = VALUES(payload), xlsx = NULL, updated_at = VALUES(updated_at)",
                (json.dumps({"meta": payload["meta"], "moved": "dashboard_files"}), now))
    con.close()
    print("Dashboard data saved to TiDB.")


def cloud_run(cfg, backfill=False):
    _STATE["run_start"] = datetime.now().isoformat(timespec="seconds")
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
        run_planning_jobs(cfg)             # 5-year Trends, rental signal, cities, weather, macro first
    except Exception as e:
        print("planning jobs:", e)
    try:
        import pnm_segments
        pnm_segments.run_segment_jobs(cfg)  # ad group segments, segment demand, competitors
    except Exception as e:
        print("segments:", e)
    try:
        run_trends_daily(cfg)              # then the bulkier daily Trends history
    except Exception as e:
        print("daily trends:", e)
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
