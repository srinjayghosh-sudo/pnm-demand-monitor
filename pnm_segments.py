"""Segments, competitors and Google Trends validation for the P&M demand monitor.

Answers three questions raised in review:
  1) Can Google Trends be trusted?  -> Validation_* sheets
       a) our own Google Ads data: eligible searches (impressions / impression share) vs Trends, per city
       b) Keyword Planner ratio test: KP volume ratio between keywords vs the Trends ratio
       c) data quality per state (low-volume states give spiky, zero-heavy Trends lines)
  2) Budget advice as two options (handled in pnm_daily.build_overview)
  3) Sub-segments within a city: Brand / Generic & SKAG / Intercity / Competitor
       -> Segments_City, Segment_Demand_Weekly, Competitor_State, Competitor_Weekly
"""
import glob, io, os, re, time
from datetime import datetime, timedelta, date
import numpy as np
import pandas as pd

SEG_ORDER = ["Brand", "Generic / SKAG", "Intercity", "Competitor"]


def _P():
    import pnm_daily as P
    return P


# ---------------------------------------------------------------- tables
def ensure_tables(con):
    cur = con.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS ads_adgroup_daily (
        date DATE, ad_group_id BIGINT, ad_group_name VARCHAR(255), campaign_id BIGINT, campaign_name VARCHAR(255),
        city VARCHAR(64), segment VARCHAR(32), impressions DOUBLE, clicks DOUBLE, cost DOUBLE, conversions DOUBLE,
        impr_share DOUBLE, rank_lost_is DOUBLE, PRIMARY KEY (date, ad_group_id))""")
    cur.execute("""CREATE TABLE IF NOT EXISTS trends_segment_weekly (
        geo VARCHAR(16), kw VARCHAR(64), week DATE, idx DOUBLE, fetched_at VARCHAR(32), PRIMARY KEY (geo, kw, week))""")
    cur.execute("""CREATE TABLE IF NOT EXISTS trends_competitor_state (
        state VARCHAR(64), kw VARCHAR(64), share DOUBLE, fetched_at VARCHAR(32), PRIMARY KEY (state, kw))""")
    cur.execute("""CREATE TABLE IF NOT EXISTS trends_competitor_weekly (
        geo VARCHAR(16), kw VARCHAR(64), week DATE, idx DOUBLE, fetched_at VARCHAR(32), PRIMARY KEY (geo, kw, week))""")


# ---------------------------------------------------------------- segment rules
def classify(campaign, adgroup, cfg):
    """Segment from ad group + campaign name. First matching rule wins; anything else is Generic / SKAG."""
    # ad group name first (most specific); "porter" only counts there, since campaign names may all carry the brand
    for text, is_camp in ((adgroup or "", False), (campaign or "", True)):
        low = text.lower()
        for seg, words in cfg["segment_rules"]:
            ws = [w for w in words if not (is_camp and w == "porter")]
            if any(re.search(rf"(^|[^a-z]){re.escape(w)}", low) for w in ws):
                return seg
    return "Generic / SKAG"


# ---------------------------------------------------------------- Google Ads: ad group level
def _adgroup_query(cfg, start, end, with_is=True):
    m = ("metrics.impressions, metrics.clicks, metrics.cost_micros, metrics.conversions"
         + (", metrics.search_impression_share, metrics.search_rank_lost_impression_share" if with_is else ""))
    return f"""
        SELECT segments.date, campaign.id, campaign.name, ad_group.id, ad_group.name, {m}
        FROM ad_group
        WHERE segments.date BETWEEN '{start:%Y-%m-%d}' AND '{end:%Y-%m-%d}'
          AND campaign.advertising_channel_type = 'SEARCH' AND campaign.status = 'ENABLED'"""


def pull_adgroups(cfg, con):
    P = _P()
    api = P._API.get("api") or P.AdsREST(cfg)
    P._API["api"] = api
    have = P.q(con, "SELECT MAX(date) AS d FROM ads_adgroup_daily")
    last = pd.to_datetime(have.d.iloc[0]) if len(have) and have.d.iloc[0] is not None else None
    end = date.today() - timedelta(days=1)
    start = (last.date() - timedelta(days=cfg["maturity_days"])) if last is not None else end - timedelta(days=cfg["segment_backfill_days"])
    rx = re.compile(cfg["campaign_name_regex"], re.I)
    rows, with_is = [], True
    for s in pd.date_range(start, end, freq=f"{cfg['chunk_days']}D"):
        e = min(s.date() + timedelta(days=cfg["chunk_days"] - 1), end)
        try:
            res = list(api.stream(_adgroup_query(cfg, s.date(), e, with_is)))
        except RuntimeError as err:
            if with_is and ("search_impression_share" in str(err) or "PROHIBITED" in str(err) or "INVALID" in str(err)):
                with_is = False                      # some accounts reject IS at ad group level: still get the rest
                res = list(api.stream(_adgroup_query(cfg, s.date(), e, False)))
            else:
                raise
        for r in res:
            c, g, m = r.get("campaign", {}), r.get("adGroup", {}), r.get("metrics", {})
            if not rx.search(c.get("name", "")):
                continue
            rows.append((r["segments"]["date"], int(g["id"]), g.get("name", "")[:255], int(c["id"]), c.get("name", "")[:255],
                         P._city_from_name(c.get("name", ""), cfg), classify(c.get("name"), g.get("name"), cfg),
                         float(m.get("impressions", 0)), float(m.get("clicks", 0)),
                         float(m.get("costMicros", 0)) / 1e6 * cfg["gst_multiplier"], float(m.get("conversions", 0)),
                         _num(m.get("searchImpressionShare")), _num(m.get("searchRankLostImpressionShare"))))
    if rows:
        con.cursor().executemany(
            "REPLACE INTO ads_adgroup_daily (date, ad_group_id, ad_group_name, campaign_id, campaign_name, city, segment, "
            "impressions, clicks, cost, conversions, impr_share, rank_lost_is) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", rows)
    # re-classify stored rows (rules may have been edited)
    st = P.q(con, "SELECT DISTINCT ad_group_id, ad_group_name, campaign_name, segment FROM ads_adgroup_daily")
    for _, r in st.iterrows():
        seg = classify(r.campaign_name, r.ad_group_name, cfg)
        if seg != r.segment:
            con.cursor().execute("UPDATE ads_adgroup_daily SET segment = %s WHERE ad_group_id = %s", (seg, int(r.ad_group_id)))
    return len(rows)


def _num(v):
    if v is None:
        return None
    try:
        return float(str(v).replace("<", "").replace(">", "")) if str(v) not in ("", "--") else None
    except ValueError:
        return None


# ---------------------------------------------------------------- Google Trends: segments + competitors
def pull_segment_trends(geo, cfg):
    P = _P()
    terms = [t for _, t in cfg["segment_trend_terms"]]

    def go():
        py = P._trends_client()
        py.build_payload(terms, timeframe="today 12-m", geo=geo)
        d = py.interest_over_time()
        if "isPartial" in d.columns:
            d = d[~d["isPartial"].astype(bool)].drop(columns="isPartial")
        return d.astype(float)
    return P._with_retry(go, f"Trends segments {geo}")


def pull_competitors(cfg):
    P = _P()
    terms = [t for _, t in cfg["competitors"]]

    def by_state():
        py = P._trends_client()
        py.build_payload(terms, timeframe="today 12-m", geo="IN")
        return py.interest_by_region(resolution="REGION", inc_low_vol=True).astype(float)

    def over_time():
        py = P._trends_client()
        py.build_payload(terms, timeframe="today 12-m", geo="IN")
        d = py.interest_over_time()
        if "isPartial" in d.columns:
            d = d[~d["isPartial"].astype(bool)].drop(columns="isPartial")
        return d.astype(float)
    return P._with_retry(by_state, "Trends competitors by state"), P._with_retry(over_time, "Trends competitors over time")


def run_segment_jobs(cfg):
    P = _P()
    con = P.db_connect(cfg)
    ensure_tables(con)
    now = datetime.now().isoformat(timespec="seconds")
    # 1) ad group data (Google Ads API, not rate-limited like Trends)
    try:
        n = pull_adgroups(cfg, con)
        P.log(con, "segments", f"ad group rows: {n}")
    except Exception as e:
        con = P.db_connect(cfg); P.log(con, "warn", f"Segments (ad group data): {str(e)[:250]}")
    # 2) segment demand per state, a few states per run
    cities = [c for c in P.q(con, "SELECT DISTINCT city FROM ads_campaign_daily")["city"] if c != "Unmapped"]
    geos = ["IN"] + sorted({cfg["city_to_geo"][c] for c in cities if c in cfg["city_to_geo"]})
    budget = cfg["segment_states_per_run"]
    for geo in geos:
        if budget <= 0:
            break
        if not P._job_due(con, f"seg:{geo}", cfg["forecast_refresh_days"]):
            continue
        budget -= 1
        try:
            d = pull_segment_trends(geo, cfg)
            con = P.db_connect(cfg); cur = con.cursor()
            cur.execute("DELETE FROM trends_segment_weekly WHERE geo = %s", (geo,))
            cur.executemany("INSERT INTO trends_segment_weekly (geo, kw, week, idx, fetched_at) VALUES (%s,%s,%s,%s,%s)",
                            [(geo, kw, w.strftime("%Y-%m-%d"), float(v), now) for w, row in d.iterrows()
                             for kw, v in row.items() if pd.notna(v)])
            P._job_done(con, f"seg:{geo}", f"{len(d)} weeks")
        except P.TrendsPaused:
            print("[trends] segments paused (rate limit); continues next run")
            break
        except Exception as e:
            con = P.db_connect(cfg); P.log(con, "warn", f"Google Trends segments {geo}: {str(e)[:200]}")
        time.sleep(15)
    # 3) competitors: who leads in each state + trend over time (India)
    if P._job_due(con, "competitors", cfg["forecast_refresh_days"]):
        try:
            st, ot = pull_competitors(cfg)
            con = P.db_connect(cfg); cur = con.cursor()
            cur.execute("DELETE FROM trends_competitor_state")
            cur.executemany("INSERT INTO trends_competitor_state (state, kw, share, fetched_at) VALUES (%s,%s,%s,%s)",
                            [(str(s)[:64], kw, float(v), now) for s, row in st.iterrows() for kw, v in row.items()])
            cur.execute("DELETE FROM trends_competitor_weekly")
            cur.executemany("INSERT INTO trends_competitor_weekly (geo, kw, week, idx, fetched_at) VALUES (%s,%s,%s,%s,%s)",
                            [("IN", kw, w.strftime("%Y-%m-%d"), float(v), now) for w, row in ot.iterrows()
                             for kw, v in row.items() if pd.notna(v)])
            P._job_done(con, "competitors", f"{len(st)} states")
        except Exception as e:
            con = P.db_connect(cfg); P.log(con, "warn", f"Google Trends competitors: {str(e)[:200]}")
    con.close()


# ---------------------------------------------------------------- Keyword Planner manual export (UI download)
def read_kp_exports(cfg):
    """Keyword Planner CSVs downloaded from the Google Ads UI (Tools > Keyword Planner > Get search volume and
    forecasts > Download). Put them in keyword_planner/ named kp_<geo>.csv, e.g. kp_IN.csv, kp_IN-KA.csv.
    Returns rows: geo, keyword, month (YYYY-MM), searches."""
    out = []
    for path in sorted(glob.glob(os.path.join(cfg["kp_manual_dir"], "kp_*.csv"))):
        geo = os.path.basename(path)[3:-4]
        raw = open(path, "rb").read()
        text = None
        for enc in ("utf-16", "utf-8-sig", "utf-8", "latin-1"):
            try:
                text = raw.decode(enc)
                if "Keyword" in text:
                    break
            except Exception:
                continue
        if not text:
            continue
        lines = text.splitlines()
        hdr = next((i for i, l in enumerate(lines) if l.split("\t")[0].strip().strip('"') == "Keyword"
                    or l.split(",")[0].strip().strip('"') == "Keyword"), None)
        if hdr is None:
            continue
        sep = "\t" if "\t" in lines[hdr] else ","
        df = pd.read_csv(io.StringIO("\n".join(lines[hdr:])), sep=sep, dtype=str)
        mcols = [c for c in df.columns if str(c).startswith("Searches:")]
        for _, r in df.iterrows():
            kw = str(r.get("Keyword", "")).strip().lower()
            if not kw or kw == "nan":
                continue
            if mcols:
                for c in mcols:
                    v = pd.to_numeric(str(r[c]).replace(",", ""), errors="coerce")
                    m = pd.to_datetime(c.replace("Searches:", "").strip(), format="%b %Y", errors="coerce")
                    if pd.notna(v) and pd.notna(m):
                        out.append((geo, kw, m.strftime("%Y-%m"), float(v)))
            else:
                v = pd.to_numeric(str(r.get("Avg. monthly searches", "")).replace(",", ""), errors="coerce")
                if pd.notna(v):
                    out.append((geo, kw, "avg", float(v)))
    return pd.DataFrame(out, columns=["geo", "keyword", "month", "searches"])


# ---------------------------------------------------------------- analysis
def _pct(a, b):
    return None if b in (0, None) or pd.isna(b) or pd.isna(a) else round((a / b - 1) * 100, 1)


def _r(a, b):
    m = pd.concat([pd.Series(a), pd.Series(b)], axis=1).dropna()
    if len(m) < 8 or m.iloc[:, 0].std() == 0 or m.iloc[:, 1].std() == 0:
        return None
    return round(float(m.iloc[:, 0].corr(m.iloc[:, 1])), 2)


def build_segment_sheets(cfg, con, ads, trends, sh):
    P = _P()
    ensure_tables(con)
    out = {}
    geo_of = dict(cfg["city_to_geo"])
    end = ads.date.max()

    # ---- 1) segments within each city
    ag = P.q(con, "SELECT * FROM ads_adgroup_daily")
    seg_rows = []
    if len(ag):
        ag["date"] = pd.to_datetime(ag["date"])
        ag["w_is"] = ag.impressions / ag.impr_share.where(ag.impr_share > 0)
        allin = ag.copy(); allin["city"] = "All India"
        ag2 = pd.concat([ag, allin])
        st = P.q(con, "SELECT geo, kw, week, idx FROM trends_segment_weekly")
        st["week"] = pd.to_datetime(st["week"]) if len(st) else st.get("week")
        term_seg = {t: s for s, t in cfg["segment_trend_terms"]}
        for city, g in ag2.groupby("city"):
            if city == "Unmapped":
                continue
            c30 = g[g.date > end - pd.Timedelta(days=30)]
            c14 = g[g.date > end - pd.Timedelta(days=14)]
            tot = c30.cost.sum()
            geo = "IN" if city == "All India" else geo_of.get(city)
            for seg in SEG_ORDER:
                a, b = c30[c30.segment == seg], c14[c14.segment == seg]
                if not len(a):
                    continue
                elig = b.w_is.sum()
                is14 = (b.impressions.sum() / elig) if elig else None
                rl = (b.rank_lost_is * b.w_is).sum() / elig if elig and b.rank_lost_is.notna().any() else None
                dem = None
                if len(st) and geo:
                    terms = [t for t, s in term_seg.items() if s == seg]
                    sg = st[(st.geo == geo) & (st.kw.isin(terms))].groupby("week").idx.sum().sort_index()
                    if len(sg) >= 12 and sg.mean() > 0:
                        dem = _pct(sg.tail(4).mean(), sg.mean())
                conv = a.conversions.sum()
                cpa = a.cost.sum() / conv if conv else None
                seg_rows.append({"city": city, "segment": seg, "cost_30d": round(a.cost.sum()),
                                 "share_of_spend_pct": round(a.cost.sum() / tot * 100, 1) if tot else None,
                                 "impressions_30d": int(a.impressions.sum()), "clicks_30d": int(a.clicks.sum()),
                                 "conversions_30d": round(conv, 1), "cpa_30d": round(cpa) if cpa else None,
                                 "impr_share_14d": round(is14, 3) if is14 else None,
                                 "lost_to_rank_14d": round(rl, 3) if rl is not None else None,
                                 "demand_last4w_vs_12m_pct": dem,
                                 "ad_groups": int(a.ad_group_id.nunique()),
                                 "suggestion": _seg_suggest(seg, dem, is14, rl)})
    out["Segments_City"] = pd.DataFrame(seg_rows)

    # segment demand weekly (per city view uses the city's state)
    st = P.q(con, "SELECT geo, kw, week, idx FROM trends_segment_weekly")
    rows = []
    if len(st):
        term_seg = {t: s for s, t in cfg["segment_trend_terms"]}
        st["segment"] = st.kw.map(term_seg)
        for city in ["All India"] + sorted(c for c in ads.city.unique() if c not in ("All India", "Unmapped")):
            geo = "IN" if city == "All India" else geo_of.get(city)
            s = st[st.geo == geo]
            for (kw, seg), g in s.groupby(["kw", "segment"]):
                for _, r in g.sort_values("week").iterrows():
                    rows.append({"city": city, "segment": seg, "term": kw, "week_start": str(r.week)[:10], "idx": r.idx})
    out["Segment_Demand_Weekly"] = pd.DataFrame(rows)

    # ---- competitors
    cs = P.q(con, "SELECT state, kw, share FROM trends_competitor_state")
    if len(cs):
        name = {t: n for n, t in cfg["competitors"]}
        cs["competitor"] = cs.kw.map(name).fillna(cs.kw)
        pv = cs.pivot_table(index="state", columns="competitor", values="share", aggfunc="sum").fillna(0)
        tot = pv.sum(axis=1)
        sharep = pv.div(tot.where(tot > 0), axis=0) * 100
        res = sharep.round(0)
        def lead(st):
            if tot[st] <= 0:
                return "Not enough data"
            top = sharep.loc[st].sort_values(ascending=False)
            return f"Close: {top.index[0]} / {top.index[1]}" if len(top) > 1 and top.iloc[0] - top.iloc[1] < 3 else top.index[0]
        res["leader"] = [lead(st) for st in pv.index]
        res["leader_share_pct"] = [round(sharep.loc[s].max()) if tot[s] > 0 else None for s in pv.index]
        our_states = {v for k, v in cfg.get("state_names", {}).items()}
        res["porter_runs_here"] = [s in our_states for s in res.index]
        out["Competitor_State"] = res.reset_index().sort_values(["porter_runs_here", "state"], ascending=[False, True])
    else:
        out["Competitor_State"] = pd.DataFrame()
    cw = P.q(con, "SELECT kw, week, idx FROM trends_competitor_weekly")
    if len(cw):
        name = {t: n for n, t in cfg["competitors"]}
        cw["competitor"] = cw.kw.map(name).fillna(cw.kw)
        cw["week_start"] = cw.week.astype(str).str[:10]
        out["Competitor_Weekly"] = cw[["week_start", "competitor", "idx"]]
    else:
        out["Competitor_Weekly"] = pd.DataFrame()

    # ---- Q1a: our own Google Ads data vs Trends
    camp = ads.copy()
    camp["w_is"] = camp.impressions / camp.impr_share.where(camp.impr_share > 0) if "impr_share" in camp else np.nan
    val, valw = [], []
    tw = sh.get("Trends_Weekly_KW", pd.DataFrame())
    dk = cfg["decision_keyword"]
    for city, g in camp.groupby("city"):
        if city == "Unmapped":
            continue
        g = g.set_index("date").sort_index()
        wk = g["w_is"].resample("W-MON", label="left", closed="left").sum(min_count=3)
        if len(tw) and "kw" in tw:
            t = tw[(tw.city == city) & (tw.kw == dk)].copy()
        else:
            t = pd.DataFrame()
        if not len(t):
            val.append({"city": city, "weeks_compared": 0, "match_level_r": None, "match_weekly_change_r": None,
                        "verdict": "Waiting for Trends data"})
            continue
        t["week_start"] = pd.to_datetime(t.week_start)
        ts = t.set_index("week_start").idx
        m = pd.concat([wk.rename("eligible"), ts.rename("trends")], axis=1).dropna().tail(52)
        m = m[m.eligible > 0]
        rl, rc = _r(m.eligible, m.trends), _r(m.eligible.pct_change(), m.trends.pct_change())
        verdict = ("Strong match" if rl is not None and rl >= 0.6 else "Moderate match" if rl is not None and rl >= 0.35
                   else "Weak match" if rl is not None else "Not enough weeks yet")
        val.append({"city": city, "weeks_compared": len(m), "match_level_r": rl, "match_weekly_change_r": rc, "verdict": verdict})
        if len(m):
            e0, t0 = m.eligible.mean(), m.trends.mean()
            for w, r in m.iterrows():
                valw.append({"city": city, "week_start": str(w)[:10], "eligible_searches": round(r.eligible),
                             "eligible_index": round(r.eligible / e0 * 100, 1), "trends_index": round(r.trends / t0 * 100, 1)})
    out["Validation_Ads"] = pd.DataFrame(val)
    out["Validation_Ads_Weekly"] = pd.DataFrame(valw)

    # ---- Q1b: Keyword Planner ratio test (Malay's method)
    kp = read_kp_exports(cfg)
    kp_api = P.q(con, "SELECT 'city' AS src, city, keyword, ym, searches FROM kp_monthly") if True else pd.DataFrame()
    t5 = P.q(con, "SELECT geo, kw, week, idx FROM trends_weekly_5y_kw")
    vk = []
    if len(t5):
        t5["week"] = pd.to_datetime(t5.week)
        t5 = t5[t5.week > t5.week.max() - pd.Timedelta(days=365)]
        t5["ym"] = t5.week.dt.strftime("%Y-%m")
    sources = []
    if len(kp):
        sources.append(("Keyword Planner export", kp.rename(columns={"month": "ym"})))
    if len(kp_api):
        rev = {v: k for k, v in geo_of.items()}
        kk = kp_api.copy()
        kk["geo"] = kk.city.map(lambda c: "IN" if c == "All India" else geo_of.get(c))
        kk = kk.dropna(subset=["geo"])
        sources.append(("Keyword Planner API", kk[["geo", "keyword", "ym", "searches"]]))
    base = cfg["decision_keyword"]
    for src, k in sources:
        for geo, g in k.groupby("geo"):
            tg = t5[t5.geo == geo] if len(t5) else pd.DataFrame()
            gm = g[g.ym != "avg"]
            kb = gm[gm.keyword == base]
            for kw in cfg["trends_keywords"]:
                if kw == base:
                    continue
                kx = gm[gm.keyword == kw]
                months = sorted(set(kb.ym) & set(kx.ym) & (set(tg.ym) if len(tg) else set()))
                if len(months) < 3 or not len(tg):
                    avg_b, avg_x = g[(g.keyword == base)].searches.mean(), g[(g.keyword == kw)].searches.mean()
                    if pd.notna(avg_b) and pd.notna(avg_x) and avg_x > 0 and len(tg):
                        trat = tg[tg.kw == base].idx.mean() / tg[tg.kw == kw].idx.mean() if tg[tg.kw == kw].idx.mean() > 0 else None
                        krat = avg_b / avg_x
                        vk.append(_kp_row(src, geo, kw, base, krat, trat, None, 12))
                    continue
                kbm = kb.set_index("ym").searches.reindex(months)
                kxm = kx.set_index("ym").searches.reindex(months)
                tb = tg[tg.kw == base].groupby("ym").idx.mean().reindex(months)
                tx = tg[tg.kw == kw].groupby("ym").idx.mean().reindex(months)
                krat = kbm.sum() / kxm.sum() if kxm.sum() > 0 else None
                trat = tb.sum() / tx.sum() if tx.sum() > 0 else None
                vk.append(_kp_row(src, geo, kw, base, krat, trat, _r(kbm.values, tb.values), len(months)))
    out["Validation_KP"] = pd.DataFrame(vk)

    # ---- Q1c: data quality per state
    dq = []
    if len(t5):
        rev = {}
        for c, g in geo_of.items():
            rev.setdefault(g, []).append(c)
        for (geo, kw), g in t5.groupby(["geo", "kw"]):
            z = float((g.idx <= 0).mean() * 100)
            cv = float(g.idx.std() / g.idx.mean()) if g.idx.mean() > 0 else None
            verdict = ("Reliable" if z < 5 and (cv is None or cv < 0.6) else
                       "Usable, a bit noisy" if z < 25 else "Too little search volume: use as rough guide only")
            dq.append({"state": geo, "cities": ", ".join(rev.get(geo, ["All India"] if geo == "IN" else [])),
                       "keyword": kw, "weeks": len(g), "zero_weeks_pct": round(z, 1),
                       "week_to_week_swing": round(cv, 2) if cv is not None else None, "verdict": verdict})
    out["Data_Quality"] = pd.DataFrame(dq)
    return out


def _kp_row(src, geo, kw, base, krat, trat, r, months):
    diff = None if not krat or not trat else round((trat / krat - 1) * 100, 1)
    verdict = ("Not enough data" if diff is None else "Matches" if abs(diff) <= 25 else
               "Roughly matches" if abs(diff) <= 50 else "Does not match")
    return {"source": src, "geo": geo, "keyword": kw, "compared_with": base, "months": months,
            "keyword_planner_ratio": round(krat, 1) if krat else None, "trends_ratio": round(trat, 1) if trat else None,
            "difference_pct": diff, "monthly_shape_r": r, "verdict": verdict}


def _seg_suggest(seg, dem, is14, rl):
    if rl is not None and rl >= 0.5:
        return "Losing half or more auctions to ad rank: review bids / tCPA and ad quality."
    if dem is not None and dem >= 10:
        return "Demand rising: make sure budget and bids can capture it."
    if dem is not None and dem <= -10:
        return "Demand falling: choose efficiency (trim) or share (hold budget, relax bids)."
    if seg == "Brand" and is14 is not None and is14 < 0.8:
        return "Brand impression share below 80%: protect brand terms first."
    return "No change needed."
