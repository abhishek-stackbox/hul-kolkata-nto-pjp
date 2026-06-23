import json, os, warnings, hashlib, time
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd

VERLAUF_DB        = "postgresql://postgres:password@localhost/verlauf"
VERLAUF_USER_FILES = os.path.expanduser(
    "~/projects/work/stackbox/verlauf/verlauf-backend/user-files"
)

warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="HUL Kolkata NTO & PJP",
    layout="wide",
    initial_sidebar_state="collapsed",
)
st.markdown("""
<style>
header[data-testid="stHeader"]     { display:none !important; }
footer                              { display:none !important; }
section[data-testid="stSidebar"]   { display:none !important; }
.block-container                    { padding:0 !important; max-width:100% !important; }
[data-testid="stAppViewContainer"] { padding:0 !important; background:#0a1929 !important; }
[data-testid="stMain"]             { padding:0 !important; }
[data-testid="stMainBlockContainer"]{ padding:0 !important; }
[data-testid="stElementContainer"] { padding:0 !important; }
[data-testid="stIFrame"]{
    position:fixed !important; top:0; left:0;
    width:100vw !important; height:100vh !important;
    border:none !important; display:block !important; z-index:9999;
}
</style>
<script>
(function(){
  function fix(){
    document.querySelectorAll('[data-testid="stIFrame"]').forEach(function(f){
      f.style.setProperty('position','fixed','important');
      f.style.setProperty('top','0','important');
      f.style.setProperty('left','0','important');
      f.style.setProperty('width','100vw','important');
      f.style.setProperty('height','100vh','important');
      f.style.setProperty('z-index','9999','important');
      f.style.setProperty('border','none','important');
    });
  }
  fix();
  new MutationObserver(fix).observe(document.body,{childList:true,subtree:true,attributes:true,attributeFilter:['style']});
  window.addEventListener('resize',fix);
})();
</script>
""", unsafe_allow_html=True)

DATA_ROOT = os.path.expanduser(
    "~/Library/CloudStorage/GoogleDrive-abhishek@stackbox.xyz"
    "/My Drive/Clients Self/HUL/Sales Route/Kolkata"
)
DATA_FILE       = f"{DATA_ROOT}/Active_Outlet_Master_Kolkata.xlsx"
ACTIVE_FILE     = f"{DATA_ROOT}/outlet_comparison_v2.xlsx"
ACTIVE_SHEET    = "Final_active_outlets"
PROPOSED_PLAN   = os.path.expanduser("~/Downloads/test2_p2_output.xlsx")
BEATS_390_FILE  = f"{DATA_ROOT}/218390/All_Beat_Designs_218390_V3.xlsx"  # still used for Existing Beats
BEATS_V4_FILE   = f"{DATA_ROOT}/218390/All_Beat_Designs_218390_V4.xlsx"
CACHE_DIR  = os.path.join(os.path.dirname(__file__), "data")
os.makedirs(CACHE_DIR, exist_ok=True)

def _cache_path(name): return os.path.join(CACHE_DIR, f"{name}.parquet")
def _json_path(name):  return os.path.join(CACHE_DIR, f"{name}.json")

def _xlsx_mtime():
    try: return max(os.path.getmtime(DATA_FILE), os.path.getmtime(ACTIVE_FILE))
    except: return 0

def _cache_valid(name):
    j = _json_path(name)
    if not os.path.exists(j):
        return False
    src_mtime = _xlsx_mtime()
    if src_mtime == 0:
        return True   # source files not accessible (cloud) — trust committed JSON
    return os.path.getmtime(j) > src_mtime

RS_COLORS = [
    "#2563eb","#dc2626","#16a34a","#ea580c","#9333ea",
    "#0891b2","#ca8a04","#db2777","#0d9488","#6d28d9",
    "#b45309","#be185d","#0284c7","#15803d","#b91c1c",
    "#7c3aed","#065f46","#1e3a5f","#78350f","#134e4a",
    "#4c1d95","#881337",
]

def _rgb(h):
    h = h.lstrip("#")
    return [int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)]

# Outlets with bad coordinates in the Existing Beats sheet; corrected from V3 Beats
_EX_COORD_FIXES = {
    "HUL-218390P98619": (22.536544, 88.326435),
}

def _fix_ex_coords(df, code_col="Code"):
    for code, (lat, lon) in _EX_COORD_FIXES.items():
        mask = df[code_col] == code
        if mask.any():
            df.loc[mask, "lat"] = lat
            df.loc[mask, "lon"] = lon


def _load_json(name):
    with open(_json_path(name)) as f:
        return json.load(f)

def _save_json(name, data):
    with open(_json_path(name), "w") as f:
        json.dump(data, f)


@st.cache_data
def load():
    if _cache_valid("outlets") and _cache_valid("rs_info") and _cache_valid("stats") and _cache_valid("excl_outlets"):
        try:
            return (
                _load_json("outlets"),
                _load_json("rs_info"),
                _load_json("boundaries"),
                _load_json("stats"),
                _load_json("excl_outlets"),
            )
        except Exception:
            pass

    df_all = pd.read_excel(ACTIVE_FILE, sheet_name=ACTIVE_SHEET, dtype=str)
    excl_raw = df_all[df_all["Exclude (Incorrect Lat-Long)"].notna()].copy()
    df = df_all[df_all["Exclude (Incorrect Lat-Long)"].isna()].copy()

    df["lat"]    = pd.to_numeric(df["Latitude"],    errors="coerce")
    df["lon"]    = pd.to_numeric(df["Longitude"],   errors="coerce")
    df["rs_lat"] = pd.to_numeric(df["RS Latitude"], errors="coerce")
    df["rs_lon"] = pd.to_numeric(df["RS Longitude"],errors="coerce")
    df = df.dropna(subset=["lat","lon"]).copy()
    df["combined_moc"] = pd.to_numeric(df["Combined MOC"], errors="coerce").fillna(0)

    # Proposed RS: from plan output file (Outlet Code → new distributor)
    proposed_rs_map = {}
    if os.path.exists(PROPOSED_PLAN):
        try:
            _plan = pd.read_excel(PROPOSED_PLAN, usecols=["code","distributor"], dtype=str)
            proposed_rs_map = dict(zip(_plan["code"].str.strip(), _plan["distributor"].str.strip()))
        except Exception:
            pass
    df["new_rs"] = df["Outlet Code"].map(proposed_rs_map).fillna(df["RS Code"])

    rs_ll = df.dropna(subset=["rs_lat","rs_lon"]).groupby("RS Code")[["rs_lat","rs_lon"]].first()
    rs_ll_dict = {code: (float(rs_ll.loc[code,"rs_lat"]), float(rs_ll.loc[code,"rs_lon"]))
                  for code in rs_ll.index}

    def _hav(la1, lo1, la2, lo2):
        from math import radians, sin, cos, sqrt, atan2
        R = 6371.0
        dlat = radians(la2-la1); dlon = radians(lo2-lo1)
        a = sin(dlat/2)**2 + cos(radians(la1))*cos(radians(la2))*sin(dlon/2)**2
        return round(R * 2 * atan2(sqrt(a), sqrt(1-a)), 1)

    counts = df.groupby(["RS Code","RS Name","Distributor Type"]).size().reset_index(name="n")
    counts["_ord"] = counts["Distributor Type"].map({"General":0,"WS":1,"Pharma":2}).fillna(3)
    counts = counts.sort_values(["_ord","n"], ascending=[True,False]).drop(columns="_ord")
    prop_counts = df.groupby("new_rs").size().to_dict()
    type_agg = df.groupby(["RS Code","Distributor Type"]).agg(
        n=("Outlet Code","count"), moc=("combined_moc","sum")
    ).reset_index()
    rs_total_moc = df.groupby("RS Code")["combined_moc"].sum().to_dict()
    _WS_CH = {"WHOLESALE", "PHARMA WHOLESALE"}
    df["_biz"] = df["primarychannel"].apply(
        lambda x: "WS" if str(x).strip().upper() in _WS_CH else "General")
    biz_agg = df.groupby(["RS Code","_biz"]).agg(
        n=("Outlet Code","count"), moc=("combined_moc","sum")
    ).reset_index()

    # Per-RS gain/loss for proposed view
    _moved    = df[df["new_rs"] != df["RS Code"]]
    _gained   = _moved.groupby("new_rs").agg(
        gained_n=("Outlet Code","count"), gained_moc=("combined_moc","sum")).to_dict("index")
    _lost     = _moved.groupby("RS Code").agg(
        lost_n=("Outlet Code","count"), lost_moc=("combined_moc","sum")).to_dict("index")
    _prop_moc = df.groupby("new_rs")["combined_moc"].sum().to_dict()

    rs_info, rs_index = [], {}
    for i, row in enumerate(counts.itertuples(index=False)):
        code  = getattr(row, "RS_Code",  None) or row[0]
        name  = getattr(row, "RS_Name",  None) or row[1]
        rtype = getattr(row, "Distributor_Type", None) or row[2]
        n     = row[-1]
        color = RS_COLORS[i % len(RS_COLORS)]
        lat = lon = None
        if code in rs_ll.index:
            lat = float(rs_ll.loc[code, "rs_lat"])
            lon = float(rs_ll.loc[code, "rs_lon"])
        ba    = biz_agg[biz_agg["RS Code"] == code]
        gen_r = ba[ba["_biz"]=="General"]
        ws_r  = ba[ba["_biz"]=="WS"]
        rs_info.append({
            "idx": i, "code": code, "name": name, "type": rtype,
            "lat": lat, "lon": lon,
            "color": color, "rgb": _rgb(color),
            "outlet_count":  int(n),
            "proposed_count":int(prop_counts.get(code, 0)),
            "moc":           round(float(rs_total_moc.get(code, 0))),
            "proposed_moc":  round(float(_prop_moc.get(code, 0)), 2),
            "gained_n":      int(_gained.get(code, {}).get("gained_n", 0)),
            "gained_moc":    round(float(_gained.get(code, {}).get("gained_moc", 0)), 2),
            "lost_n":        int(_lost.get(code, {}).get("lost_n", 0)),
            "lost_moc":      round(float(_lost.get(code, {}).get("lost_moc", 0)), 2),
            "gen_n":         int(gen_r["n"].sum()),
            "gen_moc":       round(float(gen_r["moc"].sum())),
            "ws_n":          int(ws_r["n"].sum()),
            "ws_moc":        round(float(ws_r["moc"].sum())),
        })
        rs_index[code] = i

    rs_col   = list(df.columns).index("RS Code")
    new_col  = list(df.columns).index("new_rs")
    lat_col  = list(df.columns).index("lat")
    lon_col  = list(df.columns).index("lon")
    name_col = list(df.columns).index("Outlet Name")
    cl_col   = list(df.columns).index("Classification")
    moc_col  = list(df.columns).index("combined_moc")
    ch_col   = list(df.columns).index("primarychannel")
    cp_col   = list(df.columns).index("Channel Program")
    oc_col   = list(df.columns).index("Outlet Code")
    outlets = []
    for row in df.itertuples(index=False):
        ri = rs_index.get(row[rs_col], -1)
        if ri < 0:
            continue
        ni = rs_index.get(row[new_col], ri)
        outlets.append([round(row[lat_col],5), round(row[lon_col],5), ri, ni, row[name_col],
                        str(row[cl_col]) if row[cl_col] else "",
                        round(float(row[moc_col]), 2) if row[moc_col] else 0,
                        str(row[ch_col]) if row[ch_col] else "",
                        str(row[cp_col]) if row[cp_col] else "",
                        str(row[oc_col]) if row[oc_col] else ""])

    with open(f"{DATA_ROOT}/rs_boundaries.geojson") as f:
        geo = json.load(f)
    for feat in geo["features"]:
        code = feat["properties"]["RS_CODE"]
        if code in rs_index:
            idx = rs_index[code]
            feat["properties"].update(rs_idx=idx, color=rs_info[idx]["color"],
                rgb=rs_info[idx]["rgb"], rs_type=rs_info[idx]["type"])
        else:
            feat["properties"].update(rs_idx=-1, color="#9ca3af", rgb=[156,163,175], rs_type="Unknown")

    reassigned = int((df["RS Code"] != df["new_rs"]).sum())
    gen_cnt = int(df[df["Distributor Type"]=="General"].shape[0])
    pha_cnt = int(df[df["Distributor Type"]=="Pharma"].shape[0])
    ws_cnt  = int(df[df["Distributor Type"]=="WS"].shape[0])
    gen_rs  = int(df[df["Distributor Type"]=="General"]["RS Code"].nunique())
    pha_rs  = int(df[df["Distributor Type"]=="Pharma"]["RS Code"].nunique())
    stats = {
        "total":      len(outlets),
        "general":    gen_cnt,
        "pharma":     pha_cnt,
        "ws":         ws_cnt,
        "general_rs": gen_rs,
        "pharma_rs":  pha_rs,
        "rs_count":   len(rs_info),
        "reassigned": reassigned,
    }

    excl_raw["lat2"] = pd.to_numeric(excl_raw["Latitude"],  errors="coerce")
    excl_raw["lon2"] = pd.to_numeric(excl_raw["Longitude"], errors="coerce")
    excl_raw = excl_raw.dropna(subset=["lat2","lon2"])
    excl_outlets = []
    for _, r in excl_raw.iterrows():
        rl     = rs_ll_dict.get(r["RS Code"])
        rs_lat = round(rl[0], 5) if rl else None
        rs_lon = round(rl[1], 5) if rl else None
        dist   = _hav(r.lat2, r.lon2, rl[0], rl[1]) if rl else None
        excl_outlets.append([round(r.lat2,5), round(r.lon2,5), r["Outlet Name"], r["RS Code"],
                             rs_lat, rs_lon, dist])

    _save_json("outlets",       outlets)
    _save_json("rs_info",       rs_info)
    _save_json("boundaries",    geo)
    _save_json("stats",         stats)
    _save_json("excl_outlets",  excl_outlets)
    return outlets, rs_info, geo, stats, excl_outlets


@st.cache_data
def load_dupes():
    if _cache_valid("dupe_pairs") and _cache_valid("dupe_stats"):
        try:
            return _load_json("dupe_pairs"), _load_json("dupe_stats")
        except Exception:
            pass

    dupes = pd.read_excel(
        f"{DATA_ROOT}/output/hul_kolkata_validated.xlsx",
        sheet_name="Confirmed Duplicates", dtype=str
    )
    master = pd.read_excel(DATA_FILE, sheet_name="Active Outlets", dtype=str)
    master["lat"] = pd.to_numeric(master["Latitude"],  errors="coerce")
    master["lon"] = pd.to_numeric(master["Longitude"], errors="coerce")
    ll = master.drop_duplicates("Outlet Code").set_index("Outlet Code")[["lat","lon"]]

    pairs, rs_set = [], set()
    for _, r in dupes.iterrows():
        ca, cb = r["code_a"], r["code_b"]
        if ca not in ll.index or cb not in ll.index:
            continue
        a, b = ll.loc[ca], ll.loc[cb]
        if any(pd.isna(x) for x in [a.lat, a.lon, b.lat, b.lon]):
            continue
        dist = round(float(r["dist_m"])) if str(r.get("dist_m","")) not in ("","nan","None") else 0
        pairs.append({
            "na": r["name_a"], "nb": r["name_b"],
            "ca": ca, "cb": cb,
            "rsa": r["rs_code_a"], "rsb": r.get("rs_code_b",""),
            "reason": r.get("ai_reason",""),
            "la":  round(float(a.lat),5), "loa": round(float(a.lon),5),
            "lb":  round(float(b.lat),5), "lob": round(float(b.lon),5),
            "dist": dist,
        })
        rs_set.add(r["rs_code_a"])

    pairs.sort(key=lambda p: p["dist"])
    dupe_stats = {
        "total":  len(pairs),
        "rs_aff": len(rs_set),
        "saved":  max(1, round(len(pairs)/220)),
    }

    _save_json("dupe_pairs", pairs)
    _save_json("dupe_stats", dupe_stats)
    return pairs, dupe_stats


@st.cache_data
def load_clusters():
    if _cache_valid("clusters") and _cache_valid("cluster_stats"):
        try:
            return _load_json("clusters"), _load_json("cluster_stats")
        except Exception:
            pass

    df = pd.read_excel(ACTIVE_FILE, sheet_name=ACTIVE_SHEET, dtype=str)
    df = df[df["Exclude (Incorrect Lat-Long)"].isna()].copy()
    df["lat"] = pd.to_numeric(df["Latitude"],  errors="coerce")
    df["lon"] = pd.to_numeric(df["Longitude"], errors="coerce")
    df = df.dropna(subset=["lat","lon"])

    GRID = 0.0002
    df["glat"] = (df["lat"] / GRID).round() * GRID
    df["glon"] = (df["lon"] / GRID).round() * GRID
    grid = df.groupby(["glat","glon"]).size().reset_index(name="n")
    grid = grid[grid["n"] >= 5].sort_values("n", ascending=False).reset_index(drop=True)

    clusters = [
        {"i": int(i), "lat": round(row.glat,5), "lon": round(row.glon,5), "n": int(row.n)}
        for i, row in grid.iterrows()
    ]
    cluster_stats = {"total": len(clusters), "max_n": int(grid["n"].max())}

    _save_json("clusters", clusters)
    _save_json("cluster_stats", cluster_stats)
    return clusters, cluster_stats


@st.cache_data
def load_beats():
    data_keys  = ["beats_390","beats_391","ex_beats_390","ex_beats_391",
                  "plg_info","dse_info","beat_stats"]
    cache_keys = data_keys + ["beats_v2"]
    if all(_cache_valid(k) for k in cache_keys):
        try:
            return tuple(_load_json(k) for k in data_keys)
        except Exception:
            pass

    PLG_ORDER  = [
        # New SBX PLGs
        "D","D+F","D+F+N","F","F+N","N","PP","PP-A","PP-B",
        # OFM specialists (V3 then V4 naming; deduped)
        "D-OFM","D_OFM","F-OFM","F+N_OFM","N_OFM",
        "PP-A_OFM","PP-B_OFM",
        # UNIGLOW specialists
        "D+F_UNIGLOW","PP-A_UNIGLOW","PP-B_UNIGLOW",
        # Old RS (existing) PLG names
        "DETS","FNB","NUTS","D+F+NUTS","FNB+NUTS","HUL+NUTS",
    ]
    _OFM_PLGS = {"D-OFM","D_OFM","F-OFM","F+N_OFM","N_OFM","PP-A_OFM","PP-B_OFM"}
    _UNI_PLGS = {"D+F_UNIGLOW","PP-A_UNIGLOW","PP-B_UNIGLOW"}
    _EX_PLGS  = {"DETS","FNB","NUTS","D+F+NUTS","FNB+NUTS","HUL+NUTS"}
    PLG_COLORS = {
        "D":"#2563eb","D+F":"#0891b2","D+F+N":"#0d9488",
        "F":"#16a34a","F+N":"#65a30d","N":"#ca8a04",
        "PP":"#dc2626","PP-A":"#ea580c","PP-B":"#9333ea",
        # Specialist Sub PLGs (V3) — use same color as base PLG
        "D-OFM":"#2563eb","F-OFM":"#16a34a","N_OFM":"#ca8a04",
        "D+F_UNIGLOW":"#0891b2","PP-A_OFM":"#ea580c","PP-A_UNIGLOW":"#ea580c",
        "PP-B_OFM":"#9333ea","PP-B_UNIGLOW":"#9333ea",
        # Specialist Sub PLGs (V4)
        "D_OFM":"#2563eb","F+N_OFM":"#65a30d",
        # Existing beat PLG names
        "D+F+NUTS":"#0891b2","DETS":"#2563eb","FNB":"#16a34a",
        "FNB+NUTS":"#65a30d","HUL+NUTS":"#0d9488","NUTS":"#ca8a04",
    }

    def _make_plg_idx(vals):
        seen = {v for v in vals if pd.notna(v)}
        plgs = [p for p in PLG_ORDER if p in seen]
        return {p: i for i, p in enumerate(plgs)}, plgs

    ROOT391 = f"{DATA_ROOT}/218391"

    # ── V4 Proposed 218390 (All_Beat_Designs_218390_V4.xlsx · V4 Beats) ───────
    df390p = pd.read_excel(BEATS_V4_FILE, sheet_name="V4 Beats", dtype=str)
    df390p = df390p.rename(columns={
        "Sub_PLG": "Sub PLG", "Lat": "Latitude", "Lon": "Longitude",
        "DSE_Global": "DSE", "Market (cal day)": "Market",
    })
    df390p["lat"] = pd.to_numeric(df390p["Latitude"], errors="coerce")
    df390p["lon"] = pd.to_numeric(df390p["Longitude"], errors="coerce")
    df390p["Market"] = pd.to_numeric(df390p["Market"], errors="coerce")
    # Use Sub PLG (not PLG) — correctly routes specialist DSE beats to specialist or regular PLG
    df390p["PLG_use"] = df390p["Sub PLG"].where(df390p["Sub PLG"].notna(), df390p["PLG"])
    df390p = df390p.dropna(subset=["lat","lon","PLG_use","Market"]).copy()
    plg_idx390p, _ = _make_plg_idx(df390p["PLG_use"])
    dse_vals390p   = sorted(df390p["DSE"].dropna().unique().tolist())
    dse_idx390p    = {d: i for i, d in enumerate(dse_vals390p)}
    beats_390 = [
        [round(float(r.lat),5), round(float(r.lon),5),
         plg_idx390p.get(r.PLG_use, 0), int(r.Market)-1,
         dse_idx390p.get(str(r.DSE), 0),
         int(float(str(r.Beat))) if str(r.Beat) not in ('','nan','None') else -1]
        for r in df390p.itertuples()
    ]

    # ── Existing 218390 (All_Beat_Designs_218390_V3.xlsx · Existing Beats) ────
    df390e = pd.read_excel(BEATS_390_FILE, sheet_name="Existing Beats", dtype=str)
    df390e["lat"] = pd.to_numeric(df390e["Latitude"], errors="coerce")
    df390e["lon"] = pd.to_numeric(df390e["Longitude"], errors="coerce")
    df390e["Market"] = pd.to_numeric(df390e["Market"], errors="coerce")
    _fix_ex_coords(df390e)
    df390e = df390e.dropna(subset=["lat","lon","PLG","Market"]).copy()
    plg_idx390e, _ = _make_plg_idx(df390e["PLG"])
    dse_vals390e   = sorted(df390e["DSE"].dropna().unique().tolist())
    dse_idx390e    = {d: i for i, d in enumerate(dse_vals390e)}
    ex_beats_390 = [
        [round(float(r.lat),5), round(float(r.lon),5),
         plg_idx390e.get(r.PLG, 0), int(r.Market)-1,
         dse_idx390e.get(str(r.DSE), 0),
         int(float(str(r.Beat))) if str(r.Beat) not in ('','nan','None') else -1]
        for r in df390e.itertuples()
    ]

    # ── Proposed 218391 ────────────────────────────────────────────────────────
    df391p = pd.read_excel(f"{ROOT391}/PLG_Working_218391.xlsx", dtype=str)
    df391p["lat"] = pd.to_numeric(df391p["Latitude"], errors="coerce")
    df391p["lon"] = pd.to_numeric(df391p["Longitude"], errors="coerce")
    df391p = df391p.dropna(subset=["lat","lon"]).copy()
    plg_idx391p, _ = _make_plg_idx(df391p["New_PLGs"])
    beats_391 = [
        [round(float(r.lat),5), round(float(r.lon),5),
         plg_idx391p.get(str(r.New_PLGs).split(",")[0].strip(), -1), -1, -1]
        for r in df391p.itertuples()
    ]

    # ── Existing 218391 ────────────────────────────────────────────────────────
    ex_391_frames = []
    ex_391_dir    = f"{ROOT391}/Existing"
    for fn in sorted(os.listdir(ex_391_dir)):
        if not fn.endswith(".xlsx"): continue
        plg_name = fn[:-5]
        tmp = pd.read_excel(f"{ex_391_dir}/{fn}", dtype=str)
        tmp["lat"] = pd.to_numeric(tmp["Latitude"], errors="coerce")
        tmp["lon"] = pd.to_numeric(tmp["Longitude"], errors="coerce")
        tmp["plg_name"] = plg_name
        ex_391_frames.append(tmp)
    df391e = pd.concat(ex_391_frames, ignore_index=True)
    df391e = df391e.rename(columns={"DSE Code":"dse_code"})
    df391e["lat"] = pd.to_numeric(df391e["Latitude"], errors="coerce")
    df391e["lon"] = pd.to_numeric(df391e["Longitude"], errors="coerce")
    df391e["Market"] = pd.to_numeric(df391e["Market"], errors="coerce")
    df391e = df391e.dropna(subset=["lat","lon","Market"]).copy()
    plg_idx391e, _ = _make_plg_idx(df391e["plg_name"])
    dse_vals391e   = sorted(df391e["dse_code"].dropna().unique().tolist())
    dse_idx391e    = {d: i for i, d in enumerate(dse_vals391e)}
    ex_beats_391 = [
        [round(float(r.lat),5), round(float(r.lon),5),
         plg_idx391e.get(r.plg_name, 0),
         int(r.Market)-1,
         dse_idx391e.get(str(r.dse_code), 0)]
        for r in df391e.itertuples()
    ]

    # ── Build unified PLG info (union of all sources) ──────────────────────────
    all_plg_names = set()
    for idx_map in [plg_idx390p, plg_idx390e, plg_idx391p, plg_idx391e]:
        all_plg_names.update(idx_map.keys())
    plgs     = [p for p in PLG_ORDER if p in all_plg_names]
    def _plg_group(name):
        if name in _OFM_PLGS: return "ofm"
        if name in _UNI_PLGS: return "uniglow"
        if name in _EX_PLGS:  return "existing"
        return "normal"
    plg_info = [{"idx":i,"name":p,"color":PLG_COLORS.get(p,"#6b7280"),"group":_plg_group(p)} for i,p in enumerate(plgs)]
    plg_idx  = {p: i for i, p in enumerate(plgs)}

    def _remap_plg(beats, old_idx):
        inv = {v: k for k, v in old_idx.items()}
        return [[b[0],b[1], plg_idx.get(inv.get(b[2],""), b[2]), b[3], b[4]] + list(b[5:]) for b in beats]

    beats_390    = _remap_plg(beats_390,    plg_idx390p)
    ex_beats_390 = _remap_plg(ex_beats_390, plg_idx390e)
    beats_391    = _remap_plg(beats_391,    plg_idx391p)
    ex_beats_391 = _remap_plg(ex_beats_391, plg_idx391e)

    # ── DSE info: combined from all sources ────────────────────────────────────
    all_dse = sorted(set(dse_vals390p) | set(dse_vals390e))
    dse_idx_all = {d: i for i, d in enumerate(all_dse)}
    dse_info    = [{"idx": i, "name": d} for i, d in enumerate(all_dse)]

    def _remap_dse(beats, old_idx):
        inv = {v: k for k, v in old_idx.items()}
        return [[b[0],b[1],b[2],b[3], dse_idx_all.get(inv.get(b[4],""), b[4])] + list(b[5:]) for b in beats]

    beats_390    = _remap_dse(beats_390,    dse_idx390p)
    ex_beats_390 = _remap_dse(ex_beats_390, dse_idx390e)

    # 391 existing DSE: keep as-is with 391-specific dse_idx_391
    dse_info_391 = [{"idx": i, "name": d} for i, d in enumerate(dse_vals391e)]
    _save_json("dse_info_391", dse_info_391)

    beat_stats = {
        "prop_390":  len(beats_390),
        "prop_391":  len(beats_391),
        "ex_390":    len(ex_beats_390),
        "ex_391":    len(ex_beats_391),
        "plgs":      len(plgs),
        "markets":   6,
    }

    for name, data in [("beats_390",beats_390),("beats_391",beats_391),
                       ("ex_beats_390",ex_beats_390),("ex_beats_391",ex_beats_391),
                       ("plg_info",plg_info),("dse_info",dse_info),("beat_stats",beat_stats),
                       ("beats_v2",{"version":2})]:
        _save_json(name, data)

    return beats_390, beats_391, ex_beats_390, ex_beats_391, plg_info, dse_info, beat_stats


@st.cache_data
def load_benefits():
    from scipy.spatial import ConvexHull

    cache_keys = ["benefit_stats","dse_balance_390",
                  "conflicts_ex_390","conflicts_v3_390",
                  "hull_v3_390","hull_ex_390"]

    def _bv(name):
        j = _json_path(name)
        if not os.path.exists(j): return False
        try:   bm = max(os.path.getmtime(BEATS_390_FILE), os.path.getmtime(BEATS_V4_FILE))
        except: return True
        return os.path.getmtime(j) > bm

    if all(_bv(k) for k in cache_keys):
        try:
            return tuple(_load_json(k) for k in cache_keys)
        except Exception:
            pass

    df_v3 = pd.read_excel(BEATS_V4_FILE, sheet_name="V4 Beats", dtype=str)
    df_v3 = df_v3.rename(columns={
        "Sub_PLG": "Sub PLG", "Lat": "Latitude", "Lon": "Longitude",
        "DSE_Global": "DSE", "Market (cal day)": "Market",
    })
    df_v3["Market"] = pd.to_numeric(df_v3["Market"], errors="coerce")
    df_v3["lat"]    = pd.to_numeric(df_v3["Latitude"], errors="coerce")
    df_v3["lon"]    = pd.to_numeric(df_v3["Longitude"], errors="coerce")
    df_v3["Beat"]   = pd.to_numeric(df_v3["Beat"],     errors="coerce")
    df_v3["PLG"]    = df_v3["Sub PLG"].where(df_v3["Sub PLG"].notna(), df_v3["PLG"])
    df_v3 = df_v3.dropna(subset=["lat","lon","Market","Code","DSE","Beat"]).copy()

    df_ex = pd.read_excel(BEATS_390_FILE, sheet_name="Existing Beats", dtype=str)
    df_ex["Market"] = pd.to_numeric(df_ex["Market"], errors="coerce")
    df_ex["lat"]    = pd.to_numeric(df_ex["Latitude"], errors="coerce")
    df_ex["lon"]    = pd.to_numeric(df_ex["Longitude"], errors="coerce")
    df_ex["Beat"]   = pd.to_numeric(df_ex["Beat"],     errors="coerce")
    _fix_ex_coords(df_ex)
    df_ex = df_ex.dropna(subset=["lat","lon","Market","Code","DSE","Beat"]).copy()

    def _conflicts(df):
        grp = df.groupby(["Code","Market"]).agg(
            dse_count=("DSE","nunique"), lat=("lat","first"), lon=("lon","first")
        ).reset_index()
        c = grp[grp["dse_count"] > 1]
        return [[round(float(r.lat),5), round(float(r.lon),5),
                 int(r.Market)-1, int(r.dse_count)]
                for r in c.itertuples()]

    conflicts_v3 = _conflicts(df_v3)
    conflicts_ex = _conflicts(df_ex)

    def _hulls(df):
        import numpy as np
        result = []
        for (plg, dse, market), sub in df.groupby(["PLG","DSE","Market"]):
            pts = sub[["lat","lon"]].values
            if len(pts) < 3:
                continue
            try:
                h = ConvexHull(pts)
                verts = pts[h.vertices].tolist()
                verts.append(verts[0])
                result.append({"plg":str(plg),"dse":str(dse),"market":int(market),"n":len(pts),
                                "hull":[[round(p[0],5),round(p[1],5)] for p in verts]})
            except Exception:
                pass
        return result

    hull_v3 = _hulls(df_v3)
    hull_ex  = _hulls(df_ex)

    try:
        df_dse = pd.read_excel(BEATS_V4_FILE, sheet_name="V4 DSE Summary")
        df_dse.columns = ["dse","sub_plg","outlets","beats","avg_dist","avg_area","jaccard","moc"]
        dse_balance = []
        for _, row in df_dse.dropna(subset=["dse"]).iterrows():
            dse_balance.append({
                "dse":   str(row.dse),
                "plg":   str(row.sub_plg),
                "n":     int(row.outlets)  if pd.notna(row.outlets)  else 0,
                "beats": int(row.beats)    if pd.notna(row.beats)    else 0,
                "dist":  round(float(row.avg_dist),2) if pd.notna(row.avg_dist) else 0,
                "area":  round(float(row.avg_area),2) if pd.notna(row.avg_area) else 0,
                "moc":   round(float(row.moc),2)      if pd.notna(row.moc)      else 0,
            })
    except Exception:
        dse_balance = []

    benefit_stats = {
        "same_day": {"ex_outlets":5436,"ex_occ":8171,"v3_outlets":1961,"v3_occ":2192},
        "plg_purity": {
            "ex_total_dse":107,"ex_impure":57,"ex_pure":50,
            "v3_total_dse":108,"v3_impure":0,"v3_pure":108,
            "impure_examples":[
                {"dse":"SMN00119","plgs":"PP, PP-A, PP-B","n":3},
                {"dse":"SMN00016","plgs":"PP, PP-A, PP-B","n":3},
                {"dse":"SMN00036","plgs":"D+F+NUTS, DETS, HUL+NUTS","n":3},
                {"dse":"SMN00056","plgs":"PP, PP-A, PP-B","n":3},
                {"dse":"SMN00060","plgs":"PP, PP-A, PP-B","n":3},
                {"dse":"SMN00134","plgs":"PP, PP-A, PP-B","n":3},
                {"dse":"SMN00129","plgs":"PP, PP-A, PP-B","n":3},
                {"dse":"SMN00049","plgs":"PP, PP-A, PP-B","n":3},
            ],
        },
        "balance": {
            "ex_cv":24.9,"v3_cv":15.1,"ex_avg":28,"v3_avg":27,
            "by_plg":[
                {"ex_plg":"DETS",    "v3_plg":"D",    "ex_avg":29.6,"ex_cv":22.5,"v3_avg":22.2,"v3_cv":20.2},
                {"ex_plg":"FNB",     "v3_plg":"F",    "ex_avg":24.5,"ex_cv":33.0,"v3_avg":21.0,"v3_cv":33.4},
                {"ex_plg":"NUTS",    "v3_plg":"N",    "ex_avg":24.8,"ex_cv":32.9,"v3_avg":24.6,"v3_cv":13.0},
                {"ex_plg":"FNB+NUTS","v3_plg":"F+N",  "ex_avg":27.6,"ex_cv":27.4,"v3_avg":30.0,"v3_cv":15.4},
                {"ex_plg":"D+F+NUTS","v3_plg":"D+F+N","ex_avg":27.2,"ex_cv":21.7,"v3_avg":28.3,"v3_cv":10.9},
                {"ex_plg":"HUL+NUTS","v3_plg":"D+F+N","ex_avg":26.0,"ex_cv":21.6,"v3_avg":28.3,"v3_cv":10.9},
                {"ex_plg":"PP",      "v3_plg":"PP",   "ex_avg":29.4,"ex_cv":21.2,"v3_avg":28.6,"v3_cv":12.6},
                {"ex_plg":"PP-A",    "v3_plg":"PP-A", "ex_avg":25.0,"ex_cv":26.8,"v3_avg":25.1,"v3_cv":6.7},
                {"ex_plg":"PP-B",    "v3_plg":"PP-B", "ex_avg":25.2,"ex_cv":26.1,"v3_avg":25.1,"v3_cv":6.7},
            ],
        },
        "jaccard": {
            "by_plg":[
                {"ex_plg":"DETS",    "v3_plg":"D",    "ex_beats":150,"ex_jac":0.3660,"v3_beats":54, "v3_jac":0.0},
                {"ex_plg":"FNB",     "v3_plg":"F",    "ex_beats":12, "ex_jac":0.0216,"v3_beats":12, "v3_jac":0.0},
                {"ex_plg":"NUTS",    "v3_plg":"N",    "ex_beats":12, "ex_jac":0.0214,"v3_beats":24, "v3_jac":0.0001},
                {"ex_plg":"FNB+NUTS","v3_plg":"F+N",  "ex_beats":152,"ex_jac":0.4508,"v3_beats":33, "v3_jac":0.0},
                {"ex_plg":"D+F+NUTS","v3_plg":"D+F+N","ex_beats":39, "ex_jac":0.2572,"v3_beats":198,"v3_jac":0.0},
                {"ex_plg":"HUL+NUTS","v3_plg":"D+F+N","ex_beats":43, "ex_jac":0.0990,"v3_beats":198,"v3_jac":0.0},
                {"ex_plg":"PP",      "v3_plg":"PP",   "ex_beats":134,"ex_jac":0.5168,"v3_beats":198,"v3_jac":0.0},
                {"ex_plg":"PP-A",    "v3_plg":"PP-A", "ex_beats":50, "ex_jac":0.1133,"v3_beats":48, "v3_jac":0.0},
                {"ex_plg":"PP-B",    "v3_plg":"PP-B", "ex_beats":50, "ex_jac":0.1194,"v3_beats":48, "v3_jac":0.0},
            ],
        },
    }

    for k, v in [("benefit_stats",benefit_stats),("dse_balance_390",dse_balance),
                  ("conflicts_ex_390",conflicts_ex),("conflicts_v3_390",conflicts_v3),
                  ("hull_v3_390",hull_v3),("hull_ex_390",hull_ex)]:
        _save_json(k, v)

    return benefit_stats, dse_balance, conflicts_ex, conflicts_v3, hull_v3, hull_ex


@st.cache_data
def load_rs_hulls():
    from scipy.spatial import ConvexHull
    import numpy as np
    from math import radians, sin, cos, sqrt, atan2

    cache_keys = ["hull_rs_ex", "hull_rs_prop", "rs_dist_stats"]
    if all(os.path.exists(_json_path(k)) for k in cache_keys):
        try:
            return tuple(_load_json(k) for k in cache_keys)
        except Exception:
            pass

    outlets_data  = _load_json("outlets")
    rs_info_data  = _load_json("rs_info")
    SAFOAN_IDX    = next((r["idx"] for r in rs_info_data if "safoan" in r["name"].lower()), -1)

    def hav(la1, lo1, la2, lo2):
        R = 6371.0
        dlat = radians(la2 - la1); dlon = radians(lo2 - lo1)
        a = sin(dlat/2)**2 + cos(radians(la1)) * cos(radians(la2)) * sin(dlon/2)**2
        return R * 2 * atan2(sqrt(a), sqrt(1 - a))

    ex_coords, prop_coords = {}, {}
    ex_dists,  prop_dists  = {}, {}

    for o in outlets_data:
        lat, lon, ex_idx, prop_idx = o[0], o[1], o[2], o[3]
        ex_coords.setdefault(ex_idx, []).append([lat, lon])
        prop_coords.setdefault(prop_idx, []).append([lat, lon])
        rs_e = rs_info_data[ex_idx]   if 0 <= ex_idx   < len(rs_info_data) else None
        rs_p = rs_info_data[prop_idx] if 0 <= prop_idx < len(rs_info_data) else None
        if rs_e and rs_e["lat"] and rs_e["lon"]:
            ex_dists.setdefault(ex_idx, []).append(hav(lat, lon, rs_e["lat"], rs_e["lon"]))
        if rs_p and rs_p["lat"] and rs_p["lon"] and ex_idx != SAFOAN_IDX:
            prop_dists.setdefault(prop_idx, []).append(hav(lat, lon, rs_p["lat"], rs_p["lon"]))

    def _hulls(coords_dict):
        out = []
        for rs_idx, pts in coords_dict.items():
            if len(pts) < 3:
                continue
            arr = np.array(pts)
            try:
                h = ConvexHull(arr)
                verts = arr[h.vertices].tolist()
                verts.append(verts[0])
                out.append({"rs_idx": rs_idx,
                             "points": [[round(p[0], 5), round(p[1], 5)] for p in verts]})
            except Exception:
                pass
        return out

    hull_rs_ex   = _hulls(ex_coords)
    hull_rs_prop = _hulls(prop_coords)

    rs_dist_stats = {}
    for rs in rs_info_data:
        idx = rs["idx"]
        if idx == SAFOAN_IDX:
            continue
        ed = ex_dists.get(idx, [])
        pd_ = prop_dists.get(idx, [])
        rs_dist_stats[str(idx)] = {
            "ex":   round(sum(ed)  / len(ed),  2) if ed  else None,
            "prop": round(sum(pd_) / len(pd_), 2) if pd_ else None,
        }

    for k, v in [("hull_rs_ex", hull_rs_ex), ("hull_rs_prop", hull_rs_prop),
                 ("rs_dist_stats", rs_dist_stats)]:
        _save_json(k, v)

    return hull_rs_ex, hull_rs_prop, rs_dist_stats


@st.cache_data
def load_rs_overlap():
    from shapely.geometry import Polygon
    from shapely.ops import unary_union

    cache_key = "rs_overlap"
    if os.path.exists(_json_path(cache_key)):
        try:
            return _load_json(cache_key)
        except Exception:
            pass

    hull_rs_ex   = _load_json("hull_rs_ex")
    hull_rs_prop = _load_json("hull_rs_prop")
    rs_info_data = _load_json("rs_info")
    rs_by_idx    = {r["idx"]: r for r in rs_info_data}

    def _overlap_pct(hulls, ter_type):
        polys = []
        for h in hulls:
            r = rs_by_idx.get(h["rs_idx"])
            if not r or r["type"] != ter_type:
                continue
            pts = [(p[1], p[0]) for p in h["points"]]
            if len(pts) >= 3:
                polys.append(Polygon(pts))
        if not polys:
            return None
        union_area = unary_union(polys).area
        if union_area == 0:
            return None
        overlap = sum(p.area for p in polys) - union_area
        return round(overlap / union_area * 100, 1)

    result = {}
    for ter_type in ["General", "Pharma"]:
        result[ter_type] = {
            "ex":   _overlap_pct(hull_rs_ex,   ter_type),
            "prop": _overlap_pct(hull_rs_prop,  ter_type),
        }

    _save_json(cache_key, result)
    return result


@st.cache_data
def load_beat_distances():
    cache_key = "beat_distances"
    if os.path.exists(_json_path(cache_key)):
        try:
            return _load_json(cache_key)
        except Exception:
            pass

    df_v3 = pd.read_excel(BEATS_390_FILE, sheet_name="V3 Beats", dtype=str)
    df_ex = pd.read_excel(BEATS_390_FILE, sheet_name="Existing Beats", dtype=str)
    for df in [df_v3, df_ex]:
        df["Market"]        = pd.to_numeric(df["Market"],         errors="coerce")
        df["Road Dist (km)"]= pd.to_numeric(df["Road Dist (km)"],errors="coerce")
    # V3: use Sub PLG so specialist beats group separately from regular PLG beats
    df_v3["PLG"] = df_v3["Sub PLG"].where(df_v3["Sub PLG"].notna(), df_v3["PLG"])
    df_v3 = df_v3.dropna(subset=["Market","PLG","DSE"]).copy()
    df_ex = df_ex.dropna(subset=["Market","PLG","DSE"]).copy()

    def _all(df):
        out = []
        for (plg, dse, market), sub in df.groupby(["PLG","DSE","Market"]):
            # Road Dist (km) is per-beat (same value for all outlets); take first non-null
            rd = sub["Road Dist (km)"].dropna()
            chain = round(float(rd.iloc[0]), 2) if len(rd) else None
            out.append({"plg":str(plg),"dse":str(dse),"market":int(market),
                        "n":len(sub),"chain_km":chain,"route_km":None})
        return out

    result = {"v3": _all(df_v3), "ex": _all(df_ex)}
    _save_json(cache_key, result)
    return result


@st.cache_data
def load_beat_areas():
    if os.path.exists(_json_path("beat_areas")):
        try:
            return _load_json("beat_areas")
        except Exception:
            pass

    df_v3 = pd.read_excel(BEATS_390_FILE, sheet_name="V3 Beats", dtype=str)
    df_ex = pd.read_excel(BEATS_390_FILE, sheet_name="Existing Beats", dtype=str)
    df_v3["Market"] = pd.to_numeric(df_v3["Market"], errors="coerce")
    df_v3["area"]   = pd.to_numeric(df_v3["Beat Area km²"], errors="coerce")
    df_v3["PLG"]    = df_v3["Sub PLG"].where(df_v3["Sub PLG"].notna(), df_v3["PLG"])
    df_ex["Market"] = pd.to_numeric(df_ex["Market"], errors="coerce")
    df_ex["area"]   = pd.to_numeric(df_ex["Beat Area km²"], errors="coerce")

    SPEC = {"D-OFM","F-OFM","N_OFM","D+F_UNIGLOW","PP-A_OFM","PP-A_UNIGLOW","PP-B_OFM","PP-B_UNIGLOW"}
    df_v3_u = df_v3.dropna(subset=["Market","area","PLG"]).drop_duplicates(subset=["PLG","DSE","Market"])
    df_ex_u = df_ex.dropna(subset=["Market","area","PLG"]).drop_duplicates(subset=["PLG","DSE","Market"])
    df_reg  = df_v3_u[~df_v3_u["PLG"].isin(SPEC)]
    df_spec = df_v3_u[df_v3_u["PLG"].isin(SPEC)]

    def _by_market(df):
        return {str(int(m)): round(float(a), 1)
                for m, a in df.groupby("Market")["area"].sum().items()}

    df_v3_raw = df_v3.dropna(subset=["PLG"])
    plg_counts = (df_v3_raw.groupby("PLG")
                  .agg(outlets=("Code","count"), dses=("DSE","nunique"))
                  .reset_index())
    plg_summary = [{"plg": str(r.PLG), "outlets": int(r.outlets), "dses": int(r.dses)}
                   for r in plg_counts.itertuples()]

    result = {
        "v3_regular":    _by_market(df_reg),
        "v3_specialist": _by_market(df_spec),
        "ex":            _by_market(df_ex_u),
        "plg_summary":   plg_summary,
    }
    _save_json("beat_areas", result)
    return result


@st.cache_data
def load_delivery_zones():
    cache_key = "delivery_zones"
    if os.path.exists(_json_path(cache_key)):
        try:
            return _load_json(cache_key)
        except Exception:
            pass

    from shapely.geometry import MultiPoint
    import math
    LAT_KM = 111.0
    LON_KM = 111.0 * math.cos(math.radians(22.5))

    # Market → Group A calendar day (D+F+N anchors)
    DFN_CAL = {1:5, 2:2, 3:4, 4:1, 5:3, 6:6}
    REV_DFN = {v:k for k,v in DFN_CAL.items()}
    GROUP_A = {"D+F+N","D","D+F","F","PP-A","D_OFM","PP-A_OFM","PP-A_UNIGLOW","D+F_UNIGLOW"}
    LAT_MIN, LAT_MAX, LON_MIN, LON_MAX = 22.35, 22.60, 88.20, 88.42

    def to_hull(pts):
        if len(pts) < 3:
            return pts
        mp = MultiPoint([(lon*LON_KM, lat*LAT_KM) for lat, lon in pts])
        h = mp.convex_hull
        if h.geom_type == 'Point':
            return [list(pts[0])]
        return [[round(y/LAT_KM,5), round(x/LON_KM,5)] for x,y in h.exterior.coords]

    def area_km2(pts):
        if len(pts) < 3:
            return 0.0
        mp = MultiPoint([(lon*LON_KM, lat*LAT_KM) for lat, lon in pts])
        return round(mp.convex_hull.area, 1)

    # Load V4 from beats_v4.json
    beats_v4_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "beats_v4.json")
    with open(beats_v4_path, encoding="utf-8") as f:
        bv4 = json.load(f)
    cols = bv4["cols"]
    ci = {c:i for i,c in enumerate(cols)}
    from collections import defaultdict
    v4_zone = defaultdict(list)
    for row in bv4["rows"]:
        sub_plg = row[ci["Sub_PLG"]]
        cal_day = int(row[ci["market"]])
        try:
            lat, lon = float(row[ci["latitude"]]), float(row[ci["longitude"]])
        except (TypeError, ValueError):
            continue
        if not (LAT_MIN <= lat <= LAT_MAX and LON_MIN <= lon <= LON_MAX):
            continue
        if sub_plg in GROUP_A:
            orig = REV_DFN.get(cal_day)
        else:
            ga = (cal_day - 2) % 6 + 1
            orig = REV_DFN.get(ga)
        if orig:
            v4_zone[orig].append((lat, lon))

    # Existing beats by market
    df_ex = pd.read_excel(BEATS_390_FILE, sheet_name="Existing Beats", dtype=str)
    df_ex["lat"] = pd.to_numeric(df_ex["Latitude"], errors="coerce")
    df_ex["lon"] = pd.to_numeric(df_ex["Longitude"], errors="coerce")
    df_ex["mkt"] = pd.to_numeric(df_ex["Market"], errors="coerce")
    df_ex = df_ex.dropna(subset=["lat","lon","mkt"]).drop_duplicates(subset=["Code","mkt"])
    df_ex = df_ex[(df_ex["lat"]>=LAT_MIN)&(df_ex["lat"]<=LAT_MAX)&(df_ex["lon"]>=LON_MIN)&(df_ex["lon"]<=LON_MAX)]
    ex_zone = defaultdict(list)
    for _, row in df_ex.iterrows():
        m = int(row["mkt"])
        ex_zone[m].append((row["lat"], row["lon"]))

    zones = []
    for m in range(1,7):
        pts_v4 = v4_zone[m]
        pts_ex  = ex_zone[m]
        ga = DFN_CAL[m]
        gb = ga % 6 + 1
        zones.append({
            "zone": m,
            "group_a_day": ga,
            "group_b_day": gb,
            "v4_area":  area_km2(pts_v4),
            "ex_area":  area_km2(pts_ex),
            "v4_hull":  to_hull(pts_v4),
            "ex_hull":  to_hull(pts_ex),
        })
    result = {"zones": zones}
    _save_json(cache_key, result)
    return result


@st.cache_data
def load_delivery_beats_v4():
    cache_key = "delivery_beats_v4"
    if os.path.exists(_json_path(cache_key)):
        try:
            return _load_json(cache_key)
        except Exception:
            pass

    import math
    from collections import defaultdict
    from scipy.spatial import ConvexHull
    import numpy as np

    LAT_KM = 111.0
    LON_KM = 111.0 * math.cos(math.radians(22.5))

    beats_v4_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "beats_v4.json")
    with open(beats_v4_path, encoding="utf-8") as f:
        bv4 = json.load(f)
    cols = bv4["cols"]
    ci = {c: i for i, c in enumerate(cols)}

    groups = defaultdict(list)
    for row in bv4["rows"]:
        plg = row[ci["PLG"]]
        dse = row[ci["DSE_Global"]]
        try:
            market = int(row[ci["market"]])
            lat = float(row[ci["latitude"]])
            lon = float(row[ci["longitude"]])
        except (TypeError, ValueError):
            continue
        groups[(plg, dse, market)].append([lat, lon])

    def to_hull(pts):
        if len(pts) < 3:
            return [[round(p[0], 5), round(p[1], 5)] for p in pts]
        try:
            pts_km = np.array([[p[1] * LON_KM, p[0] * LAT_KM] for p in pts])
            hull = ConvexHull(pts_km)
            return [[round(pts_km[v][1] / LAT_KM, 5), round(pts_km[v][0] / LON_KM, 5)] for v in hull.vertices]
        except Exception:
            return [[round(p[0], 5), round(p[1], 5)] for p in pts[:4]]

    result = []
    for (plg, dse, market), pts in sorted(groups.items()):
        centroid = [round(sum(p[0] for p in pts) / len(pts), 5),
                    round(sum(p[1] for p in pts) / len(pts), 5)]
        result.append({"plg": plg, "dse": dse, "market": market, "n": len(pts),
                        "centroid": centroid, "hull": to_hull(pts)})

    _save_json(cache_key, result)
    return result


outlets, rs_info, boundaries, stats, excl_outlets = load()
dupe_pairs, dupe_stats                            = load_dupes()
clusters, cluster_stats                           = load_clusters()
beats_390, beats_391, ex_beats_390, ex_beats_391, plg_info, dse_info, beat_stats = load_beats()
dse_info_391 = _load_json("dse_info_391") if os.path.exists(_json_path("dse_info_391")) else []
_, dse_balance_390, conflicts_ex_390, conflicts_v3_390, hull_v3_390, hull_ex_390 = load_benefits()
# Always use hardcoded values — @st.cache_data can serve stale JSON for benefit_stats
benefit_stats = {
    "same_day": {"ex_outlets":5436,"ex_occ":8171,"v3_outlets":1961,"v3_occ":2192},
    "plg_purity": {
        "ex_total_dse":107,"ex_impure":57,"ex_pure":50,
        "v3_total_dse":108,"v3_impure":0,"v3_pure":108,
        "impure_examples":[
            {"dse":"SMN00119","plgs":"PP, PP-A, PP-B","n":3},
            {"dse":"SMN00016","plgs":"PP, PP-A, PP-B","n":3},
            {"dse":"SMN00036","plgs":"D+F+NUTS, DETS, HUL+NUTS","n":3},
            {"dse":"SMN00056","plgs":"PP, PP-A, PP-B","n":3},
            {"dse":"SMN00060","plgs":"PP, PP-A, PP-B","n":3},
            {"dse":"SMN00134","plgs":"PP, PP-A, PP-B","n":3},
            {"dse":"SMN00129","plgs":"PP, PP-A, PP-B","n":3},
            {"dse":"SMN00049","plgs":"PP, PP-A, PP-B","n":3},
        ],
    },
    "balance": {
        "ex_cv":24.9,"v3_cv":15.1,"ex_avg":28,"v3_avg":27,
        "by_plg":[
            {"ex_plg":"DETS",    "v3_plg":"D",    "ex_avg":29.6,"ex_cv":22.5,"v3_avg":22.2,"v3_cv":20.2},
            {"ex_plg":"FNB",     "v3_plg":"F",    "ex_avg":24.5,"ex_cv":33.0,"v3_avg":21.0,"v3_cv":33.4},
            {"ex_plg":"NUTS",    "v3_plg":"N",    "ex_avg":24.8,"ex_cv":32.9,"v3_avg":24.6,"v3_cv":13.0},
            {"ex_plg":"FNB+NUTS","v3_plg":"F+N",  "ex_avg":27.6,"ex_cv":27.4,"v3_avg":30.0,"v3_cv":15.4},
            {"ex_plg":"D+F+NUTS","v3_plg":"D+F+N","ex_avg":27.2,"ex_cv":21.7,"v3_avg":28.3,"v3_cv":10.9},
            {"ex_plg":"HUL+NUTS","v3_plg":"D+F+N","ex_avg":26.0,"ex_cv":21.6,"v3_avg":28.3,"v3_cv":10.9},
            {"ex_plg":"PP",      "v3_plg":"PP",   "ex_avg":29.4,"ex_cv":21.2,"v3_avg":28.6,"v3_cv":12.6},
            {"ex_plg":"PP-A",    "v3_plg":"PP-A", "ex_avg":25.0,"ex_cv":26.8,"v3_avg":25.1,"v3_cv":6.7},
            {"ex_plg":"PP-B",    "v3_plg":"PP-B", "ex_avg":25.2,"ex_cv":26.1,"v3_avg":25.1,"v3_cv":6.7},
        ],
    },
    "jaccard": {
        "by_plg":[
            {"ex_plg":"DETS",    "v3_plg":"D",    "ex_beats":150,"ex_jac":0.3660,"v3_beats":54, "v3_jac":0.0},
            {"ex_plg":"FNB",     "v3_plg":"F",    "ex_beats":12, "ex_jac":0.0216,"v3_beats":12, "v3_jac":0.0},
            {"ex_plg":"NUTS",    "v3_plg":"N",    "ex_beats":12, "ex_jac":0.0214,"v3_beats":24, "v3_jac":0.0001},
            {"ex_plg":"FNB+NUTS","v3_plg":"F+N",  "ex_beats":152,"ex_jac":0.4508,"v3_beats":33, "v3_jac":0.0},
            {"ex_plg":"D+F+NUTS","v3_plg":"D+F+N","ex_beats":39, "ex_jac":0.2572,"v3_beats":198,"v3_jac":0.0},
            {"ex_plg":"HUL+NUTS","v3_plg":"D+F+N","ex_beats":43, "ex_jac":0.0990,"v3_beats":198,"v3_jac":0.0},
            {"ex_plg":"PP",      "v3_plg":"PP",   "ex_beats":134,"ex_jac":0.5168,"v3_beats":198,"v3_jac":0.0},
            {"ex_plg":"PP-A",    "v3_plg":"PP-A", "ex_beats":50, "ex_jac":0.1133,"v3_beats":48, "v3_jac":0.0},
            {"ex_plg":"PP-B",    "v3_plg":"PP-B", "ex_beats":50, "ex_jac":0.1194,"v3_beats":48, "v3_jac":0.0},
        ],
    },
}
hull_rs_ex, hull_rs_prop, rs_dist_stats = load_rs_hulls()
rs_overlap = load_rs_overlap()
_flagged_pharma = _load_json("flagged_pharma_outlets") or []
beat_distances   = load_beat_distances()
beat_areas       = load_beat_areas()
delivery_zones   = load_delivery_zones()
delivery_beats_v4 = load_delivery_beats_v4()

_delivery_data = {}
_delivery_json = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "delivery_data.json")
try:
    with open(_delivery_json, encoding="utf-8") as _f:
        _delivery_data = json.load(_f)
except Exception:
    pass

_beats_v4 = {}
_beats_v4_json = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "beats_v4.json")
try:
    with open(_beats_v4_json, encoding="utf-8") as _f:
        _beats_v4 = json.load(_f)
except Exception:
    pass

_KEEP = ['id','sub_id','sellers','seller_ids','plgs','outlets','value','truck','truck_color','cost','round_trip','centroid','hull']
_delivery_data_slim = {}
for _sc in ['Existing', 'Output 1']:
    _delivery_data_slim[_sc] = {}
    for _lim in ['Max 2 sellers', 'Max 3 sellers', 'Max 4 sellers']:
        _delivery_data_slim[_sc][_lim] = {}
        for _day in _delivery_data.get(_sc, {}).get(_lim, {}):
            _delivery_data_slim[_sc][_lim][_day] = [{k: b[k] for k in _KEEP if k in b} for b in _delivery_data[_sc][_lim][_day]]

# Jun2026 aligned beats — written by salesBeatGuru/HUL-KOLKATA/218390/build_jun26_app_data.py
_jun26_beats     = _load_json("beats_jun26")            or []
_jun26_plg       = _load_json("plg_info_jun26")         or []
_jun26_dse       = _load_json("dse_info_jun26")         or []
_jun26_hull      = _load_json("hull_jun26")             or []
_jun26_conflicts = _load_json("conflicts_jun26")        or []
_jun26_zones     = _load_json("delivery_zones_jun26")   or {}
_jun26_delivery  = _load_json("delivery_beats_jun26")   or {}
_jun26_trucks    = _load_json("truck_assignments_jun26") or {}
_jun26_dist      = _load_json("distances_jun26") or []

# Existing (pre-redesign) 218390+20B801 — written by build_existing_app_data.py
_ex_beats        = _load_json("existing_beats_jun26")            or []
_ex_plg          = _load_json("existing_plg_info_jun26")         or []
_ex_dse          = _load_json("existing_dse_info_jun26")         or []
_ex_hull         = _load_json("existing_hull_jun26")             or []
_ex_dist         = _load_json("existing_distances_jun26")        or []
_ex_delivery     = _load_json("existing_delivery_beats_jun26")   or []
_ex_trucks       = _load_json("existing_truck_assignments_jun26") or []
_plg_compare     = _load_json("plg_comparison_jun26") or {"rows":[],"totals":{}}
_ex_beat_meta    = _load_json("existing_beat_meta_jun26") or []
_outlet_meta     = _load_json("outlet_meta_jun26") or []
_ex_outlet_meta  = _load_json("existing_outlet_meta_jun26") or []

DATA_BLOCK = (
    "const OUTLETS    = " + json.dumps(outlets)       + ";\n"
    "const RS_INFO    = " + json.dumps(rs_info)       + ";\n"
    "const BOUNDARIES = " + json.dumps(boundaries)    + ";\n"
    "const STATS      = " + json.dumps(stats)         + ";\n"
    "const DUPE_PAIRS = " + json.dumps(dupe_pairs)    + ";\n"
    "const DUPE_STATS = " + json.dumps(dupe_stats)    + ";\n"
    "const CLUSTERS   = " + json.dumps(clusters)      + ";\n"
    "const CLUSTER_ST = " + json.dumps(cluster_stats) + ";\n"
    "const BEATS_390    = " + json.dumps(beats_390)     + ";\n"
    "const BEATS_391    = " + json.dumps(beats_391)     + ";\n"
    "const EX_BEATS_390 = " + json.dumps(ex_beats_390)  + ";\n"
    "const EX_BEATS_391 = " + json.dumps(ex_beats_391)  + ";\n"
    "const PLG_INFO     = " + json.dumps(plg_info)      + ";\n"
    "const DSE_INFO     = " + json.dumps(dse_info)      + ";\n"
    "const DSE_INFO_391 = " + json.dumps(dse_info_391)  + ";\n"
    "const BEAT_STATS   = " + json.dumps(beat_stats)    + ";\n"
    "const EXCL_OUTLETS = " + json.dumps(excl_outlets)  + ";\n"
    "const BENEFIT_STATS    = " + json.dumps(benefit_stats)    + ";\n"
    "const DSE_BALANCE_390  = " + json.dumps(dse_balance_390)  + ";\n"
    "const CONFLICTS_EX_390 = " + json.dumps(conflicts_ex_390) + ";\n"
    "const CONFLICTS_V3_390 = " + json.dumps(conflicts_v3_390) + ";\n"
    "const HULL_V3_390      = " + json.dumps(hull_v3_390)      + ";\n"
    "const HULL_EX_390      = " + json.dumps(hull_ex_390)      + ";\n"
    "const HULL_RS_EX       = " + json.dumps(hull_rs_ex)       + ";\n"
    "const HULL_RS_PROP     = " + json.dumps(hull_rs_prop)     + ";\n"
    "const RS_DIST_STATS    = " + json.dumps(rs_dist_stats)    + ";\n"
    "const RS_OVERLAP       = " + json.dumps(rs_overlap)       + ";\n"
    "const BEAT_DIST        = " + json.dumps(beat_distances)   + ";\n"
    "const BEAT_AREA        = " + json.dumps(beat_areas)       + ";\n"
    "const DELIVERY_BEATS_V4 = " + json.dumps(delivery_beats_v4) + ";\n"
    "const DELIVERY_ZONES  = " + json.dumps(delivery_zones)    + ";\n"
    "const DELIVERY_DATA = " + json.dumps(_delivery_data_slim) + ";\n"
    "const FLAGGED_PHARMA = " + json.dumps(_flagged_pharma) + ";\n"
    "const BEATS_JUN26     = " + json.dumps(_jun26_beats)     + ";\n"
    "const PLG_JUN26       = " + json.dumps(_jun26_plg)       + ";\n"
    "const DSE_JUN26       = " + json.dumps(_jun26_dse)       + ";\n"
    "const HULL_JUN26      = " + json.dumps(_jun26_hull)      + ";\n"
    "const CONFLICTS_JUN26 = " + json.dumps(_jun26_conflicts) + ";\n"
    "const ZONES_JUN26     = " + json.dumps(_jun26_zones)     + ";\n"
    "const DELIVERY_JUN26  = " + json.dumps(_jun26_delivery)  + ";\n"
    "const TRUCKS_JUN26    = " + json.dumps(_jun26_trucks)    + ";\n"
    "const DIST_JUN26      = " + json.dumps(_jun26_dist)      + ";\n"
    "const EX_BEATS_J26    = " + json.dumps(_ex_beats)         + ";\n"
    "const EX_PLG_J26      = " + json.dumps(_ex_plg)           + ";\n"
    "const EX_DSE_J26      = " + json.dumps(_ex_dse)           + ";\n"
    "const EX_HULL_J26     = " + json.dumps(_ex_hull)          + ";\n"
    "const EX_DIST_J26     = " + json.dumps(_ex_dist)          + ";\n"
    "const EX_DELIVERY_J26 = " + json.dumps(_ex_delivery)      + ";\n"
    "const EX_TRUCKS_J26   = " + json.dumps(_ex_trucks)        + ";\n"
    "const PLG_CMP_J26     = " + json.dumps(_plg_compare)       + ";\n"
    "const EX_BEAT_META    = " + json.dumps(_ex_beat_meta)      + ";\n"
    "const OUTLET_META     = " + json.dumps(_outlet_meta)       + ";\n"
    "const EX_OUTLET_META  = " + json.dumps(_ex_outlet_meta)    + ";\n"
)

# ── HTML ───────────────────────────────────────────────────────────────────────
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>HUL Kolkata NTO &amp; PJP</title>
<link href="https://unpkg.com/maplibre-gl@4.7.0/dist/maplibre-gl.css" rel="stylesheet"/>
<script src="https://unpkg.com/maplibre-gl@4.7.0/dist/maplibre-gl.js"></script>
<link href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" rel="stylesheet"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box;}
html,body{width:100%;height:100%;overflow:hidden;
  font-family:'Segoe UI',Arial,sans-serif;background:#f8fafc;}

#slides{width:100vw;height:100vh;overflow-y:scroll;overflow-x:hidden;
  scroll-snap-type:y mandatory;}
#slides::-webkit-scrollbar{display:none;}
.slide{width:100vw;height:100vh;scroll-snap-align:start;position:relative;overflow:hidden;}

.map-wrap{position:absolute;top:0;left:0;right:400px;bottom:0;}
.map-wrap .maplibregl-map{width:100% !important;height:100% !important;}
.maplibregl-ctrl-logo{display:none !important;}
.maplibregl-ctrl-attrib{font-size:10px !important;opacity:0.45;}
.maplibregl-popup{z-index:20 !important;}
.maplibregl-marker{z-index:6 !important;}

.panel{position:absolute;top:0;right:0;width:400px;height:100%;
  background:rgba(255,255,255,0.97);backdrop-filter:blur(12px);
  box-shadow:-4px 0 28px rgba(0,0,0,0.10);
  overflow-y:auto;padding:24px 20px;z-index:20;}
.panel::-webkit-scrollbar{width:4px;}
.panel::-webkit-scrollbar-thumb{background:#d1d5db;border-radius:2px;}

.panel h2{font-size:15px;font-weight:700;color:#111827;
  padding-bottom:10px;border-bottom:2px solid #e5e7eb;margin-bottom:3px;}
.p-sub{font-size:12px;color:#9ca3af;margin-bottom:12px;}
.kpi-r{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:12px;}
.kpi{background:#f9fafb;border-radius:8px;padding:10px 12px;}
.kpi .kv{font-size:20px;font-weight:700;color:#111827;}
.kpi .kl{font-size:10px;color:#9ca3af;text-transform:uppercase;letter-spacing:.6px;margin-top:2px;}

.filter-row{display:flex;gap:6px;margin-bottom:12px;flex-wrap:wrap;}
.f-chip{padding:5px 14px;border:2px solid #e5e7eb;border-radius:20px;
  background:white;cursor:pointer;font-size:12px;font-weight:600;color:#374151;transition:all .15s;}
.f-chip.active{border-color:#1565C0;background:#e3f2fd;color:#1565C0;}
.f-chip:hover:not(.active){border-color:#9ca3af;}

.beat-chip{padding:4px 10px;border:2px solid #e5e7eb;border-radius:20px;
  background:white;cursor:pointer;font-size:11px;font-weight:700;color:#374151;transition:all .15s;}
.beat-chip.active{color:white;}
.beat-chip:hover:not(.active){border-color:#9ca3af;}
/* PLG-DSE accordion tree */
.plg-tree{margin-bottom:10px;}
.plg-tree-sec{font-size:10px;font-weight:700;letter-spacing:.7px;text-transform:uppercase;
  padding:4px 4px 2px;color:#9ca3af;}
.plg-tree-sec.ofm{color:#7c3aed;}.plg-tree-sec.uni{color:#0369a1;}
.plg-item{border:1.5px solid #e5e7eb;border-radius:8px;margin-bottom:3px;overflow:hidden;}
.plg-item.sel{border-color:#1565C0;}
.plg-row{display:flex;align-items:center;padding:7px 10px;gap:8px;cursor:pointer;
  background:white;transition:background .1s;}
.plg-row:hover{background:#f8fafc;}
.plg-item.sel .plg-row{background:#eff6ff;}
.plg-radio{width:15px;height:15px;border-radius:50%;border:2px solid #d1d5db;
  flex-shrink:0;transition:all .15s;position:relative;}
.plg-radio.on{border-color:#1565C0;background:#1565C0;}
.plg-radio.on::after{content:'';position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);
  width:5px;height:5px;border-radius:50%;background:white;}
.plg-cb{width:14px;height:14px;border:2px solid #d1d5db;border-radius:3px;flex-shrink:0;
  position:relative;transition:all .15s;cursor:pointer;}
.plg-cb.on{border-color:#1565C0;background:#1565C0;}
.plg-cb.on::after{content:'✓';position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);
  color:white;font-size:9px;font-weight:900;line-height:1;}
.plg-cb.partial{border-color:#1565C0;background:#1565C0;}
.plg-cb.partial::after{content:'–';position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);
  color:white;font-size:11px;font-weight:900;line-height:1;}
.dse-plg-tag{font-size:10px;padding:1px 5px;border-radius:3px;font-weight:600;white-space:nowrap;}
.plg-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0;}
.plg-name{flex:1;font-size:12px;font-weight:700;color:#111827;}
.plg-cnt{font-size:10px;color:#9ca3af;}
.plg-chev{font-size:11px;color:#9ca3af;transition:transform .2s;flex-shrink:0;}
.plg-chev.open{transform:rotate(180deg);}
.dse-list{border-top:1px solid #f3f4f6;padding:4px 10px 8px 36px;display:none;}
.dse-list.open{display:block;}
.dse-item{display:flex;align-items:center;gap:7px;padding:3px 0;cursor:pointer;}
.dse-cb{width:14px;height:14px;border:2px solid #d1d5db;border-radius:3px;flex-shrink:0;
  position:relative;transition:all .15s;}
.dse-cb.on{border-color:#1565C0;background:#1565C0;}
.dse-cb.on::after{content:'✓';position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);
  color:white;font-size:9px;font-weight:900;line-height:1;}
.dse-cb.partial{border-color:#1565C0;background:#1565C0;}
.dse-cb.partial::after{content:'–';position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);
  color:white;font-size:11px;font-weight:900;line-height:1;}
.dse-label{font-size:11px;color:#374151;}
.plg-all-row{display:flex;align-items:center;gap:8px;padding:7px 10px;cursor:pointer;
  border:1.5px solid #e5e7eb;border-radius:8px;margin-bottom:6px;background:white;transition:background .1s;}
.plg-all-row.sel{border-color:#1565C0;background:#eff6ff;}
.plg-all-row:hover:not(.sel){background:#f8fafc;}

.toggle-row{display:flex;gap:0;margin-bottom:12px;
  border:2px solid #e5e7eb;border-radius:10px;overflow:hidden;}
.t-btn{flex:1;padding:8px 0;font-size:13px;font-weight:700;
  background:white;border:none;cursor:pointer;color:#6b7280;transition:all .15s;}
.t-btn.active{background:#1565C0;color:white;}

.dt-tbl{width:100%;border-collapse:collapse;font-size:12.5px;margin-top:4px;}
.dt-tbl th{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;
  color:#9ca3af;padding:5px 4px;border-bottom:1px solid #e5e7eb;text-align:right;}
.dt-tbl th:first-child,.dt-tbl th:nth-child(2){text-align:left;}
.dt-tbl td{padding:6px 4px;border-bottom:1px solid #f3f4f6;color:#374151;text-align:right;}
.dt-tbl td:first-child{text-align:left;}
.dt-tbl td:nth-child(2){text-align:left;max-width:160px;overflow:hidden;}
.rs-nm{display:inline-block;max-width:130px;overflow:hidden;white-space:nowrap;
  text-overflow:ellipsis;vertical-align:bottom;}
.dt-tbl tr:hover td{background:#f9fafb;cursor:pointer;}
.dc{width:9px;height:9px;border-radius:2px;display:inline-block;
  margin-right:5px;flex-shrink:0;vertical-align:middle;}

.rs-item{display:flex;align-items:center;padding:7px 6px;border-radius:8px;
  cursor:pointer;transition:background .12s;gap:9px;}
.rs-item:hover{background:#f3f4f6;}
.rs-dot{width:10px;height:10px;border-radius:3px;flex-shrink:0;}
.rs-name{font-size:12.5px;font-weight:500;color:#111827;flex:1;min-width:0;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.rs-badge{font-size:10px;font-weight:700;padding:2px 7px;border-radius:10px;
  text-transform:uppercase;letter-spacing:.4px;flex-shrink:0;white-space:nowrap;}
.rs-cnt{font-size:13px;font-weight:700;color:#374151;flex-shrink:0;
  min-width:44px;text-align:right;}

.reass-box{background:#e3f2fd;border:1px solid #90caf9;border-radius:10px;
  padding:10px 14px;margin-bottom:12px;}
.reass-box b{color:#0d47a1;font-size:15px;}

.dupe-item{padding:8px 10px;border-radius:8px;cursor:pointer;
  border-bottom:1px solid #f3f4f6;transition:background .12s;}
.dupe-item:hover{background:#fff7ed;}
.dupe-item .d-na{font-size:12px;font-weight:600;color:#111827;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.dupe-item .d-nb{font-size:11px;color:#6b7280;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.dupe-item .d-meta{font-size:10px;color:#9ca3af;margin-top:2px;}
.dupe-dist{display:inline-block;padding:1px 6px;border-radius:8px;
  font-size:10px;font-weight:700;background:#fee2e2;color:#dc2626;margin-left:6px;}

.cl-item{padding:9px 10px;cursor:pointer;border-bottom:1px solid #f3f4f6;
  transition:background .12s;}
.cl-item:hover{background:#f0f9ff;}
.cl-item.sel{background:#eff6ff;}

.dl-btn{width:100%;padding:9px 0;background:#1565C0;color:white;
  border:none;border-radius:8px;font-size:13px;font-weight:700;
  cursor:pointer;margin-bottom:12px;transition:background .15s;}
.dl-btn:hover{background:#0d47a1;}

.page-lbl{position:absolute;bottom:16px;left:16px;
  font-size:11px;font-weight:700;letter-spacing:1px;color:#6b7280;
  text-transform:uppercase;z-index:500;
  background:rgba(255,255,255,0.9);padding:4px 11px;border-radius:20px;pointer-events:none;}
.zoom-hint{position:absolute;bottom:16px;
  left:calc((100% - 400px)/2);transform:translateX(-50%);
  font-size:11px;color:#9ca3af;background:rgba(255,255,255,0.85);
  padding:4px 11px;border-radius:20px;z-index:25;pointer-events:none;white-space:nowrap;}

#nav-dots{position:fixed;bottom:20px;left:0;width:100vw;
  display:flex;justify-content:center;gap:10px;z-index:9999;}
.dot{width:8px;height:8px;border-radius:50%;background:rgba(0,0,0,0.2);
  cursor:pointer;transition:all .2s;border:1.5px solid rgba(0,0,0,0.1);}
.dot.active{background:#1565C0;transform:scale(1.35);border-color:#1565C0;}
#nav-dots.dark-mode .dot{background:rgba(255,255,255,0.35);border-color:rgba(255,255,255,0.4);}
#nav-dots.dark-mode .dot.active{background:#90caf9;border-color:#90caf9;}
#sb-logo{filter:brightness(0) opacity(0.35);}
#sb-logo.on-dark{filter:brightness(0) invert(1) opacity(0.55);}

#slide-0{
  background:linear-gradient(135deg,#0a1929 0%,#0d47a1 55%,#0a1929 100%);
  display:flex;flex-direction:column;align-items:center;
  justify-content:center;text-align:center;color:white;}
.t-badge{background:rgba(21,101,192,0.35);border:1px solid rgba(21,101,192,0.6);
  border-radius:20px;padding:7px 22px;font-size:12px;font-weight:700;
  letter-spacing:1.5px;text-transform:uppercase;color:#90caf9;margin-bottom:28px;}
.t-h1{font-size:54px;font-weight:800;line-height:1.1;margin-bottom:14px;
  background:linear-gradient(90deg,#fff 30%,#90caf9 100%);
  -webkit-background-clip:text;-webkit-text-fill-color:transparent;}
.t-sub{font-size:17px;color:#94a3b8;margin-bottom:46px;max-width:600px;line-height:1.55;}
.s-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;
  max-width:800px;width:90%;margin-bottom:36px;}
.s-box{background:rgba(255,255,255,0.07);border:1px solid rgba(255,255,255,0.12);
  border-radius:12px;padding:22px 16px;}
.s-box .sv{font-size:30px;font-weight:700;color:#f1f5f9;margin-bottom:6px;}
.s-box .sl{font-size:11px;color:#94a3b8;text-transform:uppercase;letter-spacing:.8px;}
.scroll-h{font-size:13px;color:#64748b;display:flex;align-items:center;gap:10px;}
.arr{width:26px;height:26px;border:2px solid #334155;border-radius:50%;
  display:flex;align-items:center;justify-content:center;
  animation:bounce .65s infinite alternate;}
@keyframes bounce{from{transform:translateY(0)}to{transform:translateY(5px)}}
kbd{background:#1565C0;padding:2px 7px;border-radius:3px;font-size:12px;
  color:#90caf9;font-family:monospace;}

/* Density slider */
.slider-wrap{margin-bottom:12px;}
.slider-wrap label{font-size:11px;font-weight:700;color:#6b7280;
  text-transform:uppercase;letter-spacing:.5px;display:block;margin-bottom:6px;}
.slider-row{display:flex;align-items:center;gap:10px;}
.slider-row input[type=range]{flex:1;accent-color:#1565C0;cursor:pointer;}
.slider-val{font-size:18px;font-weight:800;color:#1565C0;min-width:36px;text-align:right;}
.cluster-meta{font-size:11px;color:#6b7280;margin-bottom:8px;padding:6px 8px;
  background:#f9fafb;border-radius:6px;}


/* ── Benefit info slides (8, 10) ── */
.info-slide{overflow-y:auto;}
.info-slide .page-lbl{background:rgba(0,0,0,0.3);color:rgba(255,255,255,0.7);}
.bs-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-bottom:28px;}
.bs-card{border-radius:12px;padding:22px;text-align:center;}
.bs-card .bs-v{font-size:40px;font-weight:800;margin-bottom:6px;}
.bs-card .bs-l{font-size:11px;text-transform:uppercase;letter-spacing:.6px;line-height:1.4;}
.bal-row{display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid #f3f4f6;}
.bal-row:last-child{border-bottom:none;}
.bal-lbl{font-size:12px;font-weight:600;color:#374151;min-width:72px;}
.bar-wrap{flex:1;display:flex;flex-direction:column;gap:3px;}
.bar-ex{height:12px;border-radius:3px;background:#fca5a5;position:relative;}
.bar-v3{height:12px;border-radius:3px;background:#86efac;position:relative;}
.bar-pct{position:absolute;right:4px;top:50%;transform:translateY(-50%);font-size:9px;font-weight:700;color:#374151;}
/* ── Slide 9 Leaflet map ── */
#l9-map .leaflet-control-attribution{font-size:9px;opacity:.5;}
</style>
</head>
<body>

<div id="slides">

<!-- SLIDE 0 · TITLE -->
<div class="slide" id="slide-0">
  <img src="data:image/png;base64,__HUL_LOGO__" alt="HUL" style="position:absolute;top:14px;right:24px;height:68px;object-fit:contain;filter:brightness(1.25) drop-shadow(0 1px 4px rgba(0,0,0,0.4))">
  <div class="t-badge">HUL Calcutta Metro &middot; NTO &amp; PJP</div>
  <h1 class="t-h1">New Territory<br/>Organization</h1>
  <p class="t-sub">Outlet and distributor analysis for Calcutta Metro &nbsp;&middot;&nbsp; <span style="color:#60a5fa;font-weight:600">Prepared by Stackbox</span></p>
  <div class="s-grid" id="title-stats"></div>
  <div class="scroll-h">
    <div class="arr">&#8595;</div>
    Scroll to explore &nbsp;&middot;&nbsp;
    <kbd>&#8593;</kbd><kbd>&#8595;</kbd> keys to navigate
  </div>
</div>

<!-- SLIDE SUMMARY · KEY BENEFITS -->
<div class="slide info-slide" id="slide-summary" style="background:linear-gradient(135deg,#0a1929 0%,#1a3a5c 100%);display:flex;flex-direction:column;justify-content:center;padding:40px 48px;overflow-y:auto">
  <div class="page-lbl" style="background:rgba(0,0,0,0.3);color:rgba(255,255,255,0.7)">1 / 12 &middot; Key Benefits</div>
  <div style="max-width:820px;margin:0 auto;width:100%">
    <div style="font-size:11px;font-weight:700;color:#60a5fa;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:10px">What This Redesign Delivers</div>
    <h2 style="font-size:30px;font-weight:800;color:white;margin-bottom:6px;line-height:1.2">Route Redesign &mdash; Calcutta Metro</h2>
    <p style="font-size:13px;color:#94a3b8;margin-bottom:28px">Sales territory RS&nbsp;218390 &middot; 108 salesmen &middot; ~49,000 active outlets</p>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px">

      <div style="background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:20px 24px">
        <div style="font-size:36px;font-weight:800;color:#4ade80;margin-bottom:4px">73%</div>
        <div style="font-size:14px;font-weight:700;color:white;margin-bottom:6px">Reduction in Mirror Beats</div>
        <div style="font-size:12px;color:#94a3b8;line-height:1.5">Each salesman covers a distinct, non-overlapping area per day. Same-day visits by 2+ salesmen to the same outlet dropped from 8,171 to 2,192.</div>
      </div>

      <div style="background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:20px 24px">
        <div style="font-size:36px;font-weight:800;color:#4ade80;margin-bottom:4px">~0%</div>
        <div style="font-size:14px;font-weight:700;color:white;margin-bottom:6px">Territory Overlap (Proposed)</div>
        <div style="font-size:12px;color:#94a3b8;line-height:1.5">Existing routes have 22&ndash;52% overlap by product category &mdash; multiple salesmen covering the same geography. The new design eliminates this entirely.</div>
      </div>

      <div style="background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:20px 24px">
        <div style="font-size:36px;font-weight:800;color:#60a5fa;margin-bottom:4px">100%</div>
        <div style="font-size:14px;font-weight:700;color:white;margin-bottom:6px">Category-Specialist Salesmen</div>
        <div style="font-size:12px;color:#94a3b8;line-height:1.5">Previously 53 of 107 salesmen carried mixed product portfolios, splitting their focus. Every salesman now owns a single product category &mdash; deeper expertise, better hit rates.</div>
      </div>

      <div style="background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:20px 24px">
        <div style="font-size:36px;font-weight:800;color:#60a5fa;margin-bottom:4px">39%</div>
        <div style="font-size:14px;font-weight:700;color:white;margin-bottom:6px">Fairer Workload Distribution</div>
        <div style="font-size:12px;color:#94a3b8;line-height:1.5">Workload imbalance across salesmen reduced from 24.9% to 15.1%. Fewer overloaded routes and fewer underutilised ones &mdash; more consistent performance.</div>
      </div>

    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
      <div style="background:rgba(96,165,250,0.08);border:1px solid rgba(96,165,250,0.2);border-radius:10px;padding:14px 20px;display:flex;align-items:center;gap:16px">
        <div style="font-size:28px;font-weight:800;color:#fbbf24">27%</div>
        <div>
          <div style="font-size:13px;font-weight:700;color:white">Shorter Delivery Routes</div>
          <div style="font-size:12px;color:#94a3b8">Average route length drops from 4.4 km to 3.2 km per beat. Less travel time, lower fuel cost, more calls per day.</div>
        </div>
      </div>
      <div style="background:rgba(96,165,250,0.08);border:1px solid rgba(96,165,250,0.2);border-radius:10px;padding:14px 20px;display:flex;align-items:center;gap:16px">
        <div style="font-size:28px;font-weight:800;color:#fbbf24">2,939</div>
        <div>
          <div style="font-size:13px;font-weight:700;color:white">Duplicate Outlets Cleaned</div>
          <div style="font-size:12px;color:#94a3b8">Ghost stores and duplicates removed &mdash; equivalent to ~13 salesmen&rsquo;s worth of wasted coverage eliminated from the system.</div>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- SLIDE 1 · OUTLETS & DISTRIBUTORS -->
<div class="slide" id="slide-1">
  <div class="map-wrap" id="map-1"></div>
  <div class="page-lbl">2 / 19 &middot; Outlets &amp; Distributors</div>
  <div class="zoom-hint">Ctrl+Scroll or Pinch to zoom</div>
  <div class="panel">
    <h2>Outlets &amp; Distributors</h2>
    <p class="p-sub">Calcutta Metro &middot; all active outlets</p>
    <div class="filter-row">
      <button class="f-chip active" onclick="setFilter('ALL')">All</button>
      <button class="f-chip"        onclick="setFilter('General')">General</button>
      <button class="f-chip"        onclick="setFilter('Pharma')">Pharma</button>
      <button class="f-chip"        onclick="setFilter('WS')">WS</button>
      <button class="f-chip" id="excl-btn" onclick="toggleExcl()"
        style="color:#ef4444;border-color:#fca5a5">Excl.</button>
    </div>
    <div id="p1-excl-panel" style="display:none;margin-bottom:12px">
      <p style="font-size:11px;color:#ef4444;margin-bottom:6px">
        GPS is far from mapped RS &middot; verify retag or remove
      </p>
      <button class="dl-btn" style="background:#ef4444;margin-bottom:0" onclick="downloadExcl()">
        &#8595; Download Excluded Outlets CSV</button>
    </div>
    <div class="kpi-r" id="p1-kpis"></div>
    <table class="dt-tbl">
      <thead><tr>
        <th style="width:20px;padding:5px 4px 5px 0;text-align:center">
          <input type="checkbox" id="p1-sel-all-chk" onclick="toggleSelectAllRS()"
            style="accent-color:#1565C0;width:13px;height:13px;cursor:pointer;vertical-align:middle"/>
        </th>
        <th>Distributor</th><th>Type</th><th>Outlets</th><th>MOC</th></tr></thead>
      <tbody id="p1-tb"></tbody>
    </table>
  </div>
</div>

<!-- SLIDE 2 · TERRITORY OVERLAPS -->
<div class="slide" id="slide-2">
  <div class="map-wrap" id="map-2"></div>
  <div class="page-lbl">3 / 19 &middot; Territory Overlaps</div>
  <div class="zoom-hint">Ctrl+Scroll or Pinch to zoom</div>
  <div class="panel">
    <h2>Territory Overlaps</h2>
    <div class="toggle-row" style="margin-bottom:8px">
      <button class="t-btn active" id="t-gen" onclick="setTerType('General')">General</button>
      <button class="t-btn"        id="t-pha" onclick="setTerType('Pharma')">Pharma</button>
    </div>
    <div class="toggle-row">
      <button class="t-btn active" id="t-existing" onclick="setView('existing')">Existing</button>
      <button class="t-btn"        id="t-proposed" onclick="setView('proposed')">Proposed</button>
    </div>
    <div style="font-size:11px;color:#6b7280;line-height:1.5;margin:8px 0 4px">
      Convex hull of each distributor&rsquo;s outlets. Overlapping fills = territory overlap.
    </div>
    <div id="p2-overlap-stats" style="margin-bottom:8px"></div>
    <div class="filter-row" style="gap:4px;margin-bottom:6px">
      <button class="beat-chip active" id="p2-tg-boundary" onclick="s2tgl('boundary')" style="background:#1565C0;color:#fff;border-color:#1565C0">Hulls</button>
      <button class="beat-chip" id="p2-tg-ret" onclick="s2tgl('ret')" style="">Outlets</button>
    </div>
    <button class="dl-btn" id="p2-dl-btn" style="display:none;margin-top:4px" onclick="downloadProposed()">
      &#8595; Download Proposed Plan CSV</button>
    <table class="dt-tbl" style="margin-top:8px">
      <thead><tr>
        <th style="width:20px;padding:5px 4px 5px 0;text-align:center">
          <input type="checkbox" id="p2-sel-all-chk" onclick="toggleSelectAllRS2()"
            style="accent-color:#1565C0;width:13px;height:13px;cursor:pointer;vertical-align:middle"/>
        </th>
        <th>Distributor</th><th id="p2-col-ol">Outlets</th><th id="p2-col-moc">MOC</th></tr></thead>
      <tbody id="p2-tb"></tbody>
    </table>
  </div>
</div>

<!-- SLIDE 4 · HIGH DENSITY CLUSTERS (position 3) -->
<div class="slide" id="slide-4">
  <div class="map-wrap" id="map-4"></div>
  <div class="page-lbl">4 / 19 &middot; High Density Clusters</div>
  <div class="zoom-hint">Ctrl+Scroll or Pinch to zoom</div>
  <div class="panel">
    <h2>High Density Clusters</h2>
    <p class="p-sub">~20m grid cells &middot; slide to adjust threshold</p>
    <button class="dl-btn" onclick="downloadClusters()" style="margin-bottom:8px">&#8595; Download Cluster Data CSV</button>
    <div class="slider-wrap">
      <label>Min outlets per cluster</label>
      <div class="slider-row">
        <input type="range" min="5" max="50" value="5" id="density-slider"
          oninput="setDensity(+this.value)">
        <span class="slider-val" id="density-val">5</span>
      </div>
    </div>
    <div class="kpi-r" id="p4-kpis"></div>
    <div class="cluster-meta" id="p4-meta"></div>
    <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;
      letter-spacing:.5px;margin-bottom:6px">Top Clusters</div>
    <div id="p4-list" style="overflow-y:auto;max-height:calc(100vh - 390px)"></div>
  </div>
</div>

<!-- SLIDE 3 · DUPLICATE OUTLETS (position 2) -->
<div class="slide" id="slide-3">
  <div class="map-wrap" id="map-3"></div>
  <div class="page-lbl">5 / 19 &middot; Duplicate Outlets</div>
  <div class="zoom-hint">Ctrl+Scroll or Pinch to zoom</div>
  <div class="panel" style="overflow:hidden;display:flex;flex-direction:column;">
    <h2>Duplicate Outlets</h2>
    <p class="p-sub">AI-verified pairs &middot; same store, two entries</p>
    <div class="kpi-r" id="p3-kpis"></div>
    <button class="dl-btn" onclick="downloadDupes()">&#8595; Download Duplicate Pairs CSV</button>
    <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;
      letter-spacing:.5px;margin:4px 0 6px;flex-shrink:0">Pairs (sorted by distance)</div>
    <div id="p3-list" style="overflow-y:scroll;flex:1;min-height:0;"></div>
  </div>
</div>

<!-- SLIDE 11 · PLG RULES (position 3) -->
<div class="slide info-slide" id="slide-11" style="background:linear-gradient(135deg,#0a1929 0%,#1a3a5c 100%);overflow-y:auto">
  <div class="page-lbl">6 / 19 &middot; PLG Rules &middot; RS 218390</div>
  <div style="max-width:980px;margin:0 auto;padding:32px 28px;color:white">
    <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;color:#60a5fa;text-transform:uppercase;margin-bottom:8px">Reference &middot; RS 218390</div>
    <h2 style="font-size:26px;font-weight:800;color:white;margin-bottom:4px">New PLG Assignment Rules</h2>
    <p style="font-size:12px;color:#94a3b8;margin-bottom:16px;line-height:1.5">How existing visit patterns map to V3 PLG assignments. Each rule specifies the new beat structure for outlets that had a given visit pattern in the existing design.</p>
    <table style="width:100%;border-collapse:collapse;font-size:12px">
      <thead>
        <tr style="border-bottom:2px solid rgba(255,255,255,0.2)">
          <th style="padding:7px 8px;text-align:left;font-size:10px;color:#60a5fa;font-weight:700;text-transform:uppercase;letter-spacing:.5px;width:35%">Existing Pattern</th>
          <th style="padding:7px 8px;text-align:left;font-size:10px;color:#60a5fa;font-weight:700;text-transform:uppercase;letter-spacing:.5px;width:28%">New PLG Rule</th>
          <th style="padding:7px 8px;text-align:right;font-size:10px;color:#60a5fa;font-weight:700;text-transform:uppercase;letter-spacing:.5px;width:9%">Outlets</th>
          <th style="padding:7px 8px;text-align:left;font-size:10px;color:#60a5fa;font-weight:700;text-transform:uppercase;letter-spacing:.5px">Notes</th>
        </tr>
      </thead>
      <tbody id="p11-rules-tbl"></tbody>
    </table>
    <div style="margin-top:16px;display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div style="padding:10px 12px;background:rgba(255,255,255,0.05);border-radius:8px;font-size:11px;color:#94a3b8;line-height:1.6">
        <strong style="color:#e2e8f0">OFM</strong> = Off-Modern Trade (dedicated salesman for large-format modern retail, 12 stores/day)<br/>
        <strong style="color:#e2e8f0">Uniglow</strong> = Unilever premium brand channel (30 stores)<br/>
        <strong style="color:#e2e8f0">Unicare</strong> = Unilever personal care channel (20 stores)
      </div>
      <div style="padding:10px 12px;background:rgba(255,255,255,0.05);border-radius:8px;font-size:11px;color:#94a3b8">
        <div style="font-size:10px;font-weight:700;color:#60a5fa;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">Channel Program</div>
        <div style="display:flex;justify-content:space-between;padding:3px 0;border-bottom:1px solid rgba(255,255,255,0.08)"><span>DC</span><span style="color:#e2e8f0;font-weight:700">100 stores</span></div>
        <div style="display:flex;justify-content:space-between;padding:3px 0;border-bottom:1px solid rgba(255,255,255,0.08)"><span>OFM</span><span style="color:#e2e8f0;font-weight:700">100 stores</span><span style="color:#60a5fa;font-size:10px">20 w/ OFM program</span></div>
        <div style="display:flex;justify-content:space-between;padding:3px 0"><span>HNB</span><span style="color:#e2e8f0;font-weight:700">100 stores</span><span style="color:#60a5fa;font-size:10px">30 Uniglow + 20 Unicare</span></div>
      </div>
    </div>
  </div>
</div>

<!-- SLIDE 8 · PLG PURITY -->
<div class="slide info-slide" id="slide-8" style="background:linear-gradient(135deg,#0a1929 0%,#1a3a5c 100%)">
  <div class="page-lbl">7 / 19 &middot; PLG Purity &middot; RS 218390</div>
  <div style="max-width:860px;margin:0 auto;padding:44px 28px;color:white">
    <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;color:#60a5fa;text-transform:uppercase;margin-bottom:12px">Benefit 2 &middot; RS 218390</div>
    <h2 style="font-size:32px;font-weight:800;margin-bottom:8px;color:white">PLG Purity</h2>
    <p style="font-size:13px;color:#94a3b8;margin-bottom:24px;max-width:580px;line-height:1.6">In V4 every salesman specialises in exactly one product category. Previously 57 of 107 salesmen carried mixed portfolios across 2&ndash;3 PLG types.</p>
    <div class="bs-grid">
      <div class="bs-card" style="background:rgba(248,113,113,0.15);border:1px solid rgba(248,113,113,0.35)">
        <div class="bs-v" style="color:#f87171">57</div>
        <div class="bs-l" style="color:#94a3b8">Impure Salesmen<br/>Existing design</div>
      </div>
      <div class="bs-card" style="background:rgba(74,222,128,0.15);border:1px solid rgba(74,222,128,0.35)">
        <div class="bs-v" style="color:#4ade80">0</div>
        <div class="bs-l" style="color:#94a3b8">Impure Salesmen<br/>V4 design</div>
      </div>
      <div class="bs-card" style="background:rgba(96,165,250,0.15);border:1px solid rgba(96,165,250,0.35)">
        <div class="bs-v" style="color:#60a5fa">100%</div>
        <div class="bs-l" style="color:#94a3b8">Pure specialist<br/>coverage in V4</div>
      </div>
    </div>
    <div style="font-size:11px;font-weight:700;color:#60a5fa;letter-spacing:.8px;text-transform:uppercase;margin-bottom:8px">Mixed-portfolio salesmen in existing design (sample of 57)</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead>
        <tr style="border-bottom:1px solid rgba(255,255,255,0.15)">
          <th style="padding:7px 6px;text-align:left;font-size:10px;color:#60a5fa;font-weight:700;text-transform:uppercase;letter-spacing:.5px">Salesman</th>
          <th style="padding:7px 6px;text-align:left;font-size:10px;color:#60a5fa;font-weight:700;text-transform:uppercase;letter-spacing:.5px">PLG Types Assigned</th>
          <th style="padding:7px 6px;text-align:right;font-size:10px;color:#60a5fa;font-weight:700;text-transform:uppercase;letter-spacing:.5px">PLG Count</th>
        </tr>
      </thead>
      <tbody id="p8-impure-tbl"></tbody>
    </table>
    <div style="margin-top:18px;padding:13px;background:rgba(255,255,255,0.05);border-radius:8px;font-size:12px;color:#94a3b8;line-height:1.6">
      <strong style="color:#e2e8f0">Why it matters:</strong> Mixed-portfolio salesmen divide attention across categories, reducing depth per PLG. V3 assigns each Salesman exactly one Sub-PLG &mdash; specialist knowledge, dedicated targets, higher hit rate.
    </div>
  </div>
</div>

<!-- SLIDE 5 · BEATS -->
<div class="slide" id="slide-5" style="background:white">
  <div class="map-wrap" id="map-5"></div>
  <div class="page-lbl">8 / 19 &middot; Beats &middot; RS 218390 &amp; 218391</div>
  <div class="zoom-hint">Ctrl+Scroll or Pinch to zoom</div>
  <div class="panel" style="overflow:hidden;display:flex;flex-direction:column;padding:0;background:#fff;">
    <div style="padding:16px 18px 10px;flex:1;min-height:0;overflow-y:auto">
      <h2 style="margin-bottom:8px">Beats</h2>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px">RS</div>
      <div class="toggle-row" style="margin-bottom:10px">
        <button class="t-btn active" id="p5-rs390" onclick="setBeatsRS('218390')">218390</button>
        <button class="t-btn"        id="p5-rs391" onclick="setBeatsRS('218391')">218391</button>
      </div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px">View</div>
      <div class="toggle-row" style="margin-bottom:10px">
        <button class="t-btn active" id="p5-vproposed" onclick="setBeatsView('proposed')">Proposed</button>
        <button class="t-btn"        id="p5-vexisting" onclick="setBeatsView('existing')">Existing</button>
      </div>
      <div class="kpi-r" id="p5-kpis"></div>
      <div id="p5-colorby-section" style="display:none;margin:0 0 8px">
        <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px">Color by</div>
        <div class="toggle-row" style="margin-bottom:0">
          <button class="t-btn active" id="p5-cb-plg" onclick="setColorBy('plg')">PLG</button>
          <button class="t-btn" id="p5-cb-day" onclick="setColorBy('day')">Day</button>
        </div>
      </div>
      <div id="p5-day-section">
        <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:6px 0 4px">Filter by Day</div>
        <div class="filter-row" id="p5-day-chips" style="flex-wrap:wrap;gap:4px"></div>
      </div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:8px 0 5px" id="p5-filter-lbl">Filter by PLG &amp; Salesman</div>
      <div id="p5-plg-tree"></div>
      <div id="p5-dse-section" style="display:none">
        <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:6px 0 4px">Salesman</div>
        <div id="p5-dse-list"></div>
      </div>
      <div style="margin-top:12px;border-top:1px solid #e5e7eb;padding-top:10px">
        <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">Download</div>
        <button class="dl-btn" id="p5-v4-dl" onclick="downloadV4Beats()" style="background:#7030A0">
          &#8595; Download V4 Sales Beat CSV</button>
        <button class="dl-btn" id="p5-ex-dl" onclick="p5DownloadExisting()" style="background:#0369a1;display:none;margin-top:6px">
          &#8595; Download Existing Beats CSV</button>
      </div>
    </div>
  </div>
</div>

<!-- SLIDE 9 · JACCARD TERRITORIES (position 5) -->
<div class="slide" id="slide-9">
  <div class="map-wrap" id="l9-map"></div>
  <div class="page-lbl">9 / 19 &middot; Beat Territories &middot; RS 218390</div>
  <div class="zoom-hint">Ctrl+Scroll or Pinch to zoom</div>
  <div class="panel" style="overflow:hidden;display:flex;flex-direction:column;padding:0">
    <div style="padding:16px 18px 10px;flex:1;min-height:0;overflow-y:auto">
      <h2 style="margin-bottom:4px">Beat Territories &amp; Overlap</h2>
      <p class="p-sub" style="margin-bottom:8px">Convex hull per PLG-salesman-day &middot; overlap visible across PLGs</p>
      <div class="toggle-row" style="margin-bottom:10px">
        <button class="t-btn active" id="j9-vv3" onclick="setJ9View('v3')">Proposed beats</button>
        <button class="t-btn"        id="j9-vex" onclick="setJ9View('existing')">Existing beats</button>
      </div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:0 0 4px">Filter by day</div>
      <div class="filter-row" id="p9-day-chips" style="flex-wrap:wrap;gap:4px;margin-bottom:8px"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:6px 0 5px">Filter by PLG &amp; Salesman</div>
      <div id="p9-plg-tree"></div>
      <div id="p9-dse-list" style="max-height:120px;overflow-y:auto;margin-bottom:6px"></div>
      <div class="kpi-r" id="p9-kpis"></div>
      <div id="p9-dist-table"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:6px 0 4px">Overlap by PLG</div>
      <table class="dt-tbl">
        <thead><tr>
          <th style="text-align:left">Ex PLG &rarr; Prop</th>
          <th>Existing</th><th>Proposed</th>
        </tr></thead>
        <tbody id="p9-jac-body"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- SLIDE 7 · SAME-DAY CONFLICTS -->
<div class="slide" id="slide-7">
  <div class="map-wrap" id="map-7"></div>
  <div class="page-lbl">10 / 19 &middot; Same-Day Conflicts &middot; RS 218390</div>
  <div class="zoom-hint">Ctrl+Scroll or Pinch to zoom</div>
  <div class="panel">
    <h2>Same-Day Multi-Salesman Visits</h2>
    <p class="p-sub">Outlets visited by 2+ salesmen on the same market day</p>
    <div class="toggle-row" style="margin-bottom:10px">
      <button class="t-btn" id="s7-vex" onclick="setS7View('existing')">Existing</button>
      <button class="t-btn active" id="s7-vv3" onclick="setS7View('v3')">V4 Proposed</button>
    </div>
    <div class="kpi-r">
      <div class="kpi" style="border:1.5px solid #fee2e2">
        <div class="kv" style="color:#dc2626">5,436</div>
        <div class="kl">Existing conflict outlets</div>
        <div style="font-size:10px;color:#9ca3af;margin-top:2px">8,171 occurrences across 6 days</div>
      </div>
      <div class="kpi" style="border:1.5px solid #dcfce7">
        <div class="kv" style="color:#16a34a">1,961</div>
        <div class="kl">V4 conflict outlets</div>
        <div style="font-size:10px;color:#9ca3af;margin-top:2px">2,192 occurrences &middot; within-group by design</div>
      </div>
    </div>
    <div style="text-align:center;padding:10px;background:#f0fdf4;border-radius:8px;margin-bottom:8px">
      <div style="font-size:26px;font-weight:800;color:#16a34a">&#9660; 64%</div>
      <div style="font-size:11px;color:#6b7280">Reduction in conflict outlets vs Existing</div>
    </div>
    <div style="padding:8px;background:#eff6ff;border-radius:6px;font-size:11px;color:#374151;margin-bottom:12px;line-height:1.5">
      <strong>V4 delivery-bundling:</strong> Group A PLGs (D+F+N, D, D+F, F, PP-A) visit on day <em>b</em>, Group B (PP, F+N, N, PP-B) on day <em>b+1</em>. Delivery truck combines orders on day <em>b+2</em>. Within-group pairs sharing outlets (D+PP-A, F+N+PP-B) are intentionally same-day &mdash; 89.6% of D+F+N/PP pairs are consecutive, PP-A/PP-B 100%.
    </div>
    <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px">Legend (Market Day)</div>
    <div id="p7-legend"></div>
    <div style="font-size:11px;color:#9ca3af;margin-top:8px;line-height:1.5">Each dot = one outlet-day pair where 2+ salesmen visit on the same day. Color = visit day.</div>
  </div>
</div>

<!-- SLIDE 12 · BEAT AREA PER DAY -->
<div class="slide" id="slide-12">
  <div class="map-wrap" id="l12-map"></div>
  <div class="page-lbl">11 / 19 &middot; Beat Area per Day &middot; RS 218390</div>
  <div class="zoom-hint">Ctrl+Scroll to zoom</div>
  <div class="panel" style="overflow:hidden;display:flex;flex-direction:column;padding:0">
    <div style="padding:16px 18px 12px;flex-shrink:0;overflow-y:auto;max-height:100vh">
      <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;color:#1565C0;text-transform:uppercase;margin-bottom:6px">Benefit 5 &middot; RS 218390</div>
      <h2 style="margin-bottom:3px">Beat Area &mdash; Delivery Zone</h2>
      <p class="p-sub" style="margin-bottom:8px">Delivery coverage km&sup2; per market zone. Existing = 1-day sales; Proposed = 2-day bundled sales.</p>
      <div class="toggle-row" style="margin-bottom:8px">
        <button class="t-btn active" id="a12-vv3" onclick="setA12View('v3')">Proposed (2-day)</button>
        <button class="t-btn" id="a12-vex" onclick="setA12View('existing')">Existing (1-day)</button>
      </div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:0 0 4px">Filter by zone</div>
      <div class="filter-row" id="a12-zone-chips" style="flex-wrap:wrap;gap:4px;margin-bottom:10px"></div>
      <div class="kpi-r" style="grid-template-columns:1fr 1fr;margin-bottom:10px" id="p12-kpis"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">
        km&sup2; per delivery zone &mdash; <span style="color:#dc2626">Ex (1-day)</span> &nbsp; <span style="color:#7030A0">Prop (2-day)</span>
      </div>
      <div id="p12-chart"></div>
      <div style="margin-top:10px;padding:10px 12px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px;font-size:11px;color:#0c4a6e;line-height:1.6">
        <div style="font-size:10px;font-weight:700;color:#0369a1;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">V4 Delivery Bundling (N+2)</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:7px">
          <div style="background:rgba(255,255,255,0.6);border-radius:6px;padding:6px 8px">
            <div style="font-size:10px;font-weight:700;color:#0369a1;margin-bottom:2px">Day N &mdash; Group A</div>
            <div style="font-size:11px;color:#374151">D+F+N &nbsp; D &nbsp; D+F &nbsp; F &nbsp; PP-A</div>
          </div>
          <div style="background:rgba(255,255,255,0.6);border-radius:6px;padding:6px 8px">
            <div style="font-size:10px;font-weight:700;color:#0369a1;margin-bottom:2px">Day N+1 &mdash; Group B</div>
            <div style="font-size:11px;color:#374151">PP &nbsp; F+N &nbsp; N &nbsp; PP-B</div>
          </div>
        </div>
        <div style="margin-bottom:5px">Delivery truck combines orders from both days on <strong>day N+2</strong> &mdash; one trip per zone covers 2 days of sales.</div>
        <div style="border-top:1px solid #bae6fd;padding-top:5px;color:#374151">
          Within-group pairs (D&thinsp;+&thinsp;PP-A, F+N&thinsp;+&thinsp;PP-B) visit the same outlets on the same day intentionally &mdash;
          <strong>89.6%</strong> of D+F+N / PP pairs are on consecutive days &middot; <strong>100%</strong> of PP-A / PP-B pairs.
        </div>
      </div>
    </div>
  </div>
</div>

<!-- SLIDE 13 · DELIVERY BEATS -->
<div class="slide" id="slide-13">
  <div class="map-wrap" id="l13-map"></div>
  <div class="page-lbl">12 / 19 &middot; Delivery Beats &middot; RS 218390</div>
  <div class="zoom-hint">Ctrl+Scroll to zoom</div>
  <div class="panel" style="overflow:hidden;display:flex;flex-direction:column;padding:0">
    <div style="padding:16px 18px 12px;flex:1;min-height:0;overflow-y:auto">
      <h2 style="margin-bottom:4px">Delivery Beats</h2>
      <p class="p-sub" style="margin-bottom:10px">First-cut vehicle assignment &middot; RS 218390</p>
      <div class="toggle-row" style="margin-bottom:10px">
        <button class="t-btn active" id="db13-vex" onclick="setDB13View('existing')">Existing</button>
        <button class="t-btn" id="db13-vv4" onclick="setDB13View('proposed')">Proposed (V4 Zones)</button>
      </div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:0 0 4px">Max sellers per beat</div>
      <div class="toggle-row" style="margin-bottom:10px">
        <button class="t-btn active" id="db13-m2" onclick="setDB13Limit('Max 2 sellers')">Max 2</button>
        <button class="t-btn" id="db13-m3" onclick="setDB13Limit('Max 3 sellers')">Max 3</button>
        <button class="t-btn" id="db13-m4" onclick="setDB13Limit('Max 4 sellers')">Max 4</button>
      </div>
      <div id="db13-day-section">
        <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:0 0 4px">Filter by day</div>
        <div class="filter-row" id="db13-day-chips" style="flex-wrap:wrap;gap:4px;margin-bottom:8px"></div>
      </div>
      <div id="db13-zone-section" style="display:none">
        <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:0 0 4px">Filter by delivery zone</div>
        <div class="filter-row" id="db13-zone-chips" style="flex-wrap:wrap;gap:4px;margin-bottom:8px"></div>
      </div>
      <div class="kpi-r" id="db13-kpis"></div>
      <div id="db13-truck-legend" style="display:flex;gap:14px;margin:6px 0 10px;flex-wrap:wrap;font-size:11px"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px">Beat Summary</div>
      <table class="dt-tbl" style="font-size:11px">
        <thead><tr>
          <th style="text-align:left" id="db13-col-lbl">Day</th>
          <th>Beats</th><th>Outlets</th><th>3W</th><th>Ace</th><th>407</th><th>Cost</th>
        </tr></thead>
        <tbody id="db13-tbody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- SLIDE EXBEAT · Existing Beats Explorer (218390 + 20B801, ME BEAT) -->
<div class="slide" id="slide-exbeat" style="background:white">
  <div class="map-wrap" id="map-exbeat"></div>
  <div style="position:absolute;top:14px;left:50%;transform:translateX(-50%);z-index:1000;background:white;border:1px solid #d1d5db;border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,0.12);padding:6px 10px;display:flex;align-items:center;gap:6px;width:320px;pointer-events:auto">
    <span style="color:#6b7280;font-size:13px">🔍</span>
    <input type="text" id="exb-search" placeholder="Search outlet code / name / beat" oninput="exbSetOutletSearch(this.value)" style="flex:1;border:0;outline:none;font-size:12px;background:transparent">
    <span id="exb-search-count" style="font-size:10px;color:#9ca3af;white-space:nowrap"></span>
    <a href="javascript:void(0)" onclick="exbClearSearch()" id="exb-search-clear" style="display:none;color:#dc2626;cursor:pointer;font-size:11px">×</a>
  </div>
  <div class="page-lbl">13 / 19 &middot; Existing Beats Explorer &middot; RS 218390 + 20B801</div>
  <div class="zoom-hint">Ctrl+Scroll or Pinch to zoom</div>
  <div class="panel" style="overflow:hidden;display:flex;flex-direction:column;padding:0;background:#fff;">
    <div style="padding:16px 18px 10px;flex:1;min-height:0;overflow-y:auto">
      <h2 style="margin-bottom:4px">Existing Beats — 218390 + 20B801 (ME BEAT)</h2>
      <p class="p-sub" style="margin-bottom:8px">Multi-select filters: Day, PLG, Salesman, Beat. Outlets shown on map.</p>
      <div class="kpi-r" id="exb-kpis"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:8px 0 4px">Color by</div>
      <div class="toggle-row" style="margin-bottom:8px">
        <button class="t-btn active" id="exb-cb-plg" onclick="exbSetCB('plg')">PLG</button>
        <button class="t-btn" id="exb-cb-day" onclick="exbSetCB('day')">Day</button>
        <button class="t-btn" id="exb-cb-beat" onclick="exbSetCB('beat')">Beat</button>
      </div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:6px 0 4px">Filter by Day (multi)</div>
      <div class="filter-row" id="exb-day-chips" style="flex-wrap:wrap;gap:4px;margin-bottom:8px"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:6px 0 4px">Filter by PLG (multi)</div>
      <div class="filter-row" id="exb-plg-chips" style="flex-wrap:wrap;gap:4px;margin-bottom:8px"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:6px 0 4px">Filter by Salesman (multi, RSSP)</div>
      <input type="text" id="exb-dse-search" placeholder="Type to filter salesmen..." style="width:100%;padding:6px;border:1px solid #e5e7eb;border-radius:4px;font-size:11px;margin-bottom:4px" oninput="exbRenderDseList()">
      <div id="exb-dse-list" style="max-height:140px;overflow-y:auto;border:1px solid #f3f4f6;border-radius:4px;padding:4px;margin-bottom:8px;font-size:11px"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:6px 0 4px">Filter by Beat (multi, geographic name)</div>
      <input type="text" id="exb-beat-search" placeholder="Type to filter beats..." style="width:100%;padding:6px;border:1px solid #e5e7eb;border-radius:4px;font-size:11px;margin-bottom:4px" oninput="exbRenderBeatList()">
      <div id="exb-beat-list" style="max-height:200px;overflow-y:auto;border:1px solid #f3f4f6;border-radius:4px;padding:4px;font-size:11px"></div>
      <div style="margin-top:10px;border-top:1px solid #e5e7eb;padding-top:8px">
        <button class="dl-btn" onclick="exbClearFilters()" style="background:#6b7280">Clear All Filters</button>
      </div>
    </div>
  </div>
</div>

<!-- SLIDE JUN26-INTRO · Section divider for Jun 2026 work -->
<div class="slide info-slide" id="slide-jun26-intro" style="background:linear-gradient(135deg,#0a1929 0%,#1a3a5c 100%);display:flex;flex-direction:column;justify-content:center;padding:40px 48px;overflow-y:auto">
  <div style="max-width:900px;margin:0 auto;text-align:center;color:#fff;">
    <div style="font-size:13px;font-weight:700;letter-spacing:6px;color:#60a5fa;margin-bottom:14px;text-transform:uppercase">— Section II —</div>
    <h1 style="font-size:54px;font-weight:300;letter-spacing:1px;line-height:1.1;margin:0 0 12px;color:#fff">Jun 2026 Beat Redesign</h1>
    <div style="font-size:20px;font-weight:300;color:#cbd5e1;margin-bottom:36px">
      RS 218390 + 20B801 merged &middot; OFM &amp; UNIGLOW+UNICARE specialist channels &middot; mirror-aligned for delivery bundling
    </div>
    <div style="display:flex;justify-content:center;gap:42px;flex-wrap:wrap;margin-top:24px">
      <div><div style="font-size:36px;font-weight:600;color:#fff">7,103</div><div style="font-size:13px;color:#94a3b8;letter-spacing:1px;text-transform:uppercase">working outlets</div></div>
      <div><div style="font-size:36px;font-weight:600;color:#fff">16</div><div style="font-size:13px;color:#94a3b8;letter-spacing:1px;text-transform:uppercase">PLG buckets</div></div>
      <div><div style="font-size:36px;font-weight:600;color:#fff">105</div><div style="font-size:13px;color:#94a3b8;letter-spacing:1px;text-transform:uppercase">salesmen</div></div>
    </div>
    <div style="margin-top:34px;color:#94a3b8;font-size:14px">Slides 15–18 cover the merged Jun 2026 design</div>
  </div>
</div>

<!-- SLIDE JUN26 · ALIGNED BEATS (218390 + 20B801 merged, Jun 2026) -->
<div class="slide" id="slide-jun26" style="background:white">
  <div class="map-wrap" id="map-jun26"></div>
  <div class="map-search" style="position:absolute;top:14px;left:50%;transform:translateX(-50%);z-index:1000;background:white;border:1px solid #d1d5db;border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,0.12);padding:6px 10px;display:flex;align-items:center;gap:6px;width:320px;pointer-events:auto">
    <span style="color:#6b7280;font-size:13px">🔍</span>
    <input type="text" id="j26-search" placeholder="Search outlet code / name / beat" oninput="j26SetSearch(this.value)" style="flex:1;border:0;outline:none;font-size:12px;background:transparent">
    <span id="j26-search-count" style="font-size:10px;color:#9ca3af;white-space:nowrap"></span>
    <a href="javascript:void(0)" onclick="j26ClearSearch()" id="j26-search-clear" style="display:none;color:#dc2626;cursor:pointer;font-size:11px">×</a>
  </div>
  <div class="page-lbl">15 / 19 &middot; Aligned Beats &middot; Jun 2026 &middot; RS 218390 + 20B801</div>
  <div class="zoom-hint">Ctrl+Scroll or Pinch to zoom</div>
  <div class="panel" style="overflow:hidden;display:flex;flex-direction:column;padding:0;background:#fff;">
    <div style="padding:16px 18px 10px;flex:1;min-height:0;overflow-y:auto">
      <h2 style="margin-bottom:4px">Jun 2026 Aligned Beats</h2>
      <p class="p-sub" style="margin-bottom:8px">Merged 218390+20B801 &middot; OFM split Mon/Wed/Fri vs Tue/Thu/Sat &middot; mirror-aligned for delivery bundling</p>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:0 0 4px">View</div>
      <div class="toggle-row" style="margin-bottom:8px">
        <button class="t-btn active" id="j26-view-prop" onclick="j26SetView('proposed')">Proposed</button>
        <button class="t-btn" id="j26-view-exist" onclick="j26SetView('existing')">Existing</button>
      </div>
      <div class="kpi-r" id="j26-kpis"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:8px 0 4px">Color by</div>
      <div class="toggle-row" style="margin-bottom:8px">
        <button class="t-btn active" id="j26-cb-plg" onclick="j26SetCB('plg')">PLG</button>
        <button class="t-btn" id="j26-cb-day" onclick="j26SetCB('day')">Day</button>
      </div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:6px 0 4px">Filter by Day</div>
      <div class="filter-row" id="j26-day-chips" style="flex-wrap:wrap;gap:4px;margin-bottom:8px"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:8px 0 5px">Filter by PLG &amp; Salesman</div>
      <div id="j26-plg-tree"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:14px 0 4px">Same-Day Multi-Visits (vs prior designs)</div>
      <div id="j26-conflict-summary" style="font-size:11px;color:#374151"></div>
      <div style="margin-top:10px;border-top:1px solid #e5e7eb;padding-top:8px">
        <button class="dl-btn" onclick="j26Download()" style="background:#7030A0">
          &#8595; Download Aligned Beats CSV</button>
      </div>
    </div>
  </div>
</div>

<!-- SLIDE JUN26-CHANGES · Outlet & Visit Reduction Breakdown -->
<div class="slide" id="slide-jun26-changes" style="background:#f9fafb">
  <div style="position:absolute;inset:0;overflow-y:auto;padding:50px 60px 100px;box-sizing:border-box">
  <div class="page-lbl">16 / 19 &middot; Outlet &amp; Visit Reduction &middot; Jun 2026</div>
  <div style="max-width:1280px;width:100%;margin:0 auto">
    <h1 style="font-size:26px;font-weight:300;color:#0f172a;margin:0 0 2px">Where did the outlet changes come from?</h1>
    <div style="font-size:13px;color:#64748b;margin-bottom:14px">RS 218390 + 20B801 merged, ME BEAT existing vs Jun 2026 proposed</div>

    <!-- Headline KPIs -->
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:16px">
      <div style="background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:12px;text-align:center">
        <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.7px">Outlets</div>
        <div style="margin-top:6px"><span style="font-size:28px;font-weight:700;color:#374151">8,091</span> <span style="font-size:14px;color:#94a3b8">&rarr;</span> <span style="font-size:32px;font-weight:700;color:#0f172a">7,103</span></div>
        <div style="color:#16a34a;font-weight:700;font-size:16px;margin-top:4px">&darr; 988 (12.2%)</div>
      </div>
      <div style="background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:12px;text-align:center">
        <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.7px">Visits / week</div>
        <div style="margin-top:6px"><span style="font-size:28px;font-weight:700;color:#374151">18,988</span> <span style="font-size:14px;color:#94a3b8">&rarr;</span> <span style="font-size:32px;font-weight:700;color:#0f172a">15,747</span></div>
        <div style="color:#16a34a;font-weight:700;font-size:16px;margin-top:4px">&darr; 3,241 (17.1%)</div>
      </div>
      <div style="background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:12px;text-align:center">
        <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.7px">Avg visits / outlet</div>
        <div style="margin-top:6px"><span style="font-size:28px;font-weight:700;color:#374151">2.35</span> <span style="font-size:14px;color:#94a3b8">&rarr;</span> <span style="font-size:32px;font-weight:700;color:#0f172a">2.22</span></div>
        <div style="color:#16a34a;font-weight:700;font-size:16px;margin-top:4px">&darr; 0.13 (5.5%)</div>
      </div>
    </div>

    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
      <!-- Outlet changes table -->
      <div style="background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:14px">
        <h3 style="margin:0 0 8px;color:#0f172a;font-size:16px">Outlet Reconciliation</h3>
        <table class="dt-tbl" style="width:100%;font-size:13px">
          <thead><tr>
            <th style="text-align:left">Change</th><th>Outlets</th><th>Note</th>
          </tr></thead>
          <tbody>
            <tr><td style="text-align:left;font-weight:600">Starting (ME BEAT)</td><td><b>8,091</b></td><td style="text-align:left;color:#64748b">218390: 7,384 &middot; 20B801: 707</td></tr>
            <tr style="color:#dc2626"><td style="text-align:left">&minus; Pre-removed by client</td><td>&minus;635</td><td style="text-align:left;color:#64748b">DUMMY / inactive (not in working file)</td></tr>
            <tr style="color:#dc2626"><td style="text-align:left">&minus; Wholesale (WS) beats</td><td>&minus;248</td><td style="text-align:left;color:#64748b">Excluded from redesign</td></tr>
            <tr style="color:#dc2626"><td style="text-align:left">&minus; Moved to RS 218391 (OTR)</td><td>&minus;143</td><td style="text-align:left;color:#64748b">Outstanding-receivables transfer</td></tr>
            <tr style="color:#dc2626"><td style="text-align:left">&minus; Confirmed duplicates</td><td>&minus;93</td><td style="text-align:left;color:#64748b">Same physical store, two codes</td></tr>
            <tr style="color:#16a34a"><td style="text-align:left">+ New outlets added</td><td>+131</td><td style="text-align:left;color:#64748b">Not in ME BEAT, in master</td></tr>
            <tr style="background:#f9fafb;font-weight:700;color:#0f172a"><td style="text-align:left">Working count (Proposed)</td><td><b>7,103</b></td><td style="text-align:left;color:#64748b">Δ net &minus;988 outlets</td></tr>
          </tbody>
        </table>
      </div>

      <!-- Visit changes table -->
      <div style="background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:14px">
        <h3 style="margin:0 0 8px;color:#0f172a;font-size:16px">Visit Reduction Drivers</h3>
        <table class="dt-tbl" style="width:100%;font-size:13px">
          <thead><tr>
            <th style="text-align:left">Driver</th><th>Visits</th><th>Share</th>
          </tr></thead>
          <tbody>
            <tr><td style="text-align:left;font-weight:600">Starting visits (existing)</td><td><b>18,988</b></td><td></td></tr>
            <tr style="color:#dc2626"><td style="text-align:left">&minus; Outlet count reduction</td><td>&minus;2,319</td><td style="color:#dc2626;font-weight:700">72%</td></tr>
            <tr style="color:#dc2626"><td style="text-align:left">&minus; Visit-frequency consolidation</td><td>&minus;922</td><td style="color:#dc2626;font-weight:700">28%</td></tr>
            <tr style="background:#f9fafb;font-weight:700;color:#0f172a"><td style="text-align:left">Total reduction</td><td><b>&minus;3,257</b></td><td>100%</td></tr>
            <tr style="background:#f9fafb;font-weight:700;color:#0f172a"><td style="text-align:left">Proposed visits</td><td><b>15,747</b></td><td></td></tr>
          </tbody>
        </table>
        <div style="margin-top:12px;font-size:12px;color:#64748b;line-height:1.5">
          <b>Outlet-driven (71%):</b> 988 removed outlets × 2.35 existing avg visits.<br/>
          <b>Frequency-driven (29%):</b> Rule 9/10 consolidates 3-visit outlets to 2 visits (D+F+N | PP); F+N collapse merges some F + N pairs.
        </div>
      </div>
    </div>

    <!-- Visit distribution histogram -->
    <div style="background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:14px;margin-top:14px">
      <h3 style="margin:0 0 8px;color:#0f172a;font-size:16px">Outlets by Visit Count</h3>
      <table class="dt-tbl" style="width:100%;font-size:13px">
        <thead><tr>
          <th style="text-align:left">Visits / outlet</th><th>Existing</th><th>Proposed</th><th>Δ</th>
        </tr></thead>
        <tbody>
          <tr><td style="text-align:left;font-weight:600">1 visit</td><td>2,257</td><td>435</td><td style="color:#16a34a;font-weight:700">&minus;1,822</td></tr>
          <tr><td style="text-align:left;font-weight:600">2 visits</td><td>1,980</td><td><span style="color:#0369a1;font-weight:700">5,710</span></td><td style="color:#dc2626;font-weight:700">+3,730</td></tr>
          <tr><td style="text-align:left;font-weight:600">3 visits</td><td>2,789</td><td>0</td><td style="color:#16a34a;font-weight:700">&minus;2,789</td></tr>
          <tr><td style="text-align:left;font-weight:600">4 visits</td><td>921</td><td>898</td><td style="color:#94a3b8">&minus;23</td></tr>
          <tr><td style="text-align:left;font-weight:600">5 visits</td><td>144</td><td>60</td><td style="color:#16a34a">&minus;84</td></tr>
        </tbody>
      </table>
      <div style="margin-top:10px;font-size:12px;color:#64748b;line-height:1.5">
        The big shift: <b>3-visit outlets (-2,789) and 1-visit outlets (-1,822) consolidated into 2-visit (+3,730)</b> — driven by Rule 9/10 which sends most outlets to a single D+F+N visit + a PP visit. 4 and 5-visit categories stay roughly stable (premium PP-A/PP-B outlets and OFM channel stores).
      </div>
    </div>
  </div>
  </div>
</div>

<!-- SLIDE JUN26-TERR · Beat Territories (Jun 2026) -->
<div class="slide" id="slide-jun26-terr">
  <div class="map-wrap" id="map-jun26-terr"></div>
  <div class="page-lbl">17 / 19 &middot; Beat Territories &middot; Jun 2026</div>
  <div class="zoom-hint">Ctrl+Scroll to zoom</div>
  <div class="panel" style="overflow:hidden;display:flex;flex-direction:column;padding:0;background:#fff;">
    <div style="padding:16px 18px 10px;flex:1;min-height:0;overflow-y:auto">
      <h2 style="margin-bottom:4px">Beat Territories &amp; Overlap</h2>
      <p class="p-sub" style="margin-bottom:8px">Convex hull per PLG-salesman-day &middot; overlap visible across PLGs</p>
      <div class="toggle-row" style="margin-bottom:10px">
        <button class="t-btn active" id="jt-view-prop" onclick="jtSetView('proposed')">Proposed</button>
        <button class="t-btn" id="jt-view-exist" onclick="jtSetView('existing')">Existing</button>
      </div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:0 0 4px">Filter by day</div>
      <div class="filter-row" id="jt-day-chips" style="flex-wrap:wrap;gap:4px;margin-bottom:8px"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:6px 0 5px">Filter by PLG &amp; Salesman</div>
      <div id="jt-plg-list" style="max-height:240px;overflow-y:auto;font-size:11px"></div>
      <div class="kpi-r" id="jt-kpis" style="margin-top:8px"></div>
      <div id="jt-dist-table"></div>
      <div id="jt-area-table"></div>
    </div>
  </div>
</div>

<!-- (Slide 16 "Beat Area per Day Jun 2026" removed per user request) -->

<!-- SLIDE JUN26-DEL · Delivery Beats &amp; Truck Assignment (Jun 2026) -->
<div class="slide" id="slide-jun26-del">
  <div class="map-wrap" id="map-jun26-del"></div>
  <div style="position:absolute;top:14px;left:50%;transform:translateX(-50%);z-index:1000;background:white;border:1px solid #d1d5db;border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,0.12);padding:6px 10px;display:flex;align-items:center;gap:6px;width:320px;pointer-events:auto">
    <span style="color:#6b7280;font-size:13px">🔍</span>
    <input type="text" id="jd-search" placeholder="Search truck / beat / outlet code / name" oninput="jdSetSearch(this.value)" style="flex:1;border:0;outline:none;font-size:12px;background:transparent">
    <span id="jd-search-count" style="font-size:10px;color:#9ca3af;white-space:nowrap"></span>
    <a href="javascript:void(0)" onclick="jdClearSearch()" id="jd-search-clear" style="display:none;color:#dc2626;cursor:pointer;font-size:11px">×</a>
  </div>
  <div class="page-lbl">18 / 19 &middot; Delivery Beats &middot; Jun 2026</div>
  <div class="zoom-hint">Ctrl+Scroll to zoom</div>
  <div class="panel" style="overflow:hidden;display:flex;flex-direction:column;padding:0;background:#fff;">
    <div style="padding:16px 18px 10px;flex:1;min-height:0;overflow-y:auto">
      <h2 style="margin-bottom:4px">Delivery Beats</h2>
      <p class="p-sub" style="margin-bottom:8px">Truck assignment based on per-visit beat value</p>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:0 0 4px">View</div>
      <div class="toggle-row" style="margin-bottom:8px">
        <button class="t-btn active" id="jd-view-prop" onclick="jdSetView('proposed')">Proposed</button>
        <button class="t-btn" id="jd-view-exist" onclick="jdSetView('existing')">Existing (2 salesmen, D+1)</button>
      </div>
      <div class="kpi-r" id="jd-kpis"></div>
      <div id="jd-selection-bar" style="background:#f3f4f6;border:1px solid #e5e7eb;border-radius:6px;padding:6px 10px;margin:6px 0;font-size:11px;color:#6b7280;min-height:28px;display:flex;align-items:center;justify-content:space-between">
        <span id="jd-sel-text">No trips selected</span>
        <a href="javascript:void(0)" onclick="jdClearSelected()" id="jd-sel-clear" style="color:#dc2626;cursor:pointer;display:none">Clear</a>
      </div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:8px 0 4px">Color by</div>
      <div class="toggle-row" style="margin-bottom:8px">
        <button class="t-btn active" id="jd-cb-tt" onclick="jdSetCB('truck-type')">Truck type</button>
        <button class="t-btn" id="jd-cb-truck" onclick="jdSetCB('truck')">Truck no</button>
        <button class="t-btn" id="jd-cb-beat" onclick="jdSetCB('beat')">Beat in Truck</button>
      </div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:8px 0 4px">Filter by Delivery Day</div>
      <div class="filter-row" id="jd-day-chips" style="flex-wrap:wrap;gap:4px;margin-bottom:8px"></div>
      <div id="jd-truck-legend" style="display:flex;gap:14px;margin:6px 0 10px;flex-wrap:wrap;font-size:11px"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px">Per-day Summary</div>
      <table class="dt-tbl" style="font-size:11px;margin-bottom:10px">
        <thead><tr>
          <th style="text-align:left">Deliv Day</th><th>Trucks</th><th>Visits</th><th>Value (L)</th><th>3W</th><th>Ace</th><th>Split</th>
        </tr></thead>
        <tbody id="jd-tbody"></tbody>
      </table>
      <div id="jd-trip-table" style="max-height:400px;overflow-y:auto"></div>
      <div style="margin-top:12px;border-top:1px solid #e5e7eb;padding-top:10px">
        <button class="dl-btn" onclick="jdOpenDownload('proposed')" style="background:#7030A0">
          &#8595; Download Delivery Detail (Proposed)</button>
        <button class="dl-btn" onclick="jdOpenDownload('existing')" style="background:#1565C0;margin-top:6px">
          &#8595; Download Delivery Detail (Existing)</button>
      </div>
    </div>
  </div>
</div>

</div><!-- /#slides -->

<img src="data:image/png;base64,__SB_LOGO__" id="sb-logo" alt="Stackbox" style="position:fixed;top:14px;left:16px;height:24px;pointer-events:none;z-index:1000;transition:filter 0.3s">

<div id="nav-dots">
  <div class="dot" onclick="goTo(0)"></div>
  <div class="dot" onclick="goTo(1)"></div>
  <div class="dot" onclick="goTo(2)"></div>
  <div class="dot" onclick="goTo(3)"></div>
  <div class="dot" onclick="goTo(4)"></div>
  <div class="dot" onclick="goTo(5)"></div>
  <div class="dot" onclick="goTo(6)"></div>
  <div class="dot" onclick="goTo(7)"></div>
  <div class="dot" onclick="goTo(8)"></div>
  <div class="dot" onclick="goTo(9)"></div>
  <div class="dot" onclick="goTo(10)"></div>
  <div class="dot" onclick="goTo(11)"></div>
  <div class="dot" onclick="goTo(12)"></div>
  <div class="dot" onclick="goTo(13)"></div>
  <div class="dot" onclick="goTo(14)"></div>
  <div class="dot" onclick="goTo(15)"></div>
  <div class="dot" onclick="goTo(16)"></div>
  <div class="dot" onclick="goTo(17)"></div>
  <div class="dot" onclick="goTo(18)"></div>
</div>

<script>
__DATA_BLOCK__

// ── helpers ───────────────────────────────────────────────────────────────────
function fN(v){return(v||0).toLocaleString();}

// CARTO Voyager tiles — clean roads, minimal POI labels
const BASE_TILES=[
  'https://a.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}.png',
  'https://b.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}.png',
  'https://c.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}.png',
  'https://d.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}.png'];
const KOL_CENTER=[88.43,22.60];
const KOL_ZOOM=10;
const MAPS={};

// ── Pre-build GeoJSON ─────────────────────────────────────────────────────────
const _OUTLET_GJ={
  type:'FeatureCollection',
  features:OUTLETS.filter(o=>RS_INFO[o[2]]!=null).map(o=>({
    type:'Feature',
    geometry:{type:'Point',coordinates:[o[1],o[0]]},
    properties:{
      ri:o[2], ni:o[3],
      tp:RS_INFO[o[2]].type==='Pharma'?1:RS_INFO[o[2]].type==='WS'?2:0,
      color:RS_INFO[o[2]].color,
      ncolor:RS_INFO[o[3]]!=null?RS_INFO[o[3]].color:RS_INFO[o[2]].color,
      name:o[4]
    }
  }))
};
// Pre-group outlets by color|tp for efficient Canvas 2D batching
const _OL_GROUPS={};
_OUTLET_GJ.features.forEach(f=>{
  const k=f.properties.color+'|'+f.properties.tp;
  if(!_OL_GROUPS[k])_OL_GROUPS[k]={color:f.properties.color,tp:f.properties.tp,pts:[]};
  _OL_GROUPS[k].pts.push({lng:f.geometry.coordinates[0],lat:f.geometry.coordinates[1],nc:f.properties.ncolor});
});
const _OL_NGROUPS={};
_OUTLET_GJ.features.forEach(f=>{
  const k=f.properties.ncolor+'|'+f.properties.tp;
  if(!_OL_NGROUPS[k])_OL_NGROUPS[k]={color:f.properties.ncolor,tp:f.properties.tp,pts:[]};
  _OL_NGROUPS[k].pts.push({lng:f.geometry.coordinates[0],lat:f.geometry.coordinates[1]});
});

// ── RS marker helpers ─────────────────────────────────────────────────────────
function _rsPopupHTML(rs){
  const genLine=rs.gen_n>0
    ?'<br/><span style="color:#6b7280">General:</span> <b>'+fN(rs.gen_n)+'</b> &middot; MOC '+fN(rs.gen_moc):''
  const wsLine=rs.ws_n>0
    ?'<br/><span style="color:#6b7280">WS:</span> <b>'+fN(rs.ws_n)+'</b> &middot; MOC '+fN(rs.ws_moc):''
  return '<div style="font-size:12px">'
    +'<b>'+rs.name+'</b>'
    +'<span style="color:#9ca3af;font-size:10px;margin-left:6px">'+rs.code+'</span><br/>'
    +'<span style="color:#6b7280;font-size:11px">'+rs.type+'</span>'
    +'<div style="margin-top:6px;font-size:11px;line-height:1.8">'
    +'Total: <b>'+fN(rs.outlet_count)+'</b> &middot; MOC <b>'+fN(rs.moc)+'</b>'
    +genLine+wsLine
    +'</div></div>';
}
function _addRSMarkers(m,popup,rsFilter){
  const markers=[];
  RS_INFO.filter(r=>r.lat!=null&&r.lon!=null&&(!rsFilter||rsFilter.includes(r.code))).forEach(rs=>{
    const el=document.createElement('div');
    el.style.cssText='width:18px;height:18px;border-radius:50%;background:'+rs.color
      +';border:3px solid #fff;cursor:pointer;'
      +'box-shadow:0 2px 8px rgba(0,0,0,0.7),0 0 0 1px rgba(0,0,0,0.15);';
    new maplibregl.Marker({element:el,anchor:'center'}).setLngLat([rs.lon,rs.lat]).addTo(m);
    el.addEventListener('mouseenter',()=>{
      popup.setLngLat([rs.lon,rs.lat]).setHTML(_rsPopupHTML(rs)).addTo(m);
    });
    el.addEventListener('mouseleave',()=>popup.remove());
    markers.push({el,rs});
  });
  return markers;
}
function _filterRSMarkers(markers,typeFilter){
  markers.forEach(({el,rs})=>{
    el.style.display=(typeFilter===null||rs.type===typeFilter)?'':'none';
  });
}

// ── Cluster outlet lookup (snaps each outlet to 20m grid) ─────────────────────
const _CGRID_INV=5000; // 1/0.0002
function _gridKey(lat,lon){return Math.round(lat*_CGRID_INV)+'|'+Math.round(lon*_CGRID_INV);}
const _CL_IDX={};
CLUSTERS.forEach(c=>{_CL_IDX[_gridKey(c.lat,c.lon)]={i:c.i,n:c.n};});
const _OL_CL=OUTLETS.map(o=>{const k=_gridKey(o[0],o[1]);return _CL_IDX[k]||null;});

// Canvas 2D overlay factory — bypasses WebGL circle layers
function _makeOutletCanvas(m,dpr){
  const wrap=m.getContainer();
  const oc=document.createElement('canvas');
  oc.style.cssText='position:absolute;top:0;left:0;pointer-events:none;z-index:2;';
  const resize=()=>{
    oc.width=wrap.offsetWidth*dpr; oc.height=wrap.offsetHeight*dpr;
    oc.style.width=wrap.offsetWidth+'px'; oc.style.height=wrap.offsetHeight+'px';
  };
  resize(); wrap.appendChild(oc);
  m.on('resize',resize);
  return {canvas:oc,ctx:oc.getContext('2d'),resize};
}

function _drawGroups(m,ctx,oc,dpr,groups,tp_filter,radius_fn){
  const dpr2=dpr; ctx.clearRect(0,0,oc.width,oc.height);
  const z=m.getZoom();
  const r=radius_fn(z)*dpr2;
  const b=m.getBounds();
  const pad=0.03;
  const sLat=b.getSouth()-pad,nLat=b.getNorth()+pad,wLng=b.getWest()-pad,eLng=b.getEast()+pad;
  Object.values(groups).forEach(g=>{
    if(tp_filter!==null&&g.tp!==tp_filter)return;
    ctx.fillStyle=g.color; ctx.beginPath();
    g.pts.forEach(p=>{
      if(p.lat<sLat||p.lat>nLat||p.lng<wLng||p.lng>eLng)return;
      const pt=m.project([p.lng,p.lat]);
      ctx.moveTo(pt.x*dpr2+r,pt.y*dpr2);
      ctx.arc(pt.x*dpr2,pt.y*dpr2,r,0,Math.PI*2);
    });
    ctx.globalAlpha=0.85; ctx.fill();
  });
  ctx.globalAlpha=1;
}

const _RS_GJ={
  type:'FeatureCollection',
  features:RS_INFO.filter(r=>r.lat!=null&&r.lon!=null).map(r=>({
    type:'Feature',
    geometry:{type:'Point',coordinates:[r.lon,r.lat]},
    properties:{idx:r.idx,name:r.name,tp:r.type,color:r.color,cnt:r.outlet_count}
  }))
};

const _BEATS_BY_PLG={};
PLG_INFO.forEach(p=>{_BEATS_BY_PLG[p.idx]=[];});
BEATS_390.forEach(b=>{if(_BEATS_BY_PLG[b[2]])_BEATS_BY_PLG[b[2]].push(b);});

const MKT_COLORS=['#ef4444','#f97316','#84cc16','#22c55e','#3b82f6','#a855f7'];
const MKT_DAYS=['Mon','Tue','Wed','Thu','Fri','Sat'];

function colorMatchExpr(prop){
  const args=RS_INFO.flatMap(r=>[r.idx,r.color]);
  return['match',['get',prop],...args,'#9ca3af'];
}

// ── TITLE ─────────────────────────────────────────────────────────────────────
document.getElementById('title-stats').innerHTML=[
  [fN(STATS.total),'Active Outlets'],
  [STATS.rs_count,'Distributors'],
  [fN(DUPE_STATS.total),'Duplicate Pairs'],
  [fN(CLUSTER_ST.total),'Dense Clusters'],
].map(([v,l])=>'<div class="s-box"><div class="sv">'+v+'</div><div class="sl">'+l+'</div></div>').join('');

// ── MAP FACTORY ───────────────────────────────────────────────────────────────
function makeMap(mapId,loadFn,center,zoom){
  const el=document.getElementById(mapId);
  if(!el)return;
  let map;
  try{
    map=new maplibregl.Map({container:el,
      style:{version:8,sources:{},layers:[]},
      center:center||KOL_CENTER,zoom:zoom||KOL_ZOOM,
      scrollZoom:false,dragRotate:false,attributionControl:false});
  }catch(e){
    el.innerHTML='<div style="padding:20px;color:red">Map error: '+e+'</div>';
    return;
  }
  MAPS[mapId]={map,init:false};

  const doInit=()=>{
    if(MAPS[mapId].init)return;
    if(!map.isStyleLoaded()){setTimeout(doInit,50);return;}
    MAPS[mapId].init=true;
    map.resize();
    try{
      map.addSource('_base',{type:'raster',tiles:BASE_TILES,tileSize:256,
        attribution:'&copy; OpenStreetMap &copy; CARTO'});
      map.addLayer({id:'_base',type:'raster',source:'_base',minzoom:0,maxzoom:22});
    }catch(e){console.warn('[HUL] basemap',mapId,e);}
    try{loadFn(map);}
    catch(e){console.error('[HUL] loadFn err',mapId,e);}
    setTimeout(()=>map.resize(),300);
  };

  map.once('load',doInit);
  setTimeout(doInit,100);
  setTimeout(doInit,800);
  setTimeout(doInit,2500);

  el.addEventListener('wheel',e=>{
    if(e.ctrlKey||e.metaKey){
      e.preventDefault();e.stopPropagation();
      const d=-(e.deltaY)*(e.deltaMode===0?0.08:1.5);
      map.jumpTo({zoom:Math.max(5,Math.min(19,map.getZoom()+Math.max(-2,Math.min(2,d))))});
    }
  },{passive:false});
  el.addEventListener('touchmove',e=>{if(e.touches.length>=2)e.preventDefault();},{passive:false});
}

function flyTo(mapId,lon,lat){
  if(lon==null||lat==null)return;
  if(MAPS[mapId])MAPS[mapId].map.flyTo({center:[lon,lat],zoom:13,duration:700});
}
function resizeMap(id){const m=MAPS[id];if(m&&m.map)m.map.resize();}

// ── SLIDE 1 · OUTLETS & DISTRIBUTORS ─────────────────────────────────────────
let curFilter='ALL';
let selRS=new Set();
const _DPR=window.devicePixelRatio||1;

function initSlide1(){
  if(MAPS['map-1'])return;
  makeMap('map-1',m=>{
    const rsPopup=new maplibregl.Popup({closeButton:false,closeOnClick:false,offset:8,maxWidth:'240px'});
    const olPopup=new maplibregl.Popup({closeButton:false,closeOnClick:false,offset:6,maxWidth:'220px'});
    const {canvas:oc,ctx:ctx2}=_makeOutletCanvas(m,_DPR);
    MAPS['map-1']._oc=oc;

    function drawExcl(){
      ctx2.clearRect(0,0,oc.width,oc.height);
      const dpr=_DPR;
      const z=m.getZoom();
      const r=Math.max(2,2+(z-9)/(14-9)*7)*dpr;
      const b=m.getBounds();const pad=0.05;
      const sl=b.getSouth()-pad,nl=b.getNorth()+pad,wl=b.getWest()-pad,el=b.getEast()+pad;
      const tpF=curFilter==='General'?0:curFilter==='Pharma'?1:curFilter==='WS'?2:null;
      // Draw all normal outlets as light grey context layer
      ctx2.fillStyle='#9ca3af';ctx2.globalAlpha=0.2;
      ctx2.beginPath();
      OUTLETS.forEach(o=>{
        if(o[0]<sl||o[0]>nl||o[1]<wl||o[1]>el)return;
        const rs=RS_INFO[o[2]];if(!rs)return;
        if(tpF!==null){const tp=rs.type==='Pharma'?1:rs.type==='WS'?2:0;if(tp!==tpF)return;}
        if(selRS.size>0&&!selRS.has(o[2]))return;
        const pt=m.project([o[1],o[0]]);
        ctx2.moveTo(pt.x*dpr+r,pt.y*dpr);ctx2.arc(pt.x*dpr,pt.y*dpr,r,0,Math.PI*2);
      });
      ctx2.fill();
      // Draw excluded outlets on top with dashed lines to RS
      EXCL_OUTLETS.forEach(o=>{
        if(o[0]<sl||o[0]>nl||o[1]<wl||o[1]>el)return;
        const rs=RS_INFO.find(r=>r.code===o[3]);
        if(tpF!==null){
          if(rs){const tp=rs.type==='Pharma'?1:rs.type==='WS'?2:0;if(tp!==tpF)return;}
        }
        if(selRS.size>0&&rs&&!selRS.has(rs.idx))return;
        const pt=m.project([o[1],o[0]]);
        const ptx=pt.x*dpr,pty=pt.y*dpr;
        if(o[4]!==null&&o[5]!==null){
          const rpt=m.project([o[5],o[4]]);
          ctx2.strokeStyle='#ef4444';ctx2.lineWidth=1.5*dpr;ctx2.globalAlpha=0.4;
          ctx2.setLineDash([4*dpr,3*dpr]);
          ctx2.beginPath();ctx2.moveTo(ptx,pty);ctx2.lineTo(rpt.x*dpr,rpt.y*dpr);ctx2.stroke();
          ctx2.setLineDash([]);
        }
        ctx2.fillStyle='#ef4444';ctx2.globalAlpha=0.85;
        ctx2.beginPath();ctx2.arc(ptx,pty,4*dpr,0,Math.PI*2);ctx2.fill();
      });
      ctx2.globalAlpha=1;
    }

    function draw1(){
      if(_showExcl){drawExcl();return;}
      ctx2.clearRect(0,0,oc.width,oc.height);
      if(selRS.size===0)return;
      const tpF=curFilter==='General'?0:curFilter==='Pharma'?1:curFilter==='WS'?2:null;
      const filt1=curFilter==='ALL'?RS_INFO:RS_INFO.filter(r=>r.type===curFilter);
      if(filt1.every(r=>selRS.has(r.idx))){
        _drawGroups(m,ctx2,oc,_DPR,_OL_GROUPS,tpF,z=>Math.max(2,2+(z-9)/(14-9)*7));
      } else {
        const z=m.getZoom();const dpr=_DPR;
        const r=Math.max(2,2+(z-9)/(14-9)*7)*dpr;
        const b=m.getBounds();const pad=0.03;
        const sl=b.getSouth()-pad,nl=b.getNorth()+pad,wl=b.getWest()-pad,el=b.getEast()+pad;
        const byCol={};
        OUTLETS.forEach(o=>{
          if(!selRS.has(o[2]))return;
          const rs=RS_INFO[o[2]];if(!rs)return;
          if(tpF!==null){const tp=rs.type==='Pharma'?1:rs.type==='WS'?2:0;if(tp!==tpF)return;}
          if(o[0]<sl||o[0]>nl||o[1]<wl||o[1]>el)return;
          if(!byCol[rs.color])byCol[rs.color]=[];byCol[rs.color].push(o);
        });
        ctx2.globalAlpha=0.85;
        Object.entries(byCol).forEach(([col,pts])=>{
          ctx2.fillStyle=col;ctx2.beginPath();
          pts.forEach(o=>{
            const pt=m.project([o[1],o[0]]);
            ctx2.moveTo(pt.x*dpr+r,pt.y*dpr);ctx2.arc(pt.x*dpr,pt.y*dpr,r,0,Math.PI*2);
          });
          ctx2.fill();
        });
        ctx2.globalAlpha=1;
      }
    }
    MAPS['map-1']._draw=draw1;
    m.on('render',draw1);

    MAPS['map-1']._rsMarkers=_addRSMarkers(m,rsPopup);
    selectAllRS();

    // Outlet tooltip — separate popup so it doesn't conflict with RS marker popup
    m.on('mousemove',e=>{
      if(_showExcl||m.getZoom()<10){olPopup.remove();m.getCanvas().style.cursor='';return;}
      // suppress if mouse is over an RS HTML marker
      if(e.originalEvent.target.closest('.maplibregl-marker')){olPopup.remove();m.getCanvas().style.cursor='';return;}
      const b=m.getBounds(),pad=0.005;
      const sl=b.getSouth()-pad,nl=b.getNorth()+pad,wl=b.getWest()-pad,el=b.getEast()+pad;
      const tpF=curFilter==='General'?0:curFilter==='Pharma'?1:curFilter==='WS'?2:null;
      let best=null,bestD=18*18;
      OUTLETS.forEach(o=>{
        if(selRS.size===0||!selRS.has(o[2]))return;
        if(o[0]<sl||o[0]>nl||o[1]<wl||o[1]>el)return;
        const rs=RS_INFO[o[2]];if(!rs)return;
        if(tpF!==null){const tp=rs.type==='Pharma'?1:rs.type==='WS'?2:0;if(tp!==tpF)return;}
        const pt=m.project([o[1],o[0]]);
        const d=(pt.x-e.point.x)**2+(pt.y-e.point.y)**2;
        if(d<bestD){bestD=d;best={o,rs};}
      });
      if(best){
        const {o,rs}=best;
        m.getCanvas().style.cursor='pointer';
        const _ok=v=>v&&v!=='0'&&v!=='nan'&&v!=='None';
        const parts=[];
        if(_ok(o[7]))parts.push(o[7]);
        if(_ok(o[5]))parts.push(o[5]);
        if(_ok(o[8]))parts.push(o[8]);
        const chLine=parts.length?'<br/><span style="color:#6b7280;font-size:10px">'+parts.join(' &middot; ')+'</span>':'';
        const mocLine='<br/><span style="color:#6b7280;font-size:10px">MOC: <b>'+(+o[6]).toFixed(2)+'</b></span>';
        olPopup.setLngLat([o[1],o[0]])
          .setHTML('<div style="font-size:12px"><b>'+o[4]+'</b><br/>'
            +'<span style="color:'+rs.color+'">&#11044; '+rs.name+'</span>'
            +chLine+mocLine+'</div>')
          .addTo(m);
      } else {m.getCanvas().style.cursor='';olPopup.remove();}
    });
  });
  renderPanel1();
}

let _exclMarkers=null, _showExcl=false;
// Build a set of RS codes present in excluded outlets for RS marker filtering
const _EXCL_RS_CODES=new Set(EXCL_OUTLETS.map(o=>o[3]));

function toggleExcl(){
  _showExcl=!_showExcl;
  const btn=document.getElementById('excl-btn');
  btn.style.background=_showExcl?'#fee2e2':'';
  btn.style.borderColor=_showExcl?'#ef4444':'#fca5a5';
  btn.textContent=_showExcl?'Show All':'Excl.';
  const exclPanel=document.getElementById('p1-excl-panel');
  if(exclPanel)exclPanel.style.display=_showExcl?'':'none';
  const m=MAPS['map-1']&&MAPS['map-1'].map;
  if(!m)return;
  // Build excluded markers lazily
  if(!_exclMarkers){
    const pop=new maplibregl.Popup({closeButton:false,closeOnClick:false,offset:8,maxWidth:'220px'});
    _exclMarkers=EXCL_OUTLETS.map(o=>{
      const el=document.createElement('div');
      el.style.cssText='width:16px;height:16px;border-radius:50%;background:#ef4444;color:white;'
        +'font-size:11px;font-weight:900;display:flex;align-items:center;justify-content:center;'
        +'border:2.5px solid white;box-shadow:0 2px 6px rgba(0,0,0,0.6);cursor:pointer;z-index:10;';
      el.textContent='✕';
      new maplibregl.Marker({element:el,anchor:'center'}).setLngLat([o[1],o[0]]).addTo(m);
      el.addEventListener('mouseenter',()=>{
        const rs=RS_INFO.find(r=>r.code===o[3]);
        const distText=o[6]!=null?'<br/><b style="color:#ef4444">'+o[6]+' km from RS</b>':'';
        const rsText=rs
          ?'<span style="color:'+rs.color+'">&#11044; '+rs.name+'</span><br/>'
          :'<span style="color:#9ca3af">RS '+o[3]+'</span><br/>';
        pop.setLngLat([o[1],o[0]])
          .setHTML('<div style="font-size:12px"><b>'+o[2]+'</b><br/>'+rsText+distText
            +'<br/><span style="color:#9ca3af;font-size:10px">Incorrect GPS &middot; verify retag or remove</span></div>')
          .addTo(m);
      });
      el.addEventListener('mouseleave',()=>pop.remove());
      return el;
    });
  }
  // Apply type + selRS filter to excl markers visibility
  _exclMarkers.forEach((el,i)=>{
    if(!_showExcl){el.style.display='none';return;}
    const o=EXCL_OUTLETS[i];
    const rs=RS_INFO.find(r=>r.code===o[3]);
    const typeOk=curFilter==='ALL'||(rs&&rs.type===curFilter);
    const selOk=selRS.size===0||(rs&&selRS.has(rs.idx));
    el.style.display=(typeOk&&selOk)?'':'none';
  });
  // Show only RS markers that have excluded outlets AND match type + selRS filter
  const markers=MAPS['map-1']&&MAPS['map-1']._rsMarkers;
  if(markers)markers.forEach(({el,rs})=>{
    if(_showExcl){
      const selOk=selRS.size===0||selRS.has(rs.idx);
      el.style.display=(_EXCL_RS_CODES.has(rs.code)&&(curFilter==='ALL'||rs.type===curFilter)&&selOk)?'':'none';
    } else {
      const typeOk=curFilter==='ALL'||rs.type===curFilter;
      el.style.display=(typeOk&&(selRS.size===0||selRS.has(rs.idx)))?'':'none';
    }
  });
  // Canvas redraws via draw1 (drawExcl path when _showExcl=true)
  if(MAPS['map-1']&&MAPS['map-1']._draw)MAPS['map-1']._draw();
  if(!_showExcl)renderPanel1();
}

function setFilter(f){
  curFilter=f; selRS.clear();
  document.querySelectorAll('.f-chip').forEach((b,i)=>{
    b.classList.toggle('active',['ALL','General','Pharma','WS'][i]===f);
  });
  const markers=MAPS['map-1']&&MAPS['map-1']._rsMarkers;
  if(markers)markers.forEach(({el,rs})=>{
    if(_showExcl){
      el.style.display=(_EXCL_RS_CODES.has(rs.code)&&(f==='ALL'||rs.type===f))?'':'none';
    } else {
      const typeOk=f==='ALL'||rs.type===f;
      el.style.display=(typeOk&&(selRS.size===0||selRS.has(rs.idx)))?'':'none';
    }
  });
  if(_showExcl&&_exclMarkers){
    _exclMarkers.forEach((el,i)=>{
      const o=EXCL_OUTLETS[i];
      const rs=RS_INFO.find(r=>r.code===o[3]);
      const typeOk=f==='ALL'||(rs&&rs.type===f);
      const selOk=selRS.size===0||(rs&&selRS.has(rs.idx));
      el.style.display=(typeOk&&selOk)?'':'none';
    });
  }
  selectAllRS();
}

function toggleRS(idx){
  if(selRS.has(idx))selRS.delete(idx);else selRS.add(idx);
  const markers=MAPS['map-1']&&MAPS['map-1']._rsMarkers;
  if(markers){
    if(_showExcl){
      markers.forEach(({el,rs})=>{
        const typeOk=curFilter==='ALL'||rs.type===curFilter;
        const selOk=_EXCL_RS_CODES.has(rs.code)&&(selRS.size===0||selRS.has(rs.idx));
        el.style.display=(typeOk&&selOk)?'':'none';
      });
    } else {
      markers.forEach(({el,rs})=>{
        const typeOk=curFilter==='ALL'||rs.type===curFilter;
        el.style.display=(typeOk&&(selRS.size===0||selRS.has(rs.idx)))?'':'none';
      });
    }
  }
  if(_showExcl&&_exclMarkers){
    _exclMarkers.forEach((el,i)=>{
      const o=EXCL_OUTLETS[i];
      const rs=RS_INFO.find(r=>r.code===o[3]);
      const typeOk=curFilter==='ALL'||(rs&&rs.type===curFilter);
      const selOk=selRS.size===0||(rs&&selRS.has(rs.idx));
      el.style.display=(typeOk&&selOk)?'':'none';
    });
  }
  if(MAPS['map-1']&&MAPS['map-1']._draw)MAPS['map-1']._draw();
  renderPanel1();
}

function clearRSSelection(){
  selRS.clear();
  const markers=MAPS['map-1']&&MAPS['map-1']._rsMarkers;
  if(markers&&!_showExcl)_filterRSMarkers(markers,curFilter==='ALL'?null:curFilter);
  if(MAPS['map-1']&&MAPS['map-1']._draw)MAPS['map-1']._draw();
  renderPanel1();
}

function renderPanel1(){
  const filt=curFilter==='ALL'?RS_INFO:RS_INFO.filter(r=>r.type===curFilter);
  const activeFilt=selRS.size>0?filt.filter(r=>selRS.has(r.idx)):[];
  const totalOut=activeFilt.reduce((s,r)=>s+r.outlet_count,0);
  document.getElementById('p1-kpis').innerHTML=
    '<div class="kpi"><div class="kv">'+fN(selRS.size>0?totalOut:0)+'</div>'
   +'<div class="kl">Selected Outlets</div></div>'
   +'<div class="kpi"><div class="kv">'+filt.length+'</div><div class="kl">Distributors</div></div>';
  const sorted=[...filt].sort((a,b)=>b.outlet_count-a.outlet_count);
  const TB={General:'background:#e3f2fd;color:#1565C0',Pharma:'background:#e8f5e9;color:#2e7d32',
             WS:'background:#fff3e0;color:#e65100'};
  document.getElementById('p1-tb').innerHTML=sorted.map(rs=>{
    const isSel=selRS.size>0&&selRS.has(rs.idx);
    const chk='<input type="checkbox"'+(isSel?' checked':'')+' style="pointer-events:none;'
      +'accent-color:#1565C0;width:13px;height:13px;vertical-align:middle"/>';
    return'<tr style="cursor:pointer;'+(isSel?'background:#eff6ff;':'')+'" '
     +'onclick="toggleRS('+rs.idx+')">'
     +'<td style="padding:6px 4px 6px 0;text-align:center">'+chk+'</td>'
     +'<td><span class="dc" style="background:'+rs.color+'"></span>'
     +'<span class="rs-nm">'+rs.name+'</span>'
     +'<div style="font-size:10px;color:#9ca3af;margin-left:14px">'+rs.code+'</div></td>'
     +'<td><span class="rs-badge" style="'+(TB[rs.type]||'')+'">'+rs.type+'</span></td>'
     +'<td>'+fN(rs.outlet_count)+'</td>'
     +'<td>'+fN(rs.moc)+'</td></tr>';
  }).join('');
  const chkAll=document.getElementById('p1-sel-all-chk');
  if(chkAll){
    if(selRS.size===0){chkAll.checked=false;chkAll.indeterminate=false;}
    else if(filt.every(r=>selRS.has(r.idx))){chkAll.checked=true;chkAll.indeterminate=false;}
    else{chkAll.checked=false;chkAll.indeterminate=true;}
  }
}

function toggleSelectAllRS(){
  const filt=curFilter==='ALL'?RS_INFO:RS_INFO.filter(r=>r.type===curFilter);
  if(filt.every(rs=>selRS.has(rs.idx)))clearRSSelection();
  else selectAllRS();
}

function selectAllRS(){
  const filt=curFilter==='ALL'?RS_INFO:RS_INFO.filter(r=>r.type===curFilter);
  filt.forEach(rs=>selRS.add(rs.idx));
  const markers=MAPS['map-1']&&MAPS['map-1']._rsMarkers;
  if(markers&&!_showExcl)markers.forEach(({el,rs})=>{
    const typeOk=curFilter==='ALL'||rs.type===curFilter;
    el.style.display=typeOk?'':'none';
  });
  if(MAPS['map-1']&&MAPS['map-1']._draw)MAPS['map-1']._draw();
  renderPanel1();
}

function downloadExcl(){
  const hdr=['Outlet Name','RS Code','RS Name','Outlet Lat','Outlet Lon','RS Lat','RS Lon','Distance to RS (km)'];
  const rows=[hdr,...EXCL_OUTLETS.map(o=>{
    const rs=RS_INFO.find(r=>r.code===o[3]);
    return[o[2],o[3],rs?rs.name:'',o[0],o[1],
           o[4]!=null?o[4]:'',o[5]!=null?o[5]:'',o[6]!=null?o[6]:''];
  })];
  const csv=rows.map(r=>r.map(v=>'"'+String(v||'').replace(/"/g,'""')+'"').join(',')).join('\\r\\n');
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8;'}));
  a.download='hul_kolkata_excluded_outlets.csv';
  document.body.appendChild(a);a.click();document.body.removeChild(a);
}

function downloadV4Beats(){
  const cols=BEATS_V4.cols;
  const rows=BEATS_V4.rows;
  const csv=[cols.join(','),...rows.map(r=>r.map(v=>'"'+String(v||'').replace(/"/g,'""')+'"').join(','))].join('\\r\\n');
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8;'}));
  a.download='hul_kolkata_218390_v4_sales_beat.csv';
  document.body.appendChild(a);a.click();document.body.removeChild(a);
}

function p5DownloadExisting(){
  const MKT_DAYS=['Mon','Tue','Wed','Thu','Fri','Sat'];
  const dseInfo=DSE_INFO;
  const rows=EX_BEATS_390.filter(bt=>{
    if(curBeatDSEsNone)return false;
    if(curBeatPLGs.size>0&&!curBeatPLGs.has(PLG_INFO[bt[2]]?.name))return false;
    if(curBeatDSEs.size>0&&!curBeatDSEs.has(dseInfo[bt[4]]?.name))return false;
    return true;
  });
  const hdr=['lat','lon','plg','day','dse'];
  const lines=[hdr.join(','),...rows.map(bt=>{
    const plg=PLG_INFO[bt[2]]?.name||'';
    const day=bt[3]>=0&&bt[3]<6?MKT_DAYS[bt[3]]:'';
    const dse=dseInfo[bt[4]]?.name||'';
    return[bt[0],bt[1],'"'+plg+'"','"'+day+'"','"'+dse+'"'].join(',');
  })];
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([lines.join('\\r\\n')],{type:'text/csv;charset=utf-8;'}));
  a.download='existing_beats_218390.csv';
  document.body.appendChild(a);a.click();document.body.removeChild(a);
}
function downloadProposed(){
  function _hav(la1,lo1,la2,lo2){
    const R=6371,dLat=(la2-la1)*Math.PI/180,dLon=(lo2-lo1)*Math.PI/180;
    const a=Math.sin(dLat/2)**2+Math.cos(la1*Math.PI/180)*Math.cos(la2*Math.PI/180)*Math.sin(dLon/2)**2;
    return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
  }
  const hdr=['Outlet Code','Outlet Name','Old RS Code','Old RS Name','New RS Code','New RS Name',
             'primarychannel','Classification','Channel Program','MOC',
             'Old Dist (km)','New Dist (km)','Old Dist x MOC','New Dist x MOC','Notes'];
  const terOutlets=OUTLETS.filter(o=>{const rs=RS_INFO[o[2]];return rs&&rs.type===curTerType;});
  const dataRows=terOutlets.map(o=>{
    const oldRS=RS_INFO[o[2]],newRS=RS_INFO[o[3]];
    const moc=o[6]||0;
    const oldD=(oldRS&&oldRS.lat&&oldRS.lon)?_hav(o[0],o[1],oldRS.lat,oldRS.lon):null;
    const newD=(newRS&&newRS.lat&&newRS.lon)?_hav(o[0],o[1],newRS.lat,newRS.lon):null;
    return[o[9]||'',o[4],
           oldRS?oldRS.code:'',oldRS?oldRS.name:'',
           newRS?newRS.code:'',newRS?newRS.name:'',
           o[7]||'',o[5]||'',o[8]||'',moc,
           oldD!=null?oldD.toFixed(3):'',newD!=null?newD.toFixed(3):'',
           oldD!=null?(oldD*moc).toFixed(3):'',newD!=null?(newD*moc).toFixed(3):'',''];
  });
  if(curTerType==='Pharma'&&FLAGGED_PHARMA&&FLAGGED_PHARMA.length){
    FLAGGED_PHARMA.forEach(f=>{
      dataRows.push([f.code,f.name,f.existing_rs_code,f.existing_rs_name,'','',
                     f.channel||'',f.classification||'',f.channel_prog||'',f.moc||0,
                     '','','','','VERIFY LOCATION - excluded from territory calculations']);
    });
  }
  const rows=[hdr,...dataRows];
  const csv=rows.map(r=>r.map(v=>'"'+String(v||'').replace(/"/g,'""')+'"').join(',')).join('\\r\\n');
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8;'}));
  a.download='hul_kolkata_'+curTerType.toLowerCase()+'_proposed_plan.csv';
  document.body.appendChild(a);a.click();document.body.removeChild(a);
}

// ── SLIDE 2 · TERRITORY OVERLAPS ─────────────────────────────────────────────
let curView='existing', curTerType='General';
let selRS2=new Set();
let _s2ShowBoundary=true, _s2ShowRetailers=false;

function _applyTer2Filters(m){
  let f;
  if(selRS2.size===0){f=['literal',false];}
  else{const typeF=['==',['get','rs_type'],curTerType];
    f=['all',typeF,['in',['get','rs_idx'],['literal',[...selRS2]]]];}
  if(m.getLayer('ter-fill'))m.setFilter('ter-fill',f);
  if(m.getLayer('ter-line'))m.setFilter('ter-line',f);
}

function s2tgl(k){
  function _styleBtn(id,on){
    const b=document.getElementById(id);if(!b)return;
    b.style.background=on?'#1565C0':'';b.style.color=on?'#fff':'';b.style.borderColor=on?'#1565C0':'';
    b.classList.toggle('active',on);}
  if(k==='boundary'){
    _s2ShowBoundary=!_s2ShowBoundary;
    _styleBtn('p2-tg-boundary',_s2ShowBoundary);
    if(MAPS['map-2']&&MAPS['map-2']._draw)MAPS['map-2']._draw();
  }else if(k==='ret'){
    _s2ShowRetailers=!_s2ShowRetailers;
    _styleBtn('p2-tg-ret',_s2ShowRetailers);
    if(MAPS['map-2']&&MAPS['map-2']._draw)MAPS['map-2']._draw();
  }
}

function initSlide2(){
  if(MAPS['map-2'])return;
  makeMap('map-2',m=>{
    m.addSource('territories',{type:'geojson',data:BOUNDARIES});
    m.addLayer({id:'ter-fill',type:'fill',source:'territories',paint:{
      'fill-color':['coalesce',['get','color'],'#9ca3af'],'fill-opacity':0.12
    }});
    m.addLayer({id:'ter-line',type:'line',source:'territories',paint:{
      'line-color':['coalesce',['get','color'],'#9ca3af'],
      'line-width':2,'line-opacity':0.8
    }});
    _applyTer2Filters(m);
    const _rsm2=_addRSMarkers(m,new maplibregl.Popup({closeButton:false,closeOnClick:false,offset:8,maxWidth:'220px'}));
    _filterRSMarkers(_rsm2,curTerType);
    MAPS['map-2']._rsMarkers=_rsm2;

    const {canvas:oc,ctx:ctx2}=_makeOutletCanvas(m,_DPR);
    function _drawHulls2(){
      if(!_s2ShowBoundary)return;
      const hulls=curView==='existing'?HULL_RS_EX:HULL_RS_PROP;
      ctx2.save();
      hulls.forEach(h=>{
        const rs=RS_INFO[h.rs_idx];
        if(!rs||rs.type!==curTerType)return;
        const pts=h.points.map(p=>m.project([p[1],p[0]]));
        if(pts.length<3)return;
        ctx2.beginPath();
        pts.forEach((pt,i)=>{const x=pt.x*_DPR,y=pt.y*_DPR;i===0?ctx2.moveTo(x,y):ctx2.lineTo(x,y);});
        ctx2.closePath();
        ctx2.globalAlpha=0.13;ctx2.fillStyle=rs.color;ctx2.fill();
        ctx2.globalAlpha=0.9;ctx2.strokeStyle=rs.color;ctx2.lineWidth=2*_DPR;ctx2.stroke();
      });
      ctx2.globalAlpha=1;ctx2.restore();
    }
    function draw2(){
      ctx2.clearRect(0,0,oc.width,oc.height);
      if(!_s2ShowRetailers){_drawHulls2();return;}
      const tpF=curTerType==='General'?0:1;
      const filt=RS_INFO.filter(r=>r.type===curTerType);
      if(filt.every(r=>selRS2.has(r.idx))){
        const grps=curView==='proposed'?_OL_NGROUPS:_OL_GROUPS;
        _drawGroups(m,ctx2,oc,_DPR,grps,tpF,z=>Math.max(2,2+(z-9)/(14-9)*7));
      } else {
        const z=m.getZoom();const dpr=_DPR;
        const r=Math.max(2,2+(z-9)/(14-9)*7)*dpr;
        const b=m.getBounds();const pad=0.03;
        const sl=b.getSouth()-pad,nl=b.getNorth()+pad,wl=b.getWest()-pad,el=b.getEast()+pad;
        const byCol={};
        OUTLETS.forEach(o=>{
          const rsIdx=curView==='proposed'?o[3]:o[2];
          if(!selRS2.has(rsIdx))return;
          const rs=RS_INFO[rsIdx];if(!rs||rs.type!==curTerType)return;
          if(o[0]<sl||o[0]>nl||o[1]<wl||o[1]>el)return;
          if(!byCol[rs.color])byCol[rs.color]=[];byCol[rs.color].push(o);
        });
        ctx2.globalAlpha=0.85;
        Object.entries(byCol).forEach(([col,pts])=>{
          ctx2.fillStyle=col;ctx2.beginPath();
          pts.forEach(o=>{const pt=m.project([o[1],o[0]]);
            ctx2.moveTo(pt.x*dpr+r,pt.y*dpr);ctx2.arc(pt.x*dpr,pt.y*dpr,r,0,Math.PI*2);});
          ctx2.fill();
        });
        ctx2.globalAlpha=1;
      }
      _drawHulls2();
    }
    MAPS['map-2']._draw=draw2;
    m.on('render',draw2);

    const olPopup2=new maplibregl.Popup({closeButton:false,closeOnClick:false,offset:6,maxWidth:'260px'});
    m.on('mousemove',e=>{
      if(m.getZoom()<10){olPopup2.remove();m.getCanvas().style.cursor='';return;}
      if(e.originalEvent.target.closest('.maplibregl-marker')){olPopup2.remove();m.getCanvas().style.cursor='';return;}
      const b=m.getBounds(),pad=0.005;
      const sl=b.getSouth()-pad,nl=b.getNorth()+pad,wl=b.getWest()-pad,el=b.getEast()+pad;
      let best=null,bestD=18*18;
      OUTLETS.forEach(o=>{
        if(o[0]<sl||o[0]>nl||o[1]<wl||o[1]>el)return;
        const rsIdx=curView==='proposed'?o[3]:o[2];
        const rs=RS_INFO[rsIdx];if(!rs)return;
        if(rs.type!==curTerType)return;
        if(selRS2.size===0||!selRS2.has(rsIdx))return;
        const pt=m.project([o[1],o[0]]);
        const d=(pt.x-e.point.x)**2+(pt.y-e.point.y)**2;
        if(d<bestD){bestD=d;best={o,rs};}
      });
      if(best){
        const {o,rs}=best;
        m.getCanvas().style.cursor='pointer';
        const oldRS=RS_INFO[o[2]],newRS=RS_INFO[o[3]];
        const moved=o[2]!==o[3];
        const rsLine=moved
          ?'<span style="color:'+oldRS.color+'">&#11044; '+oldRS.name+'</span> &rarr; <span style="color:'+newRS.color+'">'+newRS.name+'</span>'
          :'<span style="color:'+rs.color+'">&#11044; '+rs.name+'</span>';
        const _ok=v=>v&&v!=='0'&&v!=='nan'&&v!=='None';
        const parts=[];if(_ok(o[7]))parts.push(o[7]);if(_ok(o[5]))parts.push(o[5]);if(_ok(o[8]))parts.push(o[8]);
        const chLine=parts.length?'<br/><span style="color:#6b7280;font-size:10px">'+parts.join(' &middot; ')+'</span>':'';
        const mocLine='<br/><span style="color:#6b7280;font-size:10px">MOC: <b>'+(+o[6]).toFixed(2)+'</b></span>';
        olPopup2.setLngLat([o[1],o[0]])
          .setHTML('<div style="font-size:12px"><b>'+o[4]+'</b><br/>'+rsLine+chLine+mocLine+'</div>')
          .addTo(m);
      } else {m.getCanvas().style.cursor='';olPopup2.remove();}
    });

    // Initialise with all RS selected so the map is populated
    selectAllRS2();
  });
  renderPanel2();
}

function setTerType(t){
  curTerType=t;selRS2.clear();
  document.getElementById('t-gen').classList.toggle('active',t==='General');
  document.getElementById('t-pha').classList.toggle('active',t==='Pharma');
  if(MAPS['map-2']&&MAPS['map-2']._rsMarkers)_filterRSMarkers(MAPS['map-2']._rsMarkers,t);
  selectAllRS2();
}

function setView(v){
  curView=v;
  if(MAPS['map-2']&&MAPS['map-2']._draw)MAPS['map-2']._draw();
  renderPanel2();
}

function renderPanel2(){
  document.getElementById('t-existing').classList.toggle('active',curView==='existing');
  document.getElementById('t-proposed').classList.toggle('active',curView==='proposed');
  const overlapEl=document.getElementById('p2-overlap-stats');
  if(overlapEl&&RS_OVERLAP){
    const t=curTerType;
    const d=RS_OVERLAP[t]||{};
    const exP=d.ex!=null?d.ex.toFixed(1)+'%':'&mdash;';
    const prP=d.prop!=null?d.prop.toFixed(1)+'%':'&mdash;';
    const exCol=d.ex==null?'#6b7280':d.ex>10?'#dc2626':d.ex>5?'#ca8a04':'#16a34a';
    const prCol=d.prop==null?'#6b7280':d.prop>10?'#dc2626':d.prop>5?'#ca8a04':'#16a34a';
    const active=curView==='existing';
    overlapEl.innerHTML='<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px">'
      +'<div style="padding:8px 10px;border-radius:8px;border:1.5px solid '+(active?'#1565C0':'#e5e7eb')+';background:'+(active?'#eff6ff':'#f9fafb')+'">'
      +'<div style="font-size:10px;color:#6b7280;margin-bottom:2px">Existing overlap</div>'
      +'<div style="font-size:20px;font-weight:800;color:'+exCol+'">'+exP+'</div>'
      +'</div>'
      +'<div style="padding:8px 10px;border-radius:8px;border:1.5px solid '+(!active?'#1565C0':'#e5e7eb')+';background:'+(!active?'#eff6ff':'#f9fafb')+'">'
      +'<div style="font-size:10px;color:#6b7280;margin-bottom:2px">Proposed overlap</div>'
      +'<div style="font-size:20px;font-weight:800;color:'+prCol+'">'+prP+'</div>'
      +'</div>'
      +'</div>'
      +'<div style="font-size:10px;color:#9ca3af;margin-top:4px">Overlap = double-covered area &divide; total footprint. '+(t)+' distributors.</div>';
  }
  const isProposed=curView==='proposed';
  const dlBtn=document.getElementById('p2-dl-btn');
  if(dlBtn)dlBtn.style.display=isProposed?'':'none';
  const colOl=document.getElementById('p2-col-ol');
  const colMoc=document.getElementById('p2-col-moc');
  if(colOl)colOl.textContent=isProposed?'Prop. Outlets':'Outlets';
  if(colMoc)colMoc.textContent=isProposed?'Prop. MOC':'MOC';
  const filt=RS_INFO.filter(r=>r.type===curTerType);
  const sorted=[...filt].sort((a,b)=>(isProposed?b.proposed_count-a.proposed_count:b.outlet_count-a.outlet_count));
  document.getElementById('p2-tb').innerHTML=sorted.map(rs=>{
    const isSel=selRS2.size===0?false:selRS2.has(rs.idx);
    const chk='<input type="checkbox"'+(isSel?' checked':'')+' style="pointer-events:none;'
      +'accent-color:#1565C0;width:13px;height:13px;vertical-align:middle"/>';
    let olCell,mocCell;
    if(isProposed){
      const gainN=rs.gained_n>0?'<span style="color:#16a34a;font-size:9px"> +'+fN(rs.gained_n)+'</span>':'';
      const lossN=rs.lost_n>0?'<span style="color:#dc2626;font-size:9px"> &minus;'+fN(rs.lost_n)+'</span>':'';
      const gainM=rs.gained_moc>0?'<span style="color:#16a34a;font-size:9px"> +'+rs.gained_moc.toFixed(1)+'</span>':'';
      const lossM=rs.lost_moc>0?'<span style="color:#dc2626;font-size:9px"> &minus;'+rs.lost_moc.toFixed(1)+'</span>':'';
      olCell=fN(rs.proposed_count)+gainN+lossN;
      mocCell=rs.proposed_moc.toFixed(1)+gainM+lossM;
    } else {
      olCell=fN(rs.outlet_count);
      mocCell=fN(rs.moc);
    }
    const ds=RS_DIST_STATS[String(rs.idx)];
    let distLine='';
    if(ds){
      function _distRow(exV2,propV2,label){
        if(isProposed){
          const d=(exV2!=null&&propV2!=null)?propV2-exV2:null;
          const dStr=d!=null?(d>=0?'+':'')+d.toFixed(2)+' km':'';
          const dCol=d!=null?(d<0?'#16a34a':'#dc2626'):'#9ca3af';
          return'<div style="font-size:9px;color:#9ca3af;margin-left:14px">'+label+': '
            +(exV2!=null?exV2.toFixed(2)+' km':'&mdash;')+' &rarr; '
            +(propV2!=null?propV2.toFixed(2)+' km':'&mdash;')
            +(dStr?'<span style="color:'+dCol+'"> ('+dStr+')</span>':'')+'</div>';
        }else{
          return'<div style="font-size:9px;color:#9ca3af;margin-left:14px">'+label+': '
            +(exV2!=null?exV2.toFixed(2)+' km':'&mdash;')+'</div>';
        }
      }
      distLine=_distRow(ds.ex,ds.prop,'dist');
      if(ds.ex_wt!=null||ds.prop_wt!=null)distLine+=_distRow(ds.ex_wt,ds.prop_wt,'wt dist');
    }
    return'<tr style="cursor:pointer;'+(isSel?'background:#eff6ff;':'')+'" '
     +'onclick="toggleRS2('+rs.idx+')">'
     +'<td style="padding:6px 4px 6px 0;text-align:center">'+chk+'</td>'
     +'<td><span class="dc" style="background:'+rs.color+'"></span>'
     +'<span class="rs-nm">'+rs.name+'</span>'
     +'<div style="font-size:10px;color:#9ca3af;margin-left:14px">'+rs.code+'</div>'
     +distLine+'</td>'
     +'<td>'+olCell+'</td>'
     +'<td>'+mocCell+'</td></tr>';
  }).join('');
  const chkAll=document.getElementById('p2-sel-all-chk');
  if(chkAll){
    if(selRS2.size===0){chkAll.checked=false;chkAll.indeterminate=false;}
    else if(filt.every(r=>selRS2.has(r.idx))){chkAll.checked=true;chkAll.indeterminate=false;}
    else{chkAll.checked=false;chkAll.indeterminate=true;}
  }
}

function _applyRS2Markers(){
  const markers=MAPS['map-2']&&MAPS['map-2']._rsMarkers;
  if(!markers)return;
  markers.forEach(({el,rs})=>{
    const typeOk=rs.type===curTerType;
    const selOk=selRS2.size===0||selRS2.has(rs.idx);
    el.style.display=(typeOk&&selOk)?'':'none';
  });
}

function toggleRS2(idx){
  if(selRS2.has(idx))selRS2.delete(idx);else selRS2.add(idx);
  const m=MAPS['map-2']&&MAPS['map-2'].map;
  if(m)_applyTer2Filters(m);
  _applyRS2Markers();
  if(MAPS['map-2']&&MAPS['map-2']._draw)MAPS['map-2']._draw();
  renderPanel2();
}

function clearRS2Selection(){
  selRS2.clear();
  const m=MAPS['map-2']&&MAPS['map-2'].map;
  if(m)_applyTer2Filters(m);
  _applyRS2Markers();
  if(MAPS['map-2']&&MAPS['map-2']._draw)MAPS['map-2']._draw();
  renderPanel2();
}

function selectAllRS2(){
  RS_INFO.filter(r=>r.type===curTerType).forEach(rs=>selRS2.add(rs.idx));
  const m=MAPS['map-2']&&MAPS['map-2'].map;
  if(m)_applyTer2Filters(m);
  _applyRS2Markers();
  if(MAPS['map-2']&&MAPS['map-2']._draw)MAPS['map-2']._draw();
  renderPanel2();
}

function toggleSelectAllRS2(){
  const filt=RS_INFO.filter(r=>r.type===curTerType);
  if(filt.every(rs=>selRS2.has(rs.idx)))clearRS2Selection();
  else selectAllRS2();
}

// ── SLIDE 3 · DUPLICATE OUTLETS ───────────────────────────────────────────────
let _sel3=-1;

function initSlide3(){
  if(MAPS['map-3'])return;
  makeMap('map-3',m=>{
    const popup=new maplibregl.Popup({closeButton:false,closeOnClick:false,offset:8,maxWidth:'240px'});
    const {canvas:oc3,ctx:ctx3}=_makeOutletCanvas(m,_DPR);
    function draw3(){
      ctx3.clearRect(0,0,oc3.width,oc3.height);
      const z=m.getZoom();
      const dpr=_DPR;
      const r=Math.max(2,1.5+(z-9)/(15-9)*7)*dpr;
      const b=m.getBounds();
      const pad=0.04;
      const sl=b.getSouth()-pad,nl=b.getNorth()+pad,wl=b.getWest()-pad,el=b.getEast()+pad;
      // Lines
      ctx3.strokeStyle='#f97316';ctx3.lineWidth=1.5*dpr;ctx3.globalAlpha=0.4;
      DUPE_PAIRS.forEach(p=>{
        if(p.la===p.lb&&p.loa===p.lob)return;
        if(p.la<sl||p.la>nl||p.loa<wl||p.loa>el)return;
        const a=m.project([p.loa,p.la]),bp=m.project([p.lob,p.lb]);
        ctx3.beginPath();ctx3.moveTo(a.x*dpr,a.y*dpr);ctx3.lineTo(bp.x*dpr,bp.y*dpr);ctx3.stroke();
      });
      // Point A (red)
      ctx3.fillStyle='#ef4444';ctx3.globalAlpha=0.85;ctx3.beginPath();
      DUPE_PAIRS.forEach(p=>{
        if(p.la<sl||p.la>nl||p.loa<wl||p.loa>el)return;
        const pt=m.project([p.loa,p.la]);
        ctx3.moveTo(pt.x*dpr+r,pt.y*dpr);ctx3.arc(pt.x*dpr,pt.y*dpr,r,0,Math.PI*2);
      });
      ctx3.fill();
      // Point B (orange)
      ctx3.fillStyle='#f97316';ctx3.beginPath();
      DUPE_PAIRS.forEach(p=>{
        if(p.lb<sl||p.lb>nl||p.lob<wl||p.lob>el)return;
        const pt=m.project([p.lob,p.lb]);
        ctx3.moveTo(pt.x*dpr+r,pt.y*dpr);ctx3.arc(pt.x*dpr,pt.y*dpr,r,0,Math.PI*2);
      });
      ctx3.fill();
      // Selection ring
      if(_sel3>=0&&_sel3<DUPE_PAIRS.length){
        const p=DUPE_PAIRS[_sel3];
        const rsel=r+3*dpr;
        ctx3.strokeStyle='#1565C0';ctx3.lineWidth=2.5*dpr;ctx3.globalAlpha=1;
        [[p.loa,p.la],[p.lob,p.lb]].forEach(([ln,la])=>{
          const pt=m.project([ln,la]);
          ctx3.beginPath();ctx3.arc(pt.x*dpr,pt.y*dpr,rsel,0,Math.PI*2);ctx3.stroke();
        });
      }
      ctx3.globalAlpha=1;
    }
    MAPS['map-3']._draw=draw3;
    m.on('render',draw3);
    // Hover tooltip via map events
    function findNearDupe(px){
      const b=m.getBounds();const pad=0.02;
      const sl=b.getSouth()-pad,nl=b.getNorth()+pad,wl=b.getWest()-pad,el=b.getEast()+pad;
      let best=null,bestD=16*16;
      DUPE_PAIRS.forEach((p,i)=>{
        if(p.la>=sl&&p.la<=nl&&p.loa>=wl&&p.loa<=el){
          const pt=m.project([p.loa,p.la]);
          const d=(pt.x-px.x)**2+(pt.y-px.y)**2;
          if(d<bestD){bestD=d;best={p,i,side:0};}
        }
        if(p.lb>=sl&&p.lb<=nl&&p.lob>=wl&&p.lob<=el){
          const pt=m.project([p.lob,p.lb]);
          const d=(pt.x-px.x)**2+(pt.y-px.y)**2;
          if(d<bestD){bestD=d;best={p,i,side:1};}
        }
      });
      return best;
    }
    m.on('mousemove',e=>{
      const nr=findNearDupe(e.point);
      if(nr){
        const {p,side}=nr;
        const [ln,la]=side===0?[p.loa,p.la]:[p.lob,p.lb];
        const nm=side===0?p.na:p.nb,vs=side===0?p.nb:p.na;
        m.getCanvas().style.cursor='pointer';
        popup.setLngLat([ln,la])
          .setHTML('<div style="font-size:12px"><b>'+nm+'</b><br/>'
            +'<span style="color:#9ca3af">vs '+vs+'</span><br/>'
            +'<span style="color:#ef4444">'+p.dist+'m &middot; RS '+p.rsa+'</span></div>')
          .addTo(m);
      } else {m.getCanvas().style.cursor='';popup.remove();}
    });
    m.on('click',e=>{const nr=findNearDupe(e.point);if(nr)highlightDupe(nr.i);});
  });
  renderPanel3();
}

function highlightDupe(i){
  _sel3=i;
  const m=MAPS['map-3']&&MAPS['map-3'].map;
  const p=DUPE_PAIRS[i];
  if(m)m.flyTo({center:[(p.loa+p.lob)/2,(p.la+p.lb)/2],zoom:18,duration:700});
  if(MAPS['map-3']&&MAPS['map-3']._draw)MAPS['map-3']._draw();
  document.querySelectorAll('.dupe-item').forEach((el,j)=>{
    el.style.background=j===i?'#fff7ed':'';
  });
}

function renderPanel3(){
  document.getElementById('p3-kpis').innerHTML=
    '<div class="kpi"><div class="kv">'+fN(DUPE_STATS.total)+'</div><div class="kl">Confirmed Pairs</div></div>'
   +'<div class="kpi"><div class="kv">'+DUPE_STATS.rs_aff+'</div><div class="kl">RS Affected</div></div>'
   +'<div class="kpi"><div class="kv">'+fN(DUPE_STATS.total)+'</div><div class="kl">Outlets to Remove</div></div>'
   +'<div class="kpi" title="'+DUPE_STATS.total+' outlets removed ÷ ~220 avg outlets/salesman"><div class="kv">~'+DUPE_STATS.saved+'</div><div class="kl">Salesmen (est.)</div></div>';
  document.getElementById('p3-list').innerHTML=DUPE_PAIRS.map((p,i)=>{
    const distTag='<span class="dupe-dist">'+p.dist+'m</span>';
    return'<div class="dupe-item" onclick="highlightDupe('+i+')">'
     +'<div class="d-na">'+p.na+distTag+'</div>'
     +'<div class="d-nb">vs '+p.nb+'</div>'
     +'<div class="d-meta">RS '+p.rsa+(p.rsb&&p.rsb!==p.rsa?' → '+p.rsb:'')+'</div>'
     +'</div>';
  }).join('');
}

function downloadDupes(){
  const hdr=['Name A','Name B','RS Code A','RS Code B','Distance (m)','AI Reason','Lat A','Lon A','Lat B','Lon B','Code A','Code B'];
  const rows=[hdr,...DUPE_PAIRS.map(p=>[p.na,p.nb,p.rsa,p.rsb,p.dist,p.reason,p.la,p.loa,p.lb,p.lob,p.ca,p.cb])];
  const csv=rows.map(r=>r.map(v=>'"'+String(v||'').replace(/"/g,'""')+'"').join(',')).join('\\r\\n');
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8;'}));
  a.download='hul_kolkata_duplicate_outlets.csv';
  document.body.appendChild(a);a.click();document.body.removeChild(a);
}

// ── SLIDE 4 · HIGH DENSITY CLUSTERS ──────────────────────────────────────────
let curDensity=5, selCluster=-1;

function clusterColor(n){
  if(n>=60)return'#7f1d1d';
  if(n>=35)return'#dc2626';
  if(n>=20)return'#f97316';
  if(n>=10)return'#fb923c';
  return'#fde68a';
}

function downloadClusters(){
  const filtered=CLUSTERS.filter(c=>c.n>=curDensity).sort((a,b)=>b.n-a.n);
  const hdr=['Cluster','Latitude','Longitude','Outlets'];
  const rows=[hdr,...filtered.map((c,i)=>[i+1,c.lat,c.lon,c.n])];
  const csv=rows.map(r=>r.join(',')).join('\\r\\n');
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));
  a.download='hul_kolkata_clusters.csv';
  document.body.appendChild(a);a.click();document.body.removeChild(a);
}

function initSlide4(){
  if(MAPS['map-4'])return;
  makeMap('map-4',m=>{
    const popup=new maplibregl.Popup({closeButton:false,closeOnClick:false,offset:10,maxWidth:'200px'});
    const {canvas:bgOc,ctx:bgCtx}=_makeOutletCanvas(m,_DPR);

    function draw4(){
      bgCtx.clearRect(0,0,bgOc.width,bgOc.height);
      const z=m.getZoom();
      const dpr=_DPR;
      const r=Math.max(3,3+(z-10)/(15-10)*7)*dpr;
      const b=m.getBounds();
      const pad=0.03;
      const sl=b.getSouth()-pad,nl=b.getNorth()+pad,wl=b.getWest()-pad,el=b.getEast()+pad;
      // Group outlets by color (cluster-colored vs grey)
      const byCol={};
      OUTLETS.forEach((o,i)=>{
        if(o[0]<sl||o[0]>nl||o[1]<wl||o[1]>el)return;
        const cl=_OL_CL[i];
        const active=cl&&cl.n>=curDensity;
        const col=active?clusterColor(cl.n):'#94a3b8';
        const alpha=active?0.82:0.13;
        const key=col+'|'+alpha;
        if(!byCol[key])byCol[key]={col,alpha,pts:[]};
        byCol[key].pts.push(o);
      });
      Object.values(byCol).forEach(g=>{
        bgCtx.fillStyle=g.col;bgCtx.globalAlpha=g.alpha;bgCtx.beginPath();
        g.pts.forEach(o=>{
          const pt=m.project([o[1],o[0]]);
          bgCtx.moveTo(pt.x*dpr+r,pt.y*dpr);bgCtx.arc(pt.x*dpr,pt.y*dpr,r,0,Math.PI*2);
        });
        bgCtx.fill();
      });
      bgCtx.globalAlpha=1;
    }
    MAPS['map-4']._draw=draw4;
    m.on('render',draw4);

    // Find nearest clustered outlet for hover/click
    function findNearCl(px){
      let best=null,bestD=20*20;
      const b=m.getBounds();const pad=0.02;
      const sl=b.getSouth()-pad,nl=b.getNorth()+pad,wl=b.getWest()-pad,el=b.getEast()+pad;
      OUTLETS.forEach((o,i)=>{
        const cl=_OL_CL[i];if(!cl||cl.n<curDensity)return;
        if(o[0]<sl||o[0]>nl||o[1]<wl||o[1]>el)return;
        const pt=m.project([o[1],o[0]]);
        const d=(pt.x-px.x)**2+(pt.y-px.y)**2;
        if(d<bestD){bestD=d;best={cl,lat:o[0],lon:o[1],name:o[4]};}
      });
      return best;
    }
    m.on('mousemove',e=>{
      const c=findNearCl(e.point);
      if(c){
        m.getCanvas().style.cursor='pointer';
        popup.setLngLat([c.lon,c.lat])
          .setHTML('<div style="font-size:12px"><b>'+c.name+'</b><br/>'
            +'<span style="color:#6b7280;font-size:11px">Cluster: '+c.cl.n+' outlets &middot; ~20m cell</span></div>').addTo(m);
      } else {m.getCanvas().style.cursor='';popup.remove();}
    });
    m.on('click',e=>{
      const c=findNearCl(e.point);
      if(c)selectCluster(c.cl.i);
    });
  });
  renderPanel4();
}

function setDensity(v){
  curDensity=v;
  document.getElementById('density-val').textContent=v;
  selCluster=-1;
  if(MAPS['map-4']&&MAPS['map-4']._draw)MAPS['map-4']._draw();
  renderPanel4();
}

function selectCluster(i){
  selCluster=i;
  const c=CLUSTERS.find(x=>x.i===i);
  if(!c)return;
  const m=MAPS['map-4']&&MAPS['map-4'].map;
  if(m)m.flyTo({center:[c.lon,c.lat],zoom:17,duration:700});
  if(MAPS['map-4']&&MAPS['map-4']._draw)MAPS['map-4']._draw();
  document.querySelectorAll('.cl-item').forEach(el=>{
    el.classList.toggle('sel',+el.dataset.idx===i);
  });
  const sel=document.querySelector('.cl-item[data-idx="'+i+'"]');
  if(sel)sel.scrollIntoView({block:'nearest',behavior:'smooth'});
}

function renderPanel4(){
  const filtered=CLUSTERS.filter(c=>c.n>=curDensity);
  const totalOutlets=filtered.reduce((s,c)=>s+c.n,0);
  document.getElementById('p4-kpis').innerHTML=
    '<div class="kpi"><div class="kv">'+fN(filtered.length)+'</div><div class="kl">Dense Areas</div></div>'
   +'<div class="kpi"><div class="kv">'+fN(totalOutlets)+'</div><div class="kl">Outlets</div></div>';
  document.getElementById('p4-meta').textContent=
    filtered.length+' clusters with '+curDensity+'+ outlets · '
    +fN(totalOutlets)+' outlets shown';
  const top=filtered.slice(0,80);
  const maxN=filtered.length>0?filtered[0].n:1;
  document.getElementById('p4-list').innerHTML=top.map((c,idx)=>{
    const pct=Math.round(c.n/maxN*100);
    const col=clusterColor(c.n);
    return'<div class="cl-item" data-idx="'+c.i+'" onclick="selectCluster('+c.i+')">'
     +'<div style="display:flex;justify-content:space-between;align-items:center">'
     +'<span style="font-size:12px;font-weight:700;color:#111827">#'+(idx+1)+
       ' <span style="color:'+col+'">&#9679;</span> '+c.n+' outlets</span>'
     +'<span style="font-size:10px;color:#9ca3af">'+c.lat.toFixed(3)+', '+c.lon.toFixed(3)+'</span>'
     +'</div>'
     +'<div style="height:3px;background:#f3f4f6;border-radius:2px;margin-top:5px">'
     +'<div style="height:3px;width:'+pct+'%;background:'+col+';border-radius:2px"></div>'
     +'</div>'
     +'</div>';
  }).join('')
  +(filtered.length>80?'<div style="padding:8px;font-size:11px;color:#9ca3af;text-align:center">Showing top 80 of '+fN(filtered.length)+'</div>':'');
}

// ── SLIDE 5 · BEATS ──────────────────────────────────────────────────────────
let curBeatsRS='218390', curBeatsView='proposed', curBeatDay='ALL';
let curBeatPLGs=new Set();     // empty+!PLGsNone=All; non-empty=specific set
let curBeatPLGsNone=false;     // proposed only: true=no PLGs shown (deselect all state)
let curBeatDSEs=new Set();     // empty+!DSEsNone=All; non-empty=specific set (existing only)
let curBeatDSEsNone=false;     // existing only: true=no salesmen shown
let colorBy='plg';             // 'plg' or 'day'
let p5ExpandedPLGs=new Set();  // PLG names with open accordion
const _SPEC_PLG_NAMES=new Set(['D-OFM','D_OFM','F-OFM','F+N_OFM','N_OFM','PP-A_OFM','PP-B_OFM','D+F_UNIGLOW','PP-A_UNIGLOW','PP-B_UNIGLOW']);

function _getBeats5(){
  if(curBeatsRS==='218390')return curBeatsView==='proposed'?BEATS_390:EX_BEATS_390;
  return curBeatsView==='proposed'?BEATS_391:EX_BEATS_391;
}
function _getBgBeats5(){return curBeatsRS==='218390'?BEATS_391:BEATS_390;}
function _getDseInfo5(){return(curBeatsRS==='218391'&&curBeatsView==='existing')?DSE_INFO_391:DSE_INFO;}
function _hasDay5(){return curBeatsRS==='218390'||(curBeatsRS==='218391'&&curBeatsView==='existing');}

// Pre-compute DSE name -> Set of PLG names for existing 218390
const _EX390_DSE_PLGS=(()=>{
  const m={};
  EX_BEATS_390.forEach(bt=>{
    const dn=DSE_INFO[bt[4]]?.name; const pn=PLG_INFO[bt[2]]?.name;
    if(dn&&pn){if(!m[dn])m[dn]=new Set();m[dn].add(pn);}
  });
  return m;
})();

function setBeatsRS(rs){
  curBeatsRS=rs;curBeatDay='ALL';curBeatPLGs=new Set();curBeatPLGsNone=false;curBeatDSEs=new Set();curBeatDSEsNone=false;p5ExpandedPLGs=new Set();
  document.getElementById('p5-rs390').classList.toggle('active',rs==='218390');
  document.getElementById('p5-rs391').classList.toggle('active',rs==='218391');
  buildBeatChips();
  if(MAPS['map-5']&&MAPS['map-5']._draw)MAPS['map-5']._draw();
  renderPanel5();
}
function setBeatsView(v){
  curBeatsView=v;curBeatDay='ALL';curBeatPLGs=new Set();curBeatPLGsNone=false;curBeatDSEs=new Set();curBeatDSEsNone=false;p5ExpandedPLGs=new Set();
  document.getElementById('p5-vproposed').classList.toggle('active',v==='proposed');
  document.getElementById('p5-vexisting').classList.toggle('active',v==='existing');
  buildBeatChips();
  if(MAPS['map-5']&&MAPS['map-5']._draw)MAPS['map-5']._draw();
  renderPanel5();
}

function initSlide5(){
  if(MAPS['map-5'])return;
  makeMap('map-5',m=>{
    _addRSMarkers(m,
      new maplibregl.Popup({closeButton:false,closeOnClick:false,offset:8,maxWidth:'220px'}),
      ['218390','218391']
    );
    const {canvas:oc5,ctx:ctx5}=_makeOutletCanvas(m,_DPR);
    function draw5(){
      ctx5.clearRect(0,0,oc5.width,oc5.height);
      const z=m.getZoom();const dpr=_DPR;
      const rBg=Math.max(1,1+(z-9)/(14-9)*3.5)*dpr;
      const rFg=Math.max(2,2+(z-9)/(14-9)*5)*dpr;
      const b=m.getBounds();const pad=0.03;
      const sl=b.getSouth()-pad,nl=b.getNorth()+pad,wl=b.getWest()-pad,el=b.getEast()+pad;
      ctx5.fillStyle='#94a3b8';ctx5.globalAlpha=0.35;ctx5.beginPath();
      _getBgBeats5().forEach(pt=>{
        if(pt[0]<sl||pt[0]>nl||pt[1]<wl||pt[1]>el)return;
        const p=m.project([pt[1],pt[0]]);
        ctx5.moveTo(p.x*dpr+rBg,p.y*dpr);ctx5.arc(p.x*dpr,p.y*dpr,rBg,0,Math.PI*2);
      });
      ctx5.fill();
      const hasDay=_hasDay5();
      const dayF=(!hasDay||curBeatDay==='ALL')?null:parseInt(curBeatDay);
      const dseInfo=_getDseInfo5();
      const isEx390draw=curBeatsRS==='218390'&&curBeatsView==='existing';
      const rows=_getBeats5().filter(bt=>{
        if(curBeatsView==='proposed'&&curBeatPLGsNone)return false;
        if(isEx390draw&&curBeatDSEsNone)return false;
        if(curBeatPLGs.size>0){
          if(!curBeatPLGs.has(PLG_INFO[bt[2]]?.name))return false;
          if(curBeatsView==='proposed'&&curBeatDSEs.size>0&&!curBeatDSEs.has(dseInfo[bt[4]]?.name))return false;
        }else{
          if(curBeatsView==='proposed'&&_SPEC_PLG_NAMES.has(PLG_INFO[bt[2]]?.name))return false;
          if(curBeatsView==='proposed'&&curBeatDSEs.size>0&&!curBeatDSEs.has(dseInfo[bt[4]]?.name))return false;
          if(isEx390draw&&curBeatDSEs.size>0&&!curBeatDSEs.has(dseInfo[bt[4]]?.name))return false;
        }
        if(dayF!==null&&bt[3]!==dayF)return false;
        return true;
      });
      const byCol={};
      rows.forEach(bt=>{
        const pi=PLG_INFO[bt[2]];
        const col=(colorBy==='day'&&hasDay&&bt[3]>=0)?MKT_COLORS[bt[3]]:(pi?pi.color:'#6b7280');
        if(!byCol[col])byCol[col]=[];byCol[col].push(bt);
      });
      ctx5.globalAlpha=0.85;
      Object.entries(byCol).forEach(([col,pts])=>{
        ctx5.fillStyle=col;ctx5.beginPath();
        pts.forEach(bt=>{
          if(bt[0]<sl||bt[0]>nl||bt[1]<wl||bt[1]>el)return;
          const p=m.project([bt[1],bt[0]]);
          ctx5.moveTo(p.x*dpr+rFg,p.y*dpr);ctx5.arc(p.x*dpr,p.y*dpr,rFg,0,Math.PI*2);
        });
        ctx5.fill();
      });
      ctx5.globalAlpha=1;
    }
    MAPS['map-5']._draw=draw5;
    m.on('render',draw5);
  },[88.25,22.50],11);
  buildBeatChips();
  renderPanel5();
}

function setColorBy(v){
  colorBy=v;
  document.getElementById('p5-cb-plg').classList.toggle('active',v==='plg');
  document.getElementById('p5-cb-day').classList.toggle('active',v==='day');
  buildBeatChips();
  if(MAPS['map-5']&&MAPS['map-5']._draw)MAPS['map-5']._draw();
}
function _p5PlgDseNames(plgName){
  const beats=_getBeats5(),dseInfo=_getDseInfo5();
  const pi=PLG_INFO.findIndex(p=>p.name===plgName);
  const s=new Set();
  beats.filter(bt=>bt[2]===pi).forEach(bt=>{const n=dseInfo[bt[4]]?.name;if(n)s.add(n);});
  return s;
}
function p5TogglePLG(plgName){
  const activePlgNames=new Set(_getBeats5().map(bt=>PLG_INFO[bt[2]]?.name).filter(Boolean));
  if(curBeatPLGsNone){
    // None mode → select just this PLG
    curBeatPLGsNone=false;curBeatPLGs=new Set([plgName]);curBeatDSEs=new Set();
    _p5PlgDseNames(plgName).forEach(dn=>curBeatDSEs.add(dn));
  }else if(curBeatPLGs.size===0){
    // All mode → deselect just this one
    curBeatPLGs=new Set([...activePlgNames].filter(n=>n!==plgName));
    curBeatDSEs=new Set();
    [...curBeatPLGs].forEach(pn=>_p5PlgDseNames(pn).forEach(dn=>curBeatDSEs.add(dn)));
    if(curBeatPLGs.size===0){curBeatPLGsNone=true;curBeatDSEs=new Set();}
  }else if(curBeatPLGs.has(plgName)){
    const dses=_p5PlgDseNames(plgName);
    curBeatPLGs.delete(plgName);
    dses.forEach(dn=>{
      const inOther=[...curBeatPLGs].some(op=>_p5PlgDseNames(op).has(dn));
      if(!inOther)curBeatDSEs.delete(dn);
    });
    if(curBeatPLGs.size===0){curBeatPLGsNone=true;curBeatDSEs=new Set();}
  }else{
    curBeatPLGs.add(plgName);
    _p5PlgDseNames(plgName).forEach(dn=>curBeatDSEs.add(dn));
    if([...activePlgNames].every(n=>curBeatPLGs.has(n))){curBeatPLGs=new Set();curBeatDSEs=new Set();}
  }
  buildBeatChips();
  if(MAPS['map-5']&&MAPS['map-5']._draw)MAPS['map-5']._draw();
}
function p5ClearPLGs(){
  if(!curBeatPLGsNone&&curBeatPLGs.size===0){
    curBeatPLGsNone=true;
  }else{
    curBeatPLGsNone=false;curBeatPLGs=new Set();curBeatDSEs=new Set();
  }
  buildBeatChips();
  if(MAPS['map-5']&&MAPS['map-5']._draw)MAPS['map-5']._draw();
}
function p5SetExistingPLG(plgName){
  curBeatDSEs=new Set();
  curBeatPLGs=plgName==='ALL'?new Set():new Set([plgName]);
  buildBeatChips();
  if(MAPS['map-5']&&MAPS['map-5']._draw)MAPS['map-5']._draw();
}
function p5ToggleExpand(plgName,ev){
  if(ev)ev.stopPropagation();
  if(p5ExpandedPLGs.has(plgName))p5ExpandedPLGs.delete(plgName);
  else p5ExpandedPLGs.add(plgName);
  buildBeatChips();
}
function _ex390VisibleDseNames(){
  const idxSet=new Set(EX_BEATS_390.map(bt=>bt[4]));
  return DSE_INFO.filter((d,i)=>{
    if(!idxSet.has(i))return false;
    if(curBeatPLGs.size===0)return true;
    const plgs=_EX390_DSE_PLGS[d.name];if(!plgs)return false;
    for(const pn of curBeatPLGs)if(plgs.has(pn))return true;
    return false;
  }).map(d=>d.name);
}
function _p5AllProposedDseNames(){
  const beats=_getBeats5(),dseInfo=_getDseInfo5();
  const idxSet=new Set(beats.map(bt=>bt[4]));
  return dseInfo.filter((_,i)=>idxSet.has(i)).map(d=>d.name);
}
function p5ToggleDSE(dseName){
  const isEx390=curBeatsRS==='218390'&&curBeatsView==='existing';
  const isProposedAll=curBeatsView==='proposed'&&curBeatPLGs.size===0&&!curBeatPLGsNone;
  if(isEx390&&curBeatDSEsNone){
    curBeatDSEsNone=false;curBeatDSEs=new Set([dseName]);
  }else if(isEx390&&curBeatDSEs.size===0){
    const all=_ex390VisibleDseNames();
    curBeatDSEs=new Set(all.filter(n=>n!==dseName));
    if(curBeatDSEs.size===0)curBeatDSEsNone=true;
  }else if(isProposedAll&&curBeatDSEs.size===0){
    const all=_p5AllProposedDseNames();
    curBeatDSEs=new Set(all.filter(n=>n!==dseName));
  }else if(curBeatDSEs.has(dseName)){
    curBeatDSEs.delete(dseName);
    if(isEx390&&curBeatDSEs.size===0)curBeatDSEsNone=true;
  }else{
    curBeatDSEs.add(dseName);
    if(isEx390){
      const all=_ex390VisibleDseNames();
      if(all.every(n=>curBeatDSEs.has(n))){curBeatDSEsNone=false;curBeatDSEs=new Set();}
    }else if(isProposedAll){
      const all=_p5AllProposedDseNames();
      if(all.every(n=>curBeatDSEs.has(n)))curBeatDSEs=new Set();
    }
  }
  buildBeatChips();
  if(MAPS['map-5']&&MAPS['map-5']._draw)MAPS['map-5']._draw();
}
function p5ResetDSEs(){
  if(curBeatsView==='proposed'){
    curBeatDSEs=new Set();
    [...curBeatPLGs].forEach(pn=>_p5PlgDseNames(pn).forEach(dn=>curBeatDSEs.add(dn)));
  }else{
    // Existing: toggle all-on vs all-off
    if(!curBeatDSEsNone&&curBeatDSEs.size===0){
      curBeatDSEsNone=true;
    }else{
      curBeatDSEsNone=false;curBeatDSEs=new Set();
    }
  }
  buildBeatChips();
  if(MAPS['map-5']&&MAPS['map-5']._draw)MAPS['map-5']._draw();
}
function _activateChip(selector,key,val,color){
  document.querySelectorAll(selector).forEach(b=>{
    const isA=b.dataset[key]===val;
    b.classList.toggle('active',isA);
    b.style.background=isA?(color||'#1565C0'):'';
    b.style.color=isA?'white':'';
    b.style.borderColor=isA?(color||'#1565C0'):'';
  });
}
function setBeatDay(day){
  curBeatDay=day;
  buildBeatChips();
  if(MAPS['map-5']&&MAPS['map-5']._draw)MAPS['map-5']._draw();
}

function renderPanel5(){
  const beats=_getBeats5();
  const rsLabel=curBeatsRS+' '+(curBeatsView==='proposed'?'Proposed':'Existing');
  const activePlgIdxSet=new Set(beats.map(bt=>bt[2]));
  const activePlgsFull=PLG_INFO.filter(p=>activePlgIdxSet.has(p.idx));
  const basePlgCount=activePlgsFull.filter(p=>!_SPEC_PLG_NAMES.has(p.name)).length;
  document.getElementById('p5-kpis').innerHTML=
    '<div class="kpi"><div class="kv">'+fN(beats.length)+'</div><div class="kl">Beat entries &middot; '+rsLabel+'</div></div>'
   +'<div class="kpi"><div class="kv">'+basePlgCount+'</div><div class="kl">PLGs</div></div>';
  const hasDay=_hasDay5();
  const isEx390=curBeatsRS==='218390'&&curBeatsView==='existing';
  const fl=document.getElementById('p5-filter-lbl');
  if(fl)fl.textContent=isEx390?'Filter by PLG':'Filter by PLG & Salesman';
  const ds=document.getElementById('p5-dse-section');
  if(ds)ds.style.display=isEx390?'':'none';
  document.getElementById('p5-colorby-section').style.display=hasDay?'':'none';
  document.getElementById('p5-day-section').style.display=hasDay?'':'none';
  const exDl=document.getElementById('p5-ex-dl');
  if(exDl)exDl.style.display=isEx390?'':'none';
}

function buildBeatChips(){
  const beats=_getBeats5();
  const dseInfo=_getDseInfo5();
  const activePlgIdxSet=new Set(beats.map(bt=>bt[2]));
  const activePlgs=PLG_INFO.filter(p=>activePlgIdxSet.has(p.idx));
  const isEx390=curBeatsRS==='218390'&&curBeatsView==='existing';

  if(isEx390){
    // Existing 218390: PLG filter chips + flat DSE list
    const allSel=curBeatPLGs.size===0;
    let html='<div class="filter-row" style="flex-wrap:wrap;gap:4px;margin-bottom:8px">';
    html+='<button class="beat-chip'+(allSel?' active':'')+'" data-n="ALL" '
      +'style="'+(allSel?'background:#1565C0;color:white;border-color:#1565C0;':'')+'" '
      +'onclick="p5SetExistingPLG(this.dataset.n)">All</button>';
    activePlgs.forEach(p=>{
      const isSel=curBeatPLGs.has(p.name);
      const col=p.color||'#1565C0';
      html+='<button class="beat-chip'+(isSel?' active':'')+'" data-n="'+p.name+'" '
        +'style="'+(isSel?'background:'+col+';color:white;border-color:'+col+';':'')+'" '
        +'onclick="p5SetExistingPLG(this.dataset.n)">'+p.name+'</button>';
    });
    html+='</div>';
    const treeEl=document.getElementById('p5-plg-tree');
    if(treeEl)treeEl.innerHTML=html;

    // Flat DSE list filtered by selected PLG chip
    const activeDseIdxSet=new Set(EX_BEATS_390.map(bt=>bt[4]));
    const visibleDses=DSE_INFO.filter((d,i)=>{
      if(!activeDseIdxSet.has(i))return false;
      if(curBeatPLGs.size===0)return true;
      const plgs=_EX390_DSE_PLGS[d.name];
      if(!plgs)return false;
      for(const pn of curBeatPLGs)if(plgs.has(pn))return true;
      return false;
    });
    const exAllDseOn=!curBeatDSEsNone&&curBeatDSEs.size===0;
    const exAllDsePartial=!curBeatDSEsNone&&curBeatDSEs.size>0;
    const exDseCb=exAllDseOn?'on':exAllDsePartial?'partial':'';
    const allDseRow='<div class="dse-item" onclick="p5ResetDSEs()" style="padding:3px 0;cursor:pointer;">'
      +'<div class="dse-cb '+exDseCb+'"></div>'
      +'<span class="dse-label" style="font-weight:700;color:'+(exAllDseOn||exAllDsePartial?'#1565C0':'#374151')+'">All Salesmen</span>'
      +'</div>';
    const dseHtml=allDseRow+visibleDses.map(d=>{
      const chk=(!curBeatDSEsNone&&curBeatDSEs.size===0)||curBeatDSEs.has(d.name);
      const plgs=_EX390_DSE_PLGS[d.name]||new Set();
      const tags=[...plgs].map(pn=>{
        const pi=PLG_INFO.find(p=>p.name===pn);
        const c=colorBy==='plg'?(pi?.color||'#6b7280'):'#9ca3af';
        return'<span class="dse-plg-tag" style="background:'+c+'22;color:'+c+';border:1px solid '+c+'44">'+pn+'</span>';
      }).join('');
      return'<div class="dse-item" data-n="'+d.name+'" onclick="p5ToggleDSE(this.dataset.n)">'
        +'<div class="dse-cb'+(chk?' on':'')+'"></div>'
        +'<span class="dse-label">'+d.name+'</span>'
        +(tags?'<span style="margin-left:4px;display:flex;gap:3px;flex-wrap:wrap;flex:1">'+tags+'</span>':'')
        +'</div>';
    }).join('');
    const dseListEl=document.getElementById('p5-dse-list');
    if(dseListEl)dseListEl.innerHTML=dseHtml||'<span style="color:#9ca3af;font-size:11px">No salesmen</span>';
  }else{
    // Proposed / 391: PLG accordion tree with checkboxes
    const normalPlgs=activePlgs.filter(p=>p.group==='normal'||p.group==='existing');
    const ofmPlgs=activePlgs.filter(p=>p.group==='ofm');
    const uniPlgs=activePlgs.filter(p=>p.group==='uniglow');

    function plgDses(plgName){
      const plgIdx=PLG_INFO.findIndex(p=>p.name===plgName);
      const idxSet=new Set(beats.filter(bt=>bt[2]===plgIdx).map(bt=>bt[4]));
      return dseInfo.filter((_,i)=>idxSet.has(i));
    }
    function makePlgItem(p){
      const isChk=curBeatPLGsNone?false:(curBeatPLGs.size===0||curBeatPLGs.has(p.name));
      const isOpen=p5ExpandedPLGs.has(p.name);
      const col=p.color||'#1565C0';
      const dses=plgDses(p.name);
      const cnt=dses.length;
      const allProposedMode=curBeatPLGs.size===0&&!curBeatPLGsNone;
      const dseHtml=isChk?dses.map(d=>{
        const chk=allProposedMode?(curBeatDSEs.size===0||curBeatDSEs.has(d.name)):curBeatDSEs.has(d.name);
        return'<div class="dse-item" data-n="'+d.name+'" onclick="p5ToggleDSE(this.dataset.n)">'
          +'<div class="dse-cb'+(chk?' on':'')+'"></div>'
          +'<span class="dse-label">'+d.name+'</span></div>';
      }).join(''):'';
      const allDseNames=dses.map(d=>d.name);
      const chkCount=allProposedMode&&curBeatDSEs.size===0?allDseNames.length:allDseNames.filter(n=>curBeatDSEs.has(n)).length;
      const plgCbClass=!isChk?'':(chkCount===allDseNames.length||allDseNames.length===0?'on':'partial');
      const dotHtml=colorBy==='plg'?'<div class="plg-dot" style="background:'+col+'"></div>':'';
      return'<div class="plg-item'+(isChk?' sel':'')+'">'
        +'<div class="plg-row">'
        +'<div class="plg-cb '+plgCbClass+'" data-n="'+p.name+'" onclick="p5TogglePLG(this.dataset.n)"></div>'
        +dotHtml
        +'<span class="plg-name">'+p.name+'</span>'
        +'<span class="plg-cnt">'+cnt+' Salesmen</span>'
        +(isChk&&cnt>0?'<span class="plg-chev'+(isOpen?' open':'')+'" data-n="'+p.name+'" onclick="p5ToggleExpand(this.dataset.n,event)">&#9660;</span>':'')
        +'</div>'
        +(isChk&&cnt>0&&isOpen?'<div class="dse-list open">'+dseHtml+'</div>':'')
        +'</div>';
    }

    const allPLGsOn=!curBeatPLGsNone&&curBeatPLGs.size===0;
    const allPLGsPartial=!curBeatPLGsNone&&curBeatPLGs.size>0;
    const allPLGsCb=allPLGsOn?'on':allPLGsPartial?'partial':'';
    let html='<div class="plg-tree">'
      +'<div class="plg-all-row'+((allPLGsOn||allPLGsPartial)?' sel':'')+'" onclick="p5ClearPLGs()">'
      +'<div class="plg-cb '+(allPLGsCb)+'"></div>'
      +'<span class="plg-name" style="color:'+((allPLGsOn||allPLGsPartial)?'#1565C0':'#374151')+'">All PLGs</span>'
      +'</div>';
    if(normalPlgs.length)normalPlgs.forEach(p=>html+=makePlgItem(p));
    if(ofmPlgs.length){html+='<div class="plg-tree-sec ofm">OFM</div>';ofmPlgs.forEach(p=>html+=makePlgItem(p));}
    if(uniPlgs.length){html+='<div class="plg-tree-sec uni">UNIGLOW</div>';uniPlgs.forEach(p=>html+=makePlgItem(p));}
    html+='</div>';
    const treeEl=document.getElementById('p5-plg-tree');
    if(treeEl)treeEl.innerHTML=html;
  }

  // Day chips — colored when colorBy==='day'
  const MKT_DAY_LABELS=['Mon','Tue','Wed','Thu','Fri','Sat'];
  const dayChipData=[{val:'ALL',label:'All',col:null},...MKT_DAY_LABELS.map((d,i)=>({val:String(i),label:d,col:MKT_COLORS[i]}))];
  const dayEl=document.getElementById('p5-day-chips');
  if(dayEl)dayEl.innerHTML=dayChipData.map(d=>{
    const isA=d.val===curBeatDay;
    const chipCol=colorBy==='day'&&d.col?d.col:'#1565C0';
    return'<button class="beat-chip'+(isA?' active':'')+'" data-day="'+d.val+'" '
      +'style="'+(isA?'background:'+chipCol+';color:white;border-color:'+chipCol+';':
        colorBy==='day'&&d.col?'border-color:'+d.col+'40;':'')+'" '
      +`onclick="setBeatDay('${d.val}')">`+d.label+'</button>';
  }).join('');
}
try{buildBeatChips();}catch(e){console.error('buildBeatChips init error:',e);}

// ── SLIDE 7 · SAME-DAY CONFLICTS ─────────────────────────────────────────────
let curS7View='v3';

function initSlide7(){
  if(MAPS['map-7'])return;
  makeMap('map-7',m=>{
    const {canvas:oc7,ctx:ctx7}=_makeOutletCanvas(m,_DPR);
    function draw7(){
      ctx7.clearRect(0,0,oc7.width,oc7.height);
      const conflicts=curS7View==='v3'?CONFLICTS_V3_390:CONFLICTS_EX_390;
      const z=m.getZoom();const dpr=_DPR;
      const r=Math.max(3,3+(z-9)/(14-9)*6)*dpr;
      const b=m.getBounds();const pad=0.03;
      const sl=b.getSouth()-pad,nl=b.getNorth()+pad,wl=b.getWest()-pad,el=b.getEast()+pad;
      const byDay={};
      conflicts.forEach(c=>{
        if(c[0]<sl||c[0]>nl||c[1]<wl||c[1]>el)return;
        const d=c[2];
        if(!byDay[d])byDay[d]=[];byDay[d].push(c);
      });
      Object.entries(byDay).forEach(([day,pts])=>{
        const color=MKT_COLORS[+day]||'#9ca3af';
        ctx7.fillStyle=color;ctx7.globalAlpha=0.85;ctx7.beginPath();
        pts.forEach(c=>{
          const pt=m.project([c[1],c[0]]);
          ctx7.moveTo(pt.x*dpr+r,pt.y*dpr);ctx7.arc(pt.x*dpr,pt.y*dpr,r,0,Math.PI*2);
        });
        ctx7.fill();
      });
      ctx7.globalAlpha=1;
    }
    MAPS['map-7']._draw=draw7;
    m.on('render',draw7);
  },[88.30,22.49],12);
  renderPanel7();
}

function setS7View(v){
  curS7View=v;
  document.getElementById('s7-vex').classList.toggle('active',v==='existing');
  document.getElementById('s7-vv3').classList.toggle('active',v==='v3');
  if(MAPS['map-7']&&MAPS['map-7']._draw)MAPS['map-7']._draw();
  renderPanel7();
}

function renderPanel7(){
  document.getElementById('p7-legend').innerHTML=
    MKT_COLORS.map((c,i)=>'<div class="rs-item" style="padding:4px 6px">'
      +'<div class="rs-dot" style="background:'+c+'"></div>'
      +'<span class="rs-name" style="font-size:11px">Market '+(i+1)+' &mdash; '+MKT_DAYS[i]+'</span>'
      +'</div>').join('');
}

// ── SLIDE 8 · PLG PURITY ─────────────────────────────────────────────────────
let _s8init=false;
function renderPanel8(){
  if(_s8init)return;_s8init=true;
  const examples=BENEFIT_STATS.plg_purity.impure_examples||[];
  document.getElementById('p8-impure-tbl').innerHTML=examples.map(r=>'<tr style="border-bottom:1px solid rgba(255,255,255,0.08)"><td style="padding:7px 6px;color:#e2e8f0;font-weight:600">'+r.dse+'</td><td style="padding:7px 6px;color:#fbbf24">'+r.plgs+'</td><td style="padding:7px 6px;text-align:right;color:#f87171;font-weight:700">'+r.n+'</td></tr>').join('');
}

// ── SLIDE 9 · BEAT TERRITORIES & OVERLAP ──────────────────────────────────────
let curJ9View='v3',curJ9Market=0;
let curJ9PLG='ALL';
let curJ9DSEs=new Set();

function _j9ActivePLGs(){
  const hulls=curJ9View==='v3'?HULL_V3_390:HULL_EX_390;
  const names=new Set(hulls.map(h=>h.plg));
  return PLG_INFO.filter(p=>names.has(p.name));
}

function _j9PlgDses(plgName){
  const hulls=curJ9View==='v3'?HULL_V3_390:HULL_EX_390;
  const dseSet=new Set(hulls.filter(h=>h.plg===plgName).map(h=>h.dse));
  return [...dseSet].sort();
}

function buildJ9Filters(){
  const activePLGs=_j9ActivePLGs();
  const allSel=curJ9PLG==='ALL';
  let html='<div class="filter-row" style="flex-wrap:wrap;gap:4px;margin-bottom:8px">';
  html+='<button class="beat-chip'+(allSel?' active':'')+'" data-pn="ALL" '
    +'style="'+(allSel?'background:#374151;color:white;border-color:#374151;':'')+'" '
    +'onclick="setJ9PLG(this.dataset.pn)">All</button>';
  activePLGs.forEach(p=>{
    const isSel=curJ9PLG===p.name;
    const col=p.color||'#1565C0';
    html+='<button class="beat-chip'+(isSel?' active':'')+'" data-pn="'+p.name+'" '
      +'style="'+(isSel?'background:'+col+';color:white;border-color:'+col+';':'')+'" '
      +'onclick="setJ9PLG(this.dataset.pn)">'+p.name+'</button>';
  });
  html+='</div>';
  document.getElementById('p9-plg-tree').innerHTML=html;

  const hulls=curJ9View==='v3'?HULL_V3_390:HULL_EX_390;
  const visibleDses=allSel
    ?[...new Set(hulls.filter(h=>!_SPEC_PLG_NAMES.has(h.plg)).map(h=>h.dse))].sort()
    :_j9PlgDses(curJ9PLG);
  const allDseOn=curJ9DSEs.size===0;
  const allDseRow='<div class="dse-item" onclick="j9ResetDSEs()" style="padding:3px 0;cursor:pointer;">'
    +'<div class="dse-cb '+(allDseOn?'on':'partial')+'"></div>'
    +'<span class="dse-label" style="font-weight:700;color:'+(allDseOn?'#1565C0':'#374151')+'">All Salesmen</span>'
    +'</div>';
  const dseRows=visibleDses.map(d=>{
    const chk=curJ9DSEs.size===0||curJ9DSEs.has(d);
    const plgOfDse=hulls.find(h=>h.dse===d)?.plg||'';
    const pi=PLG_INFO.find(p=>p.name===plgOfDse);
    const tag=pi?'<span class="dse-plg-tag" style="background:'+pi.color+'22;color:'+pi.color+';border:1px solid '+pi.color+'44">'+pi.name+'</span>':'';
    return'<div class="dse-item" data-n="'+d+'" onclick="j9ToggleDSE(this.dataset.n)">'
      +'<div class="dse-cb'+(chk?' on':'')+'"></div>'
      +'<span class="dse-label">'+d+'</span>'
      +(allSel&&tag?'<span style="margin-left:4px">'+tag+'</span>':'')
      +'</div>';
  }).join('');
  const dseEl=document.getElementById('p9-dse-list');
  if(dseEl)dseEl.innerHTML=allDseRow+dseRows;

  const days=[{l:'All',v:0},{l:'Mon',v:1},{l:'Tue',v:2},{l:'Wed',v:3},{l:'Thu',v:4},{l:'Fri',v:5},{l:'Sat',v:6}];
  document.getElementById('p9-day-chips').innerHTML=days.map(d=>{
    const on=curJ9Market===d.v;
    const st=on?'background:#374151;color:white;border-color:#374151;':'';
    return'<button class="beat-chip'+(on?' active':'')+'" style="'+st+'" data-mv="'+d.v+'" onclick="setJ9Market(+this.dataset.mv)">'+d.l+'</button>';
  }).join('');
}

function setJ9PLG(name){
  curJ9PLG=name;curJ9DSEs=new Set();
  buildJ9Filters();renderJaccard9();
}

function j9ToggleDSE(name){
  const hulls=curJ9View==='v3'?HULL_V3_390:HULL_EX_390;
  const allDses=curJ9PLG==='ALL'
    ?[...new Set(hulls.filter(h=>!_SPEC_PLG_NAMES.has(h.plg)).map(h=>h.dse))].sort()
    :_j9PlgDses(curJ9PLG);
  if(curJ9DSEs.has(name)){curJ9DSEs.delete(name);}
  else{curJ9DSEs.add(name);if(allDses.every(n=>curJ9DSEs.has(n)))curJ9DSEs=new Set();}
  buildJ9Filters();renderJaccard9();
}

function j9ResetDSEs(){curJ9DSEs=new Set();buildJ9Filters();renderJaccard9();}

function initSlide9(){
  if(MAPS['leaf-9'])return;
  const mapEl=document.getElementById('l9-map');
  const lmap=L.map(mapEl,{zoomControl:true,preferCanvas:true}).setView([22.52,88.36],12);
  L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png',
    {attribution:'&copy; OpenStreetMap &copy; CARTO',subdomains:'abcd',maxZoom:19,opacity:.9}).addTo(lmap);
  lmap.on('wheel',e=>{
    if(e.originalEvent&&(e.originalEvent.ctrlKey||e.originalEvent.metaKey)){
      e.originalEvent.preventDefault();
    }
  });
  MAPS['leaf-9']={map:lmap,lg:L.layerGroup().addTo(lmap)};
  setTimeout(()=>lmap.invalidateSize(),200);
  document.getElementById('j9-vv3').textContent='Proposed — '+HULL_V3_390.length+' beats';
  document.getElementById('j9-vex').textContent='Existing — '+HULL_EX_390.length+' beats';
  buildJ9Filters();renderJaccard9();
}

function setJ9View(v){
  curJ9View=v;curJ9Market=0;curJ9PLG='ALL';curJ9DSEs=new Set();
  document.getElementById('j9-vv3').classList.toggle('active',v==='v3');
  document.getElementById('j9-vex').classList.toggle('active',v==='existing');
  document.getElementById('j9-vv3').textContent='Proposed — '+HULL_V3_390.length+' beats';
  document.getElementById('j9-vex').textContent='Existing — '+HULL_EX_390.length+' beats';
  buildJ9Filters();renderJaccard9();
}

function setJ9Market(m){
  curJ9Market=m;
  document.querySelectorAll('#p9-day-chips .beat-chip').forEach(b=>{
    const on=+b.dataset.mv===m;
    b.classList.toggle('active',on);
    b.style.cssText=on?'background:#374151;color:white;border-color:#374151;':'';
  });
  renderJaccard9();
}

function renderJaccard9(){
  const state=MAPS['leaf-9'];if(!state)return;
  state.lg.clearLayers();
  const hulls=curJ9View==='v3'?HULL_V3_390:HULL_EX_390;
  let drawn=0;const bnds=[];
  const _J9_DNAMES=['','Mon','Tue','Wed','Thu','Fri','Sat'];
  hulls.forEach(h=>{
    if(curJ9Market!==0&&h.market!==curJ9Market)return;
    if(curJ9PLG!=='ALL'&&h.plg!==curJ9PLG)return;
    if(curJ9PLG==='ALL'&&_SPEC_PLG_NAMES.has(h.plg))return;
    if(curJ9DSEs.size>0&&!curJ9DSEs.has(h.dse))return;
    const pts=h.hull.map(p=>[p[0],p[1]]);
    L.polygon(pts,{
      color:'#374151',weight:1.5,fillColor:'#374151',fillOpacity:0.06
    }).bindTooltip((h.plg?h.plg+' - ':'')+h.dse+' - '+_J9_DNAMES[h.market]+' - '+h.n+' outlets',
      {sticky:true,direction:'top'}).addTo(state.lg);
    bnds.push(...pts);
    drawn++;
  });
  if(bnds.length>0)state.map.fitBounds(bnds,{padding:[20,20],maxZoom:14});
  const jacRow=BENEFIT_STATS.jaccard.by_plg||[];
  const jv3=(curJ9View==='v3');
  document.getElementById('p9-kpis').innerHTML='<div class="kpi"><div class="kv">'+drawn+'</div><div class="kl">'+(jv3?'Proposed':'Existing')+' beats shown</div></div>'
    +'<div class="kpi" style="background:'+(jv3?'#f0fdf4':'#fff7f7')+'"><div class="kv" style="color:'+(jv3?'#16a34a':'#dc2626')+'">'+(jv3?'~0%':'22%&ndash;52%')+'</div><div class="kl">Avg Overlap</div></div>';
  document.getElementById('p9-jac-body').innerHTML=jacRow.map(r=>'<tr><td style="text-align:left">'+r.ex_plg+' &rarr; <b>'+r.v3_plg+'</b></td>'
    +'<td style="color:#dc2626;font-weight:700">'+(r.ex_jac*100).toFixed(2)+'%</td>'
    +'<td style="color:#16a34a;font-weight:700">'+(r.v3_jac*100).toFixed(2)+'%</td>'
    +'</tr>').join('');
  renderBeatDists9();
}

function renderBeatDists9(){
  const el=document.getElementById('p9-dist-table');if(!el)return;
  if(!BEAT_DIST||!BEAT_DIST.v3)return;
  const v3=BEAT_DIST.v3||[],ex=BEAT_DIST.ex||[];
  const mktF=curJ9Market===0?null:curJ9Market;
  const plgF=curJ9PLG==='ALL'?null:curJ9PLG;
  const avg=(arr,k)=>{const v=arr.filter(d=>d[k]!=null).map(d=>d[k]);return v.length?v.reduce((a,b)=>a+b,0)/v.length:null};
  const fmt=v=>v==null?'&mdash;':v.toFixed(1)+' km';
  // Use Jaccard PLG mapping: ex_plg → v3_plg
  const jacMap=BENEFIT_STATS.jaccard.by_plg||[];
  const showJac=plgF?jacMap.filter(r=>(curJ9View==='v3'?r.v3_plg:r.ex_plg)===plgF):jacMap;
  const plgRows=showJac.map(r=>{
    const exD=ex.filter(d=>d.plg===r.ex_plg&&(mktF?d.market===mktF:true));
    const v3D=v3.filter(d=>d.plg===r.v3_plg&&(mktF?d.market===mktF:true));
    const ec=avg(exD,'chain_km'),vc=avg(v3D,'chain_km');
    const dc=vc!=null&&ec!=null?vc-ec:null;
    const col=dc==null?'#6b7280':dc<0?'#16a34a':'#dc2626';
    return'<tr><td style="text-align:left">'+r.ex_plg+' &rarr; <b>'+r.v3_plg+'</b></td>'
      +'<td>'+fmt(ec)+'</td><td>'+fmt(vc)+'</td>'
      +'<td style="color:'+col+'">'+(dc==null?'&mdash;':(dc<0?'':'+')+dc.toFixed(1))+'</td>'
      +'</tr>';
  }).join('');
  // Overall averages (exclude specialists from v3)
  const allV3=v3.filter(d=>!_SPEC_PLG_NAMES.has(d.plg)&&(mktF?d.market===mktF:true));
  const allEx=ex.filter(d=>(mktF?d.market===mktF:true));
  const eC=avg(allEx,'chain_km'),dC=avg(allV3,'chain_km');
  const dChain=dC!=null&&eC!=null?dC-eC:null;
  const colC=dChain==null?'#6b7280':dChain<0?'#16a34a':'#dc2626';
  // OFM and UNIGLOW specialist rows (proposed only — no existing equivalent)
  // plg_info chip names differ from beat_distances PLG names in two cases
  const _CHIP_TO_DIST={'D_OFM':'D-OFM','F+N_OFM':'F-OFM'};
  const _OFM_CHIPS=new Set(['D_OFM','F+N_OFM','N_OFM','PP-A_OFM','PP-B_OFM']);
  const _OFM_DIST=new Set(['D-OFM','F-OFM','N_OFM','PP-A_OFM','PP-B_OFM']);
  const _UNI_PLG=new Set(['D+F_UNIGLOW','PP-A_UNIGLOW','PP-B_UNIGLOW']);
  const showOFM=!plgF||_OFM_CHIPS.has(plgF);
  const showUNI=!plgF||_UNI_PLG.has(plgF);
  let specRows='';
  if(showOFM){
    let ofmD;
    if(plgF&&_OFM_CHIPS.has(plgF)){
      const distName=_CHIP_TO_DIST[plgF]||plgF;
      ofmD=v3.filter(d=>d.plg===distName&&(mktF?d.market===mktF:true));
    } else {
      ofmD=v3.filter(d=>_OFM_DIST.has(d.plg)&&(mktF?d.market===mktF:true));
    }
    const ofmKm=avg(ofmD,'chain_km');
    const lbl=(plgF&&_OFM_CHIPS.has(plgF))?plgF:'OFM (all)';
    if(ofmKm!=null)specRows+='<tr style="background:#faf5ff"><td style="text-align:left;color:#7c3aed;font-weight:600">&mdash; &rarr; <b>'+lbl+'</b></td>'
      +'<td style="color:#9ca3af">&mdash;</td><td>'+fmt(ofmKm)+'</td><td style="color:#9ca3af">&mdash;</td></tr>';
  }
  if(showUNI){
    const uniD=v3.filter(d=>(plgF&&_UNI_PLG.has(plgF)?d.plg===plgF:_UNI_PLG.has(d.plg))&&(mktF?d.market===mktF:true));
    const uniKm=avg(uniD,'chain_km');
    const lbl=(plgF&&_UNI_PLG.has(plgF))?plgF:'UNIGLOW (all)';
    if(uniKm!=null)specRows+='<tr style="background:#eff6ff"><td style="text-align:left;color:#0369a1;font-weight:600">&mdash; &rarr; <b>'+lbl+'</b></td>'
      +'<td style="color:#9ca3af">&mdash;</td><td>'+fmt(uniKm)+'</td><td style="color:#9ca3af">&mdash;</td></tr>';
  }
  el.innerHTML='<div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:10px 0 2px">In-Beat Route Distance (km/market day)</div>'
    +'<div style="font-size:10px;color:#9ca3af;margin-bottom:4px">Distance a salesman travels within their beat on one market day. Not a round trip from distributor.</div>'
    +'<table class="dt-tbl" style="width:100%"><thead><tr>'
    +'<th style="text-align:left">Ex PLG &rarr; Prop</th><th>Existing</th><th>Proposed</th><th>&Delta;</th>'
    +'</tr></thead><tbody>'
    +(plgF?'':'<tr style="font-weight:700;background:#f9fafb"><td style="text-align:left">All (regular)</td>'
      +'<td>'+fmt(eC)+'</td><td>'+fmt(dC)+'</td>'
      +'<td style="color:'+colC+'">'+(dChain==null?'&mdash;':(dChain<0?'':'+')+dChain.toFixed(1))+'</td>'
      +'</tr>')
    +plgRows
    +specRows
    +'</tbody></table>';
}

// ── SLIDE 11 · PLG RULES ─────────────────────────────────────────────────────
// PLG Rules data from PLG Rules.xlsx (Rules sheet)
const _PLG11_RULES=[
  {ex:'OFM Store (any visits)',    nr:'5 Visits (D, F, N, PP-A, PP-B)', n:46,   note:'12 stores/day, exclusive salesman'},
  {ex:'1 Visit',                  nr:'2 Visits (D+F+N, PP)',           n:1953,  note:''},
  {ex:'2 Visits (PP-A + PP-B)',   nr:'2 Visits (PP-A, PP-B)',          n:96,    note:''},
  {ex:'2 Visits (PP + DETS)',     nr:'2 Visits (D+F+N, PP)',           n:147,   note:''},
  {ex:'2 Visits (PP + D+F+N)',    nr:'2 Visits (D+F+N, PP)',           n:936,   note:''},
  {ex:'2 Visits (PP + FNB+NUTS)', nr:'2 Visits (D+F+N, PP)',           n:12,    note:''},
  {ex:'2 Visits (DETS + FNB+NUTS / NUTS / FNB)', nr:'2 Visits (D+F, N)', n:327, note:''},
  {ex:'2 Visits (FNB + NUTS)',    nr:'2 Visits (F, N)',                n:43,    note:''},
  {ex:'3 Visits (PP-A + PP-B + anything)', nr:'4 Visits (D, F+N, PP-A, PP-B)', n:162, note:'Channel program stores served separately'},
  {ex:'3 Visits (Non-Food)',      nr:'2 Visits (D+F+N, PP)',           n:2574,  note:''},
  {ex:'3 Visits (Food)',          nr:'2 Visits (D+F, N)',              n:25,    note:''},
  {ex:'4 Visits (PP-A + PP-B + anything)', nr:'4 Visits (D, F+N, PP-A, PP-B)', n:856, note:''},
  {ex:'4 Visits (non PP-A/PP-B)', nr:'4 Visits (D, F, N, PP)',        n:62,    note:''},
  {ex:'5 Visits',                 nr:'5 Visits (D, F, N, PP-A, PP-B)',n:144,   note:''},
  {ex:'Uniglow',                  nr:'3 Visits (D+F, PP-A, PP-B)',    n:null,  note:'15-24 outlets/day; fill with HNB stores'},
  {ex:'Unicare',                  nr:'2 Visits (D+F+N, PP)',          n:null,  note:''},
];

let _s11init=false;
function renderSlide11(){
  if(_s11init)return;_s11init=true;
  const rs='border-bottom:1px solid rgba(255,255,255,0.08)';
  const td='padding:6px 8px;color:#e2e8f0;font-size:12px;';
  const tdR=td+'text-align:right;';
  document.getElementById('p11-rules-tbl').innerHTML=_PLG11_RULES.map((r,i)=>{
    const bg=i%2===0?'':'background:rgba(255,255,255,0.03);';
    const nrParts=r.nr.match(/\(([^)]+)\)/g)||[];
    const plgBadges=nrParts.map(p=>{
      const names=p.slice(1,-1).split(',').map(s=>s.trim());
      return names.map(n=>{
        const c=PLG_INFO.find(x=>x.name===n)?.color||'#60a5fa';
        return '<span style="display:inline-block;background:'+c+'22;color:'+c+';border:1px solid '+c+'55;padding:1px 5px;border-radius:8px;font-size:10px;font-weight:700;margin:1px">'+n+'</span>';
      }).join('');
    }).join('');
    const nrText=r.nr.replace(/\([^)]+\)/g,'').trim();
    return '<tr style="'+rs+bg+'">'
      +'<td style="'+td+'">'+r.ex+'</td>'
      +'<td style="'+td+'">'+nrText+' '+plgBadges+'</td>'
      +'<td style="'+tdR+'">'+(r.n!=null?r.n.toLocaleString():'&mdash;')+'</td>'
      +'<td style="'+td+';color:#94a3b8;font-size:11px">'+r.note+'</td>'
      +'</tr>';
  }).join('');
}

// ── SLIDE 12 · BEAT AREA — DELIVERY ZONE ────────────────────────────────────
const _A12_ZONE_COLORS=['#ef4444','#f97316','#eab308','#22c55e','#3b82f6','#a855f7'];
const _REV_DFN_MAP={5:1,2:2,4:3,1:4,3:5,6:6};
const _GROUP_A_PLGS=new Set(['D+F+N','D','D+F','F','PP-A','D_OFM','PP-A_OFM','PP-A_UNIGLOW','D+F_UNIGLOW']);
let curA12View='v3',curA12Zone=0;

function _getOrigZone(plg,calDay){
  if(_GROUP_A_PLGS.has(plg))return _REV_DFN_MAP[calDay]||calDay;
  const ga=((calDay-2+6)%6)+1;
  return _REV_DFN_MAP[ga]||ga;
}

function initSlide12(){
  if(MAPS['leaf-12'])return;
  const mapEl=document.getElementById('l12-map');
  const lmap=L.map(mapEl,{zoomControl:true,preferCanvas:true}).setView([22.52,88.36],12);
  L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png',
    {attribution:'&copy; OpenStreetMap &copy; CARTO',subdomains:'abcd',maxZoom:19,opacity:.9}).addTo(lmap);
  lmap.on('wheel',e=>{if(e.originalEvent&&(e.originalEvent.ctrlKey||e.originalEvent.metaKey))e.originalEvent.preventDefault();});
  MAPS['leaf-12']={map:lmap,lg:L.layerGroup().addTo(lmap)};
  setTimeout(()=>lmap.invalidateSize(),200);
  const chips=[{label:'All',z:0},...[1,2,3,4,5,6].map(z=>({label:'Z'+z,z}))];
  document.getElementById('a12-zone-chips').innerHTML=chips.map(c=>{
    const a=c.z===0;
    return '<button class="beat-chip'+(a?' active':'')+'" style="'+(a?'background:#374151;color:white;border-color:#374151;':'')+'" onclick="setA12Zone('+c.z+')">'+c.label+'</button>';
  }).join('');
  renderArea12();
}

function setA12View(v){
  curA12View=v;
  document.getElementById('a12-vv3').classList.toggle('active',v==='v3');
  document.getElementById('a12-vex').classList.toggle('active',v==='existing');
  renderArea12();
}

function setA12Zone(z){
  curA12Zone=z;
  document.querySelectorAll('#a12-zone-chips .beat-chip').forEach((el,i)=>{
    const a=i===z;
    el.classList.toggle('active',a);
    el.style.background=a?'#374151':'';el.style.color=a?'white':'';el.style.borderColor=a?'#374151':'';
  });
  renderArea12();
}

function renderArea12(){
  const state=MAPS['leaf-12'];if(!state)return;
  state.lg.clearLayers();
  const zones=(DELIVERY_ZONES&&DELIVERY_ZONES.zones)||[];
  const beatHulls=curA12View==='existing'?HULL_EX_390:HULL_V3_390;
  const filtZones=curA12Zone===0?zones:zones.filter(z=>z.zone===curA12Zone);
  const activeZoneNums=new Set(filtZones.map(z=>z.zone));
  const bnds=[];

  // Individual beat hulls (light fill)
  beatHulls.forEach(h=>{
    const oz=curA12View==='existing'?h.market:_getOrigZone(h.plg,h.market);
    if(!activeZoneNums.has(oz))return;
    if(!h.hull||h.hull.length<3)return;
    const col=_A12_ZONE_COLORS[oz-1];
    const pts=h.hull.map(p=>[p[0],p[1]]);
    L.polygon(pts,{color:col,weight:0.8,fillColor:col,fillOpacity:0.18,interactive:true})
      .bindTooltip(h.plg+' • '+h.dse+' • '+(h.area_km2||0).toFixed(2)+' km²',{sticky:true,direction:'top'})
      .addTo(state.lg);
    bnds.push(...pts);
  });

  // Combined zone hull (dashed outline on top)
  filtZones.forEach((z,i)=>{
    const hull=curA12View==='existing'?z.ex_hull:z.v4_hull;
    if(!hull||hull.length<3)return;
    const col=_A12_ZONE_COLORS[z.zone-1];
    const pts=hull.map(p=>[p[0],p[1]]);
    const area=curA12View==='existing'?z.ex_area:z.v4_area;
    const dayLabel=curA12View==='existing'?'Day '+z.zone:'Days '+z.group_a_day+'+'+z.group_b_day;
    L.polygon(pts,{color:col,weight:2.5,fillOpacity:0,dashArray:'7,5',interactive:true})
      .bindTooltip('Zone '+z.zone+' ('+dayLabel+') total: '+area+' km²',{sticky:true,direction:'top'})
      .addTo(state.lg);
    if(curA12Zone!==0)bnds.push(...pts);
  });

  if(bnds.length>0)state.map.fitBounds(bnds,{padding:[20,20],maxZoom:14});

  // KPIs — show selected zone or overall total
  const kZones=curA12Zone===0?zones:filtZones;
  const totV4=kZones.reduce((s,z)=>s+z.v4_area,0);
  const totEx=kZones.reduce((s,z)=>s+z.ex_area,0);
  const pct=totEx>0?Math.round((totV4-totEx)/totEx*100):0;
  const pctCol=pct<0?'#16a34a':'#dc2626';
  document.getElementById('p12-kpis').innerHTML=
    '<div class="kpi" style="border:1.5px solid #fee2e2"><div class="kv" style="color:#dc2626">'+totEx.toFixed(0)+' km²</div>'
   +'<div class="kl">Existing'+(curA12Zone?(' Z'+curA12Zone):' (total)')+'</div></div>'
   +'<div class="kpi" style="border:1.5px solid #ede9fe"><div class="kv" style="color:#7030A0">'+totV4.toFixed(0)+' km²</div>'
   +'<div class="kl">Proposed'+(curA12Zone?(' Z'+curA12Zone):' (total)')+' <span style="color:'+pctCol+'">('+(pct<0?'':'+')+pct+'%)</span></div></div>';

  // Bar chart (always show all 6 zones; highlight selected)
  const maxV=Math.max(...zones.map(z=>Math.max(z.v4_area,z.ex_area)),1);
  const barH=70;
  document.getElementById('p12-chart').innerHTML='<div style="display:grid;grid-template-columns:repeat(6,1fr);gap:6px;align-items:end">'
    +zones.map((z)=>{
      const col=_A12_ZONE_COLORS[z.zone-1];
      const eV=z.ex_area,pV=z.v4_area;
      const eH=Math.max(3,Math.round(eV/maxV*barH));
      const pH=Math.max(3,Math.round(pV/maxV*barH));
      const delta=eV>0?Math.round((pV-eV)/eV*100):0;
      const dc=delta<0?'#16a34a':'#dc2626';
      const dim=curA12Zone!==0&&curA12Zone!==z.zone?'opacity:0.3;':'';
      return '<div style="text-align:center;cursor:pointer;'+dim+'" onclick="setA12Zone('+(curA12Zone===z.zone?0:z.zone)+')">'
        +'<div style="font-size:9px;font-weight:700;color:#6b7280;margin-bottom:2px">'+eV.toFixed(0)+'</div>'
        +'<div style="display:flex;gap:3px;align-items:flex-end;justify-content:center;height:'+barH+'px">'
          +'<div style="width:16px;height:'+eH+'px;background:#fca5a5;border-radius:2px 2px 0 0"></div>'
          +'<div style="width:16px;height:'+pH+'px;background:'+col+';border-radius:2px 2px 0 0"></div>'
        +'</div>'
        +'<div style="font-size:9px;font-weight:700;color:#7030A0;margin-top:2px">'+pV.toFixed(0)+'</div>'
        +'<div style="font-size:9px;color:'+dc+';font-weight:700">'+(delta<0?'':'+')+delta+'%</div>'
        +'<div style="font-size:10px;color:#6b7280;margin-top:1px">Z'+z.zone+'</div>'
        +'<div style="font-size:9px;color:#9ca3af">D'+z.group_a_day+'+'+z.group_b_day+'</div>'
        +'</div>';
    }).join('')+'</div>'
    +'<div style="font-size:10px;color:#9ca3af;margin-top:6px">&#9632; Ex (red) &nbsp;&#9632; Prop (zone color) &nbsp;% = change &nbsp;□― = zone boundary (dashed)</div>';
}

// SLIDE 13 - DELIVERY BEATS
let _db13m, _db13lg;
let _db13View='existing', _db13Limit='Max 2 sellers';
let _db13DayF=null, _db13ZoneF=null;
const _DB13_DAY=['','Mon','Tue','Wed','Thu','Fri','Sat'];
const _DB13_TRUCK_C={'3 Wheeler':'#1565C0','Tata Ace':'#388e3c','407':'#c62828'};
const _DB13_ZONE_C=['','#e91e63','#1565C0','#2e7d32','#e65100','#6a1b9a','#006064'];
function _db13fmt(v){return v>=100?'&#8377;'+(v/100).toFixed(1)+'L':'&#8377;'+v.toFixed(0)+'K';}

function initSlide13(){
  if(_db13m)return;
  const mapEl=document.getElementById('l13-map');
  _db13m=L.map(mapEl,{zoomControl:true,preferCanvas:true}).setView([22.52,88.36],12);
  L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png',
    {attribution:'&copy; OpenStreetMap &copy; CARTO',subdomains:'abcd',maxZoom:19,opacity:0.9}).addTo(_db13m);
  _db13lg=L.layerGroup().addTo(_db13m);
  setTimeout(()=>_db13m.invalidateSize(),200);
  _buildDB13Chips();
  renderDB13();
}

function setDB13View(v){
  _db13View=v;_db13DayF=null;_db13ZoneF=null;
  document.getElementById('db13-vex').classList.toggle('active',v==='existing');
  document.getElementById('db13-vv4').classList.toggle('active',v==='proposed');
  document.getElementById('db13-day-section').style.display=v==='existing'?'':'none';
  document.getElementById('db13-zone-section').style.display=v==='proposed'?'':'none';
  document.getElementById('db13-col-lbl').textContent=v==='existing'?'Day':'Zone';
  _buildDB13Chips();
  renderDB13();
}

function setDB13Limit(lim){
  _db13Limit=lim;
  ['Max 2 sellers','Max 3 sellers','Max 4 sellers'].forEach((l,i)=>{
    document.getElementById(['db13-m2','db13-m3','db13-m4'][i]).classList.toggle('active',l===lim);
  });
  renderDB13();
}

function _buildDB13Chips(){
  if(_db13View==='existing'){
    const el=document.getElementById('db13-day-chips');
    el.innerHTML='<button class="beat-chip'+(null===_db13DayF?' active':'')+'" onclick="_db13DayClick(null)">All</button>';
    for(let d=1;d<=6;d++)el.innerHTML+='<button class="beat-chip'+(d===_db13DayF?' active':'')+'" onclick="_db13DayClick('+d+')">'+_DB13_DAY[d]+'</button>';
  } else {
    const el=document.getElementById('db13-zone-chips');
    const zones=(DELIVERY_ZONES&&DELIVERY_ZONES.zones)||[];
    el.innerHTML='<button class="beat-chip'+(null===_db13ZoneF?' active':'')+'" onclick="_db13ZoneClick(null)">All</button>';
    zones.forEach(z=>{
      const col=_DB13_ZONE_C[z.zone]||'#666';
      el.innerHTML+='<button class="beat-chip'+(z.zone===_db13ZoneF?' active':'')+'" onclick="_db13ZoneClick('+z.zone+')" style="border-color:'+col+';'+(z.zone===_db13ZoneF?'background:'+col+';color:#fff;':'')+'">Z'+z.zone+' ('+_DB13_DAY[z.group_a_day]+'+'+_DB13_DAY[z.group_b_day]+')</button>';
    });
  }
}

function _db13DayClick(d){_db13DayF=d;_buildDB13Chips();renderDB13();}
function _db13ZoneClick(z){_db13ZoneF=z;_buildDB13Chips();renderDB13();}

function _db13GetBeats(){
  if(!DELIVERY_DATA)return[];
  const sc=_db13View==='existing'?'Existing':'Output 1';
  const src=(DELIVERY_DATA[sc]||{})[_db13Limit]||{};
  let beats=[];
  if(_db13View==='existing'){
    const days=_db13DayF?[String(_db13DayF)]:Object.keys(src);
    days.forEach(d=>{(src[d]||[]).forEach(b=>beats.push({...b,_day:+d}));});
  } else {
    const zones=(DELIVERY_ZONES&&DELIVERY_ZONES.zones)||[];
    const selZ=_db13ZoneF?zones.filter(z=>z.zone===_db13ZoneF):zones;
    const daySet=new Set();
    selZ.forEach(z=>{daySet.add(z.group_a_day);daySet.add(z.group_b_day);});
    daySet.forEach(d=>{(src[String(d)]||[]).forEach(b=>beats.push({...b,_day:d}));});
  }
  return beats;
}

function renderDB13Map(){
  if(!_db13lg)return;
  _db13lg.clearLayers();
  const zones=(DELIVERY_ZONES&&DELIVERY_ZONES.zones)||[];
  if(_db13View==='proposed'){
    const selZ=_db13ZoneF?zones.filter(z=>z.zone===_db13ZoneF):zones;
    selZ.forEach(z=>{
      if(!z.v4_hull||z.v4_hull.length<3)return;
      const col=_DB13_ZONE_C[z.zone]||'#666';
      L.polygon(z.v4_hull,{color:col,fillColor:col,fillOpacity:0.08,weight:2}).addTo(_db13lg);
    });
  }
  const beats=_db13GetBeats().filter(b=>(b.sub_id===null||b.sub_id==='a'));
  beats.forEach(b=>{
    if(!b.hull||b.hull.length<3)return;
    let col;
    if(_db13View==='proposed'){
      const z=zones.find(z=>z.group_a_day===b._day||z.group_b_day===b._day);
      col=z?(_DB13_ZONE_C[z.zone]||'#666'):'#666';
    } else {
      col=b.truck_color||_DB13_TRUCK_C[b.truck]||'#666';
    }
    L.polygon(b.hull,{color:col,fillColor:col,fillOpacity:0.2,weight:1.5}).addTo(_db13lg);
  });
}

function renderDB13(){
  if(!_db13m)return;
  renderDB13Map();
  renderDB13Panel();
}

function renderDB13Panel(){
  const beats=_db13GetBeats();
  const firstOnly=beats.filter(b=>(b.sub_id===null||b.sub_id==='a'));
  const zones=(DELIVERY_ZONES&&DELIVERY_ZONES.zones)||[];

  const totBeats=firstOnly.length;
  const totOutlets=firstOnly.reduce((s,b)=>s+(b.outlets||0),0);
  const totCost=beats.reduce((s,b)=>s+(b.cost||0),0);
  const truckCt={'3 Wheeler':0,'Tata Ace':0,'407':0};
  beats.forEach(b=>{if(b.truck in truckCt)truckCt[b.truck]++;});
  document.getElementById('db13-kpis').innerHTML=
    '<div class="kpi"><div class="kv">'+totBeats+'</div><div class="kl">Beats</div></div>'+
    '<div class="kpi"><div class="kv">'+totOutlets+'</div><div class="kl">Outlets</div></div>'+
    '<div class="kpi"><div class="kv">'+_db13fmt(totCost)+'</div><div class="kl">Cost/wk</div></div>';

  document.getElementById('db13-truck-legend').innerHTML=
    Object.entries(_DB13_TRUCK_C).map(([t,c])=>'<span style="display:flex;align-items:center;gap:4px"><span style="width:12px;height:12px;border-radius:2px;background:'+c+'"></span>'+t+' ('+truckCt[t]+')</span>').join('');

  const tbody=document.getElementById('db13-tbody');
  if(_db13View==='existing'){
    const days=_db13DayF?[_db13DayF]:[1,2,3,4,5,6];
    const sc='Existing';
    const src=(DELIVERY_DATA[sc]||{})[_db13Limit]||{};
    tbody.innerHTML=days.map(d=>{
      const rows=(src[String(d)]||[]);
      const firstRows=rows.filter(b=>(b.sub_id===null||b.sub_id==='a'));
      const nB=firstRows.length,nO=firstRows.reduce((s,b)=>s+(b.outlets||0),0);
      const tc={'3 Wheeler':0,'Tata Ace':0,'407':0};
      rows.forEach(b=>{if(b.truck in tc)tc[b.truck]++;});
      const cost=rows.reduce((s,b)=>s+(b.cost||0),0);
      return '<tr><td style="text-align:left;font-weight:600">'+_DB13_DAY[d]+'</td><td>'+nB+'</td><td>'+nO+'</td><td>'+tc['3 Wheeler']+'</td><td>'+tc['Tata Ace']+'</td><td>'+tc['407']+'</td><td>'+_db13fmt(cost)+'</td></tr>';
    }).join('');
  } else {
    const selZ=_db13ZoneF?zones.filter(z=>z.zone===_db13ZoneF):zones;
    const src=(DELIVERY_DATA['Output 1']||{})[_db13Limit]||{};
    tbody.innerHTML=selZ.map(z=>{
      const days=[z.group_a_day,z.group_b_day];
      const rows=days.flatMap(d=>(src[String(d)]||[]));
      const firstRows=rows.filter(b=>(b.sub_id===null||b.sub_id==='a'));
      const nB=firstRows.length,nO=firstRows.reduce((s,b)=>s+(b.outlets||0),0);
      const tc={'3 Wheeler':0,'Tata Ace':0,'407':0};
      rows.forEach(b=>{if(b.truck in tc)tc[b.truck]++;});
      const cost=rows.reduce((s,b)=>s+(b.cost||0),0);
      const col=_DB13_ZONE_C[z.zone]||'#666';
      return '<tr><td style="text-align:left;font-weight:600;color:'+col+'">Z'+z.zone+' ('+_DB13_DAY[z.group_a_day]+'+'+_DB13_DAY[z.group_b_day]+')</td><td>'+nB+'</td><td>'+nO+'</td><td>'+tc['3 Wheeler']+'</td><td>'+tc['Tata Ace']+'</td><td>'+tc['407']+'</td><td>'+_db13fmt(cost)+'</td></tr>';
    }).join('');
  }
}

// ── SLIDE EXBEAT · Existing Beats Explorer ──────────────────────────────────
let _exbm, _exblg;
let _exbDays = new Set();      // empty = all
let _exbPLGs = new Set();      // empty = all
let _exbDSEs = new Set();      // empty = all (DSE idx)
let _exbBeats = new Set();     // empty = all (beat key plg|dse|market)
let _exbCB = 'plg';
let _exbOutletSearch = '';     // outlet code/name search
function exbSetOutletSearch(v){ _exbOutletSearch = v || ''; renderEXB(); }
const _EXB_DAY = ['Mon','Tue','Wed','Thu','Fri','Sat'];
const _EXB_DAY_COL = ['#1565C0','#388e3c','#e65100','#6a1b9a','#c62828','#00838f'];

function initSlideEXB(){
  if(_exbm) return;
  _exbm = L.map('map-exbeat',{zoomControl:true,preferCanvas:true}).setView([22.52,88.36],12);
  L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png',
    {attribution:'&copy; OpenStreetMap &copy; CARTO',subdomains:'abcd',maxZoom:19,opacity:0.9}).addTo(_exbm);
  _exblg = L.layerGroup().addTo(_exbm);
  setTimeout(()=>_exbm.invalidateSize(),200);
  _exbBuildDayChips();
  _exbBuildPLGChips();
  exbRenderDseList();
  exbRenderBeatList();
  renderEXB();
}
function _exbBuildDayChips(){
  const el = document.getElementById('exb-day-chips');
  const items = [{v:null,l:'All',c:'#1565C0'},
    ...[0,1,2,3,4,5].map(d=>({v:d,l:_EXB_DAY[d],c:_EXB_DAY_COL[d]}))];
  el.innerHTML = items.map(it=>{
    const isA = it.v===null ? _exbDays.size===0 : _exbDays.has(it.v);
    return '<button class="beat-chip'+(isA?' active':'')+'" '
      +'style="'+(isA?'background:'+it.c+';color:white;border-color:'+it.c+';':'border-color:'+it.c+'40;')+'" '
      +'onclick="exbToggleDay('+JSON.stringify(it.v)+')">'+it.l+'</button>';
  }).join('');
}
function _exbBuildPLGChips(){
  const el = document.getElementById('exb-plg-chips');
  const allOn = _exbPLGs.size===0;
  let h = '<button class="beat-chip'+(allOn?' active':'')+'" '
    +'style="'+(allOn?'background:#1565C0;color:white;border-color:#1565C0;':'')+'" '
    +'onclick="exbClearPLGs()">All</button>';
  EX_PLG_J26.forEach(p=>{
    const isA = _exbPLGs.has(p.idx);
    h += '<button class="beat-chip'+(isA?' active':'')+'" '
      +'style="'+(isA?'background:'+p.color+';color:white;border-color:'+p.color+';':'border-color:'+p.color+'40;')+'" '
      +'onclick="exbTogglePLG('+p.idx+')">'+p.name+'</button>';
  });
  el.innerHTML = h;
}
function exbToggleDay(v){
  if(v===null){ _exbDays.clear(); } else {
    v = parseInt(v);
    if(_exbDays.has(v)) _exbDays.delete(v); else _exbDays.add(v);
  }
  _exbBuildDayChips(); exbRenderDseList(); exbRenderBeatList(); renderEXB();
}
function exbTogglePLG(i){
  i = parseInt(i);
  if(_exbPLGs.has(i)) _exbPLGs.delete(i); else _exbPLGs.add(i);
  _exbBuildPLGChips(); exbRenderDseList(); exbRenderBeatList(); renderEXB();
}
function exbClearPLGs(){
  _exbPLGs.clear(); _exbDSEs.clear(); _exbBeats.clear();
  _exbBuildPLGChips(); exbRenderDseList(); exbRenderBeatList(); renderEXB();
}
function exbClearFilters(){
  _exbDays.clear(); _exbPLGs.clear(); _exbDSEs.clear(); _exbBeats.clear();
  _exbBuildDayChips(); _exbBuildPLGChips();
  document.getElementById('exb-dse-search').value='';
  document.getElementById('exb-beat-search').value='';
  exbRenderDseList(); exbRenderBeatList(); renderEXB();
}
function exbSetCB(mode){
  _exbCB = mode;
  document.getElementById('exb-cb-plg').classList.toggle('active', mode==='plg');
  document.getElementById('exb-cb-day').classList.toggle('active', mode==='day');
  document.getElementById('exb-cb-beat').classList.toggle('active', mode==='beat');
  renderEXB();
}
function exbToggleDse(i){
  i = parseInt(i);
  if(_exbDSEs.has(i)) _exbDSEs.delete(i); else _exbDSEs.add(i);
  exbRenderDseList(); exbRenderBeatList(); renderEXB();
}
function exbRenderDseList(){
  const el = document.getElementById('exb-dse-list');
  if(!el) return;
  const search = (document.getElementById('exb-dse-search').value || '').toLowerCase();
  // Only show DSEs that match current PLG/Day filters
  const allDays = _exbDays.size===0;
  const allPLGs = _exbPLGs.size===0;
  const validDseIdx = new Set();
  EX_BEATS_J26.forEach(b=>{
    const [lat,lon,pi,m,di] = b;
    if(!allDays && !_exbDays.has(m)) return;
    if(!allPLGs && !_exbPLGs.has(pi)) return;
    validDseIdx.add(di);
  });
  // Build display: PLG · RSSP (unique RSSP after stripping RS prefix)
  const items = [];
  EX_DSE_J26.forEach(d=>{
    if(!validDseIdx.has(d.idx)) return;
    const sep = d.name.indexOf(':');
    const plgName = d.name.substring(0,sep);
    const rssp = d.name.substring(sep+1);
    const label = plgName + ' · ' + rssp;
    if(search && !label.toLowerCase().includes(search)) return;
    items.push({idx:d.idx, label:label});
  });
  items.sort((a,b)=>a.label.localeCompare(b.label));
  if(items.length===0){ el.innerHTML='<div style="color:#9ca3af;padding:6px">No salesmen match filters</div>'; return; }
  el.innerHTML = items.map(it=>{
    const chk = _exbDSEs.has(it.idx);
    return '<label style="display:flex;align-items:center;gap:6px;padding:2px 4px;cursor:pointer">'
      +'<input type="checkbox" '+(chk?'checked':'')+' onchange="exbToggleDse('+it.idx+')">'
      +'<span style="font-size:11px">'+it.label+'</span></label>';
  }).join('');
}
function exbToggleBeat(key){
  if(_exbBeats.has(key)) _exbBeats.delete(key); else _exbBeats.add(key);
  exbRenderBeatList(); renderEXB();
}
function exbRenderBeatList(){
  const el = document.getElementById('exb-beat-list');
  if(!el) return;
  const search = (document.getElementById('exb-beat-search').value || '').toLowerCase();
  const allDays = _exbDays.size===0;
  const allPLGs = _exbPLGs.size===0;
  const allDSEs = _exbDSEs.size===0;
  // Group beat_meta into unique geo names (deduplicate across day/PLG/DSE)
  const seen = new Map();  // key = plg|dse|market → meta
  EX_BEAT_META.forEach(m=>{
    if(!allDays && !_exbDays.has(m.market)) return;
    if(!allPLGs && !_exbPLGs.has(m.plg)) return;
    if(!allDSEs && !_exbDSEs.has(m.dse)) return;
    const key = m.plg+'|'+m.dse+'|'+m.market;
    if(!seen.has(key)) seen.set(key, m);
  });
  const items = Array.from(seen.values());
  // Optionally filter by search
  const filtered = search ? items.filter(m=>(m.name+' '+m.geo).toLowerCase().includes(search)) : items;
  // Sort: by geo name, then PLG, then day
  filtered.sort((a,b)=>{
    const c = (a.geo||'').localeCompare(b.geo||'');
    if(c) return c;
    const pa = EX_PLG_J26[a.plg]?.name || '';
    const pb = EX_PLG_J26[b.plg]?.name || '';
    const pc = pa.localeCompare(pb);
    if(pc) return pc;
    return a.market - b.market;
  });
  if(filtered.length===0){ el.innerHTML='<div style="color:#9ca3af;padding:6px">No beats match filters</div>'; return; }
  el.innerHTML = filtered.slice(0,500).map(m=>{
    const key = m.plg+'|'+m.dse+'|'+m.market;
    const chk = _exbBeats.has(key);
    const plgName = EX_PLG_J26[m.plg]?.name || '?';
    const dayName = _EXB_DAY[m.market] || '?';
    return '<label style="display:flex;align-items:center;gap:6px;padding:2px 4px;cursor:pointer">'
      +'<input type="checkbox" '+(chk?'checked':'')+' onchange="exbToggleBeat(\\''+key+'\\')">'
      +'<span style="font-size:11px"><b>'+m.name+'</b> · '+plgName+' · '+dayName+' · '+m.outlets+' outlets</span></label>';
  }).join('') + (filtered.length>500 ? '<div style="color:#9ca3af;padding:4px;font-size:10px">Showing 500 of '+filtered.length+' — narrow filters to see more</div>' : '');
}
function renderEXB(){
  if(!_exblg) return;
  _exblg.clearLayers();
  const allDays = _exbDays.size===0;
  const allPLGs = _exbPLGs.size===0;
  const allDSEs = _exbDSEs.size===0;
  const allBeats = _exbBeats.size===0;
  // Hash beat name colors so we can color by beat
  function strColor(s){ let h=0; for(let i=0;i<s.length;i++) h=((h<<5)-h)+s.charCodeAt(i); const hue=Math.abs(h)%360; return 'hsl('+hue+',60%,45%)'; }
  // Build a quick (plg, dse, market) → beat-meta key map for color/filter
  const metaMap = new Map();
  EX_BEAT_META.forEach(m=>{
    metaMap.set(m.plg+'|'+m.dse+'|'+m.market, m);
  });
  let nOutlets = 0, exbMatchCount = 0;
  const exbMatchBounds = [];
  const visibleBeats = new Set();
  EX_BEATS_J26.forEach(b=>{
    const [lat,lon,pi,mk,di,bi,oi] = b;
    if(!allDays && !_exbDays.has(mk)) return;
    if(!allPLGs && !_exbPLGs.has(pi)) return;
    if(!allDSEs && !_exbDSEs.has(di)) return;
    const beatKey = pi+'|'+di+'|'+mk;
    if(!allBeats && !_exbBeats.has(beatKey)) return;
    visibleBeats.add(beatKey);
    // Outlet metadata for tooltip + search (matches against code, name, OR beat name)
    const om = (oi !== undefined) ? (EX_OUTLET_META[oi] || null) : null;
    const beatMeta = metaMap.get(beatKey);
    const sq = _exbOutletSearch.toLowerCase().trim();
    let isMatch = false;
    if(sq){
      const hay = ((om?.code||'')+' '+(om?.name||'')+' '+(beatMeta?.name||'')+' '+(beatMeta?.geo||'')).toLowerCase();
      if(!hay.includes(sq)) return;
      isMatch = true;
    }
    let col;
    if(_exbCB==='day') col = _EXB_DAY_COL[mk];
    else if(_exbCB==='beat'){ const mm = metaMap.get(beatKey); col = mm ? strColor(mm.geo||'') : '#999'; }
    else col = EX_PLG_J26[pi]?.color || '#666';
    const dseObj = EX_DSE_J26[di] || {};
    const dseShort = (dseObj.name && dseObj.name.indexOf(':')>=0) ? dseObj.name.split(':',2)[1] : (dseObj.name||'?');
    let tip = '<b>'+(om?.name || om?.code || 'Outlet')+'</b>';
    if(om?.code) tip += '<br>'+om.code;
    tip += '<br><span style="color:#94a3b8">Beat:</span> '+(beatMeta?.name||'—');
    tip += '<br><span style="color:#94a3b8">PLG·DSE·Day:</span> '+(EX_PLG_J26[pi]?.name||'?')+' · '+dseShort+' · '+_EXB_DAY[mk];
    if(om?.ch || om?.cls) tip += '<br><span style="color:#94a3b8">Channel:</span> '+(om.ch||'')+' · '+(om.cls||'');
    if(om?.prog && om.prog !== '0') tip += '<br><span style="color:#94a3b8">Program:</span> '+om.prog;
    if(isMatch){
      L.circleMarker([lat,lon],{radius:7,color:'#1d4ed8',fillColor:col,fillOpacity:1,weight:2.5})
        .bindTooltip(tip,{sticky:true,direction:'top'}).addTo(_exblg);
      exbMatchBounds.push([lat,lon]);
      exbMatchCount++;
    } else {
      L.circleMarker([lat,lon],{radius:3,color:col,fillColor:col,fillOpacity:sq?0.25:0.75,weight:0})
        .bindTooltip(tip,{sticky:true,direction:'top'}).addTo(_exblg);
    }
    nOutlets++;
  });
  const nBeats = visibleBeats.size;
  const nPlgs = (allPLGs?EX_PLG_J26.length:_exbPLGs.size);
  document.getElementById('exb-kpis').innerHTML =
    '<div class="kpi"><div class="kpi-v">'+nOutlets.toLocaleString()+'</div><div class="kpi-l">visits shown</div></div>'
    +'<div class="kpi"><div class="kpi-v">'+nBeats+'</div><div class="kpi-l">beats</div></div>'
    +'<div class="kpi"><div class="kpi-v">'+nPlgs+'</div><div class="kpi-l">PLGs</div></div>';
  // Update search overlay UI
  const cntEl = document.getElementById('exb-search-count');
  const clrEl = document.getElementById('exb-search-clear');
  const sq = _exbOutletSearch.toLowerCase().trim();
  if(sq){
    if(cntEl) cntEl.textContent = exbMatchCount+' match'+(exbMatchCount===1?'':'es');
    if(clrEl) clrEl.style.display = 'inline';
    if(exbMatchBounds.length>0 && exbMatchCount<=200){
      try { _exbm.fitBounds(exbMatchBounds, {padding:[40,40], maxZoom:15}); } catch(e){}
    }
  } else {
    if(cntEl) cntEl.textContent = '';
    if(clrEl) clrEl.style.display = 'none';
  }
}
function exbClearSearch(){
  _exbOutletSearch='';
  const el=document.getElementById('exb-search'); if(el) el.value='';
  renderEXB();
}

// ── SLIDE JUN26 · Aligned Beats Overview (mirror of slide-5 layout) ─────────
let _j26m, _j26lg;
let _j26DayF=null;
let _j26View='proposed';        // 'proposed' | 'existing'
let _j26PLG=new Set();          // selected PLG indices (empty + !None = all)
let _j26PLGNone=false;          // true = no PLGs selected (deselect-all)
let _j26DSE=new Set();          // selected DSE indices (empty + !None = all)
let _j26Expanded=new Set();     // expanded PLG indices in tree
let _j26CB='plg';
let _j26Search='';              // outlet code/name search filter

function _j26OutletMeta(){
  return _j26View==='existing' ? (EX_OUTLET_META||[]) : (OUTLET_META||[]);
}
// Build beat lookup: (plg_idx, dse_idx, market_0idx) → beat name
// Proposed: from TRUCKS_JUN26.trucks[].beat (each truck spans visits)
// Existing: from EX_BEAT_META rows
let _j26PropBeatLookup = null;
function _j26BeatNameFor(D, pi, di, mk){
  if(_j26View==='existing'){
    const row = (EX_BEAT_META||[]).find(b=>b.plg===pi && b.dse===di && b.market===mk);
    return row ? (row.name || row.geo || '') : '';
  }
  if(!_j26PropBeatLookup){
    _j26PropBeatLookup = {};
    const trucks = (TRUCKS_JUN26 && TRUCKS_JUN26.trucks) || [];
    const dseList = D.DSE || [];
    trucks.forEach(t=>{
      (t.visits||[]).forEach(v=>{
        // v.plg = 'd+f+n' (string), v.dse = 'S001' (short), v.day = 1-6
        // Need to map to (plg_idx, dse_idx, market_0idx)
        const piMatch = D.PLG.findIndex(p=>p.name.toLowerCase().replace('ofm-','ofm_') === v.plg);
        if(piMatch<0) return;
        const diMatch = dseList.findIndex(d=>{
          if(!d.name || d.name.indexOf(':')<0) return false;
          const [pn, dn] = d.name.split(':',2);
          return pn.toLowerCase() === v.plg && dn === v.dse;
        });
        if(diMatch<0) return;
        _j26PropBeatLookup[piMatch+'|'+diMatch+'|'+(v.day-1)] = t.beat || '';
      });
    });
  }
  return _j26PropBeatLookup[pi+'|'+di+'|'+mk] || '';
}
function _j26OutletTip(b, D){
  const [lat,lon,pi,m,di,bi,oi] = b;
  const meta = (_j26OutletMeta()[oi]) || {};
  const dseObj = D.DSE[di] || {};
  const dseShort = (dseObj.name && dseObj.name.indexOf(':')>=0) ? dseObj.name.split(':',2)[1] : (dseObj.name||'?');
  const beatName = _j26BeatNameFor(D, pi, di, m);
  let s = '<b>'+(meta.name || meta.code || 'Outlet')+'</b>';
  if(meta.code) s += '<br>'+meta.code;
  if(beatName) s += '<br><span style="color:#94a3b8">Beat:</span> '+beatName;
  s += '<br><span style="color:#94a3b8">PLG·DSE·Day:</span> '+D.PLG[pi].name+' · '+dseShort+' · '+_J26_DAY[m+1];
  if(meta.ch || meta.cls) s += '<br><span style="color:#94a3b8">Channel:</span> '+(meta.ch||'')+' · '+(meta.cls||'');
  if(meta.prog && meta.prog !== '0') s += '<br><span style="color:#94a3b8">Program:</span> '+meta.prog;
  return s;
}

// View-aware data accessor — switches between proposed and existing datasets
function _j26d(){
  if(_j26View==='existing'){
    return {BEATS:EX_BEATS_J26, PLG:EX_PLG_J26, DSE:EX_DSE_J26, HULL:EX_HULL_J26, DIST:EX_DIST_J26};
  }
  return {BEATS:BEATS_JUN26, PLG:PLG_JUN26, DSE:DSE_JUN26, HULL:HULL_JUN26, DIST:DIST_JUN26};
}
function j26SetView(v){
  if(v===_j26View) return;
  _j26View=v;
  _j26PropBeatLookup = null;   // reset beat-name cache
  _j26PLG=new Set(); _j26DSE=new Set(); _j26PLGNone=false; _j26Expanded=new Set();
  document.getElementById('j26-view-prop').classList.toggle('active', v==='proposed');
  document.getElementById('j26-view-exist').classList.toggle('active', v==='existing');
  _j26BuildTree();
  renderJ26();
}
const _J26_DAY=['','Mon','Tue','Wed','Thu','Fri','Sat'];
const _J26_DAY_COL=['#1565C0','#388e3c','#e65100','#6a1b9a','#c62828','#00838f'];

// Categorize PLGs into Normal / OFM / UNIGLOW for tree sections
function _j26PLGGroup(p){
  if(p.name.startsWith('OFM-')) return 'ofm';
  if(p.name==='D+PP-A'||p.name==='F+N+PP-B') return 'uni';
  return 'normal';
}

// Map: PLG idx → list of DSEs (view-aware, case-insensitive PLG match)
function _j26DSEsByPLG(){
  const D=_j26d();
  const m={};
  D.PLG.forEach(p=>m[p.idx]=[]);
  D.DSE.forEach(d=>{
    const sep=d.name.indexOf(':');
    if(sep<0) return;
    const plgName=d.name.substring(0,sep).toLowerCase();
    const dseShort=d.name.substring(sep+1);
    const plgIdx=D.PLG.findIndex(p=>p.name.toLowerCase().replace('ofm-','ofm_')===plgName);
    if(plgIdx>=0)m[plgIdx].push({idx:d.idx,short:dseShort,name:d.name});
  });
  return m;
}

function initSlideJun26(){
  if(_j26m)return;
  _j26m=L.map('map-jun26',{zoomControl:true,preferCanvas:true}).setView([22.52,88.36],12);
  L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png',
    {attribution:'&copy; OpenStreetMap &copy; CARTO',subdomains:'abcd',maxZoom:19,opacity:0.9}).addTo(_j26m);
  _j26lg=L.layerGroup().addTo(_j26m);
  setTimeout(()=>_j26m.invalidateSize(),200);
  _j26BuildChips();
  _j26BuildTree();
  renderJ26();
}

function _j26BuildChips(){
  const el=document.getElementById('j26-day-chips');
  const items=[{v:null,l:'All',c:null},
    {v:0,l:'Mon',c:_J26_DAY_COL[0]},{v:1,l:'Tue',c:_J26_DAY_COL[1]},
    {v:2,l:'Wed',c:_J26_DAY_COL[2]},{v:3,l:'Thu',c:_J26_DAY_COL[3]},
    {v:4,l:'Fri',c:_J26_DAY_COL[4]},{v:5,l:'Sat',c:_J26_DAY_COL[5]}];
  el.innerHTML=items.map(it=>{
    const isA=it.v===_j26DayF;
    const col=_j26CB==='day'&&it.c?it.c:'#1565C0';
    return '<button class="beat-chip'+(isA?' active':'')+'" data-d="'+it.v+'" '
      +'style="'+(isA?'background:'+col+';color:white;border-color:'+col+';':
                   _j26CB==='day'&&it.c?'border-color:'+it.c+'40;':'')+'" '
      +'onclick="j26SetDay(this.dataset.d)">'+it.l+'</button>';
  }).join('');
}
function j26SetDay(v){_j26DayF=(v==='null')?null:parseInt(v);_j26BuildChips();_j26BuildTree();renderJ26();}
function j26SetSearch(v){_j26Search=v||'';renderJ26();}
function j26ClearSearch(){
  _j26Search='';
  const el=document.getElementById('j26-search'); if(el) el.value='';
  renderJ26();
}

function _j26BuildTree(){
  const D=_j26d();
  const dsesByPLG=_j26DSEsByPLG();
  const groups={normal:[],ofm:[],uni:[]};
  D.PLG.forEach(p=>{groups[_j26PLGGroup(p)].push(p);});

  // Compute checkbox state for a PLG
  function plgCbState(p){
    if(_j26PLGNone) return '';
    const allPLG=_j26PLG.size===0;
    const isPLGSel=allPLG||_j26PLG.has(p.idx);
    if(!isPLGSel) return '';
    const dses=dsesByPLG[p.idx]||[];
    if(!dses.length) return 'on';
    const allDseSelected=_j26DSE.size===0;
    if(allDseSelected) return 'on';
    const onCount=dses.filter(d=>_j26DSE.has(d.idx)).length;
    if(onCount===0) return '';
    if(onCount===dses.length) return 'on';
    return 'partial';
  }

  function plgRow(p){
    const cbCls=plgCbState(p);
    const isOpen=_j26Expanded.has(p.idx);
    const isSel=cbCls==='on'||cbCls==='partial';
    const dses=dsesByPLG[p.idx]||[];
    const cnt=dses.length;
    const dot='<div class="plg-dot" style="background:'+p.color+'"></div>';
    const dseListHtml=isSel?dses.map(d=>{
      const dseAllOn=_j26DSE.size===0;
      const chk=dseAllOn||_j26DSE.has(d.idx);
      // Distance: when "All" days selected → average per day; specific day → that day's km
      let distKm = 0, distCnt = 0;
      D.DIST.forEach(ddd=>{
        if(ddd.plg !== p.idx) return;
        if(ddd.dse !== d.idx) return;
        if(_j26DayF !== null && ddd.market !== _j26DayF) return;
        distKm += ddd.distance_km; distCnt++;
      });
      let distLabel = '';
      if(distCnt > 0){
        if(_j26DayF !== null){
          distLabel = distKm.toFixed(1) + ' km';
        } else {
          // Average per day = total / number-of-days-with-beats
          distLabel = (distKm / distCnt).toFixed(1) + ' km/day avg';
        }
        distLabel = '<span style="margin-left:auto;padding:0 6px;font-size:10px;color:#6b7280;font-weight:600">'
          + distLabel + '</span>';
      }
      return '<div class="dse-item" data-i="'+d.idx+'" onclick="j26ToggleDSE(this.dataset.i)">'
        +'<div class="dse-cb'+(chk?' on':'')+'"></div>'
        +'<span class="dse-label">'+d.short+'</span>'+distLabel+'</div>';
    }).join(''):'';
    return '<div class="plg-item'+(isSel?' sel':'')+'">'
      +'<div class="plg-row">'
      +'<div class="plg-cb '+cbCls+'" data-i="'+p.idx+'" onclick="j26TogglePLG(this.dataset.i)"></div>'
      +(_j26CB==='plg'?dot:'')
      +'<span class="plg-name">'+p.name+'</span>'
      +'<span class="plg-cnt">'+cnt+' Salesm'+(cnt===1?'an':'en')+'</span>'
      +(isSel&&cnt>0?'<span class="plg-chev'+(isOpen?' open':'')+'" data-i="'+p.idx+'" onclick="j26ToggleExpand(this.dataset.i,event)">&#9660;</span>':'')
      +'</div>'
      +(isSel&&cnt>0&&isOpen?'<div class="dse-list open">'+dseListHtml+'</div>':'')
      +'</div>';
  }

  const allOn=!_j26PLGNone && _j26PLG.size===0;
  const allPartial=!_j26PLGNone && _j26PLG.size>0;
  const allCb=allOn?'on':(allPartial?'partial':'');
  let html='<div class="plg-tree">'
    +'<div class="plg-all-row'+((allOn||allPartial)?' sel':'')+'" onclick="j26ClearPLGs()">'
    +'<div class="plg-cb '+allCb+'"></div>'
    +'<span class="plg-name" style="color:'+((allOn||allPartial)?'#1565C0':'#374151')+'">All PLGs</span>'
    +'</div>';
  groups.normal.forEach(p=>html+=plgRow(p));
  if(groups.ofm.length){html+='<div class="plg-tree-sec ofm">OFM</div>';groups.ofm.forEach(p=>html+=plgRow(p));}
  if(groups.uni.length){html+='<div class="plg-tree-sec uni">UNIGLOW+UNICARE</div>';groups.uni.forEach(p=>html+=plgRow(p));}
  html+='</div>';
  document.getElementById('j26-plg-tree').innerHTML=html;
}

function j26TogglePLG(i){
  i=parseInt(i);
  const all=_j26d().PLG.map(p=>p.idx);
  if(_j26PLGNone){
    _j26PLGNone=false;
    _j26PLG=new Set([i]);
  } else if(_j26PLG.size===0){
    // currently "All on" — switch to subset = all except this
    _j26PLG=new Set(all.filter(x=>x!==i));
    if(_j26PLG.size===0)_j26PLGNone=true;
  } else if(_j26PLG.has(i)){
    _j26PLG.delete(i);
    if(_j26PLG.size===0)_j26PLGNone=true;
  } else {
    _j26PLG.add(i);
    if(_j26PLG.size===all.length)_j26PLG.clear();
  }
  _j26DSE.clear();
  _j26BuildTree(); renderJ26();
}
function j26ClearPLGs(){
  // 3-state cycle: All on → All off → All on
  if(!_j26PLGNone && _j26PLG.size===0){
    _j26PLGNone=true;
    _j26PLG.clear();
  } else {
    _j26PLGNone=false;
    _j26PLG.clear();
    _j26DSE.clear();
  }
  _j26BuildTree(); renderJ26();
}
function j26ToggleExpand(i, ev){
  if(ev)ev.stopPropagation();
  i=parseInt(i);
  if(_j26Expanded.has(i))_j26Expanded.delete(i); else _j26Expanded.add(i);
  _j26BuildTree();
}
function j26ToggleDSE(i){
  i=parseInt(i);
  if(_j26DSE.has(i))_j26DSE.delete(i); else _j26DSE.add(i);
  _j26BuildTree(); renderJ26();
}
function j26SetCB(mode){
  _j26CB=mode;
  document.getElementById('j26-cb-plg').classList.toggle('active',mode==='plg');
  document.getElementById('j26-cb-day').classList.toggle('active',mode==='day');
  _j26BuildChips(); _j26BuildTree(); renderJ26();
}

function renderJ26(){
  if(!_j26lg)return;
  _j26lg.clearLayers();
  const D=_j26d();
  if(_j26PLGNone){
    document.getElementById('j26-kpis').innerHTML=
      '<div class="kpi"><div class="kpi-v">0</div><div class="kpi-l">visits</div></div>'
      +'<div class="kpi"><div class="kpi-v">0</div><div class="kpi-l">PLGs</div></div>'
      +'<div class="kpi"><div class="kpi-v">0</div><div class="kpi-l">salesmen</div></div>';
    _renderJ26ConflictSummary();
    return;
  }
  const allPLG=_j26PLG.size===0;
  const allDSE=_j26DSE.size===0;
  const meta = _j26OutletMeta();
  const searchQ = (_j26Search||'').toLowerCase().trim();
  let n=0, nMatches=0;
  const matchBounds=[];
  D.BEATS.forEach(b=>{
    const [lat,lon,pi,m,di,bi,oi]=b;
    if(!allPLG && !_j26PLG.has(pi))return;
    if(!allDSE && !_j26DSE.has(di))return;
    if(_j26DayF!==null && _j26DayF!==m)return;
    let isMatch = false;
    if(searchQ){
      const om = (oi!==undefined) ? meta[oi] : null;
      const beatName = _j26BeatNameFor(D, pi, di, m);
      const hay = ((om?.code||'')+' '+(om?.name||'')+' '+beatName).toLowerCase();
      if(!hay.includes(searchQ)) return;
      isMatch = true;
    }
    const col=(_j26CB==='day')?_J26_DAY_COL[m]:D.PLG[pi].color;
    if(isMatch){
      // Highlighted: larger radius, dark border, full opacity
      L.circleMarker([lat,lon],{radius:7,color:'#1d4ed8',fillColor:col,fillOpacity:1,weight:2.5})
        .bindTooltip(_j26OutletTip(b,D),{sticky:true,direction:'top'}).addTo(_j26lg);
      nMatches++; matchBounds.push([lat,lon]);
    } else {
      L.circleMarker([lat,lon],{radius:3,color:col,fillColor:col,fillOpacity:searchQ?0.25:0.7,weight:0})
        .bindTooltip(_j26OutletTip(b,D),{sticky:true,direction:'top'}).addTo(_j26lg);
    }
    n++;
  });
  // Update search-bar UI (count + clear button + fit bounds)
  const cntEl = document.getElementById('j26-search-count');
  const clrEl = document.getElementById('j26-search-clear');
  if(searchQ){
    if(cntEl) cntEl.textContent = nMatches+' match'+(nMatches===1?'':'es');
    if(clrEl) clrEl.style.display = 'inline';
    if(matchBounds.length>0 && nMatches<=200){
      try { _j26m.fitBounds(matchBounds, {padding:[40,40], maxZoom:15}); } catch(e){}
    }
  } else {
    if(cntEl) cntEl.textContent = '';
    if(clrEl) clrEl.style.display = 'none';
  }
  const plgN=allPLG?D.PLG.length:_j26PLG.size;
  // Salesman count: proposed has per-PLG S001..S00N (one PLG = one person), so each entry is unique;
  // existing has shared RSSP codes across PLGs (one person serves DETS+FNB+NUTS), so dedupe by RSSP.
  function countSalesmen(dseSrc, selSet, view){
    if(view==='existing'){
      const uniq = new Set();
      dseSrc.forEach(d=>{
        if(selSet && selSet.size>0 && !selSet.has(d.idx)) return;
        const rssp = d.name.includes(':') ? d.name.split(':',2)[1] : d.name;
        uniq.add(rssp);
      });
      return uniq.size;
    }
    // proposed — count entries (each PLG-specific S00N is a distinct person)
    if(!selSet || selSet.size===0) return dseSrc.length;
    return selSet.size;
  }
  const dseN = countSalesmen(D.DSE, allDSE ? null : _j26DSE, _j26View);
  // Avg distance per day for current view (used for both KPI + comparison delta)
  function avgPerDayFor(distArr){
    let tot=0, cnt=0;
    distArr.forEach(dd=>{
      if(!allPLG && !_j26PLG.has(dd.plg))return;
      if(!allDSE && !_j26DSE.has(dd.dse))return;
      if(_j26DayF!==null && dd.market!==_j26DayF)return;
      tot += dd.distance_km; cnt++;
    });
    return cnt===0 ? null : (tot/cnt);
  }
  const curAvg = avgPerDayFor(D.DIST);
  const distStr = curAvg===null ? '–'
    : (_j26DayF!==null ? curAvg.toFixed(1) + ' km' : curAvg.toFixed(1) + ' km/day');
  const distLbl = _j26DayF!==null ? _J26_DAY[_j26DayF+1] + ' in-beat avg' : 'in-beat km/day avg';

  // Comparison delta — proposed vs existing (only when both data sets have values for the same filter)
  let cmpHTML = '';
  if(_j26View==='proposed'){
    // proposed view — show reduction vs existing if comparable
    // Caveat: filters apply by index, but PLG/DSE indices differ across views.
    // Show overall reduction only when no PLG/DSE filter is active.
    if(allPLG && allDSE){
      let exTot=0, exCnt=0;
      EX_DIST_J26.forEach(dd=>{
        if(_j26DayF!==null && dd.market!==_j26DayF)return;
        exTot += dd.distance_km; exCnt++;
      });
      if(exCnt>0 && curAvg!==null){
        const exAvg = exTot/exCnt;
        const delta = exAvg - curAvg;
        const pct = (delta/exAvg)*100;
        const sign = delta>=0 ? '↓' : '↑';
        const col = delta>=0 ? '#15803d' : '#b91c1c';
        cmpHTML = '<div class="kpi"><div class="kpi-v" style="color:'+col+'">'+sign+' '+Math.abs(pct).toFixed(0)+'%</div>'
          +'<div class="kpi-l">vs existing ('+exAvg.toFixed(1)+' km)</div></div>';
      }
    }
  } else {
    // existing view — show what proposed reduces to (when no filter)
    if(allPLG && allDSE){
      let prTot=0, prCnt=0;
      DIST_JUN26.forEach(dd=>{
        if(_j26DayF!==null && dd.market!==_j26DayF)return;
        prTot += dd.distance_km; prCnt++;
      });
      if(prCnt>0 && curAvg!==null){
        const prAvg = prTot/prCnt;
        const delta = curAvg - prAvg;
        const pct = (delta/curAvg)*100;
        const sign = delta>=0 ? '↓' : '↑';
        const col = delta>=0 ? '#15803d' : '#b91c1c';
        cmpHTML = '<div class="kpi"><div class="kpi-v" style="color:'+col+'">'+sign+' '+Math.abs(pct).toFixed(0)+'%</div>'
          +'<div class="kpi-l">proposed = '+prAvg.toFixed(1)+' km</div></div>';
      }
    }
  }

  document.getElementById('j26-kpis').innerHTML=
    '<div class="kpi"><div class="kpi-v">'+n.toLocaleString()+'</div><div class="kpi-l">visits</div></div>'
    +'<div class="kpi"><div class="kpi-v">'+plgN+'</div><div class="kpi-l">PLGs</div></div>'
    +'<div class="kpi"><div class="kpi-v">'+dseN+'</div><div class="kpi-l">salesmen</div></div>'
    +'<div class="kpi"><div class="kpi-v">'+distStr+'</div><div class="kpi-l">'+distLbl+'</div></div>'
    +cmpHTML;
  _renderJ26ConflictSummary();
}

function _renderJ26ConflictSummary(){
  const el=document.getElementById('j26-conflict-summary');
  if(!CONFLICTS_JUN26)return;
  const total=CONFLICTS_JUN26.length;
  el.innerHTML=
    '<table class="dt-tbl" style="font-size:11px;width:100%"><thead><tr>'
    +'<th style="text-align:left">Design</th><th>Same-day visits</th>'
    +'</tr></thead><tbody>'
    +'<tr><td style="text-align:left">Existing 218390</td><td>5,436</td></tr>'
    +'<tr><td style="text-align:left">V3 (P4 design)</td><td>1,722</td></tr>'
    +'<tr><td style="text-align:left">V3 + market opt</td><td>~180</td></tr>'
    +'<tr><td style="text-align:left">V4 (specialist)</td><td>675</td></tr>'
    +'<tr style="background:#dcfce7;font-weight:600"><td style="text-align:left">Jun 2026 (aligned)</td><td style="color:#059669">'+total+'</td></tr>'
    +'</tbody></table>';
}

// ── SLIDE JUN26-TERRITORIES (mirror of slide 9 + slide 13 tree filter) ────────
let _jtm, _jtlg;
let _jtDayF=null;
let _jtView='proposed';
let _jtPLG=new Set();
let _jtPLGNone=false;
let _jtDSE=new Set();
let _jtExpanded=new Set();

function _jtd(){
  if(_jtView==='existing') return {PLG:EX_PLG_J26, DSE:EX_DSE_J26, HULL:EX_HULL_J26};
  return {PLG:PLG_JUN26, DSE:DSE_JUN26, HULL:HULL_JUN26};
}
function jtSetView(v){
  if(v===_jtView) return;
  _jtView=v;
  _jtPLG=new Set(); _jtDSE=new Set(); _jtPLGNone=false; _jtExpanded=new Set();
  document.getElementById('jt-view-prop').classList.toggle('active', v==='proposed');
  document.getElementById('jt-view-exist').classList.toggle('active', v==='existing');
  _buildJTTree(); renderJT();
}

// Polygon area (km²) via projection to local km plane + shoelace
function _hullAreaKm(pts){
  if(!pts || pts.length<3) return 0;
  const clat = pts.reduce((s,p)=>s+p[0],0)/pts.length;
  const kx = 111.0 * Math.cos(clat*Math.PI/180);
  const xy = pts.map(p=>[p[1]*kx, p[0]*111.0]);
  let a=0;
  for(let i=0;i<xy.length;i++){
    const j=(i+1)%xy.length;
    a += xy[i][0]*xy[j][1] - xy[j][0]*xy[i][1];
  }
  return Math.abs(a)/2;
}
function initSlideJT(){
  if(_jtm)return;
  _jtm=L.map('map-jun26-terr',{zoomControl:true,preferCanvas:true}).setView([22.52,88.36],12);
  L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png',
    {attribution:'&copy; OpenStreetMap &copy; CARTO',subdomains:'abcd',maxZoom:19,opacity:0.9}).addTo(_jtm);
  _jtlg=L.layerGroup().addTo(_jtm);
  setTimeout(()=>_jtm.invalidateSize(),200);
  _buildJTChips(); _buildJTTree(); renderJT();
}
function _buildJTChips(){
  const el=document.getElementById('jt-day-chips');
  const items=[{v:null,l:'All',c:'#1565C0'},...[0,1,2,3,4,5].map(d=>({v:d,l:_J26_DAY[d+1],c:_J26_DAY_COL[d]}))];
  el.innerHTML=items.map(it=>{
    const isA=it.v===_jtDayF;
    return '<button class="beat-chip'+(isA?' active':'')+'" '
      +'style="'+(isA?'background:'+it.c+';color:white;border-color:'+it.c+';':'border-color:'+it.c+'40;')+'" '
      +'onclick="_jtDay('+JSON.stringify(it.v)+')">'+it.l+'</button>';
  }).join('');
}
function _jtDay(d){_jtDayF=d;_buildJTChips();renderJT();}
function _buildJTTree(){
  const D=_jtd();
  // Use view-specific DSE-by-PLG (same logic as _j26DSEsByPLG but on view data)
  const dsesByPLG={};
  D.PLG.forEach(p=>dsesByPLG[p.idx]=[]);
  D.DSE.forEach(d=>{
    const sep=d.name.indexOf(':');
    if(sep<0) return;
    const plgName=d.name.substring(0,sep).toLowerCase();
    const dseShort=d.name.substring(sep+1);
    const plgIdx=D.PLG.findIndex(p=>p.name.toLowerCase().replace('ofm-','ofm_')===plgName);
    if(plgIdx>=0)dsesByPLG[plgIdx].push({idx:d.idx,short:dseShort,name:d.name});
  });
  const groups={normal:[],ofm:[],uni:[]};
  D.PLG.forEach(p=>{groups[_j26PLGGroup(p)].push(p);});
  function plgCb(p){
    if(_jtPLGNone)return '';
    const allPLG=_jtPLG.size===0;
    const isSel=allPLG||_jtPLG.has(p.idx);
    if(!isSel)return '';
    const ds=dsesByPLG[p.idx]||[];
    if(!ds.length)return 'on';
    if(_jtDSE.size===0)return 'on';
    const on=ds.filter(d=>_jtDSE.has(d.idx)).length;
    if(on===0)return ''; if(on===ds.length)return 'on'; return 'partial';
  }
  function plgItem(p){
    const cb=plgCb(p), isSel=cb==='on'||cb==='partial';
    const isOpen=_jtExpanded.has(p.idx);
    const ds=dsesByPLG[p.idx]||[];
    const cnt=ds.length;
    const dseHtml=isSel?ds.map(d=>{
      const chk=_jtDSE.size===0||_jtDSE.has(d.idx);
      return '<div class="dse-item" data-i="'+d.idx+'" onclick="_jtToggleDSE(this.dataset.i)">'
        +'<div class="dse-cb'+(chk?' on':'')+'"></div>'
        +'<span class="dse-label">'+d.short+'</span></div>';
    }).join(''):'';
    return '<div class="plg-item'+(isSel?' sel':'')+'">'
      +'<div class="plg-row">'
      +'<div class="plg-cb '+cb+'" data-i="'+p.idx+'" onclick="_jtTogglePLG(this.dataset.i)"></div>'
      +'<div class="plg-dot" style="background:'+p.color+'"></div>'
      +'<span class="plg-name">'+p.name+'</span>'
      +'<span class="plg-cnt">'+cnt+'</span>'
      +(isSel&&cnt>0?'<span class="plg-chev'+(isOpen?' open':'')+'" data-i="'+p.idx+'" onclick="_jtToggleExpand(this.dataset.i,event)">&#9660;</span>':'')
      +'</div>'
      +(isSel&&cnt>0&&isOpen?'<div class="dse-list open">'+dseHtml+'</div>':'')
      +'</div>';
  }
  const allOn=!_jtPLGNone && _jtPLG.size===0;
  const allPartial=!_jtPLGNone && _jtPLG.size>0;
  const allCb=allOn?'on':(allPartial?'partial':'');
  let html='<div class="plg-tree">'
    +'<div class="plg-all-row'+((allOn||allPartial)?' sel':'')+'" onclick="_jtClearPLGs()">'
    +'<div class="plg-cb '+allCb+'"></div>'
    +'<span class="plg-name" style="color:'+((allOn||allPartial)?'#1565C0':'#374151')+'">All PLGs</span>'
    +'</div>';
  groups.normal.forEach(p=>html+=plgItem(p));
  if(groups.ofm.length){html+='<div class="plg-tree-sec ofm">OFM</div>';groups.ofm.forEach(p=>html+=plgItem(p));}
  if(groups.uni.length){html+='<div class="plg-tree-sec uni">UNIGLOW+UNICARE</div>';groups.uni.forEach(p=>html+=plgItem(p));}
  html+='</div>';
  document.getElementById('jt-plg-list').innerHTML=html;
}
function _jtTogglePLG(i){
  i=parseInt(i);
  const all=_jtd().PLG.map(p=>p.idx);
  if(_jtPLGNone){_jtPLGNone=false;_jtPLG=new Set([i]);}
  else if(_jtPLG.size===0){_jtPLG=new Set(all.filter(x=>x!==i));if(_jtPLG.size===0)_jtPLGNone=true;}
  else if(_jtPLG.has(i)){_jtPLG.delete(i);if(_jtPLG.size===0)_jtPLGNone=true;}
  else {_jtPLG.add(i);if(_jtPLG.size===all.length)_jtPLG.clear();}
  _jtDSE.clear();
  _buildJTTree(); renderJT();
}
function _jtClearPLGs(){
  if(!_jtPLGNone && _jtPLG.size===0){_jtPLGNone=true;_jtPLG.clear();}
  else {_jtPLGNone=false;_jtPLG.clear();_jtDSE.clear();}
  _buildJTTree(); renderJT();
}
function _jtToggleExpand(i,ev){
  if(ev)ev.stopPropagation();
  i=parseInt(i);
  if(_jtExpanded.has(i))_jtExpanded.delete(i); else _jtExpanded.add(i);
  _buildJTTree();
}
function _jtToggleDSE(i){
  i=parseInt(i);
  if(_jtDSE.has(i))_jtDSE.delete(i); else _jtDSE.add(i);
  _buildJTTree(); renderJT();
}
function renderJT(){
  if(!_jtlg)return;
  _jtlg.clearLayers();
  const D=_jtd();
  if(_jtPLGNone){document.getElementById('jt-kpis').innerHTML='<div class="kpi"><div class="kpi-v">0</div><div class="kpi-l">beats</div></div>';return;}
  const allPLG=_jtPLG.size===0, allDSE=_jtDSE.size===0;
  let n=0;
  D.HULL.forEach(h=>{
    if(!allPLG && !_jtPLG.has(h.plg))return;
    if(!allDSE && !_jtDSE.has(h.dse))return;
    if(_jtDayF!==null && h.market!==_jtDayF)return;
    if(!h.points || h.points.length<3)return;
    const col=D.PLG[h.plg].color;
    const dseObj = D.DSE[h.dse] || {};
    const dseShort = (dseObj.name && dseObj.name.indexOf(':')>=0) ? dseObj.name.split(':',2)[1] : (dseObj.name||'?');
    const tip = '<b>'+D.PLG[h.plg].name+'</b> · '+dseShort+' · '+_J26_DAY[h.market+1];
    L.polygon(h.points,{color:col,fillColor:col,fillOpacity:0.12,weight:1.5})
      .bindTooltip(tip,{sticky:true,direction:'top'}).addTo(_jtlg);
    n++;
  });
  // Headline KPIs from precomputed totals (avg km/day + avg pairwise overlap %)
  const T = (PLG_CMP_J26 && PLG_CMP_J26.totals) || {};
  const exKm = T.ex_km, prKm = T.pr_km;
  const exOv = T.ex_overlap_pct, prOv = T.pr_overlap_pct;
  let cmpHTML='';
  if(exKm!=null && prKm!=null){
    const delta = exKm - prKm;
    const pct = (delta/exKm)*100;
    const sign = delta>=0 ? '↓' : '↑';
    const col = delta>=0 ? '#15803d' : '#b91c1c';
    cmpHTML += '<div class="kpi"><div class="kpi-v" style="color:'+col+'">'+sign+' '+Math.abs(pct).toFixed(0)+'%</div>'
      +'<div class="kpi-l">km/day '+exKm.toFixed(1)+' → '+prKm.toFixed(1)+'</div></div>';
  }
  if(exOv!=null && prOv!=null){
    const delta = exOv - prOv;
    const pct = (delta/exOv)*100;
    const sign = delta>=0 ? '↓' : '↑';
    const col = delta>=0 ? '#15803d' : '#b91c1c';
    cmpHTML += '<div class="kpi"><div class="kpi-v" style="color:'+col+'">'+sign+' '+Math.abs(pct).toFixed(0)+'%</div>'
      +'<div class="kpi-l">avg overlap '+exOv.toFixed(2)+'% → '+prOv.toFixed(2)+'%</div></div>';
  }
  document.getElementById('jt-kpis').innerHTML=
    '<div class="kpi"><div class="kpi-v">'+n+'</div><div class="kpi-l">beats shown</div></div>'
    +'<div class="kpi"><div class="kpi-v">'+(allPLG?D.PLG.length:_jtPLG.size)+'</div><div class="kpi-l">PLGs</div></div>'
    +cmpHTML;
  _renderJTTables();
}

// Per-PLG comparison tables — Ex PLG → Prop PLG mapping (mirror slide 9 layout)
function _renderJTTables(){
  const rows = (PLG_CMP_J26 && PLG_CMP_J26.rows) || [];
  const T    = (PLG_CMP_J26 && PLG_CMP_J26.totals) || {};
  const fmt = v => v==null ? '&mdash;' : v.toFixed(2);
  const fmtKm = v => v==null ? '&mdash;' : v.toFixed(1)+' km';
  const fmtPct = v => v==null ? '&mdash;' : v.toFixed(2)+'%';

  // ── Distance table ────────────────────────────────────────────────────────
  let distHtml = rows.map(r=>{
    const d = (r.ex_km!=null && r.pr_km!=null) ? (r.pr_km - r.ex_km) : null;
    const col = d==null ? '#6b7280' : d<0 ? '#16a34a' : '#dc2626';
    return '<tr><td style="text-align:left">'+r.ex_plg+' &rarr; <b>'+r.pr_plg+'</b></td>'
      +'<td>'+fmtKm(r.ex_km)+'</td><td>'+fmtKm(r.pr_km)+'</td>'
      +'<td style="color:'+col+'">'+(d==null?'&mdash;':(d<0?'':'+')+d.toFixed(1))+'</td></tr>';
  }).join('');
  const dDelta = (T.ex_km!=null && T.pr_km!=null) ? (T.pr_km - T.ex_km) : null;
  const dCol = dDelta==null ? '#6b7280' : dDelta<0 ? '#16a34a' : '#dc2626';
  const dTotalRow = '<tr style="font-weight:700;background:#f9fafb"><td style="text-align:left">TOTAL (avg/beat)</td>'
    +'<td>'+fmtKm(T.ex_km)+'</td><td>'+fmtKm(T.pr_km)+'</td>'
    +'<td style="color:'+dCol+'">'+(dDelta==null?'&mdash;':(dDelta<0?'':'+')+dDelta.toFixed(1))+'</td></tr>';
  document.getElementById('jt-dist-table').innerHTML =
    '<div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:10px 0 2px">In-Beat Distance (km/market day)</div>'
    +'<div style="font-size:10px;color:#9ca3af;margin-bottom:4px">Route through outlets only (no depot legs) — what the salesman drives within their beat</div>'
    +'<table class="dt-tbl" style="width:100%"><thead><tr>'
    +'<th style="text-align:left">Ex PLG &rarr; Prop</th><th>Existing</th><th>Proposed</th><th>&Delta;</th>'
    +'</tr></thead><tbody>'+dTotalRow+distHtml+'</tbody></table>';

  // ── Hull overlap table ────────────────────────────────────────────────────
  let ovHtml = rows.map(r=>{
    const d = (r.ex_overlap_pct!=null && r.pr_overlap_pct!=null)
      ? (r.pr_overlap_pct - r.ex_overlap_pct) : null;
    const col = d==null ? '#6b7280' : d<0 ? '#16a34a' : '#dc2626';
    return '<tr><td style="text-align:left">'+r.ex_plg+' &rarr; <b>'+r.pr_plg+'</b></td>'
      +'<td>'+fmtPct(r.ex_overlap_pct)+'</td><td>'+fmtPct(r.pr_overlap_pct)+'</td>'
      +'<td style="color:'+col+'">'+(d==null?'&mdash;':(d<0?'':'+')+d.toFixed(2)+'%')+'</td></tr>';
  }).join('');
  const oDelta = (T.ex_overlap_pct!=null && T.pr_overlap_pct!=null) ? (T.pr_overlap_pct - T.ex_overlap_pct) : null;
  const oCol = oDelta==null ? '#6b7280' : oDelta<0 ? '#16a34a' : '#dc2626';
  const oTotalRow = '<tr style="font-weight:700;background:#f9fafb"><td style="text-align:left">TOTAL (avg pairwise)</td>'
    +'<td>'+fmtPct(T.ex_overlap_pct)+'</td><td>'+fmtPct(T.pr_overlap_pct)+'</td>'
    +'<td style="color:'+oCol+'">'+(oDelta==null?'&mdash;':(oDelta<0?'':'+')+oDelta.toFixed(2)+'%')+'</td></tr>';
  document.getElementById('jt-area-table').innerHTML =
    '<div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:10px 0 2px">Avg Pairwise Hull Overlap (%)</div>'
    +'<div style="font-size:10px;color:#9ca3af;margin-bottom:4px">Avg Jaccard overlap between same-PLG hulls. Lower = more disjoint territories</div>'
    +'<table class="dt-tbl" style="width:100%"><thead><tr>'
    +'<th style="text-align:left">Ex PLG &rarr; Prop</th><th>Existing</th><th>Proposed</th><th>&Delta;</th>'
    +'</tr></thead><tbody>'+oTotalRow+ovHtml+'</tbody></table>';
}

// ── SLIDE JUN26-DELIVERY ZONES (mirror of slide 11 Beat Area per Day) ────────
let _jzm, _jzlg;
let _jzZoneF=null;
const _JZ_COL=['','#e91e63','#1565C0','#2e7d32','#e65100','#6a1b9a','#006064'];
function initSlideJZ(){
  if(_jzm)return;
  _jzm=L.map('map-jun26-zones',{zoomControl:true,preferCanvas:true}).setView([22.52,88.36],12);
  L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png',
    {attribution:'&copy; OpenStreetMap &copy; CARTO',subdomains:'abcd',maxZoom:19,opacity:0.9}).addTo(_jzm);
  _jzlg=L.layerGroup().addTo(_jzm);
  setTimeout(()=>_jzm.invalidateSize(),200);
  _jzBuildChips(); renderJZ();
}
function _jzBuildChips(){
  const el=document.getElementById('jz-zone-chips');
  const zones=(ZONES_JUN26&&ZONES_JUN26.zones)||[];
  let html='<button class="beat-chip'+(null===_jzZoneF?' active':'')+'" '
    +'style="'+(null===_jzZoneF?'background:#1565C0;color:white;border-color:#1565C0;':'')+'" '
    +'onclick="_jzZone(null)">All</button>';
  zones.forEach(z=>{
    const col=_JZ_COL[z.zone]||'#666';
    const isA=z.zone===_jzZoneF;
    html+='<button class="beat-chip'+(isA?' active':'')+'" '
      +'style="'+(isA?'background:'+col+';color:white;border-color:'+col+';':'border-color:'+col+'40;')+'" '
      +'onclick="_jzZone('+z.zone+')">Z'+z.zone+' ('+_J26_DAY[z.group_a_day]+'+'+_J26_DAY[z.group_b_day]+')</button>';
  });
  el.innerHTML=html;
}
function _jzZone(z){_jzZoneF=z;_jzBuildChips();renderJZ();}
function renderJZ(){
  if(!_jzlg)return;
  _jzlg.clearLayers();
  const zones=(ZONES_JUN26&&ZONES_JUN26.zones)||[];
  const selZones=_jzZoneF?zones.filter(z=>z.zone===_jzZoneF):zones;
  selZones.forEach(z=>{
    const col=_JZ_COL[z.zone]||'#666';
    if(z.v4_hull&&z.v4_hull.length>=3)
      L.polygon(z.v4_hull,{color:col,fillColor:col,fillOpacity:0.10,weight:2}).addTo(_jzlg);
  });
  // KPIs
  const totalArea=selZones.reduce((s,z)=>s+(z.v4_area||0),0);
  const avgArea=selZones.length?totalArea/selZones.length:0;
  document.getElementById('jz-kpis').innerHTML=
    '<div class="kpi"><div class="kpi-v">'+selZones.length+'</div><div class="kpi-l">zones</div></div>'
    +'<div class="kpi"><div class="kpi-v">'+totalArea.toFixed(1)+'</div><div class="kpi-l">total km&sup2;</div></div>'
    +'<div class="kpi"><div class="kpi-v">'+avgArea.toFixed(1)+'</div><div class="kpi-l">avg / zone</div></div>'
    +'<div class="kpi"><div class="kpi-v">2</div><div class="kpi-l">days/zone</div></div>';
  // Bar chart of zone areas
  const maxA=Math.max(...zones.map(z=>z.v4_area||0),1);
  const chart=document.getElementById('jz-chart');
  chart.innerHTML=zones.map(z=>{
    const col=_JZ_COL[z.zone]||'#666';
    const w=(z.v4_area||0)/maxA*100;
    const isA=null===_jzZoneF||z.zone===_jzZoneF;
    return '<div style="display:flex;align-items:center;gap:6px;margin:3px 0;font-size:11px;opacity:'+(isA?1:0.35)+'">'
      +'<div style="width:56px;color:'+col+';font-weight:600">Z'+z.zone+' '+_J26_DAY[z.group_a_day]+'+'+_J26_DAY[z.group_b_day]+'</div>'
      +'<div style="flex:1;background:#f3f4f6;border-radius:4px;overflow:hidden;height:14px;position:relative">'
      +'<div style="background:'+col+';width:'+w+'%;height:100%"></div>'
      +'</div>'
      +'<div style="width:50px;text-align:right">'+(z.v4_area||0).toFixed(1)+' km&sup2;</div>'
      +'</div>';
  }).join('');
}

// ── SLIDE JUN26-DELIVERY BEATS ───────────────────────────────────────────────
let _jdm, _jdlg;
let _jdDayF=null;
let _jdTruckF=null;
let _jdView='proposed';
let _jdSearch='';
// Build outlet code → meta lookup once per view, lazily
let _jdOutletByCodeCache = {proposed:null, existing:null};
function _jdOutletByCode(){
  const view = _jdView;
  if(_jdOutletByCodeCache[view]) return _jdOutletByCodeCache[view];
  const src = view==='existing' ? (EX_OUTLET_META||[]) : (OUTLET_META||[]);
  const m = {};
  src.forEach(o=>{ if(o && o.code) m[o.code] = o; });
  _jdOutletByCodeCache[view] = m;
  return m;
}
function jdSetSearch(v){_jdSearch=v||'';renderJD();}
function jdClearSearch(){_jdSearch='';const el=document.getElementById('jd-search');if(el)el.value='';renderJD();}
// Does this truck match the current search? Match on id, beat, salesman, or any outlet code/name.
function _jdTruckMatches(t, q){
  if(!q) return true;
  q = q.toLowerCase();
  if((t.id||'').toLowerCase().includes(q)) return true;
  if((t.beat||'').toLowerCase().includes(q)) return true;
  // Visits' PLG · DSE
  for(const v of (t.visits||[])){
    if(((v.plg||'')+' '+(v.dse||'')).toLowerCase().includes(q)) return true;
  }
  // Outlet codes + names
  const lookup = _jdOutletByCode();
  for(const code of (t.outlet_codes||[])){
    if(code.toLowerCase().includes(q)) return true;
    const om = lookup[code];
    if(om && (om.name||'').toLowerCase().includes(q)) return true;
  }
  return false;
}
const _JD_COL={'3 Wheeler':'#1565C0','Tata Ace':'#388e3c','Split':'#c62828','Split (>1.5L)':'#c62828'};
let _jdCB='truck-type';  // 'truck-type' | 'truck' | 'beat'
function _jdHash(s){let h=0;for(let i=0;i<s.length;i++)h=((h<<5)-h)+s.charCodeAt(i);return Math.abs(h);}
function _jdColorFor(s){const hue=_jdHash(s||'')%360;return 'hsl('+hue+',65%,45%)';}
function jdSetCB(mode){
  _jdCB=mode;
  ['cb-tt','cb-truck','cb-beat'].forEach(id=>{
    const el=document.getElementById('jd-'+id);
    if(el) el.classList.toggle('active', id==='cb-tt'?(mode==='truck-type'):id==='cb-truck'?(mode==='truck'):(mode==='beat'));
  });
  renderJD();
}

// Return trucks in a normalised shape (existing data has different field names)
function _jdTrucks(){
  if(_jdView==='existing'){
    // Normalise EX_TRUCKS_J26 (a flat array) to proposed-shape objects
    return (EX_TRUCKS_J26||[]).map(t=>{
      const truckType = t.truck_type==='Split (>1.5L)' ? 'Split' : t.truck_type;
      // Build visits list, look up DSE short-name from EX_DSE_J26 idx
      const visits = (t.members||[]).map(m=>{
        const dseEntry = EX_DSE_J26.find(d=>d.idx===m.dse);
        const dseShort = dseEntry ? dseEntry.name.split(':')[1] : String(m.dse);
        return {plg: m.plg, dse: dseShort, day: m.day, value: m.value, outlets: m.outlets};
      });
      return {
        id: t.id,
        delivery_day: t.deliv_day,
        truck: truckType,
        total_value: t.value,
        outlets_n: t.outlets,
        salesmen_n: (t.salesmen||[]).length,
        centroid: t.centroid,
        positions: t.positions || (t.members||[]).map(m=>m.centroid),
        visits: visits,
        outlet_codes: t.outlet_codes || [],
        beat: t.beat || '',
        distance_km: t.distance_km || 0,
      };
    });
  }
  return (TRUCKS_JUN26 && TRUCKS_JUN26.trucks) || [];
}
function jdSetView(v){
  if(v===_jdView) return;
  _jdView=v;
  _jdSelected.clear();
  _jdOutletByCodeCache = {proposed:null, existing:null};
  document.getElementById('jd-view-prop').classList.toggle('active', v==='proposed');
  document.getElementById('jd-view-exist').classList.toggle('active', v==='existing');
  renderJD();
}
function initSlideJD(){
  if(_jdm)return;
  _jdm=L.map('map-jun26-del',{zoomControl:true,preferCanvas:true}).setView([22.52,88.36],12);
  L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png',
    {attribution:'&copy; OpenStreetMap &copy; CARTO',subdomains:'abcd',maxZoom:19,opacity:0.9}).addTo(_jdm);
  _jdlg=L.layerGroup().addTo(_jdm);
  setTimeout(()=>_jdm.invalidateSize(),200);
  _buildJDChips(); renderJD();
}
function _buildJDChips(){
  const el=document.getElementById('jd-day-chips');
  const items=[{v:null,l:'All',c:'#1565C0'},...[1,2,3,4,5,6].map(d=>({v:d,l:_J26_DAY[d],c:_J26_DAY_COL[d-1]}))];
  el.innerHTML=items.map(it=>{
    const isA=it.v===_jdDayF;
    return '<button class="beat-chip'+(isA?' active':'')+'" '
      +'style="'+(isA?'background:'+it.c+';color:white;border-color:'+it.c+';':'border-color:'+it.c+'40;')+'" '
      +'onclick="_jdDay('+JSON.stringify(it.v)+')">'+it.l+'</button>';
  }).join('');
}
let _jdSelected=new Set();
let _jdTruckTypes=new Set(['3 Wheeler','Tata Ace','Split']);   // visible truck types
function _jdDay(d){_jdDayF=d;_jdSelected.clear();_buildJDChips();renderJD();}
function _jdTruckFilter(t){_jdTruckF=t;_jdSelected.clear();_buildJDTruckChips();renderJD();}
function jdToggleTrip(id){
  if(_jdSelected.has(id))_jdSelected.delete(id); else _jdSelected.add(id);
  renderJD();
}
function jdClearSelected(){_jdSelected.clear();renderJD();}
function jdToggleTruckType(t){
  if(_jdTruckTypes.has(t))_jdTruckTypes.delete(t); else _jdTruckTypes.add(t);
  _buildJDTruckChips();renderJD();
}
function jdToggleAllTruckTypes(){
  const all = ['3 Wheeler','Tata Ace','Split'];
  if(_jdTruckTypes.size === all.length) _jdTruckTypes.clear();
  else _jdTruckTypes = new Set(all);
  _buildJDTruckChips();renderJD();
}
function _buildJDTruckChips(){
  const el=document.getElementById('jd-truck-chips');
  if(!el)return;
  const types=['3 Wheeler','Tata Ace','Split'];
  let h='<button class="beat-chip'+(_jdTruckTypes.size===types.length?' active':'')+'" '
    +'style="'+(_jdTruckTypes.size===types.length?'background:#1565C0;color:white;border-color:#1565C0;':'')+'" '
    +'onclick="jdToggleAllTruckTypes()">All</button>';
  types.forEach(t=>{
    const c=_JD_COL[t]||'#666';
    const isA=_jdTruckTypes.has(t);
    h+='<button class="beat-chip'+(isA?' active':'')+'" '
      +'style="'+(isA?'background:'+c+';color:white;border-color:'+c+';':'border-color:'+c+'40;')+'" '
      +'onclick="jdToggleTruckType('+JSON.stringify(t)+')">'+t+'</button>';
  });
  el.innerHTML=h;
}
function renderJD(){
  if(!_jdlg)return;
  _jdlg.clearLayers();
  const trucks=_jdTrucks();
  const sq = (_jdSearch||'').trim();
  const filt=trucks.filter(t=>{
    if(_jdDayF!==null && t.delivery_day!==_jdDayF) return false;
    if(!_jdTruckTypes.has(t.truck)) return false;
    if(sq && !_jdTruckMatches(t, sq)) return false;
    return true;
  });
  function truckColor(t){
    if(_jdCB==='truck')      return _jdColorFor(t.id);
    if(_jdCB==='beat')       return _jdColorFor((t.beat||'')+'|'+t.id);
    return _JD_COL[t.truck]||'#666';
  }
  function beatColor(t, v){
    // color-by-beat-within-truck: each (PLG, DSE, day) gets its own hue
    return _jdColorFor(t.id+'|'+v.plg+'|'+v.dse+'|'+v.day);
  }
  // Compose a tooltip string for an outlet visit
  function _outletTip(t, v){
    return '<b>'+t.id+'</b> · '+t.truck+' · '+_J26_DAY[t.delivery_day]+' deliv<br>'
      + '<span style="color:#94a3b8">Visit:</span> '+v.plg.toUpperCase()+' · '+v.dse+' · '+_J26_DAY[v.day]+'<br>'
      + '<span style="color:#94a3b8">Beat:</span> '+(t.beat||'—');
  }
  function _truckTip(t){
    const vs = t.visits.map(v=>v.plg.toUpperCase()+'/'+v.dse+'/'+_J26_DAY[v.day]).join(', ');
    return '<b>'+t.id+'</b> · '+t.truck+'<br>'
      + 'Deliv: '+_J26_DAY[t.delivery_day]+'<br>'
      + 'Beat: '+(t.beat||'—')+'<br>'
      + t.salesmen_n+' salesman'+(t.salesmen_n>1?'en':'')+' · '+t.outlets_n+' outlets · '+t.total_value.toFixed(2)+'L<br>'
      + vs;
  }
  // If any trip selected, only show those trips' outlet markers
  if(_jdSelected.size>0){
    filt.forEach(t=>{
      if(!_jdSelected.has(t.id))return;
      if(_jdCB==='beat'){
        let idx=0;
        (t.visits||[]).forEach(v=>{
          const n=v.outlets;
          const col=beatColor(t, v);
          const tip=_outletTip(t, v);
          for(let i=0;i<n && idx<t.positions.length;i++,idx++){
            const p=t.positions[idx];
            L.circleMarker(p,{radius:5,color:col,fillColor:col,fillOpacity:0.85,weight:1})
              .bindTooltip(tip,{sticky:true,direction:'top'}).addTo(_jdlg);
          }
        });
      } else {
        const col=truckColor(t);
        const tip=_truckTip(t);
        (t.positions||[]).forEach(p=>{
          L.circleMarker(p,{radius:5,color:col,fillColor:col,fillOpacity:0.85,weight:1})
            .bindTooltip(tip,{sticky:true,direction:'top'}).addTo(_jdlg);
        });
      }
    });
  } else {
    // Show truck centroids; highlight matches when searching
    filt.forEach(t=>{
      const col=truckColor(t);
      const r=Math.max(4,Math.sqrt(t.outlets_n)*1.2);
      if(sq){
        L.circleMarker(t.centroid,{radius:r+3,color:'#1d4ed8',fillColor:col,fillOpacity:1,weight:2.5})
          .bindTooltip(_truckTip(t),{sticky:true,direction:'top'}).addTo(_jdlg);
      } else {
        L.circleMarker(t.centroid,{radius:r,color:col,fillColor:col,fillOpacity:0.6,weight:1})
          .bindTooltip(_truckTip(t),{sticky:true,direction:'top'}).addTo(_jdlg);
      }
    });
  }
  // Update search overlay UI + fit bounds
  const cntEl = document.getElementById('jd-search-count');
  const clrEl = document.getElementById('jd-search-clear');
  if(sq){
    if(cntEl) cntEl.textContent = filt.length+' match'+(filt.length===1?'':'es');
    if(clrEl) clrEl.style.display = 'inline';
    if(filt.length>0 && filt.length<=200){
      const bounds = filt.map(t=>t.centroid);
      try { _jdm.fitBounds(bounds, {padding:[40,40], maxZoom:15}); } catch(e){}
    }
  } else {
    if(cntEl) cntEl.textContent = '';
    if(clrEl) clrEl.style.display = 'none';
  }
  const totalVal=filt.reduce((s,t)=>s+t.total_value,0);
  // Total visit slots = sum of per-visit outlet counts (one row per outlet-PLG-day)
  const totalVisits=filt.reduce((s,t)=>{
    if(t.visits && t.visits.length){
      return s + t.visits.reduce((vs,v)=>vs + (v.outlets||0), 0);
    }
    return s + (t.outlets_n||0);
  }, 0);
  const totalOutlets=totalVisits;
  const selN=_jdSelected.size;
  const selTrucks=trucks.filter(t=>_jdSelected.has(t.id));
  const selVal=selTrucks.reduce((s,t)=>s+t.total_value,0);
  const selOutlets=selTrucks.reduce((s,t)=>{
    if(t.visits && t.visits.length){
      return s + t.visits.reduce((vs,v)=>vs + (v.outlets||0), 0);
    }
    return s + (t.outlets_n||0);
  }, 0);
  // Comparison: total truck count both views (filter ignored to keep apples-to-apples)
  let cmpHTML='';
  const propTotal = ((TRUCKS_JUN26&&TRUCKS_JUN26.trucks)||[]).length;
  const exTotal   = (EX_TRUCKS_J26||[]).length;
  if(propTotal && exTotal){
    if(_jdView==='proposed'){
      const delta = exTotal - propTotal;
      const pct = (delta/exTotal)*100;
      const sign = delta>=0 ? '↓' : '↑';
      const col = delta>=0 ? '#15803d' : '#b91c1c';
      cmpHTML = '<div class="kpi"><div class="kpi-v" style="color:'+col+'">'+sign+' '+Math.abs(pct).toFixed(0)+'%</div>'
        +'<div class="kpi-l">vs existing ('+exTotal+' trucks)</div></div>';
    } else {
      const delta = exTotal - propTotal;
      const pct = (delta/exTotal)*100;
      const col = delta>=0 ? '#15803d' : '#b91c1c';
      cmpHTML = '<div class="kpi"><div class="kpi-v" style="color:'+col+'">↓ '+pct.toFixed(0)+'%</div>'
        +'<div class="kpi-l">proposed = '+propTotal+' trucks</div></div>';
    }
  }
  const filtDist = filt.reduce((s,t)=>s+(t.distance_km||0), 0);
  const avgDist = filt.length>0 ? filtDist/filt.length : 0;
  // Cross-view totals for headline reduction comparison
  const propTrucksAll = (TRUCKS_JUN26 && TRUCKS_JUN26.trucks) || [];
  const exTrucksAll   = EX_TRUCKS_J26 || [];
  const propTotalKm = propTrucksAll.reduce((s,t)=>s+(t.distance_km||0), 0);
  const exTotalKm   = exTrucksAll.reduce((s,t)=>s+(t.distance_km||0), 0);
  let cmpLine = '';
  if(propTrucksAll.length && exTrucksAll.length){
    if(_jdView==='proposed'){
      const dN = exTrucksAll.length - propTrucksAll.length;
      const pctN = (dN/exTrucksAll.length)*100;
      const dKm = exTotalKm - propTotalKm;
      const pctKm = (dKm/exTotalKm)*100;
      cmpLine = '<span style="color:#15803d;font-weight:700">&darr; '+Math.abs(pctN).toFixed(0)+'%</span> trucks · '
        + '<span style="color:#15803d;font-weight:700">&darr; '+Math.abs(pctKm).toFixed(0)+'%</span> km <span style="color:#9ca3af">vs existing</span>';
    } else {
      cmpLine = '<span style="color:#6b7280">Proposed: '+propTrucksAll.length+' trucks · '+propTotalKm.toFixed(0)+' km/wk</span>';
    }
  }
  // Compact 3-card KPI row + comparison line under it
  document.getElementById('jd-kpis').innerHTML=
    '<div class="kpi"><div class="kpi-v">'+filt.length+'</div><div class="kpi-l">trucks · '+totalOutlets.toLocaleString()+' visits</div></div>'
    +'<div class="kpi"><div class="kpi-v">'+totalVal.toFixed(1)+'L</div><div class="kpi-l">total value</div></div>'
    +'<div class="kpi"><div class="kpi-v">'+filtDist.toFixed(0)+'</div><div class="kpi-l">km/wk &middot; avg '+avgDist.toFixed(1)+'</div></div>';
  // Update or insert the comparison line into the selection bar parent
  let cmpBar = document.getElementById('jd-cmp-bar');
  if(!cmpBar){
    const kpiEl = document.getElementById('jd-kpis');
    cmpBar = document.createElement('div');
    cmpBar.id = 'jd-cmp-bar';
    cmpBar.style.cssText = 'font-size:11px;margin:4px 0 2px;padding:4px 8px;background:#f9fafb;border-radius:4px';
    kpiEl.parentNode.insertBefore(cmpBar, kpiEl.nextSibling);
  }
  cmpBar.innerHTML = cmpLine || '';
  // Update fixed-height selection bar (no layout shift)
  const sb = document.getElementById('jd-sel-bar') || document.getElementById('jd-selection-bar');
  if(sb){
    const txt = document.getElementById('jd-sel-text');
    const clr = document.getElementById('jd-sel-clear');
    if(selN > 0){
      sb.style.background = '#dbeafe';
      sb.style.borderColor = '#93c5fd';
      sb.style.color = '#1e40af';
      txt.innerHTML = '<b>'+selN+'</b> trip'+(selN>1?'s':'')+' selected · <b>'+selOutlets+'</b> visits · <b>'+selVal.toFixed(2)+'L</b>';
      clr.style.display = 'inline';
    } else {
      sb.style.background = '#f3f4f6';
      sb.style.borderColor = '#e5e7eb';
      sb.style.color = '#6b7280';
      txt.textContent = 'No trips selected — click a row to highlight on map';
      clr.style.display = 'none';
    }
  }
  document.getElementById('jd-truck-legend').innerHTML=
    '<span style="display:inline-flex;align-items:center;gap:4px"><span style="width:10px;height:10px;background:'+_JD_COL["3 Wheeler"]+';border-radius:50%;display:inline-block"></span>3 Wheeler &le;0.6L</span>'
    +'<span style="display:inline-flex;align-items:center;gap:4px"><span style="width:10px;height:10px;background:'+_JD_COL["Tata Ace"]+';border-radius:50%;display:inline-block"></span>Tata Ace 0.6&ndash;1.5L</span>'
    +'<span style="display:inline-flex;align-items:center;gap:4px"><span style="width:10px;height:10px;background:'+_JD_COL["Split"]+';border-radius:50%;display:inline-block"></span>Split &gt;1.5L</span>';
  // Per-day summary
  const byDay={1:[],2:[],3:[],4:[],5:[],6:[]};
  filt.forEach(t=>{if(byDay[t.delivery_day])byDay[t.delivery_day].push(t);});
  document.getElementById('jd-tbody').innerHTML=Object.keys(byDay).map(d=>{
    const arr=byDay[d];
    if(!arr.length)return '';
    const tc={'3 Wheeler':0,'Tata Ace':0,'Split':0};
    arr.forEach(t=>{if(tc[t.truck]!==undefined)tc[t.truck]++;});
    const outletsSum=arr.reduce((s,t)=>{
      if(t.visits && t.visits.length){
        return s + t.visits.reduce((vs,v)=>vs + (v.outlets||0), 0);
      }
      return s + (t.outlets_n||0);
    }, 0);
    const valSum=arr.reduce((s,t)=>s+t.total_value,0);
    return '<tr><td style="text-align:left;font-weight:600">'+_J26_DAY[d]+'</td><td>'+arr.length
      +'</td><td>'+outletsSum+'</td><td>'+valSum.toFixed(1)+'</td><td>'
      +tc['3 Wheeler']+'</td><td>'+tc['Tata Ace']+'</td><td>'+tc['Split']+'</td></tr>';
  }).join('');
  // Trip detail table — checkboxes for selection
  const trip=document.getElementById('jd-trip-table');
  if(trip){
    const rows=filt.slice();
    const visIds=new Set(rows.map(t=>t.id));
    const selVisible=rows.filter(t=>_jdSelected.has(t.id)).length;
    const allCk = (selVisible===rows.length && rows.length>0);
    const noneCk = (selVisible===0);
    trip.innerHTML=
      '<table class="dt-tbl" style="font-size:10px;width:100%"><thead><tr>'
      +'<th style="text-align:center;width:24px"><input type="checkbox" id="jd-master-cb" '
        +(allCk?'checked':'')+' onchange="jdToggleAllTrips()" title="Select / deselect all visible trips"></th>'
      +'<th style="text-align:left">#</th><th style="text-align:left">Deliv</th>'
      +'<th style="text-align:left">Salesmen (PLG &middot; DSE &middot; Day)</th>'
      +'<th>Visits</th><th>Value</th><th>km</th><th>Truck</th>'
      +'</tr></thead><tbody>'
      +rows.slice(0,600).map(t=>{
        const c=_JD_COL[t.truck]||'#666';
        const isSel=_jdSelected.has(t.id);
        const visitsHtml=t.visits.map(v=>
          '<span style="display:inline-block;background:#f3f4f6;border-radius:3px;padding:1px 4px;margin:1px 2px 1px 0">'
          +v.plg.toUpperCase()+'&middot;'+v.dse+'&middot;'+_J26_DAY[v.day]
          +(v.value>0?' &middot; <b>'+v.value.toFixed(2)+'</b>':'')
          +'</span>'
        ).join('');
        return '<tr'+(isSel?' style="background:#dbeafe"':'')+'>'
          +'<td style="text-align:center"><input type="checkbox" '+(isSel?'checked':'')
          +' onchange="jdToggleTrip(\\''+t.id+'\\')"></td>'
          +'<td style="text-align:left;font-weight:600">'+t.id+'</td>'
          +'<td style="text-align:left;font-weight:600">'+_J26_DAY[t.delivery_day]+'</td>'
          +'<td style="text-align:left">'+visitsHtml+'</td>'
          +'<td>'+t.outlets_n+'</td>'
          +'<td>'+t.total_value.toFixed(2)+'L</td>'
          +'<td>'+(t.distance_km||0).toFixed(1)+'</td>'
          +'<td style="color:'+c+';font-weight:600">'+t.truck+'</td></tr>';
      }).join('')+'</tbody></table>'
      +(rows.length>600?'<div style="color:#9ca3af;font-size:10px;padding:4px">Showing 600 of '+rows.length+' — download CSV for full list</div>':'');
    const m=document.getElementById('jd-master-cb');
    if(m && !allCk && !noneCk) m.indeterminate=true;
  }
}
function jdToggleAllTrips(){
  // Determine current filt
  const trucks=_jdTrucks();
  const filt=trucks.filter(t=>{
    if(_jdDayF!==null && t.delivery_day!==_jdDayF) return false;
    if(!_jdTruckTypes.has(t.truck)) return false;
    return true;
  });
  const allCk = filt.every(t=>_jdSelected.has(t.id));
  if(allCk) filt.forEach(t=>_jdSelected.delete(t.id));
  else filt.forEach(t=>_jdSelected.add(t.id));
  renderJD();
}

// Trigger a browser download for the delivery-detail Excel. The slide is
// rendered inside a srcdoc iframe (relative URLs resolve to about:srcdoc, and
// window.open from a sandboxed srcdoc gets blocked), so we inject an
// <a download> into the TOP document and click it programmatically.
function jdOpenDownload(view){
  const fn = view === 'existing'
    ? 'delivery_detail_existing.xlsx'
    : 'delivery_detail_proposed.xlsx';
  try{
    const topWin = window.top || window.parent || window;
    const topDoc = topWin.document;
    const origin = topWin.location.origin || window.location.origin;
    const url = origin + '/app/static/' + fn;
    const a = topDoc.createElement('a');
    a.href = url;
    a.download = fn;
    a.style.display = 'none';
    topDoc.body.appendChild(a);
    a.click();
    setTimeout(()=>topDoc.body.removeChild(a), 200);
  }catch(e){
    // Fallback: try same-doc anchor (won't escape iframe but better than nothing)
    const a = document.createElement('a');
    a.href = '/app/static/' + fn;
    a.download = fn;
    a.target = '_top';
    document.body.appendChild(a);
    a.click();
    setTimeout(()=>document.body.removeChild(a), 200);
  }
}

function j26Download(){
  const rows=[['plg','outlet_code','lat','lon','market_day','salesman_idx','beat']];
  BEATS_JUN26.forEach(b=>{
    const [lat,lon,pi,m,di,bi]=b;
    const dseInfo=DSE_JUN26.find(d=>d.idx===di)||{name:''};
    const dseShort=dseInfo.name.indexOf(':')>0?dseInfo.name.substring(dseInfo.name.indexOf(':')+1):dseInfo.name;
    rows.push([PLG_JUN26[pi].name,'',lat,lon,_J26_DAY[m+1],dseShort,bi]);
  });
  const csv=rows.map(r=>r.join(',')).join('\\n');
  const blob=new Blob([csv],{type:'text/csv'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url; a.download='aligned_beats_jun26.csv'; a.click();
  URL.revokeObjectURL(url);
}

// ── NAVIGATION ─────────────────────────────────────────────────────────────────
const slidesEl=document.getElementById('slides');
const navDots=document.querySelectorAll('.dot');
const navEl=document.getElementById('nav-dots');
const TOTAL_SLIDES=19;
const DARK_SLIDES=new Set([0,1,6,7,14]);

const sbLogo=document.getElementById('sb-logo');
function _updateLogoMode(idx){if(sbLogo)sbLogo.classList.toggle('on-dark',DARK_SLIDES.has(idx));}
function goTo(n){slidesEl.scrollTo({top:n*window.innerHeight,behavior:'smooth'});}

slidesEl.addEventListener('scroll',()=>{
  const idx=Math.round(slidesEl.scrollTop/window.innerHeight);
  navDots.forEach((d,i)=>d.classList.toggle('active',i===idx));
  navEl.classList.toggle('dark-mode',DARK_SLIDES.has(idx));
  _updateLogoMode(idx);
  if(idx===12){setTimeout(()=>{initSlide13();if(_db13m)_db13m.invalidateSize();},150);}
});
_updateLogoMode(0);

let lastWheelTime=0,gestureCount=0,lastGestureTs=0,gestureSnapped=false;
slidesEl.addEventListener('wheel',e=>{
  if(e.ctrlKey||e.metaKey)return;
  if(e.target.closest('.panel'))return; // let panels scroll independently
  const now=Date.now();
  const isNewGesture=(now-lastWheelTime)>100;
  lastWheelTime=now;
  if(isNewGesture){
    gestureSnapped=false;
    gestureCount=(now-lastGestureTs<700)?gestureCount+1:1;
    lastGestureTs=now;
  }
  if(gestureCount>=3)return;
  e.preventDefault();
  if(!gestureSnapped){
    gestureSnapped=true;
    const cur=Math.round(slidesEl.scrollTop/window.innerHeight);
    const nxt=Math.max(0,Math.min(TOTAL_SLIDES-1,cur+(e.deltaY>0?1:-1)));
    if(nxt!==cur)goTo(nxt);
  }
},{passive:false});

document.body.setAttribute('tabindex','0');
document.body.focus();
document.body.addEventListener('mouseenter',()=>document.body.focus(),{passive:true});
document.addEventListener('keydown',e=>{
  const idx=Math.round(slidesEl.scrollTop/window.innerHeight);
  if(e.key==='ArrowDown'||e.key==='PageDown'){e.preventDefault();goTo(Math.min(TOTAL_SLIDES-1,idx+1));}
  if(e.key==='ArrowUp'  ||e.key==='PageUp'  ){e.preventDefault();goTo(Math.max(0,idx-1));}
});

const obs=new IntersectionObserver(entries=>{
  entries.forEach(e=>{
    if(!e.isIntersecting)return;
    if(e.target.id==='slide-1'){initSlide1();setTimeout(()=>resizeMap('map-1'),100);}
    if(e.target.id==='slide-2'){initSlide2();setTimeout(()=>resizeMap('map-2'),100);}
    if(e.target.id==='slide-3'){initSlide3();setTimeout(()=>resizeMap('map-3'),100);}
    if(e.target.id==='slide-4'){initSlide4();setTimeout(()=>resizeMap('map-4'),100);}
    if(e.target.id==='slide-5'){initSlide5();setTimeout(()=>resizeMap('map-5'),100);}
    if(e.target.id==='slide-7'){initSlide7();setTimeout(()=>resizeMap('map-7'),100);}
    if(e.target.id==='slide-8'){renderPanel8();}
    if(e.target.id==='slide-9'){initSlide9();}
    if(e.target.id==='slide-11'){renderSlide11();}
    if(e.target.id==='slide-12'){initSlide12();}
    if(e.target.id==='slide-13'){initSlide13();}
    if(e.target.id==='slide-exbeat'){initSlideEXB();setTimeout(()=>{if(_exbm)_exbm.invalidateSize();},150);}
    if(e.target.id==='slide-jun26'){initSlideJun26();setTimeout(()=>{if(_j26m)_j26m.invalidateSize();},150);}
    if(e.target.id==='slide-jun26-terr' ){initSlideJT();  setTimeout(()=>{if(_jtm) _jtm.invalidateSize();},150);}
    // slide-jun26-zones removed
    if(e.target.id==='slide-jun26-del'  ){initSlideJD();  setTimeout(()=>{if(_jdm) _jdm.invalidateSize();},150);}
  });
},{threshold:0.25,root:slidesEl});
document.querySelectorAll('.slide').forEach(s=>obs.observe(s));
setTimeout(initSlide1,400);
navEl.classList.add('dark-mode');
</script>
</body>
</html>"""

import base64 as _b64
_logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "stackbox-logo.png")
_hul_logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "hul-logo-transparent.png")
try:
    with open(_logo_path, "rb") as _lf:
        _logo_b64 = _b64.b64encode(_lf.read()).decode()
except Exception:
    _logo_b64 = ""
try:
    with open(_hul_logo_path, "rb") as _lf:
        _hul_logo_b64 = _b64.b64encode(_lf.read()).decode()
except Exception:
    _hul_logo_b64 = ""

HTML = (HTML_TEMPLATE
    .replace("__DATA_BLOCK__", DATA_BLOCK)
    .replace("__SB_LOGO__", _logo_b64)
    .replace("__HUL_LOGO__", _hul_logo_b64))
components.html(HTML, height=800, scrolling=False)
