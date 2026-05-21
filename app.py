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
[data-testid="stCustomComponentV1"]{
    position:fixed !important; top:0; left:0;
    width:100vw !important; height:100vh !important; z-index:9999;
}
[data-testid="stCustomComponentV1"] iframe{
    width:100% !important; height:100% !important;
    border:none !important; display:block !important;
}
</style>
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
        "D","D+F","D+F+N","F","F+N","N","PP","PP-A","PP-B",
        # Specialist Sub PLGs (V3 naming)
        "D-OFM","F-OFM","N_OFM","D+F_UNIGLOW",
        "PP-A_OFM","PP-A_UNIGLOW","PP-B_OFM","PP-B_UNIGLOW",
        # Specialist Sub PLGs (V4 naming — underscore separator)
        "D_OFM","FN_OFM","PP-A_OFM","PP-B_OFM","PP-A_UNIGLOW","PP-B_UNIGLOW","D+F_UNIGLOW",
    ]
    PLG_COLORS = {
        "D":"#2563eb","D+F":"#0891b2","D+F+N":"#0d9488",
        "F":"#16a34a","F+N":"#65a30d","N":"#ca8a04",
        "PP":"#dc2626","PP-A":"#ea580c","PP-B":"#9333ea",
        # Specialist Sub PLGs (V3)
        "D-OFM":"#7c3aed","F-OFM":"#059669","N_OFM":"#b45309",
        "D+F_UNIGLOW":"#0369a1","PP-A_OFM":"#be185d","PP-A_UNIGLOW":"#9333ea",
        "PP-B_OFM":"#c2410c","PP-B_UNIGLOW":"#7c3aed",
        # Specialist Sub PLGs (V4)
        "D_OFM":"#7c3aed","FN_OFM":"#059669",
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
    plg_info = [{"idx":i,"name":p,"color":PLG_COLORS.get(p,"#6b7280")} for i,p in enumerate(plgs)]
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
                {"ex_plg":"DETS",    "v3_plg":"D",    "ex_beats":150,"ex_jac":0.0021,"v3_beats":54, "v3_jac":0.0},
                {"ex_plg":"FNB",     "v3_plg":"F",    "ex_beats":12, "ex_jac":0.0062,"v3_beats":12, "v3_jac":0.0},
                {"ex_plg":"NUTS",    "v3_plg":"N",    "ex_beats":12, "ex_jac":0.0026,"v3_beats":24, "v3_jac":0.0},
                {"ex_plg":"FNB+NUTS","v3_plg":"F+N",  "ex_beats":152,"ex_jac":0.0029,"v3_beats":33, "v3_jac":0.0},
                {"ex_plg":"D+F+NUTS","v3_plg":"D+F+N","ex_beats":39, "ex_jac":0.0057,"v3_beats":198,"v3_jac":0.0},
                {"ex_plg":"HUL+NUTS","v3_plg":"D+F+N","ex_beats":43, "ex_jac":0.0026,"v3_beats":198,"v3_jac":0.0},
                {"ex_plg":"PP",      "v3_plg":"PP",   "ex_beats":134,"ex_jac":0.0028,"v3_beats":198,"v3_jac":0.0},
                {"ex_plg":"PP-A",    "v3_plg":"PP-A", "ex_beats":50, "ex_jac":0.0059,"v3_beats":48, "v3_jac":0.0},
                {"ex_plg":"PP-B",    "v3_plg":"PP-B", "ex_beats":50, "ex_jac":0.0059,"v3_beats":48, "v3_jac":0.0},
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


outlets, rs_info, boundaries, stats, excl_outlets = load()
dupe_pairs, dupe_stats                            = load_dupes()
clusters, cluster_stats                           = load_clusters()
beats_390, beats_391, ex_beats_390, ex_beats_391, plg_info, dse_info, beat_stats = load_beats()
dse_info_391 = _load_json("dse_info_391") if os.path.exists(_json_path("dse_info_391")) else []
benefit_stats, dse_balance_390, conflicts_ex_390, conflicts_v3_390, hull_v3_390, hull_ex_390 = load_benefits()
hull_rs_ex, hull_rs_prop, rs_dist_stats = load_rs_hulls()
beat_distances   = load_beat_distances()
beat_areas       = load_beat_areas()
delivery_zones   = load_delivery_zones()

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
    "const DELIVERY_DATA    = " + json.dumps(_delivery_data).replace("</", r"<\/") + ";\n"
    "const BENEFIT_STATS    = " + json.dumps(benefit_stats)    + ";\n"
    "const DSE_BALANCE_390  = " + json.dumps(dse_balance_390)  + ";\n"
    "const CONFLICTS_EX_390 = " + json.dumps(conflicts_ex_390) + ";\n"
    "const CONFLICTS_V3_390 = " + json.dumps(conflicts_v3_390) + ";\n"
    "const HULL_V3_390      = " + json.dumps(hull_v3_390)      + ";\n"
    "const HULL_EX_390      = " + json.dumps(hull_ex_390)      + ";\n"
    "const HULL_RS_EX       = " + json.dumps(hull_rs_ex)       + ";\n"
    "const HULL_RS_PROP     = " + json.dumps(hull_rs_prop)     + ";\n"
    "const RS_DIST_STATS    = " + json.dumps(rs_dist_stats)    + ";\n"
    "const BEAT_DIST        = " + json.dumps(beat_distances)   + ";\n"
    "const BEAT_AREA        = " + json.dumps(beat_areas)       + ";\n"
    "const BEATS_V4        = " + json.dumps(_beats_v4)         + ";\n"
    "const DELIVERY_ZONES  = " + json.dumps(delivery_zones)    + ";\n"
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
.t-sub{font-size:17px;color:#94a3b8;margin-bottom:46px;max-width:520px;line-height:1.55;}
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

/* ── Slide 6 delivery map ── */
#d6-map .leaflet-control-attribution{font-size:9px;opacity:.5;}

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
#d6-beat-list .d6bc{padding:8px 12px;border-bottom:1px solid #f0f2f5;display:flex;
  align-items:flex-start;gap:7px;cursor:pointer;transition:background .1s;}
#d6-beat-list .d6bc:hover{background:#f5f8ff;}
#d6-beat-list .d6bc.sel{background:#e8f0fe;border-left:3px solid #1565C0;}
#d6-beat-list .d6bc.hid{opacity:.35;}
.d6bc-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0;margin-top:3px;}
.d6bc-info{flex:1;min-width:0;}
.d6bc-title{font-weight:600;font-size:12px;color:#222;display:flex;
  align-items:center;gap:5px;flex-wrap:wrap;}
.d6bc-meta{font-size:11px;color:#777;margin-top:2px;}
.d6bc-tags{display:flex;gap:3px;flex-wrap:wrap;margin-top:3px;}
.d6tag{padding:1px 5px;border-radius:9px;font-size:10px;font-weight:600;}
.d6t-plg{background:#fff3e0;color:#e65100;}
.d6t-val{background:#e6f4ea;color:#1e7e34;}
.d6strip{display:flex;gap:3px;margin-top:4px;flex-wrap:wrap;}
.d6sdot{display:inline-flex;align-items:center;gap:3px;font-size:10px;color:#555;}
.d6dot8{width:8px;height:8px;border-radius:50%;flex-shrink:0;display:inline-block;}
.d6eye{margin-left:auto;background:none;border:none;cursor:pointer;
  font-size:14px;opacity:.6;padding:0 2px;flex-shrink:0;align-self:center;}
.d6eye:hover{opacity:1;}
#slide-6 select{width:100%;padding:5px 8px;border:1px solid #e5e7eb;border-radius:8px;
  font-size:12px;cursor:pointer;color:#374151;background:#fff;margin-bottom:8px;}
#slide-6 select:focus{outline:none;border-color:#1565C0;}
.d6tck{display:flex;align-items:center;gap:6px;font-size:12px;cursor:pointer;padding:2px 0;}
.d6tsw{width:11px;height:11px;border-radius:50%;flex-shrink:0;display:inline-block;}
#d6-stats{display:flex;border-top:1px solid #e5e7eb;border-bottom:1px solid #e5e7eb;
  flex-shrink:0;background:#f9fafb;}
#d6-stats .kpi{flex:1;text-align:center;border-radius:0;padding:7px 2px;background:transparent;}
#d6-stats .kv{font-size:14px;}
#d6-stats .kl{font-size:9px;}
.d6-lhdr{font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;
  letter-spacing:.5px;margin:8px 0 5px;}
.d6-bp{font-size:12px;min-width:200px;}
.d6-bp h3{font-size:13px;color:#1565C0;margin-bottom:5px;}
.d6-bp table{width:100%;border-collapse:collapse;}
.d6-bp td{padding:3px 0;}
.d6-bp td:first-child{color:#666;width:95px;}
.d6-bp td:last-child{font-weight:600;}
</style>
</head>
<body>

<div id="slides">

<!-- SLIDE 0 · TITLE -->
<div class="slide" id="slide-0">
  <div class="t-badge">HUL Calcutta Metro &middot; NTO &amp; PJP</div>
  <h1 class="t-h1">New Territory<br/>Organization</h1>
  <p class="t-sub">Outlet and distributor analysis for Calcutta Metro</p>
  <div class="s-grid" id="title-stats"></div>
  <div class="scroll-h">
    <div class="arr">&#8595;</div>
    Scroll to explore &nbsp;&middot;&nbsp;
    <kbd>&#8593;</kbd><kbd>&#8595;</kbd> keys to navigate
  </div>
</div>

<!-- SLIDE 1 · OUTLETS & DISTRIBUTORS -->
<div class="slide" id="slide-1">
  <div class="map-wrap" id="map-1"></div>
  <div class="page-lbl">1 / 13 &middot; Outlets &amp; Distributors</div>
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
  <div class="page-lbl">2 / 13 &middot; Territory Overlaps</div>
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
    <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:8px 0 4px">Layers</div>
    <div class="filter-row" style="gap:4px;margin-bottom:6px">
      <button class="beat-chip" id="p2-tg-boundary" onclick="s2tgl('boundary')" style="">Boundary</button>
      <button class="beat-chip active" id="p2-tg-ret" onclick="s2tgl('ret')"    style="background:#1565C0;color:#fff;border-color:#1565C0">Retailers</button>
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

<!-- SLIDE 3 · DUPLICATE OUTLETS -->
<div class="slide" id="slide-3">
  <div class="map-wrap" id="map-3"></div>
  <div class="page-lbl">3 / 13 &middot; Duplicate Outlets</div>
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

<!-- SLIDE 4 · HIGH DENSITY CLUSTERS -->
<div class="slide" id="slide-4">
  <div class="map-wrap" id="map-4"></div>
  <div class="page-lbl">4 / 13 &middot; High Density Clusters</div>
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

<!-- SLIDE 5 · BEATS -->
<div class="slide" id="slide-5">
  <div class="map-wrap" id="map-5"></div>
  <div class="page-lbl">5 / 13 &middot; Beats &middot; RS 218390 &amp; 218391</div>
  <div class="zoom-hint">Ctrl+Scroll or Pinch to zoom</div>
  <div class="panel" style="overflow:hidden;display:flex;flex-direction:column;padding:0;">
    <div style="padding:16px 18px 10px;flex-shrink:0;border-bottom:1px solid #e5e7eb;overflow-y:auto;max-height:70vh">
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
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:0 0 4px">Filter by PLG</div>
      <div class="filter-row" id="p5-chips" style="flex-wrap:wrap;gap:4px"></div>
      <div id="p5-day-section">
        <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:6px 0 4px">Filter by Day</div>
        <div class="filter-row" id="p5-day-chips" style="flex-wrap:wrap;gap:4px"></div>
      </div>
      <div id="p5-dse-section">
        <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:6px 0 4px">Filter by Salesman</div>
        <div class="filter-row" style="flex-wrap:wrap;gap:4px" id="p5-dse-chips"></div>
      </div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:6px 0 4px">Legend</div>
      <div id="p5-legend"></div>
      <div style="margin-top:12px;border-top:1px solid #e5e7eb;padding-top:10px">
        <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">Download</div>
        <button class="dl-btn" id="p5-v4-dl" onclick="downloadV4Beats()" style="background:#7030A0">
          &#8595; Download V4 Sales Beat CSV</button>
      </div>
    </div>
    <div style="flex:1;min-height:0"></div>
  </div>
</div>

<!-- SLIDE 9 · JACCARD TERRITORIES (moved to position 6) -->
<div class="slide" id="slide-9">
  <div class="map-wrap" id="l9-map"></div>
  <div class="page-lbl">6 / 13 &middot; Beat Territories &middot; RS 218390</div>
  <div class="zoom-hint">Ctrl+Scroll or Pinch to zoom</div>
  <div class="panel" style="overflow:hidden;display:flex;flex-direction:column;padding:0">
    <div style="padding:16px 18px 10px;flex-shrink:0;border-bottom:1px solid #e5e7eb;overflow-y:auto;max-height:75vh">
      <h2 style="margin-bottom:4px">Beat Territories &amp; Jaccard</h2>
      <p class="p-sub" style="margin-bottom:8px">Convex hull per PLG-salesman-day &middot; overlap visible across PLGs</p>
      <div class="toggle-row" style="margin-bottom:10px">
        <button class="t-btn active" id="j9-vv3" onclick="setJ9View('v3')">Proposed beats</button>
        <button class="t-btn"        id="j9-vex" onclick="setJ9View('existing')">Existing beats</button>
      </div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:0 0 4px">Filter by PLG / Specialist</div>
      <div class="filter-row" id="p9-plg-chips" style="flex-wrap:wrap;gap:4px;margin-bottom:8px"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:0 0 4px">Filter by day</div>
      <div class="filter-row" id="p9-dse-chips" style="flex-wrap:wrap;gap:4px;margin-bottom:8px"></div>
      <div class="kpi-r" id="p9-kpis"></div>
      <div id="p9-dist-table"></div>
      <div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:6px 0 4px">Jaccard by PLG (lower = less overlap)</div>
      <table class="dt-tbl">
        <thead><tr>
          <th style="text-align:left">Ex PLG &rarr; Prop</th>
          <th>Ex Jac %</th><th>Prop Jac %</th>
        </tr></thead>
        <tbody id="p9-jac-body"></tbody>
      </table>
    </div>
    <div style="flex:1;min-height:0"></div>
  </div>
</div>

<!-- SLIDE 6 · DELIVERY BEATS -->
<div class="slide" id="slide-6">
  <div class="map-wrap" id="d6-map"></div>
  <div class="page-lbl">7 / 13 &middot; Delivery Beats &middot; RS 218390</div>
  <div class="zoom-hint">Ctrl+Scroll or Pinch to zoom</div>
  <div class="panel" style="overflow:hidden;display:flex;flex-direction:column;padding:0;">
    <div style="padding:16px 18px 10px;flex-shrink:0;border-bottom:1px solid #e5e7eb;overflow-y:auto;max-height:55vh;">
      <h2 style="margin-bottom:4px;">Delivery Beats</h2>
      <p class="p-sub" style="margin-bottom:10px">RS 218390 &middot; beat routing scenarios</p>
      <div class="d6-lhdr" style="margin-top:0">Beat Design</div>
      <select id="d6-design-sel"></select>
      <div class="d6-lhdr">Seller Limit</div>
      <select id="d6-scen-sel"></select>
      <div class="d6-lhdr">Market Day</div>
      <div class="filter-row" id="d6-day-btns" style="flex-wrap:wrap;gap:4px;margin-bottom:2px;"></div>
      <div class="d6-lhdr">Truck Type</div>
      <div id="d6-truck-filters" style="margin-bottom:6px;"></div>
      <div class="d6-lhdr">Layers</div>
      <div class="filter-row" style="gap:4px">
        <button class="beat-chip active" id="d6-tg-hull" onclick="d6tgl('hull')" style="background:#1565C0;color:#fff;border-color:#1565C0">Zones</button>
        <button class="beat-chip active" id="d6-tg-pts"  onclick="d6tgl('pts')"  style="background:#1565C0;color:#fff;border-color:#1565C0">Outlets</button>
        <button class="beat-chip active" id="d6-tg-cent" onclick="d6tgl('cent')" style="background:#1565C0;color:#fff;border-color:#1565C0">Trucks</button>
      </div>
    </div>
    <div id="d6-stats">
      <div class="kpi"><div class="kv" id="d6-s-beats">—</div><div class="kl">Beats</div></div>
      <div class="kpi"><div class="kv" id="d6-s-trucks">—</div><div class="kl">Trucks</div></div>
      <div class="kpi"><div class="kv" id="d6-s-cost">—</div><div class="kl">Cost</div></div>
      <div class="kpi"><div class="kv" id="d6-s-val">—</div><div class="kl">MOC/Day</div></div>
      <div class="kpi"><div class="kv" id="d6-s-rt">—</div><div class="kl">RT km</div></div>
    </div>
    <div style="display:flex;align-items:center;padding:5px 12px;border-bottom:1px solid #e5e7eb;
      flex-shrink:0;gap:6px;background:#f9fafb;">
      <span id="d6-beat-list-count" style="font-size:11px;font-weight:600;color:#6b7280;flex:1">—</span>
      <button class="dl-btn" style="width:auto;padding:4px 10px;font-size:11px;margin:0"
        onclick="d6showAll()">Show All</button>
      <button class="dl-btn" style="width:auto;padding:4px 10px;font-size:11px;margin:0;background:#6b7280"
        onclick="d6hideAll()">Hide All</button>
    </div>
    <div id="d6-beat-list" style="overflow-y:auto;flex:1;min-height:0;"></div>
  </div>
</div>

<!-- SLIDE 7 · SAME-DAY CONFLICTS -->
<div class="slide" id="slide-7">
  <div class="map-wrap" id="map-7"></div>
  <div class="page-lbl">8 / 13 &middot; Same-Day Conflicts &middot; RS 218390</div>
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

<!-- SLIDE 8 · PLG PURITY -->
<div class="slide info-slide" id="slide-8" style="background:linear-gradient(135deg,#0a1929 0%,#1a3a5c 100%)">
  <div class="page-lbl">9 / 13 &middot; PLG Purity &middot; RS 218390</div>
  <div style="max-width:860px;margin:0 auto;padding:44px 28px;color:white">
    <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;color:#60a5fa;text-transform:uppercase;margin-bottom:12px">Benefit 2 &middot; RS 218390</div>
    <h2 style="font-size:32px;font-weight:800;margin-bottom:8px;color:white">PLG Purity</h2>
    <p style="font-size:13px;color:#94a3b8;margin-bottom:24px;max-width:580px;line-height:1.6">In V4 every salesman specialises in exactly one product category. Previously 57 of 107 salesmen carried mixed portfolios across 2&ndash;3 PLG types.</p>
    <div class="bs-grid">
      <div class="bs-card" style="background:rgba(248,113,113,0.15);border:1px solid rgba(248,113,113,0.35)">
        <div class="bs-v" style="color:#f87171">57</div>
        <div class="bs-l" style="color:#94a3b8">Impure DSEs<br/>Existing design</div>
      </div>
      <div class="bs-card" style="background:rgba(74,222,128,0.15);border:1px solid rgba(74,222,128,0.35)">
        <div class="bs-v" style="color:#4ade80">0</div>
        <div class="bs-l" style="color:#94a3b8">Impure DSEs<br/>V4 design</div>
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
      <strong style="color:#e2e8f0">Why it matters:</strong> Mixed-portfolio salesmen divide attention across categories, reducing depth per PLG. V3 assigns each DSE exactly one Sub-PLG &mdash; specialist knowledge, dedicated targets, higher hit rate.
    </div>
  </div>
</div>

<!-- SLIDE 10 · BEAT BALANCE -->
<div class="slide info-slide" id="slide-10" style="background:#f8fafc">
  <div class="page-lbl">10 / 13 &middot; Beat Balance &middot; RS 218390</div>
  <div style="max-width:860px;margin:0 auto;padding:36px 28px">
    <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;color:#1565C0;text-transform:uppercase;margin-bottom:10px">Benefit 4 &middot; RS 218390</div>
    <h2 style="font-size:28px;font-weight:800;color:#111827;margin-bottom:6px">Beat Balance &mdash; Workload CV</h2>
    <p style="font-size:13px;color:#6b7280;margin-bottom:20px;max-width:600px;line-height:1.6">CV = std &divide; mean &times; 100%. Lower CV = more balanced outlet load across salesmen. Target: CV &lt; 20%.</p>
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:24px">
      <div class="kpi" style="border:1.5px solid #fee2e2">
        <div class="kv" style="color:#dc2626">24.9%</div>
        <div class="kl">Overall CV &mdash; Existing</div>
      </div>
      <div class="kpi" style="border:1.5px solid #dcfce7">
        <div class="kv" style="color:#16a34a">15.1%</div>
        <div class="kl">Overall CV &mdash; V4</div>
      </div>
      <div class="kpi" style="border:1.5px solid #dbeafe">
        <div class="kv" style="color:#2563eb">&#9660; 39%</div>
        <div class="kl">Improvement in CV</div>
      </div>
    </div>
    <div style="font-size:12px;font-weight:700;color:#6b7280;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px">CV% per PLG &mdash; <span style="color:#dc2626">&#9632; Existing</span> &nbsp; <span style="color:#16a34a">&#9632; V4</span></div>
    <div id="p10-chart"></div>
    <div style="margin-top:18px;padding:13px;background:#eff6ff;border-radius:8px;font-size:12px;color:#374151;line-height:1.6">
      <strong style="color:#1565C0">Why it matters:</strong> High CV means some salesmen are overloaded while others are underutilised. V3 balances outlets per beat within each PLG, ensuring fair and predictable workloads for all salesmen.
    </div>
  </div>
</div>

<!-- SLIDE 11 · PLG RULES -->
<div class="slide info-slide" id="slide-11" style="background:linear-gradient(135deg,#0a1929 0%,#1a3a5c 100%);overflow-y:auto">
  <div class="page-lbl">11 / 13 &middot; PLG Rules &middot; RS 218390</div>
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

<!-- SLIDE 12 · BEAT AREA PER DAY -->
<div class="slide" id="slide-12">
  <div class="map-wrap" id="l12-map"></div>
  <div class="page-lbl">12 / 13 &middot; Beat Area per Day &middot; RS 218390</div>
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
      <div style="margin-top:12px;padding:10px 12px;background:#eff6ff;border-radius:8px;font-size:11px;color:#374151;line-height:1.5">
        <strong style="color:#1565C0">Why it matters:</strong> Delivery truck covers a market zone in one trip. V4 combines 2 days of sales into one compact zone &mdash; fewer truck-km, lower logistics cost.
      </div>
    </div>
  </div>
</div>

</div><!-- /#slides -->

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

function downloadProposed(){
  const hdr=['Outlet Code','Outlet Name','Old RS Code','Old RS Name','New RS Code','New RS Name',
             'primarychannel','Classification','Channel Program','MOC'];
  const rows=[hdr,...OUTLETS.map(o=>{
    const oldRS=RS_INFO[o[2]],newRS=RS_INFO[o[3]];
    return[o[9]||'',o[4],
           oldRS?oldRS.code:'',oldRS?oldRS.name:'',
           newRS?newRS.code:'',newRS?newRS.name:'',
           o[7]||'',o[5]||'',o[8]||'',o[6]||0];
  })];
  const csv=rows.map(r=>r.map(v=>'"'+String(v||'').replace(/"/g,'""')+'"').join(',')).join('\\r\\n');
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8;'}));
  a.download='hul_kolkata_proposed_plan.csv';
  document.body.appendChild(a);a.click();document.body.removeChild(a);
}

// ── SLIDE 2 · TERRITORY OVERLAPS ─────────────────────────────────────────────
let curView='existing', curTerType='General';
let selRS2=new Set();
let _s2ShowBoundary=false, _s2ShowRetailers=true;

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
      if(!_s2ShowBoundary||selRS2.size===0)return;
      const hulls=curView==='existing'?HULL_RS_EX:HULL_RS_PROP;
      ctx2.save();
      hulls.forEach(h=>{
        const rs=RS_INFO[h.rs_idx];
        if(!rs||rs.type!==curTerType||!selRS2.has(h.rs_idx))return;
        const pts=h.points.map(p=>m.project([p[1],p[0]]));
        if(pts.length<3)return;
        ctx2.beginPath();
        pts.forEach((pt,i)=>{const x=pt.x*_DPR,y=pt.y*_DPR;i===0?ctx2.moveTo(x,y):ctx2.lineTo(x,y);});
        ctx2.closePath();
        ctx2.globalAlpha=0.1;ctx2.fillStyle=rs.color;ctx2.fill();
        ctx2.globalAlpha=0.95;ctx2.strokeStyle=rs.color;ctx2.lineWidth=2.5*_DPR;ctx2.stroke();
      });
      ctx2.globalAlpha=1;ctx2.restore();
    }
    function draw2(){
      ctx2.clearRect(0,0,oc.width,oc.height);
      if(!_s2ShowRetailers||selRS2.size===0){_drawHulls2();return;}
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
      const exV=ds.ex!=null?ds.ex.toFixed(2)+' km':'&mdash;';
      if(isProposed){
        const propV=ds.prop!=null?ds.prop.toFixed(2)+' km':'&mdash;';
        const delta=(ds.ex!=null&&ds.prop!=null)?ds.prop-ds.ex:null;
        const dStr=delta!=null?(delta>=0?'+':'')+delta.toFixed(2)+' km':'';
        const dCol=delta!=null?(delta<0?'#16a34a':'#dc2626'):'#9ca3af';
        distLine='<div style="font-size:9px;color:#9ca3af;margin-left:14px">dist: '+exV+' &rarr; '+propV
          +(dStr?'<span style="color:'+dCol+'"> ('+dStr+')</span>':'')+'</div>';
      } else {
        distLine='<div style="font-size:9px;color:#9ca3af;margin-left:14px">avg dist: '+exV+'</div>';
      }
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
let curBeatsRS='218390', curBeatsView='proposed', curBeatPLG='ALL', curBeatDay='ALL', curBeatDSE='ALL';
const _SPEC_PLG_NAMES=new Set(['D-OFM','F-OFM','N_OFM','D+F_UNIGLOW','PP-A_OFM','PP-A_UNIGLOW','PP-B_OFM','PP-B_UNIGLOW']);

function _getBeats5(){
  if(curBeatsRS==='218390')return curBeatsView==='proposed'?BEATS_390:EX_BEATS_390;
  return curBeatsView==='proposed'?BEATS_391:EX_BEATS_391;
}
function _getBgBeats5(){
  return curBeatsRS==='218390'?BEATS_391:BEATS_390;
}
function _getDseInfo5(){
  return(curBeatsRS==='218391'&&curBeatsView==='existing')?DSE_INFO_391:DSE_INFO;
}
function _hasDay5(){return curBeatsRS==='218390'||(curBeatsRS==='218391'&&curBeatsView==='existing');}

function setBeatsRS(rs){
  curBeatsRS=rs;curBeatPLG='ALL';curBeatDay='ALL';curBeatDSE='ALL';
  document.getElementById('p5-rs390').classList.toggle('active',rs==='218390');
  document.getElementById('p5-rs391').classList.toggle('active',rs==='218391');
  buildBeatChips();
  if(MAPS['map-5']&&MAPS['map-5']._draw)MAPS['map-5']._draw();
  renderPanel5();
}
function setBeatsView(v){
  curBeatsView=v;curBeatDSE='ALL';curBeatDay='ALL';
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
      // Background (other RS) in grey
      ctx5.fillStyle='#94a3b8';ctx5.globalAlpha=0.35;ctx5.beginPath();
      _getBgBeats5().forEach(pt=>{
        if(pt[0]<sl||pt[0]>nl||pt[1]<wl||pt[1]>el)return;
        const p=m.project([pt[1],pt[0]]);
        ctx5.moveTo(p.x*dpr+rBg,p.y*dpr);ctx5.arc(p.x*dpr,p.y*dpr,rBg,0,Math.PI*2);
      });
      ctx5.fill();
      // Foreground (selected RS+view) — apply filters
      const hasDay=_hasDay5();
      const dayF=(!hasDay||curBeatDay==='ALL')?null:parseInt(curBeatDay);
      const dseInfo=_getDseInfo5();
      const dseF=(!hasDay||curBeatDSE==='ALL')?null:dseInfo.findIndex(d=>d.name===curBeatDSE);
      const plgF=curBeatPLG==='ALL'?null:PLG_INFO.findIndex(p=>p.name===curBeatPLG);
      const rows=_getBeats5().filter(bt=>{
        if(plgF!==null&&bt[2]!==plgF)return false;
        if(plgF!==null&&curBeatDSE==='ALL'&&_SPEC_PLG_NAMES.has(PLG_INFO[bt[2]]?.name))return false;
        if(dayF!==null&&bt[3]!==dayF)return false;
        if(dseF!==null&&bt[4]!==dseF)return false;
        return true;
      });
      const byCol={};
      rows.forEach(bt=>{
        const pi=PLG_INFO[bt[2]];
        const col=(curBeatPLG!=='ALL'&&hasDay&&bt[3]>=0)?MKT_COLORS[bt[3]]:(pi?pi.color:'#6b7280');
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

function _activateChip(selector,key,val,color){
  document.querySelectorAll(selector).forEach(b=>{
    const isA=b.dataset[key]===val;
    b.classList.toggle('active',isA);
    b.style.background=isA?(color||'#1565C0'):'';
    b.style.color=isA?'white':'';
    b.style.borderColor=isA?(color||'#1565C0'):'';
  });
}
function setBeatPLG(plg){
  curBeatPLG=plg;curBeatDSE='ALL';
  buildBeatChips();
  if(MAPS['map-5']&&MAPS['map-5']._draw)MAPS['map-5']._draw();
  renderPanel5();
}
function setBeatDay(day){
  curBeatDay=day;
  _activateChip('[data-day]','day',day,null);
  if(MAPS['map-5']&&MAPS['map-5']._draw)MAPS['map-5']._draw();
}
function setBeatDSE(dse){
  curBeatDSE=dse;curBeatPLG='ALL';
  buildBeatChips();
  if(MAPS['map-5']&&MAPS['map-5']._draw)MAPS['map-5']._draw();
}

function renderPanel5(){
  const beats=_getBeats5();
  const rsLabel=curBeatsRS+' '+(curBeatsView==='proposed'?'Proposed':'Existing');
  document.getElementById('p5-kpis').innerHTML=
    '<div class="kpi"><div class="kv">'+fN(beats.length)+'</div><div class="kl">'+rsLabel+'</div></div>'
   +'<div class="kpi"><div class="kv">'+PLG_INFO.length+'</div><div class="kl">PLGs</div></div>';
  const hasDay=_hasDay5();
  const legendHTML=(curBeatPLG==='ALL'||!hasDay)
    ?PLG_INFO.map(p=>'<div class="rs-item"><div class="rs-dot" style="background:'+p.color+'"></div>'
      +'<span class="rs-name">'+p.name+'</span></div>').join('')
    :MKT_COLORS.map((c,i)=>'<div class="rs-item"><div class="rs-dot" style="background:'+c+'"></div>'
      +'<span class="rs-name">Market '+(i+1)+' &mdash; '+MKT_DAYS[i]+'</span></div>').join('');
  document.getElementById('p5-legend').innerHTML=legendHTML;
  document.getElementById('p5-day-section').style.display=hasDay?'':'none';
  document.getElementById('p5-dse-section').style.display=hasDay?'':'none';
}

const _P5_SPECIALISTS=[..._SPEC_PLG_NAMES];
function buildBeatChips(){
  // PLG chips + specialist chips in same row
  const dseInfo=_getDseInfo5();
  const specNames=_P5_SPECIALISTS.filter(s=>dseInfo.some(d=>d.name===s));
  const plgChips=[{name:'ALL',color:'#1565C0'},...PLG_INFO];
  const plgHtml=plgChips.map(p=>{
    const isAll=p.name==='ALL';
    const isA=isAll?(curBeatPLG==='ALL'&&curBeatDSE==='ALL'):(curBeatPLG===p.name&&curBeatDSE==='ALL');
    return'<button class="beat-chip'+(isA?' active':'')+'" data-plg="'+p.name+'" '
      +'style="'+(isA?'background:#1565C0;color:white;border-color:#1565C0;':'border-color:'+(p.color||'#e5e7eb')+';color:'+(p.color||'#374151'))+';" '
      +`onclick="setBeatPLG('${p.name}')">`+p.name+'</button>';
  }).join('');
  const specHtml=specNames.map(s=>{
    const isA=curBeatDSE===s;
    return'<button class="beat-chip'+(isA?' active':'')+'" data-dse="'+s+'" '
      +'style="font-size:10px;padding:3px 8px;'+(isA?'background:#1565C0;color:white;border-color:#1565C0;':'')+';" '
      +`onclick="setBeatDSE('${s}')">`+s+'</button>';
  }).join('');
  document.getElementById('p5-chips').innerHTML=plgHtml+'<span style="color:#d1d5db;margin:0 2px">|</span>'+specHtml;
  // Day chips
  const dayNames=['Mon','Tue','Wed','Thu','Fri','Sat'];
  const dayChips=[{val:'ALL',label:'All'},...dayNames.map((d,i)=>({val:String(i),label:d}))];
  document.getElementById('p5-day-chips').innerHTML=dayChips.map(d=>{
    const isAll=d.val==='ALL';
    return'<button class="beat-chip'+(isAll?' active':'')+'" data-day="'+d.val+'" '
      +'style="'+(isAll?'background:#1565C0;color:white;border-color:#1565C0;':'')+';" '
      +`onclick="setBeatDay('${d.val}')">`+d.label+'</button>';
  }).join('');
  // DSE chips
  const dseChips=[{name:'ALL'},...dseInfo];
  document.getElementById('p5-dse-chips').innerHTML=dseChips.map(d=>{
    const isAll=d.name==='ALL';
    return'<button class="beat-chip'+(isAll?' active':'')+'" data-dse="'+d.name+'" '
      +'style="font-size:10px;padding:3px 8px;'+(isAll?'background:#1565C0;color:white;border-color:#1565C0;':'')+';" '
      +`onclick="setBeatDSE('${d.name}')">`+(isAll?'All':d.name)+'</button>';
  }).join('');
}
buildBeatChips();

// ── SLIDE 6 · DELIVERY BEATS ──────────────────────────────────────────────────
const D6_DESIGN_KEYS=Object.keys(DELIVERY_DATA);
const D6_SCEN_KEYS=D6_DESIGN_KEYS.length?Object.keys(DELIVERY_DATA[D6_DESIGN_KEYS[0]]):[];
const D6_TRUCK_COLORS={"3 Wheeler":"#2196F3","Tata Ace":"#4CAF50","407":"#FF5722"};
const D6_DEF="#607D8B";
const D6={map:null,LG:{hull:null,pts:null,cent:null},
  curDesign:D6_DESIGN_KEYS[0]||'',curScen:D6_SCEN_KEYS[0]||'',
  curDay:1,showHull:true,showPts:true,showCent:true,selIdx:null,hiddenSet:new Set()};

function d6GetBeats(){return((DELIVERY_DATA[D6.curDesign]||{})[D6.curScen]||{})[D6.curDay]||[];}
function d6ActiveTrucks(){return new Set([...document.querySelectorAll('#d6-truck-filters input:checked')].map(x=>x.dataset.truck));}

function d6BuildSelects(){
  const ds=document.getElementById('d6-design-sel');
  D6_DESIGN_KEYS.forEach(d=>{const o=document.createElement('option');o.value=o.textContent=d;ds.appendChild(o);});
  ds.value=D6.curDesign;
  ds.onchange=()=>{D6.curDesign=ds.value;D6.hiddenSet.clear();d6Render();};
  const ss=document.getElementById('d6-scen-sel');
  D6_SCEN_KEYS.forEach(s=>{const o=document.createElement('option');o.value=o.textContent=s;ss.appendChild(o);});
  ss.value=D6.curScen;
  ss.onchange=()=>{D6.curScen=ss.value;D6.hiddenSet.clear();d6Render();};
}

function d6BuildDayBtns(){
  const c=document.getElementById('d6-day-btns');
  const days=[...new Set(
    Object.values(DELIVERY_DATA).flatMap(d=>Object.values(d).flatMap(s=>Object.keys(s).map(Number).filter(n=>!isNaN(n))))
  )].sort((a,b)=>a-b);
  const dayNames=['','Mon','Tue','Wed','Thu','Fri','Sat'];
  days.forEach(d=>{
    const b=document.createElement('button');
    const isActive=d===D6.curDay;
    b.className='beat-chip'+(isActive?' active':'');
    b.style.cssText=isActive?'background:#1565C0;color:#fff;border-color:#1565C0':'';
    b.textContent=dayNames[d]||('D'+d); b.dataset.day=d;
    b.onclick=()=>{
      D6.curDay=d; D6.hiddenSet.clear();
      c.querySelectorAll('.beat-chip').forEach(x=>{
        const on=+x.dataset.day===d;
        x.classList.toggle('active',on);
        x.style.cssText=on?'background:#1565C0;color:#fff;border-color:#1565C0':'';
      });
      d6Render();
    };
    c.appendChild(b);
  });
}

function d6BuildTruckFilters(){
  const c=document.getElementById('d6-truck-filters');
  Object.entries(D6_TRUCK_COLORS).forEach(([t,color])=>{
    const lbl=document.createElement('label');
    lbl.className='d6tck';
    lbl.innerHTML=`<input type="checkbox" checked data-truck="${t}"><span class="d6tsw" style="background:${color}"></span>${t}`;
    lbl.querySelector('input').onchange=d6Render;
    c.appendChild(lbl);
  });
}

function d6tgl(k){
  if(k==='hull')D6.showHull=!D6.showHull;
  if(k==='pts') D6.showPts =!D6.showPts;
  if(k==='cent')D6.showCent=!D6.showCent;
  const on=k==='hull'?D6.showHull:k==='pts'?D6.showPts:D6.showCent;
  const btn=document.getElementById('d6-tg-'+k);
  btn.classList.toggle('active',on);
  btn.style.cssText=on?'background:#1565C0;color:#fff;border-color:#1565C0':'';
  d6Render();
}

function d6showAll(){D6.hiddenSet.clear();d6Render();}
function d6hideAll(){
  const beats=d6GetBeats(),active=d6ActiveTrucks();
  beats.forEach((b,i)=>{if(active.has(b.truck))D6.hiddenSet.add(i);});
  d6Render();
}

function d6Render(){
  if(!D6.map)return;
  D6.LG.hull.clearLayers();D6.LG.pts.clearLayers();D6.LG.cent.clearLayers();
  const beats=d6GetBeats(),active=d6ActiveTrucks();
  let tCost=0,tTrucks=0,tRT=0,nVis=0,nShown=0;
  beats.forEach((b,idx)=>{
    if(!active.has(b.truck))return;
    nVis++;
    const hidden=D6.hiddenSet.has(idx);
    if(hidden)return;
    nShown++;
    const color=b.truck_color||D6_DEF;
    const isSel=idx===D6.selIdx;
    tCost+=b.cost; tTrucks+=1; tRT+=b.round_trip;
    if(D6.showHull&&b.is_first&&b.hull&&b.hull.length>=3){
      const poly=L.polygon(b.hull,{color,weight:isSel?2.5:1.5,fillColor:color,fillOpacity:isSel?.22:.1}).addTo(D6.LG.hull);
      poly.on('click',()=>d6selectBeat(idx));
    }
    if(D6.showPts&&b.is_first&&b.seller_pts){
      b.seller_pts.forEach(sp=>{
        sp.pts.forEach(pt=>{
          L.circleMarker(pt,{radius:3,color:sp.color,weight:.6,fillColor:sp.color,fillOpacity:.75}).addTo(D6.LG.pts);
        });
      });
    }
    if(D6.showCent&&b.centroid){
      const label=b.sub_id!=null?`${b.id}${b.sub_id}`:`${b.id}`;
      const sz=label.length>2?26:22;
      const icon=L.divIcon({html:`<div style="background:${color};color:#fff;border-radius:50%;width:${sz}px;height:${sz}px;display:flex;align-items:center;justify-content:center;font-size:${sz>22?9:10}px;font-weight:700;border:2px solid rgba(255,255,255,.9);box-shadow:0 1px 4px rgba(0,0,0,.4)">${label}</div>`,iconSize:[sz,sz],iconAnchor:[sz/2,sz/2],className:''});
      L.marker(b.centroid,{icon}).bindPopup(d6buildPopup(b),{maxWidth:280}).on('click',()=>d6selectBeat(idx)).addTo(D6.LG.cent);
    }
  });
  let totalVal=0;const valSeen=new Set();
  beats.forEach((b,idx)=>{if(!active.has(b.truck)||D6.hiddenSet.has(idx))return;if(!valSeen.has(b.id)){valSeen.add(b.id);totalVal+=b.value;}});
  document.getElementById('d6-s-beats').textContent=nShown+'/'+nVis;
  document.getElementById('d6-s-trucks').textContent=tTrucks;
  document.getElementById('d6-s-cost').textContent=tCost.toFixed(1);
  document.getElementById('d6-s-val').textContent=totalVal.toFixed(2);
  document.getElementById('d6-s-rt').textContent=tRT.toFixed(0);
  d6buildBeatList(beats,active);
}

function d6buildPopup(b){
  const label=b.sub_id!=null?`${b.id}${b.sub_id}`:`${b.id}`;
  const sellerRows=b.is_first
    ?(b.seller_pts||[]).map(sp=>`<div style="display:flex;align-items:center;gap:4px;margin-top:2px;font-size:11px"><span class="d6dot8" style="background:${sp.color}"></span><span>${sp.bid} (${sp.pts.length} outlets)</span></div>`).join('')
    :`<div style="font-size:11px;color:#888;margin-top:4px">See Beat ${b.id}a for details</div>`;
  return `<div class="d6-bp"><h3>Beat ${label}</h3><table>
    <tr><td>Truck</td><td style="color:${b.truck_color};font-weight:700">${b.truck}</td></tr>
    <tr><td>Truck load</td><td>${b.truck_val.toFixed(3)} L</td></tr>
    ${b.sub_id!=null?`<tr><td>Beat total</td><td>${b.value.toFixed(3)} L</td></tr>`:''}
    <tr><td>Outlets</td><td>${b.outlets}</td></tr>
    <tr><td>Sellers</td><td>${b.sellers} · ${b.plgs}</td></tr>
    <tr><td>Round trip</td><td>${b.round_trip} km</td></tr>
    <tr><td>Rel cost</td><td>${b.cost}</td></tr>
    </table>${b.is_first?'<div style="margin-top:6px;font-size:11px;font-weight:600;color:#333">Sellers:</div>'+sellerRows:sellerRows}</div>`;
}

function d6selectBeat(idx){
  D6.selIdx=idx;d6Render();
  const card=document.querySelector('#d6-beat-list .d6bc[data-idx="'+idx+'"]');
  if(card)card.scrollIntoView({behavior:'smooth',block:'nearest'});
}

function d6buildBeatList(beats,active){
  const c=document.getElementById('d6-beat-list');c.innerHTML='';
  let n=0;
  beats.forEach((b,idx)=>{
    if(!active.has(b.truck))return; n++;
    const hidden=D6.hiddenSet.has(idx);
    const color=b.truck_color||D6_DEF;
    const label=b.sub_id!=null?`${b.id}${b.sub_id}`:`${b.id}`;
    const valLine=b.sub_id!=null?`${b.truck_val.toFixed(3)}L (of ${b.value.toFixed(3)}L total)`:`${b.value.toFixed(3)}L`;
    const strip=b.is_first?(b.seller_pts||[]).map(sp=>`<span class="d6sdot"><span class="d6dot8" style="background:${sp.color}"></span>${sp.bid.split('|')[0]}</span>`).join(''):'';
    const card=document.createElement('div');
    card.className='d6bc'+(idx===D6.selIdx?' sel':'')+(hidden?' hid':'');
    card.dataset.idx=idx;
    card.innerHTML=`<div class="d6bc-dot" style="background:${color}"></div>
      <div class="d6bc-info">
        <div class="d6bc-title"><span>Beat ${label}</span>
          <span style="background:${color}22;color:${color};border:1px solid ${color}55;padding:1px 6px;border-radius:9px;font-size:10px;font-weight:600">${b.truck}</span>
        </div>
        <div class="d6bc-meta">${b.outlets} outlets · ${valLine} · ${b.round_trip}km RT</div>
        <div class="d6bc-tags"><span class="d6tag d6t-plg">${b.plgs}</span><span class="d6tag d6t-val">${b.sellers} seller(s)</span></div>
        ${strip?`<div class="d6strip">${strip}</div>`:''}
      </div>
      <button class="d6eye" data-idx="${idx}" title="${hidden?'Show':'Hide'}">${hidden?'🚫':'👁'}</button>`;
    card.querySelector('.d6eye').onclick=e=>{
      e.stopPropagation();
      if(D6.hiddenSet.has(idx))D6.hiddenSet.delete(idx);else D6.hiddenSet.add(idx);
      d6Render();
    };
    card.onclick=e=>{
      if(e.target.classList.contains('d6eye'))return;
      D6.selIdx=idx;if(b.centroid)D6.map.panTo(b.centroid);d6Render();
    };
    c.appendChild(card);
  });
  document.getElementById('d6-beat-list-count').textContent=n+' truck(s)';
}

function initSlide6(){
  if(D6.map)return;
  D6.map=L.map('d6-map',{zoomControl:true,preferCanvas:true}).setView([22.52,88.34],13);
  L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png',
    {attribution:'© OpenStreetMap © CARTO',subdomains:'abcd',maxZoom:19,opacity:.9}).addTo(D6.map);
  D6.LG.hull=L.layerGroup().addTo(D6.map);
  D6.LG.pts =L.layerGroup().addTo(D6.map);
  D6.LG.cent=L.layerGroup().addTo(D6.map);
  d6BuildSelects();d6BuildDayBtns();d6BuildTruckFilters();d6Render();
  setTimeout(()=>D6.map.invalidateSize(),150);
}

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
  },[88.36,22.52],12);
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

// ── SLIDE 9 · JACCARD TERRITORIES ─────────────────────────────────────────────
let curJ9View='v3',curJ9Market=0,curJ9PLG='ALL',curJ9DSE='ALL';
const _J9_SPECIALISTS=[..._SPEC_PLG_NAMES];

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

const _J9_DAYS=['All','Mon','Tue','Wed','Thu','Fri','Sat'];
function buildJ9Filters(){
  const hulls=curJ9View==='v3'?HULL_V3_390:HULL_EX_390;
  const plgs=['ALL',...[...new Set(hulls.map(h=>h.plg))].sort()];
  const plgHtml=plgs.map(p=>{
    const isA=curJ9PLG===p&&curJ9DSE==='ALL';
    const st=isA?'background:#374151;color:white;border-color:#374151;':'';
    return `<button class="beat-chip${isA?' active':''}" style="${st}" onclick="setJ9PLG('${p}')">${p}</button>`;
  }).join('');
  const specHtml=_J9_SPECIALISTS.map(s=>{
    const isA=curJ9DSE===s;
    const st=isA?'background:#374151;color:white;border-color:#374151;':'';
    return `<button class="beat-chip${isA?' active':''}" style="font-size:10px;padding:3px 8px;${st}" onclick="setJ9DSE('${s}')">${s}</button>`;
  }).join('');
  document.getElementById('p9-plg-chips').innerHTML=plgHtml+'<span style="color:#d1d5db;margin:0 2px">|</span>'+specHtml;
  document.getElementById('p9-dse-chips').innerHTML=_J9_DAYS.map((d,i)=>{
    const isA=curJ9Market===i;
    const st=isA?'background:#374151;color:white;border-color:#374151;':'';
    return `<button class="beat-chip${isA?' active':''}" style="${st}" onclick="setJ9Market(${i})">${d}</button>`;
  }).join('');
}

function setJ9View(v){
  curJ9View=v;curJ9Market=0;curJ9PLG='ALL';curJ9DSE='ALL';
  document.getElementById('j9-vv3').classList.toggle('active',v==='v3');
  document.getElementById('j9-vex').classList.toggle('active',v==='existing');
  buildJ9Filters();renderJaccard9();
}

function setJ9PLG(p){
  curJ9PLG=p;curJ9DSE='ALL';
  buildJ9Filters();renderJaccard9();
}

function setJ9Market(m){
  curJ9Market=m;
  buildJ9Filters();renderJaccard9();
}

function setJ9DSE(d){
  curJ9DSE=d;curJ9PLG='ALL';
  buildJ9Filters();renderJaccard9();
}

function renderJaccard9(){
  const state=MAPS['leaf-9'];if(!state)return;
  state.lg.clearLayers();
  const hulls=curJ9View==='v3'?HULL_V3_390:HULL_EX_390;
  let drawn=0;const bnds=[];
  hulls.forEach(h=>{
    if(curJ9Market!==0&&h.market!==curJ9Market)return;
    if(curJ9PLG!=='ALL'&&h.plg!==curJ9PLG)return;
    if(curJ9PLG!=='ALL'&&curJ9DSE==='ALL'&&_SPEC_PLG_NAMES.has(h.plg))return;
    if(curJ9DSE!=='ALL'&&h.plg!==curJ9DSE)return;
    const pts=h.hull.map(p=>[p[0],p[1]]);
    L.polygon(pts,{
      color:'#374151',weight:1.5,fillColor:'#374151',fillOpacity:0.06
    }).bindTooltip((h.plg?h.plg+' - ':'')+h.dse+' - '+_J9_DAYS[h.market]+' - '+h.n+' outlets',
      {sticky:true,direction:'top'}).addTo(state.lg);
    bnds.push(...pts);
    drawn++;
  });
  if(bnds.length>0)state.map.fitBounds(bnds,{padding:[20,20],maxZoom:14});
  const jacRow=BENEFIT_STATS.jaccard.by_plg||[];
  const jv3=(curJ9View==='v3');
  document.getElementById('p9-kpis').innerHTML='<div class="kpi"><div class="kv">'+drawn+'</div><div class="kl">'+(jv3?'Proposed':'Existing')+' beats shown</div></div>'
    +'<div class="kpi" style="background:'+(jv3?'#f0fdf4':'#fff7f7')+'"><div class="kv" style="color:'+(jv3?'#16a34a':'#dc2626')+'">'+(jv3?'0.00%':'0.21%&ndash;0.62%')+'</div><div class="kl">Avg Jaccard</div></div>';
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
  const dseF=curJ9DSE==='ALL'?null:curJ9DSE;
  const plgF=curJ9PLG==='ALL'?null:curJ9PLG;
  const avg=(arr,k)=>{const v=arr.filter(d=>d[k]!=null).map(d=>d[k]);return v.length?v.reduce((a,b)=>a+b,0)/v.length:null};
  const fmt=v=>v==null?'&mdash;':v.toFixed(1)+' km';
  // Specialist selected: show only that specialist's proposed distances
  if(dseF){
    const specD=v3.filter(d=>d.plg===dseF&&(mktF?d.market===mktF:true));
    const vc=avg(specD,'chain_km');
    el.innerHTML='<div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:10px 0 4px">In Beat Distance</div>'
      +'<table class="dt-tbl" style="width:100%"><thead><tr><th style="text-align:left">Specialist</th><th>Proposed</th></tr></thead>'
      +'<tbody><tr><td style="text-align:left">'+dseF+'</td><td>'+fmt(vc)+'</td></tr></tbody></table>';
    return;
  }
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
  el.innerHTML='<div style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.5px;margin:10px 0 4px">In Beat Distance Comparison (km)</div>'
    +'<table class="dt-tbl" style="width:100%"><thead><tr>'
    +'<th style="text-align:left">Ex PLG &rarr; Prop</th><th>Existing</th><th>Proposed</th><th>&Delta;</th>'
    +'</tr></thead><tbody>'
    +(plgF?'':'<tr style="font-weight:700;background:#f9fafb"><td style="text-align:left">All</td>'
      +'<td>'+fmt(eC)+'</td><td>'+fmt(dC)+'</td>'
      +'<td style="color:'+colC+'">'+(dChain==null?'&mdash;':(dChain<0?'':'+')+dChain.toFixed(1))+'</td>'
      +'</tr>')
    +plgRows
    +'</tbody></table>';
}

// ── SLIDE 10 · BEAT BALANCE ───────────────────────────────────────────────────
let _s10init=false;
function renderPanel10(){
  if(_s10init)return;_s10init=true;
  const byPlg=BENEFIT_STATS.balance.by_plg||[];
  const maxCV=Math.max(...byPlg.map(r=>Math.max(r.ex_cv,r.v3_cv)));
  document.getElementById('p10-chart').innerHTML=byPlg.map(r=>{
    const exW=Math.round(r.ex_cv/maxCV*100);
    const v3W=Math.round(r.v3_cv/maxCV*100);
    const target=20;const tW=Math.round(target/maxCV*100);
    return'<div class="bal-row">'
      +'<div class="bal-lbl">'+r.ex_plg+'<br/><span style="font-size:10px;color:#9ca3af">&rarr;'+r.v3_plg+'</span></div>'
      +'<div class="bar-wrap">'
        +'<div style="position:relative;height:14px;background:#f3f4f6;border-radius:3px;overflow:visible">'
          +'<div style="height:100%;width:'+exW+'%;background:#fca5a5;border-radius:3px;display:flex;align-items:center;justify-content:flex-end;padding-right:4px">'
            +'<span style="font-size:9px;font-weight:700;color:#7f1d1d">'+r.ex_cv+'%</span>'
          +'</div>'
          +'<div style="position:absolute;top:0;left:'+tW+'%;height:100%;width:2px;background:#f59e0b;z-index:1;border-radius:1px" title="Target 20%"></div>'
        +'</div>'
        +'<div style="height:14px;background:#f3f4f6;border-radius:3px;display:flex;align-items:center;margin-top:2px">'
          +'<div style="height:100%;width:'+v3W+'%;background:#86efac;border-radius:3px;display:flex;align-items:center;justify-content:flex-end;padding-right:4px">'
            +'<span style="font-size:9px;font-weight:700;color:#14532d">'+r.v3_cv+'%</span>'
          +'</div>'
        +'</div>'
      +'</div>'
      +'</div>';
  }).join('')+'<div style="font-size:10px;color:#9ca3af;margin-top:8px">&#9646; Amber line = 20% target</div>';
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

// ── NAVIGATION ─────────────────────────────────────────────────────────────────
const slidesEl=document.getElementById('slides');
const navDots=document.querySelectorAll('.dot');
const navEl=document.getElementById('nav-dots');
const TOTAL_SLIDES=13;
const DARK_SLIDES=new Set([0,8,11]);

function goTo(n){slidesEl.scrollTo({top:n*window.innerHeight,behavior:'smooth'});}

slidesEl.addEventListener('scroll',()=>{
  const idx=Math.round(slidesEl.scrollTop/window.innerHeight);
  navDots.forEach((d,i)=>d.classList.toggle('active',i===idx));
  navEl.classList.toggle('dark-mode',DARK_SLIDES.has(idx));
});

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
    if(e.target.id==='slide-6'){initSlide6();}
    if(e.target.id==='slide-7'){initSlide7();setTimeout(()=>resizeMap('map-7'),100);}
    if(e.target.id==='slide-8'){renderPanel8();}
    if(e.target.id==='slide-9'){initSlide9();}
    if(e.target.id==='slide-10'){renderPanel10();}
    if(e.target.id==='slide-11'){renderSlide11();}
    if(e.target.id==='slide-12'){initSlide12();}
  });
},{threshold:0.25,root:slidesEl});
document.querySelectorAll('.slide').forEach(s=>obs.observe(s));
setTimeout(initSlide1,400);
navEl.classList.add('dark-mode');
</script>
</body>
</html>"""

HTML = HTML_TEMPLATE.replace("__DATA_BLOCK__", DATA_BLOCK)
components.html(HTML, height=800, scrolling=False)
