# HUL Kolkata NTO & PJP — Project Instructions

## What this is
Streamlit slide-deck app for HUL Calcutta Metro outlet analysis: NTO (New Territory Organisation) and PJP (beat planning) for RS territories 218390 and 218391.

Single file: `app.py`. All JS for all slides lives in the `HTML_TEMPLATE` triple-quoted string.

## Run (local)
```bash
python3 -m streamlit run app.py --server.port 8510
```

## Deploy (forge-apps)
Repo: `~/projects/work/stackbox/forge-apps`, app at `apps/hul-kolkata/`
Live at: `https://hul-kolkata.stackbox.tech`

To deploy changes: copy updated `app.py` and `data/*.json` to `apps/hul-kolkata/`, commit, PR, merge.
Note: forge-apps does NOT allow squash merge — use regular merge.
Shortcut — pull latest from this repo's remote and sync:
```bash
git pull
cp app.py ~/projects/work/stackbox/forge-apps/apps/hul-kolkata/app.py
cp data/*.json ~/projects/work/stackbox/forge-apps/apps/hul-kolkata/data/
```

**Local repo renamed:** was `hul-kolkata-nto-pjp`, now `hul-kolkata-app`.

## Source data (Google Drive)
```
~/Library/CloudStorage/GoogleDrive-abhishek@stackbox.xyz/My Drive/Clients Self/HUL/Sales Route/Kolkata/
  Active_Outlet_Master_Kolkata.xlsx      # sheet: Active Outlets — 50,540 outlets
  218390/All_Beat_Designs_218390_V3.xlsx # V3 beats for RS 218390
  rs_boundaries.geojson                  # territory boundaries
  output/hul_kolkata_validated.xlsx      # AI-verified duplicate pairs
```

## Cache (`data/*.json`)
All load functions cache to `data/*.json`. Source files are on Google Drive (not always accessible); if unreachable, app trusts committed JSON. Clear the relevant file to force regeneration.

| File | Clear when |
|------|-----------|
| `outlets.json` | Outlet master Excel changes OR outlet data structure changes |
| `rs_info.json` | RS assignment changes |
| `beats_390.json` / `beats_391.json` | Beats Excel changes |
| `ex_beats_390.json` / `ex_beats_391.json` | Existing beats data changes |
| `benefit_stats.json` | Hardcoded — edit directly in `load_benefits()` |
| `hull_rs_ex.json` / `hull_rs_prop.json` | outlets.json proposed assignments change |
| `rs_overlap.json` | hull_rs_ex or hull_rs_prop change |
| `rs_dist_stats.json` | outlets.json or rs_info lat/lon changes |
| `flagged_pharma_outlets.json` | Manually maintained — pharma outlets with bad geocodes |

## JS data structures (in `DATA_BLOCK`)
- `OUTLETS[i]` = `[lat, lon, rs_idx, new_rs_idx, outlet_name, classification, moc_2dp, primarychannel, channel_program, outlet_code]`
- `RS_INFO[i]` = `{idx, code, name, type, lat, lon, color, rgb, outlet_count, proposed_count, moc, proposed_moc, gained_n, gained_moc, lost_n, lost_moc, gen_n, gen_moc, ws_n, ws_moc}`
- `BEATS_390[i]` = `[lat, lon, plg_idx, market_0idx, dse_idx]`
- `DSE_INFO[i]` = `{idx, name}` (S001–S033 for 218390)
- `EXCL_OUTLETS[i]` = `[lat, lon, outlet_name, rs_code, rs_lat, rs_lon, dist_km]`
- `HULL_V3_390[i]` / `HULL_EX_390[i]` = `{plg, dse, market, points:[[lat,lon],...]}` (convex hulls)
- `HULL_RS_EX[i]` / `HULL_RS_PROP[i]` = `{rs_idx, points:[[lat,lon],...]}` (RS-level distributor hulls)
- `RS_OVERLAP` = `{General: {ex: 9.8, prop: 3.8}, Pharma: {ex: 10.3, prop: 1.9}}` — overlap % computed via Shapely
- `DSE_BALANCE_390[i]` = per-DSE balance metrics
- `CONFLICTS_EX_390[i]` / `CONFLICTS_V3_390[i]` = outlet-day conflict pairs
- `DELIVERY_DATA` = `{Existing, Output 1}` → `{Max 2/3/4 sellers}` → `{day '1'-'6'}` → array of beat objects `{id, sub_id, sellers, seller_ids, plgs, outlets, value, truck, truck_color, cost, round_trip, centroid, hull}`
- `DELIVERY_ZONES` = `{zones: [{zone, group_a_day, group_b_day, v4_hull, ex_hull, v4_area, ex_area}]}` — 6 zones, each combining 2 adjacent market days for delivery on Day N+2
- `FLAGGED_PHARMA` = pharma outlets excluded from all calculations (bad geocodes); appended to pharma CSV with "VERIFY LOCATION" note

## Slides (13 total, nav position 0-indexed)
`TOTAL_SLIDES=13`, `DARK_SLIDES=new Set([0,1,6,7])`

| Nav pos | Slide ID | Title | Label |
|---|---|---|---|
| 0 | slide-0 | Title / summary | (no label, dark) |
| 1 | slide-summary | Key Benefits | 1/13, dark |
| 2 | slide-1 | Outlets & Distributors | 2/13 |
| 3 | slide-2 | Territory Overlaps | 3/13 |
| 4 | slide-4 | High Density Clusters | 4/13 |
| 5 | slide-3 | Duplicate Outlets | 5/13 |
| 6 | slide-11 | PLG Rules | 6/13, dark |
| 7 | slide-8 | Benefit: PLG Purity | 7/13, dark |
| 8 | slide-5 | Proposed Beats | 8/13 |
| 9 | slide-9 | Beat Territories & Overlap | 9/13 |
| 10 | slide-12 | Beat Area — Delivery Zone | 10/13 |
| 11 | slide-7 | Benefit: Same-Day Conflicts | 11/13 |
| 12 | slide-13 | Delivery Beats | 12/13 |

Slides removed: slide-6 (Delivery Beats old) and slide-10 (Beat Balance / DSE Balance).

## Architecture gotchas

**MapLibre WebGL circle layers do NOT render in srcdoc iframe** — use Canvas 2D overlay instead. `maplibregl.Marker()` HTML elements do render.

**`\'` (backslash-escaped single quote) inside single-quoted JS strings causes `Uncaught SyntaxError: Unexpected string` in Chrome's srcdoc iframe context.** Node.js accepts it; Chrome V8 in srcdoc does not. Fix: use backtick template literals whenever a string needs to embed single quotes dynamically (e.g. `onclick="fn('${val}')`).

**En dash `–` (U+2013) in JS string literals can cause parse errors in srcdoc.** Use `&ndash;` HTML entity instead.

**Multi-line arrow functions inside `.map()` callbacks in string literals** can trigger ASI-related parse errors in srcdoc. Keep them on a single line.

**`delivery_data.json` is ~5.7 MB** — the total generated HTML is ~13.7 MB. This is expected.

**OFM PLG name mismatch:** `plg_info` chip names use underscores (`D_OFM`, `F+N_OFM`) but `beat_distances.json` uses hyphens (`D-OFM`, `F-OFM`). The `_CHIP_TO_DIST` map in `renderBeatDists9()` handles the translation. Do not "fix" either source — the mapping is intentional.

**Territory Overlaps slide:** shows RS distributor-level overlap (not PLG-level). Overlap = `(sum of individual hull areas − union area) / union area × 100`. Convex hulls default ON, outlet dots default OFF. Never rename "existing/proposed" to "before/after".

**Beat Territories distances:** `chain_km` = in-beat route distance per market day (from `Road Dist (km)` in Excel). Not a round trip from distributor. Label as "In-Beat Route Distance (km/market day)".

**Beat Area bundling block:** uses Day N / Day N+1 / day N+2 (not "Day b"). Static HTML — does not change with view state.

**`shapely` is a required dependency** — add to `requirements.txt` whenever `load_rs_overlap()` is present. Streamlit Cloud will fail with `ModuleNotFoundError` without it.

**WebGL context limit:** Slides 1–7 use 6 MapLibre (WebGL) contexts. Browsers support ~8–16 but Chrome silently fails at 7+ in srcdoc iframes — `map.on('load')` never fires. Slides 9, 12, and 13 use Leaflet (canvas 2D) to avoid this. Never add a new MapLibre map after slide-7.

**Delivery Beats slide (slide-13) uses Leaflet**, not MapLibre. `_db13m` is a Leaflet map; call `_db13m.invalidateSize()` not `_db13m.resize()`. Layer group `_db13lg` holds all beat polygons; call `_db13lg.clearLayers()` to reset.

**Delivery beat sub_id deduplication:** Multi-seller beats produce multiple records with identical `outlets` count. Filter `sub_id === null || sub_id === 'a'` for unique outlet count. `sub_id=null` = single-seller beat; `sub_id='a'` = first sub-beat of multi-seller. Cost should sum ALL sub-beats (each is a separate vehicle trip). Display cost in K/L: `≥100 → X.XL`, `<100 → XXK` (unit = ₹1,000).

**Delivery zone design:** 6 zones each pair 2 adjacent market days (Day N + N+1) for combined delivery on Day N+2. Each market day appears in exactly 2 zones (overlapping by design). KPI totals use a JS Set to deduplicate days — correct. Zone table rows intentionally show combined 2-day outlet counts.

**Pharma territory remapping:** 316 outlets reassigned per `pharma_p2_output (1).xlsx`. Two outlets removed as location outliers (`HUL-20B420P15390`, `HUL-20B774P684`) — stored in `flagged_pharma_outlets.json`, appended to pharma CSV download with "VERIFY LOCATION" note. When pharma data changes: update `outlets.json`, recalculate `rs_info.json` proposed stats, delete `hull_rs_prop.json` + `rs_overlap.json` to force regeneration.

**`downloadProposed()` filters by `curTerType`** — General and Pharma produce separate CSV files. Pharma CSV appends `FLAGGED_PHARMA` entries at the bottom with a Notes column.

## Canvas 2D overlay pattern
```javascript
_makeOutletCanvas(m, dpr)   // creates canvas, z-index:2, pointer-events:none
_drawGroups(m,ctx,oc,dpr,groups,tp_filter,radius_fn)  // batch draw by pre-grouped color
```

## Beats columns (218390, `load_beats()`)
`index, code, latitude, longitude, weight, frequency, clusterLabel, supercluster, beat, dse, market, PLG, Group, Same Day Visit`

- DSE: S001–S033 (33 salesmen)
- PLG: D, D+F, D+F+N, F, F+N, N, PP, PP-A, PP-B
- market: 1–6 (Mon=1 … Sat=6)
- pandas `itertuples` mangles leading-underscore columns — column `PLG` is accessed as `plg_name` (not `_plg_name`) after renaming in `load_beats()`

## Outlet comparison (finalised 2026-05-16)
Three beats sheets must all be unioned:
- "Outlet Service Info" — dummy filter: "New Beat Name" contains "Dummy"
- "HNB Outlet ME BEAT Mapping Phar" — dummy filter: "New Beat Name" contains "Dummy"
- "Fitara + Gipsy" — beat column is "Beat Names" (not "New Beat Name"); no dummies found

Join key: `Outlet Code` (active master) = `Outlet HUL Code` (beats sheets)

Final counts: Active master 50,540 | Beat codes (non-dummy) 49,898 | **Final active 49,424**
Comparison file: `My Drive/Clients Self/HUL/Sales Route/Kolkata/outlet_comparison_v2.xlsx`

## Data files
All `data/*.json` files are committed to the repo so the app works without Google Drive access (both Streamlit Cloud and forge-apps). Keep them in sync when source data changes.
