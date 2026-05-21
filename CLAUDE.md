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

To deploy changes: copy updated `app.py` and `data/*.json` to `apps/hul-kolkata/`, commit, PR, squash merge.
Shortcut — pull latest from this repo's remote and sync:
```bash
git pull
cp app.py ../forge-apps/apps/hul-kolkata/app.py
cp data/*.json ../forge-apps/apps/hul-kolkata/data/
```

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

## JS data structures (in `DATA_BLOCK`)
- `OUTLETS[i]` = `[lat, lon, rs_idx, new_rs_idx, outlet_name, classification, moc_2dp, primarychannel, channel_program]`
- `RS_INFO[i]` = `{idx, code, name, type, lat, lon, color, rgb, outlet_count, proposed_count, moc, gen_n, gen_moc, ws_n, ws_moc}`
- `BEATS_390[i]` = `[lat, lon, plg_idx, market_0idx, dse_idx]`
- `DSE_INFO[i]` = `{idx, name}` (S001–S033 for 218390)
- `EXCL_OUTLETS[i]` = `[lat, lon, outlet_name, rs_code, rs_lat, rs_lon, dist_km]`
- `HULL_V3_390[i]` / `HULL_EX_390[i]` = `{plg, dse, market, points:[[lat,lon],...]}` (convex hulls)
- `DSE_BALANCE_390[i]` = per-DSE balance metrics
- `CONFLICTS_EX_390[i]` / `CONFLICTS_V3_390[i]` = outlet-day conflict pairs

## Slides (10 total, 0-indexed)
| # | Title |
|---|-------|
| 0 | Title / summary stats |
| 1 | Outlets & Distributors |
| 2 | Territory Overlaps |
| 3 | Duplicate Outlets |
| 4 | High Density Clusters |
| 5 | Proposed Beats |
| 6 | Delivery Performance |
| 7 | Benefit: Same-Day Conflicts |
| 8 | Benefit: PLG Purity |
| 9 | Benefit: Territory Compactness (Jaccard) |
| 10 | Benefit: DSE Balance |

## Architecture gotchas

**MapLibre WebGL circle layers do NOT render in srcdoc iframe** — use Canvas 2D overlay instead. `maplibregl.Marker()` HTML elements do render.

**`\'` (backslash-escaped single quote) inside single-quoted JS strings causes `Uncaught SyntaxError: Unexpected string` in Chrome's srcdoc iframe context.** Node.js accepts it; Chrome V8 in srcdoc does not. Fix: use backtick template literals whenever a string needs to embed single quotes dynamically (e.g. `onclick="fn('${val}')`).

**En dash `–` (U+2013) in JS string literals can cause parse errors in srcdoc.** Use `&ndash;` HTML entity instead.

**Multi-line arrow functions inside `.map()` callbacks in string literals** can trigger ASI-related parse errors in srcdoc. Keep them on a single line.

**`delivery_data.json` is ~5.7 MB** — the total generated HTML is ~13.7 MB. This is expected.

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
