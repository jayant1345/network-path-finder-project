# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Activate virtual environment (Git Bash / WSL)
source venv/Scripts/activate

# Install dependencies
pip install -r requirements.txt

# Run the Flask server
python server.py
# → http://127.0.0.1:5001

# Quick CLI path finder (no web UI, hardcoded IPs)
python cpan.py
```

## Architecture

Single-file Flask backend (`server.py`) + single-page frontend (`templates/index.html`). No build step, no frontend framework.

### Data flow

1. `links.csv` is loaded at startup by `load_from_disk()` → `clean_and_rebuild(df)`
2. Users can upload a new CSV/Excel at runtime via `POST /upload` — triggers the same `clean_and_rebuild()` without restarting
3. `clean_and_rebuild()` produces three global state objects:
   - `df_raw` — full cleaned dataset
   - `df_down` — alarmed links (Critical/Warning), sorted by CIR descending
   - `df_10g` — 10GE links only, sorted by CIR ascending
4. Two NetworkX undirected graphs are built from **healthy links only** (alarmed links excluded):
   - `G` — all link types (10GE + GE); when both a 10GE and a GE link exist between the same node pair, 10GE wins
   - `G10` — 10GE links only; used for the guaranteed all-10GE path search

### Path finding (`POST /find_path`)

Two searches run and merge:
- **G10 search** (`cutoff=20`): finds up to 3 all-10GE paths; sorted by fewest hops → lowest avg CIR
- **G search** (`cutoff=10`): finds mixed-bandwidth paths; sorted by most 10GE links → fewest hops → lowest avg CIR; deduplicates against G10 results; fills remaining slots up to 10 total

All-10GE results always appear first. Z→A paths are the A→Z paths with nodes reversed.

### `ip_name_map`

Built from **all** links (including alarmed ones) so that autocomplete and hop labels remain complete even if a node only appears in alarmed links.

### Input data format

Required CSV/Excel columns: `A End`, `Z End`, `Bandwidth`, `CIR Utilization Ratio(%)`, `Alarm Status`, `Name`, `Update Time`, `Bandwidth Utilization Ratio(%)`.

`A End` / `Z End` values are like `10.121.24.167_SiteName_...` — the IP is extracted by splitting on `_` and taking `[0]`; the site name is `[1]`.
