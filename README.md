# CPAN Network Explorer — BSNL Gujarat NOC

A web-based network path finder and link monitoring tool built for **BSNL Gujarat's CPAN (Core Packet Access Network)**. It visualises network topology from CSV/Excel link data, finds optimal paths between any two routers, and monitors link health (alarms, CIR utilisation).

---

## Table of Contents

- [Features](#features)
- [Tech Stack](#tech-stack)
- [Project Structure](#project-structure)
- [Installation & Setup](#installation--setup)
- [How to Run](#how-to-run)
- [Input Data Format](#input-data-format)
- [Application Logic — Detailed Explanation](#application-logic--detailed-explanation)
  - [1. Data Ingestion & Cleaning](#1-data-ingestion--cleaning)
  - [2. Graph Construction](#2-graph-construction)
  - [3. Path Finding Algorithm](#3-path-finding-algorithm)
  - [4. Path Optimisation Logic](#4-path-optimisation-logic)
  - [4a. Dual-Graph Strategy for Guaranteed All-10GE Paths](#4a-dual-graph-strategy-for-guaranteed-all-10ge-paths)
  - [5. A→Z and Z→A Bidirectional Display](#5-az-and-za-bidirectional-display)
  - [6. IP Autocomplete](#6-ip-autocomplete)
  - [7. Live Data Upload & Refresh](#7-live-data-upload--refresh)
  - [8. Down Links Monitor](#8-down-links-monitor)
  - [9. 10GE Optimum List](#9-10ge-optimum-list)
- [API Reference](#api-reference)
- [UI Overview](#ui-overview)
- [Screenshots](#screenshots)

---

## Features

| Feature | Description |
|---|---|
| **Path Finder** | Find all simple paths between any two router IPs with up to 10 hops |
| **Bidirectional Paths** | Shows both A→Z and Z→A paths side by side |
| **Smart Optimisation** | Paths ranked by: max 10GE links → fewest hops → lowest avg CIR. Guaranteed all-10GE paths always surfaced first via dedicated 10GE-only graph search |
| **Per-hop Details** | Every hop shows bandwidth type (10GE/GE) and CIR utilisation % |
| **IP Autocomplete** | Type last two octets (e.g. `24.167`) and matching IPs populate instantly |
| **Live Upload** | Upload new CSV/Excel data without restarting the server |
| **Down Link Monitor** | Table of all links in Critical/Warning alarm state |
| **10GE Optimum List** | All 10GE links sorted by lowest CIR (most available capacity first) |
| **Stats Dashboard** | Live counts: Total Links, 10GE Links, Down/Alarm, Avg CIR, Last Updated |
| **Clear Button** | One-click clear of all search inputs and results |

---

## Tech Stack

| Layer | Technology |
|---|---|
| **Backend** | Python 3.x, Flask |
| **Graph Engine** | NetworkX |
| **Data Processing** | Pandas |
| **Frontend** | Vanilla HTML/CSS/JavaScript (no frameworks) |
| **Fonts** | JetBrains Mono, Syne (Google Fonts) |

---

## Project Structure

```
network-path-finder-project/
│
├── server.py              # Flask backend — all routes, graph logic, path finding
├── cpan.py                # Utility/helper script
├── links.csv              # Default network link data (auto-loaded on startup)
├── requirements.txt       # Python dependencies
│
└── templates/
    └── index.html         # Single-page frontend (HTML + CSS + JS)
```

---

## Installation & Setup

### 1. Clone the repository

```bash
git clone https://github.com/jayant1345/network-path-finder-project.git
cd network-path-finder-project
```

### 2. Create and activate virtual environment

```bash
# Create venv
python -m venv venv

# Activate — Git Bash / WSL
source venv/Scripts/activate

# Activate — Windows CMD
venv\Scripts\activate.bat

# Activate — PowerShell
venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

**Dependencies installed:**

| Package | Version | Purpose |
|---|---|---|
| Flask | 3.1.3 | Web server & routing |
| Pandas | 3.0.1 | CSV/Excel parsing & data filtering |
| NetworkX | 3.6.1 | Graph construction & path finding |
| NumPy | 2.4.3 | Numerical operations (Pandas dependency) |

---

## How to Run

```bash
python server.py
```

Open your browser and go to: **http://127.0.0.1:5001**

The server automatically loads `links.csv` from the project directory on startup.

---

## Input Data Format

The tool accepts **CSV** or **Excel (.xlsx/.xls)** files with the following required columns:

| Column | Description | Example |
|---|---|---|
| `A End` | Source device/interface | `10.121.24.167_GE0/0/1` |
| `Z End` | Destination device/interface | `10.121.24.130_GE0/0/2` |
| `Bandwidth` | Link type | `10GE` or `GE` |
| `CIR Utilization Ratio(%)` | Current bandwidth utilisation | `29.5` |
| `Alarm Status` | Link health | `Critical`, `Warning`, or normal |
| `Name` | Link name/ID | `CPAN-GJ-LINK-001` |
| `Update Time` | Last data timestamp | `2026-03-29 07:59:00` |
| `Bandwidth Utilization Ratio(%)` | Total BW utilisation | `45.2` |

> The `A End` and `Z End` fields can contain interface suffixes (e.g. `_GE0/0/1`). The tool automatically extracts only the IP address part by splitting on `_`.

---

## Application Logic — Detailed Explanation

### 1. Data Ingestion & Cleaning

**Function:** `clean_and_rebuild(df)` in `server.py`

When data is loaded (either from disk at startup or via file upload), it goes through a cleaning pipeline:

```python
# Strip all whitespace and tab characters from string columns
for col in df.columns:
    if df[col].dtype == object:
        df[col] = df[col].astype(str).str.strip().str.lstrip('\t')
```

This handles common data quality issues like:
- Leading/trailing spaces in IP addresses
- Tab characters copied from network management tools
- Inconsistent formatting in alarm status strings

After cleaning, three derived DataFrames are built:
- **`df_raw`** — full cleaned dataset
- **`df_down`** — links with `Critical` or `Warning` alarm status, sorted by CIR descending (most loaded alarms first)
- **`df_10g`** — only 10GE bandwidth links, sorted by CIR ascending (least loaded first — these are the most optimal)

---

### 2. Graph Construction

**Library:** NetworkX `nx.Graph()` (undirected graph)

After cleaning, **two graphs** are built in memory — a full mixed-bandwidth graph and a 10GE-only graph:

```python
G   = nx.Graph()   # All healthy links (10GE + GE)
G10 = nx.Graph()   # 10GE-only links — used for guaranteed all-10GE path search
```

**Key design decision — 10GE wins when multiple links exist between the same pair of nodes:**

```python
for _, row in data.iterrows():
    a, z = row['A IP'], row['Z IP']
    bw, cir = row['Bandwidth'], row['CIR Utilization Ratio(%)']
    # Prefer 10GE over GE when multiple links exist between same nodes
    if not G.has_edge(a, z) or (bw == '10GE' and G[a][z].get('bandwidth') != '10GE'):
        G.add_edge(a, z, bandwidth=bw, cir=cir)
    # G10: only 10GE edges, keep lowest CIR if duplicates
    if bw == '10GE':
        if not G10.has_edge(a, z) or cir < G10[a][z].get('cir', 999):
            G10.add_edge(a, z, bandwidth='10GE', cir=cir)
```

**Why this matters:** `nx.Graph()` only stores one edge per node pair. Without this logic, a GE link added after a 10GE link between the same nodes would silently overwrite the 10GE link, making a direct 10GE connection invisible in path finding.

**Alarmed links are excluded from both graphs:**

```python
df_healthy = df[~df['Alarm Status'].str.lower().str.contains('critical|warning', na=False)]
```

Only healthy links are added to `G` and `G10`. The IP-to-name mapping is still built from ALL links so autocomplete and hop labels remain complete even for alarmed nodes.

**What this means:**
- Each unique IP address becomes a **node** in the graph
- Each link row becomes an **edge** connecting two nodes
- The edge carries **attributes**: `bandwidth` (10GE/GE) and `cir` (utilisation %)
- The graph is **undirected** because physical network links are bidirectional

**Example:**
```
10.121.17.65 ──[10GE, CIR:28%]── 10.121.17.70
10.121.17.70 ──[GE,   CIR:12%]── 10.121.24.130
```

---

### 3. Path Finding Algorithm

**Route:** `POST /find_path`

The path finding uses a **dual-graph strategy** — two searches run in parallel and their results are merged:

**Step 1 — All-10GE guaranteed search (G10 graph, cutoff=20):**
```python
# Sort G10 paths by fewest hops, then lowest CIR (all are already 10GE)
def g10_sort_key(p):
    avg_cir = sum(G10[p[i]][p[i+1]].get('cir', 0) for i in range(len(p)-1)) / (len(p)-1)
    return (len(p), avg_cir)

all10g_paths = sorted(
    nx.all_simple_paths(G10, source, target, cutoff=20),
    key=g10_sort_key
)[:3]   # Top 3 shortest all-10GE paths
```

**Step 2 — Mixed-bandwidth search (G graph, cutoff=10):**
```python
normal_paths = []
for p in sorted(nx.all_simple_paths(G, source, target, cutoff=10), key=path_sort_key):
    if tuple(p) not in seen:   # deduplicate against all10g_paths
        normal_paths.append(p)
    if len(normal_paths) >= (10 - len(all10g_paths)):
        break
```

**Step 3 — Merge:**
```python
az_paths = all10g_paths + normal_paths   # All-10GE first, then fill remaining slots
```

**Why two separate searches?**

The all-10GE path between two nodes may be 15–20 hops long. With a single `cutoff=10` search, it would never be found. By running a separate search on `G10` (which only contains 10GE edges), we guarantee the shortest all-10GE route is always surfaced — regardless of hop count. The normal search then fills remaining result slots with the best mixed paths within 10 hops.

**Why `sorted()` instead of direct slice?**

DFS does not guarantee returning shortest paths first. Fully sorting the generator ensures the best paths always appear — even if DFS found them late in traversal.

---

### 4. Path Optimisation Logic

**The core of the tool.** The normal mixed-bandwidth search ranks paths using a 3-level sort key:

```python
def path_sort_key(p):
    count_10g = sum(1 for i in range(len(p)-1)
                    if G[p[i]][p[i+1]].get('bandwidth') == '10GE')
    avg_cir   = (sum(G[p[i]][p[i+1]].get('cir', 0) for i in range(len(p)-1))
                 / (len(p) - 1)) if len(p) > 1 else 0
    return (-count_10g, len(p), avg_cir)
```

**Priority order:**

| Priority | Criterion | Reason |
|---|---|---|
| **1st** | Maximum 10GE links (`-count_10g`, negated so higher = better) | 10GE links have 10x more capacity than GE links — all-10GE path is always best |
| **2nd** | Minimum hops (`len(p)`) | Fewer hops = lower latency, fewer failure points |
| **3rd** | Minimum average CIR utilisation (`avg_cir`) | Lower CIR = less congested path |

**Example ranking:**

```
Path A: 4 hops, 4×10GE, avg CIR 25%  ← OPTIMAL (all 10GE, then fewest hops)
Path B: 3 hops, 2×10GE, avg CIR 10%  ← 2nd (fewer hops but fewer 10GE links)
Path C: 3 hops, 2×10GE, avg CIR 20%  ← 3rd (same as B but higher CIR)
```

**Note:** All-10GE paths from the G10 dedicated search are always placed before normal mixed paths in the final result, regardless of hop count. The normal sort key is only applied within the mixed-bandwidth portion.

The results are capped at **10 paths** to keep the UI readable and the response fast.

---

### 4a. Dual-Graph Strategy for Guaranteed All-10GE Paths

**Problem:** A path that uses only 10GE links may require 15–20 hops between two distant nodes. The normal `cutoff=10` search would never discover this path, even though it is operationally superior (no GE bottlenecks, maximum capacity).

**Solution:** A second graph `G10` is built containing only 10GE edges. A separate path search runs on `G10` with a higher cutoff (20 hops), guaranteeing the shortest all-10GE route is always found.

```
┌─────────────────────────────────────────────────────────┐
│                    Path Finding Flow                     │
│                                                         │
│  G10 (10GE only, cutoff=20)                            │
│  ├─ Shortest all-10GE path → #1 result (e.g. 1 hop)   │
│  ├─ 2nd shortest all-10GE  → #2 result (e.g. 5 hops)  │
│  └─ 3rd shortest all-10GE  → #3 result (e.g. 7 hops)  │
│                                                         │
│  G (all links, cutoff=10)                              │
│  ├─ Best mixed path #1     → #4 result                 │
│  ├─ Best mixed path #2     → #5 result                 │
│  └─ ...up to 7 more        → #6–#10 results            │
└─────────────────────────────────────────────────────────┘
```

**G10 sort key (shortest hops first, since all links are already 10GE):**
```python
def g10_sort_key(p):
    avg_cir = sum(G10[p[i]][p[i+1]].get('cir', 0) for ...) / (len(p)-1)
    return (len(p), avg_cir)   # min hops → min CIR (NOT count_10g — all are 10GE)
```

Using `-count_10g` in the G10 sort key would be wrong — since all paths are 10GE, longer paths would always rank higher (more 10GE hops). The dedicated sort key ensures the shortest all-10GE path wins.

---

### 5. A→Z and Z→A Bidirectional Display

The network graph is **undirected**, meaning a path from A→Z physically works in both directions. However, to help NOC engineers visualise traffic flow in both directions, the tool displays:

- **A→Z paths** — the discovered paths as-is
- **Z→A paths** — the same paths with node order reversed

```python
za_paths = [list(reversed(p)) for p in az_paths]
```

Both sections are independently sorted with the same optimisation key, and the first card in each section is always labelled **★ OPTIMAL**.

---

### 6. IP Autocomplete

**Frontend JavaScript** — no backend dependency for filtering.

On page load, all node IPs are fetched once from `/api/nodes` and stored in memory:

```javascript
let allNodes = [];
async function loadNodes(){
    allNodes = await fetch('/api/nodes').then(r => r.json());
}
```

When a user types in the Source or Target IP box, the input is matched against all known IPs:

```javascript
const search = raw.startsWith('10.') ? raw : PREFIX + raw;
// PREFIX = '10.121.'
const matches = allNodes.filter(n => n.startsWith(search)).slice(0, 60);
```

**How it works:**
- User types `24.167` → searches for `10.121.24.167...`
- User types `10.121.24` → searches as-is (full IP prefix)
- Results show up to 60 matching IPs in a dropdown
- Arrow keys navigate, Enter selects, Escape dismisses
- Matched portion is highlighted in the dropdown using `<mark>` tags
- After a new file upload, the node list is automatically refreshed via `reloadNodes()`

---

### 7. Live Data Upload & Refresh

**Route:** `POST /upload`

Users can upload a new CSV or Excel file without restarting the server. The upload flow:

1. User selects or drags a file into the modal
2. File is sent via `multipart/form-data` `POST` to `/upload`
3. Server validates required columns are present
4. `clean_and_rebuild()` is called — this rebuilds `df_raw`, `df_down`, `df_10g`, and the NetworkX graph `G` entirely in memory
5. Response includes updated stats (total links, 10GE count, down count, avg CIR)
6. Frontend updates the stats bar, refreshes both data tables, reloads the node list for autocomplete, and shows a success toast

**Column validation:**
```python
required = {'A End', 'Z End', 'Bandwidth', 'CIR Utilization Ratio(%)', 'Alarm Status'}
missing  = required - set(df.columns)
if missing:
    return jsonify({'error': f'Missing columns: {", ".join(missing)}'}), 400
```

---

### 8. Down Links Monitor

**Route:** `GET /api/down_links`

Filters `df_down` — links where `Alarm Status` contains `critical` or `warning` (case-insensitive):

```python
df_down = df[df['Alarm Status'].str.lower().str.contains('critical|warning', na=False)]
df_down = df_down.sort_values('CIR Utilization Ratio(%)', ascending=False)
```

**Sorted by CIR descending** so the most congested alarming links appear first — the highest priority for NOC attention.

The frontend renders these with colour-coded rows:
- **Red rows** — Critical alarms
- **Amber rows** — Warning alarms

---

### 9. 10GE Optimum List

**Route:** `GET /api/optimum_10g`

Filters only `10GE` bandwidth links and sorts by **CIR ascending**:

```python
df_10g = df[df['Bandwidth'] == '10GE']
df_10g = df_10g.sort_values('CIR Utilization Ratio(%)', ascending=True)
```

**Sorted by CIR ascending** so the least loaded (most available capacity) 10GE links appear first. These are the most optimal links for traffic engineering decisions.

Each row includes a visual CIR progress bar colour-coded:
- **Green** — CIR < 50% (healthy)
- **Amber** — CIR 50–80% (moderate)
- **Red** — CIR > 80% (congested)

---

## API Reference

| Method | Endpoint | Description |
|---|---|---|
| `GET` | `/` | Serves the main single-page application |
| `GET` | `/api/meta` | Returns stats: total links, 10GE count, down count, avg CIR, last loaded |
| `GET` | `/api/nodes` | Returns sorted list of all router IP addresses in the graph |
| `GET` | `/api/down_links` | Returns top 200 links in alarm state (Critical/Warning) |
| `GET` | `/api/optimum_10g` | Returns top 300 10GE links sorted by lowest CIR |
| `POST` | `/upload` | Upload a new CSV or Excel file to refresh all data |
| `POST` | `/find_path` | Find optimised paths between source and target IP |

### `/find_path` Request

```
Content-Type: multipart/form-data
source=10.121.24.167
target=10.121.24.130
```

### `/find_path` Response

```json
{
  "source": "10.121.24.167",
  "target": "10.121.24.130",
  "az_paths": [
    {
      "nodes": ["10.121.24.167", "10.121.17.70", "10.121.24.130"],
      "hops": [
        {"from": "10.121.24.167", "to": "10.121.17.70", "bandwidth": "10GE", "cir": 28, "is_10g": true},
        {"from": "10.121.17.70", "to": "10.121.24.130", "bandwidth": "10GE", "cir": 15, "is_10g": true}
      ],
      "hop_count": 2,
      "count_10g": 2,
      "avg_cir": 21.5,
      "direction": "A→Z"
    }
  ],
  "za_paths": [ ... ]
}
```

---

## UI Overview

```
┌─────────────────────────────────────────────────────────────────┐
│  🌐 CPAN Network Explorer          [Upload & Refresh] [BSNL NOC]│
├──────────┬──────────┬──────────┬──────────┬────────────────────┤
│Total     │10GE Links│Down/Alarm│Avg CIR   │Last Updated        │
│ 4,384    │   795    │   75     │  29.2%   │29-Mar-2026 07:59   │
├──────────┴──────────┴──────────┴──────────┴────────────────────┤
│ [🔗 Path Finder] [⚡ 10GE Optimum List 795] [🔴 Down DL List 75]│
├─────────────────────────────────────────────────────────────────┤
│  // FIND NETWORK PATH                                           │
│  [Source IP ___________] [Target IP ___________] [▶ Find] [✕]  │
│                                                                 │
│  ▶ A → Z (10.121.x.x → 10.121.y.y)                            │
│  ┌─────────────────────────────────────────────┐               │
│  │ ★ OPTIMAL  2 hops  ALL 10GE  2×10GE  CIR 21%│               │
│  │ [10.121.x.x]──10GE/CIR:28%──[10.121.17.70]  │               │
│  │             ──10GE/CIR:15%──[10.121.y.y]     │               │
│  └─────────────────────────────────────────────┘               │
│                                                                 │
│  ◀ Z → A (10.121.y.y → 10.121.x.x)  [same paths reversed]     │
└─────────────────────────────────────────────────────────────────┘
```

---

## Screenshots

> Run the application and open `http://127.0.0.1:5001` to see the live UI.

---

## Author

**Jayant** — BSNL Gujarat NOC
GitHub: [@jayant1345](https://github.com/jayant1345)

---

## License

This project is for internal network operations use at BSNL Gujarat NOC.
