from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import networkx as nx
import os
import io
import shutil
import re
import time
import threading
import uuid
import zipfile
from datetime import datetime

app = Flask(__name__)

# ── Global state ───────────────────────────────────────────────────────────────
df_raw = None
df_down = None
df_10g = None
G = nx.Graph()
G10 = nx.Graph()   # 10GE-only graph for guaranteed all-10GE path search
ip_name_map = {}   # IP → human-readable node name
last_loaded = None
loaded_filename = "links.csv"

def extract_node_name(endpoint):
    """Extract human-readable site name from endpoint string like IP_SiteName_..."""
    parts = str(endpoint).strip().split('_')
    return parts[1].strip() if len(parts) > 1 else parts[0].strip()

def clean_and_rebuild(df):
    """Clean dataframe and rebuild all derived data + graph."""
    global df_raw, df_down, df_10g, G, G10, ip_name_map

    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip().str.lstrip('\t')

    df['Bandwidth']    = df['Bandwidth'].str.strip()
    df['Alarm Status'] = df['Alarm Status'].str.strip()
    df['Name']         = df['Name'].str.strip()

    df_raw  = df
    df_down = df[df['Alarm Status'].str.lower().str.contains('critical|warning', na=False)].copy()
    df_down = df_down.sort_values('CIR Utilization Ratio(%)', ascending=False)

    df_10g  = df[df['Bandwidth'] == '10GE'].copy()
    df_10g  = df_10g.sort_values('CIR Utilization Ratio(%)', ascending=True)

    # Rebuild graph — exclude alarmed (down) links so path finder only uses healthy links
    df_healthy = df[~df['Alarm Status'].str.lower().str.contains('critical|warning', na=False)].copy()
    data = df_healthy[['A End', 'Z End', 'Bandwidth', 'CIR Utilization Ratio(%)']].copy()
    data['A IP'] = data['A End'].astype(str).str.strip().str.split('_').str[0]
    data['Z IP'] = data['Z End'].astype(str).str.strip().str.split('_').str[0]

    # Build IP → node name mapping (from all links so names are always available)
    ip_name_map = {}
    for _, row in df[['A End', 'Z End']].iterrows():
        a_ip = str(row['A End']).strip().split('_')[0]
        z_ip = str(row['Z End']).strip().split('_')[0]
        if a_ip not in ip_name_map:
            ip_name_map[a_ip] = extract_node_name(row['A End'])
        if z_ip not in ip_name_map:
            ip_name_map[z_ip] = extract_node_name(row['Z End'])

    G = nx.Graph()
    G10 = nx.Graph()
    for _, row in data.iterrows():
        a, z = row['A IP'], row['Z IP']
        bw  = row['Bandwidth']
        cir = row['CIR Utilization Ratio(%)']
        # Prefer 10GE over GE when multiple links exist between same nodes
        if not G.has_edge(a, z) or (bw == '10GE' and G[a][z].get('bandwidth') != '10GE'):
            G.add_edge(a, z, bandwidth=bw, cir=cir)
        # 10GE-only graph for guaranteed all-10GE path search
        if bw == '10GE':
            if not G10.has_edge(a, z) or cir < G10[a][z].get('cir', 999):
                G10.add_edge(a, z, bandwidth='10GE', cir=cir)

# ── Initial load from disk ─────────────────────────────────────────────────────
def load_from_disk():
    global last_loaded, loaded_filename
    path = "links.csv"
    df = pd.read_csv(path)
    clean_and_rebuild(df)
    last_loaded   = datetime.fromtimestamp(os.path.getmtime(path)).strftime('%d-%b-%Y %H:%M')
    loaded_filename = os.path.basename(path)

load_from_disk()

# ── Update alarm status from Down-DL-list CSV ─────────────────────────────────
def _update_alarm_from_dl_down(dl_down_path, report_date=None):
    """
    After report generation, sync df_raw alarm status with the latest
    Down-DL-list CSV so the path finder and dashboard reflect current
    DL down state without a manual re-upload.
    last_loaded is set to the report date (not current time) so the
    dashboard shows when the DATA is from, not when the sync ran.
    """
    global last_loaded
    if df_raw is None:
        return

    try:
        try:
            dl_df = pd.read_csv(dl_down_path, dtype=str, encoding='utf-8-sig',
                                on_bad_lines='warn')
        except Exception:
            dl_df = pd.read_csv(dl_down_path, dtype=str, encoding='latin1',
                                on_bad_lines='warn')

        dl_df.columns = [c.strip() for c in dl_df.columns]
        dl_df['Name']         = dl_df['Name'].str.replace('\t', '', regex=False).str.strip()
        dl_df['Alarm Status'] = dl_df['Alarm Status'].str.replace('\t', '', regex=False).str.strip()

        # Build name → alarm_status map from the Down-DL-list
        alarm_map = dict(zip(dl_df['Name'], dl_df['Alarm Status']))

        # Update df_raw alarm status for every link that appears in the Down-DL-list
        updated = df_raw.copy()
        def _new_alarm(row):
            name = str(row.get('Name', '')).strip()
            return alarm_map.get(name, row.get('Alarm Status', ''))

        updated['Alarm Status'] = updated.apply(_new_alarm, axis=1)
        clean_and_rebuild(updated)

        # Show report date on dashboard (data is FROM that date, not today)
        if report_date:
            last_loaded = report_date.strftime('%d-%b-%Y') + ' (report)'
        else:
            last_loaded = datetime.now().strftime('%d-%b-%Y %H:%M')
    except Exception:
        pass

# ── Helpers ────────────────────────────────────────────────────────────────────
def get_edge_info(path):
    hops = []
    for i in range(len(path) - 1):
        u, v = path[i], path[i+1]
        ed   = G.get_edge_data(u, v, default={})
        bw   = ed.get('bandwidth', 'GE')
        cir  = ed.get('cir', 0)
        hops.append({'from': u, 'to': v,
                     'bandwidth': bw,
                     'cir': int(cir) if cir == cir else 0,
                     'is_10g': bw == '10GE'})
    return hops

# ── Routes ─────────────────────────────────────────────────────────────────────
@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/meta')
def meta():
    return jsonify({
        'total':    len(df_raw),
        'links_10g': len(df_10g),
        'down':     len(df_down),
        'avg_cir':  round(df_10g['CIR Utilization Ratio(%)'].mean(), 1) if len(df_10g) else 0,
        'last_loaded': last_loaded,
        'filename':  loaded_filename,
    })

@app.route('/upload', methods=['POST'])
def upload():
    global last_loaded, loaded_filename
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in request'}), 400
    f = request.files['file']
    if f.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    fname = f.filename.lower()
    try:
        if fname.endswith('.csv'):
            df = pd.read_csv(io.BytesIO(f.read()))
        elif fname.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(io.BytesIO(f.read()))
        else:
            return jsonify({'error': 'Only CSV or Excel files are supported'}), 400

        # Validate required columns
        required = {'A End', 'Z End', 'Bandwidth', 'CIR Utilization Ratio(%)', 'Alarm Status'}
        missing  = required - set(df.columns)
        if missing:
            return jsonify({'error': f'Missing columns: {", ".join(missing)}'}), 400

        clean_and_rebuild(df)
        last_loaded    = datetime.now().strftime('%d-%b-%Y %H:%M')
        loaded_filename = f.filename

        return jsonify({
            'success': True,
            'message': f'Loaded {len(df_raw):,} links from {f.filename}',
            'total':   len(df_raw),
            'links_10g': len(df_10g),
            'down':    len(df_down),
            'avg_cir': round(df_10g['CIR Utilization Ratio(%)'].mean(), 1) if len(df_10g) else 0,
            'last_loaded': last_loaded,
            'filename': loaded_filename,
        })
    except Exception as e:
        return jsonify({'error': f'Parse error: {str(e)}'}), 400

@app.route('/find_path', methods=['POST'])
def find_path():
    source = request.form.get('source', '').strip()
    target = request.form.get('target', '').strip()
    if not source or not target:
        return jsonify({'error': 'Please enter both Source and Target IP'})
    if source not in G:
        return jsonify({'error': f'Source IP "{source}" not found in network'})
    if target not in G:
        return jsonify({'error': f'Target IP "{target}" not found in network'})
    if not nx.has_path(G, source, target):
        return jsonify({'error': 'No path exists between the given nodes'})

    def build_results(paths, direction):
        out = []
        for path in paths:
            hops      = get_edge_info(path)
            count_10g = sum(1 for h in hops if h['is_10g'])
            cir_vals  = [h['cir'] for h in hops]
            avg_cir   = round(sum(cir_vals) / len(cir_vals), 1) if cir_vals else 0
            out.append({'nodes': path, 'hops': hops,
                        'hop_count': len(path) - 1,
                        'count_10g': count_10g,
                        'avg_cir':   avg_cir,
                        'direction': direction})
        return out

    # Sort candidates: max 10G links → fewest hops → lowest avg CIR
    def path_sort_key(p):
        count_10g = sum(1 for i in range(len(p)-1)
                        if G[p[i]][p[i+1]].get('bandwidth') == '10GE')
        avg_cir   = (sum(G[p[i]][p[i+1]].get('cir', 0) for i in range(len(p)-1))
                     / (len(p) - 1)) if len(p) > 1 else 0
        return (-count_10g, len(p), avg_cir)

    # Step 1: find best all-10GE paths via 10GE-only graph (no cutoff limit)
    # Sort by fewest hops then lowest CIR — NOT by count_10g (all are 10GE here)
    def g10_sort_key(p):
        avg_cir = (sum(G10[p[i]][p[i+1]].get('cir', 0) for i in range(len(p)-1))
                   / (len(p) - 1)) if len(p) > 1 else 0
        return (len(p), avg_cir)

    all10g_paths = []
    if source in G10 and target in G10 and nx.has_path(G10, source, target):
        try:
            all10g_paths = sorted(
                nx.all_simple_paths(G10, source, target, cutoff=20),
                key=g10_sort_key
            )[:3]
        except Exception:
            pass

    # Step 2: normal mixed-bandwidth search within cutoff=10
    seen = {tuple(p) for p in all10g_paths}
    normal_paths = []
    for p in sorted(nx.all_simple_paths(G, source, target, cutoff=10), key=path_sort_key):
        if tuple(p) not in seen:
            normal_paths.append(p)
            seen.add(tuple(p))
        if len(normal_paths) >= (10 - len(all10g_paths)):
            break

    # Merge: all-10GE paths first, then remaining slots filled from normal search
    az_paths = all10g_paths + normal_paths
    za_paths = [list(reversed(p)) for p in az_paths]

    az_results = build_results(az_paths, 'A→Z')
    za_results = build_results(za_paths, 'Z→A')

    return jsonify({
        'az_paths': az_results,
        'za_paths': za_results,
        'source': source,
        'target': target,
        'names': ip_name_map,
    })

@app.route('/api/nodes')
def get_nodes():
    return jsonify(sorted(G.nodes()))

@app.route('/api/node_names')
def get_node_names():
    return jsonify(ip_name_map)

@app.route('/api/down_links')
def down_links():
    rows = df_down[['Name','Bandwidth','Alarm Status',
                    'CIR Utilization Ratio(%)','A End','Z End',
                    'Update Time']].head(200).to_dict('records')
    return jsonify(rows)

@app.route('/api/optimum_10g')
def optimum_10g():
    rows = df_10g[['Name','CIR Utilization Ratio(%)',
                   'Bandwidth Utilization Ratio(%)',
                   'A End','Z End','Update Time']].head(300).to_dict('records')
    return jsonify(rows)

# ── Native pickers (runs tkinter on the server machine) ───────────────────────
_CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'report_config.json')

def _load_config():
    try:
        import json
        with open(_CONFIG_FILE) as f:
            return json.load(f)
    except Exception:
        return {}

def _save_config(data):
    import json
    cfg = _load_config()
    cfg.update(data)
    with open(_CONFIG_FILE, 'w') as f:
        json.dump(cfg, f, indent=2)

# Load saved paths on startup
_cfg = _load_config()
_custom_template_path = _cfg.get('template_path')   # persists across restarts
_custom_output_dir    = _cfg.get('output_dir')       # custom save folder (None = default reports_output)

@app.route('/report/browse_folder')
def report_browse_folder():
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder = filedialog.askdirectory(title="Select folder containing input files")
    root.destroy()
    if folder:
        return jsonify({'folder': folder.replace('/', '\\')})
    return jsonify({'folder': None})

@app.route('/report/browse_output_folder')
def report_browse_output_folder():
    global _custom_output_dir
    import tkinter as tk
    from tkinter import filedialog
    default = _custom_output_dir or os.path.join(os.path.dirname(os.path.abspath(__file__)), 'reports_output')
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder = filedialog.askdirectory(title="Select folder to save reports", initialdir=default)
    root.destroy()
    if folder:
        folder = folder.replace('/', '\\')
        _custom_output_dir = folder
        _save_config({'output_dir': folder})
        return jsonify({'folder': folder})
    return jsonify({'folder': None})

@app.route('/report/open_output_folder')
def report_open_output_folder():
    folder = _custom_output_dir or os.path.join(os.path.dirname(os.path.abspath(__file__)), 'reports_output')
    os.makedirs(folder, exist_ok=True)
    os.startfile(folder)
    return jsonify({'folder': folder})

@app.route('/report/get_output_folder')
def report_get_output_folder():
    default = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'reports_output')
    folder  = _custom_output_dir or default
    return jsonify({'folder': folder, 'is_custom': bool(_custom_output_dir)})

@app.route('/report/browse_template')
def report_browse_template():
    global _custom_template_path
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    path = filedialog.askopenfilename(
        title="Select CPAN Master Template",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    root.destroy()
    if path:
        path = path.replace('/', '\\')
        _custom_template_path = path
        _save_config({'template_path': path})
        return jsonify({'path': path, 'name': os.path.basename(path)})
    return jsonify({'path': None})

# ── CPAN Daily Report Generation ──────────────────────────────────────────────
_REPORT_GREEN_COLOR = "FF00FF00"
_REPORT_FILE_PATTERNS = {
    "card_off":   ["CARD OFFLINE"],
    "device_off": ["DEVICE OFFLINE"],
    "fan_fail":   ["FAN FAILURE"],
    "dl_fail":    ["CPAN DL FAIL REPORT", "DL FAIL REPORT"],  # prefix varies by date
    "dash_down":  ["DASH-DOWN", "DASH_DOWN"],
    "dl_down":    ["Down-DL-list"],   # Down-DL-list_dd-mm-yyyy.csv
    "dl_alarms":  ["DL-alarms"],      # DL-alarms_dd-mm-yyyy.csv
}
_REPORT_SHEET_CONFIG = {
    "card_off":   {"sheet_name": "CARD-OFF",   "data_start_row": 11,
                   "paste_col_start": 6,  "paste_col_end": 13,
                   "input_sheet": "Sheet1", "input_data_start_row": 2,
                   "input_col_start": 1, "input_col_end": 8},
    "device_off": {"sheet_name": "DEVICE-OFF", "data_start_row": 11,
                   "paste_col_start": 6,  "paste_col_end": 13,
                   "input_sheet": "Sheet1", "input_data_start_row": 2,
                   "input_col_start": 1, "input_col_end": 8},
    "fan_fail":   {"sheet_name": "FAN-FAIL",   "data_start_row": 11,
                   "paste_col_start": 6,  "paste_col_end": 14,
                   "input_sheet": "Sheet1", "input_data_start_row": 2,
                   "input_col_start": 1, "input_col_end": 9},
    "dl_fail":    {"sheet_name": "DL-FAIL",    "data_start_row": 11,
                   "paste_col_start": 7,  "paste_col_end": 17,
                   "input_sheet": "Sheet1", "input_data_start_row": 2,
                   "input_col_start": 1, "input_col_end": 11},
    # DASH-DOWN: all columns pasted dynamically (paste_col_end/input_col_end = None means all)
    "dash_down":  {"sheet_name": "DASH-DOWN", "data_start_row": 2,
                   "paste_col_start": 1,  "paste_col_end": None,
                   "input_sheet": "Sheet1", "input_data_start_row": 2,
                   "input_col_start": 1, "input_col_end": None},
}
_report_jobs = {}  # job_id -> {status, logs, output_path, error}

def _find_report_template():
    global _custom_template_path
    if _custom_template_path and os.path.isfile(_custom_template_path):
        return _custom_template_path
    base = os.path.dirname(os.path.abspath(__file__))
    for f in os.listdir(base):
        if f.upper().startswith('DAILY_CPAN_REPORTS_MASTER') and f.lower().endswith('.xlsx'):
            return os.path.join(base, f)
    return None

def _find_report_input_files(folder, date_str):
    # date_str arrives as DD-MM-YYYY from the UI
    # Filenames use inconsistent date formats across dates:
    #   DD-MM-YYYY  e.g. CARD OFFLINE 05-05-2026.xlsx
    #   DD.MM.YYYY  e.g. CARD OFFLINE 07.05.2026.xlsx
    # DASH-DOWN is a CSV, filename varies: DASH-DOWN_DDMMYYYY.csv or DASH_DOWN.csv
    found = {}
    try:
        parts = date_str.split('-')
        dd, mm, yyyy = parts[0], parts[1], parts[2]
        yy = yyyy[2:]
        # All date variants to try when matching filenames
        date_tokens = [
            f"{dd}-{mm}-{yyyy}",   # DD-MM-YYYY  (most common)
            f"{dd}.{mm}.{yyyy}",   # DD.MM.YYYY
            f"{dd}.{mm}.{yy}",     # DD.MM.YY
            f"{dd}{mm}{yyyy}",     # DDMMYYYY    (DASH-DOWN compact)
            f"{dd}{mm}{yy}",       # DDMMYY
        ]

        all_files = os.listdir(folder)
        all_xlsx  = [f for f in all_files if f.lower().endswith('.xlsx')]
        all_csv   = [f for f in all_files if f.lower().endswith('.csv')]

        _CSV_KEYS = {'dash_down', 'dl_down', 'dl_alarms'}

        for key, prefixes in _REPORT_FILE_PATTERNS.items():
            if key in _CSV_KEYS:
                # CSV-based files (also accept xlsx); dash_down may have no date in name
                pool = all_csv + [f for f in all_xlsx
                                  if any(p.upper() in f.upper() for p in prefixes)]
                # First try: match by prefix + any date token
                matches = []
                for tok in date_tokens:
                    matches = [f for f in pool
                               if any(f.upper().startswith(p.upper()) for p in prefixes)
                               and tok in f]
                    if matches:
                        break
                # Fallback: any file starting with a known prefix (no date required)
                if not matches:
                    matches = [f for f in pool
                               if any(f.upper().startswith(p.upper()) for p in prefixes)]
            else:
                # Regular xlsx files — try each date token variant
                matches = []
                for tok in date_tokens:
                    candidates = [f for f in all_xlsx if tok in f]
                    matches = [f for f in candidates
                               if any(f.upper().startswith(p.upper()) for p in prefixes)]
                    if matches:
                        break

            if not matches:
                continue
            clean = [f for f in matches if not re.search(r'\s*\(\d+\)', f)]
            found[key] = os.path.join(folder, clean[0] if clean else matches[0])
    except Exception:
        pass
    return found

def _run_report_job(job_id, file_map, output_path, report_date, links_snap=None):
    import pythoncom
    pythoncom.CoInitialize()
    job = _report_jobs[job_id]
    def log(msg):
        job['logs'].append(msg)
    try:
        from report_gen import generate_report
        abs_output = os.path.abspath(output_path)
        pdf_path   = os.path.splitext(abs_output)[0] + '.pdf'

        log("Processing input files with Python...")
        gen_logs = generate_report(file_map, report_date, abs_output,
                                   links_df=links_snap)
        for l in gen_logs:
            log(l)

        # Sync alarm status from Down-DL-list → update dashboard & path finder
        if file_map.get('dl_down') and os.path.isfile(file_map['dl_down']):
            log("Updating dashboard alarm status from Down-DL-list...")
            _update_alarm_from_dl_down(file_map['dl_down'], report_date)
            log(f"  Dashboard updated: {len(df_down)} down links, "
                f"{len(df_raw)} total links in network.")

        # Export each sheet as a separate PDF, then zip them
        log("Exporting PDFs (one per report sheet)...")
        import win32com.client
        reports_dir = os.path.dirname(abs_output)
        base_name   = os.path.splitext(os.path.basename(abs_output))[0]
        zip_path    = os.path.join(reports_dir, base_name + '_PDFs.zip')
        # Sheet name → PDF label matching existing naming convention
        _SHEET_PDF_LABEL = {
            'CARD-OFF-R':   'CARD OFFLINE REPORT',
            'DEVICE-OFF-R': 'DEVICE OFFLINE REPORT',
            'FAN-FAIL-R':   'FAN FAIL REPORT',
            'DL-FAIL-R':    'DL FAIL REPORT',
            'DASH-DOWN-R':  'DASH DOWN REPORT',
            'DL-DOWN-RT':   'DL DOWN REALTIME REPORT',
        }
        # Date in DD.MM.YYYY format for PDF filename
        pdf_date = report_date.strftime('%d.%m.%Y')
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible        = False
        excel.DisplayAlerts  = False
        excel.ScreenUpdating = False
        excel.EnableEvents   = False
        pdf_files = []
        try:
            wb_com = excel.Workbooks.Open(abs_output, UpdateLinks=False)
            for i in range(1, wb_com.Sheets.Count + 1):
                sh      = wb_com.Sheets(i)
                sh_name = sh.Name
                label   = _SHEET_PDF_LABEL.get(sh_name, sh_name)
                pdf_out = os.path.join(reports_dir, f"CPAN_{label}_{pdf_date}.pdf")
                if os.path.exists(pdf_out):
                    os.remove(pdf_out)
                # Force landscape + fit-all-columns-on-one-page-wide before export
                try:
                    ps = sh.PageSetup
                    ps.Orientation    = 2   # xlLandscape
                    ps.PaperSize      = 9   # xlPaperA4
                    ps.Zoom           = False
                    ps.FitToPagesWide = 1
                    ps.FitToPagesTall = 9999
                except Exception:
                    try:
                        sh.PageSetup.Orientation = 2
                        sh.PageSetup.Zoom = 70
                    except Exception:
                        pass
                sh.ExportAsFixedFormat(0, pdf_out, OpenAfterPublish=False)
                pdf_files.append(pdf_out)
                log(f"  PDF: {os.path.basename(pdf_out)}")
            wb_com.Close(SaveChanges=False)
        finally:
            try:
                excel.ScreenUpdating = True
                excel.EnableEvents   = True
                excel.Quit()
            except Exception:
                pass

        # Bundle into a single ZIP for easy download
        if os.path.exists(zip_path):
            os.remove(zip_path)
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for p in pdf_files:
                zf.write(p, os.path.basename(p))
        log(f"ZIP created: {len(pdf_files)} PDFs bundled.")
        job['pdf_path'] = zip_path

        log("Done! xlsx and PDF ready for download.")
        job['status']      = 'done'
        job['output_path'] = abs_output

        # Clean up uploaded temp dir (LAN mode)
        td = job.get('_temp_dir')
        if td and os.path.isdir(td):
            import shutil
            try:
                shutil.rmtree(td)
                log(f"Temp upload folder cleaned up.")
            except Exception:
                pass
    except Exception as e:
        import traceback
        job['status'] = 'error'
        job['error']  = traceback.format_exc()
        log(f"ERROR: {str(e)}")
    finally:
        pythoncom.CoUninitialize()

@app.route('/report/is_local')
def report_is_local():
    """Return whether the request is from the server machine itself."""
    remote = request.remote_addr
    return jsonify({'local': remote in ('127.0.0.1', '::1')})

@app.route('/report/upload_inputs', methods=['POST'])
def report_upload_inputs():
    """Accept input files uploaded from a LAN browser, store in a temp dir."""
    import tempfile
    files = request.files.getlist('files')
    if not files or not any(f.filename for f in files):
        return jsonify({'error': 'No files provided'}), 400
    temp_dir = tempfile.mkdtemp(prefix='cpan_inputs_')
    saved = []
    for f in files:
        fname = os.path.basename(f.filename)
        if fname:
            f.save(os.path.join(temp_dir, fname))
            saved.append(fname)
    return jsonify({'temp_dir': temp_dir, 'count': len(saved), 'files': saved})

@app.route('/report/template_status')
def report_template_status():
    # Master template no longer required — report is generated purely in Python
    return jsonify({'found': True, 'name': 'Python generator (no template needed)', 'path': None})

@app.route('/report/scan', methods=['POST'])
def report_scan():
    global _custom_template_path
    data     = request.json or {}
    folder   = (data.get('folder') or '').strip()
    date_str = (data.get('date')   or '').strip()
    if not folder or not date_str:
        return jsonify({'error': 'Folder and date required'}), 400
    if not os.path.isdir(folder):
        return jsonify({'error': f'Folder not found: {folder}'}), 400
    found  = _find_report_input_files(folder, date_str)
    result = {key: os.path.basename(found[key]) if key in found else None
              for key in _REPORT_FILE_PATTERNS}
    # Auto-detect master template in the same folder
    tmpl_name = None
    for f in os.listdir(folder):
        if f.upper().startswith('DAILY_CPAN_REPORTS_MASTER') and f.lower().endswith('.xlsx'):
            _custom_template_path = os.path.join(folder, f)
            _save_config({'template_path': _custom_template_path})
            tmpl_name = f
            break
    return jsonify({'files': result, 'template': tmpl_name})

@app.route('/report/start', methods=['POST'])
def report_start():
    data      = request.json or {}
    folder    = (data.get('folder')   or '').strip()
    temp_dir  = (data.get('temp_dir') or '').strip()   # from LAN upload flow
    date_str  = (data.get('date')     or '').strip()
    out_name  = (data.get('output')   or '').strip()

    # LAN mode: temp_dir replaces folder
    if temp_dir and os.path.isdir(temp_dir):
        folder = temp_dir

    if not folder or not date_str or not out_name:
        return jsonify({'error': 'Missing required fields'}), 400
    if not os.path.isdir(folder):
        return jsonify({'error': f'Folder not found: {folder}'}), 400

    found = _find_report_input_files(folder, date_str)
    if not found:
        return jsonify({'error': 'No input files found for the given date in that folder'}), 400
    if not out_name.lower().endswith('.xlsx'):
        out_name += '.xlsx'

    # Output always goes to server-side reports_output; LAN users download via /report/download
    default_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'reports_output')
    reports_dir = _custom_output_dir if _custom_output_dir else default_dir
    os.makedirs(reports_dir, exist_ok=True)
    output_path = os.path.join(reports_dir, out_name)

    try:
        parts = date_str.split('-')
        report_date = datetime(int(parts[2]), int(parts[1]), int(parts[0]))
    except Exception:
        return jsonify({'error': 'Invalid date format, expected DD-MM-YYYY'}), 400

    # Snapshot df_raw NOW in the main thread before handing off to the background
    # thread — prevents any pandas operation inside report_gen from touching the
    # live DataFrame that powers the path finder.
    links_snap = df_raw.copy() if df_raw is not None else None

    job_id = str(uuid.uuid4())[:8]
    _report_jobs[job_id] = {
        'status': 'running', 'logs': [], 'output_path': None,
        'pdf_path': None, 'error': None,
        '_temp_dir': temp_dir or None,
    }
    threading.Thread(target=_run_report_job,
                     args=(job_id, found, output_path, report_date, links_snap),
                     daemon=True).start()
    return jsonify({'job_id': job_id})

@app.route('/report/status/<job_id>')
def report_status(job_id):
    job = _report_jobs.get(job_id)
    if not job:
        return jsonify({'error': 'Job not found'}), 404
    return jsonify({'status': job['status'], 'logs': job['logs'], 'error': job['error'],
                    'has_pdf': bool(job.get('pdf_path'))})

@app.route('/report/download/<job_id>')
def report_download(job_id):
    job = _report_jobs.get(job_id)
    if not job or job['status'] != 'done' or not job['output_path']:
        return jsonify({'error': 'Not ready'}), 404
    path = job['output_path']
    if not os.path.exists(path):
        return jsonify({'error': 'Output file not found on server'}), 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

@app.route('/report/download_pdf/<job_id>')
def report_download_pdf(job_id):
    job = _report_jobs.get(job_id)
    if not job or job['status'] != 'done' or not job.get('pdf_path'):
        return jsonify({'error': 'PDFs not ready'}), 404
    path = job['pdf_path']
    if not os.path.exists(path):
        return jsonify({'error': 'ZIP file not found on server'}), 404
    return send_file(path, as_attachment=True,
                     download_name=os.path.basename(path),
                     mimetype='application/zip')

# ──────────────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True, use_reloader=False)
