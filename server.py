from flask import Flask, render_template, request, jsonify
import pandas as pd
import networkx as nx
import os
import io
from datetime import datetime

app = Flask(__name__)

# ── Global state ───────────────────────────────────────────────────────────────
df_raw = None
df_down = None
df_10g = None
G = nx.Graph()
ip_name_map = {}   # IP → human-readable node name
last_loaded = None
loaded_filename = "links.csv"

def extract_node_name(endpoint):
    """Extract human-readable site name from endpoint string like IP_SiteName_..."""
    parts = str(endpoint).strip().split('_')
    return parts[1].strip() if len(parts) > 1 else parts[0].strip()

def clean_and_rebuild(df):
    """Clean dataframe and rebuild all derived data + graph."""
    global df_raw, df_down, df_10g, G, ip_name_map

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

    # Rebuild graph
    data = df[['A End', 'Z End', 'Bandwidth', 'CIR Utilization Ratio(%)']].copy()
    data['A IP'] = data['A End'].astype(str).str.strip().str.split('_').str[0]
    data['Z IP'] = data['Z End'].astype(str).str.strip().str.split('_').str[0]

    # Build IP → node name mapping
    ip_name_map = {}
    for _, row in data.iterrows():
        a_ip = row['A IP']
        z_ip = row['Z IP']
        if a_ip not in ip_name_map:
            ip_name_map[a_ip] = extract_node_name(row['A End'])
        if z_ip not in ip_name_map:
            ip_name_map[z_ip] = extract_node_name(row['Z End'])

    G = nx.Graph()
    for _, row in data.iterrows():
        G.add_edge(row['A IP'], row['Z IP'],
                   bandwidth=row['Bandwidth'],
                   cir=row['CIR Utilization Ratio(%)'])

# ── Initial load from disk ─────────────────────────────────────────────────────
def load_from_disk():
    global last_loaded, loaded_filename
    path = "links.csv"
    df = pd.read_csv(path)
    clean_and_rebuild(df)
    last_loaded   = datetime.fromtimestamp(os.path.getmtime(path)).strftime('%d-%b-%Y %H:%M')
    loaded_filename = os.path.basename(path)

load_from_disk()

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

    # Sort candidates: fewest hops → max 10G links → lowest avg CIR
    def path_sort_key(p):
        count_10g = sum(1 for i in range(len(p)-1)
                        if G[p[i]][p[i+1]].get('bandwidth') == '10GE')
        avg_cir   = (sum(G[p[i]][p[i+1]].get('cir', 0) for i in range(len(p)-1))
                     / (len(p) - 1)) if len(p) > 1 else 0
        return (len(p), -count_10g, avg_cir)

    az_paths = sorted(
        nx.all_simple_paths(G, source, target, cutoff=10),
        key=path_sort_key
    )[:10]
    za_paths = [list(reversed(p)) for p in az_paths]

    az_results = build_results(az_paths, 'A→Z')
    za_results = build_results(za_paths, 'Z→A')

    sort_key = lambda x: (x['hop_count'], -x['count_10g'], x['avg_cir'])
    az_results.sort(key=sort_key)
    za_results.sort(key=sort_key)

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

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
