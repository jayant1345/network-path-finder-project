"""Extract NODE-LIST from master template and verify filter logic."""
import openpyxl, pandas as pd, os

MASTER   = r"C:\Project_AI\cpan_report\DAILY_CPAN_REPORTS_MASTER_02.04.26.xlsx"
CARD_IN  = r"C:\Project_AI\cpan_report\CARD OFFLINE 02-04-2026.xlsx"
REF      = r"C:\Project_AI\cpan_report\DAILY_CPAN_REPORTS_02.04.26.xlsx"
OUT_CSV  = r"C:\Project_AI\network-path-finder-project\node_list.csv"

# ── Extract NODE-LIST ──────────────────────────────────────────────────────
wb = openpyxl.load_workbook(MASTER, data_only=True, read_only=True)
sh = wb['NODE-LIST']

rows = []
headers = None
for r in sh.iter_rows(values_only=True):
    vals = [v for v in r]
    if not any(v for v in vals if v is not None):
        continue
    if headers is None:
        headers = [str(v).strip() if v else f'col{i}' for i, v in enumerate(vals)]
        continue
    rows.append(vals)

node_df = pd.DataFrame(rows, columns=headers)
print(f"NODE-LIST total rows: {len(node_df)}")
print(f"Columns: {list(node_df.columns)}")

# Find IP column
ip_col = next((c for c in node_df.columns if 'ip' in c.lower() or 'IP' in c), None)
# 'Node IP' column should be present
for c in node_df.columns:
    print(f"  Sample col '{c}': {node_df[c].dropna().head(3).tolist()}")

print(f"\nNode IP column: '{ip_col}'")
print(f"Unique IPs in NODE-LIST: {node_df[ip_col].dropna().nunique() if ip_col else 'N/A'}")

# Save a clean version
if ip_col:
    clean = node_df[[ip_col]].copy()
    clean.columns = ['Node IP']
    clean['Node IP'] = clean['Node IP'].astype(str).str.strip()
    clean = clean[clean['Node IP'].str.match(r'^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$')]
    clean = clean.drop_duplicates()
    print(f"Valid unique IPs after cleanup: {len(clean)}")
    clean.to_csv(OUT_CSV, index=False)
    print(f"Saved to: {OUT_CSV}")

    managed_ips = set(clean['Node IP'].tolist())

    # ── Verify: does NODE-LIST filter explain CARD-OFF-R row count? ──────────
    print()
    print("=== Verification: CARD-OFF-R ===")
    card = pd.read_excel(CARD_IN, dtype=str)
    card.columns = [c.strip() for c in card.columns]

    def parse_ct(val):
        if not val or str(val).strip() in ('nan', 'None', ''): return None
        s = str(val).strip()
        import datetime
        for fmt in ('%Y/%m/%d,%H:%M:%S', '%Y/%m/%d', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
            try:
                dt = datetime.datetime.strptime(s, fmt)
                return None if dt.year < 2010 else dt
            except ValueError:
                continue
        return None

    card['_dt'] = card['Create Time'].str.strip().apply(parse_ct)
    card_valid = card.dropna(subset=['_dt'])
    print(f"Input after date filter: {len(card_valid)}")

    card_managed = card_valid[card_valid['IP'].astype(str).str.strip().isin(managed_ips)]
    print(f"Input after NODE-LIST filter: {len(card_managed)}  [REF=200]")
    print(f"Circles in filtered set: {card_managed['CIRCLE'].value_counts().to_string()}")
    print(f"SSAs in filtered set: {card_managed['SSA'].value_counts().head(15).to_string()}")

    # ── Check for DL-FAIL ──────────────────────────────────────────────────
    print()
    print("=== Verification: DL-FAIL-R ===")
    dl = pd.read_excel(r"C:\Project_AI\cpan_report\CPAN DL FAIL REPORT  02-04-2026.xlsx", dtype=str)
    dl.columns = [c.strip() for c in dl.columns]
    dl = dl.rename(columns={' Circle': 'Circle', 'region': 'Region'})
    dl['_dt'] = dl['Create Time'].str.strip().apply(parse_ct)
    dl_valid = dl.dropna(subset=['_dt'])

    # Filter: A-end OR Z-end IP in managed_ips
    dl['a_in'] = dl['IP A END'].astype(str).str.strip().isin(managed_ips)
    dl['z_in'] = dl['IP Z END'].astype(str).str.strip().isin(managed_ips)
    dl_managed = dl_valid[dl['a_in'] | dl['z_in']]
    print(f"Input total: {len(dl_valid)}")
    print(f"After NODE-LIST filter (A or Z in managed): {len(dl_managed)}  [REF=309]")
    print(f"Circle distribution: {dl_managed['Circle'].value_counts().to_string()}")

    # Alarm status check
    if 'Alarm Status' in dl_valid.columns:
        print(f"\nAlarm Status in input: {dl_valid['Alarm Status'].value_counts().to_string()}")
    if 'Alarm Type' in dl_valid.columns:
        print(f"Alarm Type in input: {dl_valid['Alarm Type'].value_counts().to_string()}")
