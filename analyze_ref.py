"""Analyze reference master sheets to find exact filter logic."""
import openpyxl, pandas as pd
from datetime import datetime

BASE = r"C:\Project_AI\network-path-finder-project"

# Use 07.05.26 as representative sample
FOLDER   = f"{BASE}\\07.05.26"
MASTER   = f"{FOLDER}\\DAILY_CPAN_REPORTS_MASTER_07.05.26.xlsx"
CARD_IN  = f"{FOLDER}\\CARD OFFLINE 07.05.2026.xlsx"
DEV_IN   = f"{FOLDER}\\DEVICE OFFLINE 07.05.2026.xlsx"
FAN_IN   = f"{FOLDER}\\FAN FAILURE 07.05.2026.xlsx"
DL_IN    = f"{FOLDER}\\DL FAIL REPORT 07.05.2026.xlsx"

REPORT_DATE = datetime(2026, 5, 7)

print("=" * 65)
print("REFERENCE MASTER: 07.05.26")
print("=" * 65)

wb = openpyxl.load_workbook(MASTER, data_only=True)

def read_ref_sheet(wb, sh_name, hdr_row):
    sh = wb[sh_name]
    rows = []
    for r in range(hdr_row, sh.max_row + 1):
        row_vals = [c.value for c in sh[r]]
        rows.append(row_vals)
    df = pd.DataFrame(rows[1:], columns=rows[0])
    df.columns = [str(c).strip().replace('\n',' ') if c else f'c{i}'
                  for i, c in enumerate(df.columns)]
    return df

# ── CARD-OFF-R reference ────────────────────────────────────────────────────
print("\n--- CARD-OFF-R (reference) ---")
ref_card = read_ref_sheet(wb, 'CARD-OFF-R', 4)
print(f"Columns: {list(ref_card.columns)}")
print(f"Total rows: {len(ref_card)}")
print(f"Circle distribution:")
circ_col = next((c for c in ref_card.columns if 'circle' in c.lower() or c=='Circle'), None)
if circ_col:
    print(ref_card[circ_col].value_counts(dropna=False).to_string())
ssa_col = next((c for c in ref_card.columns if 'ssa' in c.lower() or c=='SSA'), None)
if ssa_col:
    print(f"SSA distribution (top 20):")
    print(ref_card[ssa_col].value_counts(dropna=False).head(20).to_string())

# ── Check input card vs ref ─────────────────────────────────────────────────
print("\n--- CARD OFFLINE input 07.05.26 ---")
card_in = pd.read_excel(CARD_IN, dtype=str)
card_in.columns = [c.strip() for c in card_in.columns]
print(f"Total rows: {len(card_in)}")
print(f"Circle distribution:")
print(card_in['CIRCLE'].value_counts().to_string())
print(f"WTR SSA distribution:")
print(card_in[card_in['CIRCLE']=='WTR']['SSA'].value_counts().to_string())

# Which input IPs are in reference?
ref_ips = set(ref_card.iloc[:,1].astype(str).str.strip())
in_ips  = set(card_in['IP'].astype(str).str.strip())
overlap = ref_ips & in_ips
print(f"\nRef unique IPs: {len(ref_ips)}")
print(f"Input unique IPs: {len(in_ips)}")
print(f"Overlap (ref IPs found in input): {len(overlap)}")

# ── FAN-FAIL-R reference ────────────────────────────────────────────────────
print("\n--- FAN-FAIL-R (reference) ---")
ref_fan = read_ref_sheet(wb, 'FAN-FAIL-R', 4)
print(f"Total rows: {len(ref_fan)}, cols: {list(ref_fan.columns)[:8]}")
circ_col = next((c for c in ref_fan.columns if 'circle' in c.lower()), None)
if circ_col:
    print(f"Circle dist: {ref_fan[circ_col].value_counts(dropna=False).to_string()}")
ssa_col2 = next((c for c in ref_fan.columns if 'region' in c.lower() or c=='REGION'), None)
if ssa_col2:
    print(f"Region/SSA dist (top 20): {ref_fan[ssa_col2].value_counts(dropna=False).head(20).to_string()}")

print("\n--- FAN FAILURE input 07.05.26 ---")
fan_in = pd.read_excel(FAN_IN, dtype=str)
fan_in.columns = [c.strip() for c in fan_in.columns]
print(f"Total rows: {len(fan_in)}")
print(f"Circle dist:")
print(fan_in['Circle'].value_counts().to_string())
print(f"WTR SSA dist:")
print(fan_in[fan_in['Circle']=='WTR']['SSA'].value_counts().to_string())

# ── DEVICE-OFF-R reference ─────────────────────────────────────────────────
print("\n--- DEVICE-OFF-R (reference) ---")
ref_dev = read_ref_sheet(wb, 'DEVICE-OFF-R', 4)
print(f"Total rows: {len(ref_dev)}, cols: {list(ref_dev.columns)[:8]}")
circ_col = next((c for c in ref_dev.columns if 'circle' in c.lower()), None)
if circ_col:
    print(f"Circle dist: {ref_dev[circ_col].value_counts(dropna=False).head(10).to_string()}")

print("\n--- DEVICE OFFLINE input 07.05.26 ---")
dev_in = pd.read_excel(DEV_IN, dtype=str)
dev_in.columns = [c.strip() for c in dev_in.columns]
print(f"Total rows: {len(dev_in)}")
print(f"Circle dist:")
print(dev_in['CIRCLE'].value_counts().to_string())

# Are all ref IPs in the input?
ref_dev_ips = set(ref_dev.iloc[:,1].astype(str).str.strip())
dev_in_ips  = set(dev_in['IP'].astype(str).str.strip())
print(f"Ref unique IPs: {len(ref_dev_ips)}, Input unique IPs: {len(dev_in_ips)}")
print(f"Ref IPs in input: {len(ref_dev_ips & dev_in_ips)}")
print(f"Ref IPs NOT in input: {len(ref_dev_ips - dev_in_ips)}")

# ── DL-FAIL-R reference ────────────────────────────────────────────────────
print("\n--- DL-FAIL-R (reference) ---")
ref_dl = read_ref_sheet(wb, 'DL-FAIL-R', 5)
print(f"Total rows: {len(ref_dl)}, cols: {list(ref_dl.columns)[:8]}")
circ_col = next((c for c in ref_dl.columns if 'circle' in c.lower()), None)
if circ_col:
    print(f"Circle dist: {ref_dl[circ_col].value_counts(dropna=False).head(10).to_string()}")

print("\n--- DL FAIL input 07.05.26 ---")
dl_in = pd.read_excel(DL_IN, dtype=str)
dl_in.columns = [c.strip() for c in dl_in.columns]
print(f"Total rows: {len(dl_in)}, cols: {list(dl_in.columns)}")
circ_col2 = next((c for c in dl_in.columns if 'circle' in c.lower()), None)
if circ_col2:
    print(f"Circle dist in input:")
    print(dl_in[circ_col2].value_counts().to_string())
