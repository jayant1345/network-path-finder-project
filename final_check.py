"""Final check: what WTR SSAs are in the inputs vs reference."""
import pandas as pd, openpyxl
from datetime import datetime

BASE = r"C:\Project_AI\network-path-finder-project\07.05.26"

# Card
card = pd.read_excel(f"{BASE}\\CARD OFFLINE 07.05.2026.xlsx", dtype=str)
card.columns = [c.strip() for c in card.columns]
print("CARD OFFLINE WTR rows SSA/REGION distribution:")
wtr_card = card[card['CIRCLE']=='WTR']
ssa_col = 'SSA' if 'SSA' in card.columns else ('REGION' if 'REGION' in card.columns else None)
if ssa_col:
    print(wtr_card[ssa_col].value_counts().to_string())
print(f"Total WTR: {len(wtr_card)}, GJ: {len(card[card['CIRCLE']=='GJ'])}")
print(f"After GJ + WTR(AM+RJ) filter: {len(card[card['CIRCLE']=='GJ'])} + {len(wtr_card[wtr_card[ssa_col].isin(['AM','RJ'])])}")

print()
# Fan
fan = pd.read_excel(f"{BASE}\\FAN FAILURE 07.05.2026.xlsx", dtype=str)
fan.columns = [c.strip() for c in fan.columns]
wtr_fan = fan[fan['Circle']=='WTR']
print("FAN FAILURE WTR rows SSA distribution:")
print(wtr_fan['SSA'].value_counts().to_string())
print(f"Total WTR: {len(wtr_fan)}, GJ: {len(fan[fan['Circle']=='GJ'])}")
print(f"After GJ + WTR(AM+RJ): {len(fan[fan['Circle']=='GJ'])} + {len(wtr_fan[wtr_fan['SSA'].isin(['AM','RJ'])])}")

print()
# Device
dev = pd.read_excel(f"{BASE}\\DEVICE OFFLINE 07.05.2026.xlsx", dtype=str)
dev.columns = [c.strip() for c in dev.columns]
print("DEVICE OFFLINE all circles:")
print(dev['CIRCLE'].value_counts().to_string())

print()
# Check master DL-FAIL input - how many GJ rows?
print("=== Reading DL-FAIL from master CARD-OFF sheet ===")
wb = openpyxl.load_workbook(f"{BASE}\\DAILY_CPAN_REPORTS_MASTER_07.05.26.xlsx", data_only=True)
sh = wb['CARD-OFF']
# Count rows where Circle formula col (col index 3) is not empty/null
total_rows = 0
gj_count = 0
wtr_am_rj = 0
for r in range(11, sh.max_row + 1):
    row = [c.value for c in sh[r]]
    if not any(v for v in row if v is not None):
        break
    total_rows += 1
    circle_code = row[3]  # formula column: 1=GJ, 2=WTR, blank=not managed
    ssa_formula = str(row[2]).strip() if row[2] else ''
    if circle_code == 1:
        gj_count += 1
    elif circle_code == 2 and ssa_formula in ('AM', 'RJ'):
        wtr_am_rj += 1
print(f"Master CARD-OFF total data rows: {total_rows}")
print(f"  Circle code=1 (GJ): {gj_count}")
print(f"  Circle code=2, SSA=AM/RJ (WTR-GJ): {wtr_am_rj}")
print(f"  Total GJ+WTR-GJ: {gj_count + wtr_am_rj}  [REF=200]")

# Check what circle codes appear
codes = {}
for r in range(11, sh.max_row + 1):
    row = [c.value for c in sh[r]]
    if not any(v for v in row if v is not None):
        break
    code = row[3]
    codes[code] = codes.get(code, 0) + 1
print(f"\nCircle code distribution in master CARD-OFF: {codes}")
