"""Investigate filter logic to match reference row counts."""
import pandas as pd
from datetime import datetime

REPORT_DATE = datetime(2026, 4, 2)

def parse_ct(val):
    if not val or str(val).strip() in ('nan', 'None', ''): return None
    s = str(val).strip()
    for fmt in ('%Y/%m/%d,%H:%M:%S', '%Y/%m/%d', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
        try:
            dt = datetime.strptime(s, fmt)
            return None if dt.year < 2010 else dt
        except ValueError:
            continue
    return None

# ── CARD-OFF-R analysis ─────────────────────────────────────────────────────
print("=== CARD-OFF-R filter analysis ===")
card = pd.read_excel(r"C:\Project_AI\cpan_report\CARD OFFLINE 02-04-2026.xlsx", dtype=str)
card.columns = [c.strip() for c in card.columns]
card['_dt'] = card['Create Time'].str.strip().apply(parse_ct)
card_valid = card.dropna(subset=['_dt'])
print(f"Input total: {len(card)}")
print(f"After date filter (year>=2010): {len(card_valid)}")
print(f"SSA distribution for WTR rows:")
wtr_card = card_valid[card_valid['CIRCLE'].str.upper() == 'WTR']
print(wtr_card['SSA'].value_counts().to_string())
print(f"\nAfter GJ only: {len(card_valid[card_valid['CIRCLE'].str.upper()=='GJ'])}")
print(f"After GJ+WTR all: {len(card_valid[card_valid['CIRCLE'].str.upper().isin({'GJ','WTR'})])}")
# WTR SSA AM + RJ only
wtr_am_rj = wtr_card[wtr_card['SSA'].isin(['AM', 'RJ'])]
gj_rows = card_valid[card_valid['CIRCLE'].str.upper() == 'GJ']
combo = len(gj_rows) + len(wtr_am_rj)
print(f"After GJ + WTR(AM+RJ only): {combo}  [REF=200]")

# ── FAN-FAIL-R analysis ──────────────────────────────────────────────────────
print()
print("=== FAN-FAIL-R filter analysis ===")
fan = pd.read_excel(r"C:\Project_AI\cpan_report\FAN FAILURE 02-04-2026.xlsx", dtype=str)
fan.columns = [c.strip() for c in fan.columns]
fan['_dt'] = fan['Create Time'].str.strip().apply(parse_ct)
fan_valid = fan.dropna(subset=['_dt'])
print(f"Input total: {len(fan)}")
print(f"After date filter: {len(fan_valid)}")
wtr_fan = fan_valid[fan_valid['Circle'].str.upper() == 'WTR']
print(f"WTR SSA distribution:")
print(wtr_fan['SSA'].value_counts().to_string())
gj_fan = fan_valid[fan_valid['Circle'].str.upper() == 'GJ']
wtr_am_rj_fan = wtr_fan[wtr_fan['SSA'].isin(['AM', 'RJ'])]
print(f"GJ only: {len(gj_fan)}")
print(f"WTR(AM+RJ) only: {len(wtr_am_rj_fan)}")
print(f"GJ + WTR(AM+RJ): {len(gj_fan)+len(wtr_am_rj_fan)}  [REF=151]")

# ── DEVICE-OFF-R analysis ────────────────────────────────────────────────────
print()
print("=== DEVICE-OFF-R analysis ===")
dev = pd.read_excel(r"C:\Project_AI\cpan_report\DEVICE OFFLINE 02-04-2026.xlsx", dtype=str)
dev.columns = [c.strip() for c in dev.columns]
dev['_dt'] = dev['Create Time'].str.strip().apply(parse_ct)
dev_valid = dev.dropna(subset=['_dt'])
print(f"Input total: {len(dev)}, after date filter: {len(dev_valid)}")
print(f"CIRCLE distribution in input:")
print(dev_valid['CIRCLE'].value_counts().to_string())
print(f"\nGJ only: {len(dev_valid[dev_valid['CIRCLE']=='GJ'])}  [REF=150]")
print("NOTE: REF has 150 rows from 26 unique GJ IPs but input only has 26 GJ rows")
print("=> The reference DEVICE-OFF-R uses DIFFERENT/EXPANDED source data")
print("=> Check if CARD-OFF GJ rows = DEVICE-OFF-R rows in reference...")
# Count unique IPs that appear in CARD-OFF with GJ circle
card_gj_ips = set(card_valid[card_valid['CIRCLE']=='GJ']['IP'].astype(str).str.strip())
dev_gj_ips  = set(dev_valid[dev_valid['CIRCLE']=='GJ']['IP'].astype(str).str.strip())
print(f"CARD-OFF GJ unique IPs: {len(card_gj_ips)}")
print(f"DEVICE-OFF GJ unique IPs: {len(dev_gj_ips)}")
overlap = card_gj_ips & dev_gj_ips
print(f"IPs in both: {len(overlap)}")
# How many CARD-OFF GJ rows for these IPs?
card_gj_rows = card_valid[(card_valid['CIRCLE']=='GJ') & (card_valid['IP'].isin(overlap))]
print(f"CARD-OFF rows for GJ IPs that are also in DEVICE-OFF: {len(card_gj_rows)}")

# ── DL-FAIL-R analysis ──────────────────────────────────────────────────────
print()
print("=== DL-FAIL-R analysis ===")
dl = pd.read_excel(r"C:\Project_AI\cpan_report\CPAN DL FAIL REPORT  02-04-2026.xlsx", dtype=str)
dl.columns = [c.strip() for c in dl.columns]
dl = dl.rename(columns={' Circle': 'Circle', 'region': 'Region'})
dl['_dt'] = dl['Create Time'].str.strip().apply(parse_ct)
dl_valid = dl.dropna(subset=['_dt'])
print(f"Input total: {len(dl)}")
print(f"After date filter (year>=2010): {len(dl_valid)}  [REF=309]")
print(f"Circle distribution after date filter:")
print(dl_valid['Circle'].value_counts().to_string())
if 'Direction' in dl_valid.columns:
    print(f"\nDirection distribution:")
    print(dl_valid['Direction'].value_counts().to_string())
    az_only = dl_valid[dl_valid['Direction'].str.upper().str.strip() == 'AZ']
    print(f"AZ only: {len(az_only)}")
print(f"\nGJ+WTR: {len(dl_valid[dl_valid['Circle'].str.upper().isin({'GJ','WTR'})])}  [REF=309]")
