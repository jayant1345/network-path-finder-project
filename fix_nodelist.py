"""
Extract BOTH old and new IPs from NODE-LIST (many nodes changed IP).
The Additional Remarks column has 'new IP-xx.xx.xx.xx' for changed nodes.
"""
import openpyxl, pandas as pd, re

MASTER = r"C:\Project_AI\network-path-finder-project\07.05.26\DAILY_CPAN_REPORTS_MASTER_07.05.26.xlsx"
OUT    = r"C:\Project_AI\network-path-finder-project\node_list.csv"

IP_RE = re.compile(r'\b(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\b')

print("Loading NODE-LIST with full IP data...")
wb = openpyxl.load_workbook(MASTER, data_only=True)
sh = wb['NODE-LIST']

all_ips = set()
hdr = None
changed_count = 0

for r in range(1, sh.max_row + 1):
    row = [c.value for c in sh[r]]
    if not any(v for v in row if v is not None):
        continue
    if hdr is None:
        hdr = [str(v).strip().replace('\n', ' ') if v else f'c{i}'
               for i, v in enumerate(row)]
        continue

    # Find column indices
    ip_idx    = next((i for i, h in enumerate(hdr) if h == 'Node IP'), None)
    name_idx  = next((i for i, h in enumerate(hdr) if h == 'Node Name'), None)
    addl_idx  = next((i for i, h in enumerate(hdr) if 'additiona' in h.lower() or 'additional' in h.lower()), None)

    # Old IP from Node IP column
    if ip_idx is not None and ip_idx < len(row) and row[ip_idx]:
        m = IP_RE.search(str(row[ip_idx]))
        if m:
            all_ips.add(m.group(1))

    # Old IP also from Node Name column (format: '10.121.15.132_NAME_...')
    if name_idx is not None and name_idx < len(row) and row[name_idx]:
        m = IP_RE.match(str(row[name_idx]).strip())
        if m:
            all_ips.add(m.group(1))

    # NEW IP from Additional Remarks column ('new IP-10.121.96.35')
    if addl_idx is not None and addl_idx < len(row) and row[addl_idx]:
        val = str(row[addl_idx])
        if 'new' in val.lower() and 'ip' in val.lower():
            for m in IP_RE.finditer(val):
                all_ips.add(m.group(1))
                changed_count += 1

print(f"Total unique IPs (old + new): {len(all_ips)}")
print(f"New/changed IPs added: {changed_count}")

# Save
df = pd.DataFrame(sorted(all_ips), columns=['Node IP'])
df.to_csv(OUT, index=False)
print(f"Saved to {OUT}")

# Verify against 07.05.26 CARD OFFLINE
print("\n=== Verification ===")
card = pd.read_excel(
    r"C:\Project_AI\network-path-finder-project\07.05.26\CARD OFFLINE 07.05.2026.xlsx",
    dtype=str)
card.columns = [c.strip() for c in card.columns]
after = card[card['IP'].astype(str).str.strip().isin(all_ips)]
print(f"CARD-OFF 07.05.26: {len(card)} rows -> after filter: {len(after)} [REF=200]")
print(f"  Circle dist: {after['CIRCLE'].value_counts().to_string()}")

# FAN FAIL
fan = pd.read_excel(
    r"C:\Project_AI\network-path-finder-project\07.05.26\FAN FAILURE 07.05.2026.xlsx",
    dtype=str)
fan.columns = [c.strip() for c in fan.columns]
fan_after = fan[fan['IP'].astype(str).str.strip().isin(all_ips)]
print(f"\nFAN-FAIL 07.05.26: {len(fan)} rows -> after filter: {len(fan_after)} [REF=150]")
print(f"  Circle dist: {fan_after['Circle'].value_counts().to_string()}")

# DEVICE OFF
dev = pd.read_excel(
    r"C:\Project_AI\network-path-finder-project\07.05.26\DEVICE OFFLINE 07.05.2026.xlsx",
    dtype=str)
dev.columns = [c.strip() for c in dev.columns]
dev_after = dev[dev['IP'].astype(str).str.strip().isin(all_ips)]
print(f"\nDEVICE-OFF 07.05.26: {len(dev)} rows -> after filter: {len(dev_after)} [REF=150]")
print(f"  Circle dist: {dev_after['CIRCLE'].value_counts().to_string()}")

# DL FAIL
dl = pd.read_excel(
    r"C:\Project_AI\network-path-finder-project\07.05.26\DL FAIL REPORT 07.05.2026.xlsx",
    dtype=str)
dl.columns = [c.strip() for c in dl.columns]
dl_ems = pd.read_csv(r"C:\Project_AI\network-path-finder-project\dl_ems_ips.csv", dtype=str)
dl_ems_ips = set(dl_ems['Node IP'].dropna().str.strip())
a_in = dl['IP A END'].astype(str).str.strip().isin(dl_ems_ips)
z_in = dl['IP Z END'].astype(str).str.strip().isin(dl_ems_ips)
dl_after = dl[a_in | z_in]
print(f"\nDL-FAIL 07.05.26: {len(dl)} rows -> after DL-EMS filter: {len(dl_after)} [REF=300]")
