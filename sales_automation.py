import pandas as pd
import openpyxl

# ============================================================
# SETTINGS — update these every week
# ============================================================

ORDER_FILE = r"C:\Users\samue\Downloads\ORDER_RECEIVE_2026.xlsx"
SALES_FILE = r"C:\Users\samue\Downloads\Yearly_Sales_Report_2026.xlsx"

# "overwrite" = replace original file
# "new"       = save a new copy, original stays safe
SAVE_MODE = "new"

# Month sheet to read from Order Receive
CURRENT_MONTH = "APR"

# Which column number is this month in Sales Report?
# C=Jan(3), D=Feb(4), E=Mar(5), F=Apr(6)
# In openpyxl columns use normal counting: A=1, B=2, C=3...
MONTH_COLUMN = 6   # F = April

# ============================================================
# NAME MAP
# Left side  = name in Order Receive, will be auto-uppercased
# Right side = name EXACTLY as written in Sales Report
# ============================================================

NAME_MAP = {
    # ALAN
    "ABEDON":           "Abedon",
    "G.TANJUNG":        "G.Tanjung",
    "LEEPANG":          "Leepang",
    "MORISEM":          "Morisem",
    "SYARIMO":          "Syarimo",
    "TUNG HUP":         "Tung Hup",        # fixed!
    "UNICO DESA":       "Unico Desa",

    # ALWIN
    "ATLANTICA":        "Atlantica",
    "BEAUFORT":         "Beaufort",
    "DESA KIM LOONG":   "Desa Kim Loong",  # fixed!
    "G.EDIBLE OIL":     "G.Edible Oil",    # fixed!
    "IOI EDIBLE/BIO":   "IOI Edible/Bio",
    "MALSA":            "Malsa",
    "MERIDIAN":         "Meridian",
    "MELALAP":          "Melalap",
    "PITAS":            "Pitas",
    "SEO":              "SEO",
    "SOOK/DALIT":       "Sook/Dalit",
    "TUARAN CRUMB":     "Tuaran Crumb",
    "YUN FOOK":         "Yun Fook",
    "BORNION (M)":      "Bornion (M)",
    "SABANG 2 (M)":     "Sabang 2 (M)",
    "DESA TALISAI (M)": "Desa Talisai (M)",

    # CHAU
    "AUMKAR":           "Aumkar",
    "BUKITMAS":         "Bukitmas",
    "BATURONG":         "Baturong",
    "BINUANG":          "Binuang",
    "ECO OIL":          "Eco Oil",
    "JEROCO 1":         "Jeroco 1",
    "JEROCO 2":         "Jeroco 2",
    "KLK PREMIER":      "KLK Premier",
    "L.PERMAI":         "L.Permai",
    "SABAHMAS":         "Sabahmas",
    "SAPANG":           "Sapang",
    "TALIWAS":          "Taliwas",
    "TOMANGGONG":       "Tomanggong",
    "YUWANG":           "Yuwang",
    "LUNGMANIS":        "Lungmanis",
    "RIMMER":           "Rimmer",
    "PINANG":           "Pinang",
    "SABAH 2":          "Sabah 2",
    "MEROTAI":          "Merotai",
    "REX":              "Rex",
    "SEBATIK":          "Sebatik",
    "BALUNG":           "Balung",
    "SG.BURUNG":        "Sg.Burung",
    "SEGARIA (M)":      "Segaria (M)",

    # DANNY
    "BERKAT":           "Berkat",
    "CASH HORSE":       "Cash Horse",
    "G.INDAH":          "G.Indah",
    "K.L SABAH":        "K.L Sabah",
    "MAYVIN":           "Mayvin",
    "PROLIFIC":         "Prolific",
    "RIBUBONUS":        "Ribubonus",
    "REKA HALUS":       "Reka Halus",
    "T.SIANG":          "T.Siang",
    "VEETAR":           "Veetar",
    "SEGAMAHA (M)":     "Segamaha (M)",

    # LEONG
    "BATU PUTIH":       "Batu Putih",
    "GLOBAL":           "Global",
    "G.MEWAH":          "G.Mewah",
    "G.TRUSHIDUP":      "G.Trushidup",
    "G.JAMBONGAN":      "G.Jambongan",
    "LCH":              "LCH",
    "L.SABAH":          "L.Sabah",
    "LCH BAY (QL2)":    "LCH Bay (QL2)",
    "SDK BAY":          "Sdk Bay",
    "SEPAGAYA":         "Sepagaya",
    "SAKILAN":          "Sakilan",

    # RYAN
    "BELL":             "Bell",
    "BORNION":          "Bornion",
    "FORTUNA":          "Fortuna",
    "L.MILLS":          "L.Mills",
    "SABANG (2)":       "Sabang (2)",
    "SG.RUKU":          "Sg.Ruku",
    "SUKAU":            "Sukau",
    "T.PANJANG":        "T.Panjang",
    "NAK (M)":          "Nak (M)",

    # TOONG
    "CAHAYA":           "Cahaya",
    "DESA TALISAI":     "Desa Talisai",
    "DUMPAS":           "Dumpas",
    "G.SABAPALM":       "G.Sabapalm",
    "PAMOL":            "Pamol",
    "SABANG 1":         "Sabang 1",
    "SRI KAMUSAN":      "Sri Kamusan",
    "SAPI":             "Sapi",
    "TERUSAN":          "Terusan",
    "MAJUMAS":          "Majumas",
    "TAWAI (M)":        "Tawai (M)",
}

# These are the MAIN sales table rows only
# We use this to avoid writing into the Job table or Budget table
# This is the list of row numbers where each salesperson's
# MAIN table starts and ends in the Sales Report
# Based on what I found in your actual file:

MAIN_TABLE_ROWS = {
    "Abedon":           range(4,  15),
    "G.Tanjung":        range(4,  15),
    "Leepang":          range(4,  15),
    "Morisem":          range(4,  15),
    "Syarimo":          range(4,  15),
    "Tung Hup":         range(4,  15),
    "Unico Desa":       range(4,  15),
    "Nak (S)":          range(4,  15),
    "Segamaha (S)":     range(4,  15),
    "Segaria (S)":      range(4,  15),
    "Tawai (S)":        range(4,  15),
    "Atlantica":        range(58, 74),
    "Beaufort":         range(58, 74),
    "Desa Kim Loong":   range(58, 74),
    "G.Edible Oil":     range(58, 74),
    "IOI Edible/Bio":   range(58, 74),
    "Malsa":            range(58, 74),
    "Meridian":         range(58, 74),
    "Melalap":          range(58, 74),
    "Pitas":            range(58, 74),
    "SEO":              range(58, 74),
    "Sook/Dalit":       range(58, 74),
    "Tuaran Crumb":     range(58, 74),
    "Yun Fook":         range(58, 74),
    "Bornion (M)":      range(58, 74),
    "Sabang 2 (M)":     range(58, 74),
    "Desa Talisai (M)": range(58, 74),
}

# ============================================================
# STEP 1 — Read Order Receive, collect each customer's total
# ============================================================

print(f"Reading Order Receive — month: {CURRENT_MONTH}")
print("=" * 50)

df = pd.read_excel(ORDER_FILE, sheet_name=CURRENT_MONTH, header=None)

results = {}
current_customer = None

for index, row in df.iterrows():
    col0 = str(row[0]).strip()
    col1 = str(row[1]).strip()
    col4 = str(row[4]).strip()
    col5 = str(row[5]).strip()

    if col0.upper() in NAME_MAP:
        current_customer = col0.upper()
    elif col1.upper() in NAME_MAP:
        current_customer = col1.upper()

    if col4 == "Total" and current_customer is not None:
        amount_str = col5.replace("$", "").replace(",", "").strip()
        try:
            amount = float(amount_str)
            mapped_name = NAME_MAP[current_customer]
            results[mapped_name] = amount
            print(f"  Found: {mapped_name:25} = RM {amount:>12,.2f}")
        except:
            pass
        current_customer = None

print("=" * 50)
print(f"Total customers found: {len(results)}")
print()

# ============================================================
# STEP 2 — Write into Sales Report
# ============================================================

print("Opening Sales Report...")
wb = openpyxl.load_workbook(SALES_FILE)
ws = wb.active

updates_made = 0
not_found = []
found_names = []

for row_idx, row in enumerate(ws.iter_rows(), start=1):
    cell_name = row[1].value  # Column B

    if cell_name is None:
        continue

    clean_name = str(cell_name).strip()

    if clean_name in results and clean_name not in found_names:
        target_cell = row[MONTH_COLUMN - 1]
        old_value = target_cell.value
        new_value = results[clean_name]

        if old_value != new_value:
            target_cell.value = new_value
            print(f"  Updated: {clean_name:25} | "
                  f"Old: {str(old_value):>12} → "
                  f"New: {new_value:>12,.2f}")
            updates_made += 1
        else:
            print(f"  Same:    {clean_name:25} | "
                  f"{new_value:>12,.2f} (no change)")

        # Mark this name as done so we skip the duplicate row
        found_names.append(clean_name)

# Check for any customers not found in Sales Report
for name in results:
    if name not in found_names:
        not_found.append(name)

# ============================================================
# STEP 3 — Save
# ============================================================

if SAVE_MODE == "overwrite":
    wb.save(SALES_FILE)
    print(f"\n✅ Saved — overwrote original file.")
elif SAVE_MODE == "new":
    new_filename = SALES_FILE.replace(
        ".xlsx", f"_{CURRENT_MONTH}_UPDATED.xlsx")
    wb.save(new_filename)
    print(f"\n✅ Saved new file: {new_filename}")

# ============================================================
# SUMMARY
# ============================================================

print()
print("=" * 50)
print(f"  Total updates made:          {updates_made}")
print(f"  Customers not found:         {len(not_found)}")
if not_found:
    for name in not_found:
        print(f"    ⚠️  {name} — check name spelling!")
print("=" * 50)