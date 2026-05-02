import pandas as pd

ORDER_FILE = r"C:\Users\samue\Downloads\ORDER RECEIVE -2026.xlsx"
SALES_FILE = r"C:\Users\samue\Downloads\Yearly Sales Report 2026(SALES_REPORT_2026).xlsx"

CURRENT_MONTH = "APR"

NAME_MAP = {
    # ALAN's customers
    "ABEDON": "Abedon",
    "G.TANJUNG": "G.Tanjung",
    "LEEPANG": "Leepang",
    "MORISEM": "Morisem",
    "SYARIMO": "Syarimo",
    "TUNG HUP": "Tung Hup",       # <-- different name! fixed here
    "UNICO DESA": "Unico Desa",

    # ALWIN's customers
    "ATLANTICA": "Atlantica",
    "BEAUFORT": "Beaufort",
    "DESA KIM LOONG": "Desa Kim Loong",  # <-- different name! fixed here
    "G.EDIBLE OIL": "G.Edible",
    "IOI EDIBLE/BIO": "IOI Edible/Bio",
    "MALSA": "Malsa",
    "MERIDIAN": "Meridian",
    "MELALAP": "Melalap",
    "PITAS": "Pitas",
    "SEO": "SEO",
    "SOOK/DALIT": "Sook/Dalit",
    "TUARAN CRUMB": "Tuaran Crumb",
    "YUN FOOK": "Yun Fook"
}

print(f"Reading Order Receive for month: {CURRENT_MONTH}")
print("=" * 50)

# Read the sheet - header=None means "don't treat row 1 as column names"
df = pd.read_excel(ORDER_FILE, sheet_name=CURRENT_MONTH, header=None)

results = {}   # This will store {customer_name: total_amount}

current_customer = None   # We'll track which customer we're inside

for index, row in df.iterrows():
    # Convert row values to strings for easy searching
    col1 = str(row[0]).strip()   # First column (salesperson or date)
    col2 = str(row[1]).strip()   # Second column (customer name sometimes)
    col4 = str(row[4]).strip()   # Fifth column (often says "Total")
    col5 = str(row[5]).strip()   # Sixth column (the amount)

    # Check if this row has a customer name we recognize
    name_upper = col1.upper()
    if name_upper in NAME_MAP or col2.upper() in NAME_MAP:
        if col1.upper() in NAME_MAP:
            current_customer = col1.upper()
        else:
            current_customer = col2.upper()

    # Check if this is a Total row
    if col4 == "Total" and current_customer is not None:
        # Clean the amount - remove $, commas, spaces
        amount_str = col5.replace("$", "").replace(",", "").strip()
        try:
            amount = float(amount_str)
            mapped_name = NAME_MAP[current_customer]
            results[mapped_name] = amount
            print(f"  Found: {mapped_name:25} = RM {amount:>12,.2f}")
        except:
            pass  # Skip if amount can't be converted to number
        current_customer = None  # Reset for next customer

print()
print("=" * 50)
print(f"Total customers found: {len(results)}")