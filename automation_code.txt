import os
import glob
import csv
import re
from datetime import datetime
import openpyxl

# Developed by Eva Karakostas using Python.
# AI tools were used to assist with logic and optimization.

# ---------------------------------
# INPUT & OUTPUT FOLDERS
# ---------------------------------

BASE_INPUT = r"INPUT_FOLDER_PATH"
BASE_OUTPUT = r"OUTPUT_FOLDER_PATH"

# ---------------------------------
# Part Configuration
# ---------------------------------

PARTS = {
    "PART_A": {
        "lin": "LIN**BP*PART-A***PO*PARTA***PL*1~",
        "shp": "SHP*01*0000*000*00000000~"
    },
    "PART_B": {
        "lin": "LIN**BP*PART-B***PO*PARTB***PL*1~",
        "shp": "SHP*01*0000*000*00000000~"
    },
    "PART_C": {
        "lin": "LIN**BP*PART-C***PO*PARTC***PL*1~",
        "shp": "SHP*01*0000*000*00000000~"
    }
}

# ---------------------------------
# Helpers
# ---------------------------------

def get_file_date_from_name(file_path):
    filename = os.path.basename(file_path)
    match = re.search(r"(\d{4}\.\d{2}\.\d{2})", filename)

    if not match:
        return None

    return match.group(1)


def get_file_date_obj(file_path):
    file_date = get_file_date_from_name(file_path)

    if not file_date:
        return None

    return datetime.strptime(file_date, "%Y.%m.%d")


def fmt_date(value):
    if isinstance(value, datetime):
        return value.strftime("%Y%m%d")

    if value:
        value = str(value).strip()

        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%Y.%m.%d"):
            try:
                return datetime.strptime(value, fmt).strftime("%Y%m%d")
            except ValueError:
                pass

    return None


def get_part_family(part_no):
    if not part_no:
        return None

    part_no = str(part_no).strip().upper()

    for part in PARTS:
        if part_no.startswith(part):
            return part

    return None


def get_line_number(item_no):
    if not item_no:
        return "1"

    item_no = str(item_no).strip()
    match = re.search(r"(\d+)", item_no)

    if match:
        return str(int(match.group(1)))

    return "1"


def make_ref_from_order(order_no, item_no):
    if not order_no:
        raise ValueError("Missing Order No.")

    order_no = str(order_no).strip().upper()

    order_match = re.search(r"(\d{4})-(\d+)", order_no)
    if not order_match:
        raise ValueError(f"Could not parse order number: {order_no}")

    base_order = order_match.group(1) + order_match.group(2)

    suffix = ""
    if "/" in order_no:
        suffix = order_no.split("/")[-1].replace("-", "")

    line_no = get_line_number(item_no)

    return f"{base_order}-{suffix}-{line_no}"


# ---------------------------------
# Find newest dated non-temp Excel file
# ---------------------------------

xlsx_files = [
    f for f in glob.glob(os.path.join(BASE_INPUT, "*.xlsx"))
    if not os.path.basename(f).startswith("~$")
]

xlsx_files = [
    f for f in xlsx_files
    if get_file_date_obj(f) is not None
]

if not xlsx_files:
    raise FileNotFoundError("No dated Order Excel files found.")

excel_file = max(xlsx_files, key=get_file_date_obj)

print("=" * 80)
print(f"Using newest dated Excel file only: {excel_file}")

file_date = get_file_date_from_name(excel_file)

today = datetime.today()

edi_short_date = today.strftime("%y%m%d")
edi_long_date = today.strftime("%Y%m%d")
today_file = today.strftime("%Y.%m.%d")

today_file = datetime.today().strftime("%Y.%m.%d")

# Create folder:
# Example: 2026.05.05 (2026.05.08 ~ 2026.05.08 Orders)

week_folder_name = f"{today_file} ({file_date} Orders)"
output_folder_path = os.path.join(BASE_OUTPUT, week_folder_name)

os.makedirs(output_folder_path, exist_ok=True)

print(f"Using Output Folder: {output_folder_path}")

# ---------------------------------
# Load Excel
# ---------------------------------

wb = openpyxl.load_workbook(excel_file, data_only=True)
ws = wb.active

# ---------------------------------
# Auto-detect Header Row
# ---------------------------------

header_row = None
headers = None

required_headers = [
    "Order No.",
    "Item No.",
    "Part No.",
    "Order Quantity",
    "Delivery Date"
]

for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
    if row:
        row_values = [str(c).strip() if c else "" for c in row]

        if all(h in row_values for h in required_headers):
            header_row = i
            headers = row_values
            break

if not header_row:
    raise ValueError(
        f"Header row not found in {excel_file}. "
        "Please confirm the Excel headers match exactly."
    )

col = {name: idx for idx, name in enumerate(headers)}

# ---------------------------------
# Collect Rows by Part
# ---------------------------------

part_rows = {part: [] for part in PARTS}

for r in ws.iter_rows(min_row=header_row + 1, values_only=True):
    part_no = r[col["Part No."]]
    part = get_part_family(part_no)

    if part not in PARTS:
        continue

    qty = r[col["Order Quantity"]]

    if not qty or qty == 0:
        continue

    deliver = fmt_date(r[col["Delivery Date"]])

    if not deliver:
        raise ValueError(f"Missing or invalid delivery date for part {part_no}")

    ref = make_ref_from_order(
        r[col["Order No."]],
        r[col["Item No."]]
    )

    part_rows[part].append({
        "qty": int(qty),
        "deliver": deliver,
        "ref": ref
    })

# ---------------------------------
# Process each part separately
# ---------------------------------

for part, rows in part_rows.items():
    cfg = PARTS[part]

    if not rows:
        print(f"Skipping {part} for {file_date} - no orders")
        continue

    dl_start_part = min(r["deliver"] for r in rows)
    dl_end_part = max(r["deliver"] for r in rows)

    edi = [
        f"ISA*00*          *00*          *ZZ*SENDER_ID     *ZZ*RECEIVER_ID     *{edi_short_date}*2121*U*00401*000000059*0*P*>~",
        f"GS*S*SENDER_ID*RECEIVER_ID*{edi_long_date}*2121*59*X*004010~",
        "ST*862*000059001~",
        f"BSS*05*{edi_long_date}000059001*{edi_long_date}*DL*{dl_start_part}*{dl_end_part}~",
        f"DTM*102*{edi_long_date}*2121~",
        "N1*SU*SUPPLIER*92*SUPPLIER_CODE~",
        "N1*ST*CUSTOMER*92*CUSTOMER_CODE~",
        "REF*DK*CUSTOMER_CODE~",
        cfg["lin"],
        "UIT*PC~"
    ]

    for r in rows:
        edi.extend([
            f"FST*{r['qty']}*C*C*{r['deliver']}~",
            f"JIT*{r['qty']}*000001~",
            f"REF*RE*{r['ref']}~"
        ])

    edi.append(cfg["shp"])

    # ---------------------------------
    # Calculate proper SE segment
    # ---------------------------------

    edi.append("CTT*1*63900~")

    st_index = next(i for i, s in enumerate(edi) if s.startswith("ST"))
    se_count = len(edi) - st_index + 1

    edi.append(f"SE*{se_count}*000059001~")
    edi.append("GE*1*59~")
    edi.append("IEA*1*000000059~")

    # ---------------------------------
    # Save as CSV
    # Example:
    # CUSTOMER_862_PART_A_2026.05.08.csv
    # ---------------------------------

    out_name = f"CUSTOMER_862_{part}_{file_date}.csv"
    out_path = os.path.join(output_folder_path, out_name)

    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        for line in edi:
            writer.writerow([line])

    print(f"Created: {out_path}")

print("=" * 80)
print("Newest EDI file created successfully.")