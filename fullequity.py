import os
import math
import subprocess
from copy import copy
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side


# ===================== FILE PATHS =====================
SOURCE_FILE_XLS = "SEBI_Monthly_Portfolio 31 JAN 2026.xls"
SOURCE_FILE_XLSX = "SEBI_Monthly_Portfolio 31 JAN 2026.xlsx"
MASTER_FILE = "IN_MF_PORTFOLIO_DETAILS_ACROSS_SCHEMES.xlsx"
OUTPUT_FILE = "SEBI_Monthly_Portfolio_Equity_All_Sections_Reconciled.xlsx"

# IMPORTANT:
# In the uploaded master file, Total Market Value (Rs.) matches SEBI's
# "Market/Fair Value\n(Rs.in Lacs)" after dividing by 100000 because 1 lac = 100000.
MARKET_VALUE_DIVISOR = 100000
QUANTITY_DIVISOR = 100000

# Trial balance workbook used only for these two remaining equity sections.
TRIAL_BALANCE_FILE = "CITI_ABC_Trial_Balance_310126.xlsx"
TRIAL_BALANCE_SHEET = "sg_in003_Y4X_300126_310126"

TRIAL_BALANCE_SECTION_ACCOUNT_CODE = {
    "Cash and Bank": "141839",
    "Margin (Future and Options)": "141350",
}

# NOTE: % to Net Assets for Cash and Bank / Margin will be added later.
TRIAL_BALANCE_COMPARE_COLS = ["Quantity"]

REMAINING_EQUITY_SECTIONS = {
    "Others",
    "TREPS / Reverse Repo",
    "Foreign Securities and/or overseas ETF(s)",
    "(a) Listed / awaiting listing on Stock Exchanges",
    "International Mutual Fund Units",
    "Exchange Traded Funds",
    "Equity Futures/Index",
    "Margin (Future and Options)",
    "Cash and Bank",
}

MASTER_SECTION_SECURITY_TYPES = {
    "TREPS / Reverse Repo": ["COLLATERALISED BORROWING AND LENDIN", "REPO"],
    "Foreign Securities and/or overseas ETF(s)": ["International Equity"],
    "(a) Listed / awaiting listing on Stock Exchanges": ["International Equity"],
    "International Mutual Fund Units": ["INVESTMENT FUNDS/MUTUAL FUNDS"],
    "Exchange Traded Funds": ["Exchange Traded Fund"],
    "Equity Futures/Index": ["EQUITY  FUTURE"],
}

NON_SECTION_LABELS = {
    "sub total",
    "total",
    "grand total",
    "net receivables / (payables)",
}

# ===================== SCHEME TYPE DICTIONARY FROM TESTTWO.PY =====================
scheme_dict = {'ABBSEIIF': 'EQUITY', 'ABSLBCF': 'EQUITY', 'ABSLCONF': 'EQUITY', 'ABSLESG': 'EQUITY',
'ABSLLDF': 'DEBT', 'ABSLMAAF': 'HYBRID', 'ABSLMCF': 'EQUITY', 'ABSLQF': 'EQUITY', 'ABSLSO': 'EQUITY',
'ABSLTNLF': 'EQUITY', 'ABSLUS03': 'FOF', 'ABSLUS10': 'FOF', 'ADVG': 'EQUITY', 'BANKETF': 'EQUITY', 'BBIF': 'DEBT',
'BBP': 'DEBT', 'BDB': 'DEBT', 'BDYP': 'EQUITY', 'BFL': 'DEBT', 'BFS': 'DEBT', 'BINFRA': 'EQUITY',
'BINTEQA': 'INTRNL', 'BMIDX50': 'EQUITY', 'BQIDX50': 'EQUITY', 'BSL95F': 'HYBRID', 'BSLAAMM': 'FOF',
'BSLADMM': 'FOF', 'BSLBBYW': 'EQUITY', 'BSLBKFS': 'EQUITY', 'BSLCBF': 'DEBT', 'BSLCM': 'DEBT', 'BSLDAAF': 'HYBRID',
'BSLEAF': 'HYBRID', 'BSLEQSF': 'HYBRID', 'BSLEQTY': 'EQUITY', 'BSLFEF': 'HYBRID', 'BSLFPAP': 'FOF', 'BSLFPCP': 'FOF',
'BSLFPPP': 'FOF', 'BSLIF': 'DEBT', 'BSLMFG': 'EQUITY', 'BSLMIFOF': 'FOF', 'BSLMTP': 'DEBT', 'BSLNMF': 'EQUITY',
'BSLONF': 'DEBT', 'BSLPHF': 'EQUITY', 'BSLR96': 'EQUITY', 'BSLRF30': 'HYBRID', 'BSLRF40': 'HYBRID',
'BSLRF50': 'HYBRID', 'BSLRF50P': 'DEBT', 'BSLSTF': 'DEBT', 'BSLTA1': 'EQUITY', 'BTOP100': 'EQUITY',
'C10YGETF': 'DEBT', 'CASH': 'DEBT', 'CBGETF': 'DEBT', 'CIGAPR26': 'DEBT', 'CIGAPR28': 'DEBT', 'CIGAPR29': 'DEBT',
'CIGAPR33': 'DEBT', 'CIGJUN27': 'DEBT', 'CIGPSA28': 'DEBT', 'CISJUN32': 'DEBT', 'COFAIETF': 'DEBT',
'CSFSD12': 'DEBT', 'CSFSD6': 'DEBT', 'CSFSI27': 'DEBT', 'CSNHFS26': 'DEBT', 'CSPAPA26': 'DEBT', 'CSPAPA27': 'DEBT',
'FTPTI': 'DEBT', 'FTPTJ': 'DEBT', 'FTPTQ': 'DEBT', 'FTPUB': 'DEBT', 'FTPUJ': 'DEBT', 'GENNEXT': 'EQUITY',
'GOLDETF': 'GOLD', 'GOLDFOF': 'OTH', 'INV': 'DEBT', 'MIDCAP': 'EQUITY', 'MIP25': 'HYBRID', 'MNC': 'EQUITY',
'N30MTETF': 'EQUITY', 'N30QTETF': 'EQUITY', 'NEWIF50': 'EQUITY', 'NHLTHETF': 'EQUITY', 'NIFTY': 'EQUITY',
'NIFTYETF': 'EQUITY', 'NIFYNX50': 'EQUITY', 'NINDDEF': 'EQUITY', 'NITETF': 'EQUITY', 'NMID150': 'EQUITY',
'NPSEETF': 'EQUITY', 'NSDAPR27': 'DEBT', 'NSDAQ100': 'FOF', 'NSDSEP27': 'DEBT', 'NSMALL50': 'EQUITY',
'NSPPBS26': 'DEBT', 'NXTIDX50': 'EQUITY', 'PLUS': 'DEBT', 'PSUEQ': 'EQUITY', 'PURE': 'EQUITY', 'SENSXETF': 'EQUITY',
'SILVRETF': 'SILVER', 'SILVRFOF': 'OTH', 'BSLGCF': 'FOF', 'BSLGRE': 'FOF'}

SEBI_COLUMNS = [
    "Name of the Instrument",
    "ISIN",
    "Industry^ / Rating",
    "Quantity",
    "Market/Fair Value\n(Rs.in Lacs)",
    "% to Net Assets",
]

MASTER_RENAME = {
    "Issuer Name": "Name of the Instrument",
    "Industry": "Industry^ / Rating",
    "Total Market Value (Rs.)": "Market/Fair Value\n(Rs.in Lacs)",
    "% to Net assests": "% to Net Assets",
}

COMPARE_COLS = [
    "Name of the Instrument",
    "Industry^ / Rating",
    "Quantity",
    "Market/Fair Value\n(Rs.in Lacs)",
    "% to Net Assets",
]

HIGHLIGHT_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")     # yellow
TOTAL_FILL = PatternFill(fill_type="solid", fgColor="F4CCCC")         # light red
HEADER_FILL = PatternFill(fill_type="solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def resolve_source_xlsx():
    """
    Return an xlsx file path that openpyxl can edit.

    openpyxl cannot edit old .xls files directly.
    So this function first looks for an already-saved .xlsx file.
    If only .xls exists, it tries:
      1) Microsoft Excel COM conversion on Windows, if Excel + pywin32 are available.
      2) LibreOffice conversion, if LibreOffice is installed and available in PATH.

    Easiest manual fix if conversion fails:
      Open the SEBI .xls file in Excel -> File -> Save As -> .xlsx
      and keep the file name as SOURCE_FILE_XLSX below.
    """
    if os.path.exists(SOURCE_FILE_XLSX):
        return SOURCE_FILE_XLSX

    if not os.path.exists(SOURCE_FILE_XLS):
        raise FileNotFoundError(
            f"Could not find {SOURCE_FILE_XLSX} or {SOURCE_FILE_XLS} in this folder."
        )

    # Method 1: Windows + Microsoft Excel installed
    try:
        import win32com.client  # pip install pywin32

        src = os.path.abspath(SOURCE_FILE_XLS)
        dst = os.path.abspath(SOURCE_FILE_XLSX)

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb_excel = excel.Workbooks.Open(src)
        # 51 = xlOpenXMLWorkbook = .xlsx
        wb_excel.SaveAs(dst, FileFormat=51)
        wb_excel.Close(False)
        excel.Quit()

        if os.path.exists(SOURCE_FILE_XLSX):
            return SOURCE_FILE_XLSX
    except Exception as excel_error:
        excel_conversion_error = excel_error

    # Method 2: LibreOffice installed and available in PATH
    libreoffice_commands = ["libreoffice", "soffice"]
    for command in libreoffice_commands:
        try:
            subprocess.run(
                [command, "--headless", "--convert-to", "xlsx", SOURCE_FILE_XLS, "--outdir", "."],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            if os.path.exists(SOURCE_FILE_XLSX):
                return SOURCE_FILE_XLSX
        except FileNotFoundError:
            continue
        except Exception as libreoffice_error:
            last_libreoffice_error = libreoffice_error

    raise RuntimeError(
        "Could not convert the SEBI .xls file to .xlsx.\n\n"
        "Fix option 1: Open 'SEBI_Monthly_Portfolio 31 JAN 2026.xls' in Excel, "
        "then Save As 'SEBI_Monthly_Portfolio 31 JAN 2026.xlsx', and run this script again.\n"
        "Fix option 2: Install pywin32 using: pip install pywin32, if Microsoft Excel is installed.\n"
        "Fix option 3: Install LibreOffice and add it to PATH.\n\n"
        f"Excel conversion error was: {excel_conversion_error}"
    )


def clean_text(value):
    if value is None:
        return ""
    return str(value).replace("\r", "\n").strip()


def normalize_isin(value):
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    return str(value).strip().upper()


def norm_header(value):
    """Normalize Excel headers so CRLF/LF/_x000D_/extra spaces do not break matching."""
    if value is None:
        return ""
    text = str(value)
    # When .xls is converted to .xlsx, Excel/openpyxl can expose carriage return as literal _x000D_.
    text = text.replace("_x000D_", "\n").replace("\r", "\n")
    lines = [part.strip() for part in text.split("\n") if part.strip()]
    return "\n".join(lines).strip()


def canonical_header(value):
    """Return the standard column name used by this script."""
    h = norm_header(value)
    flat = " ".join(h.replace("\n", " ").split()).casefold()

    aliases = {
        "name of the instrument": "Name of the Instrument",
        "isin": "ISIN",
        "industry^ / rating": "Industry^ / Rating",
        "industry^/ rating": "Industry^ / Rating",
        "industry^/rating": "Industry^ / Rating",
        "quantity": "Quantity",
        "market/fair value (rs.in lacs)": "Market/Fair Value\n(Rs.in Lacs)",
        "market/fair value (rs. in lacs)": "Market/Fair Value\n(Rs.in Lacs)",
        "% to net assets": "% to Net Assets",
    }
    return aliases.get(flat, h)


def to_number(value):
    if value is None or value == "":
        return None
    if isinstance(value, str):
        value = value.replace(",", "").replace("%", "").replace("$", "").strip()
        if value in {"", "-"}:
            return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    try:
        return float(value)
    except Exception:
        return None


def to_decimal(value):
    if value is None or value == "":
        return None

    if isinstance(value, str):
        value = value.replace(",", "").replace("%", "").replace("$", "").strip()
        if value in {"", "-"}:
            return None

    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    try:
        return Decimal(str(value).strip())
    except (InvalidOperation, ValueError):
        return None


def round_half_up(value, decimals=2):
    dec_value = to_decimal(value)

    if dec_value is None:
        return None
    
    quant = Decimal("1").scaleb(-decimals)     #decimal=2 -> Decimal("0.01")
    return dec_value.quantize(quant, rounding="ROUND_HALF_UP")


def percent_value_for_compare(value, source):
    dec_value = to_decimal(value)

    if dec_value is None:
        return None

    if source == "SEBI":
        # SEBI percentage cells are stored internally as decimal values.
        # Example formula bar 6.46% is read by Python as 0.0646.
        return dec_value * Decimal("100")

    # Master already stores percentage as normal number.
    # Example 6.46 means 6.46%.
    return dec_value


def format_percent_for_index(value, source):
    if value is None:
        return ""
    
    raw_text = str(value).strip()

    # If SEBI already has a visible percentage string like "$0.00%"
    # show it exactly as it appears in the SEBI sheet.
    if source == "SEBI" and "%" in raw_text:
        return raw_text
    
    dec_value = to_decimal(value)

    if dec_value is None:
        return ""

    if source == "SEBI":
        dec_value = dec_value * Decimal("100")

    return format(dec_value, "f").rstrip("0").rstrip(".") + "%"


def round_to_2(dec_value):
    if dec_value is None:
        return None
    return dec_value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def quantity_value_for_compare(value, source):
    dec_value = to_decimal(value)

    if dec_value is None:
        return None

    # Master / Trial Balance quantity-like values are stored in full units.
    # SEBI value is already in the comparable lakhs-style value.
    if source in {"MASTER", "TRIAL"}:
        return dec_value / Decimal(str(QUANTITY_DIVISOR))

    return dec_value


def format_quantity_for_index(value, source):
    dec_value = quantity_value_for_compare(value, source)

    if dec_value is None:
        return ""

    rounded_value = round_to_2(dec_value)

    if rounded_value is None:
        return ""

    return format(rounded_value, "f")


def section_key(value):
    return clean_text(value).casefold()


def row_match_key(row):
    isin = normalize_isin(row.get("ISIN"))
    if isin:
        return isin
    return clean_text(row.get("Name of the Instrument")).casefold()


def is_blank_cell_value(value):
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    return str(value).strip() == ""


def has_any_comparable_value(row):
    return any(not is_blank_cell_value(row.get(col)) for col in SEBI_COLUMNS if col != "Name of the Instrument")


def same_value(master_value, sebi_value, column_name):
    # Compare numeric values after preprocessing and ROUND_HALF_UP to 2 decimal places.

    if column_name == "% to Net Assets":
        mv = percent_value_for_compare(master_value, "MASTER")
        sv = percent_value_for_compare(sebi_value, "SEBI")

        if mv is None and sv is None:
            return True

        if mv is None or sv is None:
            return False

        return round_to_2(mv) == round_to_2(sv)

    if column_name == "Quantity":
        mv = quantity_value_for_compare(master_value, "MASTER")
        sv = quantity_value_for_compare(sebi_value, "SEBI")

        if mv is None and sv is None:
            return True

        if mv is None or sv is None:
            return False

        return round_to_2(mv) == round_to_2(sv)

    mv = to_decimal(master_value)
    sv = to_decimal(sebi_value)

    if mv is not None or sv is not None:
        if mv is None or sv is None:
            return False

        return round_to_2(mv) == round_to_2(sv)

    return clean_text(master_value).casefold() == clean_text(sebi_value).casefold()


def get_header_map(ws):
    """Find the SEBI header row and map columns by header names."""
    for row in range(1, ws.max_row + 1):
        values = [canonical_header(ws.cell(row, col).value) for col in range(1, ws.max_column + 1)]
        if "Name of the Instrument" in values and "ISIN" in values:
            header_row = row
            col_map = {}
            for col in range(1, ws.max_column + 1):
                header = canonical_header(ws.cell(row, col).value)
                if header in SEBI_COLUMNS:
                    col_map[header] = col
            missing = [c for c in SEBI_COLUMNS if c not in col_map]
            if missing:
                available = [norm_header(ws.cell(row, col).value) for col in range(1, ws.max_column + 1) if ws.cell(row, col).value]
                raise ValueError(f"Missing required SEBI columns {missing}. Available headers: {available}")
            return header_row, col_map
    raise ValueError(f"Header row not found in sheet {ws.title}")


def find_first_equity_section_rows(ws, header_row, name_col):
    """
    For equity schemes only:
    Start after first '(a) Listed / awaiting listing on Stock Exchange'.
    Stop at the first 'Sub Total'.
    Also capture the first 'Total' after that Sub Total.
    Nothing below this first Total is checked.
    """
    start_row = None
    subtotal_row = None
    total_row = None

    for row in range(header_row + 1, ws.max_row + 1):
        label = clean_text(ws.cell(row, name_col).value)
        label_lower = label.casefold()
        if start_row is None and "(a) listed / awaiting listing on stock" in label_lower:
            start_row = row + 1
            continue
        if start_row is not None and subtotal_row is None and label_lower == "sub total":
            subtotal_row = row
            continue
        if subtotal_row is not None and label_lower == "total":
            total_row = row
            break

    if start_row is None or subtotal_row is None:
        raise ValueError(f"First equity Sub Total section not found in sheet {ws.title}")

    # If Total is missing, still compare Sub Total and only read rows before Sub Total.
    return start_row, subtotal_row, total_row


def read_sebi_first_section(ws):
    header_row, col_map = get_header_map(ws)
    name_col = col_map["Name of the Instrument"]
    start_row, subtotal_row, total_row = find_first_equity_section_rows(ws, header_row, name_col)

    records = []
    row_by_isin = {}
    for row in range(start_row, subtotal_row):
        isin = clean_text(ws.cell(row, col_map["ISIN"]).value)
        name = clean_text(ws.cell(row, name_col).value)
        if not isin and not name:
            continue
        rec = {"Excel Row": row}
        for col in SEBI_COLUMNS:
            rec[col] = ws.cell(row, col_map[col]).value
        records.append(rec)
        if isin:
            row_by_isin[isin] = rec

    subtotal = {"Excel Row": subtotal_row}
    for col in SEBI_COLUMNS:
        subtotal[col] = ws.cell(subtotal_row, col_map[col]).value

    total = None
    if total_row:
        total = {"Excel Row": total_row}
        for col in SEBI_COLUMNS:
            total[col] = ws.cell(total_row, col_map[col]).value

    return {
        "header_row": header_row,
        "col_map": col_map,
        "records": records,
        "row_by_isin": row_by_isin,
        "subtotal": subtotal,
        "total": total,
    }


def read_master_equity(master_df, scheme_code):
    master_table = master_df[
        (master_df["Client Code"] == scheme_code)
        & (master_df["Security Type Name"].isin(["EQUITY", "PREFERRED STOCK"]))
    ].copy()

    master_table = master_table[
        [
            "Client Code",
            "Issuer Name",
            "Security Type Name",
            "ISIN",
            "Industry",
            "Quantity",
            "Total Market Value (Rs.)",
            "% to Net assests",
        ]
    ]

    master_table = master_table.rename(columns=MASTER_RENAME)
    master_table["Market/Fair Value\n(Rs.in Lacs)"] = master_table[
        "Market/Fair Value\n(Rs.in Lacs)"
        ].apply(
            lambda x: round_to_2(Decimal(str(x)) / Decimal(str(MARKET_VALUE_DIVISOR))) 
            if pd.notna(x) else None
        )

    # Master stores % as 1.23 for 1.23%.
    # Do NOT divide by 100.
    # SEBI percentage is converted only during comparison/output.
    master_table["% to Net Assets"] = pd.to_numeric(
        master_table["% to Net Assets"], errors="coerce"
    )

    master_table = master_table[
        [
            "Name of the Instrument",
            "ISIN",
            "Industry^ / Rating",
            "Quantity",
            "Market/Fair Value\n(Rs.in Lacs)",
            "% to Net Assets",
        ]
    ].reset_index(drop=True)

    records = master_table.to_dict("records")
    by_isin = {clean_text(r["ISIN"]): r for r in records if clean_text(r.get("ISIN"))}
    return master_table, records, by_isin



def read_trial_balance_df():
    if not os.path.exists(TRIAL_BALANCE_FILE):
        return pd.DataFrame()

    df = pd.read_excel(TRIAL_BALANCE_FILE, sheet_name=TRIAL_BALANCE_SHEET)
    df.columns = [str(c).strip() for c in df.columns]
    if "Account Code" in df.columns:
        df["Account Code"] = df["Account Code"].apply(lambda x: str(x).strip().split(".")[0] if pd.notna(x) else "")
    if "Client Code" in df.columns:
        df["Client Code"] = df["Client Code"].apply(lambda x: clean_text(x))
    return df


def read_sebi_remaining_sections(ws, first_total_row, col_map):
    sections = {}
    current_section = None

    def ensure_section(section_name):
        if section_name not in sections:
            sections[section_name] = {"records": [], "subtotals": [], "totals": []}
        return sections[section_name]

    for row in range(first_total_row + 1, ws.max_row + 1):
        label = clean_text(ws.cell(row, col_map["Name of the Instrument"]).value)
        label_key = section_key(label)

        if not label:
            continue

        if label_key == "grand total":
            break

        rec = {"Excel Row": row}
        for col in SEBI_COLUMNS:
            rec[col] = ws.cell(row, col_map[col]).value

        if label_key in NON_SECTION_LABELS:
            if current_section and current_section in sections:
                if label_key == "sub total":
                    sections[current_section]["subtotals"].append(rec)
                elif label_key == "total":
                    sections[current_section]["totals"].append(rec)
                    current_section = None
            continue

        # Cash and Bank / Margin are line items in the SEBI file, not empty section headers.
        if label in TRIAL_BALANCE_SECTION_ACCOUNT_CODE:
            ensure_section(label)["records"].append(rec)
            current_section = label
            continue

        # Real section heading.
        if label in REMAINING_EQUITY_SECTIONS and not has_any_comparable_value(rec):
            ensure_section(label)
            current_section = label
            continue

        # Ignore note/disclosure rows that are outside a target section.
        if current_section is None or current_section not in sections:
            continue

        # Do not treat nested parent marker "Others" as a security row.
        if label == "Others" and not has_any_comparable_value(rec):
            current_section = label
            ensure_section(label)
            continue

        sections[current_section]["records"].append(rec)

    return sections


def read_master_remaining_section(master_df, scheme_code, section_name):
    security_types = MASTER_SECTION_SECURITY_TYPES.get(section_name, [])
    if not security_types:
        return [], {}

    master_table = master_df[
        (master_df["Client Code"] == scheme_code)
        & (master_df["Security Type Name"].isin(security_types))
    ].copy()

    if master_table.empty:
        return [], {}

    master_table["Name of the Instrument"] = master_table["Issuer Name"].where(
        master_table["Issuer Name"].notna(), master_table["Security Name"]
    )
    master_table["Industry^ / Rating"] = master_table["Industry"]
    master_table["Market/Fair Value\n(Rs.in Lacs)"] = master_table["Total Market Value (Rs.)"].apply(
        lambda x: round_to_2(Decimal(str(x)) / Decimal(str(MARKET_VALUE_DIVISOR))) if pd.notna(x) else None
    )
    master_table["% to Net Assets"] = pd.to_numeric(master_table["% to Net assests"], errors="coerce")

    master_table = master_table[
        [
            "Name of the Instrument",
            "ISIN",
            "Industry^ / Rating",
            "Quantity",
            "Market/Fair Value\n(Rs.in Lacs)",
            "% to Net Assets",
        ]
    ].reset_index(drop=True)

    records = master_table.to_dict("records")
    by_key = {row_match_key(r): r for r in records if row_match_key(r)}
    return records, by_key


def read_trial_balance_section(trial_balance_df, scheme_code, section_name):
    if trial_balance_df is None or trial_balance_df.empty:
        return [], {}

    account_code = TRIAL_BALANCE_SECTION_ACCOUNT_CODE.get(section_name)
    if not account_code:
        return [], {}

    required_cols = {"Client Code", "Account Code", "Opening Balance"}
    if not required_cols.issubset(set(trial_balance_df.columns)):
        return [], {}

    filtered = trial_balance_df[
        (trial_balance_df["Client Code"].astype(str).str.strip() == scheme_code)
        & (trial_balance_df["Account Code"].astype(str).str.strip() == account_code)
    ].copy()

    if filtered.empty:
        return [], {}

    records = []
    for _, row in filtered.iterrows():
        records.append(
            {
                "Name of the Instrument": section_name,
                "ISIN": "",
                "Industry^ / Rating": "",
                # Opening Balance is compared against SEBI Quantity after /100000 and ROUND_HALF_UP.
                "Quantity": row.get("Opening Balance"),
                "Market/Fair Value\n(Rs.in Lacs)": None,
                # % to Net Assets for these sections will be implemented later.
                "% to Net Assets": None,
            }
        )

    by_key = {section_key(section_name): records[0]}
    return records, by_key


def pair_remaining_rows(master_records, sebi_records):
    pairs = []
    master_by_key = {row_match_key(r): r for r in master_records if row_match_key(r)}
    sebi_by_key = {row_match_key(r): r for r in sebi_records if row_match_key(r)}

    used_master = set()
    used_sebi = set()

    for key in sorted(set(master_by_key.keys()) & set(sebi_by_key.keys())):
        pairs.append((key, master_by_key[key], sebi_by_key[key]))
        used_master.add(key)
        used_sebi.add(key)

    unmatched_master = [r for r in master_records if row_match_key(r) not in used_master]
    unmatched_sebi = [r for r in sebi_records if row_match_key(r) not in used_sebi]

    # If both sides have exactly one unmatched row, compare one-to-one.
    # Useful for TREPS and Equity Futures where SEBI may not carry ISIN.
    if len(unmatched_master) == 1 and len(unmatched_sebi) == 1:
        key = row_match_key(unmatched_sebi[0]) or row_match_key(unmatched_master[0])
        pairs.append((key, unmatched_master[0], unmatched_sebi[0]))
        unmatched_master = []
        unmatched_sebi = []

    return pairs, unmatched_master, unmatched_sebi


def add_processed_discrepancy(discrepancies, scheme, row_type, isin, instrument, column, master_value, sebi_value, status, excel_row=None):
    add_discrepancy(discrepancies, scheme, row_type, isin, instrument, column, master_value, sebi_value, status, excel_row)


def compare_remaining_equity_sections(ws, scheme_code, master_df, trial_balance_df, sebi, all_discrepancies):
    col_map = sebi["col_map"]
    first_total_row = sebi["total"]["Excel Row"] if sebi.get("total") else None
    if not first_total_row:
        return

    sections = read_sebi_remaining_sections(ws, first_total_row, col_map)

    for section_name, section_data in sections.items():
        sebi_records = section_data.get("records", [])
        if not sebi_records:
            continue

        if section_name in TRIAL_BALANCE_SECTION_ACCOUNT_CODE:
            source_records, _ = read_trial_balance_section(trial_balance_df, scheme_code, section_name)
            compare_cols = TRIAL_BALANCE_COMPARE_COLS
        else:
            source_records, _ = read_master_remaining_section(master_df, scheme_code, section_name)
            compare_cols = COMPARE_COLS

        pairs, missing_master, missing_sebi = pair_remaining_rows(source_records, sebi_records)

        for key, master_row, sebi_row in pairs:
            isin = normalize_isin(master_row.get("ISIN")) or normalize_isin(sebi_row.get("ISIN"))
            instrument = master_row.get("Name of the Instrument") or sebi_row.get("Name of the Instrument")

            for col in compare_cols:
                # Cash and Bank / Margin % to Net Assets logic is intentionally skipped for now.
                if section_name in TRIAL_BALANCE_SECTION_ACCOUNT_CODE and col == "% to Net Assets":
                    continue

                if not same_value(master_row.get(col), sebi_row.get(col), col):
                    add_processed_discrepancy(
                        all_discrepancies,
                        scheme_code,
                        f"Remaining Section - {section_name}",
                        isin,
                        instrument,
                        col,
                        master_row.get(col),
                        sebi_row.get(col),
                        "Value mismatch",
                        sebi_row.get("Excel Row"),
                    )
                    highlight_cell(ws, sebi_row.get("Excel Row"), col_map.get(col), HIGHLIGHT_FILL)

        for master_row in missing_master:
            add_discrepancy(
                all_discrepancies,
                scheme_code,
                f"Remaining Section - {section_name}",
                normalize_isin(master_row.get("ISIN")),
                master_row.get("Name of the Instrument"),
                "Entire Row",
                "Present",
                None,
                "SEBI sheet does not have this entire row",
                None,
            )

        for sebi_row in missing_sebi:
            add_discrepancy(
                all_discrepancies,
                scheme_code,
                f"Remaining Section - {section_name}",
                normalize_isin(sebi_row.get("ISIN")),
                sebi_row.get("Name of the Instrument"),
                "Entire Row",
                None,
                "Present",
                "Master/Trial Balance does not have this entire row",
                sebi_row.get("Excel Row"),
            )
            # Do not highlight whole missing rows; only field-level mismatches are highlighted.


def add_discrepancy(discrepancies, scheme, row_type, isin, instrument, column, master_value, sebi_value, status, excel_row=None):
    discrepancies.append(
        {
            "Scheme Code": scheme,
            "Row Type": row_type,
            "ISIN": isin,
            "Name of the Instrument": instrument,
            "Column": column,
            "Master Value": master_value,
            "SEBI Value": sebi_value,
            "Status": status,
            "SEBI Excel Row": excel_row,
        }
    )


def highlight_cell(ws, row, col, fill):
    if row and col:
        ws.cell(row, col).fill = fill


def compare_scheme(ws, scheme_code, master_df, trial_balance_df, all_discrepancies):
    sebi = read_sebi_first_section(ws)
    master_table, master_records, master_by_isin = read_master_equity(master_df, scheme_code)
    sebi_by_isin = sebi["row_by_isin"]
    col_map = sebi["col_map"]

    # 1. Security row comparison by ISIN.
    all_isins = sorted(set(master_by_isin.keys()) | set(sebi_by_isin.keys()))
    for isin in all_isins:
        master_row = master_by_isin.get(isin)
        sebi_row = sebi_by_isin.get(isin)

        if master_row is None:
            add_discrepancy(
                all_discrepancies,
                scheme_code,
                "Security Row",
                isin,
                sebi_row.get("Name of the Instrument"),
                "Entire Row",
                None,
                "Present",
                "Master workbook does not have this entire row",
                sebi_row.get("Excel Row"),
            )
            # Whole row mismatch can be highlighted because SEBI has the row.
            for col in SEBI_COLUMNS:
                highlight_cell(ws, sebi_row.get("Excel Row"), col_map.get(col), HIGHLIGHT_FILL)
            continue

        if sebi_row is None:
            add_discrepancy(
                all_discrepancies,
                scheme_code,
                "Security Row",
                isin,
                master_row.get("Name of the Instrument"),
                "Entire Row",
                "Present",
                None,
                "SEBI sheet does not have this entire row",
                None,
            )
            continue

        for col in COMPARE_COLS:
            if not same_value(master_row.get(col), sebi_row.get(col), col):
                add_discrepancy(
                    all_discrepancies,
                    scheme_code,
                    "Security Row",
                    isin,
                    master_row.get("Name of the Instrument") or sebi_row.get("Name of the Instrument"),
                    col,
                    master_row.get(col),
                    sebi_row.get(col),
                    "Value mismatch",
                    sebi_row.get("Excel Row"),
                )
                highlight_cell(ws, sebi_row.get("Excel Row"), col_map.get(col), HIGHLIGHT_FILL)

    # 2. First Sub Total and first Total comparison.
    master_market_total = round(pd.to_numeric(master_table["Market/Fair Value\n(Rs.in Lacs)"], errors="coerce").fillna(0).sum(), 2)
    master_percent_total = pd.to_numeric(master_table["% to Net Assets"], errors="coerce").fillna(0).sum()

    for row_label, sebi_total_row in [("Sub Total", sebi["subtotal"]), ("Total", sebi["total"])] :
        if not sebi_total_row:
            add_discrepancy(
                all_discrepancies,
                scheme_code,
                row_label,
                "",
                row_label,
                "Entire Row",
                "Present",
                None,
                f"SEBI sheet does not have first {row_label} row",
                None,
            )
            continue

        row_no = sebi_total_row.get("Excel Row")
        market_col = "Market/Fair Value\n(Rs.in Lacs)"
        pct_col = "% to Net Assets"

        if not same_value(master_market_total, sebi_total_row.get(market_col), market_col):
            add_discrepancy(
                all_discrepancies,
                scheme_code,
                row_label,
                "",
                row_label,
                market_col,
                master_market_total,
                sebi_total_row.get(market_col),
                "Value mismatch",
                row_no,
            )
            highlight_cell(ws, row_no, col_map.get(market_col), TOTAL_FILL)

        if not same_value(master_percent_total, sebi_total_row.get(pct_col), pct_col):
            add_discrepancy(
                all_discrepancies,
                scheme_code,
                row_label,
                "",
                row_label,
                pct_col,
                master_percent_total,
                sebi_total_row.get(pct_col),
                "Value mismatch",
                row_no,
            )
            highlight_cell(ws, row_no, col_map.get(pct_col), TOTAL_FILL)

    # 3. Remaining equity sections after the first Total.
    compare_remaining_equity_sections(ws, scheme_code, master_df, trial_balance_df, sebi, all_discrepancies)


def write_index_output(wb, discrepancies):
    if "Index" not in wb.sheetnames:
        ws = wb.create_sheet("Index", 0)
    else:
        ws = wb["Index"]

    # Clear old output from E onwards only.
    for row in range(1, ws.max_row + 1):
        for col in range(5, max(ws.max_column, 20) + 1):
            cell = ws.cell(row, col)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.font = copy(ws.cell(row, 1).font)
            cell.border = Border()
            cell.alignment = Alignment()

    title_cell = ws.cell(1, 5)
    title_cell.value = "Equity Reconciliation Output - All Equity Sections"
    title_cell.font = Font(bold=True, size=12)

    headers = [
        "Scheme Code",
        "Row Type",
        "ISIN",
        "Name of the Instrument",
        "Column",
        "Master Value",
        "SEBI Value",
        "Status",
        "SEBI Excel Row",
    ]

    start_row = 3
    start_col = 5
    for i, header in enumerate(headers, start=start_col):
        cell = ws.cell(start_row, i)
        cell.value = header
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

    for r_idx, rec in enumerate(discrepancies, start=start_row + 1):
        for c_idx, header in enumerate(headers, start=start_col):
            cell = ws.cell(r_idx, c_idx)

            if rec.get("Column") == "% to Net Assets" and header == "Master Value":
                cell.value = format_percent_for_index(rec.get("Master Value"), "MASTER")

            elif rec.get("Column") == "% to Net Assets" and header == "SEBI Value":
                cell.value = format_percent_for_index(rec.get("SEBI Value"), "SEBI")

            elif rec.get("Column") == "Quantity" and header == "Master Value":
                # For Quantity mismatches, show the same processed value used for comparison.
                cell.value = format_quantity_for_index(rec.get("Master Value"), "MASTER")

            elif rec.get("Column") == "Quantity" and header == "SEBI Value":
                cell.value = format_quantity_for_index(rec.get("SEBI Value"), "SEBI")

            else:
                cell.value = rec.get(header)

            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    widths = {
        5: 14, 6: 16, 7: 18, 8: 42, 9: 30, 10: 20, 11: 20, 12: 42, 13: 14
    }
    for col, width in widths.items():
        ws.column_dimensions[ws.cell(1, col).column_letter].width = width

    ws.freeze_panes = "E4"
    ws.auto_filter.ref = f"E3:M{max(start_row + 1, start_row + len(discrepancies))}"


def main():
    source_xlsx = resolve_source_xlsx()

    wb = load_workbook(source_xlsx)
    master_df = pd.read_excel(MASTER_FILE)
    trial_balance_df = read_trial_balance_df()

    discrepancies = []
    processed = []
    skipped = []

    for sheet_name in wb.sheetnames:
        if sheet_name == "Index":
            continue
        if scheme_dict.get(sheet_name) != "EQUITY":
            continue

        ws = wb[sheet_name]
        try:
            compare_scheme(ws, sheet_name, master_df, trial_balance_df, discrepancies)
            processed.append(sheet_name)
        except Exception as exc:
            skipped.append((sheet_name, str(exc)))
            add_discrepancy(
                discrepancies,
                sheet_name,
                "Sheet Error",
                "",
                "",
                "",
                "",
                "",
                f"Could not process sheet: {exc}",
                None,
            )

    write_index_output(wb, discrepancies)
    wb.save(OUTPUT_FILE)

    print(f"Processed equity schemes: {len(processed)}")
    print(f"Discrepancies written: {len(discrepancies)}")
    if skipped:
        print("Skipped sheets:")
        for item in skipped:
            print(item)
    print(f"Output file: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()