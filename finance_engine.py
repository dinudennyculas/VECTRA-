"""
finance_engine.py  --  the calculation engine behind VECTRA
=============================================================
Engineered & Architected by Dinu Denny Culas

Every public function follows the same four-stage pipeline:

    (a) INPUT VALIDATION  - can the file be read, and does it have the required columns?
    (b) DATA CLEANING     - trim text, convert amounts, fill blanks, stop on unusable values
    (c) VECTORIZED MATH   - whole-column pandas operations, no row-by-row loops
    (d) OUTPUT            - return a tidy DataFrame

Public functions
    variance_flag(variance, is_expense)  -> "F" / "U" / "On Track" (single number or whole column)
    load_and_clean_opex(source)          -> cleaned OPEX ledger with variance columns
    get_department_summary(df_clean)     -> Department x Quarter totals with variances
    calculate_pvm_analysis(source)       -> Price-Volume-Mix (PVM) decomposition per sales row
    get_template_csv_bytes(kind)         -> blank ERP template as CSV bytes
    get_template_excel_bytes(kind)       -> blank ERP template as Excel bytes

A "source" is a file path (str or Path) or a Streamlit UploadedFile: anything pandas can read
that carries a .name ending in .csv or .xlsx.

ENTERPRISE VARIANCE ACCOUNTING CONVENTION
    Expenses : Actual > Budget = Unfavorable (U)   |  Actual < Budget = Favorable (F)
    Revenue  : Actual > Budget = Favorable (F)     |  Actual < Budget = Unfavorable (U)
    Equal    : "On Track"

Problems with a file raise ValueError with a plain-English message, so the app can show it
with st.error() instead of a traceback.

Run the self-test at the bottom with:   python finance_engine.py
"""

import io
import os
import re

import numpy as np
import pandas as pd

# ======================================================================
# THE CONTRACT: the exact column headers every ERP export must contain
# ======================================================================
OPEX_REQUIRED_COLUMNS = ["Date", "Department", "Account_Code", "Expense_Line_Item",
                         "Cost_Type", "Budget_INR", "Actual_INR"]

SALES_REQUIRED_COLUMNS = ["Quarter", "Product_Tier", "Budget_Units", "Actual_Units",
                          "Budget_Price_INR", "Actual_Price_INR"]

# Example rows for the downloadable templates (they load without errors if uploaded unchanged)
OPEX_TEMPLATE_ROWS = [
    ["2026-01-01", "Cloud Infrastructure", 5101, "Compute & Hosting", "Variable", 1800000, 1827400],
    ["2026-01-01", "G&A / Payroll", 5401, "Admin & Finance Salaries", "Fixed", 1500000, 1487200],
    ["2026-02-01", "Cloud Infrastructure", 5101, "Compute & Hosting", "Variable", 1818000, "₹19,10,500"],
    ["2026-02-01", "G&A / Payroll", 5401, "Admin & Finance Salaries", "Fixed", 1515000, 1522300],
    ["2026-03-01", "Cloud Infrastructure", 5101, "Compute & Hosting", "Variable", 1836000, 1869900],
    ["2026-03-01", "G&A / Payroll", 5401, "Admin & Finance Salaries", "Fixed", 1530000, 1530000],
]

SALES_TEMPLATE_ROWS = [
    ["Q1 2026", "Basic Tier", 4000, 4080, 2500, 2490],
    ["Q1 2026", "Pro Suite", 1200, 1230, 12000, 12150],
    ["Q1 2026", "Enterprise License", 150, 153, 150000, 150000],
    ["Q2 2026", "Basic Tier", 4200, 4310, 2500, 2530],
    ["Q2 2026", "Pro Suite", 1300, 1352, 12000, 11900],
    ["Q2 2026", "Enterprise License", 160, 165, 150000, 148500],
]

OPEX_COLUMN_GUIDE = [
    ["Date", "Yes", "Any date inside the month: a real Excel date or text like 2026-01-01.", "2026-01-01"],
    ["Department", "Yes", "Cost centre. Extra spaces are trimmed; blanks become 'Unassigned'.", "Cloud Infrastructure"],
    ["Account_Code", "Yes", "GL account code from the ERP (reference only).", "5101"],
    ["Expense_Line_Item", "Yes", "GL account / expense description.", "Compute & Hosting"],
    ["Cost_Type", "Yes", "Fixed or Variable. A blank is treated as Variable.", "Variable"],
    ["Budget_INR", "Yes", "Budget in INR. Plain numbers or text like ₹1,25,000 both work. Must not be blank.", "1800000"],
    ["Actual_INR", "Yes", "Actual in INR. A blank is filled with the Budget (zero variance).", "1827400"],
]

SALES_COLUMN_GUIDE = [
    ["Quarter", "Yes", "Must contain Q1, Q2, Q3 or Q4 (a year is optional).", "Q1 2026"],
    ["Product_Tier", "Yes", "Product, tier or SKU group.", "Enterprise License"],
    ["Budget_Units", "Yes", "Budgeted units sold.", "150"],
    ["Actual_Units", "Yes", "Actual units sold.", "153"],
    ["Budget_Price_INR", "Yes", "Budgeted selling price per unit in INR.", "150000"],
    ["Actual_Price_INR", "Yes", "Actual average selling price per unit in INR.", "150000"],
]


# ======================================================================
# SMALL BUILDING BLOCKS (each is used by more than one function)
# ======================================================================
def read_erp_file(source):
    """Read a .csv or .xlsx file given as a path or an uploaded file object."""
    if isinstance(source, (str, os.PathLike)):
        file_name = str(source)
    else:
        file_name = getattr(source, "name", "")
    extension = os.path.splitext(file_name)[1].lower()
    if extension not in (".csv", ".xlsx", ".xlsm"):
        raise ValueError(f"Unsupported file type '{extension}'. Please upload a .csv or .xlsx file.")

    if hasattr(source, "seek"):
        source.seek(0)          # uploaded files may already have been read once

    try:
        if extension == ".csv":
            try:
                return pd.read_csv(source)
            except UnicodeDecodeError:
                if hasattr(source, "seek"):
                    source.seek(0)
                return pd.read_csv(source, encoding="latin-1")     # common in older ERP exports
        return pd.read_excel(source, engine="openpyxl")             # first sheet of the workbook
    except ImportError as error:
        raise ValueError("Reading Excel files needs openpyxl. Run: pip install openpyxl") from error
    except Exception as error:
        raise ValueError(f"Could not read '{os.path.basename(file_name)}' ({error}). "
                         "Please check it is a valid .csv or .xlsx export.") from error


def check_required_columns(df, required_columns, ledger_name):
    """Match headers ignoring case and spaces, then confirm every required column exists."""
    known_names = {name.lower(): name for name in required_columns}
    renames = {}
    for column in df.columns:
        lookup_key = str(column).strip().lower().replace(" ", "_")
        # pandas renames a repeated header such as Budget_INR to Budget_INR.1; strip that suffix
        # so the repeat is caught by the duplicate check below instead of being silently ignored
        lookup_key = re.sub(r"\.\d+$", "", lookup_key)
        if lookup_key in known_names:
            renames[column] = known_names[lookup_key]
    df = df.rename(columns=renames)

    if df.columns.duplicated().any():
        raise ValueError(f"{ledger_name} has duplicate column headers. Each header must appear once.")

    missing = [name for name in required_columns if name not in df.columns]
    if missing:
        raise ValueError(f"{ledger_name} is missing required column(s): {', '.join(missing)}. "
                         f"Columns found in your file: {', '.join(map(str, df.columns))}. "
                         "Download a blank ERP template to see the exact column headers.")
    return df


def spreadsheet_rows(df, mask, limit=5):
    """First few spreadsheet row numbers (header = row 1) where mask is True, for error messages."""
    return (df.index[mask] + 2).tolist()[:limit]


def clean_amounts(series):
    """Turn text such as '₹1,25,000' into the number 125000.0 (unreadable cells become NaN)."""
    text = series.astype(str)
    for symbol in ["₹", "INR", "Rs.", "Rs", ","]:
        text = text.str.replace(symbol, "", regex=False)
    return pd.to_numeric(text.str.strip(), errors="coerce").astype(float)


def clean_text_column(df, column, fill_value):
    """Trim spaces and fill blank cells. Returns masks of the rows that were trimmed and filled."""
    raw_text = df[column].fillna("").astype(str)
    trimmed = raw_text.str.strip()
    rows_trimmed = raw_text != trimmed
    rows_blank = trimmed == ""
    df[column] = trimmed.where(~rows_blank, fill_value)
    return rows_trimmed, rows_blank


def add_fix_note(df, mask, note):
    """Append a note to the audit-trail column on every row where mask is True."""
    df.loc[mask, "Data_Fix_Note"] = df.loc[mask, "Data_Fix_Note"] + note + "; "


def variance_flag(variance, is_expense):
    """Return "F", "U" or "On Track" for a single number or a whole column of variances.

    A variance is Actual minus Budget, so a positive variance means Actual is above Budget.
    """
    if is_expense:
        positive_flag, negative_flag = "U", "F"
    else:
        positive_flag, negative_flag = "F", "U"

    flags = np.select([variance > 0, variance < 0], [positive_flag, negative_flag], default="On Track")
    if flags.ndim == 0:
        return str(flags)
    return flags


# ======================================================================
# FUNCTION 1: LOAD AND CLEAN THE OPEX LEDGER
# ======================================================================
def load_and_clean_opex(source):
    print("--- CLEANING OPEX LEDGER ---")

    # ------------------------------------------------------------------
    # (a) INPUT VALIDATION
    # ------------------------------------------------------------------
    # EXCEL EQUIVALENT: Data > Get Data > From Text/CSV (or From Workbook)
    df = read_erp_file(source)
    df = check_required_columns(df, OPEX_REQUIRED_COLUMNS, "OPEX ledger")
    if df.empty:
        raise ValueError("The OPEX ledger has column headers but no data rows.")
    print(f"Rows loaded: {len(df)}")

    # ------------------------------------------------------------------
    # (b) DATA CLEANING   (every repaired row gets a note in Data_Fix_Note)
    # ------------------------------------------------------------------
    df["Data_Fix_Note"] = ""

    # Dates.  EXCEL EQUIVALENT: =DATEVALUE(A2), then look for #VALUE! errors
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    unreadable_dates = df["Date"].isna()
    if unreadable_dates.any():
        raise ValueError(f"OPEX ledger: {unreadable_dates.sum()} row(s) have a Date that could not be read "
                         f"(spreadsheet rows {spreadsheet_rows(df, unreadable_dates)}). "
                         "Use real dates or text like 2026-01-31.")

    # Text columns.  EXCEL EQUIVALENT: =IF(TRIM(B2)="", "Unassigned", TRIM(B2))
    # "Cloud  " and "Cloud" would otherwise become two departments in a Pivot Table, and a
    # blank department would silently vanish from every total.
    dept_trimmed, dept_blank = clean_text_column(df, "Department", "Unassigned")
    item_trimmed, item_blank = clean_text_column(df, "Expense_Line_Item", "Unspecified")
    add_fix_note(df, dept_trimmed, "Trimmed Department spaces")
    add_fix_note(df, item_trimmed, "Trimmed line-item spaces")
    add_fix_note(df, dept_blank, "Missing Department set to Unassigned")
    add_fix_note(df, item_blank, "Missing line item set to Unspecified")
    print(f"Trimmed spaces: {dept_trimmed.sum()} Department rows, {item_trimmed.sum()} Expense_Line_Item rows")

    # Amounts.  EXCEL EQUIVALENT: Find & Replace "₹" and "," with nothing, then =VALUE(cell)
    converted_counts = {}
    for column in ["Budget_INR", "Actual_INR"]:
        cleaned = clean_amounts(df[column])
        was_text = pd.to_numeric(df[column], errors="coerce").isna() & cleaned.notna()
        add_fix_note(df, was_text, column.replace("_INR", "") + " text converted to number")
        converted_counts[column] = was_text.sum()
        df[column] = cleaned
    print(f"Converted text amounts to numbers: {converted_counts['Budget_INR']} Budget rows, "
          f"{converted_counts['Actual_INR']} Actual rows")

    # A missing Budget cannot be guessed: without it there is no variance to measure.
    budget_blank = df["Budget_INR"].isna()
    if budget_blank.any():
        raise ValueError(f"OPEX ledger: {budget_blank.sum()} row(s) have a blank or non-numeric Budget_INR "
                         f"(spreadsheet rows {spreadsheet_rows(df, budget_blank)}). "
                         "A budget is needed to calculate variance, so please fill these in.")

    # Cost_Type must end up as "Fixed" or "Variable".  EXCEL EQUIVALENT: =PROPER(TRIM(E2))
    cost_text = df["Cost_Type"].fillna("").astype(str)
    cost_clean = cost_text.str.strip().str.title()
    cost_blank = cost_clean == ""
    cost_recased = (cost_text != cost_clean) & ~cost_blank
    cost_invalid = ~cost_clean.isin(["Fixed", "Variable"]) & ~cost_blank
    if cost_invalid.any():
        bad_values = sorted(cost_clean[cost_invalid].unique())[:5]
        raise ValueError(f"OPEX ledger: Cost_Type must be 'Fixed' or 'Variable'. Found: {', '.join(bad_values)}.")
    add_fix_note(df, cost_recased, "Standardised Cost_Type text")
    add_fix_note(df, cost_blank, "Missing Cost_Type set to Variable")

    # Blank cells.  ASSUMPTIONS:
    #   a blank Actual means "spent exactly the Budget" (variance 0, so it can understate a real overspend)
    #   a blank Cost_Type is "Variable" (the conservative choice: the cost could have been controlled)
    # EXCEL EQUIVALENT: =IF(ISBLANK(Actual), Budget, Actual)
    actual_blank = df["Actual_INR"].isna()
    df["Actual_INR"] = df["Actual_INR"].fillna(df["Budget_INR"])
    df["Cost_Type"] = cost_clean.where(~cost_blank, "Variable")
    add_fix_note(df, actual_blank, "Missing Actual filled with Budget")
    print(f"Filled missing Actual_INR with Budget_INR: {actual_blank.sum()} rows")
    print(f"Filled missing Cost_Type with 'Variable':  {cost_blank.sum()} rows")

    # ------------------------------------------------------------------
    # (c) VECTORIZED FINANCIAL MATH
    # ------------------------------------------------------------------
    # EXCEL EQUIVALENT: =Actual - Budget      and      =(Actual - Budget) / ABS(Budget) * 100
    # Dividing by |Budget| keeps the sign of the percentage aligned with the variance.
    # A zero budget becomes NaN (not a number) instead of causing a divide-by-zero.
    df["Variance_INR"] = df["Actual_INR"] - df["Budget_INR"]
    budget_magnitude = df["Budget_INR"].abs().replace(0, np.nan)
    df["Variance_Pct"] = (df["Variance_INR"] / budget_magnitude * 100).round(2)

    # EXCEL EQUIVALENT: =IF(Variance>0,"U",IF(Variance<0,"F","On Track"))   (expense rule)
    df["Variance_Flag"] = variance_flag(df["Variance_INR"], is_expense=True)

    # EXCEL EQUIVALENT: ="Q" & ROUNDUP(MONTH(A2)/3, 0)
    df["Quarter"] = "Q" + df["Date"].dt.quarter.astype(str)

    # ------------------------------------------------------------------
    # (d) OUTPUT
    # ------------------------------------------------------------------
    df["Data_Fix_Note"] = df["Data_Fix_Note"].str.rstrip("; ")
    print(f"Rows with at least one auto-fix: {(df['Data_Fix_Note'] != '').sum()} of {len(df)}")
    print(f"Clean ledger ready: {len(df)} rows\n")
    return df


# ======================================================================
# FUNCTION 2: DEPARTMENT x QUARTER SUMMARY
# ======================================================================
def get_department_summary(df_clean):

    # (a) INPUT VALIDATION
    needed_columns = ["Department", "Quarter", "Budget_INR", "Actual_INR"]
    missing = [name for name in needed_columns if name not in df_clean.columns]
    if missing:
        raise ValueError(f"get_department_summary needs a cleaned OPEX table. Missing: {', '.join(missing)}.")

    # (b) DATA CLEANING: the input is already clean, so go straight to grouping.
    # EXCEL EQUIVALENT: Pivot Table, Rows = Department and Quarter, Values = Sum of Budget and Actual
    summary = df_clean.groupby(["Department", "Quarter"], as_index=False)[["Budget_INR", "Actual_INR"]].sum()

    # (c) VECTORIZED FINANCIAL MATH
    # Variance % is calculated on the TOTALS, never as an average of row percentages.
    summary["Variance_INR"] = summary["Actual_INR"] - summary["Budget_INR"]
    budget_magnitude = summary["Budget_INR"].abs().replace(0, np.nan)
    summary["Variance_Pct"] = (summary["Variance_INR"] / budget_magnitude * 100).round(1)
    summary["Variance_Flag"] = variance_flag(summary["Variance_INR"], is_expense=True)

    # (d) OUTPUT
    return summary


# ======================================================================
# FUNCTION 3: SALES PRICE-VOLUME-MIX (PVM) DECOMPOSITION
# ======================================================================
def calculate_pvm_analysis(source):

    # ------------------------------------------------------------------
    # (a) INPUT VALIDATION
    # ------------------------------------------------------------------
    df = read_erp_file(source)
    df = check_required_columns(df, SALES_REQUIRED_COLUMNS, "Sales PVM ledger")
    if df.empty:
        raise ValueError("The Sales PVM ledger has column headers but no data rows.")

    # ------------------------------------------------------------------
    # (b) DATA CLEANING
    # ------------------------------------------------------------------
    clean_text_column(df, "Product_Tier", "Unassigned")

    numeric_columns = ["Budget_Units", "Actual_Units", "Budget_Price_INR", "Actual_Price_INR"]
    for column in numeric_columns:
        df[column] = clean_amounts(df[column])
    unusable = df[numeric_columns].isna()
    if unusable.any().any():
        bad_columns = unusable.columns[unusable.any()].tolist()
        raise ValueError(f"Sales PVM ledger has blank or non-numeric values in: {', '.join(bad_columns)} "
                         f"(spreadsheet rows {spreadsheet_rows(df, unusable.any(axis=1))}).")

    # "Q3 2026", "FY26-Q3" and "q3" all become "Q3".  EXCEL EQUIVALENT: =MID(A2, FIND("Q", UPPER(A2)), 2)
    quarter_text = df["Quarter"].fillna("").astype(str).str.strip()
    quarter_number = quarter_text.str.extract(r"[Qq]\s*([1-4])")[0]
    if quarter_number.isna().any():
        raise ValueError("Sales PVM ledger: Quarter values must contain Q1, Q2, Q3 or Q4 (for example "
                         f"'Q3 2026'). Please check spreadsheet rows {spreadsheet_rows(df, quarter_number.isna())}.")
    df["Quarter"] = quarter_text
    df["Quarter_Short"] = "Q" + quarter_number

    # ------------------------------------------------------------------
    # (c) VECTORIZED FINANCIAL MATH
    # ------------------------------------------------------------------
    # EXCEL EQUIVALENT: =Units * Price
    df["Budget_Revenue"] = df["Budget_Units"] * df["Budget_Price_INR"]
    df["Actual_Revenue"] = df["Actual_Units"] * df["Actual_Price_INR"]
    df["Total_Sales_Variance"] = df["Actual_Revenue"] - df["Budget_Revenue"]

    # PRICE VARIANCE: the price gap, measured on the units actually sold.
    # EXCEL EQUIVALENT: =(Actual_Price - Budget_Price) * Actual_Units
    df["Price_Variance"] = (df["Actual_Price_INR"] - df["Budget_Price_INR"]) * df["Actual_Units"]

    # VOLUME VARIANCE: the unit gap, measured at the budgeted price so price changes cannot leak in.
    # EXCEL EQUIVALENT: =(Actual_Units - Budget_Units) * Budget_Price
    df["Volume_Variance"] = (df["Actual_Units"] - df["Budget_Units"]) * df["Budget_Price_INR"]

    # Price + Volume = Total for every row. (A separate Mix variance only appears
    # when several products are combined into one total.)
    # EXCEL EQUIVALENT: =IF(Variance>0,"F",IF(Variance<0,"U","On Track"))   (revenue rule)
    df["Price_Flag"] = variance_flag(df["Price_Variance"], is_expense=False)
    df["Volume_Flag"] = variance_flag(df["Volume_Variance"], is_expense=False)

    # ------------------------------------------------------------------
    # (d) OUTPUT
    # ------------------------------------------------------------------
    return df


# ======================================================================
# BLANK ERP TEMPLATES
# ======================================================================
def get_template_dataframe(kind):
    """kind is 'opex' or 'sales'."""
    if kind == "opex":
        return pd.DataFrame(OPEX_TEMPLATE_ROWS, columns=OPEX_REQUIRED_COLUMNS)
    return pd.DataFrame(SALES_TEMPLATE_ROWS, columns=SALES_REQUIRED_COLUMNS)


def get_template_csv_bytes(kind):
    # utf-8-sig lets Excel display the ₹ symbol correctly when the CSV is double-clicked
    return get_template_dataframe(kind).to_csv(index=False).encode("utf-8-sig")


def get_template_excel_bytes(kind):
    if kind == "opex":
        guide_rows, data_sheet_name = OPEX_COLUMN_GUIDE, "OPEX_Data"
    else:
        guide_rows, data_sheet_name = SALES_COLUMN_GUIDE, "Sales_Data"
    guide = pd.DataFrame(guide_rows, columns=["Column", "Required", "What to enter", "Example"])

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        # The data sheet must come first: the engine reads the first sheet of a workbook.
        get_template_dataframe(kind).to_excel(writer, sheet_name=data_sheet_name, index=False)
        guide.to_excel(writer, sheet_name="Instructions", index=False)
        for sheet in writer.sheets.values():
            for column_cells in sheet.columns:
                longest_text = max(len(str(cell.value)) for cell in column_cells)
                sheet.column_dimensions[column_cells[0].column_letter].width = min(longest_text + 3, 70)
    return buffer.getvalue()


# ======================================================================
# SELF-TEST: runs only with `python finance_engine.py`, never when the app imports this file
# ======================================================================
if __name__ == "__main__":
    pd.set_option("display.float_format", "{:,.1f}".format)
    pd.set_option("display.width", 200)

    df_clean = load_and_clean_opex("erp_opex_ledger.csv")
    dept_summary = get_department_summary(df_clean)
    print("=== DEPARTMENT SUMMARY: Q3 and Q4 ===")
    print(dept_summary[dept_summary["Quarter"].isin(["Q3", "Q4"])].to_string(index=False))

    df_pvm = calculate_pvm_analysis("sales_pvm_ledger.csv")
    print("\n=== Q3 PRICE-VOLUME-MIX (PVM) BREAKDOWN ===")
    q3_columns = ["Product_Tier", "Budget_Revenue", "Actual_Revenue", "Total_Sales_Variance",
                  "Price_Variance", "Price_Flag", "Volume_Variance", "Volume_Flag"]
    print(df_pvm.loc[df_pvm["Quarter_Short"] == "Q3", q3_columns].to_string(index=False))

    identity_gap = df_pvm["Total_Sales_Variance"] - df_pvm["Price_Variance"] - df_pvm["Volume_Variance"]
    print("\nPrice + Volume = Total for every row?", bool((identity_gap.abs() < 0.01).all()))

    print("\n=== TEMPLATE ROUND-TRIP ===")
    for kind in ["opex", "sales"]:
        for extension, make_bytes in [(".csv", get_template_csv_bytes), (".xlsx", get_template_excel_bytes)]:
            template_file = io.BytesIO(make_bytes(kind))
            template_file.name = f"template{extension}"
            if kind == "opex":
                rows = len(load_and_clean_opex(template_file))
            else:
                rows = len(calculate_pvm_analysis(template_file))
            print(f"{kind} template ({extension}) loaded: {rows} rows")

    print("\n=== MISSING-COLUMN MESSAGE ===")
    broken_file = io.BytesIO(b"Date,Department,Budget_INR\n2026-01-01,Cloud,1000\n")
    broken_file.name = "broken.csv"
    try:
        load_and_clean_opex(broken_file)
    except ValueError as error:
        print(error)