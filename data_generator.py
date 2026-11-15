"""
data_generator.py  --  STEP 1 of the FP&A Variance Dashboard
==============================================================
Generates two "ERP-style" CSV files that we will clean, analyse and
visualise in later steps:

    1. erp_opex_ledger.csv  -> Monthly Operating Expense: Budget vs Actual
    2. sales_pvm_ledger.csv -> Quarterly Sales: Price-Volume-Mix (PVM) data

HOW TO RUN:   python data_generator.py

VARIANCE SIGN CONVENTION (used in the sanity checks):
    Revenue : Actual > Budget  = Favorable (F)    | Actual < Budget = Unfavorable (U)
    Expense : Actual > Budget  = Unfavorable (U)  | Actual < Budget = Favorable (F)
"""

import numpy as np
import pandas as pd

# ----------------------------------------------------------------------
# 0. SETUP
# ----------------------------------------------------------------------
# A "seed" makes the random numbers repeatable: every time you run the
# script you get the SAME data. (Excel analogy: pasting values after
# RAND() so the numbers stop changing every time you press F9.)
rng = np.random.default_rng(42)


# ----------------------------------------------------------------------
# 1. HELPER: format a number the Indian way  (125000 -> "₹1,25,000")
# ----------------------------------------------------------------------
# Indian grouping: the last 3 digits form one group, then groups of 2.
# Real ERP exports often store amounts as TEXT like this, which breaks
# calculations until cleaned -- we will fix this in Step 2.
def to_indian_format(number):
    digits = str(int(number))
    if len(digits) <= 3:
        return "₹" + digits
    last_three = digits[-3:]      # e.g. "000"
    rest = digits[:-3]            # e.g. "125"
    groups = []
    while len(rest) > 2:          # peel off groups of 2 from the right
        groups.insert(0, rest[-2:])
        rest = rest[:-2]
    if rest:
        groups.insert(0, rest)
    return "₹" + ",".join(groups + [last_three])


# ======================================================================
# PART A: ERP OPEX LEDGER
# ======================================================================

# ----------------------------------------------------------------------
# A1. CHART OF ACCOUNTS  (the "master data" behind the ledger)
# ----------------------------------------------------------------------
# Format: (Account_Code, Expense_Line_Item, Cost_Type, Monthly_Base_Budget_INR)
chart_of_accounts = {
    "Cloud Infrastructure": [
        (5101, "Compute & Hosting",          "Variable", 1800000),
        (5102, "Data Storage & Backup",      "Variable",  600000),
        (5103, "Cloud Support & Licences",   "Fixed",     400000),
    ],
    "R&D & Engineering": [
        (5201, "Engineering Salaries",       "Fixed",    3200000),
        (5202, "Software Tools & Licences",  "Fixed",     350000),
        (5203, "Contract Developers",        "Variable",  900000),
    ],
    "Performance Marketing": [
        (5301, "Paid Search & Social Ads",   "Variable", 2000000),
        (5302, "Affiliate Commissions",      "Variable",  700000),
        (5303, "Marketing Analytics Tools",  "Fixed",     250000),
    ],
    "G&A / Payroll": [
        (5401, "Admin & Finance Salaries",   "Fixed",    1500000),
        (5402, "Office Rent & Utilities",    "Fixed",     600000),
        (5403, "Professional & Legal Fees",  "Variable",  300000),
    ],
}

# ----------------------------------------------------------------------
# A2. OVERSPEND SCENARIO  (the "story" hidden in the data)
# ----------------------------------------------------------------------
# (Department, Quarter) -> (min overspend %, max overspend %)
# Only Cloud Infrastructure and Performance Marketing overspend, and
# only in Q3 and Q4. Variable costs take the full hit; Fixed costs only
# a quarter of it (a fixed contract can't swing as much as ad spend).
overspend_ranges = {
    ("Cloud Infrastructure", 3):   (0.12, 0.20),
    ("Cloud Infrastructure", 4):   (0.18, 0.30),
    ("Performance Marketing", 3):  (0.15, 0.25),
    ("Performance Marketing", 4):  (0.25, 0.40),
}

# ----------------------------------------------------------------------
# A3. BUILD THE CLEAN LEDGER (one row per month x department x line item)
# ----------------------------------------------------------------------
months = pd.date_range("2026-01-01", periods=12, freq="MS")  # month starts
opex_rows = []

for month in months:
    quarter = (month.month - 1) // 3 + 1                     # Jan-Mar=1 ... Oct-Dec=4

    for department, line_items in chart_of_accounts.items():
        for account_code, line_item, cost_type, base_budget in line_items:

            # Budget grows ~1% each month (planned business growth)
            budget = base_budget * (1 + 0.01 * (month.month - 1))
            budget = round(budget, -3)                       # round to nearest 1,000

            # Everyday noise: Fixed costs barely move, Variable costs wobble more
            if cost_type == "Fixed":
                noise = rng.normal(0, 0.01)
            else:
                noise = rng.normal(0, 0.05)

            # Extra overspend only for the Q3/Q4 "problem" departments
            overspend = 0.0
            if (department, quarter) in overspend_ranges:
                low, high = overspend_ranges[(department, quarter)]
                overspend = rng.uniform(low, high)
                if cost_type == "Fixed":
                    overspend = overspend * 0.25

            actual = budget * (1 + noise + overspend)
            actual = round(actual, -2)                       # round to nearest 100

            opex_rows.append({
                "Date": month,
                "Department": department,
                "Account_Code": account_code,
                "Expense_Line_Item": line_item,
                "Cost_Type": cost_type,
                "Budget_INR": int(budget),
                "Actual_INR": int(actual),
            })

df_opex = pd.DataFrame(opex_rows)
df_opex = df_opex.sort_values(["Date", "Department", "Account_Code"]).reset_index(drop=True)

# ----------------------------------------------------------------------
# A4. SANITY CHECK ON THE CLEAN DATA (before we add the mess)
# ----------------------------------------------------------------------
# EXCEL EQUIVALENT: Pivot Table -> Rows: Department, Quarter | Values: Sum of Budget, Sum of Actual
df_check = df_opex.copy()
df_check["Quarter"] = df_check["Date"].dt.quarter
opex_summary = df_check.groupby(["Department", "Quarter"])[["Budget_INR", "Actual_INR"]].sum().reset_index()

# EXCEL EQUIVALENT: =Actual - Budget   and   =(Actual - Budget) / Budget
opex_summary["Variance_INR"] = opex_summary["Actual_INR"] - opex_summary["Budget_INR"]
opex_summary["Variance_Pct"] = (opex_summary["Variance_INR"] / opex_summary["Budget_INR"] * 100).round(1)

# EXCEL EQUIVALENT: =IF(Actual > Budget, "U", "F")   <- EXPENSE logic
opex_summary["F_or_U"] = np.where(opex_summary["Actual_INR"] > opex_summary["Budget_INR"], "U", "F")

print("=== OPEX SANITY CHECK (clean data, by Department & Quarter) ===")
print(opex_summary.to_string(index=False))

# ----------------------------------------------------------------------
# A5. INJECT REALISTIC ERP MESSINESS
# ----------------------------------------------------------------------
# We shuffle the row positions once, then hand out DIFFERENT slices to
# each type of mess so no row gets damaged twice.
shuffled_rows = rng.permutation(len(df_opex))

rows_with_space_in_dept   = shuffled_rows[0:8]     # trailing spaces in Department
rows_with_text_actual     = shuffled_rows[8:16]    # Actual_INR stored as "₹1,25,000"
rows_with_text_budget     = shuffled_rows[16:21]   # Budget_INR stored as "₹1,25,000"
rows_with_missing_actual  = shuffled_rows[21:23]   # Actual_INR is blank
rows_with_missing_type    = shuffled_rows[23:24]   # Cost_Type is blank

# (a) Trailing spaces: "Cloud Infrastructure" becomes "Cloud Infrastructure  "
# Later, a groupby would treat these as DIFFERENT departments!
df_opex.loc[rows_with_space_in_dept, "Department"] = (
    df_opex.loc[rows_with_space_in_dept, "Department"] + "  "
)

# (b) Numbers stored as text. First switch the columns to 'object' dtype
# (Pandas' "can hold anything" type) so strings can sit next to numbers.
df_opex["Budget_INR"] = df_opex["Budget_INR"].astype(object)
df_opex["Actual_INR"] = df_opex["Actual_INR"].astype(object)

df_opex.loc[rows_with_text_actual, "Actual_INR"] = (
    df_opex.loc[rows_with_text_actual, "Actual_INR"].apply(to_indian_format)
)
df_opex.loc[rows_with_text_budget, "Budget_INR"] = (
    df_opex.loc[rows_with_text_budget, "Budget_INR"].apply(to_indian_format)
)

# (c) Missing values (np.nan = an empty cell in the CSV)
df_opex.loc[rows_with_missing_actual, "Actual_INR"] = np.nan
df_opex.loc[rows_with_missing_type, "Cost_Type"] = np.nan

# Write dates as plain text YYYY-MM-DD, then save.
# utf-8-sig lets Excel display the ₹ symbol correctly when you double-click the CSV.
df_opex["Date"] = df_opex["Date"].dt.strftime("%Y-%m-%d")
df_opex.to_csv("erp_opex_ledger.csv", index=False, encoding="utf-8-sig")


# ======================================================================
# PART B: SALES PVM LEDGER
# ======================================================================

# ----------------------------------------------------------------------
# B1. BUDGET ASSUMPTIONS per tier
# ----------------------------------------------------------------------
# Format: Tier -> (Q1..Q4 budget units list, budget price per unit in INR)
sales_budget = {
    "Basic Tier":         ([4000, 4200, 4400, 4600],  2500),
    "Pro Suite":          ([1200, 1300, 1400, 1500], 12000),
    "Enterprise License": ([ 150,  160,  170,  180], 150000),
}

# ----------------------------------------------------------------------
# B2. ACTUAL PERFORMANCE SCENARIO per tier and quarter
# ----------------------------------------------------------------------
# (units % vs budget, price % vs budget) -- fixed numbers so the story is clear.
# Enterprise License: Q3 = deep discounting (-18% price) that pulls in a
# volume surge (+35% units). Q4 = discounts ease, volume stays strong.
# Basic and Pro: small random noise around budget (set in the loop below).
enterprise_story = {
    1: (0.02,  0.00),
    2: (0.03, -0.01),
    3: (0.35, -0.18),
    4: (0.12, -0.05),
}

quarter_labels = ["Q1 2026", "Q2 2026", "Q3 2026", "Q4 2026"]
sales_rows = []

for tier, (budget_units_list, budget_price) in sales_budget.items():
    for q_index in range(4):
        budget_units = budget_units_list[q_index]

        if tier == "Enterprise License":
            units_change, price_change = enterprise_story[q_index + 1]
        else:
            units_change = rng.normal(0.01, 0.03)    # small noise, slight upside
            price_change = rng.normal(0.00, 0.015)

        actual_units = int(round(budget_units * (1 + units_change)))
        actual_price = int(round(budget_price * (1 + price_change), -1))  # nearest 10

        sales_rows.append({
            "Quarter": quarter_labels[q_index],
            "Product_Tier": tier,
            "Budget_Units": budget_units,
            "Actual_Units": actual_units,
            "Budget_Price_INR": budget_price,
            "Actual_Price_INR": actual_price,
        })

df_sales = pd.DataFrame(sales_rows)

# ----------------------------------------------------------------------
# B3. SANITY CHECK: revenue by tier (temporary columns, NOT saved)
# ----------------------------------------------------------------------
# EXCEL EQUIVALENT: =Units * Price  in two helper columns
df_sales_check = df_sales.copy()
df_sales_check["Budget_Revenue"] = df_sales_check["Budget_Units"] * df_sales_check["Budget_Price_INR"]
df_sales_check["Actual_Revenue"] = df_sales_check["Actual_Units"] * df_sales_check["Actual_Price_INR"]

# EXCEL EQUIVALENT: =IF(Actual > Budget, "F", "U")   <- REVENUE logic (opposite of expenses!)
df_sales_check["Revenue_F_or_U"] = np.where(
    df_sales_check["Actual_Revenue"] > df_sales_check["Budget_Revenue"], "F", "U"
)
# Price effect on its own: did we sell at a higher or lower price than planned?
df_sales_check["Price_F_or_U"] = np.where(
    df_sales_check["Actual_Price_INR"] >= df_sales_check["Budget_Price_INR"], "F", "U"
)

print()
print("=== SALES SANITY CHECK (revenue by tier & quarter) ===")
print(df_sales_check[["Quarter", "Product_Tier", "Budget_Revenue", "Actual_Revenue",
                      "Revenue_F_or_U", "Price_F_or_U"]].to_string(index=False))

df_sales.to_csv("sales_pvm_ledger.csv", index=False, encoding="utf-8-sig")

# ----------------------------------------------------------------------
# DONE
# ----------------------------------------------------------------------
print()
print("Created: erp_opex_ledger.csv  ->", len(df_opex), "rows (messy on purpose!)")
print("Created: sales_pvm_ledger.csv ->", len(df_sales), "rows")