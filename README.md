# VECTRA | Enterprise Strategic FP&A & PVM Intelligence Engine

[![Live App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://vectra-fpa.streamlit.app/)
![Python](https://img.shields.io/badge/Python-3.10%2B-0A84FF?logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Vectorized-Pandas%20%2F%20NumPy-150458?logo=pandas&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-22C55E)
[![LinkedIn](https://img.shields.io/badge/Author-Dinu%20Denny%20Culas-0A66C2?logo=linkedin&logoColor=white)](https://www.linkedin.com/in/dinudenny/)

**Live Production App:** [https://vectra-fpa.streamlit.app/](https://vectra-fpa.streamlit.app/)  
**Engineered & Architected by:** [Dinu Denny Culas](https://www.linkedin.com/in/dinudenny/)

---

## Executive Overview

**VECTRA** is an open-source, deterministic Financial Planning & Analysis (FP&A) and Price-Volume-Mix (PVM) decomposition engine built in Python. 

In enterprise finance, top-line revenue growth frequently masks aggressive unit discounting, while raw ERP exports from SAP, Oracle, or NetSuite introduce schema breaks, currency string formatting (`₹10,00,000`), unclosed accruals, and zero-budget division traps. VECTRA automates the end-to-end workflow from dirty ERP ingestion to audit-verified variance bridges and flexible budget stress-testing.

---

## Core Capabilities & Financial Engineering

### 1. Automated Multi-Source ERP Ingestion & Sanitization (`finance_engine.py`)
* **Live `.csv` & `.xlsx` Ingestion:** Accepts custom general ledger (`erp_opex_ledger.csv`) and quarterly sales exports (`sales_pvm_ledger.csv`) alongside downloadable blank schema templates.
* **Deterministic Data Hygiene:** Automatically strips currency symbols (`₹`, commas), trims trailing department whitespace, imputes unclosed period actuals with baseline budget figures, defaults missing cost behaviors to conservative variable treatment, and guards against `ZeroDivisionError` on unbudgeted emergency expense lines (`Budget_INR = 0`).
* **Full Audit Trail:** Tracks and surfaces the exact row-level repair count and percentage directly in the executive KPI header.

### 2. OPEX Variance & Monthly Run-Rate Surveillance
* **Sign-Aware Variance Accounting:** Enforces strict corporate variance conventions where spending above budget is flagged **Unfavorable (`U`)** and spending below budget is flagged **Favorable (`F`)**, while revenue/profit above budget is flagged **Favorable (`F`)**.
* **Run-Rate Trajectory:** Plots continuous 12-month departmental spending trajectories (dotted budget vs. solid actuals) in Indian numbering conventions (`₹ Cr` / `₹ L`) to isolate seasonal cost spikes.

### 3. Price-Volume-Mix (PVM) Revenue Bridge
Decomposes total quarterly revenue variance across product tiers into pure operational drivers using standard management accounting formulations:
* **Price Variance:** `(Actual Price - Budget Price) × Actual Units`
* **Volume Variance:** `(Actual Units - Budget Units) × Budget Price`
* **Net Revenue Bridge:** `Price Variance + Volume Variance = Actual Revenue - Budget Revenue`

### 4. What-If Scenario Simulator (True Flexible Budgeting)
Interactive multi-variable stress testing across **Volume Growth Shock (%)**, **Price Realization Adjustment (%)**, and **Variable Cost Inflation (%)**:
* **Flexible vs. Static Budget Toggle:** When active, Variable OPEX flexes proportionally with unit volume before applying rate inflation (`Baseline Variable OPEX × (1 + Volume_Shock) × (1 + Inflation_Rate)`), while Fixed OPEX remains invariant—preventing artificial unfavorable cost variances during high-growth volume scenarios.

### 5. Glass-Box Calculation Audit (Tab 4)
* **Zero Black-Box Math:** Exposes the exact vectorized Pandas/NumPy functions powering the engine alongside a live single-row reconciliation verifier that compares hand-calculated math against engine output row by row.

---

## System Architecture

```text
VECTRA-/
├── .streamlit/
│   └── config.toml          # Cupertino Dark OLED theme configuration
├── app.py                   # Streamlit UI layer, Bento-grid CSS, & Plotly bridges
├── finance_engine.py        # Vectorized Pandas/NumPy sanitization & variance engine
├── data_generator.py        # Synthetic ERP dataset generator with intentional dirty rows
├── erp_opex_ledger.csv      # Sample 144-row Departmental OPEX General Ledger
├── sales_pvm_ledger.csv     # Sample Quarterly Product Tier PVM Revenue Ledger
└── requirements.txt         # Production cloud dependencies
```

---

## Local Installation & Quickstart

```bash
# 1. Clone the repository
git clone https://github.com/dinudennyculas/VECTRA-.git
cd VECTRA-

# 2. Install dependencies
python -m pip install -r requirements.txt

# 3. Launch VECTRA locally
python -m streamlit run app.py
```

---

## License & Attribution
Distributed under the **MIT License**. Engineered & Architected by [Dinu Denny Culas](https://www.linkedin.com/in/dinudenny/).
