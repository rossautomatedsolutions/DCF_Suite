# Single-Scenario DCF Valuation Template (Commercial-Ready)

## 1. Worksheet Tabs
1. **Inputs**
2. **Operating_Model**
3. **Valuation**
4. **Outputs**
5. **Notes**

---

## 2. Inputs per Tab (with Definitions)

### Inputs
**Section: General Assumptions**
- **Period_Start_Year**: First projection year (e.g., 2025).
- **Projection_Years**: Number of annual periods (e.g., 5).
- **Discount_Rate**: Flat annual discount rate.
- **Tax_Rate**: Effective tax rate used for NOPAT.

**Section: Revenue Assumptions**
- **Revenue_Base**: Revenue in the most recent historical year.
- **Revenue_Growth_Rate**: Annual growth rate applied to Revenue_Base.

**Section: Margin Assumptions**
- **EBIT_Margin**: EBIT as a % of revenue.
- **Depreciation_Pct_Revenue**: Depreciation as a % of revenue.

**Section: Reinvestment Proxy**
- **Capex_Pct_Revenue**: Capex as a % of revenue.
- **NWC_Pct_Revenue**: Net working capital as a % of revenue (used for change in NWC).

**Section: Terminal Value**
- **Terminal_Method**: Data validation with a single option: "Perpetuity Growth".
- **Terminal_Growth_Rate**: Long-run growth rate used in terminal value.

**Section: Per-Unit Value**
- **Units_Outstanding**: Generic units/shares outstanding.

> **Named Range Requirement**: Every input above is a named range.

### Notes
No inputs. Documentation-only.

---

## 3. Key Calculations per Tab

### Operating_Model
- **Revenue**: Prior year revenue × (1 + Revenue_Growth_Rate).
- **EBIT**: Revenue × EBIT_Margin.
- **NOPAT**: EBIT × (1 − Tax_Rate).
- **Depreciation**: Revenue × Depreciation_Pct_Revenue.
- **Capex**: Revenue × Capex_Pct_Revenue.
- **Net Working Capital**: Revenue × NWC_Pct_Revenue.
- **Change in NWC**: Current year NWC − Prior year NWC.
- **Free Cash Flow (FCF)**: NOPAT + Depreciation − Capex − Change in NWC.

### Valuation
- **Discount Factor**: 1 / (1 + Discount_Rate) ^ Year_Index.
- **Present Value of FCF**: FCF × Discount Factor.
- **Terminal Value (Perpetuity Growth)**: 
  - **TV at End of Projection**: Final Year FCF × (1 + Terminal_Growth_Rate) / (Discount_Rate − Terminal_Growth_Rate).
  - **PV of TV**: TV × Discount Factor (final year).
- **Enterprise Value**: Sum of PV of FCF + PV of TV.

### Outputs
- **Enterprise Value** (linked from Valuation).
- **Value Per Unit**: Enterprise Value / Units_Outstanding.
- **Summary Table**: Revenue, EBIT, FCF for each year (linked from Operating_Model).

---

## 4. Output Metrics
- Enterprise Value
- Present Value of Explicit FCF
- Present Value of Terminal Value
- Value Per Unit

---

## 5. Guardrails and Exclusions
- One operating scenario only.
- Flat discount rate only; no WACC decomposition.
- No balance sheet build.
- No debt schedules or financing structure.
- No multi-tranche waterfalls.
- No external data sources or APIs.
- No macros, no volatile functions, no circular references.
- No hidden constants in formulas (all assumptions in Inputs).

---

## 6. Notes/Disclaimer Content (Notes Tab)
**Purpose**
This template is designed for educational and illustrative valuation mechanics. It is intended to help users understand how inputs flow through a discounted cash flow model in a transparent, single-scenario structure.

**Important Disclosures**
- This spreadsheet is for educational use only and does not constitute investment advice.
- It does not guarantee accuracy, completeness, or suitability for any purpose.
- All outputs are mechanically derived from user inputs and do not represent forecasts.
- Users are responsible for validating assumptions and results.

**Definitions**
- **NOPAT**: Net Operating Profit After Tax; EBIT × (1 − Tax_Rate).
- **FCF**: Free Cash Flow to the Firm; NOPAT + Depreciation − Capex − Change in NWC.
- **Terminal Value**: Value of cash flows beyond the explicit forecast period using a single perpetuity growth method.

**Formatting Guidance**
- Inputs: light fill, border.
- Calculations: no fill.
- Outputs: bold label, subtle highlight.

---

## 7. Optional Variants (if needed)
**Lite Version**
- Remove depreciation, capex, and NWC sections; use a single FCF margin.

**Extended Version**
- Add a separate revenue driver block (volume × price) with input drivers.
