# Excel Model Structure

1. **Cover & Instructions** - purpose, model-use steps, circularity policy, and integrity checks.
2. **Assumptions** - Base, Upside, and Downside drivers; yellow inputs and named ranges.
3. **Customer ARR Ledger** - 36 months of customer-month source data with MRR movements.
4. **ARR Bridge** - monthly beginning ARR, new, expansion, contraction, churn, ending ARR, and zero check using `SUMIFS`.
5. **Revenue Forecast** - three driver-based cases and a selected scenario output; probability-weighted pipeline is stored in hidden source columns.
6. **Headcount & Opex Plan** - department headcount, personnel expense, and non-personnel spend; raw role data is stored in hidden source columns.
7. **Income Statement** - revenue through net income, linked to revenue and opex builds.
8. **Balance Sheet** - cash, working capital, PP&E, debt, and equity with a balance check.
9. **Cash Flow Statement** - indirect method from net income through ending cash.
10. **SaaS Metrics Dashboard** - ARR, MRR, NRR, GRR, CAC, LTV, LTV:CAC, Rule of 40, Magic Number, Burn Multiple, runway, and EBITDA margin.
11. **Budget vs Actual** - revenue volume and price/mix decomposition, departmental opex variance, and materiality flags; source actuals are stored in hidden columns.
12. **Sensitivity Analysis** - ending cash and runway grids across churn and hiring pace plus a tornado source table.
13. **Cohort LTV Analysis** - acquisition-quarter retention, CAC, LTV:CAC, and payback.

### Modeling conventions

- Blue font indicates formulas; yellow fill indicates editable assumptions.
- Cross-sheet formulas quote worksheet names.
- Financial outputs use consistent $K presentation and negative-value formatting.
- Interest uses prior-period cash; iterative calculation is not required.
- Raw source data is never manually restated in output sections.

## So what?

The architecture separates inputs, calculations, and outputs so a reviewer can trace any board metric back to an assumption or source record without relying on Python-calculated answers.
