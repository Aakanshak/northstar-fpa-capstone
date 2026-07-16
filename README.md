# Northstar SaaS - Revenue Forecasting, Budget Variance & Board Reporting

Portfolio-grade FP&A capstone for Corporate Finance, FP&A, and Strategic Finance roles. The project is deliberately **Excel/PowerPoint-native**: Python generates realistic source data and constructs the files, while the analysis, scenarios, three statements, KPIs, and sensitivities remain live Excel formulas.

## Executive findings

- Q3 revenue missed budget by **$437K**, driven by new-logo volume and close timing rather than broad pricing pressure.
- The Base case forecasts **$52.5M FY2026 revenue**, **$58.3M ending ARR**, **$6.7M ending cash**, and **12.3 months of runway**.
- FY opex is **$871K unfavorable**, led by Marketing (+$550K) and Sales (+$454K).
- Recommendation: hold incremental GTM hiring until conversion and sales-cycle duration improve for two consecutive months.

## Primary deliverables

- `03-excel-model/northstar_fpa_model.xlsx` - 13-tab, formula-driven integrated model with named assumptions, 3 statements, SaaS metrics, variance decomposition, native Excel What-If Data Tables, and cohort economics.
- `05-board-deck/northstar_q3_board_deck.pptx` - editable eight-slide board deck with native charts.
- `05-board-deck/northstar_q3_board_deck.pdf` - presentation-ready PDF export.
- `04-variance-narrative/variance_commentary.md` - root-cause commentary for material variances.
- `06-executive-memo/cfo_memo.md` - one-page CFO recommendation memo.

## Rebuild locally

Requirements: Windows, Python 3.11+, Microsoft Excel, and Microsoft PowerPoint. Office is used to recalculate formulas, create native What-If Data Tables, and export the deck PDF.

```powershell
cd northstar-fpa-capstone
python -m pip install -r requirements.txt
python .\02-data-generation\run_all.py
python .\03-excel-model\build_workbook.py
python .\05-board-deck\build_deck.py
```

Generated CSVs land in `02-data-generation/output_csvs/`. The generators use fixed random seeds, so the portfolio findings are reproducible.

## Model design highlights

- Central Base/Upside/Downside drivers for churn, expansion, contraction, conversion, sales-cycle timing, hiring, gross margin, working capital, and CAC.
- Customer-level ARR bridge built with `SUMIFS`; selected scenario uses `CHOOSE`, `INDEX/MATCH`, and named ranges.
- Income Statement -> Balance Sheet -> Cash Flow integration with a documented prior-period-cash circularity breaker.
- Revenue variance decomposed into volume and price/mix; opex variance shown by department with 5% / $50K materiality flags.
- Conditional formatting on retention, LTV:CAC, Rule of 40, and variance thresholds.
- Native two-variable Excel Data Tables for ending cash and runway sensitivity to churn and hiring pace.

## Recruiter-ready resume bullets

- Built a 13-tab, formula-driven SaaS FP&A model covering ARR, revenue, headcount, three statements, cash runway, SaaS KPIs, cohort LTV, and scenario planning for a fictional $40M ARR company.
- Diagnosed a $437K Q3 revenue miss through volume and price/mix decomposition and identified an $871K FY opex overrun concentrated in Sales and Marketing.
- Developed Base/Upside/Downside forecasts yielding a $48.2M-$55.9M FY revenue range and translated results into an eight-slide CFO/board narrative with hiring gates and downside triggers.

## Loom walkthrough

`LOOM VIDEO LINK GOES HERE ONCE RECORDED`

## So what?

This capstone demonstrates the core FP&A workflow recruiters test: build an auditable Excel model, explain actual performance, quantify forecast risk, and turn the analysis into a clear capital-allocation recommendation.
