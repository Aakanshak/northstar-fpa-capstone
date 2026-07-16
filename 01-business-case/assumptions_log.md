# Assumptions Log

All workbook drivers are centralized on the `Assumptions` tab. Yellow cells are editable inputs; the selected values flow through named ranges.

| Driver | Base | Upside | Downside | Rationale |
|---|---:|---:|---:|---|
| Scenario selector | 1 | 2 | 3 | Uses `CHOOSE`; 1 is the board planning case. |
| Monthly logo churn | 1.10% | 0.80% | 1.60% | Enterprise-weighted SaaS portfolio with downside pressure. |
| Monthly expansion | 0.65% | 0.85% | 0.40% | Seat growth and plan upgrades. |
| Monthly contraction | 0.20% | 0.14% | 0.32% | Downgrades and seat optimization. |
| Pipeline realization | 0.55x | 0.70x | 0.38x | Applied after stage probability weighting. |
| Sales-cycle timing | 1.00x | 0.92x | 1.18x | Upside closes faster; downside slips. |
| Hiring pace | 1.00x | 1.08x | 0.82x | Distinct capacity and cash tradeoff. |
| Gross margin | 80.5% | 81.5% | 78.5% | Hosting and support efficiency range. |
| Merit increase | 4.0% | 4.0% | 3.0% | Annual compensation inflation. |
| Interest on cash | 4.5% | 4.8% | 4.0% | Short-term yield on beginning cash. |
| Capex / revenue | 2.5% | 2.4% | 2.0% | Internal systems and equipment. |
| DSO | 48 days | 44 days | 55 days | Collections sensitivity. |
| DPO | 32 days | 36 days | 28 days | Vendor-payment sensitivity. |
| Opening cash | $18.0M | $18.0M | $18.0M | 1-Jan-2026 balance. |
| Opening debt | $5.0M | $5.0M | $5.0M | Existing term debt. |
| Tax rate | 21% | 21% | 21% | Blended statutory proxy. |
| CAC per logo | $42K | $39K | $49K | Fully loaded sales and marketing acquisition cost. |
| Discount rate | 12% | 11% | 14% | SaaS cost-of-capital proxy. |
| Materiality | 5% / $50K | same | same | Board variance flag if either threshold is exceeded. |

Other operating assumptions are visible source inputs: customer mix, plan pricing, stage probabilities, role costs, department non-personnel spend, and FY2026 budget phasing. Revenue is recognized ratably from average ARR. Depreciation uses a five-year proxy. Deferred revenue is 35% of monthly revenue.

### Circularity policy

The model uses a documented circularity breaker. Interest income is calculated on prior-period cash. Cash Flow calculates ending cash, and Balance Sheet links to Cash Flow. This avoids iterative calculation while preserving an economically reasonable interest estimate.

## So what?

The forecast is auditable because every outcome can be traced to a visible operational driver, and scenario differences represent real business conditions rather than arbitrary percentage overlays.
