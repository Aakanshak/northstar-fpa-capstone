"""Generate Northstar's deterministic 36-month customer ARR ledger.

Python is used only to create realistic source-system-like data. All analysis
and model outputs are calculated later with Excel formulas.
"""
from __future__ import annotations

import csv
import random
from dataclasses import dataclass
from datetime import date
from pathlib import Path

SEED = 20260716
START = date(2024, 1, 1)
MONTHS = 36
OUTPUT = Path(__file__).resolve().parent / "output_csvs" / "customer_arr_ledger.csv"
TIERS = {"SMB": (1800, 4500), "Mid-Market": (6500, 15000), "Enterprise": (24000, 65000), "Strategic": (75000, 145000)}
TIER_WEIGHTS = [0.46, 0.34, 0.16, 0.04]
CHANNELS = ["Inbound", "Outbound", "Partner", "Event"]
REPS = ["Avery", "Jordan", "Morgan", "Riley", "Taylor", "Casey", "Jamie"]

def add_months(d: date, n: int) -> date:
    x = d.month - 1 + n
    return date(d.year + x // 12, x % 12 + 1, 1)

def month_key(d: date) -> str:
    return d.strftime("%Y-%m")

@dataclass
class Customer:
    customer_id: str
    cohort: date
    tier: str
    channel: str
    rep: str
    mrr: int
    active: bool = True

def make_customer(rng: random.Random, customer_id: int, cohort: date) -> Customer:
    tier = rng.choices(list(TIERS), weights=TIER_WEIGHTS, k=1)[0]
    lo, hi = TIERS[tier]
    mrr = round(rng.triangular(lo, hi, lo * 1.35) / 100) * 100
    channel = rng.choices(CHANNELS, weights=[0.38, 0.34, 0.20, 0.08], k=1)[0]
    return Customer(f"CUST-{customer_id:04d}", cohort, tier, channel, rng.choice(REPS), mrr)

def generate_rows() -> list[dict]:
    rng = random.Random(SEED)
    customers = [make_customer(rng, i, START) for i in range(1000, 1235)]
    # Scale opening MRR to approximately $3.33M / $40M ARR.
    scale = 3_333_333 / sum(c.mrr for c in customers)
    for c in customers:
        c.mrr = max(500, round(c.mrr * scale / 100) * 100)
    next_id, rows = 1235, []
    for idx in range(MONTHS):
        month = add_months(START, idx)
        if idx:
            new_count = max(4, round(rng.gauss(6.0 + idx * 0.04, 1.3)))
            for _ in range(new_count):
                new_customer = make_customer(rng, next_id, month)
                new_customer.mrr = max(500, round(new_customer.mrr * 0.65 / 100) * 100)
                customers.append(new_customer); next_id += 1
        for c in customers:
            if not c.active or c.cohort > month:
                continue
            prior = c.mrr
            status = "new" if c.cohort == month else "active"
            if status != "new":
                roll = rng.random()
                churn_prob = 0.0095 if c.tier in ("Enterprise", "Strategic") else 0.0135
                if roll < churn_prob:
                    c.mrr, c.active, status = 0, False, "churned"
                elif roll < churn_prob + 0.025:
                    c.mrr = round(prior * rng.uniform(1.05, 1.16) / 100) * 100; status = "expanded"
                elif roll < churn_prob + 0.041:
                    c.mrr = round(prior * rng.uniform(0.82, 0.94) / 100) * 100; status = "contracted"
            movement = c.mrr if status == "new" else c.mrr - prior
            rows.append({"month": month_key(month), "customer_id": c.customer_id, "plan_tier": c.tier,
                         "acquisition_channel": c.channel, "sales_rep": c.rep, "mrr": c.mrr,
                         "status": status, "cohort_month": month_key(c.cohort), "prior_mrr": prior if status != "new" else 0,
                         "mrr_movement": movement})
    return rows

def main() -> None:
    rows = generate_rows(); OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0])); w.writeheader(); w.writerows(rows)
    print(f"customer_arr_ledger.csv: {len(rows):,} rows; opening ARR=${sum(r['mrr'] for r in rows if r['month']=='2024-01')*12:,.0f}; ending ARR=${sum(r['mrr'] for r in rows if r['month']=='2026-12')*12:,.0f}")

if __name__ == "__main__": main()
