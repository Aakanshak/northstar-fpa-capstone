"""Generate FY2026 monthly budget and actuals with driver-based variances."""
import csv, random
from pathlib import Path
SEED=4417; OUTPUT=Path(__file__).resolve().parent/"output_csvs"/"actuals_vs_budget.csv"
DEPTS=["Sales","Customer Success","Marketing","Product & Engineering","G&A"]
OPEX={"Sales":920000,"Customer Success":390000,"Marketing":520000,"Product & Engineering":1160000,"G&A":480000}
def main():
 rng=random.Random(SEED); rows=[]
 for m in range(1,13):
  season=1.06 if m in (3,6,9,12) else .98
  total_budget=3_460_000*(1.015**(m-1))*season
  rev_factor={1:.985,2:.972,3:1.018,4:1.025,5:1.012,6:.963,7:.946,8:.958,9:.981,10:1.006,11:1.018,12:1.035}[m]
  units_budget=44+round(m*.8); units_actual=round(units_budget*(rev_factor-.006))
  arpa_budget=total_budget/units_budget; arpa_actual=total_budget*rev_factor/units_actual
  for d in DEPTS:
   rb=total_budget if d=="Sales" else 0; ra=total_budget*rev_factor if d=="Sales" else 0
   ob=OPEX[d]*(1.008**(m-1)); bias={"Sales":1.035,"Customer Success":.972,"Marketing":1.085,"Product & Engineering":1.018,"G&A":.955}[d]
   oa=ob*bias*(1+rng.uniform(-.012,.012))
   rows.append({"month":f"2026-{m:02d}","department":d,"revenue_budget":round(rb),"revenue_actual":round(ra),
                "opex_budget":round(ob),"opex_actual":round(oa),"units_budget":units_budget if d=="Sales" else 0,
                "units_actual":units_actual if d=="Sales" else 0,"arpa_budget":round(arpa_budget,2) if d=="Sales" else 0,
                "arpa_actual":round(arpa_actual,2) if d=="Sales" else 0})
 OUTPUT.parent.mkdir(parents=True,exist_ok=True)
 with OUTPUT.open('w',newline='',encoding='utf-8') as f:
  w=csv.DictWriter(f,fieldnames=list(rows[0]));w.writeheader();w.writerows(rows)
 q3=[r for r in rows if r['department']=='Sales' and r['month'] in ('2026-07','2026-08','2026-09')]
 print(f"actuals_vs_budget.csv: {len(rows)} rows; Q3 revenue variance=${sum(r['revenue_actual']-r['revenue_budget'] for r in q3):,.0f}")
if __name__=='__main__': main()
