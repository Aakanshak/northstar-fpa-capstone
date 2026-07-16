"""Generate a monthly role-level headcount plan for 2024-2026."""
import csv
from datetime import date
from pathlib import Path

START=date(2024,1,1); MONTHS=36
OUTPUT=Path(__file__).resolve().parent/"output_csvs"/"headcount_plan.csv"
ROLES={
 "Sales":{"Account Executive":(32,185000,4),"Sales Development Rep":(18,105000,3),"Sales Leadership":(5,240000,1)},
 "Customer Success":{"Customer Success Manager":(18,125000,2),"Implementation":(8,112000,1)},
 "Marketing":{"Demand Generation":(9,145000,1),"Content & Brand":(6,132000,1)},
 "Product & Engineering":{"Software Engineer":(52,175000,5),"Product Manager":(10,165000,1),"Design & Research":(7,145000,1)},
 "G&A":{"Finance & Accounting":(8,142000,1),"People & Legal":(7,150000,1),"IT & Facilities":(5,125000,0)},
}
def add_months(d,n):
 x=d.month-1+n; return date(d.year+x//12,x%12+1,1)
def main():
 rows=[]
 for m in range(MONTHS):
  dt=add_months(START,m)
  for dept,roles in ROLES.items():
   for role,(base,cost,annual_hires) in roles.items():
    hires=(m*annual_hires)//12
    # Planned hiring slows in H2 2026 pending pipeline conversion.
    if m>=30: hires=max(0,hires-(m-29)//2)
    hc=base+hires
    rows.append({"month":dt.strftime('%Y-%m'),"department":dept,"role":role,"start_date":START.isoformat(),
                 "fully_loaded_annual_cost":cost,"headcount":hc,"monthly_cost":round(hc*cost/12,2)})
 OUTPUT.parent.mkdir(parents=True,exist_ok=True)
 with OUTPUT.open('w',newline='',encoding='utf-8') as f:
  w=csv.DictWriter(f,fieldnames=list(rows[0]));w.writeheader();w.writerows(rows)
 print(f"headcount_plan.csv: {len(rows)} rows; HC {sum(r['headcount'] for r in rows if r['month']=='2024-01')} -> {sum(r['headcount'] for r in rows if r['month']=='2026-12')}")
if __name__=='__main__': main()
