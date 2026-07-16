"""Generate a deterministic FY2026 bottoms-up sales pipeline."""
import csv, random
from datetime import date
from pathlib import Path
SEED=90210; OUTPUT=Path(__file__).resolve().parent/"output_csvs"/"sales_pipeline.csv"
STAGES=[("Discovery",.15),("Demo",.30),("Evaluation",.50),("Proposal",.65),("Negotiation",.80),("Commit",.92)]
REPS=["Avery","Jordan","Morgan","Riley","Taylor","Casey","Jamie"]
def main():
 rng=random.Random(SEED); rows=[]
 for i in range(420):
  seg=rng.choices(["SMB","Mid-Market","Enterprise"],[.38,.42,.20])[0]
  lo,hi={"SMB":(18000,60000),"Mid-Market":(65000,190000),"Enterprise":(200000,650000)}[seg]
  stage,prob=rng.choices(STAGES,[17,20,19,18,16,10])[0]
  month=rng.choices(range(1,13),[6,7,9,7,7,10,6,6,11,7,8,16])[0]
  rows.append({"deal_id":f"DEAL-{2000+i}","stage":stage,"probability":prob,"deal_size":round(rng.triangular(lo,hi,lo*1.35)/1000)*1000,
               "expected_close_date":date(2026,month,1).isoformat(),"sales_rep":rng.choice(REPS),"segment":seg,
               "acquisition_channel":rng.choices(["Inbound","Outbound","Partner","Event"],[.35,.38,.20,.07])[0],
               "sales_cycle_days":round(rng.gauss({"SMB":52,"Mid-Market":83,"Enterprise":128}[seg],15))})
 OUTPUT.parent.mkdir(parents=True,exist_ok=True)
 with OUTPUT.open('w',newline='',encoding='utf-8') as f:
  w=csv.DictWriter(f,fieldnames=list(rows[0]));w.writeheader();w.writerows(rows)
 weighted=sum(r['deal_size']*r['probability'] for r in rows)
 print(f"sales_pipeline.csv: {len(rows)} deals; gross=${sum(r['deal_size'] for r in rows):,.0f}; probability-weighted=${weighted:,.0f}")
if __name__=='__main__': main()
