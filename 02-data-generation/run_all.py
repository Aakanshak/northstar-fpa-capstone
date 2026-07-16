"""Run all source-data generators in the required order."""
import subprocess, sys
from pathlib import Path
HERE=Path(__file__).resolve().parent
for script in ["generate_customer_ledger.py","generate_headcount_plan.py","generate_pipeline_data.py","generate_actuals_vs_budget.py"]:
    subprocess.run([sys.executable,str(HERE/script)],check=True)
