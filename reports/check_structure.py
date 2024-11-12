import os
import pandas as pd

# The purpose of this script is to check to make sure all the files are setup properly from whoever gave me the documents

def check_structure():
    root_dir = "data"
    # ensure there is a file called Monthly Invoice Report
    monthly_invoice_report = os.path.join(root_dir, "Monthly Invoice Report.csv")
    assert os.path.exists(monthly_invoice_report), "Monthly Invoice Report.csv not found"

    # ensure that it has at least 5 rows
    df = pd.read_csv(monthly_invoice_report)
    assert len(df) >= 5, "Monthly Invoice Report.csv does not have at least 5 rows"


    # ensure a data, outputs and reports folder exist. If it doesn't exist, create it
    if not os.path.exists("data"):
        os.makedirs("data")
    if not os.path.exists("outputs"):
        os.makedirs(os.path.join(root_dir, "outputs"))
    if not os.path.exists("reports"):
        os.makedirs("reports")
