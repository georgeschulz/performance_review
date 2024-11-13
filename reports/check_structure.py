import os
import pandas as pd
import re

# The purpose of this script is to check to make sure all the files are setup properly from whoever gave me the documents

def check_structure():
    root_dir = "data"
    # iterate over the data folder and remove " (#)" pattern from the end of the file names that appear due to the files being downloaded from the web
    for file in os.listdir(root_dir):
        if re.search(r" \([0-9]+\)", file):
            new_file = re.sub(r" \([0-9]+\)", "", file)
            os.rename(os.path.join(root_dir, file), os.path.join(root_dir, new_file))

    # ensure there is a file called Monthly Invoice Report
    monthly_invoice_report = os.path.join(root_dir, "Monthly Invoice Report.csv")
    assert os.path.exists(monthly_invoice_report), "Monthly Invoice Report.csv not found"

    # ensure that it has at least 5 rows
    df = pd.read_csv(monthly_invoice_report)
    assert len(df) >= 5, "Monthly Invoice Report.csv does not have at least 5 rows"

    # ensure that Historical Invoice Report exists
    historical_invoice_report = os.path.join(root_dir, "Historical Invoice Report.csv")
    assert os.path.exists(historical_invoice_report), "Historical Invoice Report.csv not found"

    # ensure that it has at least 5 rows
    df = pd.read_csv(historical_invoice_report)
    assert len(df) >= 5, "Historical Invoice Report.csv does not have at least 5 rows"

    # ensure that Timesheets exists
    timesheets = os.path.join(root_dir, "Timesheets.xlsx")
    assert os.path.exists(timesheets), "Timesheets.xlsx not found"


    # ensure a data, outputs and reports folder exist. If it doesn't exist, create it
    if not os.path.exists("data"):
        os.makedirs("data")
    if not os.path.exists("outputs"):
        os.makedirs("outputs")
    if not os.path.exists("reports"):
        os.makedirs("reports")
