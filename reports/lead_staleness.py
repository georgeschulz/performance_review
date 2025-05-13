import pandas as pd

df = pd.read_csv("weekly_review_data/Leads-Reporting Export.csv")

# Last Touch Date
df["Last Touch Date"] = pd.to_datetime(df["Last Touch Date"])

# get only leads that are Close Status = "Open"
open_leads = df[(df["Close Status"] == "Open") | (df["Close Status"] == "Lost: Auto-Closed")]

# Calculate days since last touch for each lead
open_leads["Days Since Last Touch"] = (pd.Timestamp.now() - open_leads["Last Touch Date"]).dt.days 

# Select only the columns we want to save in the raw data
raw_data_columns = ["Customer First Name", "Customer Last Name", "Salesperson", "Lead Type", "Close Status", "Date Added", "Last Touch Date", "Days Since Last Touch"]
raw_data = open_leads[raw_data_columns]

# sort by "Salesperson" and "Days Since Last Touch"
raw_data = raw_data.sort_values(by=["Salesperson", "Days Since Last Touch"])

# Save the raw data to a CSV file
raw_data.to_csv("weekly_outputs/open_leads_staleness_raw.csv", index=False)

# Group by "Salesperson" and calculate the avg. days since last touch
summary_df = open_leads.groupby("Salesperson")["Days Since Last Touch"].mean().reset_index()

# Rename the column to "Avg. Days Since Last Touch"
summary_df = summary_df.rename(columns={"Days Since Last Touch": "Avg. Days Since Last Touch"})

# Append the summary data to the bottom of the raw data file
with open("weekly_outputs/open_leads_staleness_raw.csv", "a") as f:
    f.write("\n\nSUMMARY - AVERAGE DAYS SINCE LAST TOUCH BY SALESPERSON\n")
    summary_df.to_csv(f, index=False)

print(summary_df.head())