import pandas as pd

df = pd.read_csv("weekly_review_data/Leads-Reporting Export.csv")

# cast the Close Date column to a datetime object
df["Close Date"] = pd.to_datetime(df["Close Date"])
df["Date Added"] = pd.to_datetime(df["Date Added"], format="%m/%d/%Y %I:%M%p")

START_DATE = "2025-05-01"
END_DATE = "2025-05-31"
BEGINNING_OF_TIME = "2025-02-01"

target_sequence = [
    {
        "name": "Followup 1",
        "days_after_initial_contact": 1,
        "grace_period_hours": 12,
    },
    {
        "name": "Followup 2",
        "days_after_initial_contact": 2,
        "grace_period_hours": 12,
    },
    {
        "name": "Followup 3",
        "days_after_initial_contact": 3,
        "grace_period_hours": 12,
    },
    {
        "name": "Followup 4",
        "days_after_initial_contact": 4,
        "grace_period_hours": 12,
    },
    {
        "name": "Followup 5",
        "days_after_initial_contact": 5,
        "grace_period_hours": 12,
    },
]

eligible_close_statuses = [
    "Lost: Never Reached",
    "Lost: Reached, Then Went Cold",
    "Lost: Auto-Closed",
    "Open"
]

# determine if the leads should even recieve a followup sequence. Leads that don't recieve a followup sequence are:

# get the leads that are Disqualified
eligible_leads = df[df["Close Status"].isin(eligible_close_statuses)]

# include only leads where either Close date is null or close date is within the start and end dates
eligible_leads = eligible_leads[
    (eligible_leads["Date Added"] >= BEGINNING_OF_TIME) &
    ((eligible_leads["Close Date"].isnull()) | 
    ((eligible_leads["Close Date"].dt.strftime("%Y-%m-%d") >= START_DATE) & 
    (eligible_leads["Close Date"].dt.strftime("%Y-%m-%d") <= END_DATE)))
]

#save the eligible leads to a csv
eligible_leads.to_csv("weekly_outputs/eligible_leads.csv", index=False)


