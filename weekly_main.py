from reports.interval_close_rate import interval_close_rate
from reports.interval_sales import interval_sales
from reports.interval_cancels import interval_cancels
from reports.weekly_scorecard import weekly_scorecard_report
from reports.ctm_call_report import ctm_call_report

if __name__ == "__main__":
    # Example usage
    beginning_of_time = "2023-01-01"
    first_day_of_week = "2025-03-17"
    replacements ={
        "hussamobetter@gmail.com": "Hussam Olabi",
        "kamaalsbetter@gmail.com": "Kamaal Sherrod",
        "service@bettertermite.com": "G Schulz"
    }

    sales_reps = ["Hussam Olabi", "Kamaal Sherrod", "Rob Dively"]
    reps_for_weekly_scorecard = ["Hussam Olabi", "Kamaal Sherrod"]
    
    interval_close_rate(beginning_of_time=beginning_of_time, salespeople=sales_reps)
    interval_sales(beginning_of_time=beginning_of_time, salespeople=sales_reps)
    interval_cancels(beginning_of_time=beginning_of_time, salespeople=sales_reps)
    ctm_call_report(beginning_of_time=beginning_of_time, agents=sales_reps)
    weekly_scorecard_report(first_day_of_week, beginning_of_time, reps_for_weekly_scorecard)