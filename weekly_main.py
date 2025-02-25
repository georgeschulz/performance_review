from reports.interval_close_rate import interval_close_rate
from reports.interval_sales import interval_sales
from reports.interval_calls import interval_calls
from reports.interval_cancels import interval_cancels

if __name__ == "__main__":
    # Example usage
    beginning_of_time = "2023-01-01"
    replacements ={
        "hussamobetter@gmail.com": "Hussam Olabi",
        "kamaalsbetter@gmail.com": "Kamaal Sherrod",
        "service@bettertermite.com": "G Schulz"
    }

    sales_reps = ["Hussam Olabi", "Kamaal Sherrod", "Rob Dively"]
    
    interval_close_rate(beginning_of_time=beginning_of_time, salespeople=sales_reps)
    interval_sales(beginning_of_time=beginning_of_time, salespeople=sales_reps)
    interval_calls(beginning_of_time=beginning_of_time, replacements=replacements)
    interval_cancels(beginning_of_time=beginning_of_time, salespeople=sales_reps)