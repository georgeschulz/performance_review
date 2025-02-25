import os
import pandas as pd
from datetime import datetime, timedelta

# Import the interval report functions
from reports.interval_calls import interval_calls
from reports.interval_sales import interval_sales
from reports.interval_cancels import interval_cancels
from reports.interval_close_rate import interval_close_rate

def get_value_from_df(df, rep_col, rep_name, week_start_col, week_start_val, target_col):
    """
    Helper to get a value for a given rep from a DataFrame filtered by the week start.
    Returns 0 if the DataFrame is empty or if any column is missing.
    """
    # If the DataFrame is empty or missing any of the expected columns, just return 0.
    for col in [rep_col, week_start_col, target_col]:
        if col not in df.columns:
            return 0
    row = df[(df[rep_col] == rep_name) & (df[week_start_col] == week_start_val)]
    if not row.empty:
        return row.iloc[0][target_col]
    return 0

def get_value_by_rep(df, rep_col, rep_name, target_col):
    """
    Helper to get a value for a given rep from a DataFrame keyed by rep.
    Returns 0 if the DataFrame is empty or if the column is missing.
    """
    if rep_col not in df.columns or target_col not in df.columns:
        return 0
    row = df[df[rep_col] == rep_name]
    if not row.empty:
        return row.iloc[0][target_col]
    return 0

def weekly_scorecard_report(first_day_of_week_str, beginning_of_time="2023-01-01", sales_reps=None):
    """
    Creates a weekly scorecard Excel report (Weekly Scorecard.xlsx) that summarizes,
    for each sales rep, the following metrics for:
      - Week-to-date (WTD)
      - Month-to-date (MTD)
      - Year-to-date (YTD)
    
    Metrics include:
      - Total Calls
      - Total Sales
      - Average First Year Price
      - Cancels
      - Starts Count
      - Close Rate
      
    The final Excel file includes a header with the report's date criteria.
    """
    if sales_reps is None:
        sales_reps = ["Hussam Olabi", "Kamaal Sherrod", "Rob Dively"]

    # Convert first day of week to datetime and compute week end (assuming Monday-Sunday week)
    first_day = pd.to_datetime(first_day_of_week_str)
    week_end = first_day + timedelta(days=6)
    week_start_str = first_day.strftime('%Y-%m-%d')
    week_end_str = week_end.strftime('%Y-%m-%d')

    # --- Get Calls Data ---
    # Call interval_calls with current_date set to the week_end
    calls_data = interval_calls(beginning_of_time=beginning_of_time, agents=sales_reps, current_date=week_end_str)
    calls_weekly_df = calls_data.get('weekly', pd.DataFrame())
    calls_ytd_mtd_df = calls_data.get('ytd_mtd', pd.DataFrame())
    # For calls, the rep column is 'Agent' and the week start column is 'Week Start'
    
    # --- Get Sales Data ---
    sales_data = interval_sales(beginning_of_time=beginning_of_time, salespeople=sales_reps, current_date=week_end_str)
    sales_weekly_df = sales_data.get('weekly', pd.DataFrame())
    sales_ytd_mtd_df = sales_data.get('ytd_mtd', pd.DataFrame())
    # For sales, the rep column is 'Salesperson'
    
    # --- Get Cancels/Starts Data ---
    cancels_data = interval_cancels(beginning_of_time=beginning_of_time, salespeople=sales_reps, current_date=week_end_str)
    cancels_weekly_df = cancels_data.get('weekly', pd.DataFrame())
    cancels_ytd_mtd_df = cancels_data.get('ytd_mtd', pd.DataFrame())
    # For cancels, the rep column is 'Salesperson'
    
    # --- Get Close Rate Data ---
    close_rate_data = interval_close_rate(beginning_of_time=beginning_of_time, salespeople=sales_reps, current_date=week_end_str)
    close_rate_weekly_df = close_rate_data.get('weekly', pd.DataFrame())
    close_rate_ytd_mtd_df = close_rate_data.get('ytd_mtd', pd.DataFrame())
    # For close rate, the rep column is 'Salesperson'

    # Build a dictionary to hold the final metrics per rep and period.
    # We will have columns keyed by (rep, period) and rows for each metric.
    periods = ["Week", "MTD", "YTD"]
    # Metrics: Total Calls, Total Sales, Avg First Year Price, Cancels, Starts Count, Close Rate
    final_data = {}

    for rep in sales_reps:
        for period in periods:
            cell = {}
            # --- Total Calls ---
            if period == "Week":
                calls_val = get_value_from_df(calls_weekly_df, "Agent", rep, "Week Start", week_start_str, "Total Calls")
            elif period == "MTD":
                calls_val = get_value_by_rep(calls_ytd_mtd_df, "Agent", rep, "MTD Total Calls")
            else:  # YTD
                calls_val = get_value_by_rep(calls_ytd_mtd_df, "Agent", rep, "YTD Total Calls")
            cell["Total Calls"] = calls_val

            # --- Sales Data (Total Sales and Average Sale) ---
            if period == "Week":
                total_sales = get_value_from_df(sales_weekly_df, "Salesperson", rep, "Week Start", week_start_str, "Total Sales")
                avg_sale = get_value_from_df(sales_weekly_df, "Salesperson", rep, "Week Start", week_start_str, "Average Sale")
            elif period == "MTD":
                total_sales = get_value_by_rep(sales_ytd_mtd_df, "Salesperson", rep, "MTD Total Sales")
                avg_sale = get_value_by_rep(sales_ytd_mtd_df, "Salesperson", rep, "MTD Average Sale")
            else:  # YTD
                total_sales = get_value_by_rep(sales_ytd_mtd_df, "Salesperson", rep, "YTD Total Sales")
                avg_sale = get_value_by_rep(sales_ytd_mtd_df, "Salesperson", rep, "YTD Average Sale")
            cell["Total Sales"] = total_sales
            cell["Avg First Year Price"] = avg_sale

            # --- Cancels and Starts ---
            if period == "Week":
                cancels_val = get_value_from_df(cancels_weekly_df, "Salesperson", rep, "Week Start", week_start_str, "Cancels")
                starts_val = get_value_from_df(cancels_weekly_df, "Salesperson", rep, "Week Start", week_start_str, "Starts")
            elif period == "MTD":
                cancels_val = get_value_by_rep(cancels_ytd_mtd_df, "Salesperson", rep, "MTD Cancels")
                starts_val = get_value_by_rep(cancels_ytd_mtd_df, "Salesperson", rep, "MTD Starts")
            else:  # YTD
                cancels_val = get_value_by_rep(cancels_ytd_mtd_df, "Salesperson", rep, "YTD Cancels")
                starts_val = get_value_by_rep(cancels_ytd_mtd_df, "Salesperson", rep, "YTD Starts")
            cell["Cancels"] = cancels_val
            cell["Starts Count"] = starts_val
            
            # --- Close Rate ---
            if period == "Week":
                close_rate_val = get_value_from_df(close_rate_weekly_df, "Salesperson", rep, "Week Start", week_start_str, "Close Rate")
            elif period == "MTD":
                close_rate_val = get_value_by_rep(close_rate_ytd_mtd_df, "Salesperson", rep, "MTD Close Rate")
            else:  # YTD
                close_rate_val = get_value_by_rep(close_rate_ytd_mtd_df, "Salesperson", rep, "YTD Close Rate")
            cell["Close Rate"] = close_rate_val

            final_data[(rep, period)] = cell

    # Create a DataFrame with rows as metrics and MultiIndex columns (rep, period)
    scorecard_df = pd.DataFrame(final_data)
    # Ensure the rows are in the desired order
    scorecard_df = scorecard_df.reindex(["Total Calls", "Total Sales", "Avg First Year Price", "Cancels", "Starts Count", "Close Rate"])

    # Sort the columns by rep and period order (WTD, then MTD, then YTD)
    sorted_columns = []
    for rep in sales_reps:
        for period in periods:
            sorted_columns.append((rep, period))
    scorecard_df = scorecard_df[sorted_columns]

    # --- Write to Excel with header information ---
    output_path = os.path.join("weekly_outputs", "Weekly Scorecard.xlsx")
    writer = pd.ExcelWriter(output_path, engine="xlsxwriter")
    
    workbook = writer.book
    worksheet = writer.sheets.get("Scorecard") or workbook.add_worksheet("Scorecard")
    
    # Write the report title and date criteria at the top
    title_format = workbook.add_format({'bold': True, 'font_size': 14})
    criteria_format = workbook.add_format({'italic': True, 'font_color': '#555555'})
    worksheet.write("A1", "Weekly Scorecard Report", title_format)
    criteria_text = f"Report Date Criteria: Week Start = {week_start_str}, Week End = {week_end_str}"
    worksheet.write("A2", criteria_text, criteria_format)
    
    # Get column headers (will be used for manual writing)
    reps = scorecard_df.columns.get_level_values(0).unique()
    periods = scorecard_df.columns.get_level_values(1).unique()
    
    # Create formats for different cell types
    header_format = workbook.add_format({
        'bold': True, 
        'align': 'center',
        'border': 1  # Add border to header cells
    })
    currency_format = workbook.add_format({
        'num_format': '$#,##0.00',  # Format as currency with 2 decimal places
    })
    percentage_format = workbook.add_format({
        'num_format': '0.00"%"',  # Format as percentage with 2 decimal places
    })
    
    # Write first level of MultiIndex (rep names) with merged cells
    col_offset = 1  # Start after the "Metric" column
    worksheet.write(3, 0, "Metric", header_format)
    
    for rep in reps:
        # Calculate span (number of periods)
        span = len(periods)
        # Write rep name and merge cells across periods
        worksheet.merge_range(3, col_offset, 3, col_offset + span - 1, rep, header_format)
        
        # Write period names in the second row
        for i, period in enumerate(periods):
            worksheet.write(4, col_offset + i, period, header_format)
        
        col_offset += span
    
    # Convert scorecard_df to values array for direct writing
    data_values = scorecard_df.values
    row_labels = scorecard_df.index.tolist()
    
    # Write the row headers and data starting at row 5 (after the column headers)
    for i, label in enumerate(row_labels):
        worksheet.write(5 + i, 0, label)
        for j in range(data_values.shape[1]):
            # Apply currency format to Total Sales and Avg First Year Price
            if label in ["Total Sales", "Avg First Year Price"]:
                worksheet.write(5 + i, 1 + j, data_values[i, j], currency_format)
            # Apply percentage format to Close Rate
            elif label == "Close Rate":
                worksheet.write(5 + i, 1 + j, data_values[i, j], percentage_format)  # Remove division by 100 since the value is already a percentage
            else:
                worksheet.write(5 + i, 1 + j, data_values[i, j])
    
    # Set column widths for a nice look
    worksheet.set_column(0, 0, 25)  # Metric names column
    for col in range(1, col_offset):
        worksheet.set_column(col, col, 15)

    writer.close()