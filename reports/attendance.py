from datetime import datetime, time
import pandas as pd

def get_weekdays(df):
    # Get min and max dates from the dataset
    start_date = pd.to_datetime(df['Date']).min()
    end_date = pd.to_datetime(df['Date']).max()
    
    # Generate all dates in range
    date_range = pd.date_range(start=start_date, end=end_date)
    
    # Filter to only weekdays
    weekdays = date_range[date_range.dayofweek < 5]
    return weekdays

def attendance_report(eight_o_clock_starts=[], eight_thirty_starts=[]):
    # Read the data
    df = pd.read_excel("data/Timesheets.xlsx")
    
    # Create full name and convert date column
    df['Full Name'] = df['First Name'] + ' ' + df['Last Name']
    df['Date'] = pd.to_datetime(df['Date']).dt.date  # Convert to date only
    
    # Convert Start Time to datetime and extract time
    df['Start Time'] = pd.to_datetime(df['Start Time']).dt.time
    
    # Group by date and name, get minimum start time
    attendance = df.groupby(['Date', 'Full Name'])['Start Time'].min().reset_index()
    
    # Pivot the data
    pivot_table = attendance.pivot(
        index='Date',
        columns='Full Name',
        values='Start Time'
    )
    
    # Get all weekdays
    weekdays = get_weekdays(df)
    weekdays = weekdays.date  # Convert to date only
    
    # Reindex to include all weekdays
    pivot_table = pivot_table.reindex(weekdays)
    
    # Create Excel writer object with xlsxwriter engine
    writer = pd.ExcelWriter('outputs/Attendance.xlsx', engine='xlsxwriter')
    
    # Write the DataFrame to Excel
    pivot_table.to_excel(writer, sheet_name='Attendance')
    
    # Get the workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['Attendance']
    
    # Define formats
    red_format = workbook.add_format({
        'bg_color': '#FFD9D9',  # Light red
        'num_format': 'h:mm AM/PM'
    })
    yellow_format = workbook.add_format({
        'bg_color': '#FFFFCC',  # Light yellow
        'num_format': 'h:mm AM/PM'
    })
    green_format = workbook.add_format({
        'bg_color': '#E6FFE6',  # Light green
        'num_format': 'h:mm AM/PM'
    })
    bold_format = workbook.add_format({'bold': True})
    date_format = workbook.add_format({
        'bold': True,
        'num_format': 'ddd mm/dd/yy'
    })
    
    # Apply bold format to column headers
    for col_num, value in enumerate(['Date'] + pivot_table.columns.tolist()):
        worksheet.write(0, col_num, value, bold_format)
    
    # Apply date format to date column
    for row_num, value in enumerate(pivot_table.index):
        worksheet.write_datetime(row_num + 1, 0, datetime.combine(value, time()), date_format)
    
    # Apply conditional formatting to data cells
    for row in range(pivot_table.shape[0]):
        for col in range(pivot_table.shape[1]):
            cell_value = pivot_table.iloc[row, col]
            employee_name = pivot_table.columns[col]
            
            # Determine cutoff time based on employee group
            if employee_name in eight_o_clock_starts:
                cutoff_time = time(8, 5)  # 8:05 AM
            elif employee_name in eight_thirty_starts:
                cutoff_time = time(8, 35)  # 8:35 AM
            else:
                cutoff_time = time(7, 35)  # 7:00 AM
            
            if pd.isna(cell_value):
                worksheet.write(row + 1, col + 1, '', red_format)
            elif isinstance(cell_value, time) and cell_value > cutoff_time:
                # Convert time object to Excel time value (fraction of 24 hours)
                excel_time = (cell_value.hour + cell_value.minute/60.0 + cell_value.second/3600.0) / 24.0
                worksheet.write(row + 1, col + 1, excel_time, yellow_format)
            else:
                # Convert time object to Excel time value
                excel_time = (cell_value.hour + cell_value.minute/60.0 + cell_value.second/3600.0) / 24.0
                worksheet.write(row + 1, col + 1, excel_time, green_format)
    
    # Save the workbook
    writer.close()