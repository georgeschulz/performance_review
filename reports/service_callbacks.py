import pandas as pd
from datetime import datetime, timedelta
import openpyxl

def service_callbacks():
    # Read the Excel file
    df = pd.read_excel("data/Historical Invoice Report.xls")

    service_code_mappings = [
        ("ALL", "Bimonthly"),
        ("SG", "Quarterly"), 
        ("QPC", "Quarterly"),
        ("BIM", "Bimonthly"),
        ("CASTLE", "Bimonthly"),
        ("LS", "Mosquito"),
        ("MOQ", "Mosquito"),
        ("IN2", "Mosquito"),
        ("MOS", "Monthly")
    ]
    
    # Create mapping dictionaries
    code_to_type = {code: type_ for code, type_ in service_code_mappings}
    
    # Convert Work Date to datetime if it isn't already
    df['Work Date'] = pd.to_datetime(df['Work Date'])
    
    # Calculate date one year ago from today
    one_year_ago = datetime.now() - timedelta(days=365)
    
    # Filter for dates within the last year and only mapped service codes
    df_filtered = df[
        (df['Work Date'] >= one_year_ago) & 
        (df['Service Code'].isin(code_to_type.keys()))
    ]
    
    # Add Service Type column
    df_filtered['Service Type'] = df_filtered['Service Code'].map(code_to_type)
    
    # Create pivot table with total accounts and callbacks by service code
    callback_analysis = df_filtered.groupby(['Service Type', 'Service Code']).agg({
        'Account': [
            ('Total Accounts', 'nunique'),
        ],
        'Invoice Type': [
            ('Callbacks', lambda x: (x == 'Call Back').sum())
        ]
    })
    
    # Flatten column names
    callback_analysis.columns = callback_analysis.columns.get_level_values(1)
    
    # Calculate callback percentage per account
    callback_analysis['Callbacks per Account %'] = (callback_analysis['Callbacks'] / callback_analysis['Total Accounts'] * 100).round(1).astype(str) + '%'
    
    # Modify the service type totals to maintain consistent index structure
    service_type_totals = df_filtered.groupby('Service Type').agg({
        'Account': [
            ('Total Accounts', 'nunique'),
        ],
        'Invoice Type': [
            ('Callbacks', lambda x: (x == 'Call Back').sum())
        ]
    })
    
    # Flatten column names for totals
    service_type_totals.columns = service_type_totals.columns.get_level_values(1)
    service_type_totals['Callbacks per Account %'] = (service_type_totals['Callbacks'] / service_type_totals['Total Accounts'] * 100).round(1).astype(str) + '%'
    
    # Create a MultiIndex for the totals to match callback_analysis
    service_type_totals.index = pd.MultiIndex.from_tuples(
        [(idx, 'TOTAL') for idx in service_type_totals.index],
        names=['Service Type', 'Service Code']
    )
    
    # Combine individual codes and totals
    final_analysis = pd.concat([callback_analysis, service_type_totals])
    
    # Sort by Service Type and then by Service Code
    final_analysis = final_analysis.sort_index()
    
    # Create Excel writer object
    with pd.ExcelWriter('outputs/Service Code Callbacks.xlsx', engine='openpyxl') as writer:
        final_analysis.to_excel(writer, sheet_name='Callbacks')
        
        # Get workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Callbacks']
        
        # Create styles
        header_fill = openpyxl.styles.PatternFill(start_color='B8CCE4', end_color='B8CCE4', fill_type='solid')
        bold_font = openpyxl.styles.Font(bold=True)
        
        # Apply header fill
        for cell in worksheet[1]:
            cell.fill = header_fill
            
        # Bold the total rows
        for row in worksheet.iter_rows():
            if row[1].value == 'TOTAL':  # Service Code column
                for cell in row:
                    cell.font = bold_font
    
    return final_analysis

print(service_callbacks())