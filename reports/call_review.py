import requests
import os 
import sys

from pyairtable import Api
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

START_DATE = "2025-01-20"
END_DATE = "2025-01-20"

USERS = [
    ('G', 'USR606009BC8AF41AD25AE582744AA2BCBF', 'service@bettertermite.com')
]

def find_airtable_record(bare_phone_number):
    api = Api(os.getenv("AIRTABLE_API_KEY"))
    table = api.table("appBsty6iukfNnuEK", "tblenVnxR8q8iTGSk")
    formula = f'IF(FIND(LOWER("{bare_phone_number}"), LOWER({{Normalized Phone}})) > 0, TRUE(), FALSE())'
    records = table.all(formula=formula, view="viwQMbNHcF8DMkZz8")
    if len(records) > 0:
        return records[0]
    return None

def format_phone_number(phone_number):
    if len(phone_number) == 10:
        return f"({phone_number[:3]}) {phone_number[3:6]}-{phone_number[6:]}"
    elif len(phone_number) == 11:
        return f"({phone_number[:4]}) {phone_number[4:7]}-{phone_number[7:]}"
    else:
        return phone_number
    
def airtable_link(record):
    if record:
        # Return a tuple of the URL and display text instead of just the URL
        return (f"https://airtable.com/appBsty6iukfNnuEK/pagzNijXAeXwCMFPu/{record['id']}", "Airtable Link")
    return ("", "")

def get_call_tracking_metrics(start_date, end_date, user):
    name, rep, email = user
    url = "https://api.calltrackingmetrics.com/api/v1/accounts/533886/calls"
    params = {
        "start_date": start_date,
        "end_date": end_date,
        "agent_id": rep,
        "limit": 100
    }
    
    headers = {
        "Authorization": f"Basic {os.getenv('CALL_TRACKING_METRICS_API_KEY')}"
    }
    
    response = requests.get(url, headers=headers, params=params)
    
    if response.status_code == 200:
        data = response.json()
        formatted_data = []
        for call in data['calls']:
            record = find_airtable_record(call["caller_number_bare"])
            formatted_data.append({
                "ID": call["id"],
                "Rep": name,
                "Direction": "Inbound" if call["direction"] == "inbound" else "Outbound",
                "Call Type": call.get("custom_fields", {}).get("classification", ""),
                "Summary": call.get("custom_fields", {}).get("summary", ""),
                "Date": pd.to_datetime(call["called_at"]).strftime("%a %b %-d, %Y"),
                "Duration": call["talk_time"],
                "Bare Number": format_phone_number(call["caller_number_bare"]),
                "Source": call["source"],
                "Recording": "Recording Link" if call.get("audio") else "",
                "Review": "Review Link",
                "Airtable_Record": airtable_link(record),
                "Close Status": record.get("fields", {}).get("Close Status", "")
            })

        df = pd.DataFrame(formatted_data)
        # remove calls less than 5 seconds
        df = df[df["Duration"] > 5]
        return df
    else:
        print(f"Error: {response.status_code}")
        print(response.text)
        return None

if __name__ == "__main__":
    # Create an Excel writer object
    output_path = 'outputs/call_review.xlsx'
    os.makedirs('outputs', exist_ok=True)
    excel_writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    
    # Process each user and add their data to a separate worksheet
    for user in USERS:
        df = get_call_tracking_metrics(START_DATE, END_DATE, user)
        if df is not None:
            # Write the dataframe to a worksheet named after the user
            df.to_excel(excel_writer, sheet_name=user[0], index=False)
            
            # Get the worksheet object
            worksheet = excel_writer.sheets[user[0]]
            
            # Set a larger default row height (60 instead of 40)
            worksheet.set_default_row(height=60)
            
            # Explicitly set the header row height
            worksheet.set_row(0, 30)  # First row (header) can be shorter
            
            # Define column widths
            column_widths = {
                'A': 15,  # call_id
                'B': 15,  # rep
                'C': 15,  # direction
                'D': 15,  # call type
                'E': 80,  # summary
                'F': 20,  # date
                'G': 10,  # duration
                'H': 15,  # bare_number
                'I': 30,  # source
                'J': 15,  # recording
                'K': 15,  # review_link
                'L': 15,  # airtable_link
                'M': 15,  # close_status
            }
            
            # Create formats
            wrap_format = excel_writer.book.add_format({
                'text_wrap': True,
                'valign': 'top'  # Align text to top of cell
            })
            
            # Add new format for bigger font
            big_font_format = excel_writer.book.add_format({
                'text_wrap': True,
                'valign': 'top',
                'font_size': 12  # Increased font size
            })
            
            link_format = excel_writer.book.add_format({
                'text_wrap': True,
                'color': 'blue',
                'underline': 1,
                'align': 'center',    # Add horizontal center alignment
                'valign': 'center'    # Change from top to center vertical alignment
            })

            # Add color formats for direction
            inbound_format = excel_writer.book.add_format({
                'text_wrap': True,
                'valign': 'top',
                'bg_color': '#C6EFCE',  # Light green
                'font_color': '#006100',  # Dark green
                'font_size': 12  # Increased font size
            })
            
            outbound_format = excel_writer.book.add_format({
                'text_wrap': True,
                'valign': 'top',
                'bg_color': '#FFC7CE',  # Light red
                'font_color': '#9C0006',  # Dark red
                'font_size': 12  # Increased font size
            })
            
            # Set column widths and text wrapping
            for col, width in column_widths.items():
                if col in ['B', 'C', 'D', 'H', 'M']:  # Columns for Rep, Direction, Call Type, Bare Number, Close Status
                    worksheet.set_column(f'{col}:{col}', width, big_font_format)
                else:
                    worksheet.set_column(f'{col}:{col}', width, wrap_format)
            
            # Add hyperlinks and direction formatting for each row
            for row_num, row in enumerate(df.itertuples(), start=1):
                # Format direction cell
                if row.Direction == "Inbound":
                    worksheet.write(row_num, 2, "Inbound", inbound_format)
                else:
                    worksheet.write(row_num, 2, "Outbound", outbound_format)
                
                if row.Recording:  # Only add link if there's a recording
                    worksheet.write_url(row_num, 9, row._asdict()['Recording'], 
                                     string='Recording Link', cell_format=link_format)
                worksheet.write_url(row_num, 10, 
                                 f"https://app.calltrackingmetrics.com/calls#callNav=caller_transcription&callId={row.ID}", 
                                 string='Review Link', cell_format=link_format)
                
                # Add Airtable hyperlink if URL exists
                airtable_url, display_text = row._asdict()['Airtable_Record']
                if airtable_url:
                    worksheet.write_url(row_num, 11, airtable_url, 
                                     string=display_text, cell_format=link_format)
    
    # Save and close the Excel file
    excel_writer.close()
    
    # Open the Excel file
    if os.name == 'nt':  # Windows
        os.system(f'start excel "{output_path}"')
    elif os.name == 'posix':  # macOS or Linux
        if os.system('which open') == 0:  # macOS
            os.system(f'open "{output_path}"')
        else:  # Linux
            os.system(f'xdg-open "{output_path}"')