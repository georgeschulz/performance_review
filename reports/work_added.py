import pandas as pd

def work_added():
    df = pd.read_csv("data/Monthly Invoice Report.csv")
    
    # Calculate totals for each service class
    renewal_total = df[df['Service Class'] == 'RENEWAL']['Total'].sum()
    rtermite_total = df[df['Service Class'] == 'RTERMITE']['Total'].sum()
    
    # Calculate combined total
    total_added = renewal_total + rtermite_total
    
    # Create and return report dictionary with currency formatting
    report = {
        'Renewals': f'${renewal_total:,.2f}',
        'Termite Monitoring': f'${rtermite_total:,.2f}', 
        'Total Added': f'${total_added:,.2f}'
    }
    
    # Save report to CSV
    pd.DataFrame(report, index=[0]).to_csv('outputs/Work Added.csv', index=False)