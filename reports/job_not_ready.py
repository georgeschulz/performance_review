import pandas as pd

def job_not_ready_report(excluded_techs=[]):
    # Read the CSV file
    df = pd.read_csv("data/Work Completion - Stop Data at Start of Day.csv")
    
    # Define pretreat service codes
    pretreat_codes = ["PW", "PT", "PTE", "PWF"]
    
    # Create the report
    report = pd.DataFrame()
    
    # Count total pretreats per tech
    report['Total Pretreats'] = df[df['Service Code'].isin(pretreat_codes)].groupby('Tech 1').size()
    
    # Count job not ready pretreats
    not_ready = df[
        (df['Service Code'].isin(pretreat_codes)) & 
        (df['Reschedule Reason'] == 'Job Not Ready')
    ].groupby('Tech 1').size()
    
    report['Jobs Not Ready'] = not_ready.fillna(0)
    
    # Calculate the ratio as percentage
    report['Not Ready %'] = (report['Jobs Not Ready'] / report['Total Pretreats'] * 100).round(1).astype(str) + '%'
    
    # Add totals row
    totals = pd.Series({
        'Total Pretreats': report['Total Pretreats'].sum(),
        'Jobs Not Ready': report['Jobs Not Ready'].sum(),
        'Not Ready %': f"{(report['Jobs Not Ready'].sum() / report['Total Pretreats'].sum() * 100):.1f}%"
    }, name='TOTALS')
    
    report = pd.concat([report, totals.to_frame().T])
    
    # Save to CSV
    report.to_csv('outputs/Job Not Ready.csv')
    
    return report
