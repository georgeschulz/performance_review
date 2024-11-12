import pandas as pd

def ar_report():
    df = pd.read_csv("data/Bill To Balances.csv")
    # CINDY
    residential = df[(df['Type'] == 'R') & (df['Balance'] >= 55)] # 55 set here to exclude all in ones 
    total_60_89 = residential['60-89'].sum()
    total_90_plus = residential['90'].sum()

    total_residential_60_plus = total_60_89 + total_90_plus

    # BIANCA REGULAR COMMERCIAL
    commercial = df[(df['Type'] == 'C') & (df['Balance'] > 0)]
    commercial_60_90 = commercial['60-89'].sum()
    commercial_90_plus = commercial['90'].sum()

    total_commercial_60_plus = commercial_60_90 + commercial_90_plus

    commercial_longterm = df[(df['Type'] == 'CL') & (df['Balance'] > 0)]
    commercial_90_plus = commercial_longterm['90'].sum()

    data = {
        'Residential 60+': f'${total_residential_60_plus:,.2f}',
        'Commercial 60+': f'${total_commercial_60_plus:,.2f}',
        'Commercial Longterm': f'${commercial_90_plus:,.2f}'
    }

    pd.DataFrame(data, index=[0]).to_csv("outputs/AR Report.csv", index=False)