import pandas as pd

def production_report():
    # Read the CSV file
    df = pd.read_csv("data/Monthly Invoice Report.csv")
    
    # Combine First Name and Last Name into a single column
    df['Technician'] = df['First Name'] + ' ' + df['Last Name']
    
    # Group by Tech and Work Date, calculate daily totals
    daily_production = df.groupby(['Tech', 'Work Date'])['Total'].sum().reset_index()
    
    # Calculate total orders per tech (total rows)
    total_orders = df.groupby('Tech').size().reset_index(name='Total Orders')
    
    # Calculate total stops per tech (unique Account+WorkDate combinations)
    total_stops = df.groupby('Tech').agg({
        'Account': lambda x: len(pd.unique(x.astype(str) + df.loc[x.index, 'Work Date'].astype(str)))
    }).reset_index()
    total_stops.columns = ['Tech', 'Total Stops']
    
    # Calculate total production per tech
    tech_totals = daily_production.groupby('Tech').agg({
        'Total': ['sum', lambda x: x.count()]  # sum for total amount, count for number of days
    }).reset_index()
    
    # Rename columns for clarity
    tech_totals.columns = ['Technician', 'Total Production', 'Days Worked']
    
    # Merge in orders and stops
    tech_totals = tech_totals.merge(total_orders, left_on='Technician', right_on='Tech').drop('Tech', axis=1)
    tech_totals = tech_totals.merge(total_stops, left_on='Technician', right_on='Tech').drop('Tech', axis=1)
    
    # Calculate averages
    tech_totals['Average Production Per Day'] = tech_totals['Total Production'] / tech_totals['Days Worked']
    tech_totals['Average Orders Per Day'] = tech_totals['Total Orders'] / tech_totals['Days Worked']
    tech_totals['Average Stops Per Day'] = tech_totals['Total Stops'] / tech_totals['Days Worked']
    
    # Sort by Total Production in descending order
    tech_totals = tech_totals.sort_values('Total Production', ascending=False)
    
    # Format currency columns
    tech_totals['Total Production'] = tech_totals['Total Production'].apply(lambda x: f"${x:,.2f}")
    tech_totals['Average Production Per Day'] = tech_totals['Average Production Per Day'].apply(lambda x: f"${x:,.2f}")
    
    # Round average metrics to 2 decimal places
    tech_totals['Average Orders Per Day'] = tech_totals['Average Orders Per Day'].round(2)
    tech_totals['Average Stops Per Day'] = tech_totals['Average Stops Per Day'].round(2)
    
    tech_totals.to_csv("outputs/Production.csv", index=False)