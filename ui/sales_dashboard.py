import dash
from dash import dcc, html
import pandas as pd
import plotly.graph_objects as go
from dash.dependencies import Input, Output

# Initialize the Dash app
app = dash.Dash(__name__)

def load_data():
    """Load and prepare the close rate data from staging"""
    df = pd.read_csv('../staging/Close Rate.csv')
    # Convert percentage strings to floats
    df['Close Rate'] = df['Close Rate'].str.rstrip('%').astype(float)
    df['Disqualify Rate'] = df['Disqualify Rate'].str.rstrip('%').astype(float)
    # Convert Day to datetime
    df['Day'] = pd.to_datetime(df['Day'])
    return df

def group_data(df, group_by='day'):
    """Group the data by the specified period"""
    if group_by == 'day':
        df['Period'] = df['Day'].dt.strftime('%Y-%m-%d')
    elif group_by == 'week':
        df['Period'] = df['Day'].dt.strftime('%Y-W%U')
    elif group_by == 'month':
        df['Period'] = df['Day'].dt.strftime('%Y-%m')
    else:  # year
        df['Period'] = df['Day'].dt.strftime('%Y')
    
    # Group by Period and Salesperson
    grouped = df.groupby(['Period', 'Salesperson']).agg({
        'Won: Recurring': 'sum',
        'Won: One Time': 'sum',
        'Lost': 'sum',
        'Disqualified': 'sum',
        'Open': 'sum'
    }).reset_index()
    
    # Recalculate rates after grouping
    total_opportunities = grouped['Won: Recurring'] + grouped['Won: One Time'] + grouped['Lost']
    grouped['Close Rate'] = (grouped['Won: Recurring'] / total_opportunities * 100).round(1)
    total_closed = total_opportunities + grouped['Disqualified']
    grouped['Disqualify Rate'] = (grouped['Disqualified'] / total_closed * 100).round(1)
    
    return grouped

def create_stacked_bar(df, selected_salespeople=None):
    """Create a stacked bar chart for the close rate data"""
    if selected_salespeople:
        df = df[df['Salesperson'].isin(selected_salespeople)]
    
    # Handle empty dataframe case
    if df.empty:
        fig = go.Figure()
        fig.add_annotation(text="No data available", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        return fig
    
    fig = go.Figure()
    
    # Add bars for each category
    categories = ['Won: Recurring', 'Won: One Time', 'Lost', 'Disqualified']
    colors = ['#2ecc71', '#3498db', '#e74c3c', '#95a5a6']
    
    for cat, color in zip(categories, colors):
        fig.add_trace(go.Bar(
            name=cat,
            x=[df['Period'], df['Salesperson']],
            y=df[cat],
            marker_color=color
        ))

    # Calculate total height for each bar
    df['Total_Height'] = df[categories].sum(axis=1)
    
    # Add close rate as text trace above the bars
    fig.add_trace(go.Scatter(
        x=[f"{row['Period']} - {row['Salesperson']}" for _, row in df.iterrows()],
        y=df['Total_Height'],
        text=[f"{rate:.1f}%" for rate in df['Close Rate']],
        mode='text',
        textposition='top center',
        showlegend=False,
        textfont=dict(size=10),
        yaxis='y'
    ))

    fig.update_layout(
        barmode='stack',
        title='Sales Outcomes by Period and Salesperson',
        xaxis_title='Period / Salesperson',
        yaxis_title='Number of Leads',
        legend_title='Outcome',
        height=600,
        # Add padding to the top of the chart to accommodate labels
        yaxis=dict(
            range=[0, df['Total_Height'].max() * 1.1]  # Add 10% padding to the top
        )
    )
    
    return fig

# App layout
app.layout = html.Div([
    html.H1('Inbound Close Rate Dashboard'),
    
    # Control panel
    html.Div([
        # Period selector
        html.Div([
            html.Label('Group by:'),
            dcc.RadioItems(
                id='group-by',
                options=[
                    {'label': 'Day', 'value': 'day'},
                    {'label': 'Week', 'value': 'week'},
                    {'label': 'Month', 'value': 'month'},
                    {'label': 'Year', 'value': 'year'}
                ],
                value='month',
                inline=True
            )
        ], style={'margin': '20px'}),
        
        # Salesperson filter
        html.Div([
            html.Label('Select Salespeople:'),
            dcc.Dropdown(
                id='salesperson-filter',
                multi=True,
                placeholder='Select salespeople...'
            )
        ], style={'width': '50%', 'margin': '20px'})
    ]),
    
    # Graph
    dcc.Graph(id='stacked-bar-chart'),
    
    # Summary statistics
    html.Div(id='summary-stats', style={'margin': '20px'})
])

# Callbacks
@app.callback(
    [Output('salesperson-filter', 'options'),
     Output('salesperson-filter', 'value')],
    [Input('stacked-bar-chart', 'id')]
)
def populate_dropdown(_):
    df = load_data()
    salespeople = sorted(df['Salesperson'].unique())
    options = [{'label': sp, 'value': sp} for sp in salespeople]
    return options, salespeople

@app.callback(
    [Output('stacked-bar-chart', 'figure'),
     Output('summary-stats', 'children')],
    [Input('salesperson-filter', 'value'),
     Input('group-by', 'value')]
)
def update_graph(selected_salespeople, group_by):
    df = load_data()
    
    if not selected_salespeople:
        selected_salespeople = []
    
    # Group the data by selected period
    grouped_df = group_data(df, group_by)
    
    # Create the stacked bar chart
    fig = create_stacked_bar(grouped_df, selected_salespeople)
    
    # Calculate summary statistics
    filtered_df = grouped_df if not selected_salespeople else grouped_df[grouped_df['Salesperson'].isin(selected_salespeople)]
    avg_close_rate = filtered_df['Close Rate'].mean()
    avg_disqualify_rate = filtered_df['Disqualify Rate'].mean()
    
    summary = html.Div([
        html.H3('Summary Statistics'),
        html.P(f'Average Close Rate: {avg_close_rate:.1f}%'),
        html.P(f'Average Disqualify Rate: {avg_disqualify_rate:.1f}%')
    ])
    
    return fig, summary

if __name__ == '__main__':
    app.run_server(debug=True)