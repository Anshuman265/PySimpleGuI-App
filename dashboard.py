from distutils.log import debug
from pydoc import classname
from unicodedata import category
import dash
from dash.dependencies import Output, Input
import dash_core_components as dcc
import dash_html_components as html
import plotly.express as px
import pandas as pd


external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

df = pd.read_excel("excel_file_2.xlsx")


def show_pie_dashboard(df, name, value):
    fig_pie = px.pie(data_frame=df, names=name, values=value)
    fig_pie.show()


def show_bar_dashboard(df, x, y):
    fig_bar = px.bar(data_frame=df, x=x, y=y)
    fig_bar.show()


def show_hist_dashboard(df, x, y):
    fig_hist = px.histogram(data_frame=df, x=x, y=y)
    fig_hist.show()


fig = px.histogram(data_frame=df,
                   x='Year of Study',)
fig.update_xaxes(type='category')
"""
# Show the pie dashboard
show_pie_dashboard(df, 'Dept', 'Amount Requested')
# Show the bar dashboard
show_bar_dashboard(df, 'Dept', 'Amount Requested')
# Show the histogram dashboard
show_hist_dashboard(df, 'Dept', 'Amount Requested')
"""

# Interactive Graphing with Dash
app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

app.layout = html.Div(children=[
    html.Div([
        html.Div([
            html.Div([
                html.H4("Department Wise", style={
                    'margin': '0 auto', 'text-align': 'center', }),
                dcc.Graph(id='pie-graph', figure=px.pie(data_frame=df,
                                                        names='Dept', values='Amount Requested'))
            ], className='row'),
            html.Div([
                html.H4("Student Category", style={
                    'margin': '0 auto', 'text-align': 'center', }),
                dcc.Graph(id='pie-graph2', figure=px.pie(data_frame=df,
                                                         names='Category', values='Amount Requested'))
            ],  className='row'),
        ], className='six columns'),
        html.Div([
            html.Div([
                html.H4("Student Year Wise", style={
                    'margin': '0 auto', 'text-align': 'center'}),
                dcc.Graph(id='bar-chart2', figure=fig),
            ], className='row'),
            html.Div([
                html.H4("Region Wise", style={
                    'margin': '0 auto', 'text-align': 'center'}),
                dcc.Graph(id='histogram1', figure=px.histogram(data_frame=df,
                                                               x='Region',)),
            ], className='row'),
        ], className='six columns'),
    ])

])
if __name__ == '__main__':
    app.run_server(debug=False)

    """
    html.Div([
        html.Div([
            html.H3("Region Wise", style={
                'margin': '0 auto', 'text-align': 'center'}),
            dcc.Graph(id='choropleth-map', figure=px.choropleth(data_frame=df,
                      locations="Region", title="Region Wise Student Distribution")),
        ]),
    ], className='row'),
    """
