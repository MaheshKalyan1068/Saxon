# -*- coding: utf-8 -*-
"""
Created on Tue Jul 28 19:32:18 2020

@author: Manomay
"""


import dash
import dash_core_components as dcc
import dash_html_components as html
import dash_table
import pandas as pd

df = pd.read_csv(r'C:\Users\Manomay\Desktop\Saxon\status.csv')

app = dash.Dash(__name__)

df = pd.read_csv(r'C:\Users\Manomay\Desktop\Saxon\status.csv')

#df = pd.read_csv('https://raw.githubusercontent.com/plotly/datasets/master/solar.csv')

app.layout = html.Div([
      html.H4('SAXON Claim Bot'),
      dcc.Interval('graph-update', interval = 5, n_intervals = 0),
      dash_table.DataTable(
          id = 'table',
          data = df.to_dict('records'),
          columns=[{"name": i, "id": i} for i in df.columns])])

@app.callback(
        dash.dependencies.Output('table','data'),
        [dash.dependencies.Input('graph-update', 'n_intervals')])
def updateTable(n):
    df = pd.read_csv(r'C:\Users\Manomay\Desktop\Saxon\status.csv')
    return df.to_dict('records')

if __name__ == '__main__':
     app.run_server(debug=True)