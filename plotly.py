# Learn about API authentication here: https://plot.ly/python/getting-started
# Find your api_key here: https://plot.ly/settings/api

import warnings
warnings.filterwarnings("ignore")
from openpyxl import load_workbook
from openpyxl import Workbook
import plotly.plotly as py
import plotly.graph_objs as go

def make_a_graph():
    first, last = html_input()
    
    trace1 = go.Bar(
        x=[2548, 2549, 2550, 2551, 2552, 2553, 2554, 2555, 2556, 2557],
        y=[99.56, 93.38, 95.6, 93.96, 94.45, 93.65, 95.37, 97.56, 97.88, 97.08],
        name='Man',
        marker=dict(
            color='rgb(55, 83, 109)'
        )
    )
    trace2 = go.Bar(
        x=[2548, 2549, 2550, 2551, 2552, 2553, 2554, 2555, 2556, 2557],
        y=[29.11, 24.08, 25.73, 27.5, 26.93, 26.41, 27.54, 28.79, 26.08, 27.13],
        name='Woman',
        marker=dict(
            color='rgb(26, 118, 255)'
        )
    )
    trace3 = go.Bar(
        x=[2548, 2549, 2550, 2551, 2552, 2553, 2554, 2555, 2556, 2557],
        y=[128.67, 117.45, 121.33, 121.46, 121.38, 120.06, 122.91, 126.35, 123.96, 124.21],
        name='Total',
        marker=dict(
            color='rgb(13, 158, 188)'
        )
    )
    data = [trace1, trace2, trace3]
    layout = go.Layout(
        title='Bar graphs show total deaths from suicide in Thailand since 2548-2557.',
        xaxis=dict(
            title='Total since 2548 - 2557.',
            titlefont=dict(
                size=16,
                color='rgb(107, 107, 107)'
            ),
            tickfont=dict(
                size=14,
                color='rgb(107, 107, 107)'
            )
        ),
        yaxis=dict(
            title='Number of suicides per million population.',
            titlefont=dict(
                size=16,
                color='rgb(107, 107, 107)'
            ),
            tickfont=dict(
                size=14,
                color='rgb(107, 107, 107)'
            )
        ),
        legend=dict(
            x=0,
            y=1.0,
            bgcolor='rgba(255, 255, 255, 0)',
            bordercolor='rgba(255, 255, 255, 0)'
        ),
        barmode='group',
        bargap=0.15,
        bargroupgap=0.1
    )
    fig = go.Figure(data=data, layout=layout)
    plot_url = py.plot(fig, filename='style-bar')

make_a_graph()
