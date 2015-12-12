# Learn about API authentication here: https://plot.ly/python/getting-started
# Find your api_key here: https://plot.ly/settings/api

import warnings
warnings.filterwarnings("ignore")
from openpyxl import load_workbook
from openpyxl import Workbook
import plotly.plotly as py
import plotly.graph_objs as go

def put_input():
    print("How to input?\nEx.\n    1. 2548-2552\n    2. 2550-2557\n    3. All\n    **(1. & 2.\
The range of input have to more then 2547 and less than 2558.)**\n    **(3. If your input \
is\All, the result will display 2548-2557 and predict of 2558.)**")
    the_input = input("Please put your input. ")
    if the_input == "All":
        return ['2548', '2549', '2550', '2551', '2552', '2553', '2554', '2555', '2556', '2557', '2558']
    else:
        first, last = the_input.split("-")
        lis = [str(i) for i in range(int(first), int(last)+1)]
        return lis

def call_from_result():
    wb = load_workbook('result.xlsx')
    sh = wb['Sheet']
    year = put_input()
    man_lis, woman_lis, total_lis = [], [], []
    for i in year:
        for j in range(3, 14):
            if sh['A'+str(j)].value == i:
                man_lis.append(float(sh['B'+str(j)].value))
                woman_lis.append(float(sh['D'+str(j)].value))
                total_lis.append(float(sh['F'+str(j)].value))
    return man_lis, woman_lis, total_lis, year

def make_a_graph():
    man_lis, woman_lis, total_lis, year_range = call_from_result()
    trace1 = go.Bar(
        x=year_range,
        y=man_lis,
        name='Man',
        marker=dict(
            color='rgb(55, 83, 109)'
        )
    )
    trace2 = go.Bar(
        x=year_range,
        y=woman_lis,
        name='Woman',
        marker=dict(

            color='rgb(26, 118, 255)'
        )
    )
    trace3 = go.Bar(
        x=year_range,
        y=total_lis,
        name='Total',
        marker=dict(
            color='rgb(13, 158, 188)'
        )
    )
    data = [trace1, trace2, trace3]
    layout = go.Layout(
        title='Bar graphs show total deaths from suicide in Thailand since 2548-2557 and predict of 2558.',
        xaxis=dict(
            title='Total since 2548 - 2557 and predict of 2558.',
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
