"""Plot the graph in plot.ly with result.xlsx ."""
import warnings
warnings.filterwarnings("ignore")
from openpyxl import load_workbook
from openpyxl import Workbook
import plotly.plotly as py
import plotly.graph_objs as go

def put_input():
    """Get a input from user and repeat get input if the input is invalid."""
    print("How to input?\nEx.\n    1. 2548-2552\n    2. 2550-2557\n    3. 2557-2557\n    4. All\n    \
**(1. & 2. The range of input have to more than 2547 and less than 2558.)**\n    **(3. Show one \
year.)**\n    **(4. If your input is\All, the result will display 2548-2557 and predict of 2558.)**")
    while True: # repeat if a input is invalid.
        the_input = input("Please put your input : ")
        if the_input == "all" or the_input == "All" or the_input == "ALL":
            return ['2548', '2549', '2550', '2551', '2552', '2553', '2554', '2555', '2556', '2557', '2558']
        elif "-" not in the_input:
            print("Your input is invalid, please put your input again.")
        else:
            first, last = the_input.split("-")
            if 2548 <= int(first) <= 2557 and 2548 <= int(last) <= 2557 and int(last) >= int(first):
                lis = [str(i) for i in range(int(first), int(last)+1)]
                return lis
            else: print("Your input is invalid, plase put your input again.")

def call_from_result():
    """Call information from result.xlsx and return list of information."""
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

def title(year_range):
    """Return title that match the input."""
    if year_range == [str(i) for i in range(2548, 2559)]:
        pre_title = "since 2548-2557 and predict of 2558."
    elif len(year_range) == 1:
        pre_title = "of " + year_range[0] + "."
    else:
        first, last = year_range[0], year_range[-1]
        pre_title = "since "+ first + "-" + last + "."
    title_1 = "Bar graphs show total deaths from suicide in Thailand " + pre_title
    title_2 = "Total " + pre_title
    return title_1, title_2

def make_a_graph():
    """Run and display graph in plot.ly ."""
    man_lis, woman_lis, total_lis, year_range = call_from_result()
    title_1, title_2 = title(year_range)
    trace1 = go.Bar(
        x=year_range,
        y=man_lis,
        name='Man',
        marker=dict(
            color='rgb(55, 83, 109)'
        )
    ) # Man graph
    trace2 = go.Bar(
        x=year_range,
        y=woman_lis,
        name='Woman',
        marker=dict(
            color='rgb(26, 118, 255)'
        )
    ) # Woman graph
    trace3 = go.Bar(
        x=year_range,
        y=total_lis,
        name='Total',
        marker=dict(
            color='rgb(13, 158, 188)'
        )
    ) # Total graph
    data = [trace1, trace2, trace3]
    layout = go.Layout(
        title=title_1,
        xaxis=dict(
            title=title_2,
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

