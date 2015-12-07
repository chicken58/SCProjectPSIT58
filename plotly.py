import plotly.plotly as py
from plotly.graph_objs import *
py.sign_in('emptygiiiz', 'dmughevmvv')
trace1 = Bar(
    x=[2548, 2549, 2550, 2551, 2552, 2553, 2554, 2555, 2556, 2557],
    y=[99.56, 93.38, 95.6, 93.96, 94.45, 93.65, 95.37, 97.56, 97.88, 97.08],
    name='',
    uid='743f82'
)
trace2 = Bar(
    x=[2548, 2549, 2550, 2551, 2552, 2553, 2554, 2555, 2556, 2557],
    y=[29.11, 24.08, 25.73, 27.5, 26.93, 26.41, 27.54, 28.79, 26.08, 27.13],
    name='',
    uid='b1b664'
)
trace3 = Bar(
    x=[2548, 2549, 2550, 2551, 2552, 2553, 2554, 2555, 2556, 2557],
    y=[128.67, 117.45, 121.33, 121.46, 121.38, 120.06, 122.91, 126.35, 123.96, 124.21],
    name='',
    uid='4f5d72'
)
data = Data([trace1, trace2, trace3])
layout = Layout(
    autosize=True,
    height=536,
    width=1121,
    xaxis=XAxis(
        autorange=True,
        range=[2547.5, 2557.5],
        type='linear'
    ),
    yaxis=YAxis(
        autorange=True,
        range=[0, 135.44210526315788],
        type='linear'
    )
)
fig = Figure(data=data, layout=layout)
plot_url = py.plot(fig)
