import numpy as np
import pandas as pd
import datetime as dt
import matplotlib.pyplot as plt

data = pd.read_excel('old_data_new_cohorts.xlsx')
pr = data['Program']
kind = data['Kind']
data.fillna(0)
values = (data.iloc[:, 3:]).T # transpose
values.columns = pr

plt.rcParams['figure.figsize'] = [ 15, 8 ]
plt.rcParams['figure.dpi'] = 140
plt.rcParams['figure.facecolor'] = 'black'
plt.rcParams['axes.facecolor'] = 'black'
plt.rcParams['savefig.facecolor'] = 'black'
plt.rcParams['text.color'] = 'white'
plt.rcParams['axes.labelcolor'] = 'white'
plt.rcParams['xtick.color'] = 'white'
plt.rcParams['ytick.color'] = 'white'

values.plot.area(cmap = 'Paired', linewidth = 0)
plt.legend(bbox_to_anchor = (1.05, 1), loc = 'upper right', borderaxespad = 0)

covid = dt.datetime(2020, 3, 1)
scaleai = dt.datetime(2020, 4, 15)
pratic = dt.datetime(2021, 1, 1) # guess
present = dt.datetime(2023, 1, 1)
future = dt.datetime(2023, 6, 1)
past = dt.datetime(2015, 1, 1)

roffset = dt.timedelta(days = 10)
loffset = dt.timedelta(days = 60)

plt.axvline(x = covid, color = 'pink', alpha = 0.5, linewidth = 4)
plt.text(covid - loffset, 300, 'Start of COVID-19', rotation = 90)

plt.axvline(x = scaleai, color = 'cyan', alpha = 0.5, linewidth = 4)
plt.text(scaleai + roffset, 250, 'ScaleAI begins', rotation = -90)

plt.axvline(x = pratic, color = 'purple', alpha = 0.5, linewidth = 4)
plt.text(pratic + roffset, 200, 'PRATIC begins', rotation = -90)

plt.axvline(x = present, color = 'gray', alpha = 0.5, linewidth = 10)
plt.text(future, 400, 'Forecast', rotation = 0, fontsize = 24)
plt.text(past, 300, 'Past enrolment data', rotation = 0, fontsize = 24)

plt.title('Technology & Innovation')
plt.xlabel('Term')
plt.ylabel('Registrations')



plt.show()
