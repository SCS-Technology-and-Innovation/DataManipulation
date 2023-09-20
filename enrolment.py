import numpy as np
import pandas as pd
import datetime as dt
import matplotlib.pyplot as plt

data = pd.read_excel('old_data_new_cohorts.xlsx', sheet_name = 'F23')
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

#plt.grid(color = 'darkgray', zorder = 0)

plt.legend(bbox_to_anchor = (1.05, 1), loc = 'upper right', borderaxespad = 0)

covid = dt.datetime(2020, 3, 1)
scaleai = dt.datetime(2020, 4, 15)
pratic = dt.datetime(2021, 6, 1) # credit
pratic2 = dt.datetime(2022, 6, 1) # non-credit
present = dt.datetime(2023, 9, 20) # NOW
future = present + dt.timedelta(days = 90)
past = dt.datetime(2015, 1, 1)

roffset = dt.timedelta(days = 12)
loffset = dt.timedelta(days = 85)

fact = 'gray'
plan = 'red'

plt.axvline(x = covid, color = 'pink', alpha = 0.5, linewidth = 4)
plt.text(covid - loffset, 300, 'Start of COVID-19', rotation = 90, color = fact)

plt.axvline(x = scaleai, color = 'cyan', alpha = 0.5, linewidth = 4)
plt.text(scaleai + roffset, 250, 'ScaleAI begins', rotation = -90, color = fact)

plt.axvline(x = pratic, color = 'purple', alpha = 0.5, linewidth = 4)
plt.text(pratic + roffset, 200, 'PRATIC credit', rotation = -90, color = fact)

plt.axvline(x = pratic2, color = 'orange', alpha = 0.5, linewidth = 4)
plt.text(pratic2 - loffset, 300, 'PRATIC non-credit', rotation = 90, color = fact)

plt.axvline(x = present, color = 'gray', alpha = 0.5, linewidth = 16)
plt.text(future, 400, 'Forecast', rotation = 0, fontsize = 16)
plt.text(present - 3 * roffset, 400, 'NOW', rotation = 90, fontsize = 12)
plt.text(past, 300, 'Past enrolment data', rotation = 0, fontsize = 20)

ba = dt.datetime(2026, 5, 1)
cit = dt.datetime(2025, 5, 1)
aai = dt.datetime(2025, 6, 1) 

#plt.axvline(x = ba, color = 'green', alpha = 0.5, linewidth = 3)
#plt.text(ba - loffset, 300, 'BA revised', rotation = 90, color = plan)

plt.axvline(x = cit, color = 'yellow', alpha = 0.5, linewidth = 3)
plt.text(cit - loffset, 300, 'CIT modular', rotation = 90, color = plan)

plt.axvline(x = aai, color = 'green', alpha = 0.5, linewidth = 3)
plt.text(aai + roffset, 370, 'AAI online', rotation = -90, color = plan)

plt.title('Technology & Innovation', fontsize = 40)
plt.xlabel('Term')
plt.ylabel('Registrations')



plt.show()
