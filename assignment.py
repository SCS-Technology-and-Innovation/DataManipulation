import pandas as pd
import os
os.chdir('C:/Users/elisa/Desktop')

# load the application data
afile = pd.ExcelFile('applications.xlsx')
a = afile.parse(afile.sheet_names[0])

# load the points data
pfile = pd.ExcelFile('points.xlsx')
p = pfile.parse(pfile.sheet_names[0])

a = a.rename(columns = {'Name': 'Instructor'})
data = pd.merge(a, p, how = 'left')
data = data.drop('Irrelevant', axis = 1)
data = data.fillna(0)

# process
data['highest'] = data.groupby('Course')['Points'].transform('max')
data['chosen'] = data['Points'] == data['highest']
data['Hire'] = [ 'yes' if value else 'no' for value in data['chosen'] ]
                        
# clean out the extras
data = data.drop('highest', axis = 1)
data = data.drop('chosen', axis = 1)

# output
print('Globally\n')
print(data)

print('\nBy course\n')
data.groupby('Course').apply(print) 
