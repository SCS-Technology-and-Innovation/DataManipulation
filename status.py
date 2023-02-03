# install with pip: xlrd

import pandas as pd

classlist = pd.read_excel('InternalExpandedClassListExport.xls')

courseNames = {}
studentNames = {}
emails = {}
sections = {}

from collections import defaultdict

registrations = defaultdict(set)
adm = dict()
programs = defaultdict(set)

rows = classlist.shape[0]
#print('Processing', rows, 'rows of registration data')

foundCode = False
listing = False

# which month is which term (student registration date)
terms = {
    '01': 'W',
    '02': 'W', '03': 'S', '04': 'S', '05': 'S',
    '06': 'F', '07': 'F', '08': 'F', '09': 'F',
    '10': 'W', '11': 'W', '12': 'W'
} 

def term(date):
    parts = date.split('-')
    year = parts[0][2:]
    month = parts[1]
    termLetter = terms[month]
    return f'{termLetter}{year}'
    
for row in range(rows):
    firstCell = str(classlist.iloc[row, 0]).strip()
    if len(firstCell) > 0 and firstCell != 'nan':
        if listing: # student data
            studentCode = str(classlist.iloc[row, 6]).strip()
            studentNames[studentCode] = str(classlist.iloc[row, 5]).strip()
            timestamp = str(classlist.iloc[row, 29]).strip()
            date = timestamp.split()[0]
            registrations[(studentCode, courseCode)].add(term(date))
        if foundCode: # catch the course name
            courseNames[courseCode] = firstCell
            foundCode = False
        else:
            parts = firstCell.split()
            if len(parts) == 4 and parts[2] == '-':
                letterCode = parts[0]
                numberCode = parts[1]
                sectionNumber = parts[3]
                courseCode = f'{letterCode}{numberCode}'
                sectionCode = f'{courseCode}-{sectionNumber}'
                foundCode = True
            elif firstCell == 'Last Name':
                listing = True
    else: # nothing was there
        listing = False

admissions = pd.read_excel('applicationStatus.xls')
rows = admissions.shape[0]
#print('Processing', rows, 'rows of admissions data')

PDC = 'Professional Development Certificate in '
skip = len(PDC)

abbrv = { 'Applied Artificial Intelligence': 'AAI',
          'Data Analytics for Business': 'DAB',
          'Data Science and Machine Learning': 'DSML' }

for row in range(rows):
    firstCell = str(admissions.iloc[row, 0]).strip()
    studentID = str(admissions.iloc[row, 1]).strip()
    if len(studentID) > 0 and studentID[0] == 'X':
        status = str(admissions.iloc[row, 4]).strip()
        if status == 'Admitted':
            address = str(admissions.iloc[row, 12]).strip()
            if studentID not in studentNames:
                nameparts = str(admissions.iloc[row, 0]).strip().split(', ')
                name = f'{nameparts[0]} {nameparts[1]}'
                studentNames[studentID] = name
            when = str(admissions.iloc[row, 8]).strip().split()[0]
            ym = '/'.join(when.split('-')[:-1])[2:]
            emails[studentID] = address
            adm[pID].add((studentID, ym))
            programs[studentID].add((pID, ym))
    else:
        if len(firstCell) > 0 and firstCell != 'nan':
            if PDC in firstCell:
                programName = ' '.join(firstCell[skip:].split()[:-1])
                if programName in abbrv:
                    pID = abbrv[programName]
                    adm[pID] = set()

# mandatory or not
contents = {
    'DAB' : [ ('YCBS256', True),
              ('YCBS260', True),
              ('YCBS261', True),
              ('YCBS262', True),
              ('YCBS299', True) ] ,
    'DSML' : [ ('YCBS255', True),
               ('YCBS256', True),
               ('YCBS257', True),
               ('YCBS258', True),
               ('YCBS299', True) ],
    'AAI' : [ ('YCNG228', True),
              ('YCNG229', True),
              ('YCNG230', False),
              ('YCNG231', False),
              ('YCNG232', False),
              ('YCNG233', False),
              ('YCNG234', False),
              ('YCNG235', False) ]
} 

def courselist(courses):
    result = []
    for (c, m) in courses:
        a = '*' if m else ''
        result.append(f'{c}{a}')
    return ','.join(result)

header = 'Program,Name,ID,Email,Admitted,'

tabs = {}
for program in adm: # per program
    filename = f'{program}.csv'
    with open(filename, 'w') as target:
        courses = contents[program]
        programheader = header + courselist(courses) + ',Count'
        print(programheader, file = target)
        for (student, ad) in adm[program]: # who and when was admitted
            e = emails.get(student, 'Unavailable')
            n = studentNames.get(student, 'Unavailable')
            rd = []
            completed = 0
            for (course, mand) in courses:
                reg = registrations[(student, course)]
                if len(reg) == 0:
                    s = '---'
                else:
                    s = ' '.join(reg)
                    completed += 1
                rd.append(s)
            cs = ','.join(rd)
            print(f'{program},{n},{student},{e},{ad},{cs},{completed}', file = target)
    tabs[program] = pd.read_csv(filename)

    with pd.ExcelWriter('status.xlsx') as target:
        for program in tabs:
            tabs[program].to_excel(target, sheet_name = program, index = False)

