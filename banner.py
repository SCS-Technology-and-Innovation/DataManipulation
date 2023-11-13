import pandas as pd
banner = pd.read_excel('banner.xlsb', engine = 'pyxlsb')
testing = True
if testing:
    print(banner.shape)
    header = list(banner.columns.values) 
    print(header[5:10])
    for (field, pos) in zip(header, range(len(header))):
        if 'PRIM' in field and 'DEG' in field:
            print(pos, field)
    student = 10000
    print(banner.iat[student, 2])

from collections import defaultdict
students = defaultdict(list)
for pos, data in banner.iterrows():
    sid = data[0]
    term = int(data[1])
    cc = data[4] + data[6]
    deg = data[98]
    fac = data[95]
    students[sid].append((term, cc, deg, fac))

print(len(students))

indep = dict()
for s in students:
    keep = False
    for t in students[s]:
        if t[2] == 'No Degree' or t[3] != 'CE':
            keep = True
            break
    if keep:
        indep[s] = students[s]
 

print(len(indep))

change = dict()
for s in indep:
    keep = False
    for t in indep[s]:
        if t[2] != 'No Degree' and t[3] == 'CE':
            keep = True
            break
    if keep:
        change[s] = indep[s]


print(len(change))


programs = defaultdict(list)
for s in change:
    for (t, c, p, f) in change[s]:
        programs[p].append((s, t, c))

print(len(programs))


for p in programs:
    if 'Cert' in p:
        print(p, len(programs[p]))


courses = defaultdict(dict)
for s in change:
    prev = None
    cumul = []
    for data in change[s]:
        p = data[2] # look at the program
        if prev is not None and prev != p: # it has changed
            if 'Cert' in p: # it is a certificate            
                for old in cumul:
                    oc = old[1] # the course
                    if p not in courses[oc]:
                        courses[oc][p] = 0
                    courses[oc][p] += 1
                cumul = [] # reset the accumulation
        prev = p
        cumul.append(data)

query = [ 'CCCS', 'CMIS' ] # edit here
        
for c in courses:
    for q in query:
        if q in c:
            print(c, courses[c])


                
