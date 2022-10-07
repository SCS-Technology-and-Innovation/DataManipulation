import pandas as pd
from string import punctuation

# load T&I course list
# load the course information sheet
info = pd.read_csv('courses.csv')
info['Number'] = info['Number'].apply(str)
TI = set(info[['Code', 'Number']].apply(' '.join, axis = 1))

# load the application data
tqrfile = pd.ExcelFile('creditTQR.xlsx')
tqr = tqrfile.parse(tqrfile.sheet_names[0])

stopwords = set('in with experience cutting edge and the field related tools or a b of c such computer industry making as software'.split())

courses = dict()
degree = None
experience = None
pending = False
coursecode = None
keep = False
candidatecode = ''
for index, row in tqr.iterrows():
    entry = str(row[0]).strip()
    if entry == 'or':
        keep = True
    if len(entry) == 8: # new course
        candidatecode = entry
        if not keep:
            m = False
            d = False
            t = False
            i = False
            r = False
            c = False
            kw = set()
        else:
            keep = False
        title = ''
    if len(candidatecode) == 8:
        ct = str(row[1]).strip().replace(',', '')
        if len(title) < len(ct):
            title = ct
        degree = str(row[2]).lower()
        m = m or 'master' in degree
        d = d or 'phd' in degree
        exp = str(row[3]).lower()
        t = t or 'teaching' in exp
        r = r or 'research' in exp
        mi = 'industry' in exp
        i = i or mi
        if mi and candidatecode in TI:
            pos = exp.index('industry')
            content = exp[pos:].replace('-', ' ')
            desc = content.translate(str.maketrans('', '', punctuation))
            terms = set(desc.split())- stopwords
            kw |= terms
        other = str(row[4]).lower()
        c = c or 'certification' in other
        courses[candidatecode] = (title, m, d, t, r, i, c, kw)
        
for cc in courses:
    if cc in TI:
        (n, m, d, t, r, i, c, kw) = courses[cc]
        ifm = 'master' if m else 'n/a'
        ifd = 'doctorate' if d else 'n/a'
        ift = 'teaching' if t else 'n/a'
        ifr = 'research' if r else 'n/a'
        ifi = 'industry' if i else 'n/a'
        ifc = 'cert' if c else 'n/a'
        details = '\n'.join(kw) if len(kw) > 0 else 'n/a'
        print(f'{cc},{n},{ifm},{ifd},{ift},{ifr},{ifi},{ifc},"{details}"')

