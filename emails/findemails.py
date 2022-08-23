from PyPDF2 import PdfReader
import os
import glob

emails = set()
os.chdir('C:/Users/sschae/Downloads/emails')
for inputfile in glob.glob('*.pdf'):
    reader = PdfReader(inputfile)
    pagecount = len(reader.pages)
    for pagenumber in range(pagecount):
        page = reader.pages[pagenumber]
        content = page.extract_text()
        for word in content.split():
            if '@' in word:
                emails.add(word)

for address in emails:
    print(address)
