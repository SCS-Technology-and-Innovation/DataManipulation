from PyPDF2 import PdfReader
import os
import glob

os.chdir('C:/Users/sschae/Downloads/pdfdemo')
for inputfile in glob.glob('*.pdf'):
    print('Processing', inputfile)
    reader = PdfReader(inputfile)
    page = reader.pages[0]
    content = page.extract_text()
    print('It contains this\n', content)
