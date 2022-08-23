import os
os.chdir('C:/Users/sschae/Downloads')
from PyPDF2 import PdfReader
reader = PdfReader("demo1.pdf")
page = reader.pages[0]
content = page.extract_text()
print(content)
