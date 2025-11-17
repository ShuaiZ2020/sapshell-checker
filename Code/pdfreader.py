import pypdf
from os import path
# creating a pdf reader object
pdf_path = r'~\OneDrive\Work Yeedo\Project\2025014\CRF\YDMD_2025014_AP306物质平衡_Annotated Blank Pages (unique)_1.0_20250623.pdf'
reader = pypdf.PdfReader(path.expanduser(pdf_path))

# print the number of pages in pdf file
print(len(reader.pages))

# print the text of the first page
[i.split(' ○')[0] for i in reader.pages[122].extract_text().split('\n')[5:]]