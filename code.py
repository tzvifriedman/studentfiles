## These are all hardcoded paths and names and need to be modified on each run.
## Make sure the Template.docx and the CSV containing {lastname,firstname,grade} 
## exists in the same path as the code.

## Creates a "Graduating Class" folder, and subfolders for each elementray grade.
## This modifies each template and drops it in the K folder.

from docx import Document
import csv
import os
from csv import reader

if not os.path.exists("Graduating Class 2034"):
  os.mkdir("Graduating Class 2034")
  for i in ['K','1','2','3','4']:
    os.mkdir("Graduating Class 2034/" + i)

with open("KStudents2021.csv") as csvfile:
  csv_reader = reader(csvfile)
  header = next(csv_reader)
  for row in csv_reader:
    fullname = row[1] + row[0]
    doc = Document("Template.docx")
    replace_dict = {"first":row[1],"last":row[0]}
    for i in replace_dict:
      for p in doc.paragraphs:
        p.text=p.text.replace(i," " + replace_dict[i])
    doc.save("Graduating Class 2034/K/" + fullname + ".docx")



