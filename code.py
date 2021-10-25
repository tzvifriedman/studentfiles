## These are all hardcoded paths and names and need to be modified on each run.
## Make sure the Template.docx and the CSV containing {lastname,firstname,grade} 
## exists in the same path as the code.

## Creates a "Graduating Class" folder, and subfolders for each elementray grade.
## This modifies each template and drops it in the K folder.

from docx import Document
import csv
import os
from csv import reader

grad_path_folder = "Graduating Class 2034"
if not os.path.exists(grad_path_folder):
  os.mkdir(grad_path_folder)
  

with open("students.csv") as csvfile:
  csv_reader = reader(csvfile)
  header = next(csv_reader)
  for row in csv_reader:
    fullname = row[1] + row[0]
    fullname_path = row[0] + ", " + row[1]
    # Make folder structure for each student
    os.mkdir(grad_path_folder + "/" + fullname_path)
    for i in ["K",'1','2','3','4']:
      os.mkdir(grad_path_folder + "/" + fullname_path + "/" + i)
    
    ## Make the template
    doc = Document("template.docx")
    replace_dict = {"first":row[1],"last":row[0]}
    for i in replace_dict:
      for p in doc.paragraphs:
        p.text=p.text.replace(i," " + replace_dict[i])
    doc.save("Graduating Class 2034/" + fullname_path + "/" + fullname_path + ".docx")



