import sys
from docx import Document
import re
import json
import yaml
from collections import OrderedDict

def main():
    file = "Syllabus.docx"
    document = Document(file)
    wordDoc = document.tables
    dates = []
    assignments = []
    count = 0
    for table in wordDoc:
        for row in table.rows:
            for cell in row.cells:
                if "\n" in (cell.text).encode('utf-8') or "/" in (cell.text).encode('utf-8'):
                    #print cell.text.encode('utf-8')
                    if re.findall('\n?\w+\s\d+\/\d+\n?',cell.text.encode('utf-8')):
                        
                        dates.append((cell.text.strip()))
                    else:
                        assignments.append(cell.text.replace("\n"," "))
    duedates = OrderedDict(zip(dates, assignments))
    for x,y in  duedates.items():
        duetext = x,y
        print x.encode('utf-8')+" "+ y.encode('utf-8') 
if __name__ == "__main__":
    main()

