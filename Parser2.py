import sys
from docx import Document
import re
import json
import yaml
from collections import OrderedDict

def syllaparsed(parsed):
    #pass
    #dates = []
    #assignments = []
    #print parsed
    #if re.findall('\n?\w+\s\d+\/\d+\n?',parsed):
        #print parsed
        #dates.append(parsed.strip())
    #else:
        #assignments.append(parsed.replace("\n"," "))
    
    print parsed.strip().replace("\n"," ")
    #print dates,assignments
    #print assignments
    #duedates = OrderedDict(zip(dates,assignments))
    #print duedates
    #for h1,h2 in duedates.items():
    #    duetext = h1,h2
    #    print h1.encode('utf-8')+" "+ h2.encode('utf-8')
def main():
    syllabus = "dbsyllabus.docx"
    document = Document(syllabus)
    wordDoc = document.tables
    #dates = []
    #assignments = []
    count = 0
    for table in wordDoc:
        for row in table.rows:
                for cell in row.cells:
                     if len(row.cells) == 3:
                         if "\n" in (cell.text).encode('utf-8') or "/" in (cell.text).encode('utf-8'):
                             if re.findall('\n?\w+\s\d+\/\d+\n?',cell.text.encode('utf-8')):
                                 dates.append((cell.text.strip().encode('utf-8')))
                             else:
                                 assignments.append(((cell.text).encode('utf-8')).replace("\n"," "))
                     elif len(row.cells) == 4:
                         if "/" in (cell.text).encode('utf-8'):
                             m = re.search('\d+\/\d+',cell.text.encode('utf-8'))
                             dates.append(m.group(0))
                             s = re.split('\d+\/\d+',cell.text.encode('utf-8'))
                             for items in s:
                                 if "" != items:
                                     (assignments.append(items.strip().replace("\n"," ").replace("\xe2\x80\x93","")))
 
    duedates = OrderedDict(zip(dates, assignments))
    print (duedates)
    #for table in wordDoc:
    #    for row in table.rows:
    #        for cell in row.cells:
    #            if "\n" in (cell.text).encode('utf-8') or "/" in (cell.text).encode('utf-8'):
    #                    syllaparsed(cell.text.encode('utf-8'))
                        #if re.findall('\n?\w+\s\d+\/\d+\n?',cell.text.encode('utf-8')):
                            #print cell.text.encode('utf-8')
                        
                            #dates.append((cell.text.strip()))
                        #else:
                            #assignments.append(cell.text.replace("\n"," "))
    #print dates
    #print assignments
    #for assign in assignments:
        #print assign
    #duedates = OrderedDict(zip(dates, assignments))
    #print duedates
    
    #for x,y in  duedates.items():
    #    duetext = x,y
        #print x.encode('utf-8')+" "+ y.encode('utf-8') 
if __name__ == "__main__":
    main()

