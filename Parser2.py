import sys
from docx import Document
import re
import json
import yaml
from collections import OrderedDict

def main():
    #event = {}
    file = "Syllabus.docx"
    document = Document(file)
    wordDoc = document.tables
    dates = []
    assignments = []
    count = 0
    #datecount
    for table in wordDoc:
        for row in table.rows:
            for cell in row.cells:
                if "\n" in (cell.text).encode('utf-8'):
                    #print cell.text
                    if re.findall('\n?\w+\s\d+\/\d+\n',cell.text.encode('utf-8')):
                        #print str(cell.text)
                        #if "\n" !=  cell.text    
                        
                        dates.append((cell.text.strip()))
                        #print cell.text
                    else:
                        assignments.append(cell.text.strip())
    #print len(assignments)
    #assignments.remove("\n")
    tasks = len(assignments)
    due =  len(dates)-1 
    #print dates
    #print assignments
    #print dates
    #print assignments
    #print assignments
    duedates = OrderedDict(zip(dates, assignments))
   # print duedates
    for x in  duedates.items():
        print x
    
    #print duedates.keys()
    #datez = duedates.items()
    #print datez[0]
    #for x in datez:
    #    print type(x)
    #try:        
    #    while count != tasks:  
    #        print count
    #        task = dates[count] + assignments[count]
    #        print task
    #        count +=1
    #except IndexError:
    #    task = dates[due].extend(assignments[count])
    #    print task            
        #else:
        #    task += assignments[count]
        #    print task

                    #calendar.append((cell.text))
        #print calendar
                    #match = re.search(r'')
                    
                    #print (cell.text).encode('utf-8')
                    #print json.dumps(cell.text.encode('utf-8'), separators=(" ",":"), indent=4, sort_keys=True)
                    #if "Tue" in (cell.text).encode('utf-8'):

                        #calendar =  cell.text.encode('utf-8')
                        #print cell.text.encode('utf-8')
                    #p = re.compile(r"\n\w{3}\s\d+\/\d+\n\n.*[^/]+\n")
                    #print re.findall(r'\n\w{3}\s\d+\/\d+\n.*[^/]+', cell.text.encode('utf-8'))
                    #print p.findall(cell.text.encode('utf-8'))
                        #date =  str(cell.text.encode('utf-8')).strip("\n")+":"
                        #print date
                        #calendar.append(date)
                    #else:
                    #    hw = str(cell.text.encode('utf-8')).strip("\n")
                    #    print hw.replace("\n", ",")
                    #    print 1
                    #    assignments.append(hw)
                        #event['dateTime'] = date
                    #print json.dumps(date, separators=('\n',':'), indent=4, sort_keys=True)
                    #else:
                        #assignment =  str(cell.text.encode('utf-8')).replace("\n"," ")
                        #event['description'] = assignment
    #print event
    #print calendar
    #print assignments
    #print dict(zip (calendar,assignments))





if __name__ == "__main__":
    main()

