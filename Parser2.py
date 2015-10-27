import sys
from docx import Document

def main():

    file = "Syllabus.docx"
    document = Document(file)
    wordDoc = document.tables
    calendar = []
    for table in wordDoc:
        for row in table.rows:
            for cell in row.cells:
                if "\n" in (cell.text).encode('utf-8'):
                    print cell.text.encode('utf-8')
                






if __name__ == "__main__":
    main()

