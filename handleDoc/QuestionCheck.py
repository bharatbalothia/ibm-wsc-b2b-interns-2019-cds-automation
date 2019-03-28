from docx.api import Document
import re


def questcheck(filename):
    document = Document(filename)
    for table in document.tables:
        if "Trading Partner’s Configuration" in table.rows[0].cells[0].text:
            info ={}
            for row in table.rows:
                cells = row.cells
                if 'Mandatory' in cells[0].text and cells[1].text == "":
                    yield cells[0].text




'''
keys = None
for i, row in enumerate(table.rows):
    text = (cell.text for cell in row.cells)

    # Establish the mapping based on the first row
    # headers; these will become the keys of our dictionary
    if i == 0:
        keys = tuple(text)
        continue

    # Construct a dictionary for this row, mapping
    # keys to values for this row
    row_data = dict(zip(keys, text))
    data.append(row_data)
'''