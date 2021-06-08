#!/usr/bin/python
# -*- coding: UTF-8 -*-
 
import sys
import argparse
import docx
import json

def createParser ():
    parser = argparse.ArgumentParser()
    parser.add_argument ('-n', '--inputfile', default='example.docx')
    parser.add_argument ('-n', '--tablejson')
    parser.add_argument ('-n', '--otputfile')
    return parser

if __name__ == '__main__':
    parser = createParser()
    arguments = parser.parse_args(sys.argv[1:])

    doc = docx.Document(arguments.inputfile)
    table = json.loads(arguments.tablejson)


    doc.add_page_break()

    
    records = (
        (3, '101', 'Spam'),
        (7, '422', 'Eggs'),
        (4, '631', 'Spam, spam, eggs, and spam')
    )

    table = doc.add_table(rows=1, cols=3)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Qty'
    hdr_cells[1].text = 'Id'
    hdr_cells[2].text = 'Desc'
    for qty, id, desc in records:
        row_cells = table.add_row().cells
        row_cells[0].text = str(qty)
        row_cells[1].text = id
        row_cells[2].text = desc



    doc.save('demo.docx')