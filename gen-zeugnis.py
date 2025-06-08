#!/usr/bin/python3
# -*- coding: utf-8 -*-

# Depends: pip3 install docxtpl
#
# https://docxtpl.readthedocs.io/en/latest/

import argparse
from docxtpl import DocxTemplate
import zipfile
import csv
import os

def detectdelimiter(filename):
    headline = ''
    possible_delimiters = { ',', ';', '|', '\t'}
    with open(filename, encoding='cp1252') as f:
        headline = f.readline()
    delimiter = ';'
    count = 0
    for current_delimiter in possible_delimiters:
        if headline.count(current_delimiter) > count:
            count = headline.count(current_delimiter)
            delimiter = current_delimiter
    return delimiter

def replaceMark(mark):
    readable = mark
    if mark == '1':
        readable = 'sehr gut'
    elif mark == '2':
        readable = 'gut'
    elif mark == '3':
        readable = 'befriedigend'
    elif mark == '4':
        readable = 'ausreichend'
    elif mark == '5':
        readable = 'mangelhaft'
    elif mark == '6':
        readable = 'ungenügend'
    return readable

def render_odt_template(odtFilename, outputfilename, context):
    with zipfile.ZipFile(odtFilename) as inzip, zipfile.ZipFile(outputfilename, "w") as outzip:
        # Iterate the input files
        for inzipinfo in inzip.infolist():
            # Read input file
            with inzip.open(inzipinfo) as infile:
                if inzipinfo.filename == "content.xml":
                    content = infile.read()
                    # Modify the content of the file by replacing a string
                    for placeholder, value in context.items():
                        to_replace = '{{' + placeholder + '}}'
                        content = content.replace(to_replace.encode('utf-8'), value.encode('utf-8'))
                    # Write content
                    outzip.writestr(inzipinfo.filename, content)
                else: # Other file, dont want to modify => just copy it
                    outzip.writestr(inzipinfo.filename, infile.read())
    return

def createReport(outputFolder, filename_path, context):
    filename, file_extension = os.path.splitext(filename_path)
    if (file_extension == '.docx'):
        outputfilename = outputFolder + "/" + context['NN'] + "_" + context['VN'] + ".docx"
        doc = DocxTemplate(filename_path)
        doc.render(context)
        doc.save(outputfilename)
        print('Wrote file: >' + outputfilename + "<")
    if (file_extension == ".odt"):
        outputfilename = outputFolder + "/" + context['NN'] + "_" + context['VN'] + ".odt"
        render_odt_template(filename_path, outputfilename, context)
        print('Wrote file: >' + outputfilename + "<")
    return

def main():
    parser = argparse.ArgumentParser(description='''Generate individual reports from csv-table and docx-template.
    												Use {{<var>}} in docx for substitution. Eg {{VN}} or {{NN}}''',
                                        epilog='Version 1.0. © 2023 by Daniel Ache')
    parser.add_argument('csvfile', 
                            help="Name of the file that holds the list of marks")
    parser.add_argument('docxfile', 
                            help="Name of the template file (This can be .docx or .odt)")
    parser.add_argument('--outputfolder', default='reports',
                    help='foldername for output files')
    parser.add_argument('-mr', '--marksreadable', default=0)

    args = parser.parse_args()

    csvFilename = args.csvfile
    docxFilename = args.docxfile
    outputFolder = args.outputfolder
    substituteMarks = args.marksreadable

    if not os.path.exists(csvFilename):
        print("csv-file: >" + csvFilename + "< does not exits")
        return
    if not os.path.exists(docxFilename):
        print("docx-file: >" + docxFilename + "< does not exits")
        return
        
    if not os.path.exists(outputFolder):
        os.makedirs(outputFolder)

    csvDelimiter = detectdelimiter(csvFilename)
    print("Detected delimiter: '" + csvDelimiter + "'")

    with open(csvFilename, encoding='cp1252') as csvdatei:
        csv_reader_object = csv.reader(csvdatei, delimiter=csvDelimiter)
        iRowCount = 0
        for row in csv_reader_object:
            iRowCount+=1
            if iRowCount == 1:
                header = row
            else:
                context = { }
                i = 0
                for item in header:
                    if (header[i] == 'missed'):
                        context[header[i]] = row[i]
                    elif (header[i] == 'excused'):
                        context[header[i]] = row[i]
                    elif (header[i] == 'nonexcused'):
                        context[header[i]] = row[i]
                    else:
                        if substituteMarks:
                            context[header[i]] = replaceMark(row[i])
                        else:
                            context[header[i]] = row[i]
                    i+=1
                createReport(outputFolder, docxFilename, context)


if __name__ == "__main__":
    main()
