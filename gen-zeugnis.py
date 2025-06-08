#!/usr/bin/python3
# -*- coding: utf-8 -*-

# Depends: pip3 install docxtpl pandas openpyxl
#
# https://docxtpl.readthedocs.io/en/latest/

import argparse
from docxtpl import DocxTemplate
import zipfile       # needed for patching odt
import pandas as pd  # needed for reading xlsx and ods
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

def readMarksFile(filename_path, substituteMarks):
    marksTable = [ ]
    filename, file_extension = os.path.splitext(filename_path)
    if ((file_extension == '.ods') or (file_extension == '.xlsx')):
        # read by default 1st sheet of an excel file
        dataframe1 = pd.read_excel(filename_path)
        rowcount = dataframe1.shape[0]
        for i in range(0,rowcount):
            context = { }
            for header in dataframe1.columns:
                if (header == 'missed'):
                    context[header] = str(dataframe1.at[i, header])
                elif (header == 'excused'):
                    context[header] = str(dataframe1.at[i, header])
                elif (header == 'nonexcused'):
                    context[header] = str(dataframe1.at[i, header])
                else:
                    if substituteMarks:
                        context[header] = replaceMark(str(dataframe1.at[i, header]))
                    else:
                        context[header] = str(dataframe1.at[i, header])
            marksTable.append(context)
    if (file_extension == '.csv'):
        csvDelimiter = detectdelimiter(filename_path)
        print("Detected delimiter: '" + csvDelimiter + "'")
        
        with open(filename_path, encoding='cp1252') as csvdatei:
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
                    marksTable.append(context)

    return marksTable

def main():
    parser = argparse.ArgumentParser(
        description='Generate individual reports from a table file (CSV/XLSX/ODS) and a template file (DOCX/ODT). '
                    'Use {{<var>}} in your template for substitution, e.g. {{VN}} or {{NN}}.',
        epilog='Version 1.0. © 2023-2025 by Daniel Ache'
    )
    parser.add_argument('datafile', 
        help="Path to the file containing the list of marks (supported: .csv, .xlsx, .ods)")
    parser.add_argument('templatefile', 
        help="Path to the template file (supported: .docx, .odt)")
    parser.add_argument('--outputfolder', default='reports',
        help='Output folder for generated files (default: reports)')
    parser.add_argument('-mr', '--marksreadable', type=int, default=0,
        help='Convert numeric marks to text representation (1 = enabled, 0 = disabled, default: 0)')

    args = parser.parse_args()

    datafile = args.datafile
    templatefile = args.templatefile
    outputfolder = args.outputfolder
    marksreadable = args.marksreadable

    if not os.path.exists(datafile):
        print("Data file: >" + datafile + "< does not exist")
        return
    if not os.path.exists(templatefile):
        print("Template file: >" + templatefile + "< does not exist")
        return
        
    if not os.path.exists(outputfolder):
        os.makedirs(outputfolder)

    marksTable = readMarksFile(datafile, marksreadable)
    for context in marksTable:
        createReport(outputfolder, templatefile, context)

if __name__ == "__main__":
    main()
