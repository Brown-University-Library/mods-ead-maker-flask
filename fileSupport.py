import xlrd
import profileInterpreter
from zipfile import ZipFile
import os
import io
import uuid
from lxml import etree

CACHEDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache") + "/"
HOMEDIR = os.path.dirname(os.path.abspath(__file__)) + "/"

def getSheetNames(fileContents):
    excel = xlrd.open_workbook(file_contents=fileContents)
    sheetnames = excel.sheet_names()
    return(sheetnames)

def convertXlsxFileToDict(fileContents, sheetname):
        book = xlrd.open_workbook(file_contents=fileContents)
        sheet = book.sheet_by_name(sheetname)

        rowarray = []

        for row in range(1, sheet.nrows):
            rowdictionary = {}
            for column in range(sheet.ncols):
                #If the value is a number, turn it into a string.
                newvalue = ''
                if sheet.cell(row,column).ctype > 1:
                    newvalue = str(sheet.cell_value(row,column)).replace(';','|')
                else:
                    newvalue = sheet.cell_value(row,column).replace(';','|')

                #If the column is repeating, serialize the row values.
                if rowdictionary.get(sheet.cell_value(0,column), '') != '':
                    rowdictionary[sheet.cell_value(0,column)] = rowdictionary[sheet.cell_value(0,column)] + '|' + newvalue
                else:
                    rowdictionary[sheet.cell_value(0,column)] = newvalue
            rowarray.append(rowdictionary)

        return rowarray

def getFilenameFromRow(row, index):
    if row.get("identifierBDR"):
        return row.get("identifierBDR").replace(":","")
    else:
        return "default" + str(index)

def createZipFromExcel(excelFile, sheetName, profilePath, globalConditions):
    rows = convertXlsxFileToDict(excelFile, sheetName)
    profile = profileInterpreter.Profile(profilePath)
    profile.globalConditionsSet = globalConditions

    zipBuffer = io.BytesIO()
    zipObj = ZipFile(zipBuffer, 'w')

    for (index, row) in enumerate(rows):
        xmlString = profile.convertRowToXmlString(row)
        filename = getFilenameFromRow(row, index)

        fileBuffer = io.StringIO()

        if xmlString != None:
                fileBuffer.write(xmlString)
                zipObj.writestr(filename + ".mods", fileBuffer.getvalue())
            
    zipObj.close()
    
    return zipBuffer.getvalue(), sheetName + '.zip'

def getPreview(excelFile, sheetName, profilePath, globalConditons):
    rows = convertXlsxFileToDict(excelFile, sheetName)
    profile = profileInterpreter.Profile(profilePath)
    profile.globalConditionsSet = globalConditons
    allXmlString = ""

    for (index, row) in enumerate(rows):
        
        xmlString = profile.convertRowToXmlString(row)
        filename = getFilenameFromRow(row, index)

        if xmlString:
            allXmlString = allXmlString + "\n\n" + filename + ".mods" + "\n\n" + xmlString
            allXmlString = allXmlString.lstrip("\n\n")

    return allXmlString