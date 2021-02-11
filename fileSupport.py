import xlrd
import MODSMaker2
from zipfile import ZipFile
import os
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

def createZipFromExcel(excelFile, sheetName, profilePath, globalConditons):
    id = str(uuid.uuid4())
    rows = convertXlsxFileToDict(excelFile, sheetName)
    profile = MODSMaker2.Profile(profilePath)
    os.mkdir(CACHEDIR + id)
    zipObj = ZipFile(CACHEDIR + id + "/" + sheetName + '.zip', 'w')

    for (index, row) in enumerate(rows):
        
        xmlString = profile.convertRowToXmlString(row, {})
        filename = getFilenameFromRow(row, index)

        if xmlString != None:
            with open(CACHEDIR + id + "/" + filename + ".mods", 'w+') as f:
                f.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
                f.write(xmlString)

            zipObj.write(CACHEDIR + id + "/" + filename + ".mods", filename + ".mods")

    zipObj.close()
    with open(CACHEDIR + id + "/" + sheetName + '.zip', mode='rb') as zipdata:
        return zipdata.read(), sheetName + '.zip'

def getPreview(excelFile, sheetName, profilePath, globalConditons):
    rows = convertXlsxFileToDict(excelFile, sheetName)
    profile = MODSMaker2.Profile(profilePath)
    allXmlString = ""

    for (index, row) in enumerate(rows):
        
        xmlString = profile.convertRowToXmlString(row, {})
        filename = getFilenameFromRow(row, index)

        if xmlString:
            allXmlString = allXmlString + "\n\n" + filename + ".mods" + "\n\n" + xmlString
            allXmlString = allXmlString.lstrip("\n\n")

    return allXmlString