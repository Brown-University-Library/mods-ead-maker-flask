import xlrd
import profileInterpreter
from zipfile import ZipFile
import os
import io
import uuid
from lxml import etree

CACHEDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache") + "/"
HOMEDIR = os.path.dirname(os.path.abspath(__file__)) + "/"

def getSheetNamesFromXlsx(fileContents):
    excel = xlrd.open_workbook(file_contents=fileContents)
    sheetnames = excel.sheet_names()
    return(sheetnames)

def convertXlsxToDictList(fileContents, sheetname):
        book = xlrd.open_workbook(file_contents=fileContents)
        sheet = book.sheet_by_name(sheetname)

        rowarray = []

        for row in range(1, sheet.nrows):
            rowdictionary = {}
            for column in range(sheet.ncols):
                #If the value is a number, turn it into a string.
                newvalue = ''
                if sheet.cell(row,column).ctype > 1:
                    newvalue = str(sheet.cell_value(row,column))
                else:
                    newvalue = sheet.cell_value(row,column)

                #If the column is repeating, serialize the row values.
                if rowdictionary.get(sheet.cell_value(0,column), '') != '':
                    rowdictionary[sheet.cell_value(0,column)] = rowdictionary[sheet.cell_value(0,column)] + '|' + newvalue
                else:
                    rowdictionary[sheet.cell_value(0,column)] = newvalue
            rowarray.append(rowdictionary)

        return rowarray

def cleanStringForFilename(string):
    invalidCharacters = '<>:"/\|?*'

    for character in invalidCharacters:
        string = string.replace(character, '')
        
    return string

def getFilenameFromRow(row, index, filenameColumn):
    if row.get(filenameColumn):
        return cleanStringForFilename(row.get(filenameColumn))
    
    return "default" + str(index)

def createZipFromExcel(excelFile, sheetName, profilePath, globalConditions):
    rows = convertXlsxToDictList(excelFile, sheetName)

    zipBuffer = io.BytesIO()
    zipObj = ZipFile(zipBuffer, 'w')

    for (index, row) in enumerate(rows):
        xmlString, fileBufferValue, filename = createFileFromRow(row, index, profilePath, globalConditions)

        if xmlString is not None:
                zipObj.writestr(filename, fileBufferValue)
            
    zipObj.close()
    
    return zipBuffer.getvalue(), sheetName + '.zip'

def createFileFromRow(row, index, profilePath, globalConditions):
    profile = profileInterpreter.Profile(profilePath, globalConditions=globalConditions)

    xmlString = profile.convertRowToXmlString(row)
    filename = getFilenameFromRow(row, index, profile.profileFilenameColumn) + profile.profileFileExtension

    fileBuffer = io.StringIO()

    if xmlString is not None:
        fileBuffer.write(xmlString)
            
    return xmlString, fileBuffer.getvalue(), filename

def getPreview(excelFile, sheetName, profilePath, globalConditions):
    rows = convertXlsxToDictList(excelFile, sheetName)

    allXmlString = createPreviewFromRows(rows, profilePath, globalConditions)

    return allXmlString

def createPreviewFromRows(rows, profilePath, globalConditions):
    profile = profileInterpreter.Profile(profilePath, globalConditions=globalConditions)

    allXmlString = ""

    for (index, row) in enumerate(rows):
        
        xmlString = profile.convertRowToXmlString(row)
        filename = getFilenameFromRow(row, index, profile.profileFilenameColumn)

        if xmlString:
            allXmlString = allXmlString + "\n\n" + filename + profile.profileFileExtension + "\n\n" + xmlString
            allXmlString = allXmlString.lstrip("\n\n")
    
    return allXmlString