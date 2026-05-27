import xlrd
from lib import profileInterpreter
from lib import profileValidation
from zipfile import ZipFile
import os
import io

HOMEDIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) + "/"
CACHEDIR = os.path.join(HOMEDIR, "cache") + "/"

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

def validateRowsForProfile(rows, profilePath):
    errors = getValidationErrorsForProfile(rows, profilePath)

    if errors:
        raise profileValidation.ValidationError(errors)

    return profileInterpreter.Profile(profilePath)

def getValidationErrorsForProfile(rows, profilePath):
    profile = profileInterpreter.Profile(profilePath)
    errors = []

    for rowIndex, row in enumerate(rows):
        if not profile.shouldSkipRow(row):
            errors.extend(profileValidation.validateRow(row, rowIndex, profile.profileValidations))

    return errors

def getValidationErrorText(errors):
    return profileValidation.formatValidationErrors(errors)

def getValidationWarningText(errors):
    warningText = (
        'Validation warnings were found. Processing continued because validations are not being enforced. '
        'MODS created with failed validation may not work in the Workshop.'
    )
    errorText = getValidationErrorText(errors)

    if errorText:
        return warningText + '\n\n' + errorText

    return warningText

def getValidationResultFromExcel(excelFile, sheetName, profilePath, enforceValidations=True):
    rows = convertXlsxToDictList(excelFile, sheetName)
    errors = getValidationErrorsForProfile(rows, profilePath)

    if errors and enforceValidations:
        raise profileValidation.ValidationError(errors)

    return rows, errors

def getValidationStatusFromExcel(excelFile, sheetName, profilePath, enforceValidations=True):
    rows = convertXlsxToDictList(excelFile, sheetName)
    errors = getValidationErrorsForProfile(rows, profilePath)

    if errors and enforceValidations:
        return {
            'can_continue': False,
            'errors': errors,
            'error_text': getValidationErrorText(errors),
        }

    if errors:
        return {
            'can_continue': True,
            'warnings': errors,
            'warning_text': getValidationWarningText(errors),
        }

    return {
        'can_continue': True,
        'errors': [],
        'warnings': [],
    }

def getPreviewResult(excelFile, sheetName, profilePath, globalConditions, enforceValidations=True):
    rows, errors = getValidationResultFromExcel(excelFile, sheetName, profilePath, enforceValidations)
    preview = createPreviewFromRows(rows, profilePath, globalConditions)

    if errors:
        return {
            'preview': preview,
            'warnings': errors,
            'warning_text': getValidationWarningText(errors),
        }

    return {
        'preview': preview,
        'warnings': [],
    }

def createZipFromExcel(excelFile, sheetName, profilePath, globalConditions, enforceValidations=True):
    rows, errors = getValidationResultFromExcel(excelFile, sheetName, profilePath, enforceValidations)

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

def getPreview(excelFile, sheetName, profilePath, globalConditions, enforceValidations=True):
    previewResult = getPreviewResult(excelFile, sheetName, profilePath, globalConditions, enforceValidations)
    return previewResult['preview']

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
