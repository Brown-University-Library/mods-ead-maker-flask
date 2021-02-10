import os, csv, xlsxwriter
import lxml
from lxml import etree
from openpyxl import load_workbook
import openpyxl
from lxml.builder import ElementMaker
import string, codecs
import chardet
import datetime
import re
import xlrd
import sys
from copy import copy
from zipfile import ZipFile
import yaml

def getTermsOfAddressPrependAndAppendStripped(name):
    appendTermOfAddress = ""
    prependTermOfAddress = ""

    for textIndex, text in enumerate(name.split(',')):
        if textIndex == 0:
            appendTermsOfAddress = re.findall("(\{\{.*\}\})", text)
            if len(appendTermsOfAddress) > 0:
                appendTermOfAddress = appendTermsOfAddress[0].replace("{{","").replace("}}","")
        
        if textIndex == 1:
            prependTermsOfAddress = re.findall("(\{\{.*\}\})", text)
            if len(prependTermsOfAddress) > 0:
                prependTermOfAddress = prependTermsOfAddress[0].replace("{{","").replace("}}","")
    
    return prependTermOfAddress, appendTermOfAddress

def getValueUri(name):
    uris = re.findall("(?P<url>https?://[^\s]+)", name)

    #If there's a URI
    if len(uris) > 0:
        return normalizeString(uris[0])
    else: 
        return ""

def getNameDateRoleFromEntry(entry):

    name = ""
    date = ""
    role = ""

    for textIndex, text in enumerate(entry.split(',')):
        normalizedText = normalizeString(text)

        if normalizedText == '':
            continue
        
        if textIndex == 0:
            name = name + normalizedText + ", "
        elif hasYear(normalizedText) == True:
            date = date + normalizedText
            date = date.lstrip(',').rstrip(',')
        elif isAllLower(normalizedText) == True:
            role = text
        elif hasLetters(normalizedText) != None:
            name = name + normalizedText + " "

    return normalizeString(name).rstrip(",").lstrip(", "), normalizeString(date), normalizeString(role)


def getMetadataFromEntry(entry):
    valueUri = getValueUri(entry)
    entry = entry.replace(valueUri, "")
    value = normalizeString(entry)

    prependTermOfAddress, appendTermOfAddress = getTermsOfAddressPrependAndAppendStripped(entry)
    entry = entry.replace("{{" + prependTermOfAddress + "}}", "")
    entry = entry.replace("{{" + appendTermOfAddress + "}}", "")

    name, date, role = getNameDateRoleFromEntry(entry)

    return {"entry.value": value, "entry.valueURI":valueUri, "entry.name": name, "entry.date": date, "entry.role": role, "entry.prependTermOfAddress": prependTermOfAddress, "entry.appendTermOfAddress":appendTermOfAddress}


def normalizeString(string):
    if string != None:
        string = string.replace('\n', ' ').replace('\r', ' ')
        string = string.replace('<title>', '').replace('</title>', '')
        string = string.replace('<geogname>', '- ').replace('</geogname>', '')
        return(' '.join(str(string).split()))
    else:
        return string

def hasNumbers(s):
    return any(i.isdigit() for i in s)

def hasYear(s):
    numbercount = 0
    for i in s:
        if i.isdigit() == True:
            numbercount = numbercount + 1
    if numbercount > 3:
        return True
    else:
        return False

def hasLetters(s):
    return re.search('[a-zA-Z]', s)

def isAllLower(s):
    nonlowercase = 0
    for i in s.replace(' ', ''):
        if i.islower() == False:
            nonlowercase = nonlowercase + 1
            break
    if nonlowercase > 0:
        return False
    else:
        return True

def shouldSkipRow(row, modsKeySkips):
    for modsKeySkip in modsKeySkips:
        if row.get(modsKeySkip) != None:
            return True

key = yaml.safe_load(open("MODSkey.yaml"))

keySkips = key.get("skipif", [])
keyFields = key.get("fields", [])
keyAuthorities = key.get("authorities")
keySorts = key.get("sort", [])
keyNameSpace = key.get("elementnamespace", [])

class Element:
    def __init__(self, **entries): 
        self.__dict__.update(entries)

def clearEmptyElementsFromEtree(parentElement):
    # start cleanup
    # remove any element tails
    for element in parentElement.iter():
        element.tail = None

    # remove any line breaks or tabs in element text
        if element.text:
            if '\n' in element.text:
                element.text = element.text.replace('\n', '')
            if '\t' in element.text:
                element.text = element.text.replace('\t', '')

    # remove any remaining whitespace
    parser = etree.XMLParser(remove_blank_text=True, remove_comments=True, recover=True)
    treestring = etree.tostring(parentElement)
    clean = etree.XML(treestring, parser)

    # remove recursively empty nodes
    # found here: https://stackoverflow.com/questions/12694091/python-lxml-how-to-remove-empty-repeated-tags
    def recursively_empty(e):
        if e.text:
            return False
        return all((recursively_empty(c) for c in e.iterchildren()))

    context = etree.iterwalk(clean)
    for action, elem in context:
        parent = elem.getparent()
        if recursively_empty(elem) and parent is not None:
            parent.remove(elem)

    # remove nodes with blank attribute
    for element in clean.xpath(".//*[@*='']"):
        element.getparent().remove(element)

    # remove nodes with attribute "null"
    for element in clean.xpath(".//*[@*='null']"):
        element.getparent().remove(element)

    return clean

def createParentElement(key, keyNsMap):
    keyQName = key.get("attrqname", {})
    keyAttrQname = etree.QName(keyQName.get("uri",""),keyQName.get("tag",""))

    # keyNsMap = key.get("nsmap", {})

    keyElementNameSpace = key.get("elementnamespace", "")
    keyParentTag = key.get("parenttag", "")

    return lxml.etree.Element(keyElementNameSpace + keyParentTag, {keyAttrQname: "http://www.loc.gov/mods/v3 http://www.loc.gov/mods/v3/mods-3-7.xsd"}, nsmap=keyNsMap)

def createSubElement(parentElement, elementName, elementNameSpace, elementAttrs, elementText):
    subElement = etree.SubElement(parentElement, elementNameSpace + elementName, elementAttrs)
    subElement.text = normalizeString(elementText)

    return subElement

### Column methods

def processValueColumn(column, row):
    columnHeader = column.get("header")
    text = row.get(columnHeader, "")
    return normalizeString(text)

def processNumColumn(column, row):
    columnHeader = column.get("header")
    text = row.get(columnHeader, "").replace(".0", "")
    return normalizeString(text)

def processLowerColumn(column, row):
    columnHeader = column.get("header")
    text = row.get(columnHeader, "").lower()
    return normalizeString(text)

def processColumnTextValue(column, row):
    columnMethod = column.get("method")
    columnHeader = column.get("header")

    if columnMethod  == "value":
        return processValueColumn(column, row)
    if columnMethod  == "num":
        return processNumColumn(column, row)
    if columnMethod  == "lower":
        return processLowerColumn(column, row)

    return row.get(columnHeader,"")

### Conditional attributes

def processConditionalAttrs(conditionalAttr, row, element):
    key = conditionalAttr.get("key","")

    textKeyField = conditionalAttr.get("text","")
    text = processTextKeyField(textKeyField, row)

    if text:
        element.set(key, text)

#### Text key fields

def processTextKeyFieldValues(textKeyFieldValues, row):
    text = ""
    for value in textKeyFieldValues:
        valueType = value.get("type","")
        valueHeader = value.get("header",None)
        valueText = value.get("text","")

        if valueType == "value":
            text = text + valueText

        if valueType == "col":
            text = text + processColumnTextValue(value, row)
    
    return text

def performTextAction(textAction, text):
    if textAction.get("action") == "leftstriprightstrip":
        lstripRstripText = textAction.get("leftstriprightstriptext", "")
        return text.lstrip(lstripRstripText).rstrip(lstripRstripText)

def processTextKeyField(textKeyFields, row):
    text = ""
    for textKeyField in textKeyFields:
        textKeyFieldType = textKeyField.get("type","")
        textKeyFieldColumn = textKeyField.get("col", None)
        textKeyFieldValues = textKeyField.get("values","")

        if  textKeyFieldType == "ifpresent":
            if row.get(textKeyFieldColumn):
                newText = processTextKeyFieldValues(textKeyFieldValues, row)
                text = text + newText
        if  textKeyFieldType == "ifnotpresent":
            if row.get(textKeyFieldColumn) == None:
                newText = processTextKeyFieldValues(textKeyFieldValues, row)
                text = text + newText
        if  textKeyFieldType == "value":
            newText = processTextKeyFieldValues(textKeyFieldValues, row)
            text = text + newText
        if  textKeyFieldType == "removetext":
            newText = processTextKeyFieldValues(textKeyFieldValues, row)
            replaceStrings = textKeyField.get("removetext",[])
            for replaceString in replaceStrings:
                newText = newText.replace(replaceString, "")
            text = text + newText
        if  textKeyFieldType == "action":
            text = performTextAction(textKeyField, text)

    return text

def shouldCreateElementBasedOnCondition(condition, row):
    conditionType = condition.get("type", "")

    if conditionType == "startswith":
        column = condition.get("col","")
        rowValue = row.get(column, "")
        startsWithText = condition.get("text", "")
        if rowValue.startswith(startsWithText):
            return True

    if conditionType == "has":
        column = condition.get("col","")
        rowValue = row.get(column, "")
        hasText = condition.get("text", "")
        if hasText in rowValue:
            return True

    return False

def processElementTypeField(keyField, elementNameSpace, row, parentElement):
    elementName = keyField.get("name", "")
    elementAttrs = keyField.get("attrs", {})
    element = createSubElement(parentElement, elementName, elementNameSpace, elementAttrs, "")
    columns = keyField.get("cols", {})
    conditions = keyField.get("conditions", [])

    for condition in conditions:
        shouldCreateElement = shouldCreateElementBasedOnCondition(condition, row)
        if shouldCreateElement == False:
            return
    
    conditionalAttrs = keyField.get("conditionalattrs", [])
    for conditionalAttr in conditionalAttrs:
        processConditionalAttrs(conditionalAttr, row, element)

    childrenKeyFields = keyField.get("children", {})
    for childKeyField in childrenKeyFields:
        processElementTypeField(childKeyField, elementNameSpace, row, element)
    
    textKeyField = keyField.get("text", [])
    text = processTextKeyField(textKeyField, row)
    if text:
        element.text = text

    return element

def handleRepeatingTypeEntry(parentElement, rowString, keyElement, authority, repeatingDefaults, originalRow):
    entries = rowString.split("|")
    authorityAdditions = {}
    authorityAdditions["entry.authority"] = authority.get("authority", "")
    authorityAdditions["entry.authorityURI"] = authority.get("authorityURI", "")
    elementsCreated = []

    for entry in entries:
        print(entry)
        entryAdditions = getMetadataFromEntry(entry)
        print(entryAdditions)
        if areAllDictValuesEmpty(entryAdditions) == False:
            entryAdditions.update(authorityAdditions)
            entryAdditions.update(repeatingDefaults)
            print(originalRow)
            entryAdditions.update(originalRow)
        element = processElementTypeField(keyElement, keyNameSpace, entryAdditions, parentElement)
        elementsCreated.append(element)
    
    return elementsCreated

def areAllDictValuesEmpty(dict={}):
    keys = dict.keys()

    for key in keys:
        if dict[key] == "":
            continue
        else:
            return False

    return True

def processRepeatingTypeField(keyField, keyAuthorities, elementNameSpace, row, parentElement):
    repeatingMethod = keyField.get("method")
    colPrefixes = keyField.get("colprefix", [])
    colHeaders = keyField.get("cols", [])
    repeatingElement = keyField.get("element", {})
    repeatingDefaults = keyField.get("defaults", {})

    if repeatingMethod  == "name" or repeatingMethod == "value":
        for colPrefix in colPrefixes:
            for keyAuthority in keyAuthorities:
                colHeader = colPrefix + keyAuthority.get("suffix", "")
                rowString = row.get(colHeader, "")
                handleRepeatingTypeEntry(parentElement, rowString, repeatingElement, keyAuthority, repeatingDefaults, row)
        for colHeader in colHeaders:
            rowString = row.get(colHeader, "")
            handleRepeatingTypeEntry(parentElement, rowString, repeatingElement, {}, repeatingDefaults, row)

def processSort(parentElement, sort, keyParentTag, keyElementNameSpace):
    nameSpace = {keyParentTag: keyElementNameSpace}
    elementXpath = sort.get("elementxpath", "")
    sortByXpath = sort.get("sortbyxpath", "")

    allMatchingElements = parentElement.xpath(elementXpath, namespaces=parentElement.nsmap)
    firstElementIndex = 0
    
    if len(allMatchingElements) > 0:
        firstElementIndex = parentElement.getchildren().index(allMatchingElements[0])

    allMatchingElements = sorted(allMatchingElements, key=lambda ch: ch.xpath(sortByXpath, namespaces={keyParentTag: keyElementNameSpace.replace("{","").replace("}","")}))

    for element in allMatchingElements:
        parentElement.remove(element)

    for (index, element) in enumerate(allMatchingElements):
        parentElement.insert(firstElementIndex + index, element)

def convertExcelRowToEtree(row, globalConditions):
    keyElementNameSpace = key.get("elementnamespace", "")
    keyParentTag = key.get("parenttag", "")

    if shouldSkipRow(row, keySkips):
        return
    
    parentElement = createParentElement(key, key.get("nsmap", {}))

    for keyField in keyFields:
        print(keyField)
        keyFieldType = keyField.get('type', "")

        if keyFieldType == 'element':
            element = processElementTypeField(keyField, keyNameSpace, row, parentElement)

        if keyFieldType == "repeating":
            elements = processRepeatingTypeField(keyField, keyAuthorities, keyNameSpace, row, parentElement) 

    cleandUpFile = clearEmptyElementsFromEtree(parentElement)

    for keySort in keySorts:
        processSort(cleandUpFile, keySort, keyParentTag, keyElementNameSpace)
    
    print(cleandUpFile)
    print(etree.tostring(cleandUpFile, pretty_print=True).decode("utf-8"))
####

row = {"subjectCorpNAF":"Zacks Corp, 1991-2002|B Corp, 1992-1993 http://google.com|A Corp|","subjectNamesLocal":"Person, Other, 1992-|Berson, Other, 1992-, author","nameCorpCreatorNAF":"B Corp, 1992-1993 http://google.com|A Corp","namePersonOtherLocal":"Person, Other, 1992-|Berson, Other, 1992-, author","rightsStatementURI":"www.google.com","rightsStatementText":"In Copyright","physicalLocationNAF":"Brown University. Library http://id.loc.gov/authorities/names/n81029638","shelfLocator3":"Paper","shelfLocator3ID":"100", "shelfLocator2":"Box","shelfLocator2ID":"Hello","identifierBDR":"bdr411","callNumber":"callNo", "dateText":"10-04-2014","dateStart":"10-00-2015", "dateEnd":"10-21-2016", "rightsStatementText":"In Copyright","subjectTopicsTemporalLocal":"Mock Temporal|A Temporal", "subjectNamesLC":"Name, One, manager, 1992-1993|Name, Two, director, 1993-1994", "genreLocal": "Local Genre 1|Local Genre 2","genreFAST": "genre1 http://genre.gov|genre2", "subjectTopicsLocal": "A Local Topic|C Local Topic", "subjectTopicsFreedomNow":"B FN Topic|FN2", "subjectTitleLC":"yes https://google.com| no https://yahoo.net", "itemTitle":"whatever", "itemTitlePartNumber":"1", "itemTitlePartName":"Hello", "typeOfResource":"water", "typeOfResourceCollection":"sad", "namePersonCreatorFAST":"Fast {{the great}}, Person, job having, 1991-1992", "namePersonCreatorNAF": "Nadler, Mad, 1989-2002, little helper https://google.com| Guy {{III}}, Happy, 1928-, useful man https://facebook.com "}
globalConditions = {"includeDefaults": True}

# convertExcelRowToEtree(row, globalConditions)

# columnHeaderOverride = {"identifierBDR":"bdr:123456"}


def getHeadersFromTextFields(textFields):
    textHeaders = []
    for textField in textFields:
        textFieldValues = textField.get("values", [])
        for textFieldValue in textFieldValues:
            if textFieldValue.get("type") == "col":
                header = textFieldValue.get("header", "")
                textHeaders.append(header)
    
    return textHeaders

def getConditionalAttrsTextHeadersAndConditions(conditionalAttr, name):

    keyElementNameSpace = key.get("elementnamespace", "")
    keyParentTag = key.get("parenttag", "")

    textHeaders = []
    conditionalAttrConditions = []

    attrKey = conditionalAttr.get("key", "")
    attrTextFields = conditionalAttr.get("text", [])
    textHeaders = getHeadersFromTextFields(attrTextFields)

    for attrTextField in attrTextFields:
        attrTextType = attrTextField.get("type","")
        if attrTextType == "ifpresent":
            string = "If text is entered in the " + attrTextField.get("col","") + " column, the following attribute will be added to the <" + keyParentTag + ":" + name + "> element:"
            value = processTextKeyFieldValues(attrTextField.get("values",[]), convertArrayToDictWithMatchingKeyValues(textHeaders))
            attrString = attrKey + "='" + value + "'"
            conditionalAttrCondition = {"explanation":string, "attribute":attrString}
            conditionalAttrConditions.append(conditionalAttrCondition)
            textHeaders.append(attrTextField.get("col",""))
        if attrTextType == "ifnotpresent":
            string = "If text is not entered in the " + attrTextField.get("col","") + " column, the following attribute will be added to the <" + keyParentTag + ":" + name + "> element:"
            value = processTextKeyFieldValues(attrTextField.get("values",[]), convertArrayToDictWithMatchingKeyValues(textHeaders))
            attrString = attrKey + "='" + value + "'"
            conditionalAttrCondition = {"explanation":string, "attribute":attrString}
            conditionalAttrConditions.append(conditionalAttrCondition)
            textHeaders.append(attrTextField.get("col",""))

    return textHeaders, conditionalAttrConditions

def convertArrayToDictWithMatchingKeyValues(array):
    dictionary = {}

    for arrayItem in array:
        dictionary[arrayItem] = arrayItem
    
    return dictionary

def removeDuplicatesFromArray(array):
    array = list(dict.fromkeys(array))
    return array

def removeItemsWithPeriodFromList(array):
    returnArray = []

    for item in array:
        if "." in item:
            continue
        returnArray.append(item)
    
    return returnArray

replaceTextForExamples = ' xmlns:mods="http://www.loc.gov/mods/v3" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xlink="http://www.w3.org/1999/xlink"'

def getConditionalAttrsTextHeadersElementFromField(keyField):

    keyElementNameSpace = key.get("elementnamespace", "")
    keyParentTag = key.get("parenttag", "")
    keySampleValues = key.get("samplevalues",{})
    name = keyField.get("name", [])
    textFields = keyField.get("text", [])
    parentElement = createParentElement(key, key.get("nsmap", {}))

    textHeaders = []
    conditionalAttrConditions = []

    textHeaders = getHeadersFromTextFields(textFields)
    
    conditionalAttrs = keyField.get("conditionalattrs", [])
    for conditionalAttr in conditionalAttrs:
        conditionalAttrTextHeaders, conditionalAttrCondition = getConditionalAttrsTextHeadersAndConditions(conditionalAttr, name)
        textHeaders.extend(conditionalAttrTextHeaders)
        conditionalAttrConditions.extend(conditionalAttrCondition)
    
    for child in keyField.get("children",[]):
        childTextHeaders, childConditionalAttrsHeaders, elementString = getConditionalAttrsTextHeadersElementFromField(child)
        textHeaders.extend(childTextHeaders)
        conditionalAttrConditions.extend(childConditionalAttrsHeaders)

    textHeaderRow = convertArrayToDictWithMatchingKeyValues(textHeaders)
    textHeaderRow.update(keySampleValues)
    element = processElementTypeField(keyField, keyNameSpace, textHeaderRow, parentElement)
    cleanedElement = clearEmptyElementsFromEtree(element)
    elementString = etree.tostring(cleanedElement, pretty_print=True, encoding="UTF-8").decode("utf-8")
    elementString = elementString.replace(replaceTextForExamples, "")
    
    return removeDuplicatesFromArray(textHeaders), conditionalAttrConditions, elementString

def getConditionalAttrsTextHeadersElementFromRepeatingField(keyField):
    parentElement = createParentElement(key, key.get("nsmap", {}))
    repeatingElement = keyField.get("element", {})
    repeatingDefaults = keyField.get("defaults", {})
    colPrefixes = keyField.get("colprefix", [])
    colHeaders = keyField.get("cols", [])
    repeatingFieldMethod = keyField.get("method", "")

    columnHeaders = []
    elementsCreated = []
    rowString = ""
    sampleCol = ""
    row = {}

    if repeatingFieldMethod == "value":
        rowString = "Example one https://www.brown.edu|Example two https://www.google.com"
    if repeatingFieldMethod == "name":
        rowString = "First example identity, 1980-, contributor https://www.brown.edu|Second example identity, 1990-2000, presenter https://library.brown.edu"
    
    element = keyField.get("element",[])
    textHeaders, conditionalAttrsHeaders, singleElementString = getConditionalAttrsTextHeadersElementFromField(element)
    columnHeaders.extend(textHeaders)
    row = convertArrayToDictWithMatchingKeyValues(removeItemsWithPeriodFromList(textHeaders))
    
    for colHeader in colHeaders:
        sampleCol = colHeader
        columnHeaders.append(colHeader)
        elements = handleRepeatingTypeEntry(parentElement, rowString, repeatingElement, {}, repeatingDefaults, row)
        elementsCreated.extend(elements)

    for colPrefix in colPrefixes:
        for (index, keyAuthority) in enumerate(keyAuthorities):
            colHeader = colPrefix + keyAuthority.get("suffix", "")
            columnHeaders.append(colHeader)
            if index == 0:
                sampleCol = colHeader
                elements = handleRepeatingTypeEntry(parentElement, rowString, repeatingElement, keyAuthority, repeatingDefaults, row)
                elementsCreated.extend(elements)

    elementString = ""

    for element in elementsCreated:
            cleanedElement = clearEmptyElementsFromEtree(element)
            elementString = elementString + "\n" + etree.tostring(cleanedElement, pretty_print=True, encoding="UTF-8").decode("utf-8")
            elementString = elementString.lstrip("\n").rstrip("\n")
    
    elementString = elementString.replace(replaceTextForExamples, "")
    columnHeaders = removeDuplicatesFromArray( removeItemsWithPeriodFromList(columnHeaders))

    return columnHeaders, conditionalAttrsHeaders, elementString, rowString, sampleCol

def getFieldReviewList():
    fieldList = []

    for keyField in keyFields:
        keyFieldType = keyField.get('type', "")

        if keyFieldType == 'element':
            textHeaders, conditionalAttrsHeaders, elementString = getConditionalAttrsTextHeadersElementFromField(keyField)
            field = {"headers": textHeaders, "conditionalattrs": conditionalAttrsHeaders, "elementstring": elementString}
            fieldList.append(field)
        if keyFieldType == "repeating":
            textHeaders, conditionalAttrsHeaders, elementString, sampleEntry, sampleCol = getConditionalAttrsTextHeadersElementFromRepeatingField(keyField)
            field = {"headers": textHeaders, "conditionalattrs": conditionalAttrsHeaders, "elementstring": elementString, "sampleentry":sampleEntry, "samplecol":sampleCol}
            print(field)
            fieldList.append(field)
            # elements = processRepeatingTypeField(keyField, keyAuthorities, keyNameSpace, row, parentElement) 
    return fieldList

getFieldReviewList()




