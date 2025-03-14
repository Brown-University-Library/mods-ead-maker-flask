import csv
from lxml import etree
from lxml import objectify
import yaml
import re
import os

def convertArrayToDictWithMatchingKeyValues(array):
    dictionary = {}

    for arrayItem in array:
        dictionary[arrayItem] = arrayItem

    return dictionary

def areAllDictValuesEmpty(dict={}):
    keys = dict.keys()

    for key in keys:
        if dict[key] == "":
            continue
        else:
            return False

    return True

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

def hasNumbers(s):
    return any(i.isdigit() for i in s)

def hasYear(s):
    numbercount = 0
    for i in s:
        if i.isdigit():
            numbercount = numbercount + 1
    if numbercount > 3:
        return True
    else:
        return False

def hasLetters(s):
    return re.search('[a-zA-Z]', s) is not None

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

def getAdditionalValues(name):
    additionalValuesResults = re.findall("\[(.*: .*)\]", name)
    additionalValues = {}

    if len(additionalValuesResults) > 0:
        for result in additionalValuesResults:
            additionalValue = yaml.safe_load(result)
            if type(additionalValue) is not dict:
                continue
            for key, value in additionalValue.items():
                additionalValues["entry." + key] = value

    return additionalValues, additionalValuesResults

def getValueUri(name):
    uris = re.findall("(?P<url>https?://[^\s]+)", name)

    #If there's a URI
    if len(uris) > 0:
        return normalizeString(uris[0])
    else:
        return ""

def legacyGetNameDateRoleFromEntry(entry):
    name = ""
    date = ""
    role = ""

    for textIndex, text in enumerate(entry.split(',')):
        normalizedText = normalizeString(text)

        if normalizedText == '':
            continue

        if textIndex == 0:
            name += normalizedText + ", "
        elif hasYear(normalizedText):
            date += normalizedText
            date = date.lstrip(',').rstrip(',')
        elif normalizedText.islower():
            role = text
        elif hasLetters(normalizedText) is not None:
            name = name + normalizedText + " "

    return normalizeString(name).rstrip(",").lstrip(", "), normalizeString(date), normalizeString(role)

def getNameDateRoleFromEntry(entry, method):
    '''
    Expects a string like 'Murphy, Connor, 2023-2024, Library Technologist' or
    like 'Value' and
    returns 3 strings like 'Murphy, Connor', '2023-2024', 'Library Technologist'
    or 1 string like 'Value'
    '''
    norm_entry = normalizeString(entry)

    if 'name' not in method:
        return norm_entry, "", ""

    if method == "nameLegacy":
        return legacyGetNameDateRoleFromEntry(norm_entry)

    if not norm_entry:
        return "","",""

    name, date, role = ["", "", ""]
    parts_list = [part.strip() for part in norm_entry.split(',')]

    if method == "nameOther":
        role = parts_list.pop()
    if method == "nameCreator":
        role = "Creator"
    if re.match(r'\d{3,}',parts_list[-1]):
        date = parts_list.pop()
    name = ', '.join(parts_list)

    return name, date, role

def getKeyValueFromEntry(string):
    '''
    Turns a string like key: value into two strings, key and value
    '''
    [key,val] = string.split(":")
    return key.strip(), val.strip()

def getMetadataFromEntry(entry, method):
    valueUri = getValueUri(entry)
    entry = entry.replace(valueUri, "")
    value = normalizeString(entry)

    additionalValues, additionalValuesResults = getAdditionalValues(entry)
    for additionalValuesResult in additionalValuesResults:
        entry = entry.replace(additionalValuesResult, "")

    prependTermOfAddress, appendTermOfAddress = getTermsOfAddressPrependAndAppendStripped(entry)
    entry = entry.replace("{{" + prependTermOfAddress + "}}", "")
    entry = entry.replace("{{" + appendTermOfAddress + "}}", "")

    name = ''
    date = ''
    role = ''
    key = ''
    val = ''

    if entry:
        if "name" in method:
            name, date, role = getNameDateRoleFromEntry(entry, method)
        if method == "keyValue":
            key, val = getKeyValueFromEntry(entry)

    metadata = {
        "entry.value": value,
        "entry.valueURI":valueUri,
        "entry.name": name,
        "entry.date": date,
        "entry.role": role,
        "entry.prependTermOfAddress": prependTermOfAddress,
        "entry.appendTermOfAddress":appendTermOfAddress,
        "entry.key":key,
        "entry.val":val
    }

    if additionalValues:
        metadata.update(additionalValues)

    return metadata


def normalizeString(string):
    if string is not None:
        string = string.replace('\n', ' ').replace('\r', ' ')
        string = string.replace('<title>', '').replace('</title>', '')
        string = string.replace('<geogname>', '- ').replace('</geogname>', '')
        string = string.replace('|', "")
        return(' '.join(str(string).split()))
    else:
        return string

def clearEmptyElementsFromEtree(parentElement, keepElementXpaths):

    if parentElement is None:
        return etree.Element("blank")
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
        for xpath in keepElementXpaths:
            for element in clean.xpath(xpath.get("elementxpath", None), namespaces=parentElement.nsmap):
                print(etree.tostring(element))
            if e in clean.xpath(xpath.get("elementxpath", None), namespaces=parentElement.nsmap):
                print("Skipping because of xpath")
                return
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

class Profile():

    def __init__(self, profileDirectory, globalConditions={}):
        self.profile = yaml.safe_load(open(os.path.join( os.path.dirname(os.path.abspath(__file__)), profileDirectory)))

        self.profileSkips = self.profile.get("skipif", [])
        self.profileFields = self.profile.get("fields", [])
        self.colSuffixes = self.profile.get("authorities")
        self.profileSorts = self.profile.get("sort", [])
        self.profileKeepElementXpaths = self.profile.get("keepblanktextelements", [])
        self.profileSampleValues = self.profile.get("samplevalues", {})

        self.profileGlobalConditions = self.profile.get("globalconditions", [])
        self.globalConditionsSet = globalConditions

        self.profileFilenameColumn = self.profile.get("filenamecolumn", "")
        self.profileFileExtension = self.profile.get("fileextension", ".xml")

        self.profileNameSpace = self.profile.get("elementnamespace", [])
        self.profileParentTag = self.profile.get("parenttag", [])
        self.profileQName = self.profile.get("attrqname", [])
        self.profileSchemaLocation = self.profile.get("schemalocation", [])
        self.profileNsMap = self.profile.get("nsmap", {})

        with open(os.path.join( os.path.dirname(os.path.abspath(__file__)),"SupportedLanguages.csv")) as langCsv:
            languageToLanguageCodes = {}
            csvReader = csv.reader(langCsv)
            for row in csvReader:
                languageToLanguageCodes[row[0]] = row[1]
            self.languageCodes = languageToLanguageCodes

    def shouldSkipRow(self, row):
        for profileSkip in self.profileSkips:
            if row.get(profileSkip, "") != "":
                return True

    def createParentElement(self, nsMap):
        attrQname = etree.QName(self.profileQName.get("uri",""),self.profileQName.get("tag",""))

        return etree.Element(self.profileNameSpace + self.profileParentTag, {attrQname: self.profileSchemaLocation}, nsmap=nsMap)

    def createSubElement(self, parentElement, elementName, elementAttrs, elementText):
        subElement = etree.SubElement(parentElement, self.profileNameSpace + elementName, elementAttrs)
        subElement.text = normalizeString(elementText)

        return subElement

    def processLanguageTextValue(self, language):
        if len(normalizeString(language)) > 3:
            if normalizeString(language) in self.languageCodes:
                return self.languageCodes[normalizeString(language)]
            else:
                return language
        else:
            return language

    def processColumnTextValue(self, column, row):
        columnMethod = column.get("method")
        columnHeader = column.get("header")

        if columnMethod  == "value":
            text = row.get(columnHeader, "")
            return normalizeString(text)
        if columnMethod  == "num":
            text = row.get(columnHeader, "").replace(".0", "")
            return normalizeString(text)
        if columnMethod  == "lower":
            text = row.get(columnHeader, "").lower()
            return normalizeString(text)
        if columnMethod == "lang":
            text = row.get(columnHeader, "")
            return self.processLanguageTextValue(text)

        return row.get(columnHeader,"")

    def processConditionalAttrs(self, conditionalAttr, row, element):
        key = conditionalAttr.get("key","")

        textKeyField = conditionalAttr.get("text","")
        text = self.processTextUnit(textKeyField, row)

        if text:
            element.set(key, text)

    def processTextUnitValues(self, textUnitValues, row):
        text = ""
        for value in textUnitValues:
            valueType = value.get("type","")
            valueText = value.get("text","")

            if valueType == "value":
                text = text + valueText

            if valueType == "col":
                text = text + self.processColumnTextValue(value, row)

        return text

    def performTextAction(self, textAction, text):
        if textAction.get("action") == "leftstriprightstrip":
            lstripRstripText = textAction.get("leftstriprightstriptext", "")
            return text.lstrip(lstripRstripText).rstrip(lstripRstripText)

    def processTextUnit(self, textUnits, row):
        text = ""
        for textUnit in textUnits:
            textUnitType = textUnit.get("type","")
            textUnitColumn = textUnit.get("col", None)
            textUnitValues = textUnit.get("values","")
            textUnitText = textUnit.get("text","")

            if  textUnitType == "ifpresent":
                if row.get(textUnitColumn):
                    newText = self.processTextUnitValues(textUnitValues, row)
                    text = text + newText
            if  textUnitType == "ifnotpresent":
                if row.get(textUnitColumn) is None or row.get(textUnitColumn) == "":
                    newText = self.processTextUnitValues(textUnitValues, row)
                    text = text + newText
            if  textUnitType == "ifhas":
                if textUnitText in row.get(textUnitColumn, ""):
                    newText = self.processTextUnitValues(textUnitValues, row)
                    text = text + newText
            if  textUnitType == "value":
                newText = self.processTextUnitValues(textUnitValues, row)
                text = text + newText
            if  textUnitType == "removetext":
                newText = self.processTextUnitValues(textUnitValues, row)
                replaceStrings = textUnit.get("removetext",[])
                for replaceString in replaceStrings:
                    newText = newText.replace(replaceString, "")
                text = text + newText
            if  textUnitType == "action":
                text = self.performTextAction(textUnit, text)

        return text

    def shouldCreateElementBasedOnCondition(self, condition, row):
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

        if conditionType == "global":
            globalConditionCode = condition.get("code","")
            if self.globalConditionsSet.get(globalConditionCode):
                return True

        return False

    def processElementTypeField(self, profileField, row, parentElement):
        elementName = profileField.get("name", "")
        elementAttrs = profileField.get("attrs", {})
        element = self.createSubElement(parentElement, elementName, elementAttrs, "")
        conditions = profileField.get("conditions", [])
        entryrole = row.get('entry.role','')

        if re.search(r'&&',entryrole) and elementName == "roleTerm":
            roles:list[str] = entryrole.split('&&')
            for role in roles:
                role.strip()
                newrow = {**row}
                newrow["entry.role"] = role
                self.processElementTypeField(profileField, newrow, parentElement)
            return element

        for condition in conditions:
            shouldCreateElement = self.shouldCreateElementBasedOnCondition(condition, row)
            if not shouldCreateElement:
                return

        conditionalAttrs = profileField.get("conditionalattrs", [])
        for conditionalAttr in conditionalAttrs:
            self.processConditionalAttrs(conditionalAttr, row, element)

        childrenKeyFields = profileField.get("children", {})
        for childKeyField in childrenKeyFields:
            self.processElementTypeField(childKeyField, row, element)

        textKeyField = profileField.get("text", [])
        text = self.processTextUnit(textKeyField, row)
        if text:
            element.text = text

        return element

    def handleRepeatingEntries(self, parentElement, entriesString, profileElement, repeatingDefaults, originalRow, method):
        entries = entriesString.split("|")
        elementsCreated = []

        for entry in entries:
            entryAdditions = getMetadataFromEntry(entry, method)
            if areAllDictValuesEmpty(entryAdditions) is False:
                entryAdditions.update(repeatingDefaults)
                entryAdditions.update(originalRow)
            element = self.processElementTypeField(profileElement, entryAdditions, parentElement)
            elementsCreated.append(element)

        return elementsCreated

    def processRepeatingTypeField(self, profileField, keyAuthorities, row, parentElement):
        repeatingMethod = profileField.get("method")
        colPrefixes = profileField.get("colprefix", [])
        colSuffixes = profileField.get("colsuffixes", [])
        colHeaders = profileField.get("cols", [])
        repeatingElement = profileField.get("element", {})
        repeatingDefaults = profileField.get("defaults", {})

        if repeatingMethod in ["nameCreator", "nameOther", "keyValue", "value"]:
            for colPrefix in colPrefixes:
                for colSuffix in colSuffixes:
                    colHeader = colPrefix + colSuffix.get("suffix", "")
                    colSuffixDefaults = colSuffix.get("defaults",{})
                    colSuffixDefaults.update(repeatingDefaults)
                    rowString = row.get(colHeader, "").replace(';','|')
                    self.handleRepeatingEntries(parentElement, rowString, repeatingElement, colSuffixDefaults, row, repeatingMethod)
            for colHeader in colHeaders:
                rowString = row.get(colHeader, "").replace(';','|')
                self.handleRepeatingEntries(parentElement, rowString, repeatingElement, repeatingDefaults, row, repeatingMethod)

    def processSort(self,parentElement, sort):
        elementXpath = sort.get("elementxpath", "")
        sortByXpath = sort.get("sortbyxpath", "")

        allMatchingElements = parentElement.xpath(elementXpath, namespaces=parentElement.nsmap)
        firstElementIndex = 0

        if len(allMatchingElements) > 0:
            firstElementIndex = parentElement.getchildren().index(allMatchingElements[0])

        allMatchingElements = sorted(allMatchingElements, key=lambda ch: ch.xpath(sortByXpath, namespaces={self.profileParentTag: self.profileNameSpace.replace("{","").replace("}","")}))

        for element in allMatchingElements:
            parentElement.remove(element)

        for (index, element) in enumerate(allMatchingElements):
            parentElement.insert(firstElementIndex + index, element)

    def getColumnHeadersFromTextFields(self, textFields):
        textHeaders = []
        for textField in textFields:
            textFieldValues = textField.get("values", [])
            for textFieldValue in textFieldValues:
                if textFieldValue.get("type") == "col":
                    header = textFieldValue.get("header", "")
                    textHeaders.append(header)

        return textHeaders

    def getFieldListInfoFromConditionalAttr(self, conditionalAttr, name):
        textHeaders = []
        conditionalAttrConditions = []

        attrKey = conditionalAttr.get("key", "")
        attrTextFields = conditionalAttr.get("text", [])
        textHeaders = self.getColumnHeadersFromTextFields(attrTextFields)

        for attrTextField in attrTextFields:
            attrTextType = attrTextField.get("type","")
            if attrTextType == "ifpresent":
                string = "If text is entered in the " + attrTextField.get("col","") + " column, the following attribute will be added to the <" + self.profileParentTag + ":" + name + "> element:"
                value = self.processTextUnitValues(attrTextField.get("values",[]), convertArrayToDictWithMatchingKeyValues(textHeaders))
                attrString = attrKey + "='" + value + "'"
                conditionalAttrCondition = {"explanation":string, "attribute":attrString}
                conditionalAttrConditions.append(conditionalAttrCondition)
                textHeaders.append(attrTextField.get("col",""))
            if attrTextType == "ifnotpresent":
                string = "If text is not entered in the " + attrTextField.get("col","") + " column, the following attribute will be added to the <" + self.profileParentTag + ":" + name + "> element:"
                value = self.processTextUnitValues(attrTextField.get("values",[]), convertArrayToDictWithMatchingKeyValues(textHeaders))
                attrString = attrKey + "='" + value + "'"
                conditionalAttrCondition = {"explanation":string, "attribute":attrString}
                conditionalAttrConditions.append(conditionalAttrCondition)
                textHeaders.append(attrTextField.get("col",""))

        return textHeaders, conditionalAttrConditions

    def getFieldListInfoFromCondition(self, condition):
        if condition.get("type","") == "global":
            return 'Only appears if the "' + condition.get("code", "") + '" condition is set.'
        if condition.get("type","") == "startswith":
            return 'Only appears if the text in column "' + condition.get("col", "") + '" starts with "' + condition.get("text", "") + '".'
        if condition.get("type","") == "has":
            return 'Only appears if the text in column "' + condition.get("col", "") + '" contains "' + condition.get("text", "") + '".'

    def createExampleElementTextFromEtree(self, element):
        cleanedElement = clearEmptyElementsFromEtree(element, self.profileKeepElementXpaths)
        objectify.deannotate(cleanedElement, cleanup_namespaces=True, xsi_nil=True)
        elementString = etree.tostring(cleanedElement, pretty_print=True, encoding="UTF-8").decode("utf-8")

        return elementString

    def getFieldListInfoFromElementField(self, profileField):
        keySampleValues = self.profile.get("samplevalues",{})
        name = profileField.get("name", [])
        textFields = profileField.get("text", [])
        parentElement = self.createParentElement(self.profileNsMap)

        textHeaders = []
        conditionalAttrConditions = []
        conditions = []

        textHeaders = self.getColumnHeadersFromTextFields(textFields)

        for condition in profileField.get("conditions",[]):
            conditionText = self.getFieldListInfoFromCondition(condition)
            conditions.append(conditionText)

        conditionalAttrs = profileField.get("conditionalattrs", [])
        for conditionalAttr in conditionalAttrs:
            conditionalAttrTextHeaders, conditionalAttrCondition = self.getFieldListInfoFromConditionalAttr(conditionalAttr, name)
            textHeaders.extend(conditionalAttrTextHeaders)
            conditionalAttrConditions.extend(conditionalAttrCondition)

        for child in profileField.get("children",[]):
            childTextHeaders, childConditionalAttrsHeaders, elementString, childConditions = self.getFieldListInfoFromElementField(child)
            textHeaders.extend(childTextHeaders)
            conditionalAttrConditions.extend(childConditionalAttrsHeaders)
            conditions.extend(childConditions)

        textHeaderRow = convertArrayToDictWithMatchingKeyValues(textHeaders)
        textHeaderRow.update(keySampleValues)

        element = self.processElementTypeField(profileField, textHeaderRow, parentElement)
        elementString = self.createExampleElementTextFromEtree(element)

        return removeDuplicatesFromArray(textHeaders), conditionalAttrConditions, elementString, conditions

    def getFieldListInfoFromRepeatingField(self, profileField):
        parentElement = self.createParentElement(self.profileNsMap)
        repeatingElement = profileField.get("element", {})
        repeatingDefaults = profileField.get("defaults", {})
        colPrefixes = profileField.get("colprefix", [])
        colHeaders = profileField.get("cols", [])
        repeatingFieldMethod = profileField.get("method", "")
        colSuffixes = profileField.get("colsuffixes", [])

        columnHeaders = []
        elementsCreated = []
        conditions = []
        rowString = ""
        sampleCol = ""
        row = {}

        if repeatingFieldMethod == "value":
            rowString = "Example one https://www.brown.edu|Example two https://www.google.com"
        if repeatingFieldMethod == "nameCreator":
            rowString = "Identity1, First Example, 1980- https://www.brown.edu|Example, Second, 1900-1999 http://library.brown.edu"
        if repeatingFieldMethod == "nameOther":
            rowString = "Name, Person's, Three Commas, 1980-, Contributor&&Other Role https://www.brown.edu|Corp Name, 1900-1999, Long-time Funder http://library.brown.edu"
        if repeatingFieldMethod == "keyValue":
            rowString = "key1:value1|key2:value2"
        element = profileField.get("element",[])
        textHeaders, conditionalAttrsHeaders, singleElementString, conditions = self.getFieldListInfoFromElementField(element)
        row = convertArrayToDictWithMatchingKeyValues(removeItemsWithPeriodFromList(textHeaders))

        for colPrefix in colPrefixes:
            for (index, colSuffix) in enumerate(colSuffixes):
                colHeader = colPrefix + colSuffix.get("suffix", "")
                columnHeaders.append(colHeader)
                if index == 0:
                    sampleCol = colHeader
                    colSuffixDefaults = colSuffix.get("defaults",{})
                    colSuffixDefaults.update(repeatingDefaults)

                    if self.profileSampleValues.get(colHeader):
                        rowString = self.profileSampleValues.get(colHeader)

                    elements = self.handleRepeatingEntries(parentElement, rowString, repeatingElement, colSuffixDefaults, row, repeatingFieldMethod)
                    elementsCreated.extend(elements)

        for colHeader in colHeaders:
            sampleCol = colHeader
            columnHeaders.append(colHeader)

            if self.profileSampleValues.get(colHeader):
                rowString = self.profileSampleValues.get(colHeader)

            elements = self.handleRepeatingEntries(parentElement, rowString, repeatingElement, repeatingDefaults, row, repeatingFieldMethod)
            elementsCreated.extend(elements)

        columnHeaders.extend(textHeaders)
        elementString = ""

        for element in elementsCreated:
                singleElementString = self.createExampleElementTextFromEtree(element)
                elementString = elementString + "\n" + singleElementString
                elementString = elementString.lstrip("\n").rstrip("\n")

        columnHeaders = removeDuplicatesFromArray( removeItemsWithPeriodFromList(columnHeaders))

        return columnHeaders, conditionalAttrsHeaders, elementString, rowString, sampleCol, conditions

    def getFieldList(self):
        fieldList = []

        for globalCondition in self.profileGlobalConditions:
            self.globalConditionsSet[globalCondition.get("code","")] = True

        for profileField in self.profileFields:
            profileFieldType = profileField.get('type', "")
            if profileFieldType == 'element':
                textHeaders, conditionalAttrsHeaders, elementString, conditions = self.getFieldListInfoFromElementField(profileField)
                field = {"headers": textHeaders, "conditionalattrs": conditionalAttrsHeaders, "elementstring": elementString, "conditions": conditions}
                fieldList.append(field)
            if profileFieldType == "repeating":
                textHeaders, conditionalAttrsHeaders, elementString, sampleEntry, sampleCol, conditions = self.getFieldListInfoFromRepeatingField(profileField)
                field = {"headers": textHeaders, "conditionalattrs": conditionalAttrsHeaders, "elementstring": elementString, "sampleentry":sampleEntry, "samplecol":sampleCol, "conditions":conditions}
                fieldList.append(field)
        return fieldList

    def getAllHeaders(self):
        fieldList = self.getFieldList()
        allHeaders = []

        for field in fieldList:
            for header in field.get("headers", []):
                allHeaders.append(header)

        return removeDuplicatesFromArray(allHeaders)


    def convertRowToXmlString(self, row):
        if self.shouldSkipRow(row):
            return

        parentElement = self.createParentElement(self.profileNsMap)

        for profileField in self.profileFields:
            profileFieldType = profileField.get('type', "")

            if profileFieldType == 'element':
                self.processElementTypeField(profileField, row, parentElement)

            if profileFieldType == "repeating":
                self.processRepeatingTypeField(profileField, self.colSuffixes, row, parentElement)

        cleanedUpEtree = clearEmptyElementsFromEtree(parentElement, self.profileKeepElementXpaths)

        for profileSort in self.profileSorts:
            self.processSort(cleanedUpEtree, profileSort)

        return etree.tostring(cleanedUpEtree, pretty_print=True, encoding="unicode")