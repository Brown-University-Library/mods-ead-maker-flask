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

requiredcolumns = ["subjectTopicsFAST","Ignore", "fileTitle", "itemTitle", "subTitle", "place","dateText","dateStart","dateEnd","dateBulkStart","dateBulkEnd","dateQualifier", "shelfLocator1", "shelfLocator1ID", "shelfLocator2", "shelfLocator2ID", "shelfLocator3","shelfLocator3ID","typeOfResource","genreAAT","genreLCSH","genreLocal","genreRBGENR","extentQuantity","extentSize","extentSpeed","form","noteScope","noteHistorical","noteHistoricalClassYear","noteGeneral","language","noteAccession","identifierBDR","publisher","namePersonCreatorLC","namePersonCreatorLocal","nameCorpCreatorLC","nameCorpCreatorLocal","namePersonOtherLC","namePersonOtherLocal","subjectNamesLC","subjectNamesLocal","subjectCorpLC","subjectCorpLocal","subjectTopicsLC","subjectTopicsLocal","subjectGeoLC","subjectTemporalLC","subjectTitleLC","collection","dateTextParent","callNumber","repository","findingAid","digitalOrigin","rightsStatementText","rightsStatementURI", "useAndReproduction", "coordinates", "scale", "projection"]
authorityURIs = {"aat": "https://www.getty.edu/research/tools/vocabularies/aat/","local":"","rbgenr":"https://rbms.info/vocabularies/genre/alphabetical_list.htm","lcsh":"http://id.loc.gov/authorities/subjects.html","fast":"http://id.worldcat.org/fast","naf":"http://id.loc.gov/authorities/names.html"}
langcode = {}
langcodeopp = {}
scriptcode = {}


def getSplitCharacter(string):
    if ";" in string:
        return(";")
    else:
        return("|")

def messageToUser(messagetitle, message):
    print("")
    print(messagetitle)
    print(message)
    try:
        raw_input("Press Enter to continue . . .")
    except SyntaxError:
        print("Syntax Error")
    except TypeError:
        print("Type Error")

def multilinefield(parentelement, originalfieldname, eadfieldname):
    newelement = etree.SubElement(parentelement, eadfieldname)
    lines = cldata.get(originalfieldname, '').splitlines()
    for line in lines:
        pelement = etree.SubElement(newelement, "p")
        pelement.text = ' '.join(line.split())

def repeatingfield(parentelement, rowString, modsfieldname, modsattributes, subject, subjectattributes):

    originalparentelement = parentelement

    for namesindex, addedentry in enumerate(rowString.split("|")):

        customSubjectAttributes = subjectattributes.copy()
        customMODSattributes = modsattributes.copy()

        #Extract URI
        if subject:
            addedentry, customSubjectAttributes = getUri(addedentry, customSubjectAttributes)
        else:
            addedentry, customMODSattributes = getUri(addedentry, customMODSattributes)
        # uri = re.findall("(?P<url>https?://[^\s]+)", addedentry)

        # #If there's a URI
        # if len(uri) > 0:
        #     #Remove it from the addedentry
        #     addedentry = addedentry.replace(uri[0],"")
        #     #Add it as a valueURI attribute
        #     if subject:
        #         customSubjectAttributes["valueURI"] = normalizeString(uri[0])
        #     else:
        #         customMODSattributes["valueURI"] = normalizeString(uri[0])

        # #Add authorityURI attribute

        # if subject:

        #     authorityType = customSubjectAttributes.get("authority", "")

        #     if authorityURIs.get(authorityType):
        #         customSubjectAttributes["authorityURI"] = authorityURIs.get(authorityType)

        # else:
        #     authorityType = customMODSattributes.get("authority", "")

        #     if authorityURIs.get(authorityType):
        #         customMODSattributes["authorityURI"] = authorityURIs.get(authorityType)

        #Create field

        if subject == True:
            subjectelement = etree.SubElement(parentelement, "{http://www.loc.gov/mods/v3}subject", customSubjectAttributes)
            parentelement = subjectelement

        namecontrolaccesselement = etree.SubElement(parentelement, modsfieldname, customMODSattributes)
        namecontrolaccesselement.text = ' '.join(addedentry.split())

        parentelement = originalparentelement

def repeatingTitleSubjectField(modstop, rowString, attributes):

    for title in rowString.split('|'):

        customAttributes = attributes.copy()

        #Extract URI
        title, customAttributes = getUri(title, customAttributes)

        #Create element
        subjecttitleparentelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}subject", customAttributes)
        subjecttitleinfoelement = etree.SubElement(subjecttitleparentelement, "{http://www.loc.gov/mods/v3}titleInfo")
        subjecttitleelement = etree.SubElement(subjecttitleinfoelement, "{http://www.loc.gov/mods/v3}title")
        subjecttitleelement.text = normalizeString(title)

def getTermsOfAddressPrependAndAppend(name):
    appendTermsOfAddress = []
    prependTermsOfAddress = []

    for textIndex, text in enumerate(name.split(',')):
        if textIndex == 0:
            appendTermsOfAddress = re.findall("(\{\{.*\}\})", text)

            for appendTermOfAddress in appendTermsOfAddress:
                name = name.replace(appendTermOfAddress, "")
        
        if textIndex == 1:
            prependTermsOfAddress = re.findall("(\{\{.*\}\})", text)

            for prependTermOfAddress in prependTermsOfAddress:
                name = name.replace(prependTermOfAddress, "")
    
    return name, prependTermsOfAddress, appendTermsOfAddress

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

def getUri(name, customAttributes):
    uri = re.findall("(?P<url>https?://[^\s]+)", name)

    #If there's a URI
    if len(uri) > 0:
        #Remove it from the addedentry
        name = name.replace(uri[0],"")
        #Add it as a valueURI attribute
        customAttributes["valueURI"] = normalizeString(uri[0])

    #Add authorityURI attribute
    authorityType = customAttributes.get("authority", "")

    if authorityURIs.get(authorityType):
        customAttributes["authorityURI"] = authorityURIs.get(authorityType)
    
    return name, customAttributes

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

    return {"entry.value": value, "entry.name": name, "entry.date": date, "entry.role": role, "entry.prependTermOfAddress": prependTermOfAddress, "entry.appendTermOfAddress":appendTermOfAddress}

def getRepeatingValueMetadataFromEntry(entry):
    valueUri = getValueUri(entry)
    entry = entry.replace(valueUri, "")

    value = normalizeString(entry)

    return {"entry.value": value, "entry.valueURI": valueUri}

def repeatingnamefield(parentelement, rowString, topmodsattributes, predefinedrole, subject):
    originalparentelement = parentelement

    for nameindex, name in enumerate(rowString.split("|")):
        nametext = ""
        datetext = ""
        roletext = predefinedrole
        customAttributes = topmodsattributes.copy()

        #Extract URI
        name, customAttributes = getUri(name, customAttributes)

        #Extract termsOfAddress 
        name, prependTermsOfAddress, appendTermsOfAddress = getTermsOfAddressPrependAndAppend(name)

        for textindex, text in enumerate(name.split(',')):
            textrevised = ' '.join(text.split())

            if textrevised == '':
                continue
            
            max_index = len(normalizeString(name).split(','))-1

            if textindex == 0:
                nametext = nametext + textrevised + ", "
            elif hasYear(textrevised) == True:
                datetext = datetext + textrevised
            elif isAllLower(textrevised) == True:
                roletext = text
            elif hasLetters(textrevised) != None:
                nametext = nametext + textrevised + " "

        if nametext == '':
            continue

        if subject == True:
            subjectelement = etree.SubElement(parentelement, "{http://www.loc.gov/mods/v3}subject")
            parentelement = subjectelement

        nameelement = etree.SubElement(parentelement, "{http://www.loc.gov/mods/v3}name", customAttributes)
        for termOfAddress in prependTermsOfAddress:
            termOfAddress = termOfAddress.replace("{{","").replace("}}","")
            termOfAddressElement = etree.SubElement(nameelement, "{http://www.loc.gov/mods/v3}namePart", {"type":"termsOfAddress"})
            termOfAddressElement.text = termOfAddress
        namepart = etree.SubElement(nameelement, "{http://www.loc.gov/mods/v3}namePart")
        namepart.text = normalizeString(nametext).rstrip(',')
        for termOfAddress in appendTermsOfAddress:
            termOfAddress = termOfAddress.replace("{{","").replace("}}","")
            termOfAddressElement = etree.SubElement(nameelement, "{http://www.loc.gov/mods/v3}namePart", {"type":"termsOfAddress"})
            termOfAddressElement.text = termOfAddress
        namedatepart = etree.SubElement(nameelement, "{http://www.loc.gov/mods/v3}namePart", {"type":"date"})
        namedatepart.text = normalizeString(datetext).lstrip(',').rstrip(',')
        modsrole = etree.SubElement(nameelement, "{http://www.loc.gov/mods/v3}role")
        modsroleterm = etree.SubElement(modsrole, "{http://www.loc.gov/mods/v3}roleTerm", {"type":"text", "authority":"marcrelator"})
        modsroleterm.text = normalizeString(roletext).lstrip(',').rstrip(',')

        parentelement = originalparentelement

def normalizeString(string):
    if string != None:
        string = string.replace('\n', ' ').replace('\r', ' ')
        string = string.replace('<title>', '').replace('</title>', '')
        string = string.replace('<geogname>', '- ').replace('</geogname>', '')
        return(' '.join(str(string).split()))
    else:
        return string

def convertEncoding(from_encode,to_encode,old_filepath,target_file):
    f1=open(old_filepath)
    content2=[]
    while True:
        line=f1.readline()
        content2.append(line.decode(from_encode).encode(to_encode))
        if len(line) ==0:
            break

    f1.close()
    f2=open(target_file,'w')
    f2.writelines(content2)
    f2.close()

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



def let_user_pick(message, options):
    print("")
    print(message)
    for idx, element in enumerate(options):
        print("{}) {}".format(idx+1,element))
    i = input("Enter number: ")
    try:
        if 0 < int(i) <= len(options):
            return options[i-1]
    except:
        pass
    return None

def XLSDictReader(file, sheetname):
        book    = xlrd.open_workbook(file)
        sheet   = book.sheet_by_name(sheetname)

        rowarray = []

        for row in range(1, sheet.nrows):
            rowdictionary = {}
            for column in range(sheet.ncols):
                #If the value is a number, turn it into a string.
                newvalue = ''
                if sheet.cell(row,column).ctype > 1:
                    newvalue = str(sheet.cell_value(row,column)).replace('|',';')
                else:
                    newvalue = sheet.cell_value(row,column).replace('|',';')

                #If the column is repeating, serialize the row values.
                if rowdictionary.get(sheet.cell_value(0,column), '') != '':
                    rowdictionary[sheet.cell_value(0,column)] = rowdictionary[sheet.cell_value(0,column)] + ';' + newvalue
                else:
                    rowdictionary[sheet.cell_value(0,column)] = newvalue
            rowarray.append(rowdictionary)
        return(rowarray)

def XLSDictReaderLanguageCode(file, sheetname):
        book    = xlrd.open_workbook(file)
        sheet   = book.sheet_by_name(sheetname)

        langcode = {}

        for row in range(sheet.nrows):
            key = sheet.cell_value(row, 0)
            value = sheet.cell_value(row, 1)
            langcode[key] = value
        return(langcode)

def XLSDictReaderLanguageCodeOpp(file, sheetname):
        book    = xlrd.open_workbook(file)
        sheet   = book.sheet_by_name(sheetname)

        langcode = {}

        for row in range(sheet.nrows):
            key = sheet.cell_value(row, 1)
            value = sheet.cell_value(row, 0)
            langcode[key] = value
        return(langcode)

def XLSDictReaderScriptCode(file, sheetname):
        book    = xlrd.open_workbook(file)
        sheet   = book.sheet_by_name(sheetname)

        scriptcode = {}

        for row in range(sheet.nrows):
            key = sheet.cell_value(row, 0)
            value = sheet.cell_value(row, 2)
            scriptcode[key] = value
        return(scriptcode)


CACHEDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache") + "/"
#CACHEDIR = os.path.dirname(os.path.abspath(__file__)) + "/"
HOMEDIR = os.path.dirname(os.path.abspath(__file__)) + "/"
#HOMEDIR = os.path.dirname(os.path.abspath(__file__)) + "/"

#print("._. MODS Maker ._.")

def processExceltoMODS(chosenfile, chosensheet, id, includeDefaults):

        print("INCLUDE DEFAULTS")
        print(includeDefaults)
    #try:

        if not os.path.exists(CACHEDIR + id):
            os.mkdir(CACHEDIR + id)

        #Create zipfile to hold files
        zipObj = ZipFile(CACHEDIR + id + "/" + chosensheet + '.zip', 'w')

        langcode = XLSDictReaderLanguageCode(HOMEDIR + "SupportedLanguages.xlsx","languages xlsx")
        langcodeopp = XLSDictReaderLanguageCodeOpp(HOMEDIR + "SupportedLanguages.xlsx","languages xlsx")
        scriptcode = XLSDictReaderScriptCode(HOMEDIR + "SupportedLanguages.xlsx","languages xlsx")

        #Extract spreadsheet data to csvdata dictionary.
        csvdata = {}
        langissue = False

        #For preview
        allrecords = ""

        excel = xlrd.open_workbook(chosenfile)
        selectedsheet = excel.sheet_by_name(chosensheet)
        columnsinsheet = [str(cell.value) for cell in selectedsheet.row(0)]

        missingcolumns = []
        for column in requiredcolumns:
            if (column in columnsinsheet) == False:
                #print("Missing spreadsheet column: " + column + '\n', file=sys.stderr)
                missingcolumns.append(column)

        if len(missingcolumns) != 0:
            print("*Missing Columns Detected*" + '\n', file=sys.stderr)
            print("The columns below are missing from your spreadsheet. The script will continue without them." + '\n\n', file=sys.stderr)

            for column in missingcolumns:
                print("   " + column + '\n', file=sys.stderr)


        print('\n\n', file=sys.stderr)


        csvdata = XLSDictReader(chosenfile, chosensheet)
        chosenfile = chosensheet

        #Create the output directory and save the path to the output_path variable.
        #now = datetime.datetime.now()
        #output_path = os.path.dirname(os.path.abspath(__file__))

        #try:
        #     os.mkdir(output_path + '/'+ chosenfile + " " + now.strftime("%m-%d-%Y %H %M " + str(now.second)))
        #except OSError:
        #     print ("" )
        #else:
       #      print ("")
       #      output_path = output_path + '/'+ chosenfile + " " + now.strftime("%m-%d-%Y %H %M " + str(now.second))

        #Create the error CSV.
        #errorfile = open(output_path + '/Error Report ' + now.strftime("%m-%d-%Y %H %M " + str(now.second)) + '.csv', mode='wb')
        #errorcsvwriter = csv.writer(errorfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        #errorcsvwriter.writerow(['Spreadsheet Row', 'BDR Number', 'Column Name', 'Column Contents', 'Potential Issue'])

        #Set up namespaces and attributes for XML.
        attr_qname = etree.QName("http://www.w3.org/2001/XMLSchema-instance", "schemaLocation")
        ns_map = {"mods" : "http://www.loc.gov/mods/v3", "xsi" : "http://www.w3.org/2001/XMLSchema-instance", "xlink" : "http://www.w3.org/1999/xlink"}

        #root = etree.Element("{http://www.loc.gov/mods/v3}modsCollection", {attr_qname: "http://www.loc.gov/mods/v3 http://www.loc.gov/mods/v3/mods-3-7.xsd"}, nsmap=ns_map)
        #print(etree.tostring(root))

        amountofrecords = 0
        rowindex = 2

        #Create a MODS file for every row in the input CSV file.
        for row in csvdata:

            #Ignore rows that contain EAD-specific data or anything in the Ignore column.
            if row.get('recordgroupTitle', '') != '':
                continue
            if row.get('subgroupTitle', '') != '':
                continue
            if row.get('seriesTitle', '') != '':
                continue
            if row.get('subSeriesTitle', '') != '':
                continue
            if row.get('Ignore', '') != '':
                continue

            #Set up the top-level mods element.
            modstop = etree.Element("{http://www.loc.gov/mods/v3}mods", {attr_qname: "http://www.loc.gov/mods/v3 http://www.loc.gov/mods/v3/mods-3-7.xsd"}, nsmap=ns_map)

            #mods:titleInfo
            titleinfo = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}titleInfo")
            title = etree.SubElement(titleinfo, "{http://www.loc.gov/mods/v3}title")
            if row.get("fileTitle", '') != "":
                title.text = normalizeString(row.get("fileTitle", ''))
            #elif row.get("title", '') != "":
            #    title.text = normalizeString(row.get("title", ''))
            else:
                title.text = normalizeString(row.get("itemTitle", ''))
            #title.text = normalizeString(row.get("title"]) #' '.join(row["title", '').split())
            subtitle = etree.SubElement(titleinfo, "{http://www.loc.gov/mods/v3}subTitle")
            subtitle.text = normalizeString(row.get("subTitle", ''))

            partNumberElement = etree.SubElement(titleinfo, "{http://www.loc.gov/mods/v3}partNumber")
            partNumberElement.text = normalizeString(row.get("itemTitlePartNumber","").replace(".0",""))

            partNameElement = etree.SubElement(titleinfo, "{http://www.loc.gov/mods/v3}partName")
            partNameElement.text = normalizeString(row.get("itemTitlePartName"))

            translatedTitleInfoElement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}titleInfo", {"type":"translated"})
            translatedTitleElement = etree.SubElement(translatedTitleInfoElement, "{http://www.loc.gov/mods/v3}title")
            translatedTitleElement.text = normalizeString(row.get("itemTitleTranslated", ''))

            alternateTitleInfoElement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}titleInfo", {"type":"alternative"})
            alternateTitleElement = etree.SubElement(alternateTitleInfoElement, "{http://www.loc.gov/mods/v3}title")
            alternateTitleElement.text = normalizeString(row.get("itemTitleAlternate", ''))

            pembroketitleinfo = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}titleInfo", {"type":"alternative", "displayLabel":"Pembroke title"})
            pembroketitle = etree.SubElement(pembroketitleinfo, "{http://www.loc.gov/mods/v3}title")
            pembroketitle.text = normalizeString(row.get("itemTitleAlternatePembroke", ''))
            # normalizeString(row.get("subTitle", ''))

            #names
            repeatingnamefield(modstop, row, 'namePersonCreatorLC', {"type":"personal", "authority":"naf"}, 'creator', False, 'v')
            repeatingnamefield(modstop, row, 'namePersonCreatorFAST', {"type":"personal", "authority":"fast"}, 'creator', False, 'v')
            repeatingnamefield(modstop, row, 'namePersonCreatorLocal', {"type":"personal", "authority":"local"}, 'creator', False, 'v')
            repeatingnamefield(modstop, row, 'namePersonOtherLC', {"type":"personal", "authority":"naf"}, '', False, 'v')
            repeatingnamefield(modstop, row, 'namePersonOtherFAST', {"type":"personal", "authority":"fast"}, '', False, 'v')
            repeatingnamefield(modstop, row, 'namePersonOtherLocal', {"type":"personal", "authority":"local"}, '', False, 'v')
            repeatingnamefield(modstop, row, 'nameCorpCreatorLC', {"type":"corporate", "authority":"naf"}, 'creator', False, 'v')
            repeatingnamefield(modstop, row, 'nameCorpCreatorFAST', {"type":"corporate", "authority":"fast"}, 'creator', False, 'v')
            repeatingnamefield(modstop, row, 'nameCorpCreatorLocal', {"type":"corporate", "authority":"local"}, 'creator', False, 'v')


            #typeOfResource

            typeOfResourceAttributes = {}

            if row.get("typeOfResourceManuscript", "") != "":
                typeOfResourceAttributes['manuscript'] = "yes"
            if row.get("typeOfResourceCollection", "") != "":
                typeOfResourceAttributes['collection'] = "yes"

            typeofresource = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}typeOfResource",typeOfResourceAttributes)
            typeofresource.text = normalizeString(row.get("typeOfResource", ''))

            #genre
            repeatingfield(modstop, row, "genreAAT", "{http://www.loc.gov/mods/v3}genre", {"authority":"aat"}, False, {})
            repeatingfield(modstop, row, "genreLCSH", "{http://www.loc.gov/mods/v3}genre", {"authority":"lcsh"}, False, {})
            repeatingfield(modstop, row, "genreLocal", "{http://www.loc.gov/mods/v3}genre", {"authority":"local"}, False, {})
            repeatingfield(modstop, row, "genreRBGENR", "{http://www.loc.gov/mods/v3}genre", {"authority":"rbgenr"}, False, {})

            #note
            notescopeelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}abstract", {"type":"general", "displayLabel":"Scope and Contents note"})
            notescopeelement.text = normalizeString(row.get("noteScope", ''))
            #normalizeString(row.get("noteScope", ''))

            noteAccessionelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}note", {"type":"acquisition", "displayLabel":"Immediate form of acquisition"})
            noteAccessionelement.text = normalizeString(row.get("noteAccession", ''))
            #normalizeString(row.get("noteAccession", ''))

            noteHistoricalelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}note", {"type":"biographical/historical", "displayLabel":"Biographical/historical note"})
            noteHistoricalelement.text = normalizeString(row.get("noteHistorical", ''))
            #normalizeString(row.get("noteHistorical", ''))

            noteGeneralelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}note", {"type":"general"})
            noteGeneralelement.text = normalizeString(row.get("noteGeneral", ''))
            #normalizeString(row.get("noteGeneral", ''))

            noteHistoricalClassYearelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}note", {"type":"biographical/historical", "displayLabel":"Class year"})
            noteHistoricalClassYearelement.text = normalizeString(row.get("noteHistoricalClassYear", '')).replace('.0','')
            # normalizeString(row.get("noteHistoricalClassYear", ''))

            noteVenueelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}note", {"type":"venue"})
            noteVenueelement.text = normalizeString(row.get("noteVenue", ''))
            print(normalizeString(row.get("noteVenue", '')))


            if includeDefaults == True:
                notePreferredCitation = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}note", {"type":"preferredcitation"})
                notePreferredCitationstring = title.text # normalizeString(row.get("title", '')).rstrip('.')
                if row.get("collection", '') != "":
                    notePreferredCitationstring = notePreferredCitationstring + ", " + normalizeString(row.get("collection", ''))
                if row.get("callNumber", '') != "":
                    notePreferredCitationstring = notePreferredCitationstring + ", " + normalizeString(row.get("callNumber", ''))
                
                notePreferredCitation.text = notePreferredCitationstring + ', Brown University Library'

            #originInfo
            originInfoelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}originInfo")

            publisherelement = etree.SubElement(originInfoelement, "{http://www.loc.gov/mods/v3}publisher")
            publisherelement.text = normalizeString(row.get("publisher", ''))

            dateQualifierAttribute = {}

            if row.get("dateQualifier", '') != "":
                dateQualifierAttribute = {"qualifier": row.get("dateQualifier", '')}

            if row.get("dateStart", '') == "":
                dateQualifierAttribute["keyDate"] = "yes"

            dateCreatedelement = etree.SubElement(originInfoelement, "{http://www.loc.gov/mods/v3}dateCreated", dateQualifierAttribute)
            dateCreatedelement.text = normalizeString(row.get("dateText", '')).replace('.0','')

            dateStartelementdict = {"encoding":"w3cdtf", "keyDate":"yes", "point":"start"}
            dateStartelementdict.update(dateQualifierAttribute)
            dateStartelement = etree.SubElement(originInfoelement, "{http://www.loc.gov/mods/v3}dateCreated", dateStartelementdict)
            dateStartelement.text = ' '.join(str(row.get("dateStart", '')).split()).replace('.0','')

            dateEndelementdict = {"encoding":"w3cdtf", "point":"end"}
            dateEndelementdict.update(dateQualifierAttribute)
            dateEndelement = etree.SubElement(originInfoelement, "{http://www.loc.gov/mods/v3}dateCreated", dateEndelementdict)
            dateEndelement.text = ' '.join(str(row.get("dateEnd", '')).split()).replace('.0','')

            placeelement = etree.SubElement(originInfoelement, "{http://www.loc.gov/mods/v3}place")
            placeTermelement = etree.SubElement(placeelement, "{http://www.loc.gov/mods/v3}placeTerm", {"type":"text"})
            placeTermelement.text = normalizeString(row.get("place", ''))

            #language
            languagesplitcharacter = getSplitCharacter(row.get("language", ''))
            for language in row.get("language", '').split(languagesplitcharacter):
                languageelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}language")
                languageTermelement = etree.SubElement(languageelement, "{http://www.loc.gov/mods/v3}languageTerm", {"type":"code", "authority":"iso639-2b"})

                if len(normalizeString(language)) > 3:
                    if normalizeString(language) in langcode:
                         languageTermelement.text = langcode[normalizeString(language)]
                    else:
                         languageTermelement.text = ' '.join(language.split())
                         langissue = True
                         print('langissue: ' + language)
                else:
                    languageTermelement.text = normalizeString(language)

            #physicalDescription
            physicalDescriptionelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}physicalDescription")

            extentQuantityelement = etree.SubElement(physicalDescriptionelement, "{http://www.loc.gov/mods/v3}extent")
            extentQuantityelement.text = normalizeString(row.get("extentQuantity", ''))

            extentSizeelement = etree.SubElement(physicalDescriptionelement, "{http://www.loc.gov/mods/v3}extent")
            extentSizeelement.text = normalizeString(row.get("extentSize", ''))

            extentSpeedelement = etree.SubElement(physicalDescriptionelement, "{http://www.loc.gov/mods/v3}extent")
            extentSpeedelement.text = normalizeString(row.get("extentSpeed", ''))

            digitalOriginelement = etree.SubElement(physicalDescriptionelement, "{http://www.loc.gov/mods/v3}digitalOrigin")
            digitalOriginelement.text = normalizeString(row.get("digitalOrigin", ''))

            formelement = etree.SubElement(physicalDescriptionelement, "{http://www.loc.gov/mods/v3}form")
            formelement.text = normalizeString(row.get("form", ''))

            #accessCondition
            useAndReproductionelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}accessCondition", {"type":"use and reproduction"})
            useAndReproductionelement.text = normalizeString(row.get("useAndReproduction", ''))

            rightsStatementelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}accessCondition", {"type":"rights statement","{http://www.w3.org/1999/xlink}href":normalizeString(row.get("rightsStatementURI", ''))})
            rightsStatementelement.text = normalizeString(row.get("rightsStatementText", ''))

            if includeDefaults == True:
                restrictionOnAccesselement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}accessCondition", {"type":"restriction on access"})
                restrictionOnAccesselement.text = "Collection is open for research."

            #subject
            repeatingnamefield(modstop, row, 'subjectNamesLC', {"type":"personal", "authority":"naf"}, '', True, 'v')
            repeatingnamefield(modstop, row, 'subjectNamesFAST', {"type":"personal", "authority":"fast"}, '', True, 'v')
            repeatingnamefield(modstop, row, 'subjectNamesLocal', {"type":"personal", "authority":"local"}, '', True, 'v')
            repeatingnamefield(modstop, row, 'subjectCorpLC', {"type":"corporate", "authority":"naf"}, '', True, 'v')
            repeatingnamefield(modstop, row, 'subjectCorpFAST', {"type":"corporate", "authority":"fast"}, '', True, 'v')
            repeatingnamefield(modstop, row, 'subjectCorpLocal', {"type":"corporate", "authority":"local"}, '', True, 'v')

            repeatingfield(modstop, row, "subjectTopicsLC", "{http://www.loc.gov/mods/v3}topic", {}, True, {"authority":"lcsh"})
            repeatingfield(modstop, row, "subjectTopicsLocal", "{http://www.loc.gov/mods/v3}topic", {}, True, {"authority":"local"})
            repeatingfield(modstop, row, "subjectTopicsLocalFreedomNow", "{http://www.loc.gov/mods/v3}topic", {}, True, {"authority":"local", "displayLabel":"Freedom Now! keyword"})
            repeatingfield(modstop, row, "subjectTopicsFAST", "{http://www.loc.gov/mods/v3}topic", {}, True, {"authority":"fast"})
            repeatingfield(modstop, row, "subjectGeoLC", "{http://www.loc.gov/mods/v3}geographic", {}, True, {"authority":"lcsh"})
            repeatingfield(modstop, row, "subjectGeoFAST", "{http://www.loc.gov/mods/v3}geographic", {}, True, {"authority":"fast"})
            repeatingfield(modstop, row, "subjectTemporalLC", "{http://www.loc.gov/mods/v3}temporal", {}, True, {"authority":"lcsh"})
            repeatingfield(modstop, row, "subjectTemporalFAST", "{http://www.loc.gov/mods/v3}temporal", {}, True, {"authority":"fast"})
            repeatingfield(modstop, row, "subjectTemporalLocal", "{http://www.loc.gov/mods/v3}temporal", {}, True, {"authority":"local"})

            repeatingTitleSubjectField(modstop, row, "subjectTitleLocal", {"authority":"local"})
            repeatingTitleSubjectField(modstop, row, "subjectTitleLC", {"authority":"lcsh"})
            repeatingTitleSubjectField(modstop, row, "subjectTitleFAST", {"authority":"fast"})

            #cartographic
            subjectelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}subject")
            cartographicselement = etree.SubElement(subjectelement, "{http://www.loc.gov/mods/v3}cartographics")
            cartographicExtensionelement = etree.SubElement(cartographicselement, "{http://www.loc.gov/mods/v3}cartographicExtension")

            coordinateselement = etree.SubElement(cartographicExtensionelement, "{http://www.loc.gov/mods/v3}coordinates")
            coordinateselement.text = normalizeString(row.get("coordinates", ''))

            scaleelement = etree.SubElement(cartographicExtensionelement, "{http://www.loc.gov/mods/v3}scale")
            scaleelement.text = normalizeString(row.get("scale", ''))

            projectionelement = etree.SubElement(cartographicExtensionelement, "{http://www.loc.gov/mods/v3}projection")
            projectionelement.text = normalizeString(row.get("projection", ''))

            #collection
            relatedItemelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}relatedItem", {"type":"host"})

            hosttitleInfoelement = etree.SubElement(relatedItemelement, "{http://www.loc.gov/mods/v3}titleInfo")
            hosttitleelement = etree.SubElement(hosttitleInfoelement, "{http://www.loc.gov/mods/v3}title")
            hosttitleelement.text = normalizeString(row.get("collection")) # normalizeString(row.get("collection", ''))

            hostoriginInfoelement = etree.SubElement(relatedItemelement, "{http://www.loc.gov/mods/v3}originInfo")
            hostdateCreatedelement = etree.SubElement(hostoriginInfoelement, "{http://www.loc.gov/mods/v3}dateCreated")
            hostdateCreatedelement.text = normalizeString(row.get("dateTextParent", '')).replace('.0','')

            hostidentifierelement = etree.SubElement(relatedItemelement, "{http://www.loc.gov/mods/v3}identifier", {"type":"local"})
            hostidentifierelement.text = normalizeString(row.get("callNumber", ''))

            hostlocationelement = etree.SubElement(relatedItemelement, "{http://www.loc.gov/mods/v3}location")

            hostphysicalLocationelement = etree.SubElement(hostlocationelement, "{http://www.loc.gov/mods/v3}physicalLocation")
            hostphysicalLocationelement.text = normalizeString(row.get("repository", ''))

            hosturlelement = etree.SubElement(hostlocationelement, "{http://www.loc.gov/mods/v3}url")
            hosturlelement.text = normalizeString(row.get("findingAid", ''))

            hostholdingSimpleelement = etree.SubElement(hostlocationelement, "{http://www.loc.gov/mods/v3}holdingSimple")
            hostcopyInformationelement = etree.SubElement(hostholdingSimpleelement, "{http://www.loc.gov/mods/v3}copyInformation")
            hostshelfLocatorelement = etree.SubElement(hostcopyInformationelement, "{http://www.loc.gov/mods/v3}shelfLocator")

            shelfLocatorstring = ""

            if row.get("shelfLocator1", '') != "":
                shelfLocatorstring = ' '.join(row.get("shelfLocator1",'').split()) + ' ' + ' '.join(str(row.get("shelfLocator1ID",'')).split()).replace('.0','')
            if row.get("shelfLocator2", '') != "":
                shelfLocatorstring = shelfLocatorstring + ', ' + ' '.join(row.get("shelfLocator2",'').split()) + ' ' + ' '.join(str(row.get("shelfLocator2ID",'')).split()).replace('.0','')
            if row.get("shelfLocator3", '') != "":
                shelfLocatorstring = shelfLocatorstring + ', ' + ' '.join(row.get("shelfLocator3",'').split()) + ' ' + ' '.join(str(row.get("shelfLocator3ID",'')).split()).replace('.0','')

            shelfLocatorstring = shelfLocatorstring.lstrip(', ')
            hostshelfLocatorelement.text = ' '.join(shelfLocatorstring.split())

            #Additional location fields
            locationElement = etree.SubElement(modstop,"{http://www.loc.gov/mods/v3}location")
            repeatingfield(locationElement, row, "physicalLocationLC", "{http://www.loc.gov/mods/v3}physicalLocation", {"authority":"naf"}, False, {} )
            holdingSimpleElement = etree.SubElement(locationElement, "{http://www.loc.gov/mods/v3}holdingSimple")
            copyInformationElement = etree.SubElement(holdingSimpleElement, "{http://www.loc.gov/mods/v3}copyInformation")

            if row.get("physicalLocationLC", "") != "":
                shelfLocator1Element = etree.SubElement(copyInformationElement, "{http://www.loc.gov/mods/v3}note", {"type": row.get("shelfLocator1","").lower() + " name"})
                shelfLocator1Element.text = normalizeString(row.get("shelfLocator1", "") + " " + row.get("shelfLocator1ID","").replace('.0',''))
                shelfLocator2Element = etree.SubElement(copyInformationElement, "{http://www.loc.gov/mods/v3}note", {"type": row.get("shelfLocator2","").lower() + " title"})
                shelfLocator2Element.text = normalizeString(row.get("shelfLocator2", "") + " " + row.get("shelfLocator2ID","").replace('.0',''))
                shelfLocator3Element = etree.SubElement(copyInformationElement, "{http://www.loc.gov/mods/v3}note", {"type": row.get("shelfLocator3","").lower() + " name"})
                shelfLocator3Element.text = normalizeString(row.get("shelfLocator3", "") + " " + row.get("shelfLocator3ID","").replace('.0',''))

            #If identifierBDR has a bdr number in it:
            if row.get("identifierBDR", '').startswith('bdr'):
                #identifiers
                BDRPIDIdentifierelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}identifier", {"type":"local","displayLabel":"BDR_PID"})
                BDRPIDIdentifierelement.text = 'bdr:'+ normalizeString(row.get("identifierBDR", '')).lstrip('bdr').replace(':','')

                MODSIDIdentifierelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}identifier", {"type":"local","displayLabel":"MODS_ID"})
                MODSIDIdentifierelement.text = 'bdr'+ normalizeString(row.get("identifierBDR", '')).lstrip('bdr').replace(':','')

            #lastnote
            if includeDefaults == True:
                digitalObjectMadeelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}note", {"displayLabel":"Digital object made available by"})
                digitalObjectMadeelement.text = "Brown University Library, John Hay Library, University Archives and Manuscripts, Box A, Brown University, Providence, RI, 02912, U.S.A., (http://library.brown.edu/)"

            ##Alphabetize certain kinds of elements
            #mods:name in [namePersonCreatorLC, namePersonCreatorLocal] [nameCorpCreatorLC, nameCorpCreatorLocal] [namePersonOtherLC, namePersonOtherLocal] order
            firstnameelementindex = 0

            firstnameelement = modstop.find('{http://www.loc.gov/mods/v3}name')
            try:
                firstnameelementindex = modstop.getchildren().index(firstnameelement)
            except:
                print("No names to alphabetize.")

            #print(str(firstnameelementindex))
            ns = {'mods': 'http://www.loc.gov/mods/v3'}
            allnameelements = modstop.findall("{http://www.loc.gov/mods/v3}name", ns)

            if len(allnameelements) > 0:

                #Remove all discovered elements from the document
                for element in allnameelements:
                    modstop.remove(element)

                #Sort the captured name elements alphabetically
                allnameelements = sorted(allnameelements, key=lambda ch: ch.xpath("mods:namePart/text()", namespaces={'mods': 'http://www.loc.gov/mods/v3'}))

                #Reorganize by personal and corp
                personnamecreatorelements = []
                corpnamecreatorelements = []

                personnameotherelements = []
                corpnameotherelements = []

                for element in allnameelements:
                    #print(etree.tostring(element))
                    #print(element.attrib)
                    roletermtext = ''
                    creatorrole = False

                    try:
                        roletermtext = element.find("{http://www.loc.gov/mods/v3}role/{http://www.loc.gov/mods/v3}roleTerm", ns).text
                    except:
                        print("No role text.")

                    if roletermtext == 'creator':
                        creatorrole = True

                    if element.attrib.get('type') == 'personal' and creatorrole == True:
                        personnamecreatorelements.append(element)
                        #print("personal creator")
                    if element.attrib.get('type') == 'corporate' and creatorrole == True:
                        corpnamecreatorelements.append(element)
                        #print("corporate creator")
                    if element.attrib.get('type') == 'personal' and creatorrole == False:
                        personnameotherelements.append(element)
                        #print("personal other")
                    if element.attrib.get('type') == 'corporate' and creatorrole == False:
                        corpnameotherelements.append(element)
                        #print("corporate other")

                #Reappend
                for element in personnamecreatorelements:
                    #print(etree.tostring(element).decode("utf-8"))
                    modstop.insert(firstnameelementindex, element)
                    firstnameelementindex += 1
                for element in corpnamecreatorelements:
                    #print(etree.tostring(element).decode("utf-8"))
                    modstop.insert(firstnameelementindex, element)
                    firstnameelementindex += 1
                for element in personnameotherelements:
                    #print(etree.tostring(element).decode("utf-8"))
                    modstop.insert(firstnameelementindex, element)
                    firstnameelementindex += 1
                for element in corpnameotherelements:
                    #print(etree.tostring(element).decode("utf-8"))
                    modstop.insert(firstnameelementindex, element)
                    firstnameelementindex += 1

            #mods:subject
            firstsubjectelementindex = 0

            firstsubjectelement = modstop.find('{http://www.loc.gov/mods/v3}subject')

            try:
                firstsubjectelementindex = modstop.getchildren().index(firstsubjectelement)
            except:
                print("No subjects to alphabetize.")

            #print(str(firstsubjectelementindex))

            allsubjectnameelemements = modstop.findall("{http://www.loc.gov/mods/v3}subject[{http://www.loc.gov/mods/v3}name]", ns)

            allsubjectnameelemements = sorted(allsubjectnameelemements, key=lambda ch: ch.xpath("mods:name/mods:namePart/text()", namespaces={'mods': 'http://www.loc.gov/mods/v3'}))

            if len(allsubjectnameelemements) > 0:
                #Remove all discovered elements from the document
                for element in allsubjectnameelemements:
                    modstop.remove(element)

                personnamesubjects = []
                corpnamesubjects = []

                for element in allsubjectnameelemements:
                    #print(etree.tostring(element))

                    if element.getchildren()[0].attrib.get('type') == 'personal':
                        personnamesubjects.append(element)
                        #print("personal subject")
                    if element.getchildren()[0].attrib.get('type') == 'corporate':
                        corpnamesubjects.append(element)
                        #print("corporate subject")

                #Reappend
                for element in personnamesubjects:
                    #print(etree.tostring(element).decode("utf-8"))
                    modstop.insert(firstsubjectelementindex, element)
                    firstsubjectelementindex += 1
                for element in corpnamesubjects:
                    #print(etree.tostring(element).decode("utf-8"))
                    modstop.insert(firstsubjectelementindex, element)
                    firstsubjectelementindex += 1

            #mods:subject topics
            allsubjecttopicelemements = modstop.findall("{http://www.loc.gov/mods/v3}subject[{http://www.loc.gov/mods/v3}topic]", ns)

            allsubjecttopicelemements = sorted(allsubjecttopicelemements, key=lambda ch: ch.xpath("mods:topic/text()", namespaces={'mods': 'http://www.loc.gov/mods/v3'}))

            for element in allsubjecttopicelemements:
                modstop.remove(element)

            for element in allsubjecttopicelemements:
                #print(etree.tostring(element))
                modstop.insert(firstsubjectelementindex, element)
                firstsubjectelementindex += 1

            #mods:subject title
            allsubjecttitleelemements = modstop.findall("{http://www.loc.gov/mods/v3}subject[{http://www.loc.gov/mods/v3}titleInfo]", ns)

            allsubjecttitleelemements = sorted(allsubjecttitleelemements, key=lambda ch: ch.xpath("mods:titleInfo/mods:title/text()", namespaces={'mods': 'http://www.loc.gov/mods/v3'}))

            for element in allsubjecttitleelemements:
                modstop.remove(element)

            for element in allsubjecttitleelemements:
                #print(etree.tostring(element))
                modstop.insert(firstsubjectelementindex, element)
                firstsubjectelementindex += 1

            #mods:subject geo
            allsubjectgeoelemements = modstop.findall("{http://www.loc.gov/mods/v3}subject[{http://www.loc.gov/mods/v3}geographic]", ns)

            allsubjectgeoelemements = sorted(allsubjectgeoelemements, key=lambda ch: ch.xpath("mods:geographic/text()", namespaces={'mods': 'http://www.loc.gov/mods/v3'}))

            for element in allsubjectgeoelemements:
                modstop.remove(element)

            for element in allsubjectgeoelemements:
                #print(etree.tostring(element))
                modstop.insert(firstsubjectelementindex, element)
                firstsubjectelementindex += 1

            #mods:subject temporal
            allsubjecttemporalelemements = modstop.findall("{http://www.loc.gov/mods/v3}subject[{http://www.loc.gov/mods/v3}temporal]", ns)

            allsubjecttemporalelemements = sorted(allsubjecttemporalelemements, key=lambda ch: ch.xpath("mods:temporal/text()", namespaces={'mods': 'http://www.loc.gov/mods/v3'}))

            for element in allsubjecttemporalelemements:
                modstop.remove(element)

            for element in allsubjecttemporalelemements:
                #print(etree.tostring(element))
                modstop.insert(firstsubjectelementindex, element)
                firstsubjectelementindex += 1


            # start cleanup
            # remove any element tails
            for element in modstop.iter():
                element.tail = None

            # remove any line breaks or tabs in element text
                if element.text:
                    if '\n' in element.text:
                        element.text = element.text.replace('\n', '')
                    if '\t' in element.text:
                        element.text = element.text.replace('\t', '')

            # remove any remaining whitespace
            parser = etree.XMLParser(remove_blank_text=True, remove_comments=True, recover=True)
            treestring = etree.tostring(modstop)
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

            allrecords = allrecords + "\n" + etree.tostring(clean, pretty_print = True, encoding="unicode")
            allrecords.lstrip("\n")

            filename = row.get("identifierBDR", '')

            if filename == "":
                filename = "default" + str(rowindex)

            with open(CACHEDIR + id + "/" + filename.replace(':','') + ".mods", 'w+') as f:
                f.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
                f.write(etree.tostring(clean, pretty_print = True, encoding="unicode"))
                #zipObj.write(f.read())
                print( "Writing " + filename + ".mods" + "\n", file=sys.stderr)


            zipObj.write(CACHEDIR + id + "/" + filename.replace(':','') + ".mods", filename.replace(':','') + ".mods")
            rowindex = rowindex + 1
            amountofrecords = amountofrecords + 1

        returndict = {}

        returndict["filename"] = chosensheet + ".zip"
        returndict["error"] = False
        returndict["allrecords"] = allrecords

        #Return the zipped files
        #zipObj.close()

        zipObj.close()

        with open(CACHEDIR + id + "/" + chosensheet + '.zip', mode='rb') as zipdata:
            return zipdata.read(), returndict

        print('\n\n', file=sys.stderr)


        if langissue == True:
            print( "*Language Field Error*\nThere were one or more issues with language fields in your spreadsheet. Please check your spelling in all language fields. You may also manually correct your XML file, consult the SupportedLanguages.xlsx file in the data folder for supported languages, and/or adjust the SupportedLanguages.xlsx spreadsheet to suit your project. \n", file=sys.stderr)

            print('\n\n', file=sys.stderr)

        #errorfile.close()
        if amountofrecords > 1:
            print("***Operation Complete*** \n" + str(amountofrecords) + " MODS records were written." + ".\n", file=sys.stderr)

            print('\n\n', file=sys.stderr)

            print( u"\u2606 \u2606 \u2606", file=sys.stderr)
            print('\n\n', file=sys.stderr)

        else:
            print("***Operation Complete***\n" + str(amountofrecords) + " MODS record was written." + ".\n", file=sys.stderr)

            print('\n\n', file=sys.stderr)

            print( u"   \u2606 \u2606 \u2606", file=sys.stderr)
            print('\n\n\n', file=sys.stderr)

        #MyGUI.recordsfolder = output_path
        #print(MyGUI.recordsfolder)
        #MyGUI.outputfolderbutton['state'] = 'normal'
    #except Exception, e:
    #    print("Process failed with the following error: " + e + ".\n\n", file=sys.stderr)
    #


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

# attrQname = etree.QName("http://www.w3.org/2001/XMLSchema-instance", "schemaLocation")
# nsMap = {"mods" : "http://www.loc.gov/mods/v3", "xsi" : "http://www.w3.org/2001/XMLSchema-instance", "xlink" : "http://www.w3.org/1999/xlink"}
# elementNameSpace = "{http://www.loc.gov/mods/v3}"

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

def createParentElement(key):
    keyQName = key.get("attrqname", {})
    keyAttrQname = etree.QName(keyQName.get("uri",""),keyQName.get("tag",""))

    keyNsMap = key.get("nsmap", {})

    keyElementNameSpace = key.get("elementnamespace", "")
    keyParentTag = key.get("parenttag", "")

    return lxml.etree.Element(keyElementNameSpace + keyParentTag, {keyAttrQname: "http://www.loc.gov/mods/v3 http://www.loc.gov/mods/v3/mods-3-7.xsd"}, nsmap=keyNsMap)

def createSubElement(parentElement, elementName, elementNameSpace, elementAttrs, elementText):
    print(elementNameSpace)
    print(elementName)
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
        print(etree.tostring(element))
        parentElement.insert(firstElementIndex + index, element)

    print("....")

def convertExcelRowToEtree(row, globalConditions):
    keyElementNameSpace = key.get("elementnamespace", "")
    keyParentTag = key.get("parenttag", "")

    if shouldSkipRow(row, keySkips):
        return
    
    parentElement = createParentElement(key)

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

convertExcelRowToEtree(row, globalConditions)

# def getAllColumnHeaders(keyFields):

#     parentElement = parentElement = createParentElement(key)

#     for keyField in keyFields:
#         keyFieldType = keyField.get('type', "")

#         if keyFieldType == 'element':
#             textFields = keyField.get("text", [])
#             for textField in textFields:
#                 textFieldValues = textField.get("values", [])
#                 for textFieldValue in textFieldValues:
#                     if textFieldValue.get("type") == "col":
#                         print(textFieldValue.get("header"))
#                         header = textFieldValue.get("header", "")
                        
#                         element = processElementTypeField(keyField, keyNameSpace, {header: header}, parentElement)
#                         print(etree.tostring(element, pretty_print=True).decode("utf-8"))
            
#         if keyFieldType == "repeating":
#             pass
#             # elements = processRepeatingTypeField(keyField, keyAuthorities, keyNameSpace, row, parentElement) 

