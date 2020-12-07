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

def repeatingfield(parentelement, refdict, originalfieldname, modsfieldname, modsattributes, subject, subjectattributes):
    splitcharacter = ""
    originalparentelement = parentelement

    if ";" in refdict.get(originalfieldname, ''):
        splitcharacter = ";"
    else:
        splitcharacter = "|"

    for namesindex, addedentry in enumerate(refdict.get(originalfieldname, '').split(splitcharacter)):

        customSubjectAttributes = subjectattributes.copy()
        customMODSattributes = modsattributes.copy()

        #Extract URI

        uri = re.findall("(?P<url>https?://[^\s]+)", addedentry)

        #If there's a URI
        if len(uri) > 0:
            #Remove it from the addedentry
            addedentry = addedentry.replace(uri[0],"")
            #Add it as a valueURI attribute
            if subject:
                customSubjectAttributes["valueURI"] = xmltext(uri[0])
            else:
                customMODSattributes["valueURI"] = xmltext(uri[0])

        #Add authorityURI attribute

        if subject:

            authorityType = customSubjectAttributes.get("authority", "")

            if authorityURIs.get(authorityType):
                customSubjectAttributes["authorityURI"] = authorityURIs.get(authorityType)

        else:
            authorityType = customMODSattributes.get("authority", "")

            if authorityURIs.get(authorityType):
                customMODSattributes["authorityURI"] = authorityURIs.get(authorityType)

        #Create field

        if subject == True:
            subjectelement = etree.SubElement(parentelement, "{http://www.loc.gov/mods/v3}subject", customSubjectAttributes)
            parentelement = subjectelement

        namecontrolaccesselement = etree.SubElement(parentelement, modsfieldname, customMODSattributes)
        namecontrolaccesselement.text = ' '.join(addedentry.replace("|d", "").replace("|e", "").split())

        parentelement = originalparentelement

def repeatingTitleSubjectField(modstop, row, originalfieldname, attributes):

    for title in row.get(originalfieldname, '').split(';'):

        customAttributes = attributes.copy()

        #Extract URI

        uri = re.findall("(?P<url>https?://[^\s]+)", title)

        #If there's a URI
        if len(uri) > 0:
            #Remove it from the addedentry
            title = title.replace(uri[0],"")
            #Add it as a valueURI attribute
            customAttributes["valueURI"] = xmltext(uri[0])

            print("custom attributes")
            print(customAttributes)
            print("reg attributes")
            print(attributes)

        #Add authorityURI attribute

        authorityType = customAttributes.get("authority", "")

        if authorityURIs.get(authorityType):
            customAttributes["authorityURI"] = authorityURIs.get(authorityType)

        #Create element
        subjecttitleparentelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}subject", customAttributes)
        subjecttitleinfoelement = etree.SubElement(subjecttitleparentelement, "{http://www.loc.gov/mods/v3}titleInfo")
        subjecttitleelement = etree.SubElement(subjecttitleinfoelement, "{http://www.loc.gov/mods/v3}title")
        subjecttitleelement.text = xmltext(title)


def repeatingnamefield(parentelement, refdict, originalfieldname, topmodsattributes, predefinedrole, subject, splitcharacter):
    originalparentelement = parentelement
    if splitcharacter == 'v':
        if ";" in refdict.get(originalfieldname, ''):
            splitcharacter = ";"
        else:
            splitcharacter = "|"

    for nameindex, name in enumerate(refdict.get(originalfieldname, '').split(splitcharacter)):
        nametext = ""
        datetext = ""
        roletext = predefinedrole

        customAttributes = topmodsattributes.copy()

        #Extract URI

        uri = re.findall("(?P<url>https?://[^\s]+)", name)

        #If there's a URI
        if len(uri) > 0:
            #Remove it from the addedentry
            name = name.replace(uri[0],"")
            #Add it as a valueURI attribute
            customAttributes["valueURI"] = xmltext(uri[0])

        #Add authorityURI attribute

        authorityType = customAttributes.get("authority", "")

        if authorityURIs.get(authorityType):
            customAttributes["authorityURI"] = authorityURIs.get(authorityType)


        for textindex, text in enumerate(name.split(',')):
            textrevised = ' '.join(text.split()).replace('|d', '').replace('|e','')

            if textrevised == '':
                continue

            max_index = len(xmltext(name).split(','))-1

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
        namepart = etree.SubElement(nameelement, "{http://www.loc.gov/mods/v3}namePart")
        namepart.text = xmltext(nametext).rstrip(',')
        namedatepart = etree.SubElement(nameelement, "{http://www.loc.gov/mods/v3}namePart", {"type":"date"})
        namedatepart.text = xmltext(datetext).lstrip(',').rstrip(',').replace('|d','')
        modsrole = etree.SubElement(nameelement, "{http://www.loc.gov/mods/v3}role")
        modsroleterm = etree.SubElement(modsrole, "{http://www.loc.gov/mods/v3}roleTerm", {"type":"text", "authority":"marcrelator"})
        modsroleterm.text = xmltext(roletext).lstrip(',').rstrip(',').replace('|e','')

        parentelement = originalparentelement

def xmltext(text):
    if text != None:
        text = text.replace('\n', ' ').replace('\r', ' ')
        text = text.replace('<title>', '').replace('</title>', '')
        text = text.replace('<geogname>', '- ').replace('</geogname>', '')
        return(' '.join(str(text).split()))
    else:
        return text

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
                    newvalue = str(sheet.cell_value(row,column)).replace('|d', '').replace('|e', '').replace('|',';')
                else:
                    newvalue = sheet.cell_value(row,column).replace('|d', '').replace('|e', '').replace('|',';')

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


CACHEDIR = os.path.join(os.getcwd(), "cache") + "/"
#CACHEDIR = os.getcwd() + "/"
HOMEDIR = os.getcwd() + "/"
#HOMEDIR = os.getcwd() + "/"

#print("._. MODS Maker ._.")

def processExceltoMODS(chosenfile, chosensheet, id):
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
        #output_path = os.getcwd()

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
                title.text = xmltext(row.get("fileTitle", ''))
            #elif row.get("title", '') != "":
            #    title.text = xmltext(row.get("title", ''))
            else:
                title.text = xmltext(row.get("itemTitle", ''))
            #title.text = xmltext(row.get("title"]) #' '.join(row["title", '').split())
            subtitle = etree.SubElement(titleinfo, "{http://www.loc.gov/mods/v3}subTitle")
            subtitle.text = xmltext(row.get("subTitle", ''))

            pembroketitleinfo = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}titleInfo", {"type":"alternative", "displayLabel":"Pembroke title"})
            pembroketitle = etree.SubElement(pembroketitleinfo, "{http://www.loc.gov/mods/v3}title")
            pembroketitle.text = xmltext(row.get("itemTitleAlternatePembroke", ''))
            # xmltext(row.get("subTitle", ''))

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
            typeofresource = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}typeOfResource")
            typeofresource.text = xmltext(row.get("typeOfResource", ''))

            #genre
            repeatingfield(modstop, row, "genreAAT", "{http://www.loc.gov/mods/v3}genre", {"authority":"aat"}, False, {})
            repeatingfield(modstop, row, "genreLCSH", "{http://www.loc.gov/mods/v3}genre", {"authority":"lcsh"}, False, {})
            repeatingfield(modstop, row, "genreLocal", "{http://www.loc.gov/mods/v3}genre", {"authority":"local"}, False, {})
            repeatingfield(modstop, row, "genreRBGENR", "{http://www.loc.gov/mods/v3}genre", {"authority":"rbgenr"}, False, {})

            #note
            notescopeelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}abstract", {"type":"general", "displayLabel":"Scope and Contents note"})
            notescopeelement.text = xmltext(row.get("noteScope", ''))
            #xmltext(row.get("noteScope", ''))

            noteAccessionelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}note", {"type":"acquisition", "displayLabel":"Immediate form of acquisition"})
            noteAccessionelement.text = xmltext(row.get("noteAccession", ''))
            #xmltext(row.get("noteAccession", ''))

            noteHistoricalelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}note", {"type":"biographical/historical", "displayLabel":"Biographical/historical note"})
            noteHistoricalelement.text = xmltext(row.get("noteHistorical", ''))
            #xmltext(row.get("noteHistorical", ''))

            noteGeneralelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}note", {"type":"general"})
            noteGeneralelement.text = xmltext(row.get("noteGeneral", ''))
            #xmltext(row.get("noteGeneral", ''))

            noteHistoricalClassYearelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}note", {"type":"biographical/historical", "displayLabel":"Class year"})
            noteHistoricalClassYearelement.text = xmltext(row.get("noteHistoricalClassYear", '')).replace('.0','')
            # xmltext(row.get("noteHistoricalClassYear", ''))

            noteVenueelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}note", {"type":"venue"})
            noteVenueelement.text = xmltext(row.get("noteVenue", ''))
            print(xmltext(row.get("noteVenue", '')))


            if row.get("noPreferredCitation", "") == "":
                notePreferredCitation = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}note", {"type":"preferredcitation"})
                notePreferredCitationstring = title.text # xmltext(row.get("title", '')).rstrip('.')
                if row.get("collection", '') != "":
                    notePreferredCitationstring = notePreferredCitationstring + ", " + xmltext(row.get("collection", ''))
                if row.get("callNumber", '') != "":
                    notePreferredCitationstring = notePreferredCitationstring + ", " + xmltext(row.get("callNumber", ''))
                
                notePreferredCitation.text = notePreferredCitationstring + ', Brown University Library'

            #originInfo
            originInfoelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}originInfo")

            publisherelement = etree.SubElement(originInfoelement, "{http://www.loc.gov/mods/v3}publisher")
            publisherelement.text = xmltext(row.get("publisher", ''))

            dateQualifierAttribute = {}

            if row.get("dateQualifier", '') != "":
                dateQualifierAttribute = {"qualifier": row.get("dateQualifier", '')}

            if row.get("dateStart", '') == "":
                dateQualifierAttribute["keyDate"] = "yes"

            dateCreatedelement = etree.SubElement(originInfoelement, "{http://www.loc.gov/mods/v3}dateCreated", dateQualifierAttribute)
            dateCreatedelement.text = xmltext(row.get("dateText", '')).replace('.0','')

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
            placeTermelement.text = xmltext(row.get("place", ''))

            #language
            languagesplitcharacter = getSplitCharacter(row.get("language", ''))
            for language in row.get("language", '').split(languagesplitcharacter):
                languageelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}language")
                languageTermelement = etree.SubElement(languageelement, "{http://www.loc.gov/mods/v3}languageTerm", {"type":"code", "authority":"iso639-2b"})

                if len(xmltext(language)) > 3:
                    if xmltext(language) in langcode:
                         languageTermelement.text = langcode[xmltext(language)]
                    else:
                         languageTermelement.text = ' '.join(language.split())
                         langissue = True
                         print('langissue: ' + language)
                else:
                    languageTermelement.text = xmltext(language)

            #physicalDescription
            physicalDescriptionelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}physicalDescription")

            extentQuantityelement = etree.SubElement(physicalDescriptionelement, "{http://www.loc.gov/mods/v3}extent")
            extentQuantityelement.text = xmltext(row.get("extentQuantity", ''))

            extentSizeelement = etree.SubElement(physicalDescriptionelement, "{http://www.loc.gov/mods/v3}extent")
            extentSizeelement.text = xmltext(row.get("extentSize", ''))

            extentSpeedelement = etree.SubElement(physicalDescriptionelement, "{http://www.loc.gov/mods/v3}extent")
            extentSpeedelement.text = xmltext(row.get("extentSpeed", ''))

            digitalOriginelement = etree.SubElement(physicalDescriptionelement, "{http://www.loc.gov/mods/v3}digitalOrigin")
            digitalOriginelement.text = xmltext(row.get("digitalOrigin", ''))

            formelement = etree.SubElement(physicalDescriptionelement, "{http://www.loc.gov/mods/v3}form")
            formelement.text = xmltext(row.get("form", ''))

            #accessCondition
            useAndReproductionelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}accessCondition", {"type":"use and reproduction"})
            useAndReproductionelement.text = xmltext(row.get("useAndReproduction", ''))

            rightsStatementelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}accessCondition", {"type":"rights statement","{http://www.w3.org/1999/xlink}href":xmltext(row.get("rightsStatementURI", ''))})
            rightsStatementelement.text = xmltext(row.get("rightsStatementText", ''))

            if xmltext(row.get('notOpenForResearch','')) == '':
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
            coordinateselement.text = xmltext(row.get("coordinates", ''))

            scaleelement = etree.SubElement(cartographicExtensionelement, "{http://www.loc.gov/mods/v3}scale")
            scaleelement.text = xmltext(row.get("scale", ''))

            projectionelement = etree.SubElement(cartographicExtensionelement, "{http://www.loc.gov/mods/v3}projection")
            projectionelement.text = xmltext(row.get("projection", ''))

            #collection
            relatedItemelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}relatedItem", {"type":"host"})

            hosttitleInfoelement = etree.SubElement(relatedItemelement, "{http://www.loc.gov/mods/v3}titleInfo")
            hosttitleelement = etree.SubElement(hosttitleInfoelement, "{http://www.loc.gov/mods/v3}title")
            hosttitleelement.text = xmltext(row.get("collection")) # xmltext(row.get("collection", ''))

            hostoriginInfoelement = etree.SubElement(relatedItemelement, "{http://www.loc.gov/mods/v3}originInfo")
            hostdateCreatedelement = etree.SubElement(hostoriginInfoelement, "{http://www.loc.gov/mods/v3}dateCreated")
            hostdateCreatedelement.text = xmltext(row.get("dateTextParent", '')).replace('.0','')

            hostidentifierelement = etree.SubElement(relatedItemelement, "{http://www.loc.gov/mods/v3}identifier", {"type":"local"})
            hostidentifierelement.text = xmltext(row.get("callNumber", ''))

            hostlocationelement = etree.SubElement(relatedItemelement, "{http://www.loc.gov/mods/v3}location")

            hostphysicalLocationelement = etree.SubElement(hostlocationelement, "{http://www.loc.gov/mods/v3}physicalLocation")
            hostphysicalLocationelement.text = xmltext(row.get("repository", ''))

            hosturlelement = etree.SubElement(hostlocationelement, "{http://www.loc.gov/mods/v3}url")
            hosturlelement.text = xmltext(row.get("findingAid", ''))

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
                shelfLocator1Element.text = xmltext(row.get("shelfLocator1", "") + " " + row.get("shelfLocator1ID","").replace('.0',''))
                shelfLocator2Element = etree.SubElement(copyInformationElement, "{http://www.loc.gov/mods/v3}note", {"type": row.get("shelfLocator2","").lower() + " title"})
                shelfLocator2Element.text = xmltext(row.get("shelfLocator2", "") + " " + row.get("shelfLocator2ID","").replace('.0',''))
                shelfLocator3Element = etree.SubElement(copyInformationElement, "{http://www.loc.gov/mods/v3}note", {"type": row.get("shelfLocator3","").lower() + " name"})
                shelfLocator3Element.text = xmltext(row.get("shelfLocator3", "") + " " + row.get("shelfLocator3ID","").replace('.0',''))

            #If identifierBDR has a bdr number in it:
            if row.get("identifierBDR", '').startswith('bdr'):
                #identifiers
                BDRPIDIdentifierelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}identifier", {"type":"local","displayLabel":"BDR_PID"})
                BDRPIDIdentifierelement.text = 'bdr:'+ xmltext(row.get("identifierBDR", '')).lstrip('bdr').replace(':','')

                MODSIDIdentifierelement = etree.SubElement(modstop, "{http://www.loc.gov/mods/v3}identifier", {"type":"local","displayLabel":"MODS_ID"})
                MODSIDIdentifierelement.text = 'bdr'+ xmltext(row.get("identifierBDR", '')).lstrip('bdr').replace(':','')

            #lastnote
            if row.get("noDigitalObjectMade", "") == "":
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
                if recursively_empty(elem):
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
