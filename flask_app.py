from flask import jsonify, Flask, make_response, request, render_template, redirect, url_for
import flask
from legacy.EADMaker import processExceltoEAD
from legacy.EADMaker import getSheetNames
from lib import fileSupport
from lib import profileInterpreter
from lib import profileValidation
import sys
import uuid
import os
import json
from glob import glob

app = Flask(__name__)
app.config["DEBUG"] = True

def getBooleanFromJson(value, default=True):
    if value is None:
        return default

    if isinstance(value, bool):
        return value

    return str(value).lower() in ['1', 'true', 'yes', 'on']

def getEnforceValidationsFromForm(form):
    values = form.getlist('enforce_validations')

    if not values:
        return True

    return 'true' in [str(value).lower() for value in values]

@app.errorhandler(404)
def handle404(error):
    return redirect(url_for("modsMakerHome", profileFilename="modsprofile"))

#------MODS------

@app.route("/modsmaker", methods=["GET", "POST"])
def modsMakerRedirect():
    return redirect(url_for("modsMakerHome", profileFilename="modsprofile"))

@app.route("/modsmaker/<string:profileFilename>", methods=["GET", "POST"])
def modsMakerHome(profileFilename):
    if request.method == "POST":
        input_file = request.files["input_file"]
        filename = request.files["input_file"].filename
        selectedSheet = request.form.get('sheetlist')
        enforceValidations = getEnforceValidationsFromForm(request.form)

        globalConditions = {}
        for formInput in request.form:
            if formInput not in ['sheetlist', 'enforce_validations']:
                globalConditions[formInput] = True

        if ".xlsx" in filename:
            try:
                zipFile, filename = fileSupport.createZipFromExcel(
                    input_file.read(),
                    selectedSheet,
                    os.path.join("profiles", profileFilename + ".yaml"),
                    globalConditions,
                    enforceValidations,
                )
            except profileValidation.ValidationError as error:
                return render_template('error.html', error=fileSupport.getValidationErrorText(error.errors), title="Error")
            response = make_response(zipFile)
            response.headers["Content-Disposition"] = "attachment; filename=" + filename
            return response

        else:
            return render_template('error.html', error="Please go back and select a .XLSX Excel file to proceed.", title="Error")

    else:
        profile = profileInterpreter.Profile(os.path.join("profiles", profileFilename + ".yaml"))
        return render_template('mods/modsFileSelect.html', profilename=profileFilename, globalconditions=profile.profileGlobalConditions, title="MODS Maker")

@app.route("/processfileupload", methods=["POST"])
def processNewFile():
    if request.method == "POST":
        inputFile = request.files.get("xlsx_file")
        fileName = inputFile.filename
        sheetNames = fileSupport.getSheetNamesFromXlsx(inputFile.read())
        if ".xlsx" in fileName:
            data = {"filename":fileName, "sheetnames": sheetNames}
            return jsonify(data)
        else:
            return render_template('error.html', error="Please go back and select a .XLSX Excel file to proceed.", title="Error")

@app.route("/modsmaker/getpreview", methods=["POST"])
def modsMakerGetPreview():
    if request.method == "POST":
        inputFile = request.files.get("xlsx_file")
        requestDict = json.loads(request.form["data"])
        sheetName = requestDict.get("sheetname")
        profileFilename = requestDict.get("profile")
        globalConditions = requestDict.get("globalconditions", {})
        enforceValidations = getBooleanFromJson(requestDict.get("enforce_validations"), True)
        try:
            previewResult = fileSupport.getPreviewResult(
                inputFile.read(),
                sheetName,
                os.path.join("profiles", profileFilename + ".yaml"),
                globalConditions,
                enforceValidations,
            )
        except profileValidation.ValidationError as error:
            return(jsonify({"errors": error.errors, "error_text": fileSupport.getValidationErrorText(error.errors)}))

        if previewResult.get('warnings'):
            return jsonify(previewResult)

        return(jsonify(previewResult['preview']))

@app.route("/modsmaker/validate", methods=["POST"])
def modsMakerValidate():
    if request.method == "POST":
        inputFile = request.files.get("xlsx_file") or request.files.get("input_file")
        requestDict = json.loads(request.form["data"])
        sheetName = requestDict.get("sheetname")
        profileFilename = requestDict.get("profile")
        enforceValidations = getBooleanFromJson(requestDict.get("enforce_validations"), True)
        validationStatus = fileSupport.getValidationStatusFromExcel(
            inputFile.read(),
            sheetName,
            os.path.join("profiles", profileFilename + ".yaml"),
            enforceValidations,
        )
        return jsonify(validationStatus)

@app.route("/modsmakerapi", methods=["GET", "POST"])
def modsMakerAPI():
    if request.method == "POST":
        pass
    else:
        return "ERROR"

#------EAD------

@app.route("/eadmaker", methods=["GET", "POST"])
def eadMakerHome():
    if request.method == "POST":
        #print(request.get_data(), file=sys.stderr)
        id = str(uuid.uuid4())
        input_file = request.files["input_file"]
        filename = request.files["input_file"].filename

        if ".xlsx" in filename:
            filename = filename.replace("/", " ").replace("\\", " ")
            input_file.save(os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), "legacy", "cache"), id + ".xlsx"))
            return redirect("eadmaker/renderead/" + filename + "/" + id)
        else:
            return render_template('error.html', error="Please go back and select a .XLSX Excel file to proceed.", title="Error")

    else:
        return render_template('ead/eadFileSelect.html', title="EAD Maker")

@app.route("/eadmaker/renderead/<string:filename>/<string:id>", methods=["GET", "POST"])
def eadMakerSelectSheet(filename, id):
    if request.method == "POST":
        print("GET requested", file=sys.stderr)
        select = request.form.get('sheetlist')
        output_data, returndict = processExceltoEAD(os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), "legacy", "cache"), id + ".xlsx"), select, id)
        response = make_response(output_data)
        response.headers["Content-Disposition"] = "attachment; filename=" + returndict["filename"]
        return response
    else:
        sheetnames = getSheetNames(os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), "legacy", "cache"), id + ".xlsx"))
        return render_template('ead/eadAction.html', sheets=sheetnames, publicfilename=filename, id=id, filename=filename, title="EAD Maker")

@app.route("/eadmaker/getpreview", methods=["GET", "POST"])
def eadMakerGetPreview():
    if request.method == "POST":
        print(request.get_json())
        requestDict = request.get_json()
        id = requestDict.get("id")
        select = requestDict.get("sheetname")
        output_data, returndict = processExceltoEAD(os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), "legacy", "cache"), id + ".xlsx"), select, id)
        return(jsonify(returndict["allrecords"]))

####Profiles

@app.route("/profiles/<string:profileFilename>", methods=["GET"])
def displayProfile(profileFilename):
    if request.method == "GET":
        profile = profileInterpreter.Profile(os.path.join("profiles", profileFilename + ".yaml"))
        fieldList = profile.getFieldList()
        yaml = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "profiles", profileFilename + ".yaml")).read()
            
        return render_template('profiles/profile.html', fieldList=fieldList, profileFilename=profileFilename, yaml=yaml, title="Profiles")

@app.route("/profiles/", methods=["GET"])
def displayProfileList():
    if request.method == "GET":
        files = glob(os.path.join(os.path.dirname(os.path.abspath(__file__)), "profiles", "*.yaml"))
        profileList = []
        for file in files:
            profileList.append(os.path.basename(file).replace(".yaml", ""))
        return render_template('profiles/profiles.html', profiles=profileList, title="Profiles")

@app.route("/profiles/downloadprofile/<string:profileFilename>", methods=["GET"])
def downloadYaml(profileFilename):
    path = os.path.join("profiles", profileFilename + ".yaml")
    return flask.send_file(path, as_attachment=True)

####Forms

@app.route("/forms/", methods=["GET"])
def displayFormsList():
    if request.method == "GET":
        files = glob(os.path.join(os.path.dirname(os.path.abspath(__file__)), "profiles", "*.yaml"))
        profileList = []
        for file in files:
            profileList.append(os.path.basename(file).replace(".yaml", ""))
        return render_template('forms/forms.html', profiles=profileList, title="Forms")

@app.route("/forms/profile/<string:profileFilename>", methods=["GET", "POST"])
def displayProfileForm(profileFilename):
    if request.method == "POST":
        metadata = request.form.to_dict()
        globalConditions = dict.fromkeys(metadata, True)
        xmlFile, filename = fileSupport.createFileFromRow(metadata, 0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "profiles", profileFilename + ".yaml"), globalConditions)
        response = make_response(xmlFile)
        response.headers["Content-Disposition"] = "attachment; filename=" + filename
        return response
    if request.method == "GET":
        profile = profileInterpreter.Profile(os.path.join("profiles", profileFilename + ".yaml"))
        fieldList = profile.getFieldList()

        allHeaders = {}

        for index, field in enumerate(fieldList):
            headers = field.get("headers", [])

            for header in headers:

                if allHeaders.get(header, None):
                    headers.remove(header)
                    fieldList[index]["headers"] = headers
                else:
                   allHeaders[header] = "here"

        filenameColumn = profile.profileFilenameColumn
            
        return render_template('forms/form.html', fieldList=fieldList, profilename=profileFilename, globalconditions=profile.profileGlobalConditions, filenameColumn=filenameColumn, title="Forms")

@app.route("/forms/profile/<string:profileFilename>/preview", methods=["POST"])
def getFormPreview(profileFilename):
    if request.method == "POST":
        metadata = request.form.to_dict()
        globalConditions = dict.fromkeys(metadata, True)
        preview = fileSupport.createPreviewFromRows([metadata], os.path.join(os.path.dirname(os.path.abspath(__file__)), "profiles", profileFilename + ".yaml"), globalConditions)
        return jsonify(preview)

@app.route("/resources", methods=["GET"])
def renderResources():
    return render_template('resources.html', title="Resources")
