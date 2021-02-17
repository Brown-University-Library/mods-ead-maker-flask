from flask import jsonify, Flask, make_response, request, render_template, redirect, g, url_for
import flask
from legacy.EADMaker import processExceltoEAD
from legacy.EADMaker import getSheetNames
import profileInterpreter
import sys
import uuid
import os
import json
import fileSupport
from glob import glob

app = Flask(__name__)
app.config["DEBUG"] = True

@app.route("/", methods=["GET", "POST"])
def redirectToEADMaker():
    return redirect(url_for('modsMakerHome'))

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

        globalConditions = {}
        for formInput in request.form:
            globalConditions[formInput] = True

        if ".xlsx" in filename:
            zipFile, filename = fileSupport.createZipFromExcel(input_file.read(), selectedSheet, os.path.join("profiles", profileFilename + ".yaml"),globalConditions)
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
        sheetNames = fileSupport.getSheetNames(inputFile.read())
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
        globalConditions = requestDict.get("globalconditions", {})
        preview = fileSupport.getPreview(inputFile.read(), sheetName, "profiles/modsprofile.yaml", globalConditions)
        return(jsonify(preview))

@app.route("/modsmakerapi", methods=["GET", "POST"])
def modsMakerAPI():
    if request.method == "POST":
        pass
    else:
        return "ERROR"

####Extra

@app.route("/profiles/<string:profileFilename>", methods=["GET"])
def displayProfile(profileFilename):
    if request.method == "GET":
        modsMaker = profileInterpreter.Profile(os.path.join("profiles", profileFilename + ".yaml"))
        fieldList = modsMaker.getFieldList()
        yaml = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "profiles", profileFilename + ".yaml")).read()
            
        return render_template('profiles/profile.html', fieldList=fieldList, profilename=profileFilename, yaml=yaml, title="Profiles")

@app.route("/profiles/", methods=["GET"])
def modsMakerDisplayProfiles():
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

@app.route("/resources", methods=["GET"])
def renderResources():
    return render_template('resources.html', title="Resources")