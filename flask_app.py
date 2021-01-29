from flask import jsonify, Flask, make_response, request, render_template, redirect, g, url_for

import EADMaker
from EADMaker import processExceltoEAD
from EADMaker import getSheetNames
from MODSMaker import processExceltoMODS
import sys
import uuid
import os
import json

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
            input_file.save(os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache"), id + ".xlsx"))
            return redirect("eadmaker/renderead/" + filename + "/" + id)
        else:
            return render_template('error.html', error="Please go back and select a .XLSX Excel file to proceed.", title="Error")

    else:
        return render_template('home.html', title="EAD Maker")

@app.route("/eadmaker/renderead/<string:filename>/<string:id>", methods=["GET", "POST"])
def eadMakerSelectSheet(filename, id):
    print("TRYING TO RENDER RENDERTEMPLATE", file=sys.stderr)
    #print(g.sheetnames, file=sys.stderr)
    if request.method == "POST":
        print("GET requested", file=sys.stderr)
        select = request.form.get('sheetlist')
        #print(select, file=sys.stderr)
        output_data, returndict = processExceltoEAD(os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache"), id + ".xlsx"), select, id)
        #print(output_data, file=sys.stderr)
        response = make_response(output_data)
        response.headers["Content-Disposition"] = "attachment; filename=" + returndict["filename"]
        return response
    else:
        sheetnames = getSheetNames(os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache"), id + ".xlsx"))
        return render_template('resultspage.html', sheets=sheetnames, publicfilename=filename, id=id, filename=filename, title="EAD Maker")

@app.route("/eadmakerapi", methods=["GET", "POST"])
def eadMakerAPI():
    if request.method == "POST":
        print("REQUESTREQUEST", file=sys.stderr)
        #print(request.get_data(), file=sys.stderr)
        id = str(uuid.uuid4())
        input_file = request.files['file']
        filename = request.files['file'].filename
        filename = filename.replace("/", " ").replace("\\", " ")
        #print(input_file)
        if ".xlsx" in filename:
            input_file.save(os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache"), id + ".xlsx"))
            #input_data = input_file.stream.read()
            return "eadmaker/renderead/" + filename + "/" + id
        else:
            return render_template('error.html', error="Please go back and select a .XLSX Excel file to proceed.", title="Error")
    else:
        return "ERROR"

#------MODS------

@app.route("/modsmaker", methods=["GET", "POST"])
def modsMakerHome():
    if request.method == "POST":
        #print(request.get_data(), file=sys.stderr)
        id = str(uuid.uuid4())
        input_file = request.files["input_file"]
        filename = request.files["input_file"].filename
        filename = filename.replace("/", " ").replace("\\", " ")
        #print(input_file)
        if ".xlsx" in filename:
            input_file.save(os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache"), id + ".xlsx"))
            #input_data = input_file.stream.read()
            return redirect("modsmaker/rendermods/" + filename + "/" + id)
        else:
            return render_template('error.html', error="Please go back and select a .XLSX Excel file to proceed.", title="Error")
    else:
        return render_template('MODSfileselect.html', title="MODS Maker")

@app.route("/processfileupload", methods=["POST"])
def processNewFile():
    if request.method == "POST":
        fileUid = str(uuid.uuid4())
        inputFile = request.files.get("xlsx_file")
        fileName = inputFile.filename
        if ".xlsx" in fileName:
            filePath = os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache"), fileUid + ".xlsx")
            inputFile.save(os.path.join(filePath))
            data = {"filename":fileName, "sheetnames": getSheetNames(filePath), "uid": fileUid}
            return jsonify(data)
        else:
            return render_template('error.html', error="Please go back and select a .XLSX Excel file to proceed.", title="Error")

@app.route("/modsmaker/returnmods/<string:id>", methods=["POST"])
def modsMakerReturnMods(id):
    #print(g.sheetnames, file=sys.stderr)
    if request.method == "POST":
        print("GET requested", file=sys.stderr)
        select = request.form.get('sheetlist')
        includeDefaults = True
        if request.form.get('defaultsCheckbox', None) == None:
            includeDefaults = False
        #print(select, file=sys.stderr)
        output_data, returndict = processExceltoMODS(os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache"), id + ".xlsx"), select, id, includeDefaults)
        #print(output_data, file=sys.stderr)
        response = make_response(output_data)
        response.headers["Content-Disposition"] = "attachment; filename=" + returndict["filename"]
        return response

@app.route("/modsmaker/getpreview", methods=["GET", "POST"])
def modsMakerGetPreview():
    if request.method == "POST":
        print(request.get_json())
        requestDict = request.get_json()
        print(requestDict)
        id = requestDict.get("id")
        select = requestDict.get("sheetname")
        includeDefaults = requestDict.get('includedefaults', True)
        output_data, returndict = processExceltoMODS(os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache"), id + ".xlsx"), select, id, includeDefaults)
        return(jsonify(returndict["allrecords"]))

@app.route("/eadmaker/getpreview", methods=["GET", "POST"])
def eadMakerGetPreview():
    if request.method == "POST":
        print(request.get_json())
        requestDict = request.get_json()
        id = requestDict.get("id")
        select = requestDict.get("sheetname")
        output_data, returndict = processExceltoEAD(os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache"), id + ".xlsx"), select, id)
        return(jsonify(returndict["allrecords"]))


@app.route("/modsmakerapi", methods=["GET", "POST"])
def modsMakerAPI():
    if request.method == "POST":
        print("REQUESTREQUEST", file=sys.stderr)
        #print(request.get_data(), file=sys.stderr)
        id = str(uuid.uuid4())
        input_file = request.files['file']
        filename = request.files['file'].filename
        filename = filename.replace("/", " ").replace("\\", " ")
        #print(input_file)
        if ".xlsx" in filename:
            input_file.save(os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache"), id + ".xlsx"))
            #input_data = input_file.stream.read()
            return "modsmaker/rendermods/" + filename + "/" + id
        else:
            return render_template('error.html', error="Please go back and select a .XLSX Excel file to proceed.", title="Error")

    else:
        return "ERROR"

@app.route("/resources", methods=["GET"])
def renderResources():
    return render_template('resources.html', title="Resources")