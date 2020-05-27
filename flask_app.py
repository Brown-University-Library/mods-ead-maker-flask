from flask import Flask, make_response, request, render_template, redirect, g, url_for

import EADMaker
from EADMaker import processExceltoEAD
from EADMaker import getSheetNames
from MODSMaker import processExceltoMODS
import sys
import uuid

app = Flask(__name__)
app.config["DEBUG"] = True

@app.route("/", methods=["GET", "POST"])
def redirectToEADMaker():
    return redirect(url_for('eadMakerHome'))

@app.route("/eadmaker", methods=["GET", "POST"])
def eadMakerHome():
    if request.method == "POST":
        #print(request.get_data(), file=sys.stderr)
        id = str(uuid.uuid4())
        input_file = request.files["input_file"]
        filename = request.files["input_file"].filename

        if ".xlsx" in filename:
            filename = filename.replace("/", " ").replace("\\", " ")
            #print(input_file)
            input_file.save("/home/codyross/eadmaker/cache/" + id + ".xlsx")
            #input_data = input_file.stream.read()
            return redirect("eadmaker/renderead/" + filename + "/" + id)
        else:
            return render_template('error.html', error="Uploaded file must be a .XLSX Excel file.")
        #return render_template('resultspage.html', sheets=sheetnames, publicfilename=request.files["input_file"].name, privatefilename="/home/codyross/eadmaker/cache/" + id + ".xlsx")
        #output_data = processExceltoEAD("/home/codyross/eadmaker/cache/input.xlsx", "seed-list (3)")
        #print(output_data, file=sys.stderr)
        #response = make_response(output_data)
        #response.headers["Content-Disposition"] = "attachment; filename=result.xml"
        #return response

    else:
        return render_template('home.html')

@app.route("/eadmaker/renderead/<string:filename>/<string:id>", methods=["GET", "POST"])
def eadMakerSelectSheet(filename, id):
    print("TRYING TO RENDER RENDERTEMPLATE", file=sys.stderr)
    #print(g.sheetnames, file=sys.stderr)
    if request.method == "POST":
        print("GET requested", file=sys.stderr)
        select = request.form.get('sheetlist')
        #print(select, file=sys.stderr)
        output_data, returndict = processExceltoEAD("/home/codyross/eadmaker/cache/" + id + ".xlsx", select, id)
        #print(output_data, file=sys.stderr)
        response = make_response(output_data)
        response.headers["Content-Disposition"] = "attachment; filename=" + returndict["filename"]
        return response
    else:
        sheetnames = getSheetNames("/home/codyross/eadmaker/cache/" + id + ".xlsx")
        return render_template('resultspage.html', sheets=sheetnames, publicfilename=filename, id=id, filename=filename)

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
            input_file.save("/home/codyross/eadmaker/cache/" + id + ".xlsx")
            #input_data = input_file.stream.read()
            return "eadmaker/renderead/" + filename + "/" + id
        else:
            return render_template('error.html', error="Uploaded file must be a .XLSX Excel file.")

        #return render_template('resultspage.html', sheets=sheetnames, publicfilename=request.files["input_file"].name, privatefilename="/home/codyross/eadmaker/cache/" + id + ".xlsx")
        #output_data = processExceltoEAD("/home/codyross/eadmaker/cache/input.xlsx", "seed-list (3)")
        #print(output_data, file=sys.stderr)
        #response = make_response(output_data)
        #response.headers["Content-Disposition"] = "attachment; filename=result.xml"
        #return response

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
            input_file.save("/home/codyross/eadmaker/cache/" + id + ".xlsx")
            #input_data = input_file.stream.read()
            return redirect("modsmaker/rendermods/" + filename + "/" + id)
        else:
            return render_template('error.html', error="Uploaded file must be a .XLSX Excel file.")
        #return render_template('resultspage.html', sheets=sheetnames, publicfilename=request.files["input_file"].name, privatefilename="/home/codyross/eadmaker/cache/" + id + ".xlsx")
        #output_data = processExceltoEAD("/home/codyross/eadmaker/cache/input.xlsx", "seed-list (3)")
        #print(output_data, file=sys.stderr)
        #response = make_response(output_data)
        #response.headers["Content-Disposition"] = "attachment; filename=result.xml"
        #return response

    else:
        return render_template('homeMODS.html')

@app.route("/modsmaker/rendermods/<string:filename>/<string:id>", methods=["GET", "POST"])
def modsMakerSelectSheet(filename, id):
    print("TRYING TO RENDER RENDERTEMPLATE", file=sys.stderr)
    #print(g.sheetnames, file=sys.stderr)
    if request.method == "POST":
        print("GET requested", file=sys.stderr)
        select = request.form.get('sheetlist')
        #print(select, file=sys.stderr)
        output_data, returndict = processExceltoMODS("/home/codyross/eadmaker/cache/" + id + ".xlsx", select, id)
        #print(output_data, file=sys.stderr)
        response = make_response(output_data)
        response.headers["Content-Disposition"] = "attachment; filename=" + returndict["filename"]
        return response
    else:
        sheetnames = getSheetNames("/home/codyross/eadmaker/cache/" + id + ".xlsx")
        return render_template('resultspageMODS.html', sheets=sheetnames, publicfilename=filename, id=id, filename=filename)

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
            input_file.save("/home/codyross/eadmaker/cache/" + id + ".xlsx")
            #input_data = input_file.stream.read()
            return "modsmaker/rendermods/" + filename + "/" + id
        else:
            return render_template('error.html', error="Uploaded file must be a .XLSX Excel file.")
        #return render_template('resultspage.html', sheets=sheetnames, publicfilename=request.files["input_file"].name, privatefilename="/home/codyross/eadmaker/cache/" + id + ".xlsx")
        #output_data = processExceltoEAD("/home/codyross/eadmaker/cache/input.xlsx", "seed-list (3)")
        #print(output_data, file=sys.stderr)
        #response = make_response(output_data)
        #response.headers["Content-Disposition"] = "attachment; filename=result.xml"
        #return response

    else:
        return "ERROR"
