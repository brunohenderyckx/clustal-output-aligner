from flask import Flask, Response, render_template, request
from openpyxl import Workbook
from openpyxl.styles import Font, colors
from openpyxl.styles.fills import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.writer.excel import save_virtual_workbook

import functions as functions

app = Flask(__name__)


@app.route("/")
@app.route("/protein-aligner")
def index():
    return render_template("index.html")


@app.route("/execute-protein", methods=["POST"])
def execute_P():
    # reading in the inputs from the html file
    alignment_input = request.form['alignment_input']
    matching_rules = functions.create_matching_dict(request.form['match_dictionary'])

    # executing the protein aligner function based on the wrap option
    alignment_file = functions.protein_aligner_single(alignment_input, matching_rules)
    

    # saves the excel workbook based on the excel file in memory from the alignment_file workbook in memory
    return Response(
        save_virtual_workbook(alignment_file),
        headers={
            'Content-Disposition': 'attachment; filename=Clustal_output.xlsx',
            'Content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
    )


@app.route("/execute-nucleic", methods=["POST"])
def execute_N():
    # reading in the inputs from the html file
    alignment_input = request.form['alignment_input']
    matching_rules = ""

    # executing the protein aligner function based on the wrap option
    alignment_file = functions.protein_aligner_single(alignment_input, matching_rules)
    
    # saves the excel workbook based on the excel file in memory from the alignment_file workbook in memory
    return Response(
        save_virtual_workbook(alignment_file),
        headers={
            'Content-Disposition': 'attachment; filename=Clustal_output.xlsx',
            'Content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
    )


@app.route('/nucleic-acid-aligner')
def nucleicacid():
    return render_template('/nucleic-acid.html')


@app.route('/help')
def help():
    return render_template('/help.html')


@app.route('/about')
def about():
    return render_template('/about.html')
