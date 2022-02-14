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
def execute():
    # reading in the inputs from the html file
    alignment_input = request.form['alignment_input']
    alignment_number = int(request.form['number_input'])
    output_wrap = request.form['wrapornot']
    matching_rules = functions.create_matching_dict(request.form['match_dictionary'])

    print(matching_rules)

    # executing the protein aligner function based on the wrap option
    if output_wrap == "output_wrap":
        alignment_file = functions.protein_aligner_wrap(
            alignment_input, alignment_number)
    elif output_wrap == "output_single":
        alignment_file = functions.protein_aligner_single(
            alignment_input, matching_rules)
    else:
        return None

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
