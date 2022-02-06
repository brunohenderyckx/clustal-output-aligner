from flask import Flask, render_template, request, Response
from openpyxl.writer.excel import save_virtual_workbook
from openpyxl import Workbook
from openpyxl.styles.fills import PatternFill
from openpyxl.styles import colors, Font
from openpyxl.utils import get_column_letter


def align_data(alignment_input, alignment_number, output_wrap):
    """
    Reads in Clustal Alignment Data and returns a openpyxl worksheet object
    """

    # defining 2 types of font to populate the excel with
    stft = Font(name='Consolas', size=10.5)
    ft = Font(name='Consolas', size=10.5, color=colors.WHITE)

    # takes the Clustal data and splits it into a list of rows if there's at least 5 characters
    source_data = alignment_input.splitlines()
    source_data = [x for x in source_data if len(x)>5]

    # generates a list of row indexes that will be used to compare alignments against
    index_of_first_row = [x for x in range(0, 100, alignment_number + 1)]

    # create and activagte excel worksheet
    workbook = Workbook()
    sheet = workbook.active

    # figures out what the column-indices are for the alignment comparisons
    column_range = []
    for index, character in enumerate(source_data[0]):
        if index != 0 and (character.isupper() or character=="-") and (" " in source_data[0][:index]): # only upper characters or -
            column_range.append(index)

    # generate all row indices and loop over the rows
    for i in range(0,len(source_data)): #
        row_match_counter = [0,0,0]

        # save the name to column A
        sheet.cell(row=i+2, column=1).value = source_data[i][0:min(column_range)].strip()
        sheet.cell(row=i+2, column=1).font = stft

        # generate only non-name column indices and loop over the columns
        for j in range(min(column_range), len(source_data[i]) - min(column_range)):

            # save the output to excel with the standard font
            sheet.cell(row=i+2, column=j+2).value = source_data[i][j]
            sheet.cell(row=i+2, column=j+2).font = stft

            if j>= min(column_range) and j<= max(column_range):

                # first check which row to compare against, 0 or 11
                if i in index_of_first_row:
                    comparator_row = i

                # then check if the row is the comparator row, if it isn't check if it a perfect or fuzzy match
                if i == comparator_row and source_data[i][j] != "-" and source_data[i][j] != " ":
                    sheet.cell(row=i+2, column=j+2).fill = PatternFill(fgColor="494544", fill_type = "solid") # Dark color
                    sheet.cell(row=i+2, column=j+2).font = ft
                    row_match_counter[0] += 1

                else:
                    try:
                        if ((source_data[comparator_row][j] == source_data[i][j]) and (source_data[i][j] != " ") and (source_data[i][j] != "-")):
                            sheet.cell(row=i+2, column=j+2).fill = PatternFill(fgColor="494544", fill_type = "solid") # Dark color
                            sheet.cell(row=i+2, column=j+2).font = ft
                            row_match_counter[0] += 1

                        elif (source_data[i][j] in ("D","E","N","Q") and source_data[comparator_row][j] in ("D","E","N","Q")) or \
                        (source_data[i][j] in ("K","R","H") and source_data[comparator_row][j] in ("K","R","H")) or \
                        (source_data[i][j] in ("F","W","Y") and source_data[comparator_row][j] in ("F","W","Y")) or \
                        (source_data[i][j] in ("V","I","L","M") and source_data[comparator_row][j] in ("V","I","L","M")) or \
                        (source_data[i][j] in ("S","T") and source_data[comparator_row][j] in ("S","T")):
                            sheet.cell(row=i+2, column=j+2).fill = PatternFill(fgColor="7A7675", fill_type = "solid") # Light color
                            sheet.cell(row=i+2, column=j+2).font = ft
                            row_match_counter[1] += 1

                        elif (source_data[i][j] != " ") and (source_data[i][j] != "-"):
                            row_match_counter[2] += 1

                        else:
                            pass

                    except:
                        pass

        # determines the
        summary_column = 100

        sheet.cell(row=i+2, column=summary_column).value = str(row_match_counter[0])
        sheet.cell(row=i+2, column=summary_column + 1).value = str(row_match_counter[1])
        sheet.cell(row=i+2, column=summary_column + 2).value = str(row_match_counter[2])
        sheet.cell(row=i+2, column=summary_column + 3).value = sum(row_match_counter)

    sheet.cell(row=1,column=1).value = "Name"
    sheet.cell(row=1,column=summary_column).value = "# Match"
    sheet.cell(row=1,column=summary_column + 1).value = "# Fuzzy"
    sheet.cell(row=1,column=summary_column + 2).value = "# No Match"
    sheet.cell(row=1,column=summary_column + 3).value = "Total"

    # set the width of the columns to optimize viewing display
    i = get_column_letter(1)
    sheet.column_dimensions[i].width = 20

    column = 2
    while column < 601:

        if column > 99:
            i = get_column_letter(column)
            sheet.column_dimensions[i].width = 3
            column += 1
        else:
            i = get_column_letter(column)
            sheet.column_dimensions[i].width = 1.5
            column += 1

    sheet.sheet_view.showGridLines = False
    return workbook


app = Flask(__name__)


@app.route("/")
@app.route("/protein-aligner")
def index():
    return render_template("index.html")

@app.route("/execute-protein", methods=["POST"])
def execute():
    alignment_input = request.form['alignment_input']
    alignment_number = int(request.form['number_input'])
    output_wrap = request.form['alignment_input']

    alignment_file = align_data(alignment_input, alignment_number, output_wrap)
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
