from flask import Flask, Response, render_template, request
from openpyxl import Workbook
from openpyxl.styles import Font, colors
from openpyxl.styles.fills import PatternFill
from openpyxl.utils import get_column_letter


def soft_check(cell, ref_cell, matching_rules=[("D", "E", "N", "Q"),
                                               ("K", "R", "H"),
                                               ("F", "W", "Y"),
                                               ("V", "I", "L", "M"),
                                               ("S", "T")]):
    """ Checks if the cell value matches the matching list of tuples of the reference cell

    Args:
        cell (string): A row of Clustal output
        ref_cell (string): A row of Clustal output
        matching_dict (dictionary): A row of Clustal output

    Returns:
        Bool: Returns True if the values or soft matches
    """
    for group in matching_rules:
        if cell in group and ref_cell in group:
            return True
    else:
        return False


def create_matching_dict():

    return None


def row_protein_length(row):
    """ Returns length of protein  string in Clustal rows

    Args:
        row (list): A row of Clustal output

    Returns:
        length: An integer for the number of protein characters in the row
    """
    length = 0
    print(row)
    for index, character in enumerate(row):
        if index != 0 and \
            (character.isupper() or character == "-") and \
                (" " in row[:index]):
            print(character, ": Match")
            length += 1

    return length


def convert_raw_clustal(source_data, cut_of_number=5):
    """
    Converts a clustal output string into a list of rows we want to work with.
    Empty lines or lines with only non-alphabetical characters get removed.
    Cut of number determines how many characters the rows minimally need to have to be considered.
    """

    # takes the Clustal data and splits it into a list of rows if there's at least 5 characters
    source_data = source_data.splitlines()
    source_data = [x for x in source_data if len(x) > cut_of_number]

    # removes the list item if there's no upper case character
    rows_to_remove = []
    for x in source_data:
        if any(ele.isupper() for ele in x) == False:
            rows_to_remove.append(x)

    # need to remove the list items if they are in the index list
    source_data = [item for item in source_data if item not in rows_to_remove]

    return source_data


def protein_aligner_single(alignment_input, alignment_number):
    """
    Generates the excel file with the data and color the alignment in a single row per protein
    """
    PADDING_LEFT = 1
    PADDING_RIGHT = 1
    PADDING_BOTTOM = 1
    PADDING_TOP = 1
    EXCEL_OFFSET = 1
    stft = Font(name='Consolas', size=10.5)
    ft = Font(name='Consolas', size=10.5, color=colors.WHITE)
    HARD_COLOR = "494544"
    SOFT_COLOR = "7A7675"

    # generates the Clustal data
    source_data = convert_raw_clustal(alignment_input, 5)

    # create and activate excel worksheet
    workbook = Workbook()
    sheet = workbook.active

    # figures out what the column-indices are for the alignment comparisons
    column_range = []
    for index, character in enumerate(source_data[0]):
        if index != 0 and \
            (character.isupper() or character == "-") and \
                (" " in source_data[0][:index]):  # only upper characters or -
            column_range.append(index)

    # generate a dictionary with unique series names and empty lists to store source_data rows in
    species = {}
    for row in source_data:
        name = row[0:min(column_range)].strip()
        if name and not (name in species):
            species[name] = []

    # save the names to column A
    for i, key in enumerate(species):
        sheet.cell(row=i + PADDING_TOP + EXCEL_OFFSET, column=1).value = key
        sheet.cell(row=i + PADDING_TOP + EXCEL_OFFSET, column=1).font = stft
        species[key].append(int(i) + PADDING_TOP + EXCEL_OFFSET)

    # generate all row indices and loop over the rows
    for i in range(0, len(source_data)):

        # saves the species_name of the row we are at
        # so we can check the dictionary
        row_name = source_data[i][0:min(column_range)].strip()

        if row_name not in species:  # checks if the row name is in the dictionary or not
            continue

        excel_row = min(species[row_name])
        series_rows_processed = len(species[row_name])

        print("For", source_data[i], row_protein_length(source_data[i]))

        # generates a range from the minimum index in our protein column range
        # to the maximum + 1 index except if the row is smaller than that
        #
        for j in range(min(column_range),
                       min(
            max(column_range) + 1,
            row_protein_length(source_data[i]) + min(column_range))
        ):

            # save the output to excel with the standard font
            # now we don't need to increment the row
            # we instead need to add it to the row in the dictionary
            try:
                char_to_write = source_data[i][j]

                col_to_write = j + EXCEL_OFFSET + \
                    (len(column_range) * (series_rows_processed - 1)) - \
                    (min(column_range) - 1)

                sheet.cell(row=excel_row,
                           column=col_to_write).value = char_to_write
                sheet.cell(row=excel_row, column=col_to_write).font = stft
                
            except:
                pass


        species[row_name].append(int(i + EXCEL_OFFSET + PADDING_RIGHT))

    """
    Colouring based on matching for Excel Cells
    """
    # how many columns?
    last_empty_column = len(list(sheet.columns))

    # TO DO
    # IMPLEMENT FORM FOR USER TO DEFINE THE MATCHING RULES THEMSELVES
    #

    for i in range(EXCEL_OFFSET + 1, len(species) + EXCEL_OFFSET + 1):
        row_match_counter = [0, 0, 0]
        for j in range(2, last_empty_column + EXCEL_OFFSET):
            cell = sheet.cell(row=i, column=j).value
            ref_cell = sheet.cell(row=2, column=j).value
            print(ref_cell, cell)

            # compare if hard match
            if cell == ref_cell and ref_cell != "-":
                sheet.cell(row=i, column=j).fill = PatternFill(
                    fgColor=HARD_COLOR, fill_type="solid")
                sheet.cell(row=i, column=j).font = ft
                row_match_counter[0] += 1

            # compare if soft match
            # TO DO - PASS MATCHING RULES ONCE IMPLEMENTED
            elif soft_check(cell, ref_cell) and ref_cell != "-":
                sheet.cell(row=i, column=j).fill = PatternFill(
                    fgColor=SOFT_COLOR, fill_type="solid")
                sheet.cell(row=i, column=j).font = ft
                row_match_counter[1] += 1

            elif ref_cell == "-" or cell == "-":
                pass

            # no match
            else:
                row_match_counter[2] += 1

        # save counter aggregates reference species
        for index, value in enumerate(row_match_counter):
            sheet.cell(row=i, column=last_empty_column +
                       EXCEL_OFFSET + PADDING_RIGHT + index).value = value

        sheet.cell(row=i, column=last_empty_column + EXCEL_OFFSET +
                   PADDING_RIGHT + 3).value = sum(row_match_counter)

    """
    Format the rest of the excel file
    """
    # print Count header
    sheet.cell(row=1, column=1).value = "Name"

    header_counter_names = ['# Match', '# Fuzzy', '# No Match', 'Total']
    for i, header_name in enumerate(header_counter_names):
        sheet.cell(row=1, column=last_empty_column + EXCEL_OFFSET +
                   PADDING_RIGHT + i).value = header_name

    # set the width of the columns to optimize viewing display
    i = get_column_letter(1)
    sheet.column_dimensions[i].width = 20
    sheet.column_dimensions[i].width = 20

    column = 2
    while column < (last_empty_column + len(header_counter_names) + 1):

        if column > last_empty_column + PADDING_RIGHT:
            i = get_column_letter(column)
            sheet.column_dimensions[i].width = 12
            column += 1

        else:
            i = get_column_letter(column)
            sheet.column_dimensions[i].width = 1.5
            column += 1

    sheet.sheet_view.showGridLines = False
    return workbook


def protein_aligner_wrap(alignment_input, alignment_number):
    """
    Generates the excel file with the data and color the alignment
    """

    # Defining constants and variables
    PADDING_LEFT = 1
    PADDING_RIGHT = 1
    PADDING_BELOW = 1
    stft = Font(name='Consolas', size=10.5)  # font for excel
    ft = Font(name='Consolas', size=10.5, color=colors.WHITE)  # font for excel

    # create and activate excel worksheet
    workbook = Workbook()
    sheet = workbook.active

    # converst the Clustal data in
    source_data = convert_raw_clustal(alignment_input, 5)

    # generates a list of row indexes that will be used to compare alignments against
    index_of_first_row = [x for x in range(0, 100, alignment_number + 1)]

    # figures out what the column-indices are for the alignment comparisons
    column_range = []
    for index, character in enumerate(source_data[0]):
        # only upper characters or -
        if index != 0 and (character.isupper() or character == "-") and (" " in source_data[0][:index]):
            column_range.append(index)

    # generate all row indices and loop over the rows
    for i in range(0, len(source_data)):
        row_match_counter = [0, 0, 0]

        # save the name to column A
        sheet.cell(
            row=i+2, column=1).value = source_data[i][0:min(column_range)].strip()
        sheet.cell(row=i+2, column=1).font = stft

        # generate only non-name column indices and loop over the columns
        for j in range(min(column_range), len(source_data[i]) - min(column_range)):

            # save the output to excel with the standard font
            sheet.cell(row=i+2, column=j+2).value = source_data[i][j]
            sheet.cell(row=i+2, column=j+2).font = stft

            if j >= min(column_range) and j <= max(column_range):

                # first check which row to compare against, 0 or 11
                if i in index_of_first_row:
                    comparator_row = i

                # then check if the row is the comparator row, if it isn't check if it a perfect or fuzzy match
                if i == comparator_row and source_data[i][j] != "-" and source_data[i][j] != " ":
                    sheet.cell(row=i+2, column=j+2).fill = PatternFill(
                        fgColor="494544", fill_type="solid")  # Dark color
                    sheet.cell(row=i+2, column=j+2).font = ft
                    row_match_counter[0] += 1

                else:
                    try:
                        if ((source_data[comparator_row][j] == source_data[i][j]) and (source_data[i][j] != " ") and (source_data[i][j] != "-")):
                            sheet.cell(row=i+2, column=j+2).fill = PatternFill(
                                fgColor="494544", fill_type="solid")  # Dark color
                            sheet.cell(row=i+2, column=j+2).font = ft
                            row_match_counter[0] += 1

                        elif (source_data[i][j] in ("D", "E", "N", "Q") and source_data[comparator_row][j] in ("D", "E", "N", "Q")) or \
                            (source_data[i][j] in ("K", "R", "H") and source_data[comparator_row][j] in ("K", "R", "H")) or \
                            (source_data[i][j] in ("F", "W", "Y") and source_data[comparator_row][j] in ("F", "W", "Y")) or \
                            (source_data[i][j] in ("V", "I", "L", "M") and source_data[comparator_row][j] in ("V", "I", "L", "M")) or \
                                (source_data[i][j] in ("S", "T") and source_data[comparator_row][j] in ("S", "T")):
                            sheet.cell(row=i+2, column=j+2).fill = PatternFill(
                                fgColor="7A7675", fill_type="solid")  # Light color
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

        sheet.cell(
            row=i+2, column=summary_column).value = str(row_match_counter[0])
        sheet.cell(row=i+2, column=summary_column +
                   1).value = str(row_match_counter[1])
        sheet.cell(row=i+2, column=summary_column +
                   2).value = str(row_match_counter[2])
        sheet.cell(row=i+2, column=summary_column +
                   3).value = sum(row_match_counter)

    sheet.cell(row=1, column=1).value = "Name"
    sheet.cell(row=1, column=summary_column).value = "# Match"
    sheet.cell(row=1, column=summary_column + 1).value = "# Fuzzy"
    sheet.cell(row=1, column=summary_column + 2).value = "# No Match"
    sheet.cell(row=1, column=summary_column + 3).value = "Total"

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
