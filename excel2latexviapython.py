# EXCEL TO LATEX VIA PYTHON
########################################################################################################################
#
# This code takes an excel file with tables in each worksheet and outputs the TeX code for each table as a separate file
# ready to be imported into a LaTeX document.


# USER INPUT AND SETTINGS
# ======================================================================================================================
#

# Define the path of the excel file containing all the tables
input_excel_filename = 'D:/Users/Kirker/Dropbox/My research/Productivity-From-New-Workers/Paper/tables/' +  \
                       'prod_spillover_tables_4_paper.xlsx'

# Select the directory/folder to save the resulting TeX files to
output_dir = 'D:/Users/Kirker/Dropbox/My research/Productivity-From-New-Workers/Paper/tables/'

# Define user settings in a dictionary
# ------------------------------------
#
# Current options:
#
#    booktabs: True/False
#        True = use the booktabs package functions to make prettier horizontal lines 
#        False = Use standard \hlines
#
#   includetabular: True/False
#        Should the code include the tabular environment code around the table, or just return the table rows.
#
#   roundtodp: True/False
#       Should all numbers in the table be rounded to a set number of d.p.?
#
#   numdp: scalar
#       Define how many d.p. to round numbers to if roundtpdp=True

usr_settings = {'booktabs': True, 'includetabular': True, 'roundtodp': True, 'numdp': 3}

# End of user input
# ======================================================================================================================

# PREAMBLE
# ======================================================================================================================

# Load in required packages
import openpyxl  # Package for reading excel files (.xlsx) into Python
from itertools import compress
from sub_functions import e2l  # Model functions to simulate the LBH model


# Print output so user can follow progress.
print('EXCEL 2 LATEX VIA PYTHON')
print('Creates .TeX table files from excel file.')
print('Source file: ' + input_excel_filename)

# Load in the Excel workbook/file
workbook = openpyxl.load_workbook(filename=input_excel_filename)

print(' ')
print('Output directory: ' + output_dir)
print(' ')
print('User settings:')
print('    booktabs: ' + str(usr_settings['booktabs']))
print('    includetabular: ' + str(usr_settings['includetabular']))
print('    roundtodp: ' + str(usr_settings['roundtodp']))
print('    numdp: ' + str(usr_settings['numdp']))
print(' ')
print(' ')
print('Creating TeX tables:')


# MAIN CODE
# ======================================================================================================================

for sheet_name in workbook.get_sheet_names():  # Loop over every worksheet (tab) within the workbook

    # Get the worksheet object for this iteration
    sheet = workbook[sheet_name]

    max_col2use = sheet.max_column
    max_row2use = sheet.max_row

    start_col_idx = 0
    start_row_idx = 0

    end_col_idx = sheet.max_column - 1
    end_row_idx = sheet.max_row - 1

    # Trim off any empty columns at the end of the table
    for col_num in range(sheet.max_column-1, -1, -1):

        if e2l.all_nones(sheet.columns[col_num]):
            # Trim the column for the sheet
            end_col_idx = end_col_idx -1
        else:
            # current final column has value
            break

    # Trim off any empty rows at the end of the table
    for row_num in range(sheet.max_row-1, -1, -1):

        if e2l.all_nones(sheet.rows[row_num]):
            # Trim the column for the sheet
            end_row_idx = end_row_idx - 1
        else:
            # current final column has value
            break

    # Trim off any empty columns at the start of the table
    for col_num in range(0, sheet.max_column):

        if e2l.all_nones(sheet.columns[col_num]):
            # Trim the column for the sheet
            start_col_idx = start_col_idx + 1
        else:
            # current final column has value
            break

    # Trim off any empty rows at the start of the table
    for row_num in range(0, sheet.max_row):

        if e2l.all_nones(sheet.rows[row_num]):
            # Trim the column for the sheet
            start_row_idx = start_row_idx + 1
        else:
            # current final column has value
            break

    start_cell_label = sheet.rows[start_row_idx][start_col_idx].column + str(sheet.rows[start_row_idx][start_col_idx].row)
    end_cell_label = sheet.rows[end_row_idx][end_col_idx].column + str(sheet.rows[end_row_idx][end_col_idx].row)

    num_cols = end_col_idx - start_col_idx + 1
    num_rows = end_row_idx - start_row_idx + 1

    # Trim sheet down to just the range we care about and store this in a tuple
    table_tuple = tuple(sheet[start_cell_label:end_cell_label])

    # Print the name of the table file that is being created this iteration and the excel cells being used to create it
    print('    ' + sheet_name + '.tex    ' + sheet.rows[start_row_idx][start_col_idx].column + str(sheet.rows[start_row_idx][start_col_idx].row) + ':'
          + sheet.rows[end_row_idx][end_col_idx].column + str(sheet.rows[end_row_idx][end_col_idx].row))

    # Find any merged cells within this particular worksheet
    merged_details_list = e2l.get_merged_cells(sheet)

    # Apply an offset for the fact that our table might not start in A1, so the row and column references might be off
    merged_details_list[0] = [x - start_row_idx for x in merged_details_list[0]]  # Start_row
    merged_details_list[1] = [x - start_col_idx for x in merged_details_list[1]]  # Start_col
    merged_details_list[2] = [x - start_row_idx for x in merged_details_list[2]]  # end_row
    merged_details_list[3] = [x - start_col_idx for x in merged_details_list[3]]  # end_col

    # Create .tex output file we will write to
    file = open(output_dir + sheet_name + '.tex', 'w')  #

    # If the user requested the booktabs options, add a reminder (as a LaTeX comment) to the top of the table that the
    # user will need to load up the package in the preamble of their file.
    if usr_settings['booktabs']:
        file.write('% Note: make sure \\usepackage{booktabs} is included in the preamble \n')



    if usr_settings['includetabular']:  # Create tabular preamble

        col_align_str = ""  # Preallocate string

        for colnum in range(0, num_cols):  # For each column, create vertical dividers and aligns

            # Create column to analyze from the table
            col2a = e2l.create_column(table_tuple, colnum)

            if e2l.check_for_vline(col2a, 'left'):  # check to see if there is a vline left of column
                col_align_str += '|'

            # Choose the alignment (l,c,r) of the column based on the majority of alignments in the column's cells
            col_align_str += e2l.pick_col_text_alignment(col2a)

            if e2l.check_for_vline(col2a, 'right'):  # check to see if there is a vline right of column
                col_align_str += '|'

        begin_str = "\\begin{tabular}{" + str(col_align_str) + "} \n"

        file.write(begin_str)




    for row_num in range(0, num_rows):  # For each row in the table's body


        # Check to see if row contains any multicolumns/rows

        # Generate list of True/False values to see if they match the row
        elem_picker = [True if item in [row_num] else False for item in merged_details_list[0]]

        # Pick out the column number and mutlicolumn/row details corresponding to this row
        merge_start_cols = list(compress(merged_details_list[1], elem_picker))
        merge_end_cols = list(compress(merged_details_list[3], elem_picker))
        merge_match_det = list(compress(merged_details_list[4], elem_picker))


        # If there is a horizontal rule across all cells at the top, add it to the table
        hrule_str = e2l.create_horzrule_code(table_tuple[row_num], 'top', usr_settings, top_row=False, bottom_row=False)

        # If using booktabs, and this is the first row, use toprule rather than midrule
        if (row_num == 0) & usr_settings['booktabs']:
            hrule_str = hrule_str.replace('\\midrule', '\\toprule')

        file.write(hrule_str)


        # Get string of rows contents
        str_2_write = e2l.tupple2latexstring(table_tuple[row_num], usr_settings, [merge_start_cols, merge_end_cols, merge_match_det])

        # This is now done inside of tupple2latexstring to avoid rounding color hex numbers
        # If we need to round numbers in the row, do so
        #if usr_settings['roundtodp']:
        #    str_2_write = e2l.round_num_in_str(str_2_write, usr_settings['numdp'])

        # Write row string to file
        file.write(str_2_write)



        hrule_str = e2l.create_horzrule_code(table_tuple[row_num], 'bottom', usr_settings, top_row=False, bottom_row=False)

        if (row_num == num_rows - 1) & usr_settings['booktabs']:
            hrule_str = hrule_str.replace('\\midrule', '\\bottomrule')

        file.write(hrule_str)



    if usr_settings['includetabular']:
        # Close table environment
        file.write("\\end{tabular}")

    file.close()  # Close off the current .tex file (completing the creation of the table code)

print(' ')
print('Code has completed running')
