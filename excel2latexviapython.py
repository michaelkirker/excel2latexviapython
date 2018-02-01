# Load in required packages
import openpyxl  # Package for reading excel files (.xlsx) into Python
from itertools import compress
from sub_functions import e2l  # Model functions to simulate the LBH model

# EXCEL TO LATEX VIA PYTHON
########################################################################################################################
#
# This code takes an excel file with tables in each worksheet and outputs the TeX code for each table as a separate file
# ready to be imported into a LaTeX document.
#
#
# USER INPUT AND SETTINGS
# ======================================================================================================================
#
# Input the full path and file name of the excel file containing your tables
excel_filename = 'D:/Users/Kirker/Google Drive/Git Repositories/excel2latexviapython/Example/' +  \
                       'example_tables.xlsx'
#
# Select the directory/folder to save the resulting TeX files to
set_output_dir = 'D:/Users/Kirker/Google Drive/Git Repositories/excel2latexviapython/Example/'
#
# Define user settings in a dictionary
# ====================================
#
# Current options:
#
#    booktabs: True/False
#        True = use the booktabs package functions to make prettier horizontal lines 
#        False = Use standard \hlines
#
#   includetabular: True/False
#        Should the code include the tabular environment code around the table (\begin{tabular}, \end{tabular}),
#        or just return the table rows.
#
#   roundtodp: True/False
#       Should all numbers in the table be rounded to a set number of d.p.?
#
#   numdp: scalar
#       Define how many d.p. to round numbers to if roundtpdp=True
input_usr_settings = {'booktabs': True, 'includetabular': True, 'roundtodp': True, 'numdp': 3}
#
# End of user input
# ======================================================================================================================


def excel2latexviapython(input_excel_filename, output_dir, booktabs=True, includetabular=True, roundtodp=True, numdp=3):
    usr_settings = {'booktabs': booktabs, 'includetabular': includetabular, 'roundtodp': roundtodp, 'numdp': numdp}

    # PREAMBLE
    # ==================================================================================================================

    # Print output so user can follow progress.
    print('EXCEL 2 LATEX VIA PYTHON')
    print('Creates .TeX table files from excel file.')
    print('\nSource file:      ' + input_excel_filename)

    # Load in the Excel workbook/file
    workbook = openpyxl.load_workbook(filename=input_excel_filename, data_only=True)

    print('Output directory: ' + output_dir + '\n')
    print('User settings:')
    print('    booktabs: ' + str(usr_settings['booktabs']))
    print('    includetabular: ' + str(usr_settings['includetabular']))
    print('    roundtodp: ' + str(usr_settings['roundtodp']))
    print('    numdp: ' + str(usr_settings['numdp']))
    print('\n')
    print('Starting to create TeX tables (output name, table location within excel sheet')

    # MAIN CODE
    # ==================================================================================================================

    for sheet_name in workbook.get_sheet_names():  # Loop over every worksheet/tab within the input workbook

        # Get the worksheet object for this iteration of the loop
        sheet = workbook[sheet_name]

        # The table within the sheet may not start in cell A1. So find the location of the upper-left and bottom-right
        # corner cells of the table within the sheet
        start_row_idx, start_col_idx, end_row_idx, end_col_idx = e2l.get_table_dimensions(sheet)

        # Get the excel cell labels of the upper-left and bottom-right cells of the table
        start_cell_label = list(sheet.rows)[start_row_idx][start_col_idx].column + str(list(sheet.rows)[start_row_idx]
                                                                                       [start_col_idx].row)
        end_cell_label = list(sheet.rows)[end_row_idx][end_col_idx].column + str(list(sheet.rows)[end_row_idx]
                                                                                 [end_col_idx].row)

        # Get the number of columns and rows in the table
        num_cols = end_col_idx - start_col_idx + 1
        num_rows = end_row_idx - start_row_idx + 1

        # Trim sheet object down to just the range we care about and store this in a tuple
        table_tuple = tuple(sheet[start_cell_label:end_cell_label])

        # Print to the terminal the name of the table file that is being created this iteration and the excel cells
        # being used to create it
        print('    ' + sheet_name + '.tex    ' + list(sheet.rows)[start_row_idx][start_col_idx].column +
              str(list(sheet.rows)[start_row_idx][start_col_idx].row) + ':'
              + list(sheet.rows)[end_row_idx][end_col_idx].column + str(list(sheet.rows)[end_row_idx][end_col_idx].row))

        # Create .tex file we will write to
        file = open(output_dir + sheet_name + '.tex', 'w')

        # Preamble of the individual table
        # --------------------------------

        # If the user requested the booktabs options, add a reminder (as a LaTeX comment) to the top of the table that
        # the user will need to load up the package in the preamble of their file.
        if usr_settings['booktabs']:
            file.write('% Note: make sure \\usepackage{booktabs} is included in the preamble \n')

        file.write('% Note: If your table contains colors, make sure \\usepackage[table]{xcolor} is included in the '
                   'preamble \n')

        # If the user wants the table rows wrapped in the tabular environment, write the start of the begin environment
        # command to the output tex file
        if usr_settings['includetabular']:

            col_align_str = "\\begin{tabular}{"  # Preallocate string

            # For each column of the table, append to "col_align_str" any vertical dividers and alignment code for the
            # column
            for colnum in range(0, num_cols):

                # Create column to analyze from the table
                col2a = e2l.create_column(table_tuple, colnum)

                # check to see if there is a vline left of column
                if e2l.check_for_vline(col2a, 'left'):
                    col_align_str += '|'

                # Choose the alignment (l,c,r) of the column based on the majority of alignments in the column's cells
                col_align_str += e2l.pick_col_text_alignment(col2a)

                # check to see if there is a vline right of column
                if e2l.check_for_vline(col2a, 'right'):
                    col_align_str += '|'

            # Create code to write to tex output file
            begin_str = str(col_align_str) + "} \n"

            # Write the \begin{tabular}{*} code to the tex file
            file.write(begin_str)

        # Body of the individual table
        # ----------------------------

        # Find any merged cells within this particular worksheet
        merged_details_list = e2l.get_merged_cells(sheet)

        # Adjust the merged_details_list values for the fact that the table might not start in cell A1
        merged_details_list[0] = [x - start_row_idx for x in merged_details_list[0]]  # start_row
        merged_details_list[1] = [x - start_col_idx for x in merged_details_list[1]]  # start_col
        merged_details_list[2] = [x - start_row_idx for x in merged_details_list[2]]  # end_row
        merged_details_list[3] = [x - start_col_idx for x in merged_details_list[3]]  # end_col

        # For each row in the table's body create a string containing the tex code for that row and write to the output
        # file
        for row_num in range(0, num_rows):

            # Generate list of True/False values to see if they match the row
            elem_picker = [True if item in [row_num] else False for item in merged_details_list[0]]

            # Pick out the column number and mutlicolumn/row details corresponding to this row
            merge_start_cols = list(compress(merged_details_list[1], elem_picker))
            merge_end_cols = list(compress(merged_details_list[3], elem_picker))
            merge_match_det = list(compress(merged_details_list[4], elem_picker))

            # If there is a horizontal rule across all cells at the top, add it to the table
            hrule_str = e2l.create_horzrule_code(table_tuple[row_num], 'top', merge_start_cols, merge_end_cols,
                                                 usr_settings)

            # If user requested booktabs, and this is the first row, use toprule rather than midrule
            if (row_num == 0) & usr_settings['booktabs']:
                hrule_str = hrule_str.replace('\\midrule', '\\toprule')

            file.write(hrule_str)

            # Get string of rows contents
            str_2_write = e2l.tupple2latexstring(table_tuple[row_num], usr_settings, [merge_start_cols, merge_end_cols,
                                                                                      merge_match_det])

            # Write row string to file
            file.write(str_2_write)

            # Add any horizontal rule below the row
            hrule_str = e2l.create_horzrule_code(table_tuple[row_num], 'bottom', merge_start_cols, merge_end_cols,
                                                 usr_settings)

            # If user requested booktabs, and this is the final row, use bottomrule rather than midrule
            if (row_num == num_rows - 1) & usr_settings['booktabs']:
                hrule_str = hrule_str.replace('\\midrule', '\\bottomrule')

            file.write(hrule_str)

        # Postamble of the individual table
        # ---------------------------------
        if usr_settings['includetabular']:
            # User has requested tabular environment wrapped around the table rows, so end the table
            file.write("\\end{tabular}")

        file.close()  # Close off the current .tex file (completing the creation of the table code)

    print('\nCode has completed running')


# Run the function
excel2latexviapython(excel_filename, set_output_dir, booktabs=input_usr_settings['booktabs'],
                     includetabular=input_usr_settings['includetabular'], roundtodp=input_usr_settings['roundtodp'],
                     numdp=input_usr_settings['numdp'])
