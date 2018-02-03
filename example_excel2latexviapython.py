# EXCEL TO LATEX VIA PYTHON
########################################################################################################################
#
# This code takes an excel file with tables in each worksheet and outputs the TeX code for each table as a separate file
# ready to be imported into a LaTeX document.
#
import e2lvp  # Model functions to simulate the LBH model


# USER INPUT AND SETTINGS
# ======================================================================================================================

# Input the full path and file name of the excel file containing your tables
excel_filename = 'D:/Users/Kirker/Google Drive/Git Repositories/excel2latexviapython/Example/' +  \
                       'example_tables.xlsx'

# Select the directory/folder to save the resulting TeX files to
set_output_dir = 'D:/Users/Kirker/Google Drive/Git Repositories/excel2latexviapython/Example/'


# Details of optional inputs:
# ===========================
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
#
#   makepdf: True/False
#       Make a PDF document containing all the tables. Useful for checking output quickly. Note, requires
#       includetabular=True

# Run the function
e2lvp.excel2latexviapython(excel_filename, set_output_dir, booktabs=True, includetabular=True, roundtodp=True, 
                           numdp=3, makepdf=True)
