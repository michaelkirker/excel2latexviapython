# Excel2LaTeXviaPython #

ALPHA RELEASE - NOT READY FOR THE PUBLIC YET

Creating good looking tables within LaTeX is a time consuming process. Software like Excel is much better suited to producing nicely formatted tables quickly and easily. The python code in this repository allows you to use Excel to design and format your tables, and then produce the LaTeX code to replicate those table in your LaTeX document.

There already exists an Excel add-in ([Excel2LaTeX](https://www.ctan.org/tex-archive/support/excel2latex/ )) that can produce LaTeX tables from Excel files. My approach has the following advantages over this previous approach:

1. Can run on any platform. Users have reported Runtime problems using the Excel2LaTeX add-in with recent versions of Excel for the Mac (It runs fine on Windows). 
2. This code features more automation. Running this code one will automatically update all the tables you want to use in your paper. You do not need to convert each table one at a time.


This code is currently developed for Python 3, and runs using a Jupyter Notebook. I plan on iterating on the code to improve it over time as I encounter the need for different tables in my own research work. Suggestions for new features and other feedback can be sent to me using the contact details provided below.



## Contact details ##

* Developed by: Michael Kirker

* Email: <mkirker@uchicago.edu>

* Website: [michaelkirker.net](http://michaelkirker.net)

* Git repository: [https://github.com/michaelkirker/excel2latexviapython](https://github.com/michaelkirker/excel2latexviapython)




## Repository structure ##

- /Example/
	- Folder containing the an excel file featuring example tables you can run the python code on
- /sub_functions/
	- functions written specifically for this code
- excel2latexviapython.py
	- Main file to run


## Features that can currently be included ##

The following list are the main formatting features that can be included in the tables.

- Text formatting: Bold, italicized, color (requires xcolor package)
- Horizontal and vertical rules
- Booktabs rules (for better looking tables)
- Horizontal alignment of text in columns (based on a majority of the rows)
- Merged cells (with their own horizontal alignment choice)
- Automatic rounding of values to set number of d.p.


Note that currently if a cell has a bottom rule, and the cell below has a top rule, you will get two horizontal lines in your LaTeX table code. However, in your excel file it will look like there is only one line.

Thinking about having the user put all their tables in one excel file, one per sheet, and the code looping over all sheets. The name of the file saved would be read from the sheet's name.

## List of features to add ##
The following features are on the to-do list for the near term.


- background color
- Made the code a stand-alone program
	- Move from a Jupyter notebook to a standard Python script file
	- Look at possibly doing a simple GUI to take the user's inputs
- Allow for Multirows (vertically merged cells)
- Some sort of log file for when the code runs into problems with your excel file.
	- Error checking of the file to make sure the code will run properly
- Improve the formatting/readability of the produced LaTeX code by automatically adjusting the number of tabs each used to separate the columns in each row.
- Automatically adjust for special characters (like greater than sign).

## Notes on running the code ##

Often when external programs (such as Stata) are used to generate the excel file, the content of each cell will be formatted as `="CONTENT"` rather than just `CONTENT`.

If you use a program to generate the excel file (such as Stata), it might be possible that each cell has the following format `="*"` rather than just `*`. This code will currently print the equals sign and double quotation marks. The easiest way to get around this is to highlight all the cells in excel, copy the cells, and then paste the values back into the table.

If you want to include special LaTeX characters, you need to remember to include the `\` (backslash) command before the character. Also, if you want to include mathematical notation, you need to enclose the code in `$` (dollar signs).


It appears that you can have the Excel Workbook open when running the code and it doesnt affect things.


Cant leave empty looking cells elsewhere


## How to run this code ##


### Step 1) Format your tables in excel ###
Create an Excel `.xlsx` file (Workbook) and input all your tables into the Workbook. Check that you have done the following things:

- Each table is on a separate worksheet (tab)
- The worksheets are name matches the file names you want to use for each table.
- Apply the formatting to your tables (bold/italicized text, horizontal or vertical rules, merged cells etc)



### Step 2) Set the settings inside the python file ###

At the start of the python file `FILENAME`, the user need to set the following:

- `input_excel_filename` - The full path (including extension) of the excel file.
- `output_dir` - The full path of the folder the produced TeX files should be outputted to.
- `usr_settings` - Python dictionary of options to use when processing the table.


The following user settings are available inside of `usr_settings`:

- booktabs - `True/False` 
	- If `True`, \toprule, \bottomrule, \midrule will be used rather than \hrule (requires the booktabs package to be used in your LaTeX document.
- includetabular - `True/False`
	- Should the code output the tabular environment environment around each table, or just output the individual rows of each table
- roundtodp - `True/False`
	- Apply rounding to all numbers in the table
- numdp - `scalar`
	- How many decimal places to round to if roundtodp=True


### Step 3) Run the python code ###


### Step 4) Input the produced LaTeX table code(s) into your LaTeX document ###


**Option 1:** Copy and paste the produced LaTex Code from the output files into your document

**Option 2:** In you LaTeX document, use the `\input{FILE_PATH/FILE}` command to have LaTeX automatically import the produced table coded when you compile the document.

Option 2 is useful when you will be updating your tables in the future and you want the LaTeX document to always grab the latest version of the table each time it is compiled. 


## Example ##

The folder `/Example/` provides an Excel file (`example_tables.xlsx`) that demonstrates the features of the code by presenting several differently formatted tables. Run the python code using this excel file as the input and output directory as the `/Example/` folder. Then compile the file `output_all_tables.tex` to see all the tables in a single LaTeX document.