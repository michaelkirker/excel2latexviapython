# Excel2LaTeXviaPython #

ALPHA RELEASE - NOT READY FOR THE PUBLIC YET

Creating good looking tables within LaTeX is a time consuming process. Software like Excel is much better suited to producing nicely formatted tables quickly and easily. The python code in this repository allows you to use Excel to design and format your tables, and then produce the LaTeX code to replicate those table in your LaTeX document.

There already exists an Excel add-in ([Excel2LaTeX](https://www.ctan.org/tex-archive/support/excel2latex/) )that can produce LaTeX tables from Excel files. My approach has the following advantages over this previous approach:

1. Can run on any platform. Users have reported Runtime problems using the Excel2LaTeX add-in with recent versions of Excel for the Mac (It runs fine on Windows). Python is able to run on Windows and Mac equally as well. 
2. This code features more automation. Running this code one can automatically update all the tables you want to use in your paper. You do not need to convert each table one at a time
3. More formatting options for tables. This code can (will) feature more formatting options allowing you more creative control over the final look of your tables.

This code is currently developed for Python 3, and runs using a Jupyter Notebook. I plan on iterating on the code to improve it over time as I encounter the need for different tables in my own research work. Suggestions for new features and other feedback can be sent to me using the contact details provided below.



## Contact details ##

* Developed by: Michael Kirker

* Email: <mkirker@uchicago.edu>

* Website: [michaelkirker.net](http://michaelkirker.net)

* Git repository: [https://github.com/michaelkirker/excel2latexviapython](https://github.com/michaelkirker/excel2latexviapython)


## Features that can currently be included ##

The following list are the main formatting features that can be included in the tables.

- Bold text in cells
- Italicized text in cells
- Horizontal and vertical rules
- Booktabs rules (for better looking tables)
- Horizontal alignment of text in columns (based on a majority of the rows)
- Merged cells (with their own horizontal alignment choice)
- Rounding of individual cells to different numbers of decimal places.

Note that currently if a cell has a bottom rule, and the cell below has a top rule, you will get two horizontal lines in your LaTeX table code. However, in your excel file it will look like there is only one line.

Thinking about having the user put all their tables in one excel file, one per sheet, and the code looping over all sheets. The name of the file saved would be read from the sheet's name.

## List of features to add ##
The following features are on the to-do list for the near term.


- background & text color
- Made the code a stand-alone program
	- Move from a Jupyter notebook to a standard Python script file
	- Look at possibly doing a simple GUI to take the user's inputs
- Allow for Multirows (vertically merged cells)
- Some sort of log file for when the code runs into problems with your excel file.
	- Error checking of the file to make sure the code will run properly
- Improve the formatting/readability of the produced LaTeX code by automatically adjusting the number of tabs each used to separate the columns in each row.
- Automatically adjust for special characters (like greater than sign).

## Notes on running the code ##

If you use a program to generate the excel file (such as Stata), it might be possible that each cell has the following format `="*"` rather than just `*`. This code will currently print the equals sign and double quotation marks. The easiest way to get around this is to highlight all the cells in excel, copy the cells, and then paste the values back into the table.

If you want to include special LaTeX characters, you need to remember to include the `\` (backslash) command before the character. Also, if you want to include mathematical notation, you need to enclose the code in `$` (dollar signs).


It appears that you can have the Excel Workbook open when running the code and it doesnt affect things.


Cant leave empty looking cells elsewhere


## How to run this code ##


### Step 1) Format your tables in excel ###
Create an Exel `.xlsx` file/Workbook and input all your tables into the Workbook. Check that you have done the following things:

- Each table is on a separate worksheet (tab)
- The worksheets are renamed to the names of the file names you want to use for each table.
- Apply the formatting to your tables (bold/italicized text, horizontal or vertical rules, merged cells etc)



### Step 2) Set the settings inside the Jupyter notebook ###

Near the top of the Jupyter notebook containing the Python code is a cell for user input. Set the input file, output directory, and user settings options.

Run the notebook code.

### Step 3) Input the produced LaTeX code into your file ###


**Option 1:** Copy and paste the produced LaTex Code from the output files into your document

**Option 2:** In you LaTeX document, use the `\input{FILE LOCATION}` command to have LaTeX automatically import the produced table coded when you compile the document. 


## Example ##

The folder `/Example/` provides an Excel file (`example_tables.xlsx') that demonstrates the features of the code by presenting several differently formatted tables.