# Excel2LaTeXviaPython #

Creating good looking tables within LaTeX is a time consuming process. Software like Excel is much better suited to producing nicely formatted tables quickly and easily. The python code in this repository allows you to use Excel to design and format your tables, and then produce the LaTeX code to replicate those table in your LaTeX document by running a single function.

There already exists an Excel add-in ([Excel2LaTeX](https://www.ctan.org/tex-archive/support/excel2latex/ )) that can produce LaTeX tables from Excel files. Excel2LaTeXviaPython has the following advantages over Excel2LaTeX:

1. Can run on any platform. Users have reported Runtime problems using the Excel2LaTeX add-in with recent versions of Excel for the Mac (It runs fine on Windows). 
2. This code features more automation. Running this code one will automatically update all the tables you want to use in your paper. You do not need to convert each table one at a time.


This code is developed for Python 3.

## Contact details ##

* Developed by: Michael Kirker

* Email: <mkirker@uchicago.edu>

* Website: [michaelkirker.net](http://michaelkirker.net)

* Git repository: [https://github.com/michaelkirker/excel2latexviapython](https://github.com/michaelkirker/excel2latexviapython)




## Repository structure ##

- /Example/
  - Folder containing the an excel file featuring example tables you can run the python code on.
- e2lvp.py
  - Main file containing the functions for Excel2LaTeXviaPython.
- example_excel2latexviapython.py
  - File to run to demonstrate the function
- gui_excel2latexviapython.py
  - Script to launch an optional GUI to run the main function


## Creating the Excel File Input

The main input required for this code is an excel workbook (a file ending in .xlsx). When creating your excel file, the requirements for use in this function are as follows

1. Each worksheet (tab) contains only one table
2. The name of each worksheet (tab) is named to match the names of the .tex files you wanted produced as output

In this repository you can find  `/Example/example_tables.xlsx` which provides an example excel workbook filled with example tables.

### Creating Tables on Individual Worksheets ###

On each worksheet you can format your table using Excel's inbuilt tools to format text, merge cells, draw boundary lines and boxes, etc. You can also include latex math code if you wrap it in the `$$` environment (e.g. `$x>2$`). Not that if you include the ampersand (&) character in your table, you will need to write `\&` as otherwise the LaTeX interpreter will think you have an extra cell break in your table and the code will not compile.

The most common formatting features include:

- Text formatting
  - Bold
  - italicized
  - color*
- Colored background fills for individual cells*
- Horizontal and vertical rules**
  - Horizontal lines may span the entire width of the table, or only certain columns
- Horizontal alignment of text in columns (left/center/right)
- Merged cells 


[*] Currently, only the "standard colors" or "more colors" options in Excel return colors in the LaTeX code, and the "Theme colors" in the dropdown excel menus do not work. The inbuilt theme colors do not return a nice color hex code when parsed. So it is currently not possible to convert these to a color that LaTeX could interpret. The way these cases are currently handled is to ignore the Theme color choice and return either black text or a plain background.

[**] Note that currently if a cell has a bottom rule, and the cell below has a top rule, you will get two horizontal lines in your LaTeX table code. However, in your excel file it will look like there is only one line.


## How to run this code ##

Note that you can have the Excel workbook open on your computer when running the python code. However, make sure you save any changes before running the function as the function will read in the last saved version of the file, not the version that is currently open (which may differ).

### Option 1: Standalone function

Include the file `e2lvp.py` in the directory of your project

The functions can be imported into your code using:

​     `import e2lvp`

It appears that you can have the Excel Workbook open when running the code and it doesnt affect things.

`e2lvp.excel2latexviapython(excel_filename, set_output_dir, booktabs=True, includetabular=True, roundtodp=True, numdp=3, makepdf=True)`

The inputs into the `excel2latexviapython` function are as follows:

- `input_excel_filename` [string] The full path (including extension) of the excel file.
- `output_dir` [string] The full path of the folder the produced TeX files should be outputted to.
- `booktabs` - [True/False]  Should the package booktabs be used to create nice looking horizontal rules? If False, standard \hrule is used.


- `includetabular` [True/False] Should the code output the tabular environment environment around each table, or just output the individual rows of each table?
- `roundtodp` [True/False] Apply rounding to all numbers in the table?
- `numdp` [scalar]` How many decimal places to round to if `roundtodp=True`
- `makepdf` [True/False]  Should the function also create simple LaTeX and PDF file aggregating all the tables? This is useful if you want to have one place to quickly check all your tables to make sure the output is correct

The file `example_excel2latexviapython.py` found in the main directory of the repository provides an example of using this function.

### Option 2: GUI

Running the file `gui_excel2latexviapython.py` to launch the GUI interface to the function. From there you can directly select all the inputs to the function. The window remains open after executing so you can easily re-run the code with the same inputs if you make any changes to the tables within the Excel file.


### Using the outputted LaTeX code ###

In you LaTeX document, use the `\input{FILE_PATH/FILE}` command to have LaTeX automatically import the produced table coded when you compile the document. Note that if the files are in the same folder as your main TeX file, you only need to use `\input{FILE}`

Alternatively, you can copy and paste the LaTeX code from each output file into you main document. Using the `\input{}` method has the advantage of always pulling the most recent version of the tables when you compile your document. This makes it relatively easy to update your excel tables with new numbers and quickly get it into your document (just by re-running the python code).


## Example ##

The folder `/Example/` provides an Excel file (`example_tables.xlsx`) that demonstrates the features of the code by presenting several differently formatted tables. Run the python code `example_excel2latexviapython.py` imports this Excel workbook and outputs the results back to the `/Example/` folder.