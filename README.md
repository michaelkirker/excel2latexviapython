# Excel2LaTeXviaPython #
Python code to create LaTeX tables from Microsoft Excel files.


Written for Python 3.

The code is currently being developed in a Jupyter Notebook. But it will eventually be moved to a stand-alone file, or even its own executable to make things easy.

## Contact details ##

Michael Kirker

* Email: <mkirker@uchicago.edu>

* Website: [michaelkirker.net](http://michaelkirker.net)

* Git repository: [https://github.com/michaelkirker/excel2latexviapython](https://github.com/michaelkirker/excel2latexviapython)


## Features that can currently be included ##

- Bold text in cells
- Italicized text in cells
- Horizontal rules that go across the entire width of the table
- Vertical rules that go the entire height of the table
- Allow horizontal alignment of each column (based on a majority of the rows)
- Displays the same number of d.p. if the user uses excel to round off a number

Note that currently if a cell has a bottom rule, and the cell below has a top rule, you will get two horizontal lines in your LaTeX table code.

Thinking about having the user put all their tables in one excel file, one per sheet, and the code looping over all sheets. The name of the file saved would be read from the sheet's name.

## List of features to add ##
The following features are on the to-do list for the near term.


- Control for merged cells
- Different line thicknesses
- Ignore empty cells for hrule etc choices
	- Now will make this handle rules across only part of the table
- background & text color
- Move to pure python script
- Make a standalone program


## Notes on running the code ##

If you use a program to generate the excel file (such as Stata), it might be possible that each cell has the following format `="*"` rather than just `*`. This code will currently print the equals sign and double quotation marks. The easiest way to get around this is to highlight all the cells in excel, copy the cells, and then paste the values back into the table.

If you want to include special LaTeX characters, you need to remember to incude the `\` (backslash) command before the character. Also, if you want to include mathematical notation, you need to enclose the code in `$` (dollar signs).


It appears that you can have the Excel Workbook open when running the code and it doesnt affect things.


Cant leave empty looking cells elsewhere




## Example ##