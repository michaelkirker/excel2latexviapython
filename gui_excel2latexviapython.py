from tkinter import *
from tkinter import filedialog
from sub_functions import e2lvp  # Model functions to simulate the LBH model


window = Tk()

window.title("GUI for Excel2LaTeXviaPython")


lbl = Label(window, text="This GUI allows you to run the Excel2LaTeXviaPython script")
lbl.grid(column=0, row=0)


# Get the input file
########################################################################################################################
input_row = 2

# Text Description
lbl_input_file = Label(window, text="Select the excel file:")
lbl_input_file.grid(column=0, row=input_row)

# Pre-allocate file name text block
lbl_input_file_name = Label(window, text="")
lbl_input_file_name.grid(column=2, row=input_row)


# Define open excel file button
def clicked():  # What happens when button is clicked
    file = filedialog.askopenfilename()
    lbl_input_file_name.configure(text="File = " + file)


btn = Button(window, text="Open", command=clicked)
btn.grid(column=1, row=input_row)

# Get the output folder
########################################################################################################################

output_row = 3

# Description
lbl_input_file = Label(window, text="Select the output folder:")
lbl_input_file.grid(column=0, row=output_row)


# Pre-allocate folder name
lbl_output_folder_name = Label(window, text="")
lbl_output_folder_name.grid(column=2, row=output_row)


# Define open excel file button
def clicked_out():  # What happens when button is clicked
    file = filedialog.askdirectory()
    lbl_output_folder_name.configure(text="Folder = " + file)


btn_out_folder = Button(window, text="Open", command=clicked_out)
btn_out_folder.grid(column=1, row=output_row)


# Code options
########################################################################################################################

row_options_base = output_row+2

lbl_options = Label(window, text="Specify the options you wish to use")
lbl_options.grid(column=0, row=row_options_base)

chk_state_te = BooleanVar()
chk_state_te.set(True)  # set check state

chk_te = Checkbutton(window, text='Include Tabular environment', var=chk_state_te, onvalue=True, offvalue=False)
chk_te.grid(column=0, row=row_options_base+1)

chk_state_bt = BooleanVar()
chk_state_bt.set(True)  # set check state

chk_bt = Checkbutton(window, text='Booktabs', var=chk_state_bt, onvalue=True, offvalue=False)
chk_bt.grid(column=0, row=row_options_base+2)

chk_state_rnd = BooleanVar()
chk_state_rnd.set(True)  # set check state

chk_rnd = Checkbutton(window, text='Round to specific number of d.p.', var=chk_state_rnd, onvalue=True, offvalue=False)
chk_rnd.grid(column=0, row=row_options_base+3)


txt_numdp = Entry(window, width=5)
txt_numdp.grid(column=0, row=row_options_base+4)
lbl = Label(window, text="Number of decimal places to round to")
lbl.grid(column=1, row=row_options_base+4)
txt_numdp.insert(10, "3")



chk_state_pdf = BooleanVar()
chk_state_pdf.set(False)  # set check state

chk_pdf = Checkbutton(window, text='Make PDF of tables', var=chk_state_pdf, onvalue=True, offvalue=False)
chk_pdf.grid(column=0, row=row_options_base+5)

# Output file:
########################################################################################################################
execute_row = row_options_base+5+3

lbl_execute = Label(window, text="Run Excel2LaTeXviaPython")
lbl_execute.grid(column=0, row=execute_row)


def clicked_execute():  # What happens when button is clicked
    e2lvp.excel2latexviapython(lbl_input_file_name["text"][7:], lbl_output_folder_name["text"][9:], booktabs=chk_state_bt.get(),
                         includetabular=chk_state_te.get(), roundtodp=chk_state_rnd.get(), numdp=int(txt_numdp.get()),
                         makepdf=chk_state_pdf.get())


btn = Button(window, text="Execute", command=clicked_execute)

btn.grid(column=0, row=execute_row+1)

window.mainloop()
