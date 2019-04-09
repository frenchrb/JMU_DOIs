from tkinter import *
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import ttk
import os
import subprocess
import Bepress2DataCite
import MintDOIs


class NewStdOut():
    def __init__(self, oldStdOut, textArea):
        self.oldStdOut = oldStdOut
        self.textArea = textArea

    def write(self, x):
        # self.oldStdOut.write(x)
        self.textArea.insert(END, x)
        self.textArea.see(END)
        self.textArea.update_idletasks()
    
    def flush(self):
        self.oldStdOut.flush()


def openXML():
    global openpath_xml
    openpath_xml = filedialog.askopenfilename(title='Select file',
                                              filetypes=(("XML Spreadsheet 2003","*.xml"), ("All files","*.*")))
    # print('openpath_xml: ' + openpath_xml)


def openXLS():
    global openpath_xls
    openpath_xls = filedialog.askopenfilename(title='Select file',
                                              filetypes=(("Excel 97-2003 Workbook","*.xls"), ("All files","*.*")))
    # print('openpath_xls: ' + openpath_xls)


def clear_openpath_xls(event):
    global openpath_xls
    try:
        del openpath_xls
    except NameError:
        return


def confirmProduction():
    production_check = simpledialog.askstring('Warning: Production DOIs', 'You have selected Production DOIs. Production DOIs cannot be deleted.\n\nTo confirm your selection, please type "Yes, I\'m sure!" in the box below and click OK.\n')
    if production_check == 'Yes, I\'m sure!':
        print('Production DOIs selection confirmed.')
        print()
        return True
    else:
        return False


def runBepress2DataCite():
    if not setname.get():
        print('Please enter a collection setname.')
        return
    try:
        openpath_xls
    except NameError:
        print('Please select a spreadsheet file.')
        return
    if tab1_radiobutton_production_var.get():
        if not confirmProduction():
            # Reset radiobuttons to Draft DOIs
            tab1_radiobutton_production.deselect()
            tab1_radiobutton_draft.select()
            return
    
    args = []
    args.append(setname.get())
    args.append(openpath_xls)
    if tab1_radiobutton_production_var.get():
        args.append('--production')
    Bepress2DataCite.main(args)
    
    # Reset radiobuttons to Draft DOIs
    tab1_radiobutton_production.deselect()
    tab1_radiobutton_draft.select()
    # root.destroy()


def runMintDOIs():
    try:
        openpath_xls
    except NameError:
        print('Please select a spreadsheet file.')
        return
    if tab2_radiobutton_production_var.get():
        if not confirmProduction():
            # Reset radiobuttons to Draft DOIs
            tab2_radiobutton_production.deselect()
            tab2_radiobutton_draft.select()
            return
    
    args = []
    args.append(openpath_xls)
    if tab2_radiobutton_production_var.get():
        args.append('--production')
    MintDOIs.main(args)
    
    # Reset radiobuttons to Draft DOIs
    tab2_radiobutton_production.deselect()
    tab2_radiobutton_draft.select()
    # root.destroy()


def open_folder():
    subprocess.Popen('explorer ' + os.getcwd())


def main():
    root = Tk()
    root.wm_title('JMU DOI Minter')
    root.resizable(width=FALSE, height=FALSE)
    
    # Tabs
    notebook = ttk.Notebook(root)
    tab1 = ttk.Frame(notebook)
    tab2 = ttk.Frame(notebook)
    notebook.add(tab1, text='1. Create Metadata')
    notebook.add(tab2, text='2. Mint DOIs')
    notebook.grid(row=0, column=0, sticky=NW)
    notebook.bind('<<NotebookTabChanged>>', clear_openpath_xls)

    # Tab 1
    tab1_frame = Frame(tab1)
    tab1_frame.grid()
    
    # Create tab 1 widgets
    label_entrybox_setname = Label(tab1_frame, text='Enter bepress collection setname (e.g., "diss201019")')
    global setname
    setname = StringVar()
    entrybox_setname = Entry(tab1_frame, textvariable=setname, width=20)
    label_button_open1 = Label(tab1_frame,
                                  text='Select bepress spreadsheet saved as "Excel 97-2003 Workbook (*.xls)"')
    # textarea_open_xml = Text(tab1_frame, height=1, width=40)
    button_open1 = Button(tab1_frame, text='Browse...', height=1, width=10, command=openXLS)
    tab1_radiobutton_frame = Frame(tab1_frame, relief='groove', borderwidth=2)
    global tab1_radiobutton_production_var
    tab1_radiobutton_production_var = IntVar()
    global tab1_radiobutton_draft
    global tab1_radiobutton_production
    tab1_radiobutton_draft = Radiobutton(tab1_radiobutton_frame, text='Draft DOIs',
                                         variable=tab1_radiobutton_production_var, value=False)
    tab1_radiobutton_production = Radiobutton(tab1_radiobutton_frame, text='Production DOIs',
                                              variable=tab1_radiobutton_production_var, value=True)
    button_run = Button(tab1_frame, text='Generate Metadata', height=2, width=25, command=runBepress2DataCite)
    button_open_folder1 = Button(tab1_frame, text='Open folder in Explorer', height=1, width=25, command=open_folder)
    
    # Lay out tab 1 widgets
    label_entrybox_setname.grid(row=1, column=0, sticky=W, padx=5, pady=10)
    entrybox_setname.grid(row=1, column=1, padx=5, pady=10)
    label_button_open1.grid(row=2, column=0, columnspan=2, sticky=W, padx=5, pady=10)
    # textarea_open_xml.grid(row=3, column=0, padx=5, pady=10)
    button_open1.grid(row=3, column=0, sticky=W, padx=5)
    tab1_radiobutton_frame.grid(row=4, column=0, sticky=W, padx=5, pady=10)
    tab1_radiobutton_draft.grid(row=0, column=0, sticky=W)
    tab1_radiobutton_production.grid(row=1, column=0, sticky=W)
    button_run.grid(row=7, column=0, columnspan=2, pady=25)
    button_open_folder1.grid(row=8, column=0, columnspan=2, pady=25)

    # Tab 2
    tab2_frame = Frame(tab2)
    tab2_frame.grid()
    
    # Create tab 2 widgets
    label_button_open2 = Label(tab2_frame,
                                  text='Select bepress spreadsheet saved as "Excel 97-2003 Workbook (*.xls)"')
    # textarea_open_xls = Text(tab2_frame, height=1, width=40)
    button_open2 = Button(tab2_frame, text='Browse...', height=1, width=10, command=openXLS)
    tab2_radiobutton_frame = Frame(tab2_frame, relief='groove', borderwidth=2)
    global tab2_radiobutton_production_var
    tab2_radiobutton_production_var = IntVar()
    global tab2_radiobutton_draft
    global tab2_radiobutton_production
    tab2_radiobutton_draft = Radiobutton(tab2_radiobutton_frame, text='Draft DOIs',
                                         variable=tab2_radiobutton_production_var, value=False)
    tab2_radiobutton_production = Radiobutton(tab2_radiobutton_frame, text='Production DOIs',
                                              variable=tab2_radiobutton_production_var, value=True)
    button_run = Button(tab2_frame, text='Mint DOIs', height=2, width=25, command=runMintDOIs)
    button_open_folder2 = Button(tab2_frame, text='Open folder in Explorer', height=1, width=25, command=open_folder)
    
    # Lay out tab 2 widgets
    label_button_open2.grid(row=2, column=0, columnspan=2, sticky=W, padx=5, pady=10)
    # textarea_open_xls.grid(row=3, column=0, padx=5, pady=10)
    button_open2.grid(row=3, column=0, sticky=W, padx=5)
    tab2_radiobutton_frame.grid(row=4, column=0, sticky=W, padx=5, pady=10)
    tab2_radiobutton_draft.grid(row=0, column=0, sticky=W)
    tab2_radiobutton_production.grid(row=1, column=0, sticky=W)
    button_run.grid(row=5, column=0, columnspan=2, pady=25)
    button_open_folder2.grid(row=8, column=0, columnspan=2, pady=25)

    # Text area
    text_area = Text(root, height=40, width=70)
    text_area.grid(row=0, column=2, rowspan=5)
    text_area.update_idletasks()
    
    scrollbar = Scrollbar(root)
    text_area.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=text_area.yview)
    scrollbar.grid(row=0, column=3, rowspan=5, sticky='NS')

    save_stdout = sys.stdout
    sys.stdout = NewStdOut(save_stdout, text_area)
    
    save_stderr = sys.stderr
    sys.stderr = NewStdOut(save_stderr, text_area)
    
    root.mainloop()


if __name__ == '__main__':
    main()
