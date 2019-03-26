from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import Bepress2DataCiteDrafts
import MintDOIsDrafts

class newStdOut():
    def __init__(self, oldStdOut, textArea): #remove oldStdOut
        self.oldStdOut = oldStdOut #remove if not using old stdout
        self.textArea = textArea

    def write(self, x):
        #self.oldStdOut.write(x)  #this prints to the old stdout (cmd line)
        self.textArea.insert(END, x)
        self.textArea.see(END)
        self.textArea.update_idletasks()
    
    def flush(self):
        self.oldStdOut.flush() #replace this line with pass if removing old stdout stuff

def openXML():
    global openpath_xml
    openpath_xml = filedialog.askopenfilename(title = 'Select file', filetypes = (("XML Spreadsheet 2003","*.xml"),("All files","*.*")))
    #print('openpath_xml: ' + openpath_xml)

def openXLS():
    global openpath_xls
    openpath_xls = filedialog.askopenfilename(title = 'Select file', filetypes = (("Excel 97-2003 Workbook","*.xls"),("All files","*.*")))
    #print('openpath_xls: ' + openpath_xls)

def runBepress2DataCiteDrafts():
    #print()
    if not setname.get():
        print('Please enter collection setname.')
    try:
        openpath_xml
    except NameError:
        print('Please select spreadsheet file.')
    args = []
    args.append(setname.get())
    args.append(openpath_xml)
    # if production.get() == 1:
        # args.append('--production')
    #print(args)
    Bepress2DataCiteDrafts.main(args)
    # root.destroy()
    
def runMintDOIsDrafts():
    #print()
    try:
        openpath_xls
    except NameError:
        print('Please select spreadsheet file.')
    args = []
    args.append(openpath_xls)
    # if production.get() == 1:
        # args.append('--production')
    #print(args)
    MintDOIsDrafts.main(args)
    # root.destroy()

def main():
    root = Tk()
    root.wm_title('JMU DOI Minter')
    root.resizable(width = FALSE, height = FALSE)
    
    #Tabs
    notebook = ttk.Notebook(root)
    tab1 = ttk.Frame(notebook)
    tab2 = ttk.Frame(notebook)
    notebook.add(tab1, text = '1. Create Metadata')
    notebook.add(tab2, text = '2. Mint DOIs')
    notebook.grid(row = 0, column = 0, sticky = NW)
    
    
    #Tab 1
    tab1_frame = Frame(tab1)
    tab1_frame.grid()
    
    #Create tab 1 widgets
    label_entrybox_setname = Label(tab1_frame, text='Enter bepress collection setname (e.g., "diss201019")')
    global setname
    setname = StringVar()
    entrybox_setname = Entry(tab1_frame, textvariable = setname, width = 20)
    label_button_open_xml = Label(tab1_frame, text='Select bepress spreadsheet saved as "XML Spreadsheet 2003" (*.xml)"')
    textarea_open_xml = Text(tab1_frame, height = 1, width = 40)
    button_open_xml = Button(tab1_frame, text='Browse...', height = 1, width = 10, command=openXML)
    button_run = Button(tab1_frame, text='Run Bepress2DataCiteDrafts', height=2, width=25, command=runBepress2DataCiteDrafts)
    
    #Lay out tab 1 widgets
    label_entrybox_setname.grid(row = 1, column = 0, sticky = W, padx = 5, pady = 10)
    entrybox_setname.grid(row = 1, column = 1, padx = 5, pady = 10)
    label_button_open_xml.grid(row = 2, column = 0, columnspan = 2, sticky = W, padx = 5, pady = 10)
    textarea_open_xml.grid(row = 3, column = 0, padx = 5, pady = 10)
    button_open_xml.grid(row = 3, column = 1)
    button_run.grid(row = 4, column = 0, columnspan = 2, pady = 25)

    
    #Tab 2
    tab2_frame = Frame(tab2)
    tab2_frame.grid()
    
    #Create tab 2 widgets
    label_button_open_xls = Label(tab2_frame, text='Select bepress spreadsheet saved as "Excel 97-2003 Workbook " (*.xls)"')
    textarea_open_xls = Text(tab2_frame, height = 1, width = 40)
    button_open_xls = Button(tab2_frame, text='Browse...', height = 1, width = 10, command=openXLS)
    button_run = Button(tab2_frame, text='Run MintDOIsDrafts', height=2, width=25, command=runMintDOIsDrafts)
    
    #Lay out tab 2 widgets
    label_button_open_xls.grid(row = 2, column = 0, columnspan = 2, sticky = W, padx = 5, pady = 10)
    textarea_open_xls.grid(row = 3, column = 0, padx = 5, pady = 10)
    button_open_xls.grid(row = 3, column = 1)
    button_run.grid(row = 4, column = 0, columnspan = 2, pady = 25)
    
    
    #Text area
    text_area = Text(root, height = 50, width = 70)
    text_area.grid(row = 0, column = 2, rowspan = 5)
    text_area.update_idletasks()
    
    scrollbar = Scrollbar(root)
    text_area.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=text_area.yview)
    scrollbar.grid(row=0, column=3, rowspan=5, sticky='NS')

    save_stdout = sys.stdout #remove
    sys.stdout = newStdOut(save_stdout, text_area) #remove save_stdout

    root.mainloop()
    
if __name__ == '__main__':
    main()