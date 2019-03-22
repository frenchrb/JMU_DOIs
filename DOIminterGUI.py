from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import Bepress2DataCiteDrafts
import MintDOIsDrafts

class newStdOut():
    def __init__(self, oldStdOut, textArea): #remove oldStdOut
        self.oldStdOut = oldStdOut #remove if not using old stdout
        self.textArea = textArea
        #open logfile

    def write(self, x):
        #self.oldStdOut.write(x)  #this prints to the old stdout (cmd line)
        self.textArea.insert(END, x)
        self.textArea.see(END)
        self.textArea.update_idletasks()
        #write to logfile
    
    def flush(self):
        self.oldStdOut.flush() #replace this line with pass if removing old stdout stuff
        #close logfile
#probably have to do the same with standarderror if you want to see those messages
#sys.stderror = newstdout() ??

def open():
    global openpath
    openpath = filedialog.askopenfilename()
    #print('openpath: ' + openpath)

def runBepress2DataCiteDrafts():
    print()
    args = []
    args.append(setname.get())
    args.append(openpath)
    # if production.get() == 1:
        # args.append('--production')
    #print(args)
    Bepress2DataCiteDrafts.main(args)
    # root.destroy()
    
def main():
    root = Tk()
    root.wm_title('JMU DOI Minter')
    root.resizable(width = FALSE, height = FALSE)
    
    label_entrybox_setname = Label(root, text='Enter bepress collection setname (e.g., "diss201019")')
    label_entrybox_setname.grid(row = 0, column = 0, sticky = W)
    
    global setname
    setname = StringVar()
    entrybox_setname = Entry(root, textvariable = setname, width = 20)
    entrybox_setname.grid(row = 0, column = 1)
    
    
    label_button_open = Label(root, text='Select bepress spreadsheet saved as "XML Spreadsheet 2003" (*.xml)"')
    label_button_open.grid(row = 1, column = 0, columnspan = 2, sticky = W)
    
    textarea_open = Text(root, height = 1, width = 40)
    textarea_open.grid(row = 2, column = 0)
    
    button_open = Button(root, text='Browse...', height = 1, width = 10, command=open)
    button_open.grid(row = 2, column = 1)
    
    
    button_run = Button(root, text='Run Bepress2DataCiteDrafts', height=2, width=25, command=runBepress2DataCiteDrafts)
    button_run.grid(row = 3, column = 0, columnspan = 2)

    # button_C = Button(root, text='Button C', height=2, width=25, command=buttonC)
    # button_C.grid(row=2, column=0)

    # button_D = Button(root, text='Function A', height=2, width=25, command=buttonD)
    # button_D.grid(row=3, column=0)

    #global text_area
    text_area = Text(root, height = 50, width = 70)
    text_area.grid(row = 0, column = 2, rowspan = 4)
    text_area.update_idletasks()
    
    scrollbar = Scrollbar(root)
    text_area.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=text_area.yview)
    scrollbar.grid(row=0, column=3, rowspan=4, sticky='NS')

    save_stdout = sys.stdout #remove
    sys.stdout = newStdOut(save_stdout, text_area) #remove save_stdout

    root.mainloop()

    #subprocess.call - use p=subprocess.popen instead, including stdout=subprocess.PIPE and stderr=subprocess.PIPE
    #stdout,stderr=p.communicate()
    
    #or subprocess.check_call?
    
if __name__ == '__main__':
    main()