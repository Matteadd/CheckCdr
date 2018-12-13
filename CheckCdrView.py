import sys
import tkinter.ttk as ttk
import tkinter as tk
from tkinter import messagebox as tkMessageBox
from tkinter import filedialog
import CheckCdrControl

def vp_start_gui():
    '''Starting point when module is the main routine.'''
    global val, w, root
    root = tk.Tk()
    top = Toplevel1 (root)
    root.mainloop()

w = None
def create_Toplevel1(root, *args, **kwargs):
    '''Starting point when module is imported by another program.'''
    global w, w_win, rt
    rt = root
    w = tk.Toplevel (root)
    top = Toplevel1 (w)
    return (w, top)

def destroy_Toplevel1():
    global w
    w.destroy()
    w = None


class Toplevel1:


    path=[None]*5


    def __init__(self, top=None):
        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''
        _bgcolor = '#d9d9d9'  # X11 color: 'gray85'
        _fgcolor = '#000000'  # X11 color: 'black'
        _compcolor = '#d9d9d9' # X11 color: 'gray85'
        _ana1color = '#d9d9d9' # X11 color: 'gray85'
        _ana2color = '#d9d9d9' # X11 color: 'gray85'
        self.style = ttk.Style()
        if sys.platform == "win32":
            self.style.theme_use('winnative')
        self.style.configure('.',background=_bgcolor)
        self.style.configure('.',foreground=_fgcolor)
        self.style.configure('.',font="TkDefaultFont")
        self.style.map('.',background=[('selected', _compcolor), ('active',_ana2color)])

        # top.geometry("330x160+700+132")
        top.title("Check CDR Offline")
        top.configure(background="#d9d9d9")
        top.resizable(False, False)

        self.TButton1 = ttk.Button(top,text="CHECK THE FILE",command= lambda:checkCdrControl(self.path),width=30)
        self.TButton1.grid(row=5, column=0, columnspan=3 )
        self.TButton1.configure(takefocus="")

        self.lblCdr1 = ttk.Label(top, relief="groove", width=40)
        self.lblCdr1.grid(row=0, column=1)

        self.lblCdr2 = ttk.Label(top, relief="groove", width=40)
        self.lblCdr2.grid(row=1, column=1)

        self.lblCdr3 = ttk.Label(top, relief="groove", width=40)
        self.lblCdr3.grid(row=2, column=1)

        self.lblCdr4 = ttk.Label(top, relief="groove", width=40)
        self.lblCdr4.grid(row=3, column=1)

        self.lblCdr5 = ttk.Label(top, relief="groove", width=40)
        self.lblCdr5.grid(row=4, column=1)

        self.btnCdr1=ttk.Button(top,text="SELECT...",command = lambda:self.openFilePicker(self.lblCdr1, 0),width=10)
        self.btnCdr1.grid(row=0, column=0)

        self.btnCdr2=ttk.Button(top,text="SELECT...",command = lambda:self.openFilePicker(self.lblCdr2, 1),width=10)
        self.btnCdr2.grid(row=1, column=0)

        self.btnCdr3=ttk.Button(top,text="SELECT...",command = lambda:self.openFilePicker(self.lblCdr3, 2),width=10)
        self.btnCdr3.grid(row=2, column=0)

        self.btnCdr4=ttk.Button(top,text="SELECT...",command = lambda:self.openFilePicker(self.lblCdr4, 3),width=10)
        self.btnCdr4.grid(row=3, column=0)

        self.btnCdr5=ttk.Button(top,text="SELECT...",command = lambda:self.openFilePicker(self.lblCdr5, 4),width=10)
        self.btnCdr5.grid(row=4, column=0)

        self.btnReset1=ttk.Button(top,text="X",command = lambda:self.cleanPath(self.lblCdr1, 0),width=10)
        self.btnReset1.grid(row=0, column=2)

        self.btnReset2=ttk.Button(top,text="X",command = lambda:self.cleanPath(self.lblCdr2, 1),width=10)
        self.btnReset2.grid(row=1, column=2)

        self.btnReset3=ttk.Button(top,text="X",command = lambda:self.cleanPath(self.lblCdr3, 2),width=10)
        self.btnReset3.grid(row=2, column=2)

        self.btnReset4=ttk.Button(top,text="X",command = lambda:self.cleanPath(self.lblCdr4, 3),width=10)
        self.btnReset4.grid(row=3, column=2)

        self.btnReset5=ttk.Button(top,text="X",command = lambda:self.cleanPath(self.lblCdr5, 4),width=10)
        self.btnReset5.grid(row=4, column=2)

    def openFilePicker(self,label,pos):
        path=filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xlsx","*.xlsx"),("all files","*.*")))
        if path != "":
            self.path[pos]=path
            tempPath=path.split("/")
            label['text']=tempPath[-1]
        else:
            label['text']=""
            self.path[pos]=None

        pass

    def cleanPath(self,label,pos):
        label['text']=""
        self.path[pos]=None

def checkCdrControl(path):
    # print(path)
    Control= CheckCdrControl.CheckCdrControl(path)
    pass

if __name__ == '__main__':
    vp_start_gui()
    pass
