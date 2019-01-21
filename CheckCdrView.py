import sys
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox as tkMessageBox
from tkinter import filedialog
import CheckCdrControl


def vp_start_gui():
    '''Starting point when module is the main routine.'''
    global val, w, root
    root = tk.Tk()
    top = Toplevel1(root)
    root.mainloop()


w = None


def create_Toplevel1(root, *args, **kwargs):
    '''Starting point when module is imported by another program.'''
    global w, w_win, rt
    rt = root
    w = tk.Toplevel(root)
    top = Toplevel1(w)
    return (w, top)


def destroy_Toplevel1():
    global w
    w.destroy()
    w = None


class Toplevel1:
    nSite = 3
    path = [None] * nSite
    lblCdr = []
    btnCdr = []
    btnReset = []

    def __init__(self, top=None):

        self.style = ttk.Style()
        self.style.configure('check.TButton', font=("Impact", "14"))
        self.style.configure(style="sel.TButton", font=("arial", "10", "bold"))
        self.style.configure(style="lblCdr.TLabel", font=("arial", "10", "bold"))

        # top.geometry("330x160+700+132")
        top.title("Check CDR Offline")
        # top.configure(background="#d9d9d9")
        top.resizable(False, False)

        self.checkButton = ttk.Button(top, style="check.TButton", text="CHECK THE FILE",
                                      command=lambda: checkCdrControl(self.path), )
        self.checkButton.grid(row=6, column=0, ipady=5, sticky="w,e", columnspan=3)

        for n in range(self.nSite):
            self.lblCdr.append(ttk.Label(top, relief="groove", width=40, style="lblCdr.TLabel", ))

        self.btnCdr.append(ttk.Button(top, text="SELECT...", style="sel.TButton",
                                      command=lambda: self.openFilePicker(label=self.lblCdr[0], pos=0)))
        self.btnReset.append(
            ttk.Button(top, text="X", style="sel.TButton", command=lambda: self.cleanPath(self.lblCdr[0], 0)))
        self.btnCdr.append(ttk.Button(top, text="SELECT...", style="sel.TButton",
                                      command=lambda: self.openFilePicker(label=self.lblCdr[1], pos=1)))
        self.btnReset.append(
            ttk.Button(top, text="X", style="sel.TButton", command=lambda: self.cleanPath(self.lblCdr[1], 1)))
        self.btnCdr.append(ttk.Button(top, text="SELECT...", style="sel.TButton",
                                      command=lambda: self.openFilePicker(label=self.lblCdr[2], pos=2)))
        self.btnReset.append(
            ttk.Button(top, text="X", style="sel.TButton", command=lambda: self.cleanPath(self.lblCdr[2], 2)))

        for n in range(self.nSite):
            self.lblCdr[n].grid(row=n, column=1, sticky="w,e", padx=5, pady=5)
            self.btnCdr[n].grid(row=n, column=0, sticky="w,e")
            self.btnReset[n].grid(row=n, column=2, sticky="w,e")

    def openFilePicker(self, label, pos):
        path = filedialog.askopenfilename(initialdir="/", title="Select file",
                                          filetypes=(("xlsx", "*.xlsx"), ("all files", "*.*")))
        if path != "":
            self.path[pos] = path
            tempPath = path.split("/")
            label['text'] = tempPath[-1]
        else:
            label['text'] = ""
            self.path[pos] = None

        pass

    def cleanPath(self, label, pos):
        label['text'] = ""
        self.path[pos] = None


def checkCdrControl(paths):
    Control = CheckCdrControl.CheckCdrControl(paths)
    pass


if __name__ == '__main__':
    vp_start_gui()
    pass
