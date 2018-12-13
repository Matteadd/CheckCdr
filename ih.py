from tkinter import ttk
import tkinter

root = tkinter.Tk()

ttk.Style().configure(style="pers.TLabel", padding=6, relief="flat",
   background="#000000")

btn = ttk.Button(text="Sample", style="pers.TLabel"  )
btn.pack()

root.mainloop()
