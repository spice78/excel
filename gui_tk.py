from tkinter import *
from tkinter.filedialog import askopenfilename

base = Tk()

def browsefunc():
    filename = askopenfilename(filetypes=(("excel files","*.xlsx"),("All files","*.*")))
    print(filename)

but_first = Button(base, text="File first", width=20, height=10, font=24, command=browsefunc)
but_last = Button(base, text="File last", width=20, height=10, font=24)
but_done = Button(base, text="finish", width=20, height=10, font=24)

but_first.grid(row=1, column=1)
but_last.grid(row=1, column=3)
but_done.grid(row=2, column=2)

base.mainloop()