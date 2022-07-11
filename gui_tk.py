from tkinter import *
from tkinter.filedialog import LoadFileDialog, askopenfilename, askopenfilenames, asksaveasfile
from tracemalloc import start

base = Tk()

def browsefunc():
    global filename
    filename = askopenfilename(filetypes=(("excel files","*.xlsx"),("All files","*.*")))
    #filename = askopenfilenames(filetypes=(("excel files","*.xlsx"),("All files","*.*")))
    #filename = asksaveasfile()
    #print(filename)

def browsefunc2():
    """ i = 0
    while i < len(filename):
        print(filename[i])
        i += 1 """
    print(filename)

but_first = Button(base, text="File first", width=20, height=10, font=24, command=browsefunc)
but_last = Button(base, text="File last", width=20, height=10, font=24, command=browsefunc2)
but_done = Button(base, text="finish", width=20, height=10, font=24)

but_first.grid(row=1, column=1)
but_last.grid(row=1, column=3)
but_done.grid(row=2, column=2)

base.mainloop()