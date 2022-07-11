#!python
import openpyxl as xl
from tkinter import *
from tkinter.filedialog import askopenfilenames, asksaveasfile

COPY_COLS = 3
status = 1

base = Tk()
base.title("Мержит файлы в один, за раз")
base.resizable(False, False)
base.geometry("600x400")


def save_file():
    global folder
    folder = asksaveasfile(filetypes=(("excel files","*.xlsx"), ("All files","*.*")))

def browsefunc():    
    global status
    
    r = 0

    filename1 = askopenfilenames(filetypes=(("excel files","*.xlsx"), ("All files","*.*")))
                
    while r < len(filename1)-1:
        if status:
            wb1 = xl.load_workbook(filename1[r+1]) 
            ws1 = wb1.worksheets[0]
            rows_wb1 = ws1.max_row

            wb2 = xl.load_workbook(filename1[r]) 
            ws2 = wb2.worksheets[0]
            rows_wb2 = ws2.max_row      

            i = 1    
            for c in range(1, COPY_COLS+1):
                for i in range(0, rows_wb1):
                    ws2.cell(rows_wb2+i+1, c).value = ws1.cell(i+3, c).value
            r += 1
            status = 0

            wb1.close    
            wb2.save(folder.name)
            wb2.close  
        else:
            wb1 = xl.load_workbook(filename1[r+1]) 
            ws1 = wb1.worksheets[0]
            rows_wb1 = ws1.max_row

            wb2 = xl.load_workbook(folder.name) 
            ws2 = wb2.worksheets[0]
            rows_wb2 = ws2.max_row      

            i = 1    
            for c in range(1, COPY_COLS+1):
                for i in range(0, rows_wb1):
                    ws2.cell(rows_wb2+i+1, c).value = ws1.cell(i+3, c).value
            r += 1
            status = 0

            wb1.close    
            wb2.save(folder.name)
            wb2.close

    print("Я кончил")
   
but_save = Button(
            base, 
            text="Выбор файла куда",
            bg='black', fg='white', width=20, height=5, font=('Times New Roman', 18),
            justify='left', command=save_file)
but_first = Button(
            base, 
            text="Выбор файлов которые",
            bg='white', fg='black', width=20, height=5, font=('Times New Roman', 18), 
            justify='left', command=browsefunc
            )

but_save.pack(fill=BOTH, expand=TRUE)
but_first.pack(fill=BOTH, expand=TRUE)

base.mainloop()
