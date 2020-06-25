import tkinter as tk
import tkinter.ttk as ttk
import xlwings as xw
from win32com import client

class ExcelSheet():
    def __init__(self, workbook, sheet):
        self.wb = xw.Book(workbook)
        self.wsOrig = self.wb.sheets(sheet)
        self.shadowsheet = sheet + '_shadow'
        self.wsShadow = self.createShadow(self.shadowsheet, workbook)

    def createShadow(self, shadowsheet, workbook):
        try:
            self.wb.sheets.add(shadowsheet)
            self.hideFile(workbook, shadowsheet)
            print('Shadowsheet created')
        except:
            pass
        finally:
            return self.wb.sheets(shadowsheet)

    def hideFile(self, workbook, shadowsheet):
        self.xl = client.Dispatch("Excel.Application")
        self.wb = self.xl.Workbooks.Open(
            r'C:\Users\nikla\OneDrive\Desktop\Projekt_ExcelPython\test.xlsm')
        # 2: very hidden; 1:visible; 0:hidden
        self.wb.Worksheets(shadowsheet).Visible = 2


class Buttons():
    def __init__(self, frame, width, funcs):
        self.buttonsframe = tk.Frame(frame)
        self.buttonsframe.pack()

        self.addButton = self.buttonFormula('Add Data Set', width, funcs[0])
        self.addButton.frame.grid(row=0, column=0, padx=2, pady=2)

        self.updateButton = self.buttonFormula(
            'Update Data Set', width, funcs[1])
        self.updateButton.frame.grid(row=0, column=1, padx=2, pady=2)

        self.deleteButton = self.buttonFormula(
            'Delete Data Set', width, funcs[2])
        self.deleteButton.frame.grid(row=0, column=2, padx=2, pady=2)

        self.clearButton = self.buttonFormula('Clear Entries', width, funcs[3])
        self.clearButton.frame.grid(row=0, column=3, padx=2, pady=2)

        self.placeholder = tk.Frame(self.buttonsframe, width=int(width/7))
        self.placeholder.grid(row=0, column=4, padx=2, pady=2)

        self.saveButton = self.buttonFormula('Save Sets', width, funcs[4])
        self.saveButton.frame.grid(row=0, column=5, padx=2, pady=2)

        self.loadButton = self.buttonFormula('Manage Sets', width, funcs[5])
        self.loadButton.frame.grid(row=0, column=6, padx=2, pady=2)

    def buttonFormula(self, text, width, function):
        self.frame = tk.Frame(self.buttonsframe, width=int(width/7))
        self.button = tk.Button(self.frame, text=str(text), command=function)
        self.button.pack()
        return self


class Entries():
    def __init__(self, midFrame, height, width, isIndexRange, text):
        self.entryFormula(midFrame, height, width, isIndexRange, text)

    def entryFormula(self, midFrame, mframesubh, mframesubw, isIndexRange, text):
        self.frame = ttk.Frame(midFrame, width=int(
            mframesubw/3), height=mframesubh)

        self.labelframe = tk.Frame(self.frame,
                                   width=int(mframesubw/3), height=int(mframesubh*3/7))
        self.labelframe.grid(row=0, column=0)
        self.label = tk.Label(self.labelframe, text=str(text))
        self.label.pack()

        self.entryframe = tk.Frame(self.frame,
                                   width=int(mframesubw/3), height=int(mframesubh*2/7))
        self.entryframe.grid(row=1, column=0)
        self.entry = tk.Entry(self.entryframe)
        self.entry.pack()

        if isIndexRange:
            self.entryframe2 = tk.Frame(self.frame,
                                        width=int(mframesubw/3), height=int(mframesubh*2/7))
            self.entryframe2.grid(row=2, column=0)
            self.entry2 = tk.Entry(self.entryframe2)
            self.entry2.pack()
        else:
            self.placeholderframe = tk.Frame(self.frame,
                                             width=int(mframesubw/3), height=int(mframesubh*2/7))
            self.placeholderframe.grid(row=2, column=0)
