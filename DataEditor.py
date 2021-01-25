import tkinter as tk
import tkinter.ttk as ttk
import xlwings as xw
import sys
from win32com import client
import re
from DataEditorClasses import Buttons, Entries, ExcelSheet


class GUI(tk.Frame):
    def __init__(self, worksheet = None, *args, **kwargs):
        tk.Frame.__init__(self, *args, **kwargs)
        self.wsheet = worksheet
        print(self.wsheet)
        try:
            self.excel = ExcelSheet('test.xlsm', self.wsheet)
        except:
            self.excel = ExcelSheet('test.xlsm', 'Testabelle1')
        self.ws = self.excel.wsShadow
        self.functions = Functions(self.excel.wb, self.ws)
        self.frames(600, 360)

    def frames(self, fwidth, fheight):
        self.topFrame = ttk.Frame(self, width=fwidth, height=int(fheight*0.5))
        self.topFrame.grid(row=0, column=0, padx=5, pady=5)
        self.topFrameContent(fwidth)

        self.midFrame = ttk.Frame(self, width=fwidth)
        self.midFrame.grid(row=1, column=0, padx=5, pady=5)
        self.midFrameContent(fwidth, int(fheight*0.2))

        self.botFrame = ttk.Frame(self,  width=fwidth, height=int(fheight*0.2))
        self.botFrame.grid(row=2, column=0, padx=5, pady=5)
        self.botFrameContent(fwidth)

        self.saveFrame = tk.Frame(
            self, width=fwidth, height=int(fheight*0.2))
        self.saveFrameContent()

        self.manageFrame = tk.Frame(
            self, width=fwidth, height=int(fheight*0.2))
        self.manageFrameContent()

    def topFrameContent(self, fwidth):
        self.treeview = ttk.Treeview(self.topFrame)
        self.treeview['columns'] = ('cellrange', 'indexrange')

        self.columns = ['#0', 'cellrange', 'indexrange']
        self.headings = ['Name', 'Cell Range', 'Index Range(s)']

        for (x, y) in zip(self.columns, self.headings):
            self.treeview.heading(x, text=y, anchor='w')
            self.treeview.column(x, anchor='w', width=int(fwidth/3))

        self.treeview.insert("", 0, text="<Add New Data Item>")
        self.functions.loadTree(self.treeview)

        self.treeview.grid(row=1, column=1)

        self.treeview.bind('<<TreeviewSelect>>', self.showFocused)

    def showFocused(self, *args):
        self.curItem = self.treeview.focus()
        self.ind = self.treeview.index(self.curItem)
        self.functions.clearEntries(self.entriesList)

        try:
            self.nameinput.entry.insert(0, self.ws.range((self.ind, 1)).value)
            self.crinput.entry.insert(0, self.ws.range((self.ind, 2)).value)
            self.irinput.entry.insert(0, self.ws.range((self.ind, 3)).value)
            self.irinput.entry2.insert(0, self.ws.range((self.ind, 4)).value)
        except:
            pass

    def midFrameContent(self, mframew, mframeh):
        self.nameinput = Entries(
            self.midFrame, mframeh, mframew, False, 'Name:')
        self.nameinput.frame.grid(row=0, column=0)

        self.crinput = Entries(self.midFrame, mframeh,
                               mframew, False, 'Cell Range:')
        self.crinput.frame.grid(row=0, column=1)

        self.irinput = Entries(self.midFrame, mframeh,
                               mframew, True, 'Index Range(s)')
        self.irinput.frame.grid(row=0, column=2)

    def botFrameContent(self, width):
        self.entriesList = [self.nameinput.entry, self.crinput.entry,
                            self.irinput.entry, self.irinput.entry2]

        self.buttonFunctions = [lambda:self.functions.addData(self.entriesList, self.treeview),
                                lambda:self.functions.updateData(
                                    self.treeview, self.entriesList),
                                lambda:self.functions.deleteAndClear(
                                    self.treeview,  self.entriesList),
                                lambda:self.functions.clearEntries(
                                    self.entriesList),
                                lambda:self.functions.saveSet(
                                    self.saveFrame, self.midFrame, self.botFrame),
                                self.functions.loadTree(self.treeview)]
        self.buttons = Buttons(self.botFrame, width, self.buttonFunctions)

    def saveFrameContent(self):
        self.saveEntry = tk.Entry(self.saveFrame)
        self.saveEntry.grid(row=0, column=0, padx=5, pady=5)
        self.saveEntry.insert(0, 'Insert Name Here')

        self.saveConfirmButton = tk.Button(self.saveFrame, text='Save Data Sets', command=lambda: self.functions.confirmSave(
            self.saveFrame, self.saveEntry, self.combo))
        self.saveConfirmButton.grid(row=0, column=1, padx=5, pady=5)

        self.saveCancel = tk.Button(self.saveFrame, text='Cancel', command=lambda: self.functions.cancelSave(
            self.saveFrame, self.midFrame, self.botFrame))
        self.saveCancel.grid(row=0, column=2)

    def manageFrameContent(self):
        self.combo = ttk.Combobox(
            self.saveFrame, state='readonly')
        self.functions.setCombobox(self.combo)
        self.combo.grid(row=0, column=4, padx=5, pady=5)

        self.manageDeleteButton = tk.Button(
            self.saveFrame, text='Delete Set', command=lambda: self.functions.deleteSet(self.combo))
        self.manageDeleteButton.grid(row=0, column=5, padx=5, pady=5)

        self.manageLoadButton = tk.Button(
            self.saveFrame, text='Load Set', command=lambda: self.functions.loadSet(self.combo, self.treeview))
        self.manageLoadButton.grid(row=0, column=6, padx=5, pady=5)


class Functions():
    def __init__(self, wb, ws):
        self.wb = wb
        self.ws = ws

    def addData(self, entryList, tree, col=0):
        self.count = self.counterDown()
        self.double = False
        for x in range(self.count):
            if self.ws.range((x+1, 1)).value == entryList[0].get():
                print('This Name Was Already Used, Please Take Another One!')
                self.double = True
                break
        if entryList[0].get() and entryList[1].get() and self.double == False and self.checkCellRangeInput(entryList[1].get()):
            self.column = 1 + col
            for x in entryList:
                self.ws.range((self.count, self.column)).value = x.get()
                self.column = self.column + 1

            self.wb.save()
            self.clearEntries(entryList)
            self.loadTree(tree, 1)

    def updateData(self, tree, entryList):
        self.curItem = tree.focus()
        self.ind = tree.index(self.curItem)

        self.deleteData(tree, entryList)
        self.addData(entryList, tree)

    def deleteData(self, tree, entryList):
        self.curItem = tree.focus()
        self.ind = tree.index(self.curItem)
        self.moveUp = self.counterDown() - self.ind - 1
        self.lastRow = self.counterDown()-1

        for x in range(4):
            self.ws.range((self.ind, x+1)).value = None

        for x in range(self.moveUp):
            for y in range(4):
                self.ws.range(
                    (self.ind+x, y+1)).value = self.ws.range((self.ind+x+1, y+1)).value

        for x in range(4):
            self.ws.range((self.lastRow, x+1)).value = None

        self.loadTree(tree)

    def deleteAndClear(self, tree, entryList):
        self.deleteData(tree, entryList)
        self.clearEntries(entryList)

    def clearEntries(self, entryList):
        for x in entryList:
            x.delete(0, tk.END)
            x.insert(0, "")

    def saveSet(self, saveFrame, midFrame, botFrame):
        saveFrame.grid(row=2, column=0)
        midFrame.grid_forget()
        botFrame.grid_forget()

    def confirmSave(self, saveFrame, entry, combo):
        self.addCol = self.counterSide()-1
        if entry.get() != 'Insert Name Here':
            for x in range(4):
                for y in range(self.counterDown()):
                    self.ws.range((y+1, x+1+self.addCol)
                                  ).value = self.ws.range((y+1, x+1)).value

            self.ws.range((self.counterDown(), self.addCol+1)
                          ).value = entry.get()
            self.setCombobox(combo)
            self.wb.save()
        else:
            print('Please Give a Name')

    def cancelSave(self, saveFrame, midFrame, botFrame):
        saveFrame.grid_forget()
        midFrame.grid(row=1, column=0, padx=5, pady=5)
        botFrame.grid(row=2, column=0, padx=5, pady=5)

    def loadSet(self, combo, tree):
        for sets in range(len(self.savedSets)):
            if combo.get() == self.savedSets[sets][0]:
                for rows in range(self.savedSets[sets][1]-1):
                    for cols in range(4):
                        self.ws.range(
                                (1+rows, cols+1)).value = self.ws.range((1+rows, self.savedSets[sets][2]+cols)).value
                        if self.counterDown() >= self.savedSets[sets][1]:
                            for x in range(self.counterDown()+1-self.savedSets[sets][1]):
                                self.ws.range((self.savedSets[sets][1]-x+2, 4-cols)).value = None

        self.loadTree(tree)

    def setCombobox(self, combo):
        self.savedSets = []
        self.savedSetsNames = []

        for x in range(int((self.counterSide()-1)/5)):
            if x != 0:
                self.savedSets.append([self.ws.range(self.counterDown(
                    x*5)-1, x*5+1).value, self.counterDown(x*5)-1, x*5+1])
                self.savedSetsNames.append(self.ws.range(
                    self.counterDown(x*5)-1, x*5+1).value)
        combo.set('Please Select a Data Set to Delete')
        combo.configure(value=self.savedSetsNames)

    def deleteSet(self, combo):
        for sets in range(len(self.savedSets)):
            if combo.get() == self.savedSets[sets][0]:
                for cols in range(4):
                    for rows in range(self.savedSets[sets][1]):
                        for turns in range(len(self.savedSets)-sets):
                            self.ws.range((self.savedSets[sets][1]-rows, self.savedSets[sets][2]+cols+turns*5)).value = self.ws.range(
                                (self.savedSets[sets][1]-rows, self.savedSets[sets][2]+cols+5+turns*5)).value
        self.setCombobox(combo)

    def loadTree(self, tree, x=0):
        for i in tree.get_children():
            tree.delete(i)
        tree.insert("", 0, text="<Add New Data Item>")
        for x in range(self.counterDown()-1+x):
            self.phlist = []
            for i in range(4):
                if self.ws.range((x+1, i+1)).value != None:
                    self.phlist.append(self.ws.range((x+1, i+1)).value)
                else:
                    self.phlist.append("")
            if self.phlist[2] == "" or self.phlist[3] == "":
                self.irvalue = str(self.phlist[2] + "" + str(self.phlist[3]))
            else:
                self.irvalue = str(self.phlist[2]) + ', ' + str(self.phlist[3])

            tree.insert(
                "", x+1, text=str(self.phlist[0]), value=(str(self.phlist[1]), self.irvalue))

    def counterDown(self, side=0):
        self.counter = 1
        while True:
            if self.ws.range((self.counter, 1+side)).value == None:
                return self.counter
            else:
                self.counter = self.counter + 1

    def counterSide(self):
        self.counter = 1
        while True:
            if self.ws.range((1, self.counter)).value == None:
                return self.counter
            else:
                self.counter = self.counter + 5

    def checkCellRangeInput(self, input):
        self.pattern = re.compile(
            r'^\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?(?:,\s*(?:\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?))*$')
        print(input)
        if self.pattern.match(str(input)):
            return True
        return False


if __name__ == "__main__":
    root = tk.Tk()
    try:
        GUI(root, sys.argv[1]).pack(side="top", fill="both", expand=True)
    except:
        GUI(root).pack(side="top", fill="both", expand=True)
   
    root.mainloop()