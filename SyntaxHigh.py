import tkinter as tk
import xlwings as xw
from DataEditorClasses import ExcelSheet

class TextLineNumbers(tk.Canvas):
    def __init__(self, *args, **kwargs):
        tk.Canvas.__init__(self, *args, **kwargs)
        self.textwidget = None

    def attach(self, text_widget):
        self.textwidget = text_widget

    def redraw(self, *args):
        '''redraw line numbers'''
        self.delete("all")

        i = self.textwidget.index("@0,0")
        while True :
            dline= self.textwidget.dlineinfo(i)
            if dline is None: break
            y = dline[1]
            linenum = str(i).split(".")[0]
            self.create_text(2,y,anchor="nw", text=linenum)
            i = self.textwidget.index("%s+1line" % i)

class CustomText(tk.Text):
    def __init__(self, *args, **kwargs):
        tk.Text.__init__(self, *args, **kwargs)

        # create a proxy for the underlying widget
        self._orig = self._w + "_orig"
        self.tk.call("rename", self._w, self._orig)
        self.tk.createcommand(self._w, self._proxy)

    def _proxy(self, *args):
        # let the actual widget perform the requested action
        cmd = (self._orig,) + args
        result = self.tk.call(cmd)

        # generate an event if something was added or deleted,
        # or the cursor position changed
        if (args[0] in ("insert", "replace", "delete") or 
            args[0:3] == ("mark", "set", "insert") or
            args[0:2] == ("xview", "moveto") or
            args[0:2] == ("xview", "scroll") or
            args[0:2] == ("yview", "moveto") or
            args[0:2] == ("yview", "scroll")
        ):
            self.event_generate("<<Change>>", when="tail")

        # return what the actual widget returned
        return result      

class Textfeld(tk.Frame):
    def __init__(self, *args, **kwargs):
        tk.Frame.__init__(self, *args, **kwargs)
        self.text = CustomText(self)
        self.vsb = tk.Scrollbar(orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=self.vsb.set)
        self.text.tag_configure("bigfont", font=("Helvetica", "24", "bold"))
        self.linenumbers = TextLineNumbers(self, width=30)
        self.linenumbers.attach(self.text)

        self.vsb.pack(side="right", fill="y")
        self.linenumbers.pack(side="left", fill="y")
        self.text.pack(side="right", fill="both", expand=True)

        self.text.bind("<<Change>>", self._on_change)
        self.text.bind("<Configure>", self._on_change)

        self.loadText()

        '''
        self.text.insert("end", "blue\n red\nthree\n")
        self.text.insert("end", "test\n",)
        self.text.insert("end", "five\n")'''

        self.closeBtn = tk.Button(self, text="Save", command=self.saveInput)
        self.closeBtn.pack()

    def hightlightSyntax(self): 
        self.keywords = ["keyword", "test"]
        self.colors = ["red", "blue", "green", "yellow"]

        self.checkKeywords(self.keywords, 'keys')
        self.checkKeywords(self.colors, 'color')
        
        self.text.tag_configure('keys', foreground='red', font='helvecita 12 bold')
        self.text.tag_configure('color', foreground='blue')

    def checkKeywords(self, list, tag):
        self.text.tag_remove(tag, '1.0', tk.END)
        for i in list:
            self.idx = '1.0'
            while True:
                self.idx = self.text.search(i, self.idx, nocase=1, stopindex=tk.END)
                if not self.idx: break
                self.lastidx = '%s+%dc' % (self.idx, len(i))
                self.text.tag_add(tag, self.idx, self.lastidx)
                self.idx = self.lastidx
                self.text.see(self.idx)

    def loadText(self):
        self.excel = ExcelSheet('test.xlsm', 'Testabelle1')
        self.ws = self.excel.wsOrig
        self.shadow = self.excel.wsShadow
        self.rowCount = 1.0
        while True:
            if self.shadow.range((self.rowCount, 5)).value == None:
                break
            self.text.insert(self.rowCount, self.shadow.range((self.rowCount, 5)).value)
            self.rowCount = self.rowCount + 1

    def saveInput(self):
        self.deleteOld()
        self.rowCount = 1.0
        while True:
            if self.text.get(self.rowCount, self.rowCount+1) == '':
                break
            self.shadow.range((self.rowCount, 5)).value = self.text.get(self.rowCount, self.rowCount+1)
            self.rowCount = self.rowCount + 1.0

    def deleteOld(self):
        self.rowCount = 1.0
        while True:
            if self.shadow.range((self.rowCount, 5)).value == None:
                break
            self.shadow.range((self.rowCount, 5)).value = None
            self.rowCount = self.rowCount + 1.0

    def _on_change(self, event):
        self.linenumbers.redraw()
        self.hightlightSyntax()

if __name__ == "__main__":
    root = tk.Tk()
    Textfeld(root).pack(side="top", fill="both", expand=True)
    root.mainloop()