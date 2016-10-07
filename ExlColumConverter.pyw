import tkinter as inter
import string

from tkinter import messagebox as msgbox

class gui(inter.Tk):
    def __init__(self,parent):
        inter.Tk.__init__(self,parent)
        self.parent = parent
        
    def build(self):
        #lables
        self.asDez = inter.Label(self, text="DECIMAL", font=("Arial",12),justify='left', anchor='w')
        self.asDez.place(x = 25, y=15, width = 300, height = 20)
        self.asExl = inter.Label(self, text="EXCEL-ROW", font=("Arial",12),justify='left',anchor='w')
        self.asExl.place(x = 25, y=90, width = 300, height = 20)
        #entrys
        self.enDez = inter.Entry(self, font=("Arial",12))
        self.enDez.place(x = 25, y=40, width = 300, height = 25)
        self.enExl = inter.Entry(self, font=("arial",12))
        self.enExl.place(x = 25, y=115, width = 300, height = 25)
        
    def run(self):
        self.enDez.bind("<Any-KeyPress>", self.DezConvert)
        self.enExl.bind("<Any-KeyPress>", self.ExlConvert)
        
    def DezConvert(self, event):
        #convert the Dezimalnumber to Excel Column and write it in enExl
        DezStr = self.enDez.get() + event.char
        if basics.is_Int(DezStr) and DezStr != "":
            DezInt = int(DezStr)

            OutputExl = ""
            while DezInt > 26:
                OutputExl = chr(DezInt %  26 +64) + OutputExl
                DezInt = DezInt // 26
            OutputExl = chr(DezInt +64) + OutputExl
            self.enExl.delete(0,'end')
            self.enExl.insert(0,OutputExl)
            
        else:
            self.enExl.delete(0,'end')
            
    def ExlConvert(self, event):
        #convert the Excelnumber to Dezimalnumber and write it in enDez
        ExlStr = self.enExl.get() + event.char
        
        if ExlStr.isalpha() and ExlStr != "":
            ExlStr = ExlStr.upper()
            DezOut = 0
            StrLen = len(ExlStr)
            
            for i in range(StrLen,0,-1):
                #convert the string to number and from 26 system in dec system
                DezOut = DezOut + ((ord(ExlStr[:1])-64)*(26**(i-1)))
                ExlStr = ExlStr[1:]
            
            self.enDez.delete(0,"end")
            self.enDez.insert(0,DezOut)
        else:
            self.enDez.delete(0,"end")

#basic gennerall classes
class basics():
    #check if sth is a string
    def is_Int(IntString):
        try:
            int(IntString)
            return True
        except ValueError:
            return False

#gui launcher
if __name__ =="__main__":
    box = gui(None)
    box.title("Converter")
    box.geometry("350x165")
    box.resizable(False,False)
    box.build()
    box.run()
    box.mainloop()
    exit()

