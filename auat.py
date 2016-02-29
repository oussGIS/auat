import xlrd
from Tkinter import *
from tkFileDialog import *
import tkMessageBox
from qgis.core import *
from qgis.gui import *
from PyQt4.QtCore import *
from PyQt4 import *

class XlsxClass:
    def __init__(self, parent):

        global v1, v2, testQuit, fm, path, chemin
        
        v2 = StringVar()
        v1 = StringVar()
        testQuit = 0

        self.myParent = parent
        self.myParent.wm_attributes('-topmost', 1)
        self.myParent.resizable(width=FALSE, height=FALSE)
        self.myParent.minsize(width=600, height=600)
        self.myParent.maxsize(width=600, height=600)
        self.myParent.title("Jointure en un Shapefile et un classeur Excel")

        self.fm = Frame(self.myParent)
        self.fm.pack()
        
        
        self.labSel1 = Label(self.fm, fg='red', relief=GROOVE)
        self.labSel2 = Label(self.fm, fg='red', relief=GROOVE)
        
        self.lab1 = Label(self.fm, text='Le 1er champs de jointure')
        self.lab2 = Label(self.fm, text='Le 2eme champs de jointure')

        self.btnXlsx = Button(self.fm, text='Open excel', height=5, width=10, command=self.openExcel)
        self.btnXlsx.grid(row=0, column=2, padx = 10, pady=10)

        self.btnCop = Button(self.fm, text='Copie', height=5, width=10, command=self.Copie)
        self.btnCop.grid(row=0, column=1, padx = 10, pady=10)

        self.btnShp = Button(self.fm, text='Open shpfile', height=5, width=10, command=self.openShp)
        self.btnShp.grid(row=0, column=0, padx = 10, pady=10)
        
        self.btnQuit = Button(self.fm, text='Quit', height=5, width=10, command=self.Quit)
        self.btnQuit.grid(row=1, column=1, padx = 10, pady=10)

        
        

        
        





    def openShp(self):
        global path, couche, listbox1, field_names, test, nomT, nom, champs1
        nomT = []
        nom = ""
        champs1 = ""
        self.path = askopenfilename(parent=self.myParent, filetypes=[("ShapeFiles", "*.shp")])
        nomT = self.path.split("/")
        nom = nomT[len(nomT)-1].replace(".shp", "")
        couche = QgsVectorLayer(self.path, nom, "ogr")

        QgsMapLayerRegistry.instance().addMapLayer(couche) 
        fields = couche.pendingFields()  
        field_names = [field.name() for field in fields]
        testQuit = len(field_names) 
        for item in field_names:
            self.radio = Radiobutton(self.fm, text=item, variable=v1, value=item, bg="white", command=self.Test1)
            self.radio.grid(row=field_names.index(item)+6, column=0, sticky=W, padx=10)
        self.lab1.grid(row=3, column=0, sticky=W, padx = 10, pady=10)


    def openExcel(self):
        global sheet, workbook, chemin, listbox2, champs2
        champs2 = ""
        self.chemin = askopenfilename(parent=self.myParent, filetypes=[("Fichier Excel", "*.xlsx"), ("Fichier Excel (1997- 2003)", "*.xls")])
        workbook = xlrd.open_workbook(self.chemin)
        sheet = workbook.sheet_by_index(0)
        for c in range(sheet.ncols):
            self.radio = Radiobutton(self.fm, text=sheet.cell_value(0, c), variable=v2, value=sheet.cell_value(0, c), command=self.Test2)
            self.radio.grid(row=c+6, column=2, sticky=W, padx=10)
        self.lab2.grid(row=3, column=2, sticky=E, padx = 10, pady=10)


    def Copie(self):
        global index, index1, index2
        v = sheet.nrows - 1
        if(v != couche.dataProvider().featureCount()):
           self.msgNbre = tkMessageBox.showerror("Erreur", "le nombre d'objets de votre fichier SHP ne correspond pas au nombre d'enregistrements de votre fichier Excel !")
        elif(v == couche.dataProvider().featureCount()):
            if(champs1 == "" or champs2 == ""):
                self.msgSel = tkMessageBox.showerror("Erreur", "veuillez selectionner les deux champs de jointure !")
            else:
                index = couche.fieldNameIndex("type")
                for i in range(sheet.ncols):
                    couche.startEditing()
                    fournisseur = couche.dataProvider()
                    fournisseur.addAttributes([QgsField(sheet.cell_value(0,i), QVariant.String)])
                couche.commitChanges()

                index1 = couche.fieldNameIndex(champs1)
    
                for obj in couche.getFeatures():
                    val1 = obj.attributes()[index1]
                    for c in range(sheet.ncols):
                        if (sheet.cell_value(0, c) == champs2):
                            index2 = c
                            continue
                    for r in range(sheet.nrows):
                        if (val1 == sheet.cell_value(r, index2)):
                            for c in range(sheet.ncols):
                                index = couche.fieldNameIndex(sheet.cell_value(0,c))
                                attrs = {index: sheet.cell_value(r, c)}
                                fournisseur.changeAttributeValues({ obj.id() : attrs })
                couche.commitChanges()
                couche.updateExtents()


    def Test1(self):
        global champs1
        champs1 = str(v1.get())
        self.labSel1.grid(row=15, column=0, sticky=W, padx = 10, pady=10)
        self.labSel1.configure(text=champs1)
    

    def Test2(self):
        global champs2
        champs2 = str(v2.get())
        self.labSel2.grid(row=15, column=2, sticky=W, padx = 10, pady=10)
        self.labSel2.configure(text=champs2)
    

    def Quit(self):
        if (testQuit == 0):
            self.myParent.destroy()
        else:
            field_names = []
            QgsMapLayerRegistry.instance().removeMapLayer(couche.id())
            self.myParent.destroy()




root = Tk()
app = XlsxClass(root)
root.mainloop()
