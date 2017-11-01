from docx import Document
from docx.shared import Inches
import win32com.client as win32 
import sys 

wordApp = win32.gencache.EnsureDispatch('Word.Application') #create a word application object
wordApp.Visible = False # hide the word application
doc = wordApp.Documents.Open("template.doc") # opening the template file


nombre = input("ingrese nombre ")
direc = input("ingrese direccion ")
rut = input ("ingrese rut ")
otsech = input("ingrgese ot sech ")
otasc = input("ingrse ot asc ")
modelo = input("ingrese de modelo ")
serie = input("ingrese serie ")
f1 = input("ingrese fecha primer ingreso ")
f2 = input("ingrese fecha segundo ingreso ")
fc = input("ingrese fecha de compra ")
p1 = input("ingrese problema 1a vez ")
p2 = input("ingrese problema 2a vez ")
falla1 = input ("ingrese 1a falla ")
falla2 = input ("ingrese 2a falla ")
r1 = input ("ingrese reparacion 1a vez ")
r2 = input("ingrese reparacion 2a vez ")


rng=doc.Bookmarks("nombre").Range # change the string Name to whatever name of your bookmarks
rng.InsertAfter(nombre)
rng=doc.Bookmarks("direccion").Range 
rng.InsertAfter(direc)
rng=doc.Bookmarks("rut").Range 
rng.InsertAfter(rut)
rng=doc.Bookmarks("otsech").Range 
rng.InsertAfter(otsech)
rng=doc.Bookmarks("otasc").Range 
rng.InsertAfter(otasc)
rng=doc.Bookmarks("modelo").Range 
rng.InsertAfter(modelo)
rng=doc.Bookmarks("serie").Range 
rng.InsertAfter(serie)
rng=doc.Bookmarks("f1").Range 
rng.InsertAfter(f1)
rng=doc.Bookmarks("f2").Range 
rng.InsertAfter(f2)
rng=doc.Bookmarks("fc").Range
rng.InsertAfter(fc)
rng=doc.Bookmarks("i1").Range 
rng.InsertAfter(p1)
rng=doc.Bookmarks("i2").Range 
rng.InsertAfter(p2)
rng=doc.Bookmarks("d1").Range
rng.InsertAfter(falla1)
rng=doc.Bookmarks("d2").Range 
rng.InsertAfter(falla2)
rng=doc.Bookmarks("r1").Range
rng.InsertAfter(r1)
rng=doc.Bookmarks("r2").Range 
rng.InsertAfter(r2)

#agregando imagenes

rng=doc.Bookmarks("image1").Range
rng.InlineShapes.AddPicture("c:/ima/{}/1.jpg" .format(otasc))

rng=doc.Bookmarks("image2").Range
rng.InlineShapes.AddPicture("c:/ima/{}/2.jpg" .format(otasc))

rng=doc.Bookmarks("image3").Range
rng.InlineShapes.AddPicture("c:/ima/{}/3.jpg" .format(otasc))



#guardando documento ....
print ('guardando documento')


doc.SaveAs("informe {}.doc" .format(otasc))