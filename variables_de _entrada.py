from docx import Document
from docx.shared import Inches
import win32com.client as win32 
import sys 

document = Document('informe.docx')

nombre = input("ingrese nombre")
rut = input ("ingrese rut")
otsech = input("ingrgese ot sech")
otasc = input("ingrse ot asc")
modelo = input("ingrese de modelo")
serie = input("ingrese serie")
f1 = input("ingrese fecha primer ingreso")
f2 = input("ingrese fecha segundo ingreso")
fc = input("ingrese fecha de compra")
p1 = input("ingrese problema 1a vez")
p2 = input("ingrese problema 2a vez")
falla1 = input ("ingrese 1a falla")
falla2 = input ("ingrese 2a falla")
r1 = input ("ingrese reparacion 1a vez")
r2 = input("ingrese reparacion 2a vez")

paragraph = document.add_paragraph(nombre)


document.save ('informe {}.docx' .format(otasc) )