# Sofía Múnera Medina
# Python
# POO con archivos .xslx
# Fuentes:


import pandas as pd
from pandas import ExcelWriter
import openpyxl 

class Excel:

  def __init__(self):

    self.pais0='España'
    self.pais1='Estados Unidos'
    self.pais2='China'
  
  def leer(self):

    libro1=openpyxl.load_workbook('Libro1.xlsx')

    hoja1=libro1['Hoja1']

    capitales=hoja1['A2':'A4']
    idiomas=hoja1['B2':'B4']

    paises=[self.pais0,self.pais1,self.pais2]

    copia1=pd.DataFrame({'Pais':paises,'Capital':capitales,'Idioma':idiomas})

    archivo=ExcelWriter('copia1.xlsx')
    copia1.to_excel(archivo,'Hoja Copia',index=False)
    archivo.save()
    archivo.close()


a=Excel()
a.leer()