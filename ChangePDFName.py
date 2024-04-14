# -*- coding: utf-8 -*-
"""
Created on Wed Apr  5 08:54:56 2023

@author: miguel.gomez
"""

import os
import fnmatch
import pandas as pd
import pdfquery
from bs4 import BeautifulSoup
import PySimpleGUI as sg
import sys


def window():
    """
        Establece la configuración de la ventana
        Inputs: 2 datos
            1. la ruta de la carpeta donde se encuentren los
                PDF a los que se les cambiara el nombre
                NOTA: No es necesario agregar la ruta con dos backslash '\\'
                    El programa lo hace solo
                Ejemplo:
                    C:\\Users\\user\\Desktop\\Confirmaciones Flujo
            
            2. El nombre "general" de los PDF sin la numeración
                Ejemplo:
                    Confirmaciones Flujo-
    """
    try:
        sg.theme('SandyBeach')     
        layout = [
            [sg.Text("Ingrese la ruta donde se encuentre la carpeta con los PDF")],
            [sg.Text('Ruta', size =(15, 1)), sg.InputText()],
            [sg.Text("Ingrese nombre PDF con guion al final Ejemplo: 'Nombre-'")],
            [sg.Text('Nombre PDF', size =(15, 1)), sg.InputText()],
            [sg.Submit(), sg.Cancel()]
        ]
        
        window = sg.Window("ChangePDFName 3000", layout)
        event, values = window.read()
        """
            values es un array que contene las respuestas del usuario
                [0] -> ruta
                [1] -> nombre de los PDF
        """
        main(values)
        window.close()
        sys.exit()
    
    except Exception as e:
        print(e)
        window.close()
        sys.exit()

def main(values):
    """
        main se encarga principalmente del formateo de los datos ingresados
        
    """
    root = str(values[0])
    
    # Formatea de un backlash a dos para que pueda ser leído por Python
    root = str(root).replace("\\", "\\\\")
    
    #root es la variable principal que contendra la ruta ya formateada
    root = root
    
    # Cuenta el úmero de PDFs en la carpeta seleccionada
    number_of_files = len(fnmatch.filter(os.listdir(root), "*.pdf"))
    
    #name_file contendrá el nombre general de los PDF
    name_file = str(values[1])
    extractNames(root, number_of_files, name_file)

def extractNames(root, number_of_files, name_file):
    """
        extractNames se encargá de convertir los PDF a archivos XML
        de forma que pueda obtenerse toda su información.
        
        Primero obtiene los nombres originales de los PDF y los guarda
        en una lista.
            Ejemplo:
                Confirmaciones Flujo-1.pdf
        
        Después en los XML se almacena toda la informacion del PDF
        y solamente se esocge el valor buscado 'Folio: ############'
        
        Después, se almacenan esos nombres si la palabra 'Folio' en una lista
        añadiendo/concatenando también la extension '.pdf' al final del
        valor del Folio
            Ejemplo:
                CM_ALTAMD_PIT-010_MAKC-PV0335-20.pdf
        
        Generando al final dos listas:
            name_file_concat_list: Nombres originales
            name_file_corrected: Nombres Corregidos (Nombres de los folio)
        
    """
    #variables que almacenan las extensiones para no re escribir
    pdf_exten = ".pdf"
    xml_exten = ".xml"
    
    #listas vacias donde se almacenaran los nombres originales y corregidos
    name_file_concat_list = []
    name_file_corrected = []
    
    #for de longitud 1 hasta el numero total de PDFs en la carpeta
    for i in range(1, number_of_files + 1):
        
        #Concatena el nombre original del pdf + numeracion (i) + extension
        name_file_concat = name_file + str(i) + pdf_exten
        
        #se añaden a una lista
        name_file_concat_list.append(name_file_concat)
        
        #Se carga el PDF (nombre de PDF con nmeracion, A-1, A-2, A-3, etc.)
        #Busca archivo PDF en la carpeta con el dato Ruta + nombre
        #Osea busca desde la ruta, no solamente el puro nombre
        #NOTA: Si no se busca con ruta marca ERROR
        pdf = pdfquery.PDFQuery(root + "\\\\" + name_file_concat)
        pdf.load()
        
        #Crea el XML con el nombre del archivo original + numeracion (i) + extension
        name_file_concat = name_file + str(i) + xml_exten
        pdf.tree.write(name_file_concat, pretty_print = True)
        
        #Busca en el PDF el tag 'LTTextBoxHorizontal' para recuperar TODOS los datos
        btree = BeautifulSoup(open(name_file_concat), "lxml-xml")
        texttags= btree.find_all("LTTextBoxHorizontal")
        
        """
            1.- De toda la informacion guardada del xml, solamente
                se selecciona la posición 1, ya que es ahí donde
                esta la info del Folio.
            
            2.- Se borra la palabra 'Folio: ' de lo que recuperado de la
                etiqueta Folio
            
            3.- Se formatea el folio por si tiene backlash se cambia
            a guion bajo, esto porque al cambiar automaticamente los nombres,
            Windows no permite nombres con backslash en el nombre.
            
            4.- Por ultimo se borra un espacio en blanco que algunos nombres
                tenian por defecto
            
            5.- Los puntos 3 y 4 son opcionales, osea que se ejecuta
                solamente si el dato contiene esos caracteres.
        """
        folio_temp = texttags.pop(1) #1
        folio_temp = folio_temp.string
        folio_temp = folio_temp.replace("Folio: ", "") #2
        folio_temp = folio_temp.replace("/", "-") #3
        folio_temp = folio_temp[:len(folio_temp)-1] #4
        
        #Se añade a la lisat de nombres ya corregidos mas la extension PDF
        name_file_corrected.append(folio_temp + ".pdf")
        
        #Se cierra la carga de archivos PDF
        pdf.file.close()
        
        #Se borran el archivo XML generado en la iteracion
        #ya que por el momento para otra cosa no se utiliza
        os.remove(name_file_concat)
    
    changeNames(root, number_of_files, name_file_concat_list,
                name_file_corrected)

def changeNames(root, number_of_files, name_file_concat_list,
                name_file_corrected):
    """
        changeNames se encarga de cambiar los nombres de los archivos
        usando las dos listas generadas previamente
            1.- name_file_concat_list = nombres originales
            2.- name_file_corrected = nombres corregidos
    """
    
    for i in range(0, number_of_files):
        oldName = os.path.join(root, name_file_concat_list[i]) #1
        newName = os.path.join(root, name_file_corrected[i]) #2
        os.rename(oldName, newName)

window()