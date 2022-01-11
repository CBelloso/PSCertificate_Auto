#Importando librerias
import win32com.client 
import os 

#Abriendo Photoshop
psApp = win32com.client.Dispatch("Photoshop.Application")

#Abriendo Archivo

psApp.Open(r"C:\Users\mauri\Documents\VSCode\PythonPS\DiplomaAutoCADPES2021.psd")

#Opciones de Exportacion
options = win32com.client.Dispatch("Photoshop.PDFSaveOptions")
options.optimizeForWeb = True
options.preserveEditing = False
options.jpegQuality = 12 

#Diccionario
dict_sample = {"Primero":"Belloso Soriano",
               "Segundo":"Carlos Mauricio",
               "Tercero":"Carlos Mauricio Belloso Soriano"}

#Ruta de Almacenamiento
pdffile = "C:/Users/mauri/Documents/VSCode/PythonPS/"

#Cambiando capa de texto
doc = psApp.Application.ActiveDocument 

layer_facts = doc.ArtLayers["Nombre"]
text_of_layer = layer_facts.TextItem
#text_of_layer.contents = "Carlos Belloso"

for key,value in dict_sample.items():
    
    # Replace Text of text Layer
    
    text_of_layer.contents = value
    
    fileName = pdffile + key + ".pdf"  
    doc.SaveAs(SaveIn=fileName, Options=options)


