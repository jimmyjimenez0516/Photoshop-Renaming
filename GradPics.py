"""" This imports all the libraries need for the project) """
import win32com.client
import os
import csv
from PIL import Image
import photoshop.api as ps
"""" This imports all the libraries need for the project) """



"""" This is the naming for the student name on PS and file name) """
f= open('csv file goes here','r') #insert the path of your CSV so that it can pull the name the files correctly
reader = csv.reader(f)
gradnames = {}
for row in reader:
    gradnames[row[0]]= row[4]
"""" This is the naming for the student name on PS and file name) """


"""" This opens PhotoShop with the correct Template and creates the variable for changing the text as well as includes FilePath for export """
psApp = win32com.client.Dispatch("Photoshop.Application")
psApp.Open(r"C:\Users\path to psd file goes here")
doc = psApp.Application.ActiveDocument
layer_facts = doc.ArtLayers["layer name goes here"]
text_of_layer = layer_facts.TextItem
options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
options.Format = 6   # JPEG
options.Quality = 100 # Value from 0-100
exportRoot = r"C:\Users\pathgoeshere"
"""" This open PhotoShop with the correct Template and creates the variable for changing the text as well as includes FilePath for export """

"""" This iterate through the whole dictionary which holds all of the names """
for key,value in gradnames.items():
    text_of_layer.contents = value
    filename= exportRoot+key+" "+value+".jpg"
    doc.Export(ExportIn=filename, ExportAs=2, Options=options)
"""" This iterate through the whole dictionary which holds all of the names"""
