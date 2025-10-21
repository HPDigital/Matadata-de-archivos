"""
Metadata de archivos
"""

#!/usr/bin/env python
# coding: utf-8

# In[7]:


import os
from datetime import datetime


# In[8]:


carpeta = r"C:\Users\HP\Downloads\Class_4545208_Assignment_45557830(2)\Brenda Lizarraga Velasco-8632757"


# In[9]:


metadatos = os.stat(carpeta)


# In[10]:


metadatos


# In[15]:


def timeConvert(atime):
  dt = atime
  newtime = datetime.fromtimestamp(dt)
  return newtime.date()

def sizeFormat(size):
    newform = format(size/1024, ".2f")
    return newform + " KB"

def createFileRecords(somepath):
    #dictionary
    firstDict = {}


    for name in os.listdir(somepath): 

        filepath = os.path.join(somepath, name)

        #main library that holds stats
        stats = os.stat(filepath)

        attrs = {
            'File Name': name,
            'Size (KB)': sizeFormat(stats.st_size),
            'Creation Date': timeConvert(stats.st_ctime),
            'Modified Date': timeConvert(stats.st_mtime),
            'Last Access Date': timeConvert(stats.st_atime),

        }


        firstDict[name] = attrs 


    return firstDict


def printDir(somepath):
 dictOfDicts = createFileRecords(somepath)

 for n, a in dictOfDicts.items():

    print(f"Displaying for file: '{n}':")
    for i, j in a.items():
        print(f"{i}: {j}")
    print()



if __name__ == "__main__":

    printDir(r"C:\Users\HP\Downloads\Class_4545208_Assignment_45557823(3)\Sebastian Mirko Rocha Azurduy-8664390")


# In[14]:


from openpyxl import load_workbook
ruta=r"C:\Users\HP\Downloads\Class_4545208_Assignment_45557823(3)\Sebastian Mirko Rocha Azurduy-8664390\Archivo_EXCEL_para_entregables_semana_4_(1).xlsx"
wb = load_workbook(ruta)
wb.properties


# In[ ]:




