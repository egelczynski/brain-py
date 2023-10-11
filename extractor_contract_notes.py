#!/usr/bin/env python
# coding: utf-8

# In[116]:


import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import re
import os

def generador_extracto(ruta_carpeta, nombre_extracto, destino):
    ruta_carpeta = r"{}".format(ruta_carpeta)
    nombre_extracto = nombre_extracto + ".xlsx"
    destino = r"{}\\".format(destino)
    
    def extracto(archivo):
        pdf = pdfplumber.open(archivo)
        return [page.extract_text().split("\n") for page in pdf.pages]
    
    def transformador(ext):
        textos = ["Name of Investor", "Capital Commitment", "Capital Call Value", "Shares issued", 'Outstanding Capital Commitment']
        data = {}
    
        for elemento in ext:
            for texto in textos:
                if texto in elemento:
                    key, value = elemento.split(":")
                    key = key.strip()
                    value = value.strip()
                    data.update({key:value})

        df = pd.DataFrame([data])[["Name of Investor", "Shares issued", "Capital Call Value", "Capital Commitment", "Outstanding Capital Commitment"]]

        if any(key in data for key in textos):
            df = pd.DataFrame([data])[["Name of Investor", "Shares issued", "Capital Call Value", "Capital Commitment", "Outstanding Capital Commitment"]]
        
            def str_to_float(s):
                numeric_str = re.sub(r'[^\d.]', '', s)
                return float(numeric_str)

            numeric_columns = ['Capital Commitment', 'Capital Call Value', 'Shares issued', 'Outstanding Capital Commitment']
            for col in numeric_columns:
                df[col] = df[col].apply(str_to_float)

            return df
        else:
            return pd.DataFrame()


    extracciones = []
    for filename in os.listdir(ruta_carpeta):
        if filename.endswith(".pdf"):
            file_path = os.path.join(ruta_carpeta, filename)
            extract_data = extracto(file_path)[0]
            df = transformador(extract_data)
            extracciones.append(df)
        if len(extracciones) == sum(1 for filename in os.listdir(ruta_carpeta) if filename.endswith(".pdf")):
            break
    
    df_final = pd.concat(extracciones, ignore_index=True)
        
    writer = pd.ExcelWriter(destino+nombre_extracto, engine="xlsxwriter")
    df_final.to_excel(writer, index=False)
    
    writer.save()
    writer.close()

class PDFToExcel:
    def __init__(self, master):
        self.master = master
        master.title("Extractor de PDF - Fondos")

        self.ruta_carpeta = tk.StringVar()
        self.nombre_extracto = tk.StringVar()
        self.destino = tk.StringVar()
        
        #Ruta de busqueda
        self.origen_label = tk.Label(master, text="Origen:")
        self.origen_label.grid(row=0, column=0)

        self.origen_entry = tk.Entry(master, textvariable=self.ruta_carpeta)
        self.origen_entry.grid(row=0, column=1)

        self.browse_button = tk.Button(master, text="Buscar", command=self.buscar_origen)
        self.browse_button.grid(row=0, column=2)

        # Nombre del extracto
        self.busqueda_label = tk.Label(master, text="Nombre del extracto:")
        self.busqueda_label.grid(row=1, column=0)

        self.busqueda_entry = tk.Entry(master, textvariable=self.nombre_extracto)
        self.busqueda_entry.grid(row=1, column=1)

        
        #Ruta de destino del extracto
        self.directorio_label = tk.Label(master, text="Destino:")
        self.directorio_label.grid(row=2, column=0)

        self.directorio_entry = tk.Entry(master, textvariable=self.destino)
        self.directorio_entry.grid(row=2, column=1)

        self.browse_button = tk.Button(master, text="Buscar", command=self.buscar_directorio)
        self.browse_button.grid(row=2, column=2)
        
        #Confirmación
        self.renombrar_button = tk.Button(master, text="Generar extracción", command=self.extractor)
        self.renombrar_button.grid(row=4, column=0, columnspan=3)

    def buscar_origen(self):
        origen = filedialog.askdirectory()
        self.ruta_carpeta.set(origen)
    
    def buscar_directorio(self):
        destino = filedialog.askdirectory()
        self.destino.set(destino)

    def extractor(self):
        ruta_carpeta = self.ruta_carpeta.get()
        nombre_extracto = self.nombre_extracto.get()
        destino = self.destino.get()
        generador_extracto(ruta_carpeta, nombre_extracto, destino)
        messagebox.showinfo(title="Extracción", message="Generado correctamente")

root = tk.Tk()
app = PDFToExcel(root)
root.mainloop()

