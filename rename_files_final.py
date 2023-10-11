#!/usr/bin/env python
# coding: utf-8

# In[2]:


import os
import tkinter as tk
from tkinter import filedialog, messagebox

def buscar_y_renombrar_archivos(ruta_directorio, termino_busqueda, termino_reemplazo):
    for nombre_archivo in os.listdir(ruta_directorio):
        ruta_archivo = os.path.join(ruta_directorio, nombre_archivo)
        if os.path.isdir(ruta_archivo):
            buscar_y_renombrar_archivos(ruta_archivo, termino_busqueda, termino_reemplazo)
        elif termino_busqueda in nombre_archivo:
            new_nombre_archivo = nombre_archivo.replace(termino_busqueda, termino_reemplazo)
            new_ruta_archivo = os.path.join(ruta_directorio, new_nombre_archivo)
            os.rename(ruta_archivo, new_ruta_archivo)
        

class FileRenamerApp:
    def __init__(self, master):
        self.master = master
        master.title("Renombrador de Archivos")

        self.ruta_directorio = tk.StringVar()
        self.termino_busqueda = tk.StringVar()
        self.termino_reemplazo = tk.StringVar()

        self.directorio_label = tk.Label(master, text="Directorio:")
        self.directorio_label.grid(row=0, column=0)

        self.directorio_entry = tk.Entry(master, textvariable=self.ruta_directorio)
        self.directorio_entry.grid(row=0, column=1)

        self.browse_button = tk.Button(master, text="Buscar", command=self.buscar_directorio)
        self.browse_button.grid(row=0, column=2)

        self.busqueda_label = tk.Label(master, text="Término de búsqueda:")
        self.busqueda_label.grid(row=1, column=0)

        self.busqueda_entry = tk.Entry(master, textvariable=self.termino_busqueda)
        self.busqueda_entry.grid(row=1, column=1)

        self.reemplazo_label = tk.Label(master, text="Término de reemplazo:")
        self.reemplazo_label.grid(row=2, column=0)

        self.reemplazo_entry = tk.Entry(master, textvariable=self.termino_reemplazo)
        self.reemplazo_entry.grid(row=2, column=1)

        self.renombrar_button = tk.Button(master, text="Renombrar archivos", command=self.renombrar_archivos)
        self.renombrar_button.grid(row=4, column=0, columnspan=3)

    def buscar_directorio(self):
        ruta_directorio = filedialog.askdirectory()
        self.ruta_directorio.set(ruta_directorio)

    def renombrar_archivos(self):
        ruta_directorio = self.ruta_directorio.get()
        termino_busqueda = self.termino_busqueda.get()
        termino_reemplazo = self.termino_reemplazo.get()
        buscar_y_renombrar_archivos(ruta_directorio, termino_busqueda, termino_reemplazo)
        messagebox.showinfo("Hecho", "Renombrado de archivos completado.")

root = tk.Tk()
app = FileRenamerApp(root)
root.mainloop()

