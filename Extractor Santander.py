#!/usr/bin/env python
# coding: utf-8

# In[17]:


import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def generador_extracto(ruta_archivo, nombre_extracto, destino):
    ruta_archivo = r"{}".format(ruta_archivo)
    nombre_extracto = nombre_extracto + ".xlsx"
    destino = r"{}\\".format(destino)
    
    def extracto(ruta_archivo):
        pdf = pdfplumber.open(ruta_archivo)
        return [page.extract_tables() for page in pdf.pages][3:-2]
    
    #Funciones ingles
    def monedas_eng(ext):
        lista = [elemento for page in ext for tabla in page for elemento in tabla if "Account Number" in tabla[0] and "Accrued\ninterest" not in tabla[0]]
                    
        if len(lista) >= 1:
            df = pd.DataFrame(lista[1:], columns=lista[0])
            df = df[~df.iloc[:, 0].isin([""])]
        else:
            df = pd.DataFrame()
        return df

    def renta_fija_eng(ext):
        lista = [elemento for page in ext for tabla in page for elemento in tabla if "Nominal" in tabla[0] and "ISIN" not in tabla[0]]
        
        if len(lista) >= 1:
            df = pd.DataFrame(lista[1:], columns=lista[0])
            df = df[~df.iloc[:, 0].isin(["Total", ""])]
        else:
            df = pd.DataFrame()
        
        if len(df) >= 1:
            df.reset_index(drop=True, inplace=True)
            df.columns = pd.Series(df.columns.values).apply(lambda x: x.replace("\n", " ") if isinstance(x, str) else x)
            df = df[["","Nominal","Current Price YTM Annual", "Market Value", "Accrued Interest"]]
            for index in range(len(df.values)):
                df.values[index] = pd.Series(df.values[index]).apply(lambda x: x.replace("\n", " ") if isinstance(x, str) else x)
    
            df["Nominal"] = df["Nominal"].apply(lambda x: x[:x.find(".")].replace(",",""))
            df["Current Price YTM Annual"] = df["Current Price YTM Annual"].apply(lambda x: x[:x.find(" ")])
            df["Market Value"] = df["Market Value"].apply(lambda x: x.replace(",",""))
            df["Accrued Interest"] = df["Accrued Interest"].apply(lambda x: x.replace(",",""))
    
            df[["Nominal", "Current Price YTM Annual", "Market Value", "Accrued Interest"]] = df[["Nominal", "Current Price YTM Annual", "Market Value", "Accrued Interest"]].astype(float)
        else:
            None
                    
        return df
    
    def renta_variable_eng(ext):
        lista = [elemento for page in ext for tabla in page for elemento in tabla if "NAV" in tabla[0] and "ISIN" not in tabla[0]]
                    
        if len(lista) >= 1:
            df = pd.DataFrame(lista[1:], columns=lista[0])
            df = df[~df.iloc[:, 0].isin(["Total", ""])]
        else:
            df = pd.DataFrame()
        
        if len(df) >= 1:
            df.reset_index(drop=True, inplace=True)
            df.columns = pd.Series(df.columns.values).apply(lambda x: x.replace("\n", " ") if isinstance(x, str) else x)
            df = df[["","Quantity","NAV","Market Value"]]
            
            df["Quantity"] = df["Quantity"].apply(lambda x: x.replace(",","") if isinstance(x, str) else x)
            df["NAV"] = df["NAV"].apply(lambda x: x.replace(",","") if isinstance(x, str) else x)
            df["Market Value"] = df["Market Value"].apply(lambda x: x.replace(",","") if isinstance(x, str) else x)
    
            df[["Quantity","NAV","Market Value"]] = df[["Quantity","NAV","Market Value"]].astype(float)
        else:
            None
        
        return df
    
    def acciones_eng(ext):
        lista = [elemento for page in ext for tabla in page for elemento in tabla if "Number of\nshares" in tabla[0] and "ISIN" not in tabla[0]]
                    
        if len(lista) >= 1:
            df = pd.DataFrame(lista[1:], columns=lista[0])
            df = df[~df.iloc[:, 0].isin(["Total", ""])]
        else:
            df = pd.DataFrame()
        
        if len(df) >= 1:
            df.reset_index(drop=True, inplace=True)
            df.columns = pd.Series(df.columns.values).apply(lambda x: x.replace("\n", " ") if isinstance(x, str) else x)
            df = df[["","Number of shares","Last price","Market Value"]]
            
            df["Number of shares"] = df["Number of shares"].apply(lambda x: x.replace(",","") if isinstance(x, str) else x)
            df["Last price"] = df["Last price"].apply(lambda x: x.replace(",","") if isinstance(x, str) else x)
            df["Market Value"] = df["Market Value"].apply(lambda x: x.replace(",","") if isinstance(x, str) else x)
        
            df[["Number of shares","Last price","Market Value"]] = df[["Number of shares","Last price","Market Value"]].astype(float)
        else:
            df = pd.DataFrame()
            
        return df
            
    def movimientos_eng(ext):
        lista = [elemento for page in ext for tabla in page for elemento in tabla if "Detail" in tabla[0]]
                    
        if len(lista) >= 1:
            df = pd.DataFrame(lista[1:], columns=lista[0])
            df = df[~df.iloc[:, 0].isin(["Total", "","Detail"])]
        else:
            df = pd.DataFrame()
        return df
        
    def ajustar_movs_eng(tabla):
        replace_dict = {"JAN":"01", "FEB":"02","MAR":"03","APR":"04","MAY":"05","JUN":"06","JUL":"07","AUG":"08","SEP":"09", "OCT":"10", "NOV":"11", "DEC":"12"}
                        
        if len(tabla) < 1:
            return tabla
        else:
            for column in tabla.columns:
                if column == "Value Date":
                    for row in tabla.index:
                        if tabla.loc[row, column] == "-":
                            None
                        else:
                            tabla.loc[row, column] = tabla.loc[row, column].replace("-","/")
                            tabla.loc[row, column] = tabla.loc[row, column].replace(tabla.loc[row, column][3:6], str(replace_dict[tabla.loc[row, column][3:6]]))
                    
                elif column == "Deposit" or column == "Withdraws":
                    tabla[column] = tabla[column].apply(lambda x: x.replace(",",""))

                elif column == "Detail":
                    tabla[column] = tabla[column].apply(lambda x: x.replace("\n"," "))

                elif column == "ccy.":
                    None

                else:
                    tabla.drop(column, axis=1, inplace=True)

            for row in tabla.index:
                if tabla.loc[row, "Deposit"] == "-":
                    tabla.loc[row,"Deposit"] = 0
                else:
                    tabla.loc[row,"Deposit"] = float(tabla.loc[row,"Deposit"])
                    
                if tabla.loc[row,"Withdraws"] == "-":
                    tabla.loc[row,"Withdraws"] = 0
                else:
                    tabla.loc[row,"Withdraws"] = float(tabla.loc[row,"Withdraws"])

            tabla["Monto"] = tabla["Deposit"] + tabla["Withdraws"]
            tabla = tabla[["Value Date","Detail","Monto","ccy."]]

        return tabla    
    
    #Funciones español
    def monedas(ext):
        lista = [elemento for page in ext for tabla in page for elemento in tabla if "Número de cuenta" in tabla[0] and "Intereses\ndevengados" not in tabla[0]]
                    
        if len(lista) >= 1:
            df = pd.DataFrame(lista[1:], columns=lista[0])
            df = df[~df.iloc[:, 0].isin([""])]
        else:
            df = pd.DataFrame()
        return df

    def renta_fija(ext):
        lista = [elemento for page in ext for tabla in page for elemento in tabla if "Nominal" in tabla[0] and "ISIN" not in tabla[0]]
        
        if len(lista) >= 1:
            df = pd.DataFrame(lista[1:], columns=lista[0])
            df = df[~df.iloc[:, 0].isin(["Total", ""])]
        else:
            df = pd.DataFrame()
        
        if len(df) >= 1:
            df.reset_index(drop=True, inplace=True)
            df.columns = pd.Series(df.columns.values).apply(lambda x: x.replace("\n", " ") if isinstance(x, str) else x)
            df = df[["","Nominal","Precio actual YTM actual", "Valor de mercado", "Intereses devengados"]]
            for index in range(len(df.values)):
                df.values[index] = pd.Series(df.values[index]).apply(lambda x: x.replace("\n", " ") if isinstance(x, str) else x)
    
            df["Nominal"] = df["Nominal"].apply(lambda x: x[:x.find(".")].replace(",",""))
            df["Precio actual YTM actual"] = df["Precio actual YTM actual"].apply(lambda x: x[:x.find(" ")])
            df["Valor de mercado"] = df["Valor de mercado"].apply(lambda x: x.replace(",",""))
            df["Intereses devengados"] = df["Intereses devengados"].apply(lambda x: x.replace(",",""))
    
            df[["Nominal", "Precio actual YTM actual", "Valor de mercado", "Intereses devengados"]] = df[["Nominal", "Precio actual YTM actual", "Valor de mercado", "Intereses devengados"]].astype(float)
        else:
            None
                    
        return df
    
    def renta_variable(ext):
        lista = [elemento for page in ext for tabla in page for elemento in tabla if "Número de\nparticipaciones" in tabla[0] and "ISIN" not in tabla[0]]
            
        if len(lista) >= 1:
            df = pd.DataFrame(lista[1:], columns=lista[0])
            df = df[~df.iloc[:, 0].isin(["Total", ""])]
        else:
            df = pd.DataFrame()
        
        if len(df) >= 1:
            df.reset_index(drop=True, inplace=True)
            df.columns = pd.Series(df.columns.values).apply(lambda x: x.replace("\n", " ") if isinstance(x, str) else x)
            df = df[["","Número de participaciones","NAV","Valor de mercado"]]
            
            df["Número de participaciones"] = df["Número de participaciones"].apply(lambda x: x.replace(",","") if isinstance(x, str) else x)
            df["NAV"] = df["NAV"].apply(lambda x: x.replace(",","") if isinstance(x, str) else x)
            df["Valor de mercado"] = df["Valor de mercado"].apply(lambda x: x.replace(",","") if isinstance(x, str) else x)
    
            df[["Número de participaciones","NAV","Valor de mercado"]] = df[["Número de participaciones","NAV","Valor de mercado"]].astype(float)
        else:
            None
        
        return df
    
    def acciones(ext):
        lista = [elemento for page in ext for tabla in page for elemento in tabla if "Número de\nacciones" in tabla[0] and "ISIN" not in tabla[0]]
        
        if len(lista) >= 1:
            df = pd.DataFrame(lista[1:], columns=lista[0])
            df = df[~df.iloc[:, 0].isin(["Total", ""])]
        else:
            df = pd.DataFrame()
        
        if len(df) >= 1:
            df.reset_index(drop=True, inplace=True)
            df.columns = pd.Series(df.columns.values).apply(lambda x: x.replace("\n", " ") if isinstance(x, str) else x)
            df = df[["","Número de acciones","Último precio","Valor de mercado"]]
            
            df["Número de acciones"] = df["Número de acciones"].apply(lambda x: x.replace(",","") if isinstance(x, str) else x)
            df["Último precio"] = df["Último precio"].apply(lambda x: x.replace(",","") if isinstance(x, str) else x)
            df["Valor de mercado"] = df["Valor de mercado"].apply(lambda x: x.replace(",","") if isinstance(x, str) else x)
        
            df[["Número de acciones","Último precio","Valor de mercado"]] = df[["Número de acciones","Último precio","Valor de mercado"]].astype(float)
        else:
            df = pd.DataFrame()
            
        return df
            
    def movimientos(ext):
        lista = [elemento for page in ext for tabla in page for elemento in tabla if "Detalle" in tabla[0]]
        
        if len(lista) >= 1:
            df = pd.DataFrame(lista[1:], columns=lista[0])
            df = df[~df.iloc[:, 0].isin(["Total", "","Detalle"])]
        else:
            df = pd.DataFrame()
        return df
        
    def ajustar_movs(tabla):
        replace_dict = {"ENE":"01", "FEB":"02", "MAR":"03", "ABR":"04", "MAY":"05", "JUN":"06", "JUL":"07", "AGO":"08", "SEP":"09", "OCT":"10", "NOV":"11", "DIC":"12"}
            
        if len(tabla) < 1:
            return tabla
        else:
            for column in tabla.columns:
                if column == "Fecha valor":
                    for row in tabla.index:
                        if tabla.loc[row, column] == "-":
                            None
                        else:
                            tabla.loc[row, column] = tabla.loc[row, column].replace("-","/")
                            tabla.loc[row, column] = tabla.loc[row, column].replace(tabla.loc[row, column][3:6], str(replace_dict[tabla.loc[row, column][3:6]]))
                    
                elif column == "Abonos" or column == "Débitos":
                    tabla[column] = tabla[column].apply(lambda x: x.replace(",",""))

                elif column == "Detalle":
                    tabla[column] = tabla[column].apply(lambda x: x.replace("\n"," "))

                elif column == "Divisa":
                    None

                else:
                    tabla.drop(column, axis=1, inplace=True)

            for row in tabla.index:
                if tabla.loc[row, "Abonos"] == "-":
                    tabla.loc[row,"Abonos"] = 0
                else:
                    tabla.loc[row,"Abonos"] = float(tabla.loc[row,"Abonos"])
                    
                if tabla.loc[row,"Débitos"] == "-":
                    tabla.loc[row,"Débitos"] = 0
                else:
                    tabla.loc[row,"Débitos"] = float(tabla.loc[row,"Débitos"])

            tabla["Monto"] = tabla["Abonos"] + tabla["Débitos"]
            tabla = tabla[["Fecha valor","Detalle","Monto","Divisa"]]

        return tabla
        
    if "Account Activity" in extracto(ruta_archivo)[0][0][0][0]:
        df1 = renta_fija_eng(extracto(ruta_archivo))
        df2 = renta_variable_eng(extracto(ruta_archivo))
        df3 = acciones_eng(extracto(ruta_archivo))
        df5 = monedas_eng(extracto(ruta_archivo))
        if len(movimientos_eng(extracto(ruta_archivo))) != 0:
            df4 = ajustar_movs_eng(movimientos_eng(extracto(ruta_archivo)))
        else:
            df4 = pd.DataFrame()
        
    
    else:
        df1 = renta_fija(extracto(ruta_archivo))
        df2 = renta_variable(extracto(ruta_archivo))
        df3 = acciones(extracto(ruta_archivo))
        df5 = monedas(extracto(ruta_archivo))
        if len(movimientos(extracto(ruta_archivo))) != 0:
            df4 = ajustar_movs(movimientos(extracto(ruta_archivo)))
        else:
            df4 = pd.DataFrame()
        

        
    writer = pd.ExcelWriter(destino+nombre_extracto, engine="xlsxwriter")
    df1.to_excel(writer, sheet_name="Renta fija")
    df2.to_excel(writer, sheet_name="Renta variable")
    df3.to_excel(writer, sheet_name="Acciones")
    df4.to_excel(writer, sheet_name="Movimientos")
    df5.to_excel(writer, sheet_name="Monedas")

    formato_millares = writer.book.add_format({'num_format': '#,##0.00'})

    for column in ["Renta fija","Renta variable","Acciones"]:
        hoja = writer.sheets[column]
        hoja.set_column('B:F', None, formato_millares)
    
    writer.save()
    writer.close()

class PDFToExcel:
    def __init__(self, master):
        self.master = master
        master.title("Extractor de PDF")

        self.ruta_archivo = tk.StringVar()
        self.nombre_extracto = tk.StringVar()
        self.destino = tk.StringVar()
        
        #Ruta del PDF
        self.directorio_label = tk.Label(master, text="Ruta del PDF:")
        self.directorio_label.grid(row=0, column=0)

        self.directorio_entry = tk.Entry(master, textvariable=self.ruta_archivo)
        self.directorio_entry.grid(row=0, column=1)

        self.browse_button = tk.Button(master, text="Buscar", command=self.buscar_archivo)
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
        self.renombrar_button = tk.Button(master, text="Generar extracto", command=self.extractor)
        self.renombrar_button.grid(row=4, column=0, columnspan=3)

    def buscar_archivo(self):
        ruta_archivo = filedialog.askopenfilename()
        self.ruta_archivo.set(ruta_archivo)
        
    def buscar_directorio(self):
        destino = filedialog.askdirectory()
        self.destino.set(destino)

    def extractor(self):
        ruta_archivo = self.ruta_archivo.get()
        nombre_extracto = self.nombre_extracto.get()
        destino = self.destino.get()
        generador_extracto(ruta_archivo, nombre_extracto, destino)
        messagebox.showinfo(title="Extracción", message="Generado correctamente")

root = tk.Tk()
app = PDFToExcel(root)
root.mainloop()

