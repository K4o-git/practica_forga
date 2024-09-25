import os
import sys
from pathlib import Path

import pandas as pd
import tksheet
from docxtpl import DocxTemplate
import win32api
import win32print
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tksheet import Sheet


class win(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title("Rellenar documentos para imprimir")
        self.resizable(False, False)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.frame = tk.Frame(self)
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_rowconfigure(0, weight=1)
        self.sheet = tksheet.Sheet(self, width=980, height=450, total_columns=4, total_rows=15 , show_x_scrollbar=False, show_y_scrollbar=True)
        self.sheet.change_theme("dark")
        self.sheet.column_width(column=0, width=140)
        self.sheet.column_width(column=1, width=380)
        self.sheet.column_width(column=2, width=380)
        self.sheet.column_width(column=3, width=24)
        self.sheet.headers(["DNI", "Nome", "Apelidos", "✅"])
        self.sheet.checkbox("D", checked=False)
        self.sheet.enable_bindings()
        self.sheet.disable_bindings("column_width_resize")
        self.frame.grid(row=1, column=0, sticky="nswe")
        self.sheet.grid(row=1, column=0, sticky="nswe")


        #CAMPOS DE TEXTO PARA LOS DATOS DEL CURSO
        self.frame2 = tk.Frame(self)
        self.frame2.grid_columnconfigure(2, weight=0)
        self.frame2.grid(row=0, column=0, sticky="nswe", padx=30, pady=10)

        et_num_curso = Label(self.frame2, text="Nº do curso")
        et_num_curso.grid(row=0, column=0, padx=10, pady=10)
        num_curso  = Entry(self.frame2, width=42)
        num_curso.grid(row=0, column=1, padx=10, pady=10)

        et_nom_curso = Label(self.frame2, text="Nome do curso")
        et_nom_curso.grid(row=1, column=0, padx=10, pady=10)
        nom_curso = Entry(self.frame2, width=42)
        nom_curso.grid(row=1, column=1, padx=10, pady=10)

        et_censo = Label(self.frame2, text="Nº Censo")
        et_censo.grid(row=2, column=0, padx=10, pady=10)
        censo = Entry(self.frame2, width=42)
        censo.grid(row=2, column=1, padx=10, pady=10)

        et_centro = Label(self.frame2, text="Centro")
        et_centro.grid(row=3, column=0, padx=10, pady=10)
        centro = Entry(self.frame2, width=42)
        centro.grid(row=3, column=1, padx=10, pady=10)

        titulo_tabla = Label(self.frame2, text="LISTADO DE ALUMNOS", font=('bold', 12, 'underline')).grid(row=4, column=3)


        #BOTONES DE SELECCIÓN DE DOCUMENTOS
        self.frame3 = tk.Frame(self.frame2)
        self.frame3.grid(row=0, column=4, padx=10, sticky="nswe", rowspan=4)

        self.checkbox1 = ttk.Checkbutton(self.frame3, text="Ficha de alumno").grid(row=0, sticky="w")
        self.checkbox2 = ttk.Checkbutton(self.frame3, text="Dereitos e deberes").grid(row=1, sticky="w")
        self.checkbox3 = ttk.Checkbutton(self.frame3, text="Protección de datos").grid(row=2, sticky="w")
        self.checkbox4 = ttk.Checkbutton(self.frame3, text="Rexistro pegada dixital").grid(row=3, sticky="w")
        self.checkbox5 = ttk.Checkbutton(self.frame3, text="Información bolsas").grid(row=4, sticky="w")
        self.checkbox6 = ttk.Checkbutton(self.frame3, text="Modelo autorización datos persoais").grid(row=5, sticky="w")
        self.checkbox7 = ttk.Checkbutton(self.frame3, text="Modelo autorización datos persoais 2").grid(row=6, sticky="w")
        self.checkbox8 = ttk.Checkbutton(self.frame3, text="Modelo autorización rexistro pegada dixital_gal").grid(row=7, sticky="w")

        botonOK = Button(self, width=20, text="LISTO", command=lambda: print('aaa')).grid(row=2, column=0, sticky="n", rowspan=2)

        barra_menus = Menu()
        menu = Menu(barra_menus, tearoff=False)
        barra_menus.add_cascade(menu=menu, label="Archivo")
        menu.add_command(label="Guardar datos", command=lambda: print('nada'))
        menu.add_command(label="Cargar datos guardados", command=lambda: print('nada'))
        self.config(menu=barra_menus)

        #print(self.sheet.get_data())

app = win()
app.mainloop()

def prueba_generar_documentos():
    archivo_alumnos = open('alumnos.csv', encoding="utf-8")
    for linea in archivo_alumnos:
        if (linea.split(";")[0].strip() == 'Nome'): continue
        nome = linea.split(";")[0].strip()
        apelidos = linea.split(";")[1].strip()
        dni = linea.split(";")[2].strip()
        id_curso = ''
        nome_curso = ''
        data_alta = ''
        censo = ''

        context = {
            "DNI": dni,
            "NOME": nome,
            "APELIDOS": apelidos,
            "ID_CURSO": id_curso,
            "NOME_CURSO": nome_curso,
            "DATA_ALTA": data_alta,
            "CENSO": censo
        }
        # Crea las subcarpetas para cada alumno y cuarda dentro sus documentos modificados
        os.makedirs(str("generados/" + apelidos + " " + nome))
        for file in os.listdir("./plantillas"):
            if file.endswith("docx"):
                documento_path = Path(__file__).parent / str("plantillas/" + file)
                doc = DocxTemplate(documento_path)
                doc.render(context)
                doc.save(Path(__file__).parent / str("generados/" + apelidos + " " + nome + "/" + file))
                aaa = open(Path(__file__).parent / str("generados/" + apelidos + " " + nome + "/" + file), encoding="utf-8")
                print(aaa.name)
                #win32api.ShellExecute(0, 'print', aaa.name, f'/d:"{win32print.GetDefaultPrinter()}"', '.', 0)
            else:
                print("fallo")
