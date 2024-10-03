import csv
import os
import shutil
import sys
import io
from pathlib import Path

import pandas as pd
import tksheet
from docxtpl import DocxTemplate
import win32api
import win32print
import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog
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
        self.sheet = tksheet.Sheet(self, width=1100, height=450, total_columns=4, total_rows=15 , show_x_scrollbar=False, show_y_scrollbar=True)
        self.sheet.column_width(column=0, width=140)
        self.sheet.column_width(column=1, width=440)
        self.sheet.column_width(column=2, width=440)
        self.sheet.column_width(column=3, width=24)
        self.sheet.headers(["DNI", "Nome", "Apelidos", "ok"])
        self.sheet.checkbox("D", checked=False)
        self.sheet.enable_bindings()
        self.sheet.disable_bindings("column_width_resize", "rc_insert_column", "rc_delete_column")
        self.sheet.popup_menu_add_command("Cargar archivo csv", self.open_csv)
        self.sheet.popup_menu_add_command("Guardar datos a CSV", self.save_sheet)
        self.frame.grid(row=1, column=0, sticky="nswe")
        self.sheet.grid(row=1, column=0, sticky="nswe")

        self.sheet_span = self.sheet.span(
            header=True,
            index=False,
            hdisp=False,
            idisp=False,
        )

        #CAMPOS DE TEXTO PARA LOS DATOS DEL CURSO
        self.frame2 = tk.Frame(self)
        self.frame2.grid_columnconfigure(2, weight=0)
        self.frame2.grid(row=0, column=0, sticky="nswe", padx=30, pady=20)

        et_num_curso = Label(self.frame2, text="Nº do curso")
        et_num_curso.grid(row=0, column=0, padx=10, pady=2)
        self.num_curso = Entry(self.frame2, width=54)
        self.num_curso.grid(row=0, column=1, padx=10, pady=2)

        et_cod_curso = Label(self.frame2, text="Cod. especialidade")
        et_cod_curso.grid(row=1, column=0, padx=10, pady=2)
        self.cod_curso = Entry(self.frame2, width=54)
        self.cod_curso.grid(row=1, column=1, padx=10, pady=2)

        et_nom_curso = Label(self.frame2, text="Nome do curso")
        et_nom_curso.grid(row=2, column=0, padx=10, pady=2)
        self.nom_curso = Entry(self.frame2, width=54)
        self.nom_curso.grid(row=2, column=1, padx=10, pady=2)

        et_centro = Label(self.frame2, text="Centro")
        et_centro.grid(row=3, column=0, padx=10, pady=2)
        self.centro = Entry(self.frame2, width=54)
        self.centro.grid(row=3, column=1, padx=10, pady=2)

        et_censo = Label(self.frame2, text="Nº Censo")
        et_censo.grid(row=4, column=0, padx=10, pady=2)
        self.censo = Entry(self.frame2, width=54)
        self.censo.grid(row=4, column=1, padx=10, pady=2)

        titulo_tabla = Label(self.frame2, text="LISTADO DE ALUMNOS", font=('bold', 12, 'underline')).grid(row=5, column=3)

        #BOTONES DE SELECCIÓN DE DOCUMENTOS
        self.frame3 = tk.Frame(self.frame2)
        self.frame3.grid(row=0, column=4, padx=10, sticky="nswe", rowspan=4)

        self.docCheck = []
        self.c1 = tk.IntVar()
        self.c2 = tk.IntVar()
        self.c3 = tk.IntVar()
        self.c4 = tk.IntVar()
        self.c5 = tk.IntVar()
        self.c6 = tk.IntVar()
        self.c7 = tk.IntVar()
        self.c8 = tk.IntVar()
        self.docCheck.append(self.c1)
        self.docCheck.append(self.c2)
        self.docCheck.append(self.c3)
        self.docCheck.append(self.c4)
        self.docCheck.append(self.c5)
        self.docCheck.append(self.c6)
        self.docCheck.append(self.c7)
        self.docCheck.append(self.c8)
        self.check1 = ttk.Checkbutton(self.frame3, variable=self.c1, text="Ficha de alumno").grid(row=0, pady=2, sticky="w")
        self.check2 = ttk.Checkbutton(self.frame3, variable=self.c2, text="Dereitos e deberes").grid(row=1, pady=2, sticky="w")
        self.check3 = ttk.Checkbutton(self.frame3, variable=self.c3, text="Protección de datos").grid(row=2, pady=2, sticky="w")
        self.check4 = ttk.Checkbutton(self.frame3, variable=self.c4, text="Rexistro pegada dixital").grid(row=3, pady=2, sticky="w")
        self.check5 = ttk.Checkbutton(self.frame3, variable=self.c5, text="Información bolsas").grid(row=4, pady=2, sticky="w")
        self.check6 = ttk.Checkbutton(self.frame3, variable=self.c6, text="Modelo autorización datos persoais - Narón").grid(row=5, pady=2, sticky="w")
        self.check7 = ttk.Checkbutton(self.frame3, variable=self.c7, text="Modelo autorización datos persoais - Ames").grid(row=6, pady=2, sticky="w")
        self.check8 = ttk.Checkbutton(self.frame3, variable=self.c8, text="Modelo autorización rexistro pegada dixital_gal").grid(row=7, pady=2, sticky="w")

        botonOK = Button(self, width=30, height=2, text="LISTO", command=lambda: self.prueba_impresion()).grid(row=2, column=0, sticky="n", rowspan=2, pady=20)

        barra_menus = Menu()
        menu = Menu(barra_menus, tearoff=False)
        menu2 = Menu(barra_menus, tearoff=False)
        barra_menus.add_cascade(menu=menu, label="Archivo")
        barra_menus.add_cascade(menu=menu2, label="Opciones")
        menu.add_command(label="Guardar datos", command=lambda: self.save_sheet())
        menu.add_command(label="Cargar datos guardados", command=lambda: self.open_csv())
        menu.add_command(label="Salir", command=lambda: sys.exit())
        self.config(menu=barra_menus)
        submenu = Menu(menu2, tearoff=False)
        menu2.add_command(label="Seleccionar todos", command=lambda: {self.marcar()})

    def marcar(self):
        for ok in self.docCheck:
            if (ok.get()):
                ok.set(0)
            else:
                ok.set(1)
            print(ok.get())

    #Guarda los datos de la tabla en el estado actual, excepto los valores de la tabla de checkbox.
    def save_sheet(self):
        filepath = filedialog.asksaveasfilename(
            parent=self,
            title="Save sheet as",
            filetypes=[("CSV File", ".csv"), ("TSV File", ".tsv")],
            defaultextension=".csv",
            confirmoverwrite=True,
        )
        if not filepath or not filepath.lower().endswith((".csv", ".tsv")):
            return
        try:
            file = open(filepath, 'w+')
            #file.write("DNI;Nome;Apelidos;ok\n")
            for row in self.sheet.data:
                check = ''
                if (row[3]): check = str(row[3])
                file.write(row[0] + ";" + row[1] + ";" + row[2] + ";" + check + "\n")
        except FileNotFoundError:
            print('ERROR')

    #Carga los datos en la tabla. Ya no da error al seleccionar o deseleccionar en la columna de checkbox
    def open_csv(self):
        filepath = filedialog.askopenfilename(parent=self, title="Cargar archivo CSV")
        if not filepath or not filepath.lower().endswith((".csv", ".tsv")):
            return
        try:
            with open(Path(filepath), "r") as file:
                self.sheet.data.clear()
                i = 1
                for linea in file:
                    dni = linea.split(';')[0]
                    nome = linea.split(';')[1]
                    apelidos = linea.split(';')[2]
                    self.sheet["A" + str(i)].data = [dni, nome, apelidos, False]
                    i += 1
        except Exception as error:
            print(error)
            return

    #Función para rellenar los ducumentos a partir de los datos de la tabla y los campos de texto. Por ahora no va con selección, se generan todos.
    def prueba_generar_documentos(self):
        for row in self.sheet.data:
            try:
                if (row[0] == "") or (row[3] == False): continue
                dni = row[0]
                nome = row[1]
                apelidos = row[2]
                id_curso = self.num_curso.get()
                nome_curso = self.nom_curso.get()
                centro = ''
                censo = self.censo.get()

                context = {
                    "DNI": dni,
                    "NOME": nome,
                    "APELIDOS": apelidos,
                    "ID_CURSO": id_curso,
                    "NOME_CURSO": nome_curso,
                    "CENTRO": centro,
                    "CENSO": censo
                }
                # Crea las subcarpetas para cada alumno y cuarda dentro sus documentos modificados
                os.makedirs(str("generados/" + apelidos + " " + nome), exist_ok=True)
                for file in os.listdir("./plantillas"):
                    if file.endswith("docx"):
                        documento_path = Path(__file__).parent / str("plantillas/" + file)
                        doc = DocxTemplate(documento_path)
                        doc.render(context)
                        doc.save(Path(__file__).parent / str("generados/" + apelidos + " " + nome + "/" + file))
                        archivo = open(Path(__file__).parent / str("generados/" + apelidos + " " + nome + "/" + file), encoding="utf-8")
                        print(archivo.name)
                        #win32api.ShellExecute(0, 'print', archivo.name, f'/d:"{win32print.GetDefaultPrinter()}"', '.', 0)
                    else:
                        print("fallo")
            except Exception as e:
                print(e)

    def prueba_impresion(self):
        self.prueba_generar_documentos()
        docs = []
        for row in self.sheet.data:
            nome = row[1]
            apelidos = row[2]
            if (row[0] == "") or (row[3] == False): continue
            if (self.c1.get() == 1): docs.append(str("generados/" + apelidos + " " + nome) + "/Plantilla_Ficha_alumn_AFD.docx")
            if (self.c2.get() == 1): docs.append(str("generados/" + apelidos + " " + nome) + "/Plantilla_dereitos-deberes_2.docx")
            if (self.c3.get() == 1): docs.append(str("generados/" + apelidos + " " + nome) + "/Plantilla_proteccion-datos.docx")
            if (self.c4.get() == 1): docs.append(str("generados/" + apelidos + " " + nome) + "/Plantilla_rexistro-pegada_2.docx")
            if (self.c5.get() == 1): docs.append(str("generados/" + apelidos + " " + nome) + "/Plantilla_Informacion-Bolsas.docx")
            if (self.c6.get() == 1): docs.append(str("generados/" + apelidos + " " + nome) + "/Plantilla_Modelo autorización datos persoais.docx")
            if (self.c7.get() == 1): docs.append(str("generados/" + apelidos + " " + nome) + "/Plantilla_Modelo autorización datos persoais_2.docx")
            if (self.c8.get() == 1): docs.append(str("generados/" + apelidos + " " + nome) + "/Plantilla_Modelo autorización rexistro pegada dixital_gal.docx")
            os.makedirs(str("imprimir/" + apelidos + " " + nome), exist_ok=True)
            for doc in docs:
                shutil.copy(doc, "imprimir/" + apelidos + " " + nome)

app = win()
app.mainloop()
