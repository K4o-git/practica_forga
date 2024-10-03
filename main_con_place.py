import csv
import os
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
from tkcalendar import DateEntry

class win(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title("Rellenar documentos para imprimir")
        self.resizable(False, False)
        self.geometry("1100x750")

        barra_menus = Menu()
        menu = Menu(barra_menus, tearoff=False)
        menu2 = Menu(barra_menus, tearoff=False)
        barra_menus.add_cascade(menu=menu, label="Arquivo")
        barra_menus.add_cascade(menu=menu2, label="Opcions")
        menu.add_command(label="Gardar datos alumnos", command=lambda: self.save_sheet())
        menu.add_command(label="Cargar datos alumnos gardados", command=lambda: self.open_csv())
        menu.add_command(label="Guardar datos curso", command=lambda: self.gardar_datos_curso())
        menu.add_command(label="Cargar datos curso gardados", command=lambda: self.cargar_datos_curso())
        menu.add_command(label="Saír", command=lambda: sys.exit())
        self.config(menu=barra_menus)

        et_cod_ac_form = Label(self, text="Cód. Acc. Form.")
        et_cod_ac_form.place(x=30, y=20)
        self.cod_ac_form  = Entry(self, width=30)
        self.cod_ac_form.place(x=130, y=20)

        et_cod_espec = Label(self, text="Cód. Espec.")
        et_cod_espec.place(x=30, y=51)
        self.cod_espec= Entry(self, width=30)
        self.cod_espec.place(x=130, y=51)

        et_nom_curso = Label(self, text="Nome do curso")
        et_nom_curso.place(x=30, y=82)
        self.nom_curso = Entry(self, width=93)
        self.nom_curso.place(x=130, y=82)

        et_centro = Label(self, text="Centro")
        et_centro.place(x=30, y=113)
        self.centro = ttk.Combobox(self, width=90,
                                   values=["forGA Narón, sito en Polígono Río Pozo - Avda. Ferreiros, 173-174 15578 Narón",
                                           "forGA Compostela (Ames), sito en Polígono Industrial Novo Milladoiro Rúa Palmeiras 21A - 15895 Ames",
                                           "forGA Vigo, sito en Avda. Aeroporto 92 Baixo 36206 Vigo",
                                           "forGA A Coruña (Oleiros), sito en Avda. Rosalía de Castro, 94 15172 Oleiros",
                                           "forGA Ourense, sito en Río Mao, 27 baixo 32001 Ourense",
                                           "forGA Pontevedra (Ponte Caldelas), sito en Polígono Industrial A Reigosa Parcela 27 36828 Ponte Caldelas",
                                           ]
                                     )
        self.centro.place(x=130, y=113)
        
        # et_datas_curso = Label(self, text="Datas curso:")
        # et_datas_curso.place(x=30, y=160)
        # self.datas_curso = Entry(self, width=42,)
        # self.datas_curso.place(x=130, y=160)
        # self.datas_curso.insert(0, "23 de setembro de 2024 a 23 setembro de 2025")

        et_data_inicio = Label(self, text="Data inicio")
        et_data_inicio.place(x=30, y=144)
        self.data_inicio = DateEntry(self, width=11)
        self.data_inicio.place(x=130, y=144)
        self.data_inicio.config(date_pattern= "dd/mm/yyyy")

        et_data_finalizacion = Label(self, text="Data finalización")
        et_data_finalizacion.place(x=260, y=144)
        self.data_finalizacion = DateEntry(self, width=11)
        self.data_finalizacion.place(x=360, y=144)
        self.data_finalizacion.config(date_pattern= "dd/mm/yyyy")

        et_num_censo = Label(self, text="Número Censo")
        et_num_censo.place(x=30, y=175)
        self.num_censo = Entry(self, width=30)
        self.num_censo.place(x=130, y=175)

        self.meses = ["xaneiro", "febreiro", "marzo", "abril", "maio", "xuño", "xullo", "agosto", "setembro", "outubro", "novembro", "decembro"]

        self.documentos=["", "", "", "", "", ""]
        self.documentos[0]= "Ficha_de_alumno.docx"
        self.documentos[1]= "Dereitos_e_deberes.docx"
        self.documentos[2]= "Protección_de_datos.docx"
        self.documentos[3]= "Autorización_pegada_dixital.docx"
        self.documentos[4]= "Información_de_bolsas.docx"
        self.documentos[5]= "Autorización_datos_persoais.docx"
        #self.documentos[6]= "MC_Modelo autorización datos persoais_2.docx"
        #self.documentos[7]= "MC_Modelo autorización rexistro pegada dixital_gal.docx"

        self.checkbox_variable=[tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar(), tk.IntVar()]

        Label(self, text="FORMULARIOS", font='Helvetica 10 bold underline').place(x=870, y=5)
        self.checkbox1 = ttk.Checkbutton(self, text="Ficha de alumno", variable= self.checkbox_variable[0]).place(x=830, y=30)
        self.checkbox2 = ttk.Checkbutton(self, text="Dereitos e deberes", variable= self.checkbox_variable[1]).place(x=830, y=55)
        self.checkbox3 = ttk.Checkbutton(self, text="Protección de datos", variable= self.checkbox_variable[2]).place(x=830, y=80)
        self.checkbox4 = ttk.Checkbutton(self, text="Autorización pegada dixital", variable= self.checkbox_variable[3]).place(x=830, y=105)
        self.checkbox5 = ttk.Checkbutton(self, text="Información de bolsas", variable= self.checkbox_variable[4]).place(x=830, y=130)
        self.checkbox6 = ttk.Checkbutton(self, text="Autorización datos persoais", variable= self.checkbox_variable[5]).place(x=830, y=155)
        #self.checkbox7 = ttk.Checkbutton(self, text="Modelo autorización datos persoais 2(Ames)", variable= self.checkbox_variable[6]).place(x=680, y=135)
        #self.checkbox8 = ttk.Checkbutton(self, text="Mod. aut. rex. pegada dixital_gal(Ames)", variable= self.checkbox_variable[7]).place(x=680, y=155)

        self.checkbox_variable[0].set(0)
        self.checkbox_variable[1].set(0)
        self.checkbox_variable[2].set(0)
        self.checkbox_variable[3].set(0)
        self.checkbox_variable[4].set(0)
        self.checkbox_variable[5].set(0)
        #self.checkbox_variable[6].set(0)
        #self.checkbox_variable[7].set(0)

        self.bot_marcar_todos_formularios = ttk.Button(self, text="Seleccionar todos", command= self.marcar_todos_documentos).place(x=800, y=185)
        self.bot_desmarcar_todos_formularios = ttk.Button(self, text="Deseleccionar todos", command= self.desmarcar_todos_documentos).place(x=925, y=185)

        titulo_tabla = Label(self, text="LISTADO DE ALUMNOS", font=('bold', 12, 'underline'))
        titulo_tabla.place(x=450, y=220)

        self.sheet = tksheet.Sheet(self, width=1100, height=450, total_columns=4, total_rows=15, show_x_scrollbar=False, show_y_scrollbar=True)
        self.sheet.column_width(column=0, width=140)
        self.sheet.column_width(column=1, width=440)
        self.sheet.column_width(column=2, width=440)
        self.sheet.column_width(column=3, width=24)
        self.sheet.headers(["DNI", "Nome", "Apelidos", ""])
        self.sheet.checkbox("D", checked=False)
        self.sheet.enable_bindings()
        self.sheet.disable_bindings("column_width_resize", "rc_insert_column", "rc_delete_column")
        self.sheet.popup_menu_add_command("Cargar archivo csv", self.open_csv)
        self.sheet.popup_menu_add_command("Gardar datos a CSV", self.save_sheet)
        self.sheet.place(x=0, y=250)

        self.sheet_span = self.sheet.span(
            header=True,
            index=False,
            hdisp=False,
            idisp=False,
        )

        self.bot_marcar_todos_alum = ttk.Button(self, text="Seleccionar todos", command= self.marcar_todos_alumnos).place(x=725, y=712)
        self.bot_desmarcar_todos_alum = ttk.Button(self, text="Deseleccionar todos", command= self.desmarcar_todos_alumnos).place(x=845, y=712)

        self.valeirar_taboa = ttk.Button(self, width=20, text="Baleirar táboa", command=lambda: self.valeirar_listado_de_alumnos()).place(x=120, y= 712)
        self.botonOK = ttk.Button(self, width=20, text="Cubrir Documentos", command=lambda: self.prueba_generar_documentos()).place(x=320, y= 712)
        

    def marcar_todos_documentos(self):
        self.checkbox_variable[0].set(1)
        self.checkbox_variable[1].set(1)
        self.checkbox_variable[2].set(1)
        self.checkbox_variable[3].set(1)
        self.checkbox_variable[4].set(1)
        self.checkbox_variable[5].set(1)
        #self.checkbox_variable[6].set(1)
        #self.checkbox_variable[7].set(1)
        print(self.data_inicio.get_date().year)
        print(self.meses[self.data_inicio.get_date().month-1])
        print(self.data_inicio.get_date().day)

    def desmarcar_todos_documentos(self):
        self.checkbox_variable[0].set(0)
        self.checkbox_variable[1].set(0)
        self.checkbox_variable[2].set(0)
        self.checkbox_variable[3].set(0)
        self.checkbox_variable[4].set(0)
        self.checkbox_variable[5].set(0)
        #self.checkbox_variable[6].set(0)
        #self.checkbox_variable[7].set(0)

    
    def marcar_todos_alumnos(self):
        n=1
        for row in self.sheet.data:
            if self.sheet["A"+str(n)].data:
                self.sheet["D"+str(n)].data=True
            n += 1
            print(n)

    def desmarcar_todos_alumnos(self):
        n=1
        for row in self.sheet.data:
            self.sheet["D"+str(n)].data=False
            n += 1

    def valeirar_listado_de_alumnos(self):
        n=1
        for row in self.sheet.data:
            self.sheet["A"+str(n)].data=["", "", ""]
            n += 1
        
    def save_sheet(self):
        filepath = filedialog.asksaveasfilename(
            parent=self,
            title="Save sheet as",
            filetypes=[("CSV File", ".csv"), ("TSV File", ".tsv")],
            defaultextension=".csv",
            confirmoverwrite=True
        )
        if not filepath or not filepath.lower().endswith((".csv", ".tsv")):
            return
        try:
            file = open(filepath, 'w+')
            ##file.write("DNI;Nome;Apelidos;ok\n")
            #file.write(self.num_curso.get() + ";" + self.nom_curso.get() + ";" + "\n")
            for row in self.sheet.data:
                check = ''
                if (row[3]): check = str(row[3])
                file.write(row[0] + ";" + row[1] + ";" + row[2] + ";" + check + "\n")
        except FileNotFoundError:
            print('ERROR')

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
        
    def gardar_datos_curso(self):
        filepath = filedialog.asksaveasfilename(
            parent=self,
            title="Save sheet as",
            filetypes=[("CSV File", ".csv"), ("TSV File", ".tsv")],
            defaultextension=".csv",
            confirmoverwrite=True
        )
        if not filepath or not filepath.lower().endswith((".csv", ".tsv")):
            return
        try:
            file = open(filepath, 'w+')
            file.write(self.cod_ac_form.get() + ";" 
                       + self.cod_espec.get() + ";" 
                       + self.nom_curso.get() + ";"
                       + self.centro.get() + ";" 
                       #+ self.datas_curso.get() + ";"
                       + self.data_inicio.get() + ";"
                       + self.data_finalizacion.get() + ";"
                       + self.num_censo.get() + ";"  
                       )
        except FileNotFoundError:
            print('ERROR')

    def cargar_datos_curso(self):
        filepath = filedialog.askopenfilename(parent=self, title="Cargar archivo CSV")
        if not filepath or not filepath.lower().endswith((".csv", ".tsv")):
            return
        try:
            with open(Path(filepath), "r") as file:
                self.cod_ac_form.delete(0, tk.END) 
                self.cod_espec.delete(0, tk.END) 
                self.nom_curso.delete(0, tk.END)
                self.centro.delete(0, tk.END) 
                #self.datas_curso.delete(0, tk.END)
                self.num_censo.delete(0, tk.END) 
                for linea in file: 
                    self.cod_ac_form.insert(0, linea.split(';')[0]) 
                    self.cod_espec.insert(0, linea.split(';')[1])
                    self.nom_curso.insert(0, linea.split(';')[2])
                    self.centro.insert(0, linea.split(';')[3])
                    #self.datas_curso.insert(0, linea.split(';')[4])
                    self.data_inicio.insert(0, linea.split(';')[4])
                    self.data_finalizacion.insert(0, linea.split(';')[5])
                    self.num_censo.insert(0, linea.split(';')[6])
                
        except Exception as error:
            print(error)
            return   
        
    def prueba_generar_documentos(self):
        for row in self.sheet.data:
            try:
                if (row[0] == "") or (row[3] == None): break
                if row[3]:
                    dni = row[0]
                    nome = row[1]
                    apelidos = row[2]
                    cod_ac_form = self.cod_ac_form.get()
                    cod_espec = self.cod_espec.get()
                    nome_curso = self.nom_curso.get()
                    centro = self.centro.get()
                    #datas_curso = self.datas_curso.get()
                    dia_inicio = self.data_inicio.get_date().day
                    mes_inicio = self.meses[self.data_inicio.get_date().month-1]
                    ano_inicio = self.data_inicio.get_date().year
                    dia_finalizacion = self.data_finalizacion.get_date().day
                    mes_finalizacion= self.meses[self.data_finalizacion.get_date().month-1]
                    ano_finalizacion = self.data_finalizacion.get_date().year
                    num_censo = self.num_censo.get()

                    context = {
                        "DNI": dni,
                        "NOME": nome,
                        "APELIDOS": apelidos,
                        "COD_AC_FORM": cod_ac_form,
                        "COD_ESPEC": cod_espec,
                        "NOME_CURSO": nome_curso,
                        "CENTRO": centro,
                        #"DATAS_CURSO": datas_curso,
                        "DIA_INICIO": dia_inicio,
                        "MES_INICIO": mes_inicio,
                        "ANO_INICIO": ano_inicio,
                        "DIA_FINALIZACION": dia_finalizacion,
                        "MES_FINALIZACION": mes_finalizacion,
                        "ANO_FINALIZACION": ano_finalizacion,
                        "NUM_CENSO": num_censo,
                    }
                    # Crea las subcarpetas para cada alumno y cuarda dentro sus documentos modificados
                    os.makedirs(str("xerados/" + apelidos + " " + nome))
                    for file in os.listdir("modelos"):
                        for n in range(6):
                            if str(file) == self.documentos[n] and self.checkbox_variable[n].get():
                                if file.endswith("docx"):
                                    documento_path = Path(__file__).parent / str("modelos/" + file)
                                    doc = DocxTemplate(documento_path)
                                    doc.render(context)
                                    doc.save(Path(__file__).parent / str("xerados/" + apelidos + " " + nome + "/" + file))
                                    aaa = open(Path(__file__).parent / str("xerados/" + apelidos + " " + nome + "/" + file), encoding="utf-8")
                                    print(aaa.name)
                                    #win32api.ShellExecute(0, 'print', aaa.name, f'/d:"{win32print.GetDefaultPrinter()}"', '.', 0)
                                else:
                                    print("fallo")
            except Exception as e:
                print(e)

app = win()
app.mainloop()


               