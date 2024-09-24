import os
import sys
from pathlib import Path
from docxtpl import DocxTemplate
import win32api
import win32print

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




