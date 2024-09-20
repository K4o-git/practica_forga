import os
import sys
from conexion import Conexion
from pathlib import Path
from docxtpl import DocxTemplate

#Datos de conexión a la base de datos
host='localhost'
port='3306'
db='cursos_forga'
user='root'
password=''

con = Conexion(host,port,db,user,password).conectar()

#Consulta que saca todos los alumnos del curso indicado y rellena cada documento con sus datos, si existen.
cursor = con.cursor()
sql = "SELECT a.dni, a.nome, a.apelidos, ac.id_curso, c.nome as 'nome_curso', c.data_alta, c.n_censo FROM alumno AS a JOIN alumno_curso AS ac ON a.dni = ac.dni_alumno JOIN curso AS c ON c.id_curso = ac.id_curso WHERE ac.id_curso = 'IFCD0112';"
cursor.execute(sql)
lista = cursor.fetchall()
if (lista.__len__() == 0):
    print("No hay datos que mostrar")
else:
    for c in lista:
        dni = str(c[0])
        nome = str(c[1])
        apelidos = str(c[2])
        id_curso = str(c[3])
        nome_curso = str(c[4])
        data_alta = str(c[5])
        censo = str(c[6])

        context = {
            "DNI": dni,
            "NOME": nome,
            "APELIDOS": apelidos,
            "ID_CURSO" : id_curso,
            "NOME_CURSO" : nome_curso,
            "DATA_ALTA" : data_alta,
            "CENSO" : censo
        }
        # Crea las subcarpetas para cada alumno y cuarda dentro sus documentos modificados
        os.makedirs(str("generados/" + apelidos + " " + nome))
        for file in os.listdir("./plantillas"):
            if file.endswith("docx"):
                documento_path = Path(__file__).parent / str("plantillas/" + file)
                doc = DocxTemplate(documento_path)
                doc.render(context)
                doc.save(Path(__file__).parent / str("generados/" + apelidos + " " + nome + "/" + file))
            else:
                print("fallo")
cursor.close()
con.close()

