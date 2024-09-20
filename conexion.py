import mysql.connector

class Conexion:

    host = ''
    port  = ''
    db = ''
    user = ''
    password = ''

    def __init__(self, host, port, db, user, password):
        self.host = host
        self.port = port
        self.db = db
        self.user = user
        self.password = password

    def conectar(self):
        conexion = None
        try :
            conexion = mysql.connector.connect(host=self.host, port=self.port, database=self.db, user=self.user, password=self.password)
        except Exception as e:
            print('Error al conectase a la Base de datos')
        return conexion

    def desconectar(self):
        pass