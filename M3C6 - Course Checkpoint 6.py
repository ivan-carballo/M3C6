from random import *
from openpyxl import *
from openpyxl.compat import *
from openpyxl.utils import *


class User:
    
    def __init__(self, nombre_usuario):
        self.nombre_usuario = nombre_usuario
        print(f'Su nombre es "{self.nombre_usuario}".')
        conteo = 0


    
    def pass_buscar(self, conteo):
        self.conteo = conteo

        # Carga de Excel con su hoja
        wb = load_workbook('users.xlsx')
        sheet = wb['datos']
        pass_antigua = 'vacio'

        # Crear lista de usuarios activos
        cells_usuario = sheet['A1':'A50']
        list_usuarios = []
        for row in cells_usuario:
            for cell in row:
                list_usuarios += [cell.value]

        # Crear lista de contraseñas activas
        cells_pass = sheet['B1':'B50']
        list_password = []
        for row in cells_pass:
            for cell in row:
                list_password += [cell.value]

        # busqueda comparativa entre el usuario introducido con su contraseña
        for usuario in list_usuarios:
            if usuario == self.nombre_usuario:
                posicion = list_usuarios.index(self.nombre_usuario)
                pass_antigua = list_password[posicion]

                if posicion >= 0:
                    posicion_bloqueo = posicion + 1
                    posicion_bloqueo = 'C' + str(posicion_bloqueo)
                    bloqueado = sheet[posicion_bloqueo]
               

        # Dialogo a traves del terminal con el usuario para saber las siguientes acciones con el cambio de contraseña
        if pass_antigua != 'vacio': # Controlar si el usuario existe o aun no se ha creado
                            
            if bloqueado.value == 'Bloqueado': # Si el usuario tiene bloqueo avisar y no hacer nada mas
                print('Su usuario esta bloqueado, contacte con el administrador del sistema.')
                
            else:
                pass_login = input('Introduzca su contraseña - ')

                if str(pass_login) == str(pass_antigua): # Comprobar que la contraseña sea la correcta
                    pass_cambiar = input('La contraseña introducida es valida, ¿Desea cambiarla por una nueva? (si o no) - ')

                    while pass_cambiar != 'si' or pass_cambiar != 'no': # Preguntar si se desea cambiar la contraseña
                        if pass_cambiar == 'si':
                            user.pass_nuevo(self.nombre_usuario, 'si', posicion)
                            break
                        elif pass_cambiar == 'no':
                            print('Su contraseña no sera modificada')
                            break
                        else:
                            pass_cambiar = input('Debe indicar si o no, vuelva a escribir su respuesta - ')

                else: # Hacer un contador de veces que ha introducido mal la contraseña
                    if self.conteo < 2:
                        self.conteo += 1
                        print('La contraseña no es valida, vuelva a intentarlo')
                        user.pass_buscar(self.conteo)
                    else:
                        print('La contraseña ha sido introducida de forma incorrecta tres veces, su usuario va a ser bloqueado')
                        posicion += 1
                        posicion = 'C' + str(posicion)
                        sheet[posicion] = 'Bloqueado'
                        wb.save('users.xlsx')

   
        else: # Si el usuario no existe se crea uno nuevo
            print('Su nombre de usuario no existe, se creara uno nuevo a continuacion')
            user.pass_nuevo(self.nombre_usuario, 'no', 0)
        
        return ''
    


    # Aqui se gestiona la creacion de una nueva contraseña
    def pass_nuevo(self, nombre_usuario, existente, posicion):
        self.nombre_usuario = nombre_usuario
        self.existente = existente
        self.posicion = posicion

        pass_nueva_metodo = input('¿Desea una contraseña manual o aleatoria? - ')

        if pass_nueva_metodo == "aleatoria": # Crear contraseña aleatoria
            pass_nueva = randint(100000, 999999)
            print(f'Gracias, su nueva contraseña es "{pass_nueva}"') 
            user.guardar_xlsx(self.nombre_usuario, pass_nueva, self.posicion, self.existente) 

        elif pass_nueva_metodo == "manual": # Crear contraseña a mano
            pass_nueva = input('Introduzca su nueva contraseña - ')

            if len(str(pass_nueva)) != 6: # Comprobar que la contraseña tenga 6 caracteres
                while len(str(pass_nueva)) != 6:
                    pass_nueva = input('La nueva contraseña debe tener seis caracteres. Vuelva a introducir una nueva contraseña - ')

                print(f'Gracias, su nueva contraseña es "{pass_nueva}"')
                user.guardar_xlsx(self.nombre_usuario, pass_nueva, self.posicion, self.existente) 
            else:
                print(f'Gracias, su nueva contraseña es "{pass_nueva}"')
                user.guardar_xlsx(self.nombre_usuario, pass_nueva, self.posicion, self.existente) 

        else:
            print('No ha escrito correctamente el metodo de nueva contraseña')
            user.pass_nuevo(self.nombre_usuario, self.existente, self.posicion)

        return ''
    

    
    # Aqui se guardan los datos de usuario y contraseña en el Excel
    def guardar_xlsx(self, nombre_usuario, pass_nueva, posicion, existente):
        self.nombre_usuario = nombre_usuario
        self.pass_nueva = pass_nueva
        self.posicion = posicion
        self.existente = existente

        wb = load_workbook('users.xlsx')
        sheet = wb['datos']
 
        # Usuario recurrente, solo se cambia la contraseña
        if self.existente == 'no':
            guardar_nuevo = [(self.nombre_usuario, self.pass_nueva)]

            for user_nuevo in guardar_nuevo:
                sheet.append(user_nuevo)
                wb.save('users.xlsx')
                print(f'Su usuario {self.nombre_usuario} ya ha sido añadido correctamente.')

        # Usuario nuevo, se guarda nombre y contraseña
        else:
            posicion += 1
            posicion = 'B' + str(posicion)
            sheet[posicion] = self.pass_nueva
            wb.save('users.xlsx')
            print(f'Su contraseña se ha modificado correctamente.')

        return ''
    


nombre_in = input('¿Cual es su nombre? - ')

user = User(nombre_in)
print(user.pass_buscar(0))