#ARCHIVO: funcgeneral.py
import os
import datetime
from contextlib import nullcontext

#GENERAL.GESTION DE TRAZAS
class Trazaslg:
    def __init__(self, nombre_archivo):
        if nombre_archivo is not nullcontext:
            self.archivo_log = nombre_archivo
        else:
            self.archivo_log = "trazas.log"

    # Funcion iniciar_traza
    def iniciar_traza(self, mensaje):
        fecha = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(self.archivo_log, "a") as f:
            f.write(f"[{fecha}] Inicio: {mensaje}\n")

    #Funcion finalizar_traza
    def finalizar_traza(self, mensaje):
        fecha = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(self.archivo_log, "a") as f:
            f.write(f"[{fecha}] Fin: {mensaje}\n")

    #Funcion registrar_msg_traza
    def registrar_msg_traza(self, mensaje):
        fecha = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(self.archivo_log, "a") as f:
            f.write(f"[{fecha}] Mensaje: {mensaje}\n")

    #Funcion registrar_err_traza
    def registrar_err_traza(self, mensaje):
        fecha = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(self.archivo_log, "a") as f:
            f.write(f"[{fecha}] Error: {mensaje}\n")

    #Funcion cerrar_traza        
    def cerrar_traza(self):
        self.archivo_log.close()