#======================================================================================
#======================================================================================
#Librerias utilizadas
#Pandas - Limpieza, filtrado de DataFrames
#Streamlit - Interfaz gráfica
import pandas as pd
import streamlit as st

from google.oauth2.service_account import Credentials
from gspread_pandas import Spread, Client
#--------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------
#Evaluar las listas dentro del DataFrame
from ast import literal_eval
#--------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------
#Listas para utilizas en el documento
#Empresa/Hub
e_hub = sorted(["EasyGo","Lets Advertise","Administración","RRHH","TPP","TPP Extreme","TPP Fénix","TPP ULTRA","TOM"])
#--------------------------------------------------------------------------------------
#Puestos registrados
nombre_puesto = sorted(['Gerente Contable y Auditor Regional','Asistente Contable','Intendencia y Mensajería','Jefe Local de Contabilidad','Programador/a','Encargado del Dpto. de IT ','Asistente de IT', 'Project Manager','Front-End Developer','Asesor Comercial','Gerente de Desarrollo, Easy Go','Diseñador/a Gráfico/a','Asistente Lets Advertise','SEO','Gerente de Medios','Científico de Datos','Gerente de Recursos Humanos','Asistente de Recursos Humanos','Asistente TOM','Social Media Manager','Ejecutiva de Cuentas','Asistente de Cuentas','Gerente de División TOM','Trafficker','Gerente de Ventas','Gerente de División TPP','Gerente de Operaciones','Gerente Financiero Administrativo','Gerente Comercial','Gerente de Nuevos Negocios','Copywritter'])
#--------------------------------------------------------------------------------------
#Nombres de las columnas nuevo documento
name_columns = ["Nombres", "Apellidos", "Género", "Fecha de Nacimiento","Empresa/Hub", "Email", "Puesto", "Lugar Diversificado", "Nombre Diversificado", "Estado Diversificado", "Lugar Licenciatura", "Nombre Licenciatura", "Estado Licenciatura", "Semestre", "Lugar Maestría/Posgrado", "Nombre Maestría/Posgrado", "Estado Maestría/Posgrado", "Lugar Cursos/Diplomados/Certificaciones", "Nombre Cursos/Diplomados/Certificaciones", "Estado Cursos/Diplomados/Certificaciones", "Completo"]
#--------------------------------------------------------------------------------------
#Nombres de las columnas nuevo documento
#name_columns = ["Nombres", "Apellidos", "Género", "Fecha de Nacimiento","Empresa/Hub", "Email", "Puesto", "Lugar Diversificado", "Nombre Diversificado", "Estado Diversificado", "Lugar Licenciatura", "Nombre Licenciatura", "Estado Licenciatura", "Lugar Maestría/Posgrado", "Nombre Maestría/Posgrado", "Estado Maestría/Posgrado", "Lugar Cursos/Diplomados/Certificaciones", "Nombre Cursos/Diplomados/Certificaciones", "Estado Cursos/Diplomados/Certificaciones", "Completo"]
#--------------------------------------------------------------------------------------
#Configuración de la página para que esta sea ancha
st.set_page_config(layout="wide")
#--------------------------------------------------------------------------------------
#Configuración para ocultar menu de hamburguesa y pie de página
hide_st_style = """
                <style>
                #MainMenu {visibility: hidden;}
                footer {visibility: hidden;}
                </style>
                """
st.markdown(hide_st_style, unsafe_allow_html = True)
#--------------------------------------------------------------------------------------
#Session state
if 'df_filtro' not in st.session_state:
    st.session_state.df_filtro = pd.DataFrame(columns = name_columns)
