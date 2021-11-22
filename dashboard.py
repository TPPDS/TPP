#======================================================================================
#======================================================================================
#Librerias utilizadas
#Pandas - Limpieza, filtrado de DataFrames
#Streamlit - Interfaz gráfica
import pandas as pd
import streamlit as st
#--------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------
#Plotly - Gráficas de insights
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
#--------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------
#Evaluar las listas dentro del DataFrame
from ast import literal_eval
#--------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------
#Crear reportes de excel
import xlsxwriter
import io
from io import BytesIO
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
name_columns = ["Nombres", "Apellidos", "Género", "Fecha Nacimiento","Empresa/Hub", "Email", "Puesto", "Lugar Diversificado", "Nombre Diversificado", "Estado Diversificado", "Lugar Licenciatura", "Nombre Licenciatura", "Estado Licenciatura", "Lugar Maestría/Posgrado", "Nombre Maestría/Posgrado", "Estado Maestría/Posgrado", "Lugar Cursos/Diplomados/Certificaciones", "Nombre Cursos/Diplomados/Certificaciones", "Estado Cursos/Diplomados/Certificaciones", "Completo"]
#--------------------------------------------------------------------------------------
#Configuración de la página para que esta sea ancha
st.set_page_config(layout="wide", page_title = "Hello")
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
#======================================================================================
#======================================================================================
#Funciones
def column_to_list(name_c, df):
    df_eh = pd.DataFrame(columns = [name_c])
    df_eh[name_c] = df[name_c].astype(str)
    df_eh[name_c] = df_eh[name_c].apply(literal_eval)
    get_index = []
    for i, l in enumerate(df_eh[name_c]):
        if l[0] == filter_eh:
            get_index.append(i)
    return get_index
#--------------------------------------------------------------------------------------
def get_count_courses(df_filtro):
    df_work = df_filtro
    df_work["Nombre Completo"] = df_work["Nombres"] + " " + df_work["Apellidos"]
    name_c = "Nombre Cursos/Diplomados/Certificaciones"
    df_courses = pd.DataFrame(columns = [name_c])
    df_courses[name_c] = df_work[name_c].astype(str)
    df_courses[name_c] = df_courses[name_c].apply(literal_eval)
    get_count = []
    for i, l in enumerate(df_courses[name_c]):
        if l:
            get_count.append(len(l))
        else:
            get_count.append(0)
    df_work["Cantidad de Cursos/Diplomados/Certificaciones"] = get_count
    return df_work
#--------------------------------------------------------------------------------------
def get_count_state(df_filtro, name_c):
    df_work = df_filtro
    states = ["Terminado", "Cierre de Pensum", "En Curso"]
    cant_terminado = 0
    cant_cierre = 0
    cant_curso = 0
    get_count = []
    for index, value_s in enumerate(states):
        df_states = pd.DataFrame(columns = [value_s])
        df_states[name_c] = df_work[name_c].astype(str)
        df_states[name_c] = df_states[name_c].apply(literal_eval)
        cant_terminado = 0
        for i, l in enumerate(df_states[name_c]):
            if l:
                if l[0] == value_s:
                    cant_terminado += 1
        get_count.append(cant_terminado)
    return get_count

#--------------------------------------------------------------------------------------
def column_to_list_general(name_c, filter_eh, df):
    df_eh = pd.DataFrame(columns = [name_c])
    df_eh[name_c] = df[name_c].astype(str)
    df_eh[name_c] = df_eh[name_c].apply(literal_eval)
    get_index = []
    for i, l in enumerate(df_eh[name_c]):
        if l[0] == filter_eh:
            get_index.append(i)
    return get_index, len(get_index)
#--------------------------------------------------------------------------------------
def search_words(df, words, tipo):
    df_columns = df.columns
    df_columns = df_columns.drop(["Fecha de Nacimiento"])
    df_search = pd.DataFrame(columns = name_columns)
    if tipo == "Exacto":
        words = list(words.split(" "))
        for columns in df_columns:
            for i, w in enumerate(words):
                df_new = df[df[columns].str.contains(w, case = False)]
                if df_new.empty != True:
                    df_search = df_search.append(df_new,ignore_index = True)
        df_search = df_search.drop_duplicates()
    elif tipo == "Todas las coincidencias":
        words = words.replace(" ", "|")
        for columns in df_columns:
            df_new = df[df[columns].str.contains(words, case = False)]
            if df_new.empty != True:
                df_search = df_search.append(df_new,ignore_index = True)
        df_search = df_search.drop_duplicates()
    return df_search
#--------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------
def chart_1(df_general, df_filtro, label_name):
    values = [len(df_general) - len(df_filtro), len(df_filtro)]
    labels = ["Total",label_name]
    colors = ['gold', 'mediumturquoise', 'darkorange', 'lightgreen']
    fig = go.Figure(data = go.Pie(values=values,labels=labels,hole=0.6,marker_colors=colors))
    fig.update_traces(hoverinfo='label+value+percent',textinfo='percent',marker=dict(colors=colors, line=dict(color='#000000', width=1)))
    fig.update_layout(title_text="Cantidad de Colaboradores",title_font=dict(family='Verdana',color='darkred'),showlegend=True)
    fig.update_layout(margin=dict(b=0, l=0, r=0))
    fig.add_annotation(text=label_name+"<br>"+str(len(df_filtro)),font=dict(family='Verdana', color='black'), showarrow=False)
    fig.update_layout(legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1))
    return fig
#--------------------------------------------------------------------------------------
def chart_2(df_filtro, pos_legend):
    cant_F = len(df_filtro[df_filtro["Género"]=="F"])
    cant_M = len(df_filtro[df_filtro["Género"]=="M"])
    values = [cant_F, cant_M]
    labels = ["F","M"]
    colors = ['gold', 'mediumturquoise', 'darkorange', 'lightgreen']
    fig = go.Figure(data = go.Pie(values=values,labels=labels,hole=0.6,marker_colors=colors))
    fig.update_traces(hoverinfo='label+value+percent',textinfo='percent',marker=dict(colors=colors, line=dict(color='#000000', width=1)))
    fig.update_layout(title_text="Género",title_font=dict(family='Verdana',color='darkred'), showlegend=True)
    fig.update_layout(margin=dict(b=0, l=0, r=0))
    fig.add_annotation(text="F - "+str(cant_F)+"<br>"+"M - "+str(cant_M),font=dict(family='Verdana', color='black'), showarrow=False)
    if pos_legend == "arriba":
        fig.update_layout(legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1))
    return fig
#--------------------------------------------------------------------------------------
def chart_3(df_filtro):
    df_filtro = df_filtro.sort_values(by=["Cantidad de Cursos/Diplomados/Certificaciones"], ascending=False)
    fig = px.bar(df_filtro, x = "Nombre Completo", y = "Cantidad de Cursos/Diplomados/Certificaciones", color="Cantidad de Cursos/Diplomados/Certificaciones")
    fig.update_layout(title_text="Cantidad de<br>Cursos/Diplomados/Certificaciones",title_font=dict(family='Verdana',color='darkred'), showlegend=True)
    fig.update_layout(xaxis_tickangle=-45, yaxis_title="Cantidad")
    fig.update_traces(marker_line_color='rgb(8,48,107)', marker_line_width=0.5)
    fig.update_layout(coloraxis_showscale=False)
    fig.update_yaxes(rangemode="tozero")
    return fig

def chart_4(df_general, df_filtro, label_name):
    values = [len(df_general) - len(df_filtro), len(df_filtro)]
    labels = ["Total",label_name]
    colors = ['gold', 'mediumturquoise', 'darkorange', 'lightgreen']
    fig = go.Figure(data = go.Pie(values=values,labels=labels,hole=0.6,marker_colors=colors))
    fig.update_traces(hoverinfo='label+value+percent',textinfo='percent',marker=dict(colors=colors, line=dict(color='#000000', width=1)))
    fig.update_layout(title_text="Cantidad de Colaboradores",title_font=dict(family='Verdana',color='darkred'),showlegend=True)
    fig.update_layout(margin=dict(b=0, l=0, r=0))
    fig.add_annotation(text=label_name+"<br>"+str(len(df_filtro)),font=dict(family='Verdana', color='black'), showarrow=False)
    fig.update_layout(legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1))
    return fig

def chart_5(df_filtro):
    cant_F = len(df_filtro[df_filtro["Género"]=="F"])
    cant_M = len(df_filtro[df_filtro["Género"]=="M"])
    values = [cant_F, cant_M]
    labels = ["F","M"]
    colors = ['gold', 'mediumturquoise', 'darkorange', 'lightgreen']
    fig = go.Figure(data = go.Pie(values=values,labels=labels,hole=0.6,marker_colors=colors))
    fig.update_traces(hoverinfo='label+value+percent',textinfo='percent',marker=dict(colors=colors, line=dict(color='#000000', width=1)))
    fig.update_layout(title_text="Género",title_font=dict(family='Verdana',color='darkred'), showlegend=True)
    fig.update_layout(margin=dict(b=0, l=0, r=0))
    fig.add_annotation(text="F - "+str(cant_F)+"<br>"+"M - "+str(cant_M),font=dict(family='Verdana', color='black'), showarrow=False)
    fig.update_layout(legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1))
    return fig

def chart_6(values, labels, df, label_name):
    colors = ['gold', 'mediumturquoise', 'darkorange', 'lightgreen']
    fig = go.Figure(data = go.Pie(values=values,labels=labels,hole=0.6,marker_colors=colors))
    fig.update_traces(hoverinfo='label+value+percent',textinfo='percent',marker=dict(colors=colors, line=dict(color='#000000', width=1)))
    fig.update_layout(title_text="Cantidad de Colaboradores",title_font=dict(family='Verdana',color='darkred'),showlegend=True)
    fig.update_layout(margin=dict(b=0, l=0, r=0))
    fig.add_annotation(text=label_name+"<br>"+str(len(df)),font=dict(family='Verdana', color='black'), showarrow=False)
    return fig

def chart_7(l_states, text_c, pos_legend):
    values = [l_states[0], l_states[1], l_states[2]]
    labels = ["Terminado","Cierre Pensum", "En Curso"]
    colors = ['gold', 'mediumturquoise', 'darkorange', 'lightgreen']
    fig = go.Figure(data = go.Pie(values=values,labels=labels,hole=0.6,marker_colors=colors))
    fig.update_traces(hoverinfo='label+value+percent',textinfo='percent',marker=dict(colors=colors, line=dict(color='#000000', width=1)))
    fig.update_layout(title_text=text_c,title_font=dict(family='Verdana',color='darkred'), showlegend=True)
    fig.update_layout(margin=dict(b=0, l=0, r=0))
    fig.add_annotation(text=labels[0]+" - "+str(l_states[0])+"<br>"+labels[1]+" - "+str(l_states[1])+"<br>"+labels[2]+" - "+str(l_states[2]),font=dict(family='Verdana', color='black'), showarrow=False)
    fig.update_traces(textposition='inside')
    if pos_legend == "arriba":
        fig.update_layout(legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1), uniformtext_mode='hide')
    return fig
#======================================================================================
#======================================================================================
#Subir documento y f_1 indica que el documento se subió exitosamente
df_upload = st.file_uploader("Subir documento", [".csv", ".xlsx"])
f_1 = False
if df_upload:
    try:
        df_all = pd.read_csv(df_upload)
        df_all["Fecha de Nacimiento"] = pd.to_datetime(df_all["Fecha de Nacimiento"]).dt.date
        f_1 = True
    except:
        df_all = pd.read_excel(df_upload)
        df_all["Fecha de Nacimiento"] = pd.to_datetime(df_all["Fecha de Nacimiento"]).dt.date
        f_1 = True
if f_1:
    if st.checkbox("Explorar documento"):
        st.dataframe(df_all)
    #======================================================================================
    #======================================================================================
    #Sidebar - Filtrado de datos
    filter_type = st.sidebar.radio("Tipo de Filtro", ["Ninguno","Empresa/Hub","Coincidencia de palabras"])
    if filter_type == "Ninguno":
        st.session_state.df_filtro = df_all
    elif filter_type == "Empresa/Hub":
        st.sidebar.markdown("***")
        filter_eh = st.sidebar.selectbox("Seleccionar Empresa/Hub", e_hub)
        get_index = column_to_list(filter_type, df_all)
        st.session_state.df_filtro = df_all.iloc[get_index]
    elif filter_type == "Coincidencia de palabras":
        st.sidebar.markdown("***")
        palabras = st.sidebar.text_input("Escribir palabra")
        filter_eh = palabras
        st.session_state.df_filtro = search_words(df_all, palabras, "Todas las coincidencias")
#======================================================================================
#======================================================================================
#======================================================================================
#======================================================================================
#General dashboard
st.markdown("***")
st.subheader("Insights")
general_insights = st.expander("Insights", expanded = False)
with general_insights:
    if f_1:
        c1, c2 = st.columns((1, 1))
        with c1:
            if filter_type == "Ninguno":
                cantidad_list = []
                for index, value in enumerate(e_hub):
                    _, cantidad = column_to_list_general("Empresa/Hub", value, df_all)
                    cantidad_list.append(cantidad)
                fig_general = chart_6(cantidad_list, e_hub, df_all, "Total")
                st.plotly_chart(fig_general, use_container_width=True)
            else:
                fig_proporcion = chart_1(df_all, st.session_state.df_filtro, filter_eh)
                st.plotly_chart(fig_proporcion, use_container_width = True)
        with c2:
            if filter_type == "Ninguno":
                fig_genero = chart_2(st.session_state.df_filtro, "")
                st.plotly_chart(fig_genero, use_container_width=True)
            else:
                fig_genero = chart_2(st.session_state.df_filtro, "arriba")
                st.plotly_chart(fig_genero,use_container_width=True)
        c3, c4, c5= st.columns((1,1,1))
        with c3:
            list_count = get_count_state(st.session_state.df_filtro, "Estado Diversificado")
            fig_div = chart_7(list_count, "Diversificado", "arriba")
            st.plotly_chart(fig_div,use_container_width=True)
        with c4:
            list_count = get_count_state(st.session_state.df_filtro, "Estado Licenciatura")
            fig_lic = chart_7(list_count, "Licenciatura", "arriba")
            st.plotly_chart(fig_lic,use_container_width=True)
        with c5:
            list_count = get_count_state(st.session_state.df_filtro, "Estado Maestría/Posgrado")
            fig_maestria = chart_7(list_count, "Maestría/Posgrado", "arriba")
            st.plotly_chart(fig_maestria,use_container_width=True)
        if filter_type == "Ninguno":
            st.plotly_chart(chart_3(get_count_courses(df_all)), use_container_width=True)
        if filter_type == "Empresa/Hub":
            st.plotly_chart(chart_3(get_count_courses(st.session_state.df_filtro)), use_container_width=True)
        if filter_type == "Coincidencia de palabras":
            st.plotly_chart(chart_3(get_count_courses(st.session_state.df_filtro)), use_container_width=True)
    else:
        st.markdown("#### Cargar archivo")
#======================================================================================
#======================================================================================
#Funciones para descargar reporte
def list_to_string(l):
    new_string = ",\n".join(l)
    return new_string

def column_to_list_report(name_c, df):
    df_eh = pd.DataFrame(columns = [name_c])
    df_eh[name_c] = df[name_c].astype(str)
    df_eh[name_c] = df_eh[name_c].apply(literal_eval)
    df_eh[name_c] = df_eh[name_c].apply(list_to_string)
    return df_eh[name_c]

def ordenar_columnas(df):
    change_columns = ["Empresa/Hub","Puesto","Lugar Diversificado","Nombre Diversificado","Estado Diversificado","Lugar Licenciatura","Nombre Licenciatura","Estado Licenciatura","Lugar Maestría/Posgrado","Nombre Maestría/Posgrado","Estado Maestría/Posgrado","Lugar Cursos/Diplomados/Certificaciones","Nombre Cursos/Diplomados/Certificaciones","Estado Cursos/Diplomados/Certificaciones"]
    for value in change_columns:
        df[value] = column_to_list_report(value, df)
    df = df.reindex(["Nombre Completo","Nombres","Apellidos","Género", "Fecha de Nacimiento","Empresa/Hub","Email","Puesto","Lugar Diversificado","Nombre Diversificado","Estado Diversificado","Lugar Licenciatura","Nombre Licenciatura","Estado Licenciatura","Lugar Maestría/Posgrado","Nombre Maestría/Posgrado","Estado Maestría/Posgrado","Lugar Cursos/Diplomados/Certificaciones","Nombre Cursos/Diplomados/Certificaciones","Estado Cursos/Diplomados/Certificaciones","Cantidad de Cursos/Diplomados/Certificaciones","Completo"], axis=1)
    df = df.sort_values(by="Cantidad de Cursos/Diplomados/Certificaciones", ascending=False)
    return df

def get_count_state_report(df_filtro, name_c):
    df_work = df_filtro
    states = ["Terminado", "Cierre de Pensum", "En Curso"]
    get_count = []
    for index, value_s in enumerate(states):
        df_states = pd.DataFrame(columns = [value_s])
        df_states[name_c] = df_work[name_c].astype(str)
        cant_terminado = 0
        for i, l in enumerate(df_states[name_c]):
            if l == value_s:
                cant_terminado += 1
        get_count.append(cant_terminado)
    return get_count

def get_count_courses_report(df_filtro):
    df_work = df_filtro
    df_work["Nombre Completo"] = df_work["Nombres"] + " " + df_work["Apellidos"]
    name_c = "Nombre Cursos/Diplomados/Certificaciones"
    df_courses = pd.DataFrame(columns = [name_c])
    df_courses[name_c] = df_work[name_c].astype(str)
    df_courses[name_c] = df_courses[name_c].apply(literal_eval)
    get_count = []
    for i, l in enumerate(df_courses[name_c]):
        if l:
            get_count.append(len(l))
        else:
            get_count.append(0)
    df_work["Cantidad de Cursos/Diplomados/Certificaciones"] = get_count
    return df_work
#======================================================================================
#======================================================================================
st.markdown("***")
st.subheader("Descargar reporte")
descargar_reporte = st.expander("Descargar", expanded = False)
with descargar_reporte:
    if f_1:
        r1, r2 = st.columns((1, 1))
        with r1:
            report_type = st.radio("Tipo de reporte", ["Todo", "Por Empresa/Hub"])
        with r2:
            buffer = io.BytesIO()
            df_ordenado = ordenar_columnas(df_all)
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                if report_type == "Todo":
                    df_ordenado.to_excel(writer, sheet_name='Todo', index = False)
                    workbook  = writer.book
                    worksheet = writer.sheets['Todo']
                    worksheet.set_zoom(50)
                    worksheet.conditional_format(1,20,len(df_ordenado)+1,20, {'type': 'data_bar', 'data_bar_2010': True, 'bar_color': '#ffa200'})
                    (max_row, max_col) = df_ordenado.shape
                    column_settings = [{'header': column} for column in df_ordenado.columns]
                    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings, 'style': 'Table Style Light 11'})
                    worksheet.set_column(0, max_col - 1, 12)
                    worksheet = workbook.add_worksheet('Insights')
                    fig_g = chart_6(cantidad_list, e_hub, df_all, "Total")
                    fig_ge = chart_2(df_all, "")
                    #image_data = BytesIO(fig_general.to_image(format="png", engine = 'kaleido'))
                    image_data = BytesIO(fig_g.to_image(format="png", engine = 'kaleido'))
                    worksheet.insert_image(0, 0, 'plotly.png', {'image_data': image_data})
                    image_data = BytesIO(fig_ge.to_image(format="png", engine = 'kaleido'))
                    worksheet.insert_image(0, 11, 'plotly.png', {'image_data': image_data})
                    count = 0
                    for index, x in enumerate(["Diversificado", "Licenciatura", "Maestría/Posgrado"]):
                        #st.write("Estado "+x)
                        list_count = get_count_state_report(df_all, "Estado "+x)
                        #st.write(list_count)
                        fig_coso = chart_7(list_count, x, "arriba")
                        image_data = BytesIO(fig_coso.to_image(format="png", engine = 'kaleido'))
                        worksheet.insert_image(27, count, 'plotly.png', {'image_data': image_data})
                        count += 11
                    fig_coso = chart_3(df_all)
                    image_data = BytesIO(fig_coso.to_image(format="png", engine = 'kaleido', width=2000))
                    worksheet.insert_image(54, 0, 'plotly.png', {'image_data': image_data})
                    worksheet.set_zoom(25)
                    writer.save()
                if report_type == "Por Empresa/Hub":
                    for hojas in e_hub:
                        df_filtrado = df_ordenado[df_ordenado["Empresa/Hub"]==hojas]
                        df_filtrado.to_excel(writer, sheet_name=str(hojas), index = False)
                        workbook  = writer.book
                        worksheet = writer.sheets[hojas]
                        worksheet.set_zoom(50)
                        worksheet.conditional_format(1,20,len(df_filtrado)+1,20, {'type': 'data_bar', 'data_bar_2010': True, 'bar_color': '#ffa200'})
                        (max_row, max_col) = df_filtrado.shape
                        column_settings = [{'header': column} for column in df_filtrado.columns]
                        worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings, 'style': 'Table Style Light 11'})
                        worksheet.set_column(0, max_col - 1, 12)
                        worksheet = workbook.add_worksheet('Insights '+hojas)
                        fig_g = chart_1(df_all, df_filtrado, hojas)
                        fig_ge = chart_2(df_filtrado, "")
                        #image_data = BytesIO(fig_general.to_image(format="png", engine = 'kaleido'))
                        image_data = BytesIO(fig_g.to_image(format="png", engine = 'kaleido'))
                        worksheet.insert_image(0, 0, 'plotly.png', {'image_data': image_data})
                        image_data = BytesIO(fig_ge.to_image(format="png", engine = 'kaleido'))
                        worksheet.insert_image(0, 11, 'plotly.png', {'image_data': image_data})
                        count = 0
                        for index, x in enumerate(["Diversificado", "Licenciatura", "Maestría/Posgrado"]):
                            list_count = get_count_state_report(df_filtrado, "Estado "+x)
                            fig_coso = chart_7(list_count, x, "arriba")
                            image_data = BytesIO(fig_coso.to_image(format="png", engine = 'kaleido'))
                            worksheet.insert_image(27, count, 'plotly.png', {'image_data': image_data})
                            count += 11
                        fig_coso = chart_3(df_filtrado)
                        image_data = BytesIO(fig_coso.to_image(format="png", engine = 'kaleido'))
                        worksheet.insert_image(54, 0, 'plotly.png', {'image_data': image_data})
                        worksheet.set_zoom(25)
                    writer.save()
                st.download_button("Descargar", buffer, "Reporte.xlsx")
        #st.write("<style>div.row-widget.stRadio > div{flex-direction:row;}</style>", unsafe_allow_html = True)
    else:
        st.markdown("#### Cargar archivo")
