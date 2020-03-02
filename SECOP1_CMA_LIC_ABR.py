# -*- coding: utf-8 -*-
"""
Created on Thu Jun 21 21:53:12 2018

@author: INALCON SAS
"""


from bs4 import BeautifulSoup
import urllib3
import re
import pandas
import numpy
from datetime import datetime
from datetime import timedelta
import xlrd
import math
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import xlsxwriter

#crear listas para guardar datos
lista = []
test_list = []
totalresultados = []
find_estado = ['Borrado','Convocad']
find_estado_2 = '|'.join(find_estado)
#PATH = 'D:/Servidor INDGER/4. Buscador de Contratos/6. SECOP I/CMA_LIC_ABR/' # Recuerde cambiar la ruta de la carpeta, cambiar \ por / #
PATH = 'C:/Users/90044338/Documents/Administrativo/SECOP/MCA_ABR_LIC/'
prev_link = 'https://www.contratos.gov.co'

#Solicita las fechas para la busqueda en SECOP 1
fi = (datetime.now() + timedelta(days=-30)).strftime('%d/%m/%Y')
ff = (datetime.now() + timedelta(days=15)).strftime('%d/%m/%Y')
ar = datetime.now().strftime('%d_%m_%Y__%H:%M')
x = 1

# Conexión a URL del SECOP para obtener primero el npumero de paginas a decargar
url = 'https://www.contratos.gov.co/consultas/resultadosConsulta.do?&departamento=&entidad=&paginaObjetivo=%(contador)s&fechaInicial=%(fechainicial)s&ctl00$ContentPlaceHolder1$hidIdOrgV=-1&desdeFomulario=true&registrosXPagina=50&estado=0&ctl00$ContentPlaceHolder1$hidIdEmpresaVenta=-1&ctl00$ContentPlaceHolder1$hidNombreProveedor=-1&ctl00$ContentPlaceHolder1$hidRedir=&ctl00$ContentPlaceHolder1$hidIDProducto=-1&ctl00$ContentPlaceHolder1$hidNombreDemandante=-1&cuantia=0&ctl00$ContentPlaceHolder1$hidNombreProducto=-1&ctl00$ContentPlaceHolder1$hidIdEmpresaC=0&ctl00$ContentPlaceHolder1$hidIDProductoNoIngresado=-1&ctl00$ContentPlaceHolder1$hidRangoMaximoFecha=&fechaFinal=%(fechafinal)s&ctl00$ContentPlaceHolder1$hidIdOrgC=-1&objeto=&tipoProceso=&ctl00$ContentPlaceHolder1$hidIDRubro=-1&municipio=0&numeroProceso=' % dict(contador = x, fechainicial = fi, fechafinal = ff)

http = urllib3.PoolManager()
ini = http.request('GET', url)
bs = BeautifulSoup(ini.data)

# obtener el numero de paginas de resultados
resultados = bs.findAll('p', {'class':'resumenResultados'})
for pag in resultados:
	totalresultados.append(pag.text)
	textresult = totalresultados[0]
	paginas = re.findall('([0-9]+)', textresult)
	totalregistros = int(paginas[0])
	numpaginas = totalregistros/50

numpaginas = math.ceil(numpaginas)
print (numpaginas)

# Obtiene los datos de cada una de las páginas de resultados del SECOP
while x <= numpaginas:
    
    url = 'https://www.contratos.gov.co/consultas/resultadosConsulta.do?&departamento=&entidad=&paginaObjetivo=%(contador)s&fechaInicial=%(fechainicial)s&ctl00$ContentPlaceHolder1$hidIdOrgV=-1&desdeFomulario=true&registrosXPagina=50&estado=0&ctl00$ContentPlaceHolder1$hidIdEmpresaVenta=-1&ctl00$ContentPlaceHolder1$hidNombreProveedor=-1&ctl00$ContentPlaceHolder1$hidRedir=&ctl00$ContentPlaceHolder1$hidIDProducto=-1&ctl00$ContentPlaceHolder1$hidNombreDemandante=-1&cuantia=0&ctl00$ContentPlaceHolder1$hidNombreProducto=-1&ctl00$ContentPlaceHolder1$hidIdEmpresaC=0&ctl00$ContentPlaceHolder1$hidIDProductoNoIngresado=-1&ctl00$ContentPlaceHolder1$hidRangoMaximoFecha=&fechaFinal=%(fechafinal)s&ctl00$ContentPlaceHolder1$hidIdOrgC=-1&objeto=&tipoProceso=&ctl00$ContentPlaceHolder1$hidIDRubro=-1&municipio=0&numeroProceso=' % dict(contador = x, fechainicial = fi, fechafinal = ff)
    ini = http.request('GET', url)
    bs = BeautifulSoup(ini.data)

	# para sacar valores del html
    impares = bs.findAll('td', {'class':'tablaslistOdd'})
    pares = bs.findAll('td', {'class':'tablaslistEven'})
    
	# ingresas textos obtenidos del html a la lista 
    for par in pares:
        test_list.append(par)
    for impar in impares:
        test_list.append(impar)
        
    print (x)
    x = x + 1

print (len(test_list))

#Crear Dataframe con los resultados de la busqueda en el SECOP
cabeceras = ['ID','Num_Proceso','Tipo_Proceso','Estado','Entidad','Objeto','Ubicacion','Cuantia','Fecha']
base_all = pandas.DataFrame(numpy.array(test_list).reshape(-1,len(cabeceras)), columns = cabeceras)

#Preproceso Dataframe
#Convertir variables a tipo de dato String
base_all = base_all.astype(str)

#Eliminar nelines en columna de objeto
base_all['Objeto'] = base_all['Objeto'].map(lambda x: x.replace('\n','').strip())

#Extraer valores de cada una de las variables 
base_all['ID'] = base_all['ID'].str.extract('(\>.*\<)')
base_all['Link'] = base_all['Num_Proceso'].str.extract('(\'.*\')')
base_all['Num_Proceso'] = base_all['Num_Proceso'].str.extract('(\>.*\<)')
base_all['Tipo_Proceso'] = base_all['Tipo_Proceso'].str.extract('(\>.*\<)')
base_all['Estado'] = base_all['Estado'].str.extract('(\>.*\<)')
base_all['Entidad'] = base_all['Entidad'].str.extract('(\>.*\<)')
base_all['Objeto'] = base_all['Objeto'].str.extract('(\>.*\<)')
base_all['Departamento'] = base_all['Ubicacion'].str.extract('(\>.*\</b)')
base_all['Municipio'] = base_all['Ubicacion'].str.extract('(\:.*\</td)')
base_all['Cuantia'] = base_all['Cuantia'].str.extract('(\$.*\<)')
base_all['Estado_actual'] = base_all['Fecha'].str.extract('(\>[A-Z].*\</b)')
base_all['Fecha'] = base_all['Fecha'].str.extract('(\>[0-9].*\<)')

#Eliminar caracteres innecesarios
base_all['ID'] = base_all['ID'].str.replace('>','')
base_all['ID'] = base_all['ID'].str.replace('<','')

base_all['Num_Proceso'] = base_all['Num_Proceso'].str.replace('>','')
base_all['Num_Proceso'] = base_all['Num_Proceso'].str.replace('<','')

base_all['Tipo_Proceso'] = base_all['Tipo_Proceso'].str.replace('>','')
base_all['Tipo_Proceso'] = base_all['Tipo_Proceso'].str.replace('<','')

base_all['Estado'] = base_all['Estado'].str.replace('>','')
base_all['Estado'] = base_all['Estado'].str.replace('<','')

base_all['Entidad'] = base_all['Entidad'].str.replace('>','')
base_all['Entidad'] = base_all['Entidad'].str.replace('<','')

base_all['Objeto'] = base_all['Objeto'].str.replace('>','')
base_all['Objeto'] = base_all['Objeto'].str.replace('<','')

base_all['Cuantia'] = base_all['Cuantia'].str.replace('>','')
base_all['Cuantia'] = base_all['Cuantia'].str.replace('<','')
base_all['Cuantia'] = base_all['Cuantia'].str.rstrip('0')
base_all['Cuantia'] = base_all['Cuantia'].str.replace(',','')
base_all['Cuantia'] = base_all['Cuantia'].str.replace('.','')
base_all['Cuantia'] = base_all['Cuantia'].str.replace('$','')
base_all['Cuantia'] = base_all['Cuantia'].astype('int64')

base_all['Fecha'] = base_all['Fecha'].str.replace('>','')
base_all['Fecha'] = base_all['Fecha'].str.replace('<','')
base_all['dia'] = base_all['Fecha'].str.extract('(^[0-9]{2})')
base_all['mes'] = base_all['Fecha'].str.extract('(-.*-)')
base_all['mes'] = base_all['mes'].str.replace('-','')

# Para cambiar meses (optimizar)
base_all.loc[base_all.mes == 'ENE','mes'] = '01'
base_all.loc[base_all.mes == 'FEB','mes'] = '02'
base_all.loc[base_all.mes == 'MAR','mes'] = '03'
base_all.loc[base_all.mes == 'ABR','mes'] = '04'
base_all.loc[base_all.mes == 'MAY','mes'] = '05'
base_all.loc[base_all.mes == 'JUN','mes'] = '06'
base_all.loc[base_all.mes == 'JUL','mes'] = '07'
base_all.loc[base_all.mes == 'AGO','mes'] = '08'
base_all.loc[base_all.mes == 'SEP','mes'] = '09'
base_all.loc[base_all.mes == 'OCT','mes'] = '10'
base_all.loc[base_all.mes == 'NOV','mes'] = '11'
base_all.loc[base_all.mes == 'DIC','mes'] = '12'

base_all['ano'] = base_all['Fecha'].str.extract('([0-9]{2}$)')

base_all['Fecha'] = base_all['dia'] +'-'+ base_all['mes'] +'-'+ '20' + base_all['ano']
base_all['Fecha'] = pandas.to_datetime(base_all.Fecha, format = '%d-%m-%Y')

base_all['Link'] = base_all['Link'].str.lstrip("'")
base_all['Link'] = base_all['Link'].str.rstrip("'")
base_all['Link'] = prev_link + base_all['Link']

base_all['Departamento'] = base_all['Departamento'].str.replace("><b>",'')
base_all['Departamento'] = base_all['Departamento'].str.replace("</b",'')
base_all['Departamento'] = base_all['Departamento'].str.replace("<br/",' ')
base_all['Departamento'] = base_all['Departamento'].str.replace(">",'')
base_all['Departamento'] = base_all['Departamento'].str.replace(":",'')

base_all['Municipio'] = base_all['Municipio'].str.replace("</td",'')
base_all['Municipio'] = base_all['Municipio'].str.replace(":",'')
base_all['Municipio'] = base_all['Municipio'].map(lambda x: x.strip())

base_all['Estado_actual'] = base_all['Estado_actual'].str.replace('>','')
base_all['Estado_actual'] = base_all['Estado_actual'].str.replace('</b','')

# Ordenar Dataframe
base_all = base_all.drop(columns = ['Ubicacion'])
base_all = base_all[['ID','Num_Proceso','Tipo_Proceso','Estado','Entidad','Objeto','Departamento','Municipio','Cuantia','Estado_actual','Fecha','Link']]

# Obtener historico de resultados
resultados_total_CMA = pandas.read_excel(PATH + 'resultados_total_CMA.xls', header = 0)

resultados_total_licitacion = pandas.read_excel(PATH + 'resultados_total_Licitacion.xls', header = 0)

resultados_total_abreviadas = pandas.read_excel(PATH + 'resultados_total_Abreviada.xls', header = 0)

# Filtrar de acuerdo a criterios de Min cuantia
resultados_CMA = base_all[(base_all.Tipo_Proceso == "Concurso de Méritos Abierto") &
                          (base_all.Estado.str.contains(find_estado_2)) 
                          & (base_all.Cuantia < 1000000000)]


resultados_licitacion = base_all[(base_all.Tipo_Proceso == "Licitación Pública") &
                          (base_all.Estado.str.contains(find_estado_2)) 
                          & (base_all.Cuantia < 1200000000)]

resultados_abreviada = base_all[(base_all.Tipo_Proceso == "Selección Abreviada de Menor Cuantía (Ley 1150 de 2007)") &
                          (base_all.Estado.str.contains(find_estado_2))]

# Eliminar procesos terminados anormalmente

resultados_CMA = resultados_CMA[resultados_CMA.Estado != "Terminado Anormalmente después de Convocado"]

resultados_licitacion = resultados_licitacion[resultados_licitacion.Estado != "Terminado Anormalmente después de Convocado"]

resultados_abreviada = resultados_abreviada[resultados_abreviada.Estado != "Terminado Anormalmente después de Convocado"]

##########################################################
###########             MERITOS ABIERTOS
##########################################################
#Verificación de nuevos resultados concurso de meritos abierto

resultados_hoy_CMA = pandas.merge(resultados_CMA,resultados_total_CMA, 
                              how = 'left', 
                              on = ['Num_Proceso','Link','Entidad'])

resultados_hoy_CMA = resultados_hoy_CMA[resultados_hoy_CMA.ID_y.isnull()]
resultados_hoy_CMA = resultados_hoy_CMA.iloc[:,:12]

header_totals = resultados_total_CMA.columns
resultados_hoy_CMA.columns = header_totals
resultados_total_CMA = resultados_total_CMA.append(resultados_hoy_CMA)
resultados_total_CMA.to_excel(PATH + 'resultados_total2_CMA.xls', index = False)

# Escribir resultados del día
writer = pandas.ExcelWriter(PATH + 'resultados_hoy_CMA' + ar + '.xlsx',engine = 'xlsxwriter')
resultados_hoy_CMA.to_excel(writer, sheet_name = 'Resumen_resultados')
workbook  = writer.book
worksheet = writer.sheets['Resumen_resultados']
formatdate = workbook.add_format({'valign': 'vcenter'})
format1 = workbook.add_format({'text_wrap': True,
                               'valign': 'vcenter'})
format2 = workbook.add_format({'num_format': '$ #,##0.00',
                               'valign': 'vcenter'})
formatall = workbook.add_format({'valign': 'vcenter',
                                 'align': 'center'})                               

worksheet.set_column('A:O',None,formatall)
worksheet.set_column('B:B',5)
worksheet.set_column('C:C',15,format1)
worksheet.set_column('D:D',15,format1)
worksheet.set_column('E:E',9.45)
worksheet.set_column('F:F',25,format1)
worksheet.set_column('G:G',53,format1)
worksheet.set_column('H:H',18, format1)
worksheet.set_column('I:I',12.35,format1)
worksheet.set_column('E:E',9.45)
worksheet.set_column('J:J',16,format2)
worksheet.set_column('L:L',17.30, formatdate)
worksheet.set_column('M:M',85,format1)
worksheet.set_column('A:A',None,None,{'hidden': True})
worksheet.set_column('K:K',None,None,{'hidden': True})

writer.save()

# resultados_hoy_CMA.to_excel(PATH + 'resultados_hoy_CMA' + ar + '.xls', index = False)

##########################################################
###########             LICITACIONES
##########################################################
#Verificación de nuevos resultados licitaciones

resultados_hoy_licitacion = pandas.merge(resultados_licitacion,resultados_total_licitacion, 
                              how = 'left', 
                              on = ['Num_Proceso','Link','Entidad'])

resultados_hoy_licitacion = resultados_hoy_licitacion[resultados_hoy_licitacion.ID_y.isnull()]
resultados_hoy_licitacion = resultados_hoy_licitacion.iloc[:,:12]

header_totals = resultados_total_licitacion.columns
resultados_hoy_licitacion.columns = header_totals
resultados_total_licitacion = resultados_total_licitacion.append(resultados_hoy_licitacion)
resultados_total_licitacion.to_excel(PATH + 'resultados_total_Licitacion.xls', index = False)

# Escribir resultados del día
writer = pandas.ExcelWriter(PATH + 'resultados_hoy_licitacion' + ar + '.xlsx',engine = 'xlsxwriter')
resultados_hoy_licitacion.to_excel(writer, sheet_name = 'Resumen_resultados')
workbook  = writer.book
worksheet = writer.sheets['Resumen_resultados']
formatdate = workbook.add_format({'valign': 'vcenter'})
format1 = workbook.add_format({'text_wrap': True,
                               'valign': 'vcenter'})
format2 = workbook.add_format({'num_format': '$ #,##0.00',
                               'valign': 'vcenter'})
formatall = workbook.add_format({'valign': 'vcenter',
                                 'align': 'center'})                               

worksheet.set_column('A:O',None,formatall)
worksheet.set_column('B:B',5)
worksheet.set_column('C:C',15,format1)
worksheet.set_column('D:D',15,format1)
worksheet.set_column('E:E',9.45)
worksheet.set_column('F:F',25,format1)
worksheet.set_column('G:G',53,format1)
worksheet.set_column('H:H',18, format1)
worksheet.set_column('I:I',12.35,format1)
worksheet.set_column('E:E',9.45)
worksheet.set_column('J:J',16,format2)
worksheet.set_column('L:L',17.30, formatdate)
worksheet.set_column('M:M',85,format1)
worksheet.set_column('A:A',None,None,{'hidden': True})
worksheet.set_column('K:K',None,None,{'hidden': True})

writer.save()

#resultados_hoy_licitacion.to_excel(PATH + 'resultados_hoy_licitacion' + ar + '.xls', index = False)

##########################################################
###########             ABREVIADAS DE MENOR CUANTIA
##########################################################

#Verificación de nuevos resultados licitaciones

resultados_hoy_abreviadas = pandas.merge(resultados_abreviada,resultados_total_abreviadas, 
                              how = 'left', 
                              on = ['Num_Proceso','Link','Entidad'])

resultados_hoy_abreviadas = resultados_hoy_abreviadas[resultados_hoy_abreviadas.ID_y.isnull()]
resultados_hoy_abreviadas = resultados_hoy_abreviadas.iloc[:,:12]

header_totals = resultados_total_abreviadas.columns
resultados_hoy_abreviadas.columns = header_totals
resultados_total_abreviadas = resultados_total_abreviadas.append(resultados_hoy_abreviadas)
resultados_total_abreviadas.to_excel(PATH + 'resultados_total_Abreviada.xls', index = False)

# Escribir resultados del día
writer = pandas.ExcelWriter(PATH + 'resultados_hoy_abreviadas' + ar + '.xlsx',engine = 'xlsxwriter')
resultados_hoy_abreviadas.to_excel(writer, sheet_name = 'Resumen_resultados')
workbook  = writer.book
worksheet = writer.sheets['Resumen_resultados']
formatdate = workbook.add_format({'valign': 'vcenter'})
format1 = workbook.add_format({'text_wrap': True,
                               'valign': 'vcenter'})
format2 = workbook.add_format({'num_format': '$ #,##0.00',
                               'valign': 'vcenter'})
formatall = workbook.add_format({'valign': 'vcenter',
                                 'align': 'center'})                               

worksheet.set_column('A:O',None,formatall)
worksheet.set_column('B:B',5)
worksheet.set_column('C:C',15,format1)
worksheet.set_column('D:D',15,format1)
worksheet.set_column('E:E',9.45)
worksheet.set_column('F:F',25,format1)
worksheet.set_column('G:G',53,format1)
worksheet.set_column('H:H',18, format1)
worksheet.set_column('I:I',12.35,format1)
worksheet.set_column('E:E',9.45)
worksheet.set_column('J:J',16,format2)
worksheet.set_column('L:L',17.30, formatdate)
worksheet.set_column('M:M',85,format1)
worksheet.set_column('A:A',None,None,{'hidden': True})
worksheet.set_column('K:K',None,None,{'hidden': True})

writer.save()

#resultados_hoy_abreviadas.to_excel(PATH + 'resultados_hoy_abreviadas' + ar + '.xls', index = False)

#Enviar email con los resultados del día actualizados
#####################
# Contratacion de méritos abiertos
#####################

asunto = 'Resultados SECOP Concurso de Meritos Abiertos ' + ar
COMMASPACE = ', '

def main():
    sender = 'indgersas@gmail.com'
    gmail_password = '10752302521'
    recipients = ['glchristianfer@gmail.com','licitaciones@ingecol.co','hecohsas@gmail.com','haguzmanl@unal.edu.co']
    body = 'Se encontraron los siguientes resultados para concurso de meritos abiertos'
    
    # Create the enclosing (outer) message
    outer = MIMEMultipart()
    outer['Subject'] = asunto
    outer['To'] = COMMASPACE.join(recipients)
    outer['From'] = sender
    outer.preamble = 'Resultados.\n'
    outer.attach(MIMEText(body))

    # List of attachments
    attachments = [PATH + 'resultados_hoy_CMA' + ar + '.xls']

    # Add the attachments to the message
    for file in attachments:
        try:
            with open(file, 'rb') as fp:
                msg = MIMEBase('application', "octet-stream")
                msg.set_payload(fp.read())
            encoders.encode_base64(msg)
            msg.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
            outer.attach(msg)
        except:
            print("Unable to open one of the attachments. Error: ")
            raise

    composed = outer.as_string()

    # Send the email
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as s:
            s.ehlo()
            s.starttls()
            s.ehlo()
            s.login(sender, gmail_password)
            s.sendmail(sender, recipients, composed, body)
            s.close()
        print("Email sent!")
    except:
        print("Unable to send the email. Error: ")
        raise

if __name__ == '__main__':
    main()


#####################
# Licitaciones
#####################

def main():
    sender = 'indgersas@gmail.com'
    gmail_password = '10752302521'
    recipients = ['glchristianfer@gmail.com','licitaciones@ingecol.co','hecohsas@gmail.com','haguzmanl@unal.edu.co']
    body = 'Se encontraron los siguientes resultados para procesos de Licitacion Publica'
    
    # Create the enclosing (outer) message
    outer = MIMEMultipart()
    outer['Subject'] = asunto
    outer['To'] = COMMASPACE.join(recipients)
    outer['From'] = sender
    outer.preamble = 'Resultados.\n'
    outer.attach(MIMEText(body))

    # List of attachments
    attachments = [PATH + 'resultados_hoy_licitacion' + ar + '.xls']

    # Add the attachments to the message
    for file in attachments:
        try:
            with open(file, 'rb') as fp:
                msg = MIMEBase('application', "octet-stream")
                msg.set_payload(fp.read())
            encoders.encode_base64(msg)
            msg.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
            outer.attach(msg)
        except:
            print("Unable to open one of the attachments. Error: ")
            raise

    composed = outer.as_string()

    # Send the email
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as s:
            s.ehlo()
            s.starttls()
            s.ehlo()
            s.login(sender, gmail_password)
            s.sendmail(sender, recipients, composed, body)
            s.close()
        print("Email sent!")
    except:
        print("Unable to send the email. Error: ")
        raise

if __name__ == '__main__':
    main()


#####################
# Selección abreviada
#####################


def main():
    sender = 'indgersas@gmail.com'
    gmail_password = '10752302521'
    recipients = ['glchristianfer@gmail.com','licitaciones@ingecol.co','hecohsas@gmail.com','haguzmanl@unal.edu.co']
    body = 'Se encontraron los siguientes resultados para Seleccion abreviada de menor cuantia'
    
    # Create the enclosing (outer) message
    outer = MIMEMultipart()
    outer['Subject'] = asunto
    outer['To'] = COMMASPACE.join(recipients)
    outer['From'] = sender
    outer.preamble = 'Resultados.\n'
    outer.attach(MIMEText(body))

    # List of attachments
    attachments = [PATH + 'resultados_hoy_abreviadas' + ar + '.xls']

    # Add the attachments to the message
    for file in attachments:
        try:
            with open(file, 'rb') as fp:
                msg = MIMEBase('application', "octet-stream")
                msg.set_payload(fp.read())
            encoders.encode_base64(msg)
            msg.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
            outer.attach(msg)
        except:
            print("Unable to open one of the attachments. Error: ")
            raise

    composed = outer.as_string()

    # Send the email
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as s:
            s.ehlo()
            s.starttls()
            s.ehlo()
            s.login(sender, gmail_password)
            s.sendmail(sender, recipients, composed, body)
            s.close()
        print("Email sent!")
    except:
        print("Unable to send the email. Error: ")
        raise

if __name__ == '__main__':
    main()