# -*- coding: utf-8 -*-
"""
Created on Sun Apr 29 17:10:20 2018

@author:
"""
from bs4 import BeautifulSoup
import urllib3
import re
import pandas
import numpy
from datetime import datetime
from datetime import timedelta
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
find_depart = ['Huil','Vall','Cauc','Quind','Tol','Putuma','Caquet','Cundinamar','Nariño','Boyac','Bogot']
find_depart_2 = '|'.join(find_depart)
find_objet = ['obra','alcantarilla','acueduct','consultor','interventor','construcci','saneamient',
              'adecuaci','restauraci','diseñ','topograf','estudi','puent','red','espacio','infrestruct',
              'proyect']
find_objet = '|'.join(find_objet)
# PATH = 'D:/Servidor INDGER/4. Buscador de Contratos/6. SECOP I/Minima/'
PATH = 'C:/Users/90044338/Documents/Administrativo/SECOP/'

prev_link = 'https://www.contratos.gov.co'

#Solicita las fechas para la busqueda en SECOP 1
fi = (datetime.now() + timedelta(days=-2)).strftime('%d/%m/%Y')
ff = (datetime.now() + timedelta(days=7)).strftime('%d/%m/%Y')
ar = datetime.now().strftime('%d_%m_%Y__%H_%M')
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
base_all['Objeto'] = base_all['Objeto'].map(lambda x: x.lower())

base_all['Cuantia'] = base_all['Cuantia'].str.replace('>','')
base_all['Cuantia'] = base_all['Cuantia'].str.replace('<','')
base_all['Cuantia'] = base_all['Cuantia'].str.rstrip('0')
base_all['Cuantia'] = base_all['Cuantia'].str.replace(',','')
base_all['Cuantia'] = base_all['Cuantia'].str.replace('.','')
base_all['Cuantia'] = base_all['Cuantia'].str.replace('$','')
base_all['Cuantia'] = base_all['Cuantia'].astype('int64')

base_all['Fecha'] = base_all['Fecha'].str.replace('>','')
base_all['Fecha'] = base_all['Fecha'].str.replace('<','')
base_all['Fecha'] = pandas.to_datetime(base_all.Fecha)

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
resultados_total = pandas.read_excel(PATH + 'resultados_total.xls', header = 0)

# Filtrar de acuerdo a criterios de Min cuantia
resultado_1 = base_all[(base_all.Tipo_Proceso == "Contratación Mínima Cuantía") &
                          (base_all.Estado == "Convocado") 
                          & (base_all.Departamento.str.contains(find_depart_2)) 
                          & (base_all.Cuantia > 6000000)
                          & (base_all.Objeto.str.contains(find_objet))]

# Verificación de nuevos resultados

resultados_hoy = pandas.merge(resultado_1,resultados_total, 
                              how = 'left', 
                              on = ['Num_Proceso','Link','Entidad'])

resultados_hoy = resultados_hoy[resultados_hoy.ID_y.isnull()]
resultados_hoy = resultados_hoy.iloc[:,:12]

header_totals = resultados_total.columns
resultados_hoy.columns = header_totals
resultados_total = resultados_total.append(resultados_hoy)
resultados_total.to_excel(PATH + 'resultados_total.xls', index = False)

##### Escribir resultados del día en excel
writer = pandas.ExcelWriter(PATH + 'Resultados_hoy' + ar + '.xlsx',engine = 'xlsxwriter')

resultado_1.to_excel(writer, sheet_name = 'Resumen_resultados')

workbook  = writer.book
worksheet = writer.sheets['Resumen_resultados']

formatdate = workbook.add_format({'num_format': 'dd/mm/yyyy',
                                  'valign': 'vcenter'})
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
worksheet.set_column('H:H',12.35)
worksheet.set_column('I:I',12.35,format1)
worksheet.set_column('E:E',9.45)
worksheet.set_column('J:J',16,format2)
worksheet.set_column('L:L',17.30)
worksheet.set_column('M:M',78,formatall)

worksheet.set_column('A:A',None,None,{'hidden': True})
worksheet.set_column('K:K',None,None,{'hidden': True})

writer.save()

# resultados_hoy.to_excel(PATH + 'resultados_hoy' + ar + '.xls', index = False)

#Enviar email con los resultados del día actualizados

asunto = 'Resultados SECOP ' + ar
COMMASPACE = ', '

def main():
    sender = 'indgersas@gmail.com'
    gmail_password = '10752302521'
    recipients = ['glchristianfer@gmail.com','licitaciones@ingecol.co','hecohsas@gmail.com','haguzmanl@unal.edu.co']
    body = 'Se encontraron los siguientes resultados'
    
    # Create the enclosing (outer) message
    outer = MIMEMultipart()
    outer['Subject'] = asunto
    outer['To'] = COMMASPACE.join(recipients)
    outer['From'] = sender
    outer.preamble = 'Resultados.\n'
    outer.attach(MIMEText(body))

    # List of attachments
    attachments = [PATH + 'resultados_hoy' + ar + '.xls']

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