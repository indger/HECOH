# -*- coding: utf-8 -*-
"""
Created on Wed Aug 15 08:52:06 2018

@author: 90044338
"""

from bs4 import BeautifulSoup
import urllib3
import re
import pandas as pd
import numpy
from datetime import datetime
from datetime import timedelta
from selenium import webdriver
import os
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

PATH_SII = 'D:/Servidor INDGER/4. Buscador de Contratos/5. SECOP II/'
url = "https://community.secop.gov.co/Public/Tendering/ContractNoticeManagement/Index?currentLanguage=es-CO&Page=login&Country=CO&SkinName=CCE"

gecko = os.path.normpath(os.path.join(os.path.dirname('C:/Python36/Scripts/'), 'chromedriver'))

# PATH Oficina: C:/Users/90044338/Documents/German/SECOP/5. SECOP II/
# PATH Casa: D:/Servidor INDGER/4. Buscador de Contratos/5. SECOP II/
# C:/Program Files/Python36/Scripts/
# C:/Python36/Scripts/

# obtener las fechas para la búsqueda
fi = (datetime.now() + timedelta(days=-5)).strftime('%d/%m/%Y')
ff = (datetime.now() + timedelta(days=2)).strftime('%d/%m/%Y')
fechainicio = fi + ' 12:00 AM'
fechafin = ff + ' 12:00 AM'

# crear variable a incluir en el nombre del archivo de resultados del dia
ar = datetime.now().strftime('%d_%m_%Y')

# Lista con los tipos de procesos a busar
minima = ['Concurso de méritos abierto']
pre_table = []

for tipo in minima:  

    binary = FirefoxBinary(r'C:\Program Files (x86)\Mozilla Firefox\firefox.exe')
    driver = webdriver.Chrome(executable_path=gecko+'.exe')
    # create a new Firefox session
    driver.implicitly_wait(30)
    driver.get(url)
    
    # Click en busqueda avanzada
    Busqueda_avanzada = driver.find_element_by_id('lnkAdvancedSearch') 
    Busqueda_avanzada.click() #click en Busqueda avanzada
    
    # Coloca fecha Inicial
    driver.execute_script("document.getElementById('dtmbPublishDateFrom_txt').setAttribute('Value','%(fechainicio)s')" % locals())
    
    # Coloca Fecha Final
    driver.execute_script("document.getElementById('dtmbPublishDateTo_txt').setAttribute('Value','%(fechafin)s')" % locals())
    
    # Buscar por tipo de proceso
    driver.find_element_by_xpath("//select[@name='VB_selProcedureType']/option[text()='%(tipo)s']" % locals()).click()
    
    # Click en boton Buscar
    driver.implicitly_wait(30)
    Buscar = driver.find_element_by_id('btnSearchButton') 
    driver.execute_script("arguments[0].click();", Buscar)
    
    bs = BeautifulSoup(driver.page_source, 'html')
    
    # Buscar key de la pagina
    UR_key = 'tblMainTable_trRowMiddle_tdCell1_tblForm_trGridRow_tdCell1_grdResultList_Paginator_goToPage_MoreItems'
    pagkey = bs.findAll('a', {'id':UR_key})
    pagkey = str(pagkey[0])
    pagkey = re.findall(r'\?.*\{', pagkey)[0]
    pagkey = pagkey.replace('?mkey=','')
    pagkey = pagkey.replace("', {",'')
    
    Startidx = 0
    Endidx = 0
    pagnumber = 0
    Starindex = 1
    Endindex = 5
    
    # Crea la url base para cada una de las paginas de resultados
    url_inicio_1 = 'https://community.secop.gov.co/Public/Tendering/ContractNoticeManagement/ResultListGoToPage?mkey='
    url_inicio_2 = '&startIdx='
    url_inicio_3 = '&endIdx='
    url_inicio_4 = '&pageNumber='
    url_inicio_5 = '&perspective=All&initAction=Index&externalId=&logicalId=&fromMarketplace=&authorityVat=&allWords2Search=&startIndex='
    url_inicio_6 = '&endIndex='
    url_inicio_7 = '&currentPagingStyle=0&displayAdvancedParams=&orderParam=RequestOnlinePublishingDateDESC&searchExecuted=True&authorityName=&reference=&description=&mainCategory=&mainCategoryText=&categorizationSystemCode=&country=&region=&regulation=&requestStatus=&publishDateFrom='
    url_inicio_8 = '&publishDateTo='
    url_inicio_9 = '&tendersDeadlineFrom=&tendersDeadlineTo=&openDateFrom=&openDateTo='
    
    url_inicio = url_inicio_1 + pagkey + url_inicio_2 + str(Startidx) + url_inicio_3 + str(Endidx) + url_inicio_4 + str(pagnumber) + url_inicio_5 + str(Starindex) + url_inicio_6 + str(Endindex) + url_inicio_7 + fechainicio + url_inicio_8 + fechafin + url_inicio_9
    
    # Conexion con cada uno de las urls
    driver.get(url_inicio)
    bs_1 = BeautifulSoup(driver.page_source, 'html')
    
    # obtener cantidad de procesos por pagina
    procesos = bs_1.findAll('td', {'id':'tblMainTable_trRowMiddle_tdCell1_tblForm_trGridRow_tdCell1_grdResultListtd_thAuthorityNameCol'})
    
    i = 0
    
    # ciclo por cada una de las hojas de resultados, se obtienen ID entidad, fecha de publicacion y codigo para descarga individual
    while len(procesos) > 0:
        
        url_inicio = url_inicio_1 + pagkey + url_inicio_2 + str(Startidx) + url_inicio_3 + str(Endidx) + url_inicio_4 + str(pagnumber) + url_inicio_5 + str(Starindex) + url_inicio_6 + str(Endindex) + url_inicio_7 + fechainicio + url_inicio_8 + fechafin + url_inicio_9
        
        driver.get(url_inicio)
        bs_1 = BeautifulSoup(driver.page_source, 'html')
    
        procesos = bs_1.findAll('td', {'id':'tblMainTable_trRowMiddle_tdCell1_tblForm_trGridRow_tdCell1_grdResultListtd_thAuthorityNameCol'})
        
        for x in procesos:
        
            UR_entidad = 'tblMainTable_trRowMiddle_tdCell1_tblForm_trGridRow_tdCell1_grdResultList_tdAuthorityNameCol_spnMatchingResultAuthorityName_' + str(i)
            UR_fecha_publi = 'dtmbRequestOnlinePublishingDate_' + str(i) + '_txt'
            UR_codigo = 'tblMainTable_trRowMiddle_tdCell1_tblForm_trGridRow_tdCell1_grdResultList_tdDetailColumn_lnkDetailLink_' + str(i)
        
            pre_table.append(str(i))
            entidad = bs_1.findAll('span', {'id':UR_entidad})
            if len(entidad) == 1:
                pre_table.append(entidad[0].text)
            else:
                pre_table.append('Error')
            
            fecha_publi = bs_1.findAll('span', {'id':UR_fecha_publi})
            if len(fecha_publi) == 1:
                pre_table.append(fecha_publi[0].text)
            else:
                pre_table.append('Error')
            
            codigo = bs_1.findAll('a', {'id':UR_codigo})
            if len(codigo) == 1:
                pre_table.append(str(codigo[0]))
            else:
                pre_table.append('Error')
                break
            i = i + 1
        
        if len(codigo) == 0:
            break
              
        Startidx = Startidx + 5
        pagnumber = pagnumber + 1
        
        if pagnumber == 1:
            Endidx = 9
        else:
            Endidx = Endidx + 5
        
        if pagnumber == 1:
            Starindex = 1
        else:
            Starindex = Starindex + 5
        
        if pagnumber == 1:
            Endindex = 5
        else:
            Endindex = Endindex + 5

#####################################################
# Fin del ciclo por cada uno de los tipos de procesos
#####################################################

# Crea el primer data Frame con los resultados iniciales  
titulos_ini = ['ID','Entidad','Fecha_Publicacion','Codigo']
tabla_ini = pd.DataFrame(numpy.array(pre_table).reshape(-1,len(titulos_ini)), columns = titulos_ini)

# Eliminar errores
tabla_ini = tabla_ini[tabla_ini.Codigo != 'Error']

# tab_codigos = pd.DataFrame({'valores':codigos})
# tab_codigos = tab_codigos.astype(str)
tabla_ini['Codigo'] = tabla_ini['Codigo'].str.extract('(\'CO.*\d\')')
tabla_ini['Codigo'] = tabla_ini['Codigo'].str.replace("'",'')

# tab_codigos = tab_codigos['codigo']
# tab_codigos = pd.DataFrame(tab_codigos)
# tab_codigos = tab_codigos.drop_duplicates()

# Inicio y Final de la url del Link de los procesos encontrados
Inicio_url = 'https://community.secop.gov.co/Public/Tendering/OpportunityDetail/Index?noticeUID='
fin_url = '&isFromPublicArea=True&isModal=true&asPopupView=true'

# Crear el Link del proceso con el código del proceso
tabla_ini['Link'] = Inicio_url + tabla_ini['Codigo'] + fin_url

# Leer resultados totales SECOP II para procesos minimas cuantias
resultados_total_SII_MA = pd.read_excel(PATH_SII + 'resultados_total_SII_MA.xls', header = 0)

#Cruzar base de resultados total, con la bse de resultados del dia
tabla_ini = pd.merge(tabla_ini,resultados_total_SII_MA, 
                              how = 'left', 
                              on = ['Link','Entidad'])

# Eliminar los procesos que ya habían salido en una busqueda anterior, osea que ya estaban en la tabla de resultados total
tabla_ini = tabla_ini[tabla_ini.ID_y.isnull()]
tabla_ini = tabla_ini.iloc[:,:5]

# Cambiar los encabezados
titulos_ini_2 = ['ID','Entidad','Fecha_Publicacion','Codigo','Link']
tabla_ini.columns = titulos_ini_2

#############################################################################
######## Segundo proceso descargar info de cada uno de los proesos ##########
#############################################################################

# Ruta html para buscar cada una de las variables del Data Frame
f_num_proceso = 'fdsRequestSummaryInfo_tblDetail_trRowRef_tdCell2_spnRequestReference'
f_tip_proceso = 'fdsRequestSummaryInfo_tblDetail_trRowProcedureType_tdCell2_spnProcedureType'
f_estado = 'fdsRequestSummaryInfo_tblDetail_trRowState_tdCell2_spnState'
f_entidad = 'fdsRequestSummaryInfo_tblDetail_trRowBuyer_tdCell1_ctzBusinessCard' # Name class
f_objeto = 'fdsRequestSummaryInfo_tblDetail_trRowDescription_tdCell2_spnDescription'
f_departamento = 'fdsObjectOfTheContract_tblDetail_trRowPlaceOfWorks_tdCell2_spnspnPlaceOfWorks'
f_municipio = 'fdsObjectOfTheContract_tblDetail_trRowPlaceOfWorks_tdCell2_spnspnPlaceOfWorks'
f_cuantia = 'cbxBasePriceValue'
f_estado_actual = 'fdsRequestSummaryInfo_tblDetail_trRowPhase_tdCell2_spnPhase'

# Crear lista para guardar datos de cada uno de los procesos
registros = []
a = 0
busquedas = [30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60]

# Loop para entrar a cada uno de los links e ir obteniendo los datos de los procesos
for x in tabla_ini['Link']:
    url = x
    http = urllib3.PoolManager()
    ini = http.request('GET', url)
    bs = BeautifulSoup(ini.data)
    
    registros.append(str(a))
    registros.append(url)
    num_proceso = bs.findAll('span', {'id':f_num_proceso})
    if len(num_proceso) == 1:
        registros.append(num_proceso[0].text)
    else:
        registros.append('No encontrado')
        
    tip_proceso = bs.findAll('span', {'id':f_tip_proceso})
    if len(tip_proceso) == 1:
        registros.append(tip_proceso[0].text)
    else:
        registros.append('No encontrado')
        
    estado = bs.findAll('span', {'id':f_estado})
    if len(estado) == 1:
        registros.append(estado[0].text)
    else:
        registros.append('No encontrado')
    
    objeto = bs.findAll('span', {'id':f_objeto})
    if len(objeto) == 1:
        registros.append(objeto[0].text)
    else:
        registros.append('No encontrado')

    departamento = bs.findAll('span', {'id':f_departamento})
    if len(departamento) == 1:
        registros.append(departamento[0].text)
    else:
        registros.append('No encontrado')

    municipio = bs.findAll('span', {'id':f_municipio})
    if len(municipio) == 1:
        registros.append(municipio[0].text)
    else:
        registros.append('No encontrado')
    
    cuantia = bs.findAll('span', {'id':f_cuantia})
    if len(cuantia) == 1:
        registros.append(cuantia[0].text)
    else:
        registros.append('0')

    estado_actual = bs.findAll('span', {'id':f_estado_actual})
    if len(estado_actual) == 1:
        registros.append(estado_actual[0].text)
    else:
        registros.append('No encontrado')
    
    evalu = len(registros)
    for i in busquedas:
        pimes = bs.findAll('tr', {'id':'trScheduleDateRow_' + str(i)})
        if len(pimes) == 1:
            if 'Deadline to require SME limitation' in pimes[0].text:
                registros.append(pimes[0].text)
                break
            else:
                pass
        else:
            pass
    if evalu == len(registros):
        registros.append('NA')
    else:
        pass
    
    evalu = len(registros)
    for i in busquedas:
        interes = bs.findAll('tr', {'id':'trScheduleDateRow_' + str(i)})
        if len(interes) == 1:
            if 'Deadline to show interest' in interes[0].text:
                registros.append(interes[0].text)
                break
            else:
                pass
        else:
            pass
    if evalu == len(registros):
        registros.append('NA')
    else:
        pass
    
    evalu = len(registros)
    for i in busquedas:
        sorteo = bs.findAll('tr', {'id':'trScheduleDateRow_' + str(i)})
        if len(sorteo) == 1:
            if 'Lottery Date' in sorteo[0].text:
                registros.append(sorteo[0].text)
                break
            else:
                pass
        else:
            pass
    if evalu == len(registros):
        registros.append('NA')
    else:
        pass
    
    evalu = len(registros)
    for i in busquedas:
        pubsorteo = bs.findAll('tr', {'id':'trScheduleDateRow_' + str(i)})
        if len(pubsorteo) == 1:
            if 'Lottery Publication' in pubsorteo[0].text:
                registros.append(pubsorteo[0].text)
                break
            else:
                pass
        else:
            pass
    if evalu == len(registros):
        registros.append('NA')
    else:
        pass
    
    evalu = len(registros)
    for i in busquedas:
        oferta = bs.findAll('tr', {'id':'trScheduleDateRow_' + str(i)})
        if len(oferta) == 1:
            if 'Due date for receiving replies' in oferta[0].text:
                registros.append(oferta[0].text)
                break
            else:
                pass
        else:
            pass
    if evalu == len(registros):
        registros.append('NA')
    else:
        pass
    
    evalu = len(registros)
    for i in busquedas:
        apertura = bs.findAll('tr', {'id':'trScheduleDateRow_' + str(i)})
        if len(apertura) == 1:
            if 'Opening replies date' in apertura[0].text:
                registros.append(apertura[0].text)
                break
            else:
                pass
        else:
            pass
    if evalu == len(registros):
        registros.append('NA')
    else:
        pass
        
    print(a)
    a = a + 1

# Data Frame con información complementaria de los procesos
titulos_fin = ['ID','Link','Num_Proceso','Tipo_Proceso','Estado','Objeto', 'Departamento', 'Municipio', 'Cuantia', 'Estado_actual',
               'Plazo MI pymes','Plazo MI','Sorteo','Precalificados','Presentacion ofertas','Apertura ofertas']
tabla_fin = pd.DataFrame(numpy.array(registros).reshape(-1,len(titulos_fin)), columns = titulos_fin)

# Complementar los dos Data Frame descargados
tabla_com = pd.merge(tabla_ini,tabla_fin, 
                              how = 'inner', 
                              on = ['Link'])

# Corregir datos de variables
tabla_com['Objeto'] = tabla_com['Objeto'].map(lambda x: x.replace('\n','').strip())
tabla_com['Departamento'] = tabla_com['Departamento'].map(lambda x: x.replace('\n','').strip())
tabla_com['Municipio'] = tabla_com['Municipio'].map(lambda x: x.replace('\n','').strip())

tabla_com['Cuantia'] = tabla_com['Cuantia'].str.replace(',','')
tabla_com['Cuantia'] = tabla_com['Cuantia'].str.replace(' COP','')
tabla_com['Cuantia'] = tabla_com['Cuantia'].astype('float64')

tabla_com['Fecha'] = tabla_com['Fecha_Publicacion'].str.extract('(^\d{2,}\/\d{2,}\/\d{4,})')
tabla_com = tabla_com.drop(columns = ['Fecha_Publicacion'])

tabla_com['Presentacion ofertas'] = tabla_com['Presentacion ofertas'].str.extract('(\(\d.*\()')
tabla_com['Presentacion ofertas'] = tabla_com['Presentacion ofertas'].str.replace('(','')
tabla_com['Presentacion ofertas'] = tabla_com['Presentacion ofertas'].str.replace(')','')

tabla_com['Plazo MI pymes'] = tabla_com['Plazo MI pymes'].str.extract('(\(\d.*\()')
tabla_com['Plazo MI pymes'] = tabla_com['Plazo MI pymes'].str.replace('(','')
tabla_com['Plazo MI pymes'] = tabla_com['Plazo MI pymes'].str.replace(')','')

tabla_com['Plazo MI'] = tabla_com['Plazo MI'].str.extract('(\(\d.*\()')
tabla_com['Plazo MI'] = tabla_com['Plazo MI'].str.replace('(','')
tabla_com['Plazo MI'] = tabla_com['Plazo MI'].str.replace(')','')

tabla_com['Sorteo'] = tabla_com['Sorteo'].str.extract('(\(\d.*\()')
tabla_com['Sorteo'] = tabla_com['Sorteo'].str.replace('(','')
tabla_com['Sorteo'] = tabla_com['Sorteo'].str.replace(')','')

tabla_com['Precalificados'] = tabla_com['Precalificados'].str.extract('(\(\d.*\()')
tabla_com['Precalificados'] = tabla_com['Precalificados'].str.replace('(','')
tabla_com['Precalificados'] = tabla_com['Precalificados'].str.replace(')','')

tabla_com['Apertura ofertas'] = tabla_com['Apertura ofertas'].str.extract('(\(\d.*\()')
tabla_com['Apertura ofertas'] = tabla_com['Apertura ofertas'].str.replace('(','')
tabla_com['Apertura ofertas'] = tabla_com['Apertura ofertas'].str.replace(')','')

tabla_com['ID'] = tabla_com['ID_y']
tabla_com = tabla_com.drop(columns = ['ID_x','ID_y'])

# Ordenar Data Frame
cols = ['ID', 'Num_Proceso', 'Tipo_Proceso', 'Estado', 'Entidad', 'Objeto', 'Departamento', 'Municipio', 'Cuantia', 'Estado_actual',
        'Fecha', 'Link', 'Codigo','Plazo MI pymes','Plazo MI','Sorteo','Precalificados','Presentacion ofertas','Apertura ofertas']
tabla_com = tabla_com[cols]

# Filtrar resultados para obtener solo los procesos de minima cuantia
resultados_MA_hoy_SII = tabla_com[(tabla_com.Estado == "Published")]
    
# Incluir los resultados nuevos del día, en el historico de resultados total
resultados_total_SII_MA = resultados_total_SII_MA.append(tabla_com)

# Exportar en excel las tablas de resultados total y resultados del dia
resultados_total_SII_MA.to_excel(PATH_SII + 'resultados_total_SII_MA.xls', index = False)
resultados_MA_hoy_SII.to_excel(PATH_SII + 'resultados_MA_hoy_SII' + ar +'.xls', index = False)

#Enviar email con los resultados del día actualizados

asunto = 'Resultados SECOP II Méritos abiertos.' + ar
COMMASPACE = ', '

def main():
    sender = 'indgersas@gmail.com'
    gmail_password = '10752302521'
    recipients = ['glchristianfer@gmail.com','licitaciones@ingecol.co','hecohsas@gmail.com','haguzmanl@unal.edu.co']
    body = 'Se encontraron los siguientes resultados para procesos de Meritos abiertos en el SECOP II'
    
    # Create the enclosing (outer) message
    outer = MIMEMultipart()
    outer['Subject'] = asunto
    outer['To'] = COMMASPACE.join(recipients)
    outer['From'] = sender
    outer.preamble = 'Resultados.\n'
    outer.attach(MIMEText(body))

    # List of attachments
    attachments = [PATH_SII + 'resultados_MA_hoy_SII' + ar + '.xls']

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
