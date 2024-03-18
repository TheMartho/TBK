from datetime import datetime
import time

from unittest import mock

import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains

from zipfile import ZipFile

import os
import bd_tbk

import os
from dotenv import load_dotenv

load_dotenv()
now = datetime.now()

def func_descarga_archivos():
    new     = str(datetime.now().strftime("%Y-%m-%d"))
    file    = 'PreciosFex_(' + new + ')' + '.xls'
    url     = 'https://privado.transbank.cl/'
    uss_    = os.getenv('ACCESS_TBK_USER')
    pass_   = os.getenv('ACCESS_TBK_PASS')
    options = webdriver.EdgeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--safebrowsing-disable-download-protection");
    options.add_argument("safebrowsing-disable-extension-blacklist");
    driver = webdriver.Edge(options)
    driver.get(url)
    time.sleep(2)
    user = driver.find_element(By.XPATH,'/html/body/div[7]/section/div/div/div[2]/div/div[1]/section/div/div/div/div[2]/div/div/div[3]/form/div[1]/div[1]/div/input')
    user.send_keys(uss_)
    passw = driver.find_element(By.XPATH,'/html/body/div[7]/section/div/div/div[2]/div/div[1]/section/div/div/div/div[2]/div/div/div[3]/form/div[1]/div[3]/input')
    passw.send_keys(pass_)
    time.sleep(2)    
    boton_acceso = driver.find_element(By.XPATH,'/html/body/div[7]/section/div/div/div[2]/div/div[1]/section/div/div/div/div[2]/div/div/div[3]/form/div[2]/button')
    boton_acceso.click()
    time.sleep(20)
    boton_trx = driver.find_element(By.XPATH,'/html/body/header/div[2]/section/div/div/div/div/div/div/ul/li[6]/a/span')
    boton_trx.click()
    time.sleep(2)
    boton_liquidacion = driver.find_element(By.XPATH,'/html/body/header/div[2]/section/div/div/div/div/div/div/ul/li[6]/ul/li[1]/a/span')
    boton_liquidacion.click()
    time.sleep(2)
    boton_descarga_masiva = driver.find_element(By.XPATH,'/html/body/header/div[3]/section/div/div/div/div/div/ul/li[4]/a')
    boton_descarga_masiva.click()
    time.sleep(2)
    main_window = driver.current_window_handle
    iframe = driver.find_element(By.ID, "_cl_tbk_portal10_iframe_web_Portal10IframeWebPortlet_INSTANCE_iwOacqqxsuVA_")
    driver.switch_to.frame(iframe)
    time.sleep(1)
    fecha_desde = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div/div/div/div/div/div[1]/div/div[1]/div/div[1]/div/form/table/tbody/tr[2]/td[4]/table/tbody/tr/td[1]/input')
    dt = datetime.now()
    if dt.day >9:
        fecha_desde.send_keys(Keys.BACKSPACE)
        fecha_desde.send_keys(Keys.BACKSPACE)        
    else:
        fecha_desde.send_keys(Keys.BACKSPACE)       
    fecha_desde.send_keys(1)
    time.sleep(1)
    boton_local = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div/div/div/div/div/div[1]/div/div[2]/div/div[2]/div/div[2]/div/div/div/img') 
    boton_local.click()
    time.sleep(4)
    #boton_desmarcar = driver.find_element(By.XPATH, '/html/body/div[4]/div/div[2]/div/div/div/div[2]/div/div[2]/div/table/tbody/tr/td')
    #boton_desmarcar.click()  
    #time.sleep(1)
    boton_desmarcar = driver.find_element(By.XPATH, '/html/body/div[4]/div/div[2]/div/div/div/div[2]/div/div[2]/div/table/tbody/tr/td')
    boton_desmarcar.click()
    time.sleep(1)  
    for i in range (0,10):
        try:
            boton_MC1 = driver.find_element(By.CSS_SELECTOR, '.txpms-option:nth-child('+str(i)+') > input')
            inner_text = boton_MC1.get_attribute("outerHTML")
            if '07737092' in inner_text:
                boton_MC1.click()
                #print(inner_text)
            elif '07737084' in inner_text:
                boton_MC1.click()
                #print(inner_text)
        except NoSuchElementException:
            boton_MC1 = 0    
    #boton_MC1 = driver.find_element(By.XPATH, '/html/body/div[4]/div/div[2]/div/div/div/div[1]/div/div[2]/div/div/div/div[6]/input')
    #boton_MC1.click()
    #time.sleep(1)
    #boton_MC2 = driver.find_element(By.XPATH, '/html/body/div[4]/div/div[2]/div/div/div/div[1]/div/div[2]/div/div/div/div[7]/input')
    #boton_MC2.click()
    time.sleep(2)
    confirm = driver.find_element(By.XPATH, '/html/body/div[4]/div/div[2]/div/div/div/div[3]/div/div[1]/div/table/tbody/tr/td')
    confirm.click()
    time.sleep(1)
    solicitar = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div/div/div/div/div/div[1]/div/div[3]/div/div[1]/div/table/tbody/tr/td')
    solicitar.click()
    time.sleep(1)
    confirm = driver.find_element(By.XPATH, '/html/body/div[6]/div/div[1]/div/div[2]/div/div[2]/div/div[1]/div/table/tbody/tr/td')
    confirm.click()
    time.sleep(1)
    driver.switch_to.default_content()
    doc_electronico = driver.find_element(By.XPATH, '/html/body/header/div[2]/section/div/div/div/div/div/div/ul/li[9]/a/span')
    doc_electronico.click()
    time.sleep(20)
    #iframe = driver.find_element(By.ID, '_cl_tbk_iframe_web_IframeCrossWebPortlet_portal30')
    #driver.switch_to.frame(iframe)
    #iframe2 = driver.find_element(By.XPATH, '/html/body/iframe')
    #driver.switch_to.frame(iframe2)
    main_window = driver.current_window_handle
    iframe = driver.find_element(By.XPATH, '/html/body/div[8]/section/div/div/div/div/div[1]/section/div/div/div/iframe')
    driver.switch_to.frame(iframe)
    iframe2 = driver.find_element(By.XPATH, '/html/body/iframe')
    driver.switch_to.frame(iframe2)
    time.sleep(1)
    export = driver.find_element(By.XPATH, '/html/body/div[1]/section/div[3]/div/a[2]/p[2]')
    export.click()
    time.sleep(1)
    dt = datetime.now()
    donwload_day = driver.find_element(By.XPATH,'/html/body/div[1]/section/div[3]/div/div/div[2]/div[1]/input')
    donwload_day.send_keys(dt.day)
    time.sleep(1)
    busqueda = driver.find_element(By.XPATH, '/html/body/div[1]/section/div[3]/div/div/div[4]/button')
    busqueda.click()
    time.sleep(1)
    sorting_file = driver.find_element(By.XPATH, '/html/body/div[1]/section/div[4]/div[4]/div[1]/div/table/thead/tr/th[1]')
    sorting_file.click()
    time.sleep(25)
    nombre = driver.find_element(By.XPATH, '/html/body/div[1]/section/div[4]/div[4]/div[1]/div/table/tbody/tr[1]/td[4]')
    nombre_archivo = nombre.text
    nombre2 = driver.find_element(By.XPATH,'/html/body/div[1]/section/div[4]/div[4]/div[1]/div/table/tbody/tr[2]/td[4]')
    nombre_archivo2 = nombre2.text

    dt = datetime.now()
    dia_uno = "1"
    dia_actual = str(dt.day)
    mes = str(datetime.now().strftime("%m/%Y"))
    espacio = " al "
    espacio2 = " de "
    espacio3 = " "
    #reconocemos cual nombre es credito y cual de debito
    if 'credito' in nombre_archivo:
        b = "Extracción Masiva Credito pesos del "
        c = "Extracción Masiva Débito pesos del "
        full_name_debito    = c + dia_uno + espacio + dia_actual + espacio2 + mes + espacio3
        full_name_credito   = b + dia_uno + espacio + dia_actual + espacio2 + mes + espacio3
        nombre_final_debito  = nombre_archivo2.replace(full_name_debito,"")
        nombre_final_credito = nombre_archivo.replace(full_name_credito,"")
    else:
        c = "Extracción Masiva Credito pesos del "
        b = "Extracción Masiva Débito pesos del "
        full_name_debito    = b + dia_uno + espacio + dia_actual + espacio2 + mes + espacio3
        full_name_credito   = c + dia_uno + espacio + dia_actual + espacio2 + mes + espacio3
        nombre_final_debito  = nombre_archivo.replace(full_name_debito,"")
        nombre_final_credito = nombre_archivo2.replace(full_name_credito,"")
    #
    #
    nombre_archivo = nombre_archivo[55:] # se selecciona el nombre del archivo como queda guardado en carpeta
    time.sleep(1)
    download_file = driver.find_element(By.XPATH,'/html/body/div[1]/section/div[4]/div[4]/div[1]/div/table/tbody/tr[1]/td[6]/table/tbody/tr[1]/td/a/i')
    download_file.click()
    time.sleep(5)
    download_file2 = driver.find_element(By.XPATH,'/html/body/div[1]/section/div[4]/div[4]/div[1]/div/table/tbody/tr[2]/td[6]/table/tbody/tr[1]/td/a/i')
    download_file2.click()
    time.sleep(20)
    # hay dias en que viene el archivo con formato .dat
    nombre_final_credito_dat = nombre_final_credito.replace("zip","dat")
    nombre_final_debito_dat = nombre_final_debito.replace("zip","dat")
    archivo_comprimido_debito = os.getenv('RUTA_CARPETA') + nombre_final_debito
    archivo_comprimido_credito = os.getenv('RUTA_CARPETA') + nombre_final_credito    
    #Extraemos el archivo de debito que viene con formato zip
    with ZipFile(archivo_comprimido_debito, "r") as zip:
        zip.printdir()
        zip.extractall(os.getenv('RUTA_CARPETA'))

    #Extraemos el archivo de debito que viene con formato zip
    with ZipFile(archivo_comprimido_credito, "r") as zip:
        zip.printdir()
        zip.extractall(os.getenv('RUTA_CARPETA'))
    #Borramos el archivo zip
    time.sleep(5)    
    os.remove(archivo_comprimido_debito)
    os.remove(archivo_comprimido_credito)
    bd_tbk.delete_credito_t()
    bd_tbk.delete_debito_t()
    time.sleep(5)    
    insert_excel_to_DWH_debito(nombre_final_debito_dat)
    time.sleep(10)
    insert_excel_to_DWH_credito(nombre_final_credito_dat)

def insert_excel_to_DWH_debito(file_name):
    path = os.getenv('RUTA_CARPETA') + file_name # Cambiar luego la ruta donde quedara el archivo final.
    if os.path.exists(path):
        n = 38 # your line count.
        df = pd.read_csv(path,encoding='latin1' ,skiprows=n, delimiter=';', dtype=str )
        df = df.rename(columns={'Tipo Transacción': 'Tipo_trx','Fecha Venta':'Fecha_venta','Tipo Tarjeta':'Tipo_Tarjeta','Identificador':'Identificador','Tipo Venta':'Tipo_Venta','Código Autorización Venta':'Codigo_autorizacion_venta',
                                'Nº Cuota':'Numero_cuota','Monto Transacción':'Monto_TRX','Monto Afecto':'Monto_afecto','Comisión e IVA Comisión':'Comision_e_iva_comision','Monto Exento':'Monto_exento','N° Boleta':'Numero_boleta','Monto Anulación':'Monto_anulacion',
                                'Monto Retenido':'Monto_retenido','Devolución Comisión':'Devolucion_comision','Monto Retención':'Monto_retencion','Motivo':'Motivo','Período de Cobro':'Periodo_de_cobro','Detalle de cobros u observación':'Detalle_cobros_u_observacion',
                                'Monto':'Monto','IVA':'IVA','Fecha Abono':'Fecha_abono','Cuenta de Abono':'Cuenta_de_abono','Local':'Local'})    
        df2 = df[["Tipo_trx","Fecha_venta","Tipo_Tarjeta","Identificador","Tipo_Venta","Codigo_autorizacion_venta","Numero_cuota","Monto_TRX","Monto_afecto","Comision_e_iva_comision",
        "Monto_exento","Numero_boleta","Monto_anulacion","Monto_retenido","Devolucion_comision","Monto_retencion","Motivo","Periodo_de_cobro","Detalle_cobros_u_observacion",
        "Monto","IVA","Fecha_abono","Cuenta_de_abono","Local"]]
        df2["Tipo_trx"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Fecha_venta"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Tipo_Tarjeta"].fillna("", inplace = True) #Reemplazamos los valores null    
        df2["Identificador"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Tipo_Venta"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Codigo_autorizacion_venta"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Numero_cuota"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Monto_TRX"].fillna("0", inplace = True) #Reemplazamos los valores null     
        df2["Monto_afecto"].fillna("0", inplace = True) #Reemplazamos los valores null      
        df2["Comision_e_iva_comision"].fillna("0", inplace = True) #Reemplazamos los valores null 
        df2["Monto_exento"].fillna("", inplace = True) #Reemplazamos los valores null 
        df2["Numero_boleta"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Monto_anulacion"].fillna("0", inplace = True)
        df2["Monto_retenido"].fillna("0", inplace = True)        
        df2["Monto_retenido"].fillna(0) #Reemplazamos los valores null 
        df2["Devolucion_comision"].fillna("", inplace = True) #Reemplazamos los valores null 
        df2["Monto_retencion"].fillna("", inplace = True) #Reemplazamos los valores null 
        df2["Motivo"].fillna("", inplace = True) #Reemplazamos los valores null           
        df2["Periodo_de_cobro"].fillna("", inplace = True) #Reemplazamos los valores null        
        df2["Detalle_cobros_u_observacion"].fillna("", inplace = True) #Reemplazamos los valores null    
        df2["Monto"].fillna("", inplace = True) #Reemplazamos los valores null    
        df2["IVA"].fillna("", inplace = True) #Reemplazamos los valores null    
        df2["Fecha_abono"].fillna("", inplace = True) #Reemplazamos los valores null       
        df2["Cuenta_de_abono"].fillna("", inplace = True) #Reemplazamos los valores null           
        df2["Local"].fillna("", inplace = True) #Reemplazamos los valores null
        #df2["Tipo_documento"].fillna("", inplace = True) #Reemplazamos los valores null        

        #Seteo de formato de cada variable              
        df2['Tipo_trx']                     = df2['Tipo_trx'].astype(str)
        #df2['Fecha_venta']                  = df2['Fecha_venta'].astype('datetime64[ns]')
        df2['Tipo_Tarjeta']                 = df2['Tipo_Tarjeta'].astype(str)
        df2['Identificador']                = df2['Identificador'].astype(str)
        df2['Tipo_Venta']                   = df2['Tipo_Venta'].astype(str)
        df2['Codigo_autorizacion_venta']    = df2['Codigo_autorizacion_venta'].astype(str)
        df2['Numero_cuota']                 = df2['Numero_cuota'].astype(str)
        #df2['Monto_TRX']                    = df2["Monto_TRX"].replace(".","") #Reemplazamos los valores null  
        #df2['Monto_TRX']                    = df2["Monto_TRX"].replace(",","") #Reemplazamos los valores null           
        df2['Monto_TRX']                    = df2['Monto_TRX'].astype(str)
        df2['Monto_afecto']                 = df2['Monto_afecto'].astype(str)
        df2['Comision_e_iva_comision']      = df2['Comision_e_iva_comision'].astype(str)
        df2['Monto_exento']                 = df2['Monto_exento'].astype(str)
        df2['Numero_boleta']                = df2['Numero_boleta'].astype(str)
        df2['Monto_anulacion']              = df2['Monto_anulacion'].astype(str)
        df2['Monto_retenido']               = df2['Monto_retenido'].astype(str)
        df2['Devolucion_comision']          = df2['Devolucion_comision'].astype(str)
        df2['Monto_retencion']              = df2['Monto_retencion'].astype(str)
        df2['Motivo']                       = df2['Motivo'].fillna("0")
        df2['Motivo']                       = df2['Motivo'].astype(str)
        df2['Periodo_de_cobro']             = df2['Periodo_de_cobro'].fillna("0")
        df2['Periodo_de_cobro']             = df2['Periodo_de_cobro'].astype(str)
        df2['Detalle_cobros_u_observacion'] = df2['Detalle_cobros_u_observacion'].fillna("0")
        df2['Detalle_cobros_u_observacion'] = df2['Detalle_cobros_u_observacion'].astype(str)
        df2['Monto']                        = df2['Monto'].fillna("0")
        df2['Monto']                        = df2['Monto'].astype(str)
        df2['IVA']                          = df2['IVA'].fillna("0")
        df2['IVA']                          = df2['IVA'].astype(str)

        df2['Fecha_abono']                  = df2['Fecha_abono'].astype(str)
        df2['Cuenta_de_abono']              = df2['Cuenta_de_abono'].astype(str)
        df2['Local']                        = df2['Local'].astype(str)
        df2['Tipo_documento']               = ''#df2['Tipo_documento'].astype(str)
        now = datetime.now()
        current_time = now.strftime("%H:%M:%S")
        print("Tiempo inicio insercion BD =", current_time)
        bd_tbk.insert_TRXTBK_debito_fastSql(df2)
        now = datetime.now()
        current_time_end = now.strftime("%H:%M:%S")
        print("Tiempo Fin insercion BD =", current_time_end)


def insert_excel_to_DWH_credito(file_name):
    path = os.getenv('RUTA_CARPETA') + file_name # Cambiar luego la ruta donde quedara el archivo final.
    if os.path.exists(path):
        n = 38 # your line count.
        dtype_dict = {'Comisión Adicional e IVA Comisión Adicional': 'str', 'Devolución Comisión Adicional e IVA Comisión': 'str', 'Monto Retención': 'str', 'Período de Cobro':'str'}
        df = pd.read_csv(path,encoding='latin1' ,skiprows=n, delimiter=';', dtype=dtype_dict )
        df = df.rename(columns={'Tipo Transacción': 'Tipo_trx','Fecha Venta':'Fecha_venta','Tipo Tarjeta':'Tipo_Tarjeta','Identificador':'Identificador','Tipo Cuota':'Tipo_Cuota','Monto Original Venta':'Monto_original_venta',
                                'Código Autorización Venta':'Codigo_autorizacion_venta','Nº Cuota':'Numero_cuota','Monto Para Abono':'Monto_para_abono','Comisión e IVA Comisión':'Comision_e_iva_comision',
                                'Comisión Adicional e IVA Comisión Adicional':'Comision_adicional_e_iva_comision_adicional','N° Boleta':'Numero_boleta','Monto Anulación':'Monto_anulacion','Devolución Comisión e IVA Comisión':'Devolucion_comision_e_iva_comision',
                                'Devolución Comisión Adicional e IVA Comisión':'Devolucion_comision_adicional_e_iva_comision','Monto Retención':'Monto_retencion','Período de Cobro':'Periodo_de_cobro','Motivo':'Motivo',
                                'Detalle de cobros u observación':'Detalle_cobros_u_observacion','Monto':'Monto','IVA':'IVA','Fecha Abono':'Fecha_abono','Cuenta de Abono':'Cuenta_de_abono','Local':'Local'})    
        df2 = df[["Tipo_trx","Fecha_venta","Tipo_Tarjeta","Identificador","Tipo_Cuota","Monto_original_venta","Codigo_autorizacion_venta","Numero_cuota","Monto_para_abono","Comision_e_iva_comision","Comision_adicional_e_iva_comision_adicional","Numero_boleta",
                "Monto_anulacion","Devolucion_comision_e_iva_comision","Devolucion_comision_adicional_e_iva_comision","Monto_retencion","Periodo_de_cobro","Motivo","Detalle_cobros_u_observacion","Monto","IVA","Fecha_abono","Cuenta_de_abono","Local"]]
        df2["Tipo_trx"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Fecha_venta"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Tipo_Tarjeta"].fillna("", inplace = True) #Reemplazamos los valores null    
        df2["Identificador"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Tipo_Cuota"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Monto_original_venta"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Codigo_autorizacion_venta"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Numero_cuota"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Monto_para_abono"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Comision_e_iva_comision"].fillna("", inplace = True) #Reemplazamos los valores null      
        df2["Comision_adicional_e_iva_comision_adicional"].fillna("", inplace = True) #Reemplazamos los valores null 
        df2["Numero_boleta"].fillna("", inplace = True) #Reemplazamos los valores null 
        df2["Monto_anulacion"].fillna("", inplace = True) #Reemplazamos los valores null
        df2["Monto_anulacion"].fillna("0", inplace = True)
        df2["Devolucion_comision_e_iva_comision"].fillna(0) #Reemplazamos los valores null 
        df2["Devolucion_comision_adicional_e_iva_comision"].fillna("", inplace = True) #Reemplazamos los valores null 
        df2["Monto_retencion"].fillna("", inplace = True) #Reemplazamos los valores null 
        df2["Periodo_de_cobro"].fillna("", inplace = True) #Reemplazamos los valores null 
        df2["Motivo"].fillna("", inplace = True) #Reemplazamos los valores null          
        df2["Detalle_cobros_u_observacion"].fillna("", inplace = True) #Reemplazamos los valores null    
        df2["Monto"].fillna("", inplace = True) #Reemplazamos los valores null    
        df2["IVA"].fillna("", inplace = True) #Reemplazamos los valores null    
        df2["Fecha_abono"].fillna("", inplace = True) #Reemplazamos los valores null       
        df2["Cuenta_de_abono"].fillna("", inplace = True) #Reemplazamos los valores null           
        df2["Local"].fillna("", inplace = True) #Reemplazamos los valores null

        #Seteo de formato de cada variable              
        df2['Tipo_trx']                     = df2['Tipo_trx'].astype(str)
        #df2['Fecha_venta']                  = df2['Fecha_venta'].astype('datetime64[ns]')
        df2['Tipo_Tarjeta']                 = df2['Tipo_Tarjeta'].astype(str)
        df2['Identificador']                = df2['Identificador'].astype(str)
        df2['Tipo_Cuota']                   = df2['Tipo_Cuota'].astype(str)
        df2['Monto_original_venta']         = df2['Monto_original_venta'].astype(str)
        df2['Codigo_autorizacion_venta']    = df2['Codigo_autorizacion_venta'].astype(str)
        df2['Numero_cuota']                 = df2['Numero_cuota'].astype(str)
        df2['Monto_para_abono']             = df2['Monto_para_abono'].astype(str)
        #df2['Comision_e_iva_comision']      = df2['Comision_e_iva_comision'].fillna(int(0), inplace=True) 
        df2['Comision_e_iva_comision']      = df2['Comision_e_iva_comision'].astype(str)
        df2['Comision_adicional_e_iva_comision_adicional']  = df2['Comision_adicional_e_iva_comision_adicional'].astype(str)
        df2['Numero_boleta']                = df2['Numero_boleta'].astype(str)
        df2['Monto_anulacion']              = df2['Monto_anulacion'].fillna("0")  
        df2['Monto_anulacion']              = df2['Monto_anulacion'].astype(str)
        df2['Devolucion_comision_e_iva_comision']            = df2['Devolucion_comision_e_iva_comision'].fillna(0)    
        df2['Devolucion_comision_e_iva_comision']            = df2['Devolucion_comision_e_iva_comision'].astype(int)    
        df2['Devolucion_comision_adicional_e_iva_comision']  = df2['Devolucion_comision_adicional_e_iva_comision'].fillna("0")     
        df2['Devolucion_comision_adicional_e_iva_comision']  = df2['Devolucion_comision_adicional_e_iva_comision'].astype(str)
        df2['Monto_retencion']              = df2['Monto_retencion'].fillna("0")      
        df2['Monto_retencion']              = df2['Monto_retencion'].astype(str)
        df2['Periodo_de_cobro']             = df2['Periodo_de_cobro'].fillna("0")     
        df2['Periodo_de_cobro']             = df2['Periodo_de_cobro'].astype(str)
        df2['Motivo']                       = df2['Motivo'].fillna("0")
        df2['Motivo']                       = df2['Motivo'].astype(str)
        df2['Detalle_cobros_u_observacion'] = df2['Detalle_cobros_u_observacion'].fillna("0")      
        df2['Detalle_cobros_u_observacion'] = df2['Detalle_cobros_u_observacion'].astype(str)
        df2['Monto']                        = df2['Monto'].fillna("0")     
        df2['Monto']                        = df2['Monto'].astype(str)
        df2['IVA']                          = df2['IVA'].fillna("0")       
        df2['IVA']                          = df2['IVA'].astype(str)

        df2['Fecha_abono']                  = df2['Fecha_abono'].astype(str)
        df2['Cuenta_de_abono']              = df2['Cuenta_de_abono'].astype(str)
        df2['Local']                        = df2['Local'].astype(str)
        current_time = now.strftime("%H:%M:%S")
        print("Tiempo inicio insercion BD =", current_time)
        bd_tbk.insert_TRXTBK_credito_fastSql(df2)
        current_time_end = now.strftime("%H:%M:%S")
        print("Tiempo Fin insercion BD =", current_time_end)


def retry():
    minimo = 0
    maximo = 4
    while minimo <= maximo:
        try:
            func_descarga_archivos()
            break
        except:
            minimo += 1
        if minimo == maximo:
            send.envio_email("[RPA TBK TRX DUPLICADAS] Reintentos ", "Se llego al limite de reintentos : " + str(maximo))
            log = "[RPA TBK TRX DUPLICADAS] Reintentos ", "Se llego al limite de reintentos : " + str(maximo)
            logconfig.log_info(log)
            break

retry()