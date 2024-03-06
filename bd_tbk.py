import pyodbc

import pyodbc
from fast_to_sql import fast_to_sql

from datetime import datetime
import time

server = 'dc1-vm-02027'
database = 'T.I.'
username = 'saenex'
password = 'ENEX.2017'

now = datetime.now()

cnxn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';ENCRYPT=no;UID=' + username + ';PWD=' + password)

def insert_TRXTBK_debito_fastSql(df):
    fecha_carga     = now.strftime('%Y-%m-%d %H:%M:%S')
    archivo_data    = ""
    cursor = cnxn.cursor()    
    try:
        create_statement = fast_to_sql(df, "debito_New$_t", cnxn, if_exists="append",  temp=False)
    except Exception as e:
        cnxn.rollback()
        print("ERROR INSERT", e)
    finally:
        cnxn.commit()
        #cnxn.close()


def insert_TRXTBK_credito_fastSql(df):
    fecha_carga     = now.strftime('%Y-%m-%d %H:%M:%S')
    archivo_data    = ""
    cursor = cnxn.cursor()    
    try:
        create_statement = fast_to_sql(df, "credito_New$_t", cnxn, if_exists="append",  temp=False)
    except Exception as e:
        cnxn.rollback()
        print("ERROR INSERT", e)
    finally:
        cnxn.commit()
        cnxn.close()


def insert_TRXTBK_credito(df):
    fecha_carga     = now.strftime('%Y-%m-%d %H:%M:%S')
    archivo_data    = ""
    cursor = cnxn.cursor()
    try:
        for index, row in df.iterrows():
            valores = ( row.Tipo_trx, row.Fecha_venta, row.Tipo_Tarjeta, row.Identificador, row.Tipo_Cuota, row.Monto_original_venta,
                        row.Codigo_autorizacion_venta, row.Numero_cuota, row.Monto_para_abono, row.Comision_e_iva_comision, row.Comision_adicional_e_iva_comision_adicional,
                        row.Numero_boleta, row.Monto_anulacion, row.Devolucion_comision_e_iva_comision, row.Devolucion_comision_adicional_e_iva_comision, row.Monto_retencion,
                        row.Periodo_de_cobro, row.Motivo, row.Detalle_cobros_u_observacion, row.Monto, row.IVA, row.Fecha_abono,
                        row.Cuenta_de_abono, row.Local, archivo_data ,fecha_carga)
            cursor.execute('INSERT INTO dbo.credito_New$_t (Tipo_Transaccion,Fecha_Venta,Tipo_Tarjeta,Identificador,Tipo_Cuota,'
                +'monto_Original_Venta,C_digo_Autorizaci_n_Venta, N_Cuota, Monto_Para_Abono,Comisi_n_e_IVA_Comisi_n, Comisi_n_Adicional_e_IVA_Comisi_n_Adicional,'
                +'N_Boleta,Monto_Anulaci_n,Devoluci_n_Comisi_n_e_IVA_Comisi_n, Devoluci_n_Comisi_n_Adicional_e_IVA_Comisi_n, Monto_Retenci_n, Per_odo_de_Cobro, Motivo,'
                +'Detalle_de_cobros_u_observaci_n, Monto, IVA, Fecha_Abono, Cuenta_de_Abono, Local, archivo, Fecha_carga) '
                +'VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)', valores)
    except Exception as e:
        cursor.rollback()
        print("ERROR INSERT", e)
    finally:
        cursor.commit()
        cursor.close()

 


def insert_TRXTBK_debito(df):
    fecha_carga     = now.strftime('%Y-%m-%d %H:%M:%S')
    archivo_data    = ""
    cursor = cnxn.cursor()
    try:
        for index, row in df.iterrows():
            valores = ( row.Tipo_trx, row.Fecha_venta, row.Tipo_Tarjeta, row.Identificador, row.Tipo_Venta,
                        row.Codigo_autorizacion_venta, row.Numero_cuota, row.Monto_TRX, row.Monto_afecto, row.Comision_e_iva_comision,
                        row.Monto_exento, row.Numero_boleta , row.Monto_anulacion, row.Monto_retenido, row.Devolucion_comision, row.Monto_retencion,
                        row.Motivo, row.Periodo_de_cobro, row.Detalle_cobros_u_observacion, row.Monto, row.IVA, row.Fecha_abono,
                        row.Cuenta_de_abono, row.Local, row.Tipo_documento , archivo_data ,fecha_carga)
            cursor.execute('INSERT INTO dbo.debito_New$_t (Tipo_Transaccion,Fecha_Venta,Tipo_Tarjeta,Identificador,Tipo_Venta,'
                +'C_digo_Autorizaci_n_Venta, N_Cuota, monto_Transaccion,Monto_Afecto, Comisi_n_e_IVA_Comisi_n,'
                +'Monto_Exento,N_Boleta,Monto_Anulaci_n, Monto_Retenido, Devoluci_n_Comisi_n, Monto_Retenci_n, Motivo, Per_odo_de_Cobro,'
                +'Detalle_de_cobros_u_observaci_n, Monto, IVA, Fecha_Abono, Cuenta_de_Abono, Local, Tipo_Documento, archivo, Fecha_carga) '
                +'VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)', valores)
    except Exception as e:
        cursor.rollback()
        print("ERROR INSERT", e)
    finally:
        cursor.commit()
        cursor.close()
  
