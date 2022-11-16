# -*- coding: utf-8 -*-

import os
from datetime import datetime
import pandas as pd

download_path = "C:\\Users\\ASUS\\Downloads\\"
param_file = r"C:\Z_Proyectos\descargas_DIAN\Parametros.xlsx"

df_param = pd.read_excel(param_file, index_col=0)
start_date = df_param.loc['Fecha inicio','Valor'].strftime('%Y/%m/%d')
end_date = df_param.loc['Fecha inicio','Valor'].strftime('%Y/%m/%d')
zips_paths = df_param.loc['Ruta ZIPs','Valor']

params = {}
params['Execution'] = df_param.loc['Tipo de ejecución','Valor']
params['From_xlsx'] = df_param.loc['Descarga desde excel','Valor']
params['xlsx_path'] = df_param.loc['Ruta listado excel','Valor']
params['zips_path'] = zips_paths
params['dates_str'] = start_date + ' - ' + end_date
params['url'] = df_param.loc['URL DIAN','Valor']

url_received = r"https://catalogo-vpfe.dian.gov.co/Document/Received"
download_doc_url = r"https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFiles?trackId="

today = datetime.today().strftime('%Y.%m.%d.%H.%M')
dest_path = os.path.join(zips_paths, today)
dest_zip_path = os.path.join(dest_path, 'Archivos_zip')
contado_path = os.path.join(dest_path, 'Contado')
credito_path = os.path.join(dest_path, 'Credito')
error_path = os.path.join(dest_path, 'Error')
logs_path = os.path.join(dest_path, 'Logs')
folders = [dest_path, dest_zip_path, contado_path, credito_path, error_path, logs_path]

fieldsToSearch = {'Datos del Documento': {'Código Único de Factura - CUFE': [2, 3],
                                          'Número de Factura': 3, 'Forma de pago': 3,
                                          'Fecha de Emisión': 3, 'Medio de Pago': 3,
                                          'Fecha de Vencimiento': 3, 'Tipo de Operación': 3},
                  'Datos del Emisor': {'Razón Social': 3, 'Nombre Comercial': 3,
                                       'Nit del Emisor': 3, 'País': 2, 'Tipo de Contribuyente': 3,
                                       'Departamento': 3, 'Régimen Fiscal': 1, 'Municipio': 3,
                                       'Responsabilidad tributaria': 3, 'Dirección': 3,
                                       'Teléfono': 3, 'Correo': 3},
                  'Datos del Adquiriente': {'Razón Social': 3, 'Nombre Comercial': 3,
                                            'Nit del Emisor': 3, 'País': 3,
                                            'Tipo de Contribuyente': 3, 'Departamento': 3,
                                            'Régimen Fiscal': 3, 'Municipio': 3,
                                            'Responsabilidad tributaria': 3,
                                            'Dirección': 3, 'Teléfono': 3, 'Correo': 3}}