# -*- coding: utf-8 -*-

import os
from datetime import datetime

param_file = r"C:\Z_Proyectos\descargas_DIAN\Parametros.xlsx" ##CAMBIAR
doc_list_path = r"C:\Z_Proyectos\descargas_DIAN\Documentos.xlsx" ##CAMBIAR
download_path = "C:\\Users\\ASUS\\Downloads\\" ##CAMBIAR
zips_paths = r"Y:\02. CONTABILIDAD\1 CONTADOR\2 CLAUDIA  PATRICIA GOMEZ SALAZAR\CLAUDIA FACT ELECTRONICAS 2022\RADIAN\10. 31 DE OCTUBRE DEL 2022"

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