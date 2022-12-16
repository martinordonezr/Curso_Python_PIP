# -*- coding: utf-8 -*-
"""
Created on Fri Jun 17 08:40:29 2022

@author: B39670
"""

import smtplib
from email.mime.multipart import MIMEMultipart
#from email.mime.base import MIMEBase
from email.mime.text import MIMEText
#from email import encoders
import pandas as pd

data = {'nombre' : ['Comisiones Diarias - TLV PRODUCTOS', 'Comisiones Diarias - BPE Tipo C', 'Comisiones Diarias - BPE Tipo B',  
                   'Comisiones Diarias - CONV', 'Comisiones Diarias - TLV CONV - BPE'], 
        'celda' : ['A1','A1','A1','A1','A1'],
        'hoja' : ['VENTA SV','Comisiones - EV BPE','Comisiones - SV BPE','Comisiones - SV Conv',
                  'SV Conv'],
        'destino' : ['C:\\Users/B39670/OneDrive - Interbank/1. Reportes/05. Team Performance FFVV/4. Televentas/1. Comisiones/3. TLV Productos/1. Detalle Jefe - Productos',
                     'C:\\Users/B39670/OneDrive - Interbank/1. Reportes/05. Team Performance FFVV/4. Televentas/1. Comisiones/2. TLV Convenios, BPE/3. Detalle Supervisor - BPE - Tipo C',
                     'C:\\Users/B39670/OneDrive - Interbank/1. Reportes/05. Team Performance FFVV/4. Televentas/1. Comisiones/2. TLV Convenios, BPE/4. Detalle Supervisor - BPE - Tipo B',
                     'C:\\Users/B39670/OneDrive - Interbank/1. Reportes/05. Team Performance FFVV/4. Televentas/1. Comisiones/2. TLV Convenios, BPE/2. Detalle Supervisor - Convenios',
                     'C:\\Users/B39670/OneDrive - Interbank/1. Reportes/05. Team Performance FFVV/4. Televentas/1. Comisiones/2. TLV Convenios, BPE/1. Detalle Jefe - BPE, Convenios']
        } 

df = pd.DataFrame(data)


## FILE TO SEND AND ITS PATH
#filename = 'some_file.csv'
#SourcePathName  = 'C:/reports/' + filename 

msg = MIMEMultipart()
msg['From'] = 'mordoez@intercorp.com.pe'
#msg['From'] = 'martinordonezr@outlook.com'
#recipients = ['john.doe@example.com', 'john.smith@example.co.uk']
#msg['To'] = ", ".join(recipients)
msg['To'] = 'mordoez@intercorp.com.pe'
msg['Subject'] = 'Report Update'
#body = 'Body of the message goes in here'

html_table = """<html>
<head>
<style>
table {  border-collapse: collapse;  width: autowidth;}
th {  text-align: center;  padding: 8px;}
td {  text-align: left;  padding: 8px;}
tr:nth-child(even) {background-color: #EDEDED; color:black; font-family:Arial; font-size:12; border:.5px solid black; padding:1; margin:0; border-collapse: collapse; padding:3;}
tr:nth-child(odd) {background-color: #DBDBDB; color:black; font-family:Arial; font-size:12; border:.5px solid black; padding:1; margin:0; border-collapse: collapse; padding:3;}
</style>
</head>
<body>
<table border="0.5" class="dataframe">
  <thead>""" + df.iloc[:,[0,1,2]].to_html(index=False).replace('<th>','<th style = "background-color: #006666; color:white; font-family:Arial; font-size:12; border:.5px solid silver;">')

body = '''
        <html></h1>Buenos días estimados,<br><br> 
        </h1>Se actualizaron los reportes a la fecha<br><br>
        ''' + html_table + '''
        <br><a href="https://interbankpe-my.sharepoint.com/personal/reportesretail_intercorp_com_pe/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Freportesretail%5Fintercorp%5Fcom%5Fpe%2FDocuments%2F1%2E%20Reportes%2F05%2E%20Team%20Performance%20FFVV%2F4%2E%20Televentas%2F1%2E%20Comisiones%2F2%2E%20TLV%20Convenios%2C%20BPE%2F3%2E%20Detalle%20Supervisor%20%2D%20BPE%20%2D%20Tipo%20C&FolderCTID=0x01200019956A29F4F1A9479771CB2C5988C476">BPE_C</a>
        </h1><br> Saludos</h1></html>
        '''
msg.attach(MIMEText(body, 'html'))
#msg.attach(MIMEText(body, 'plain'))

## ATTACHMENT PART OF THE CODE IS HERE
#attachment = open(SourcePathName, 'rb')
#part = MIMEBase('application', "octet-stream")
#part.set_payload((attachment).read())
#encoders.encode_base64(part)
#part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
#msg.attach(part)

server = smtplib.SMTP('smtp.office365.com', 587)  ### put your relevant SMTP here
server.ehlo()
server.starttls()
server.ehlo()
#server.login('martinordonezr@outlook.com', 'C@Toliquito123$')  ### if applicable#server.login('mordoez@intercorp.com.pe', 'G@laxy22')  ### if applicable
server.login('mordoez@intercorp.com.pe', 'G@laxy22')  ### if applicable
server.send_message(msg)
server.quit()