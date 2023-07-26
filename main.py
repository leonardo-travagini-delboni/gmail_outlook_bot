""" GMAIL & OUTLOOK MASS EMAIL SENDER USING PYTHON AND SQL """

############################################################# CREDITS ##############################################################
# GMAIL & OUTLOOK EMAIL SENDER USING PYTHON AND SQL
# Developed by Leonardo Travagini Delboni
# Version 1.0 - Since July 2023
#####################################################################################################################################

############################
# LIBRARIES AND IMPORTINGS #
############################

# Libraries and modules:
import pandas as pd
import time
import sys
import warnings
import os
import logging
import datetime
import setproctitle
import random
import yagmail
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Importings from common folder:
from functions import record, bot_telegram_sendtext, get_db, update_db
from settings import warning_telegram_id, sucess_telegram_id

# Email settings:
from settings import USER_GMAIL, PASSWORD_GMAIL, USER_OUTLOOK, PASSWORD_OUTLOOK

# Config Imports:
from config import STD_SUBJECT, STD_BODY, ATTACHMENTS, common_path

# Common folder settings:
sys.path.insert(0, common_path)

################################
# LOGGING AND WARNING SETTINGS #
################################

# Turning off the iloc and mixed types in the same column warnings:
warnings.filterwarnings('ignore', category=pd.errors.SettingWithCopyWarning)

# Turning off the chained assignment warnings:
pd.options.mode.chained_assignment = None

# Setting the logging file:
log_files = 'logging'
filename = os.path.basename(__file__)
nome_arquivo, extensao_arquivo = os.path.splitext(filename)

# Creating the logging folder if not exists:
if not os.path.exists(log_files):
    os.makedirs(log_files)

# Removing the default logging (avoidind future logging handlers warnings):
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

# Setting the logging format:
format_logging = "%(asctime)s - %(levelname)s - %(module)s - %(lineno)d - %(message)s"
logging.basicConfig(level = logging.INFO, filename = f'{log_files}/{nome_arquivo}.log', format = format_logging)

################################
# MAIN FUNCTION - MAIN PROGRAM #
################################

# Main function:
def main(wait_time, tbl, municipio, uf):

    try:
        # Extracing email IT companies database:
        record(f'Extraindo dados do banco de dados...', 'blue')
        df, status_get = get_db(tbl)
        df = df.loc[df['status'] == 'a enviar']
        df = df.drop_duplicates().reset_index(drop=True)

        # Filtering as per the wished city:
        if municipio != '':
            record(f'Filtrando empresas de {municipio}...', 'blue')
            df = df.loc[df['municipio'] == municipio]
        if uf != '':
            record(f'Filtrando empresas de {uf}...', 'blue')
            df = df.loc[df['uf'] == uf]
        df = df.reset_index(drop=True)
        
        # Checking if the database is empty:
        df = df.drop_duplicates(subset=['email'], keep='first')
        if df.empty:
            record(f'No data found in database. - {df}', 'red')
            return False
        record(f'Foram encontradas {len(df)} empresas a enviar para esses criterios.', 'blue')
        
        # Sending the emails in loop:
        max_number = len(df)
        contador = 1
        record(f'Iniciando o envio de e-mails para {len(df)} empresas...', 'green')

        # Initianting the loop:
        for index, row in df.iterrows():

            # Recording the progress:
            record(f'\nIniciando etapa {contador} de total de {max_number}...', 'blue')

            # Extracting data from the row:
            nome_fantasia = row['nome_fantasia']
            email = row['email']

            # Getting the email subject and body:
            email_subject = get_email_subject(nome_fantasia)
            email_body = get_email_body(nome_fantasia)

            # Getting new wait time:
            random_float_1 = random.uniform(0.5, 1.2)
            random_float_2 = random.uniform(0.5, 1.2)
            wait_time = wait_time * random_float_1 * random_float_2

            # Extracting the email as lower case:
            email_lower = email.lower()
            
            # Sending through GMAIL (the index is even):
            if index % 2 == 0:
                provedor = 'gmail'
                status = send_email_gmail(USER_GMAIL, PASSWORD_GMAIL, email_subject, email_body, ATTACHMENTS, email_lower)

            # Sending through OUTLOOK (the index is odd):
            else:
                provedor = 'outlook'
                status = send_email_outlook(USER_OUTLOOK, PASSWORD_OUTLOOK, email_subject, email_body, ATTACHMENTS, email_lower)
            
            # Updating the database:
            if status:
                record(f'E-mail enviado via {provedor} para {email} com sucesso!', 'blue')
                bot_telegram_sendtext(f'Sucesso via {provedor}: {email}', sucess_telegram_id)
            else:
                record(f'Erro ao enviar e-mail via {provedor} para {email}!', 'red')
                bot_telegram_sendtext(f'Falha via {provedor}: {email}', warning_telegram_id)
            
            # Updating row:
            record(f'Atualizando linha no banco de dados...', 'yellow')
            now = str( datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
            set_dict = {"timestamp":now, "status": f'enviado_{provedor}'}
            where_dict = {"municipio": municipio, "uf": uf, "email": email}
            status_update = update_db(tbl, set_dict, where_dict)

            # Reccording accordingly:
            if status_update:
                record(f'Linha atualizada com sucesso!', 'blue')
            else:
                record(f'Erro ao atualizar linha!', 'red')

            # Waiting until next e-mai:
            contador += 1
            record(f'Esperando {round(wait_time/60, 2)} minutos para enviar o próximo e-mail...', 'yellow')
            time.sleep(wait_time)
        
        # Concluding the process:
        record(f'Envio de e-mails concluido!', 'green')
        return True
    
    # Excepting errors:
    except Exception as e:
        record(f'Erro ao enviar e-mails: {e}', 'red')
        return False

#################################
# COMPONENT FUNCTIONS - PROGRAM #
#################################

# GMAIL - Single email sender function:
def send_email_gmail(user, password, subject, body, attachments, receiver_email):
    try:
        record(f'Enviando e-mail via GMAIL para {receiver_email}...')
        yag = yagmail.SMTP(user=user, password=password)
        yag.send(to=receiver_email, subject=subject, contents=body, attachments=attachments)
        record(f'E-mail enviado para {receiver_email} via GMAIL com sucesso!', 'green')
        return True
    except Exception as e:
        record(f'Erro ao enviar e-mail para {receiver_email} via GMAIL: {e}', 'red')
        return False

# OUTLOOK - Single email sender function:
def send_email_outlook(user, password, subject, body, attachments, receiver_email):

    # Initializing the MIMEMultipart object:
    msg = MIMEMultipart()
    msg['From'] = user
    msg['To'] = receiver_email
    msg['Subject'] = subject

    # Body:
    msg.attach(MIMEText(body, 'plain'))

    # Attachments:
    for file in attachments:
        with open(file, 'rb') as attachment_file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment_file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % file)
        msg.attach(part)

    # Conection to the Outlook SMTP server:
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()

    # Sending the email:
    try:
        # Sending the email:
        record(f'Enviando e-mail via OUTLOOK para {receiver_email}...')
        server.login(user, password)
        text = msg.as_string()
        server.sendmail(user, receiver_email, text)
        record(f'E-mail enviado para {receiver_email} via OUTLOOK com sucesso!', 'green')
        return True

    # Excepting errors:
    except Exception as e:
        record(f'Erro ao enviar e-mail para {receiver_email} via OUTLOOK: {e}', 'red')
        return False

    # Finally:
    finally:
        server.quit()

###############################
# INPUTS INSERTED BY THE USER #
###############################

# Getting the email subject:
def get_email_subject(nome_fantasia=''):
    if nome_fantasia != '' and nome_fantasia != None and nome_fantasia != 'nan' and nome_fantasia != 'NaN' and nome_fantasia != 'NAN' and nome_fantasia != '':
        return f'{STD_SUBJECT} - {nome_fantasia}'
    else:
        return STD_SUBJECT

# Getting the email body:
def get_email_body(nome_fantasia=''):
    if nome_fantasia != '' and nome_fantasia != None and nome_fantasia != 'nan' and nome_fantasia != 'NaN' and nome_fantasia != 'NAN'and nome_fantasia != '':
        vocativo = f'Prezado(a) Responsável por Contratações da {nome_fantasia},'
    else:
        vocativo = 'Prezado(a) Responsável por Contratações,'
    email_body = vocativo + STD_BODY
    return email_body

##################################
# EXECUTING FUNCTION - MAIN CODE #
##################################

# Executing the function:
if __name__ == '__main__':

    # Naming process for control using HTOP:
    setproctitle.setproctitle('monitor_bot_email_sender')

    # Initiating the email sender bot:
    status_main = main(wait_time=1800, tbl='tbl_ti_companies_emails', municipio='São Paulo', uf='SP')

    # Sending telegram warning accordingly:
    if status_main:
        record(f'Envio de e-mails concluido!', 'green')
        bot_telegram_sendtext(f'Envio de e-mails concluido!', warning_telegram_id)
    else:
        record(f'Erro ao enviar e-mails!', 'red')
        bot_telegram_sendtext(f'Erro ao enviar e-mails!', warning_telegram_id)
