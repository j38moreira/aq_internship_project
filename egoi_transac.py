import http.client
import json
import pyodbc
from db_config import DB_CONFIG, TABLE_CONFIG
from egoi_config import API_KEY, BODY_EMAIL
import logging
import time

egoitransac_conn = http.client.HTTPSConnection("slingshot.egoiapp.com", timeout=30)
duracao_sleep = 1

headers = {
    'Apikey': API_KEY,
    'Content-Type': 'application/json',
}

conn = pyodbc.connect(
    f"DRIVER={DB_CONFIG['DRIVER']};"
    f"SERVER={DB_CONFIG['SERVER']};"
    f"DATABASE={DB_CONFIG['DATABASE']};"
    f"TRUSTED_CONNECTION={DB_CONFIG['TRUSTED_CONNECTION']};"
)

cursor = conn.cursor()

def ObterTemplate():
    template_obtido = False
    payload = ''
    
    time.sleep(duracao_sleep)
    
    obtemplate_endpoint = "/api/v2/email/templates"
    egoitransac_conn.request("GET", obtemplate_endpoint, payload, headers)
    obtemplate_response = egoitransac_conn.getresponse()
    data = obtemplate_response.read()
    
    if obtemplate_response.status == 200:
        egoitransac_data = json.loads(data.decode('utf-8'))
        
        for template in egoitransac_data:
            e_templateName = template.get('templateName')
            e_id_template = template.get('id')
            e_subject = template.get('subject')

            cursor.execute(
                f"SELECT 1 FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE6']}] WHERE id_template = ?",
                (e_id_template,)
            )
            
            template_existente = cursor.fetchone()
            
            if not template_existente:
                
                cursor.execute(
                    f"INSERT INTO [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE6']}] (id_template, templateName, subjectEmail) VALUES (?, ?, ?)",
                        (e_id_template, e_templateName, e_subject,)    
                )
                conn.commit()
                template_obtido = True
                logging.info(f"Template {e_templateName} | {e_subject} adicionado à base de dados com sucesso")
    
    else:
        detalhe_erro = json.loads(data.decode('utf-8'))
        logging.error(f"Erro ao obter template. Detalhes: {detalhe_erro}")
        print('Ver py_log.log | Status: ', obtemplate_response.status)            
        
    return template_obtido

def CriarTemplate():
    template_criado = False
    
    cursor.execute(
        f"SELECT subjectEmail, templateName FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE6']}] WHERE id_template IS NULL;"
    )
    resCamp = cursor.fetchall()
    
    payload = {
        "subject": resCamp[0][0],
        "htmlBody": BODY_EMAIL,
        "templateName": resCamp[0][1]
    }
    
    time.sleep(duracao_sleep)
    
    criartemp_endpoint = "/api/v2/email/templates"
    egoitransac_conn.request("POST", criartemp_endpoint, json.dumps(payload), headers)
    criartemp_response = egoitransac_conn.getresponse()
    data = criartemp_response.read()
    
    if criartemp_response.status == 201:
        egoitransac_data = json.loads(data.decode('utf-8'))
        e_tempId = egoitransac_data.get('templateId')
        cursor.execute(
            f"""
            UPDATE [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE6']}]
            SET id_template = '{e_tempId}'
            WHERE templateName = '{resCamp[0][1]}';
            """
        )

        conn.commit()   
        logging.info(f"Template criado com sucesso no sistema E-goi.")
        template_criado = True
        
    else:
        detalhe_erro = json.loads(data.decode('utf-8'))
        logging.error(f"Erro ao criar template. Detalhes: {detalhe_erro}")
        print('Ver py_log.log | Status: ', criartemp_response.status)   
        
    return template_criado

def EnviarEmailTemplate(resEmail_ee):
    resp = resEmail_ee[0][0]
    
    cursor.execute(
        f"SELECT subjectEmail FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE6']}] WHERE id_template = 1;"
    )
    resTemp = cursor.fetchall()
    resT = resTemp[0][0]
    enviar_email_temp = False
    
    payload = {
        "senderId": "2",
        "subject": resT,
        "to": resp,
        "templateId": "1",
        "scheduleTo": "2024-03-14 13:20:00",
    }
    
    time.sleep(duracao_sleep)
    
    envtemplate_endpoint = "/api/v2/email/messages/action/send/single"
    egoitransac_conn.request("POST", envtemplate_endpoint, json.dumps(payload), headers)
    envtemplate_response = egoitransac_conn.getresponse()
    data = envtemplate_response.read()
    
    if envtemplate_response.status == 200:
        logging.info(f"Email enviado com sucesso.")
        print(f"Email enviado com sucesso.")
        enviar_email_temp = True
        
    else:
        try:
            detalhe_erro = json.loads(data.decode('utf-8'))
            logging.error(f"Erro ao enviar email. Detalhes: {detalhe_erro}")
            print(f"Erro ao enviar email. Detalhes: {detalhe_erro}")
        except json.JSONDecodeError:
            logging.error(f"Erro ao decodificar JSON na resposta do servidor.")
            print('Ver py_log.log | Status: ', envtemplate_response.status)
    
    return enviar_email_temp
        
def ValidarEmail(email):
    vEmail = False
    payload = {
        "email": email,
        "isMx": True,
        "isTemporaryEmail": True
    }
    
    time.sleep(duracao_sleep)
    
    valemail_endpoint = '/api/v2/utilities/emailValidator'
    egoitransac_conn.request("POST", valemail_endpoint, json.dumps(payload), headers)
    valemail_response = egoitransac_conn.getresponse()
    data = valemail_response.read()
    
    if valemail_response.status == 200:
        egoitransac_data = json.loads(data.decode('utf-8'))
        e_emailvalid = egoitransac_data.get('emailValid')
        e_emailMx = egoitransac_data.get('mx')
        e_emailtemp = egoitransac_data.get('emailTemporary')

        
        if e_emailvalid == 'true' and e_emailMx == 'true' and e_emailtemp == 'false':
            vEmail = True
        else:
            vEmail = False

    else:
        try:
            detalhe_erro = json.loads(data.decode('utf-8'))
            logging.error(f"Erro ao enviar email. Detalhes: {detalhe_erro}")
            print(f"Erro ao enviar email. Detalhes: {detalhe_erro}")
        except json.JSONDecodeError:
            logging.error(f"Erro ao decodificar JSON na resposta do servidor.")
            print('Ver py_log.log | Status: ', valemail_response.status)
    
    return vEmail