import http.client
import json
from db_config import DB_CONFIG, TABLE_CONFIG
import pyodbc
from egoi_config import API_KEY, LIST_ID, BODY_EMAIL, CAMPAIGN_HASH
import logging
import time

egoi_conn = http.client.HTTPSConnection("api.egoiapp.com", timeout = 30)
duracao_sleep = 1

headers = {
    'Apikey': API_KEY,
    'Content-Type': 'application/json'
}

conn = pyodbc.connect(
    f"DRIVER={DB_CONFIG['DRIVER']};"
    f"SERVER={DB_CONFIG['SERVER']};"
    f"DATABASE={DB_CONFIG['DATABASE']};"
    f"TRUSTED_CONNECTION={DB_CONFIG['TRUSTED_CONNECTION']};"
)

cursor = conn.cursor()

logging.basicConfig(level=logging.INFO, filename='py_log.log', filemode='w', format='%(asctime)s %(levelname)s %(message)s')

def CriarContactosEgoi():
    cursor.execute(
        f"SELECT * FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] WHERE rgpd_cliente = 1"
    )
    rCliente = cursor.fetchall()

    cliente_adicionado = False
    
    for cliente in rCliente:
        cursor.execute(
            f"SELECT id_egoi_cliente FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE2']}] WHERE id_cliente = ?", (cliente.id_cliente,)
        )
        id_existente = cursor.fetchone()

        if not id_existente:
            contactos_payload = {
                "base": {
                    "status": "active",
                    "first_name": cliente.nome_cliente,
                    "last_name": cliente.apelido_cliente,
                    "birth_date": str(cliente.data_nasc),
                    "language": cliente.idioma_cliente,
                    "email": cliente.email_cliente,
                    "cellphone": "351-" + str(cliente.telemovel_cliente)
                }
            }

            # time.sleep(duracao_sleep)
            
            criarcont_endpoint = f"/lists/{LIST_ID}/contacts"
            egoi_conn.request("POST", criarcont_endpoint, json.dumps(contactos_payload), headers)
            criarcontacto_response = egoi_conn.getresponse()
            data = criarcontacto_response.read()

            if criarcontacto_response.status == 201:
                egoi_response_data = json.loads(data.decode('utf-8'))
                egoi_contact_id = egoi_response_data.get('contact_id')

                cursor.execute(
                    f"INSERT INTO [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE2']}] (contact_id, id_cliente) VALUES (?, ?)",
                    (egoi_contact_id, cliente.id_cliente)
                )
                conn.commit()
                logging.info(f"Contacto {cliente.email_cliente} criado com sucesso no sistema E-goi.")
                cliente_adicionado = True
                
            else:
                detalhe_erro = json.loads(data.read().decode('utf-8'))
                logging.error(f"Erro ao criar contacto {cliente.email_cliente}. Detalhes: {detalhe_erro}")
                print('Ver py_log.log | Status: ', criarcontacto_response.status) 
    
    return cliente_adicionado

def EsquecerContactosEgoi(id_cliente):
    cursor.execute(
        f"SELECT contact_id FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE2']}] WHERE id_cliente = ?", (id_cliente,)
    )
    id_esquecido = cursor.fetchone()
    
    cliente_esquecido = False    
    
    if id_esquecido:
        payload = {
            "contacts": [id_esquecido[0]]
        } 
        
        # time.sleep(duracao_sleep)
        
        forgetcont_endpoint = f"/lists/{LIST_ID}/contacts/actions/forget"
        egoi_conn.request("POST", forgetcont_endpoint, json.dumps(payload), headers)
        forgetcontacto_response = egoi_conn.getresponse()
        data = forgetcontacto_response.read()
        
        if forgetcontacto_response.status == 202:
            egoi_response_data = json.loads(data.decode('utf-8'))
            egoi_result = egoi_response_data.get('result')
            
            if egoi_result == 'success':
                cursor.execute(
                    f"DELETE FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE2']}] WHERE id_cliente = ?", (id_cliente,)
                )
                cursor.execute(
                    f"UPDATE [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] SET rgpd_cliente = '0' WHERE id_cliente = ?", (id_cliente,)
                )
                conn.commit()
                logging.info(f"Contacto {id_esquecido[0]} esquecido com sucesso.")
                cliente_esquecido = True
        else:
            detalhe_erro = json.loads(data.decode('utf-8'))
            logging.error(f"Erro ao atribuir uma tag a um cliente. Detalhes: {detalhe_erro}")
            print('Ver py_log.log | Status: ', forgetcontacto_response.status)  

def AtualizartagsDB():
    tag_adicionada = False
    
    payload = ''
    atuatag_endpoint = "/tags?order=asc&order_by=tag_id"
    
    # time.sleep(duracao_sleep)
    
    egoi_conn.request("GET", atuatag_endpoint, json.dumps(payload), headers)      
    atuatag_response = egoi_conn.getresponse()
    data = atuatag_response.read()
    
    if atuatag_response.status == 200:
        egoi_response_data = json.loads(data.decode('utf-8'))
        egoi_tags = egoi_response_data.get('items', [])

        for tag in egoi_tags:
            egoi_tag_id = tag.get('tag_id')
            egoi_nome_tag = tag.get('name').title().replace('_', ' ')
            egoi_cor_tag = tag.get('color')
            
            cursor.execute(
                f"SELECT 1 FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE3']}] WHERE tag_id = ?",
                (egoi_tag_id,)
            )
            tag_existente = cursor.fetchone()
            
            if not tag_existente:
                cursor.execute(
                    f"INSERT INTO [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE3']}] (tag_id, nome_tag, cor_tag) VALUES (?, ?, ?)",
                    (egoi_tag_id, egoi_nome_tag, egoi_cor_tag)
                )
                conn.commit()
                tag_adicionada = True
                logging.info(f"Tag {egoi_tag_id} | {egoi_nome_tag} adicionada à base de dados com sucesso")
            
    else:
        detalhe_erro = json.loads(data.decode('utf-8'))
        logging.error(f"Erro ao atribuir uma tag a um cliente. Detalhes: {detalhe_erro}")
        print('Ver py_log.log | Status: ', atuatag_response.status)
        
    return tag_adicionada

def AtualizarTagsEgoi():
    cursor.execute(
        f"SELECT * FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE3']}]"
    )
    rTag = cursor.fetchall()
    
    tag_adicionada = False
    
    for tag in rTag:
        if tag.tag_id == 'NULL':
            payload = {
                "name": tag.nome_tag,
                "color": tag.cor_tag
            }
            
            # time.sleep(duracao_sleep)
            
            criartag_endpoint = "/tags"
            egoi_conn.request("POST", criartag_endpoint, json.dumps(payload), headers)
            criartag_response = egoi_conn.getresponse()
            data = criartag_response.read()
            
            if criartag_response.status == 201:
                egoi_response_data = json.loads(data.decode('utf-8'))
                egoi_tag_id = egoi_response_data.get('tag_id')
                
                cursor.execute(
                    f"UPDATE [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE3']}] SET tag_id = ? WHERE nome_tag = ?",
                    (egoi_tag_id, tag.nome_tag)
                )
                conn.commit()
                tag_adicionada = True

            else:
                detalhe_erro = json.loads(data.decode('utf-8'))
                logging.error(f"Erro ao atribuir uma tag a um cliente. Detalhes: {detalhe_erro}")
                print('Ver py_log.log | Status: ', criartag_response.status)
                
    return tag_adicionada

def AttachTagContacto():
    attach_tag = False
    cursor.execute(
        f"""
        SELECT
            ec.id_egoi_cliente,
            ec.contact_id,
            et.tag_id,
            ect.id_egoi_tag
        FROM
            [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE2']}] AS ec
            JOIN [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE4']}] AS ect
                ON ec.id_egoi_cliente = ect.id_egoi_cliente
            JOIN [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE3']}] AS et
                ON ect.id_egoi_tag = et.id_egoi_tag
        WHERE
            ect.cliente_tagged = 0;
        """
    )
    tagCliente = cursor.fetchall()
    
    for tagClientes in tagCliente:
        payload = {
            "contacts": [tagClientes[1]],
            "tag_id": tagClientes[2] 
        }
        
        # time.sleep(duracao_sleep)
        
        attagcliente_endpoint = f"/lists/{LIST_ID}/contacts/actions/attach-tag"
        egoi_conn.request("POST", attagcliente_endpoint, json.dumps(payload), headers)
        attagcliente_response = egoi_conn.getresponse()
        data = attagcliente_response.read()

        if attagcliente_response.status == 202:
            cursor.execute(
                f"""
                UPDATE [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE4']}]
                SET cliente_tagged = 1
                WHERE id_egoi_cliente = {tagClientes[0]} and id_egoi_tag = {tagClientes[3]};
                """
            )
            conn.commit()
            attach_tag = True
            
        else:
            detalhe_erro = json.loads(data.decode('utf-8'))
            logging.error(f"Erro ao atribuir uma tag a um cliente. Detalhes: {detalhe_erro}")
            print('Ver py_log.log | Status: ', attagcliente_response.status)

    return attach_tag

def DetachTagContacto(id_cliente_egoi, tag_id):
    detach_tag = True
    cursor.execute(
        f"SELECT contact_id FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE2']}] WHERE id_egoi_cliente = ?",
        (id_cliente_egoi,)
    )
    contact_ids = [row[0] for row in cursor.fetchall()]

    payload = {
        "contacts": contact_ids,
        "tag_id": tag_id
    }
    
    # time.sleep(duracao_sleep)

    dettagcliente_endpoint = f"/lists/{LIST_ID}/contacts/actions/detach-tag"
    egoi_conn.request("POST", dettagcliente_endpoint, json.dumps(payload), headers)
    dettagcliente_response = egoi_conn.getresponse()
    data = dettagcliente_response.read().decode('utf-8')

    if not dettagcliente_response.status == 202:
        detalhe_erro = json.loads(data)
        logging.error(f"Erro ao remover uma tag de um cliente. Detalhes: {detalhe_erro}")
        print('Ver py_log.log | Status: ', dettagcliente_response.status)
        detach_tag = False
    
    return detach_tag

def EnviarEmail(resEmail_ee):
    resp = resEmail_ee[0][0]
     
    cursor.execute(
        f"""
        SELECT
            ec.contact_id, 
            c.email_cliente
        FROM
            [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] AS c
            JOIN [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE2']}] AS ec
                ON ec.id_cliente = c.id_cliente
        WHERE
            c.id_cliente = ? and c.rgpd_cliente = '1';
        """,
        (resp,)
    )
    res = cursor.fetchone()
    
    email_enviado = False    
    
    if res:
        payload = {
            "list_id": 21,
            "segments": {
                "type": "contact",
                "data": res[1]	
                },
            "unique_contacts_only": False,
            "limit_hour": {
                "hour_start": "09:30",
                "hour_end": "10:00"
            }
        }
        
        # time.sleep(duracao_sleep)
        
        enviaremail_endpoint = f"/campaigns/email/{CAMPAIGN_HASH}/actions/send"
        egoi_conn.request("POST", enviaremail_endpoint, json.dumps(payload), headers)
        enviaremail_response = egoi_conn.getresponse()
        data = enviaremail_response.read()
        
        if enviaremail_response.status == 200 or enviaremail_response.status == 201:
            logging.info(f"Email enviado com sucesso.")
            email_enviado = True
        else:
            try:
                detalhe_erro = json.loads(data.decode('utf-8'))
                logging.error(f"Erro ao enviar email. Detalhes: {detalhe_erro}")
            except json.JSONDecodeError:
                logging.error(f"Erro ao decodificar JSON na resposta do servidor.")
            
            print('Ver py_log.log | Status: ', enviaremail_response.status)

    return email_enviado

       
def ObterCampanhaBD():
    campanha_adicionada = False
    
    payload = ''
    obtercampanha_endpoint = f"/campaigns?channel=email&created_min=2024-02-01%2000%3A00%3A00&created_max=2024-08-08%2000%3A00%3A00&updated_min=2024-02-01%2000%3A00%3A00&updated_max=2024-08-08%2000%3A00%3A00&start_date_min=2024-02-01%2000%3A00%3A00&start_date_max=2024-08-08%2000%3A00%3A00&end_date_min=2024-02-01%2000%3A00%3A00&end_date_max=2024-08-08%2000%3A00%3A00&schedule_date_min=2024-02-01%2000%3A00%3A00&schedule_date_max=2024-08-08%2000%3A00%3A00&limit=30&order=asc&order_by=created"
    
    # time.sleep(duracao_sleep)
    
    egoi_conn.request("GET", obtercampanha_endpoint, json.dumps(payload), headers)
    obtercampanha_response = egoi_conn.getresponse()
    data = obtercampanha_response.read()
    
    if obtercampanha_response.status == 200:
        egoi_response_data = json.loads(data.decode('utf-8'))
        egoi_campanhas = egoi_response_data.get('items', [])
        
        for campanha in egoi_campanhas:
            egoi_campaign_hash = campanha.get('campaign_hash')
            egoi_campaigns_channel = campanha.get('channel')
            egoi_campaigns_internal_name = campanha.get('internal_name')
            egoi_campaigns_list_id = campanha.get('list_id')
            egoi_campaigns_title = campanha.get('title')
            

            cursor.execute(
                f"SELECT 1 FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE5']}] WHERE internal_name = ?",
                (egoi_campaigns_internal_name,)
            )
            
            campanha_existente = cursor.fetchone()
        
            if not campanha_existente:
                cursor.execute(
                    f"INSERT INTO [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE5']}] (campaign_hash, channel, internal_name, list_id, title) VALUES (?, ?, ?, ?, ?)",
                        (egoi_campaign_hash, egoi_campaigns_channel, egoi_campaigns_internal_name, egoi_campaigns_list_id, egoi_campaigns_title,)    
                )
                conn.commit()
                campanha_adicionada = True
                logging.info(f"Campanha {egoi_campaign_hash} | {egoi_campaigns_internal_name} adicionada à base de dados com sucesso")
            
    else:
        detalhe_erro = json.loads(data.decode('utf-8'))
        logging.error(f"Erro ao obter campanhas. Detalhes: {detalhe_erro}")
        print('Ver py_log.log | Status: ', obtercampanha_response.status)
        
    return campanha_adicionada

def CriarCampanhaEgoi():
    campanha_criada = False
    
    cursor.execute(
        f"SELECT internal_name FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE5']}] WHERE campaign_hash IS NULL;"
    )
    resCamp = cursor.fetchall()
    
    payload = {
            "list_id": LIST_ID,
            "internal_name": resCamp[0][0],
            "content": {
                "type": "html",
                "body": BODY_EMAIL,
            },
            "sender_id": 2,
    }
    
    # time.sleep(duracao_sleep)
    
    criarcamp_endpoint = "/campaigns/email"
    egoi_conn.request("POST", criarcamp_endpoint, json.dumps(payload), headers)
    criarcamp_response = egoi_conn.getresponse()
    data = criarcamp_response.read()
    
    if criarcamp_response.status == 201:
        egoi_response_data = json.loads(data.decode('utf-8'))
        egoi_chash = egoi_response_data.get('campaign_hash')
        cursor.execute(
            f"""
            UPDATE [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE5']}]
            SET campaign_hash = '{egoi_chash}'
            WHERE internal_name = '{resCamp[0][0]}';
            """
        )

        conn.commit()   
        logging.info(f"Campanha criada com sucesso no sistema E-goi.")
        campanha_criada = True
        
    else:
        detalhe_erro = json.loads(data.read().decode('utf-8'))
        logging.error(f"Erro ao criar campanha. Detalhes: {detalhe_erro}")
        print('Ver py_log.log | Status: ', criarcamp_response.status) 
        
    return campanha_criada

def AtualizarClienteEgoi(contact_id, id_cliente, nomeEdit, apelidoEdit, emailEdit, telemovelEdit):
    payload = {
        "base": {
            "status": "active",
            "first_name": nomeEdit,
            "last_name": apelidoEdit,
            "email": emailEdit,
            "cellphone": "351-" + str(telemovelEdit)
        }
    }
    
    atualizarcont_endpoint = f"/lists/{LIST_ID}/contacts/{contact_id}"
    egoi_conn.request("PATCH", atualizarcont_endpoint, json.dumps(payload), headers)
    atualizarcontacto_response = egoi_conn.getresponse()
    data = atualizarcontacto_response.read()
    
    if atualizarcontacto_response.status == 200:
        logging.info(f"Contacto {emailEdit} atualizado com sucesso no sistema E-goi.")
        
    else:
        detalhe_erro = json.loads(data.decode('utf-8'))
        logging.error(f"Erro ao atualizar contacto com o id {id_cliente}. Detalhes: {detalhe_erro}")
        print('Ver py_log.log | Status: ', atualizarcontacto_response.status) 
