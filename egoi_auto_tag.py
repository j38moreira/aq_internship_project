import pyodbc
from db_config import DB_CONFIG, TABLE_CONFIG
import logging
import time
import egoi
import datetime

logging.basicConfig(level=logging.INFO, filename = 'py_log.log', filemode = 'w', format = '%(asctime)s %(levelname)s %(message)s')

duracao_sleep = 1

conn = pyodbc.connect(
    f"DRIVER={DB_CONFIG['DRIVER']};"
    f"SERVER={DB_CONFIG['SERVER']};"
    f"DATABASE={DB_CONFIG['DATABASE']};"
    f"TRUSTED_CONNECTION={DB_CONFIG['TRUSTED_CONNECTION']};"
)

cursor = conn.cursor()

id_tag = {
        'Informatica': 29,
        'Smartphones': 30,
        'Televisores Recetores': 31,
        'Electrodomesticos': 32,
        'Electronica': 33,
        'Electrico Iluminacao': 34,
        'Outras': 35
    }

def AutoTag():
    try:
        current_date = datetime.date.today()
        
        cursor.execute(
            f"SELECT MAX(data_cliques) FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE7']}] WHERE data_cliques <= '{current_date}'"
        )

        dataRes = cursor.fetchall()
        
        if dataRes[0][0] is not None:
            relevant_date = dataRes[0][0]
        else:
            relevant_date = current_date
                                
        cursor.execute(
                f"SELECT ec.id_egoi_cliente, nc.categoria_interagida "
                f"FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE2']}] AS ec "
                f"JOIN [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] AS c ON ec.id_cliente = c.id_cliente "
                f"JOIN [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE7']}] AS nc ON nc.id_cliente = c.id_cliente "
                f"WHERE c.rgpd_cliente = 1 AND nc.num_cliques > 10"
                f"AND nc.data_cliques = '{relevant_date}'"
        )
        restag = cursor.fetchall()

        for id_egoi_cliente_tag, categoria_tag in restag:
            id_egoi_tag = id_tag.get(categoria_tag)
            
            if id_egoi_tag:
                cursor.execute(
                    f"SELECT * FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE4']}] WHERE id_egoi_cliente = ? AND id_egoi_tag = ?",
                    (id_egoi_cliente_tag, id_egoi_tag,)
                )
                existe = cursor.fetchone()

                if not existe:
                    cursor.execute(
                        f"INSERT INTO [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE4']}] (id_egoi_cliente, id_egoi_tag) VALUES (?, ?)",
                        (id_egoi_cliente_tag, id_egoi_tag,)
                    )
                    conn.commit()
                    time.sleep(duracao_sleep)
                    egoi.AttachTagContacto()
        
        cursor.execute(
            f"SELECT ec.id_egoi_cliente, nc.categoria_interagida "
            f"FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE2']}] AS ec "
            f"JOIN [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE1']}] AS c ON ec.id_cliente = c.id_cliente "
            f"JOIN [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE7']}] AS nc ON nc.id_cliente = c.id_cliente "
            f"WHERE c.rgpd_cliente = 1 AND nc.num_cliques < 10"
            f"AND nc.data_cliques = '{relevant_date}'"
        )
        res = cursor.fetchall()     
        
        for id_egoi_cliente, categoria in res:
            id_egoi_tag = id_tag.get(categoria)
            
            if id_egoi_tag:
                cursor.execute(
                    f"SELECT * FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE4']}] WHERE id_egoi_cliente = ? AND id_egoi_tag = ? AND cliente_tagged = 1",
                    (id_egoi_cliente, id_egoi_tag,)
                )
                existe = cursor.fetchone()
                
                if existe is not None and existe:
                    cursor.execute(
                        f"SELECT et.tag_id FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE4']}] AS ct JOIN [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE3']}] AS et ON et.id_egoi_tag = ct.id_egoi_tag WHERE id_egoi_cliente = ? AND et.id_egoi_tag = ?",
                        (id_egoi_cliente, id_egoi_tag)
                    )
                    res = cursor.fetchone()
                    
                    time.sleep(duracao_sleep)
                    egoi.DetachTagContacto(id_egoi_cliente, res[0])
                    
                    cursor.execute(
                        f"DELETE FROM [{DB_CONFIG['DATABASE']}].[{TABLE_CONFIG['SCHEMA']}].[{TABLE_CONFIG['TABLE4']}] WHERE id_egoi_cliente = ? AND id_egoi_tag = ?",
                        (id_egoi_cliente, id_egoi_tag)
                    )
                    conn.commit()
                                
    except Exception as e:
        logging.error('Erro', exc_info=True)
        print('Ver py_log.log')