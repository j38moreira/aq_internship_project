DB_CONFIG = {
    'DRIVER': '{SQL Server}',
    'SERVER': 'LAPTOP-1DR8H4K3\SQLEXPRESS', #MUDAR NA APRESENTAÇÃO DO PROJETO -> NOME DO CONECTAR SQL SERVER
    'DATABASE': 'aquariodb',
    'TRUSTED_CONNECTION': 'yes',
}

TABLE_CONFIG = {
    'SCHEMA': 'dbo',
    'TABLE1': 'cliente', #tabela cliente
    'TABLE2': 'egoi_cliente', #tabela com os ids dos clientes pela parte da e-goi
    'TABLE3': 'egoi_tag', #tabela das tags que existem na egoi e outras adicionadas para adicionar à e-goi
    'TABLE4': 'egoi_cliente_tag', #tabela tags que estao ligadas ao cliente
    'TABLE5': 'campanha', #tabela campanha que existem na egoi e outras adicionadas para adicionar à e-goi
    'TABLE6': 'template', #tabela com os templates dos emails que existem na egoi
    'TABLE7': 'cliente_cliques', #tabela simulacao cliques dos clientes
}