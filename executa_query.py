# Fontes
# https://stackoverflow.com/questions/34614070/how-to-execute-query-saved-in-ms-access-using-pyodbc
# https://codereview.stackexchange.com/questions/79008/parse-a-config-file-and-add-to-command-line-arguments-using-argparse-in-python

import argparse
import yaml
import pyodbc

def execute():
    print(parse_args(create_parser()).db_name)
def create_parser():
    parser = argparse.ArgumentParser()

    g = parser.add_argument_group('Processa queries')
    g.add_argument(
        '--config-file',
        dest='config_file',
        type=argparse.FileType(mode='r'))
    return parser

def parse_args(parser):
    args = parser.parse_args()
    if args.config_file:
        data = yaml.load(args.config_file)
        delattr(args, 'config_file')
        arg_dict = args.__dict__
        for key, value in data.items():
            if isinstance(value, list):
                for v in value:
                    arg_dict[key].append(v)
            else:
                arg_dict[key] = value
    return args
def conectar(db_name):
    connStr = (
    "Driver={aaMicrosoft Access Driver (*.mdb, *.accdb)};"
    "DBQ=%s;" %db_name
    )
    connection = pyodbc.connect(connStr)
    return connection
def executar_queries(queries):
    sqls = ["""{CALL %s}""" % q for q in queries.split(',')]
    print(sqls)
    #Executa Todas queries
    for sql in sqls:
        connection.execute(sql)
    #crsr.close()
    connection.commit()
    connection.close()

# ler configurações
config = parse_args(create_parser())
try:
    connection = conectar(config.db_name)
    print('Conectou com sucesso')
except:
    print("erro na conexao")
    if 'Microsoft Access Driver (*.mdb, *.accdb)' not in [x for x in pyodbc.drivers()]:
        print('Instalar','https://www.microsoft.com/en-us/download/details.aspx?id=54920')

# Executar queries
try:
    executar_queries(config.queries)
    print('Executou queries com sucesso')
except:
    print("erro nas queries")