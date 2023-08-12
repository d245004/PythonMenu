import oracledb
import os

un = os.environ.get('newhaimsweb')
pw = os.environ.get('newhaims')
cs = os.environ.get('Autonet_03-PC:1521/orcl')

with oracledb.connect(user=un, password=pw, dsn=cs) as connection:
    with connection.cursor() as cursor:
        sql = """select sysdate from dual"""
        for r in cursor.execute(sql):
            print(r)