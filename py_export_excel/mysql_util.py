import mysql.connector
from mysql.connector import Error
import excel_util

connection = None
h = None
u = None
p = None
d = None

def connect(host, user, password, database):
    global connection, h, u, p, d
    h = host
    u = user
    p = password
    d = database

    try:
        connection = mysql.connector.connect(host=host, user=user, password=password)
        if connection.is_connected():
            print("Connecteed to MySQL Server")
            execute(f"CREATE DATABASE IF NOT EXISTS {database} DEFAULT CHARSET=utf8mb4;")
            execute(f"USE {database};")
            return True
        print("Connect to MySQL Server failed")
    except Error as e:
        print("Error while connecting to MySQL", e)
    return False

def is_connected():
    return connection.is_connected()

def reconnect():
    return connect(h, u, p, d)

def execute(command):
    # print("=====", command)
    if not connection.is_connected():
        print("connection is not connected, command not execute:\n", command)
    cursor = connection.cursor()
    cursor.execute(command)
    cursor.close()

    if not connection.is_connected():
        print(command, " >> cause connection close.")
        reconnect()
    return True

def dispose():
    if connection.is_connected():
        connection.close()
        print("MySQL connection is closed")
    else:
        print("MySQL connection is not connected")

def create_table(table_name, fields, types, names, comments):
    field_list = []
    for i in range(len(fields)):
        comment = comments[i]
        if comment is None:
            comment = ""
        field_list.append(f"`{fields[i]}` {types[i]} COMMENT '{names[i]}{comment}'")
    execute(f"DROP TABLE IF EXISTS {table_name};")
    execute(f"CREATE TABLE {table_name} ({','.join(field_list)}, PRIMARY KEY (`id`)) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;")

def delete_table(table_name):
    execute(f"DROP TABLE IF EXISTS {table_name};")

def rename_table(old_name, new_name):
    execute(f"DROP TABLE IF EXISTS {new_name};")
    execute(f"ALTER TABLE {old_name} RENAME TO {new_name};")

def record_to_str_list(record, types):
    return [excel_util.get_str_value_by_type(record[i], types[i]) for i in range(len(record))]

def insert_data(table_name, fields, types, records):
    fields = [f"`{str}`" for str in fields]
    fields_str = ", ".join(fields)
    
    record_str_list = []
    for record in records:
        new_record = record_to_str_list(record, types)
        record_str_list.append("({})".format(",".join(new_record)))
    values_str = ",\n".join(record_str_list)
    execute(f"INSERT INTO {table_name} ({fields_str}) VALUES {values_str};")

    # for i in range(0, len(record_str_list), 1000):
    #     end_num = i + 1000
    #     if end_num > len(record_str_list):
    #         end_num = len(record_str_list)
    #     values_str = ",\n".join(record_str_list[i:end_num])
    #     execute(f"INSERT INTO {table_name}\n ({fields_str})\n VALUES\n {values_str};")
