import openpyxl
import excel_util
import time
import mysql_util as db
import sys
import glob
import re
from os import path

def upload(path, db):
    workbook_object = openpyxl.load_workbook(path, data_only=True)
    sheet_obj = workbook_object.active

    if excel_util.is_not_valid_table(sheet_obj):
        print(">>>>不是有效配置")
        return

    # 读取表头数据
    column_indices = excel_util.get_header_indices(sheet_obj, header_row=4, start_column=2)
    if len(column_indices) <= 1:
        print(">>>>没有表头字段")
        return

    types = excel_util.get_row_data(sheet_obj, 3, column_indices)
    fields = excel_util.get_row_data(sheet_obj, 4, column_indices)
    names = excel_util.get_row_data(sheet_obj, 1, column_indices)
    comments = excel_util.get_row_comments(sheet_obj, 1, column_indices)

    # 配表类型转为数据库类型
    sql_types = excel_util.get_sql_types(types)

    try:
        # 创建临时表
        excel_file_name = re.search(r"\\([^\\/]+?)\.\w+?$", path).group(1)
        tmp_table_name = f"__{int(time.time())}_{excel_file_name}"
        db.create_table(tmp_table_name, fields, sql_types, names, comments)

        # 插入数据
        end_row = excel_util.get_end_row(sheet_obj)
        records = excel_util.get_records(sheet_obj, start_row=5, end_row=end_row, column_indices=column_indices)
        db.insert_data(tmp_table_name, fields, types, records)

        # 重命名表为正式名字
        db.rename_table(tmp_table_name, excel_file_name)
    finally:
        if not db.is_connected():
            if db.reconnect():
                db.delete_table(tmp_table_name)
            else:
                print("临时表删除失败:", tmp_table_name)
        else:
            db.delete_table(tmp_table_name)
############################################################################
"""# 单个文件测试用
host   = '192.168.0.239'
user   = 'plan'
psw    = '123456'
db_name= 'wind_basetw_config'

files = ['G:\\wind2\\client_branch\\策划数值填表\\wind2_config\\DiaryNote.xlsx']
if db.connect(host, user, psw, db_name):
    for path in files:
        try:
            print("submit:", path)
            upload(path, db)
        except BaseException as e:
            print("except:", path, "\n", e)
            
print('提交完成')"""
############################################################################
#主逻辑开始
host   = sys.argv[1]
user   = sys.argv[2]
psw    = sys.argv[3]
db_name= sys.argv[4]
folder = sys.argv[5]

# 检查参数
if not path.exists(folder):
    raise ValueError("path not exist : " + folder)

files = None
if path.isfile(folder):
    files = [folder]
elif path.isdir(folder):
    files = [file for file in glob.glob(folder + "/**/*", recursive=True) if re.match(r"[^~]+\.xls[xm]?$", file)]

    is_duplicate = False
    names = {}
    for path in files:
        excel_file_name = re.search(r"\\([^\\/]+?)\.\w+?$", path).group(1)
        if names.get(excel_file_name) is not None:
            print("重复文件 :" + excel_file_name)
            is_duplicate = True
        names[excel_file_name] = True
    if is_duplicate:
        raise ValueError("有重复文件")
else:
    raise ValueError("arg1 is not path or file :" + folder)

error_files = []

exclude_files = ["\\SceneAirWall.xlsm", "\\SceneEjector.xlsm", "\\SceneMonster.xlsm", "\\SceneNpc.xlsm", "\\SceneTrap.xlsm", "\\DungeonLevel.xlsm",
                 "\\SceneAirWall.xlsx", "\\SceneEjector.xlsx", "\\SceneMonster.xlsx", "\\SceneNpc.xlsx", "\\SceneTrap.xlsx", "\\DungeonLevel.xlsx", ]

if db.connect(host, user, psw, db_name):
    for path in files:
        is_exclude = False
        for exclude_file in exclude_files:
            if path.endswith(exclude_file):
                is_exclude = True
                break
        if not is_exclude and ("Translate\\客户端" in path or "Translate\\策划" in path):
            is_exclude = True
        if is_exclude:
            print("忽略:", path)
            continue

        try:
            print("submit:", path)
            upload(path, db)
        except BaseException as e:
            print("except:", path, "\n", e)
            error_files.append(path)

db.dispose()

print("\n\n==============提交完成==============\n\n")

if len(error_files) > 0:
    print("\n".join(error_files))
    print("以上文件提交失败\n")
