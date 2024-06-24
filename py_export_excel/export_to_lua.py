import sys
import re
import glob
import openpyxl
import excel_util
import traceback
from os import path
import time

def calc_default_value_dict(fields, records, types):
    field_count = len(records[0])
    field_name_to_default_value = {}
    for field_index in range(1, field_count):
        value_to_count = {}
        for row in range(len(records)):
            value = records[row][field_index]
            if value not in value_to_count:
                value_to_count[value] = 1
            else:
                value_to_count[value] = value_to_count[value] + 1
        
        default_value = None
        max_count = 9
        for value, count in value_to_count.items():
            if count > max_count:
                max_count = count
                default_value = value

        if max_count > 9:
            field_name = fields[field_index]
            field_name_to_default_value[field_name] = excel_util.get_str_value_by_type_for_lua(default_value, types[field_index])

    return field_name_to_default_value

def export_to_lua_file(file_path, export_dir):
    file_name = re.search(r"\\([^\\/]+?)\.\w+?$", path).group(1)
    export_file_path = export_dir + "\\" + file_name + ".lua"

    workbook_object = openpyxl.load_workbook(path, data_only=True)
    sheet_obj = workbook_object.active

    if excel_util.is_not_valid_table(sheet_obj):
        print(">>>>不是有效配置")
        return

    # 读取表头数据
    column_indices = excel_util.get_header_indices(sheet_obj, header_row=2, start_column=2)
    if len(column_indices) == 0:
        print(">>>>没有表头字段")
        return

    with open(export_file_path, 'w', encoding='utf-8', newline='') as f:
        names = excel_util.get_row_data(sheet_obj, 1, column_indices)
        names_comment = excel_util.get_row_comments(sheet_obj, 1, column_indices)
        types = excel_util.get_row_data(sheet_obj, 3, column_indices)
        fields = excel_util.get_row_data(sheet_obj, 2, column_indices)
        end_row = excel_util.get_end_row(sheet_obj)
        records = excel_util.get_records(sheet_obj, 5, end_row, column_indices)
        id_index = None
        for i in range(len(fields)):
            if fields[i] == 'id':
                id_index = i
                break
        
        if id_index is None:
            raise ValueError("id字段未找到")

        #计算每列的默认值
        field_to_default_value = calc_default_value_dict(fields, records, types)
    
        #写入字段表头信息
        name_and_comment_list = []
        for i in range(len(names)):
            if names_comment[i] == "":
                name_and_comment_list.append(f"【{fields[i]}】：{names[i]}")
            else:
                name_and_comment_list.append(f"【{fields[i]}】：{names[i]}\n批注：{names_comment[i]}")

        f.write("--[[\n表名：{}\n字段名：\n{}\n]]\n".format(file_name, "\n".join(name_and_comment_list)))

        #写入数据
        f.write(f"local Config_{file_name} = \n{{\n")
        row = 5
        for record in records:
            #控制列逻辑开始
            cell_obj = sheet_obj.cell(row, 1)
            #如果控制列第一列为空或者为END则输出 --[[PageName_]] 标识
            if row == 5:
                if cell_obj.value == None:
                    f.write(f"--[[PageName_]]\n")
                elif cell_obj.value == "END":
                    f.write(f"--[[PageName_END]]\n")
            #如果控制列有分页号则输出 --[[PageName_n]] 标识
            if cell_obj.value != None and cell_obj.value != "END":
                f.write(f"\n--[[PageName_{cell_obj.value}]]\n")
            row = row + 1
            #控制列逻辑结束

            key_value_str_list = [None if (fields[i] in field_to_default_value and excel_util.get_str_value_by_type_for_lua(record[i], types[i]) == field_to_default_value[fields[i]]) else f"{fields[i]}={excel_util.get_str_value_by_type_for_lua(record[i], types[i])}" for i in range(len(record))]
            key_value_str_list = [value for value in key_value_str_list if value != None]
            id = excel_util.get_str_value_by_type_for_lua(record[id_index], types[id_index])
            f.write("[{}]={{{},\n}};\n".format(id, ",\n".join(key_value_str_list)))
        f.write(f"\n--[[PageName_END]]\n}};")

        #写入元表信息
        if len(field_to_default_value.keys()) > 0:
            metatable_kvstr_list = [f"{field}={default_value}" for field, default_value in field_to_default_value.items()]
            f.write("\n\nlocal DefaultValueTable = \n{{\n{},\n}};\n\n".format(",\n".join(metatable_kvstr_list)))
            f.write("local base = {\n")
            f.write("\t__index = DefaultValueTable\n")
            f.write("}\n\n")
            f.write(f"for _, v in pairs(Config_{file_name}) do\n\tsetmetatable(v, base)\nend\n")

        #写入返回值
        f.write(f"\nreturn Config_{file_name}")
##################################################################################################
'''#测试用
files = ['G:/wind2/TC_branch/策划数值填表/wind2_config/EquipDungeon.xlsx']
export_dir = 'G:/wind2/TC_branch/LuaScripts/lua_source/config/configlogic'

export_to_lua_file(path, export_dir)'''

##################################################################################################
# main
folder = sys.argv[1]
export_dir = sys.argv[2]

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

for path in files:
    is_exclude = False
    if not is_exclude and ("Translate\\客户端" in path or "Translate\\策划" in path):
        is_exclude = True
    if is_exclude:
        print("忽略:", path)
        continue
    try:
        print("export:", path)
        export_to_lua_file(path, export_dir)
    except BaseException as e:
        traceback.print_exc()
        error_files.append((path, e))

print("\n\n==============导出完成==============\n\n")

if len(error_files) > 0:
    for path, error in error_files:
        print(f"{path}\t\t{error}")
    print("以上文件提交失败\n")

time.sleep(60)
