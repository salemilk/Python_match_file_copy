# -*- coding: utf-8 -*-
from excelutils import read_dict
import shutil

data_list = read_dict("Copy of Food Contact Audit List.xlsx", sheet_name="Sheet2")
file_list = []
for i in data_list:
    file_list.append(i['name'] + '.pdf')
print("需要查找的文件列表为:", file_list)
for name in file_list:
    try:
        shutil.copy("archive/{}".format(name), "copy")
        print("成功拷贝文件:{}".format(name))
    except FileNotFoundError as e:
        print("文件不存在：{}".format(name))
        continue
