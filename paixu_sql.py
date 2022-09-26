from posixpath import split
import xlrd
from collections import Counter

# 这个脚本实现的是通过excle表格的内容输出修改数据库的sql语句，修改条件是名字和年龄
# P.S.重复的名字将根据出现的先后顺序输出在最前面，修改数据库之后需要进入抽pad平台手动调试


# 打开excel表格
data_excel1 = xlrd.open_workbook('./1.xls')
data_excel2 = xlrd.open_workbook('./2.xls')

# 获取所有sheet名称
names1 = data_excel1.sheet_names()
names2 = data_excel2.sheet_names()
# names为需要查看数据类型的变量
print(names1)
print(names2)
 

table1 = data_excel1.sheet_by_name(sheet_name='Sheet1')    # 通过名称获取
table2 = data_excel2.sheet_by_name(sheet_name='Sheet1')
print(table1) 
print(table2) 
 
 
# # 返回某列中所有单元格的数据组成的列表
cols_name1=table1.col_values(0,start_rowx=0,end_rowx=None)
# 获取第二列的内容
cols_job1=table1.col_values(1,start_rowx=0,end_rowx=None)
# # 返回某列中所有单元格的数据组成的列表
cols_job2=table2.col_values(0,start_rowx=0,end_rowx=None)
index=0
for a in cols_job2:
    for b , c in zip(cols_job1,cols_name1):
        if str(a) in str(b):
            index = index+1;
            # 0占位占位到5位
            e = str(index).zfill(5)
            print('UPDATE a01 SET number = '+str(e)+' WHERE name = \''+str(c)+'\' and number = '+str(b))
