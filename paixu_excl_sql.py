from posixpath import split
import xlrd
from collections import Counter

# 这个脚本实现的是通过excle表格的内容输出修改数据库的sql语句，修改条件是名字和年龄
# P.S.重复的名字将根据出现的先后顺序输出在最前面，修改数据库之后需要进入抽pad平台手动调试


# 打开excel表格
# data_excel = xlrd.open_workbook('D:/vscode/text/tool/tool.xls')
# data_excel = xlrd.open_workbook('./tool.xls')
data_excel = xlrd.open_workbook('./tool.xlsx')


# 获取所有sheet名称
names = data_excel.sheet_names()
# names为需要查看数据类型的变量
print(type(names))
print(names)
 
# 获取book中的sheet工作表的三种方法,返回一个xlrd.sheet.Sheet()对象
# table = data_excel.sheets()[0]     # 通过索引顺序获取sheet
# table = data_excel.sheet_by_index(sheetx=0)     # 通过索引顺序获取sheet
table = data_excel.sheet_by_name(sheet_name='Sheet1')    # 通过名称获取
print(table) 
 
# excel工作表的行列操作
# n_rows = table.nrows    # 获取该sheet中的有效行数
# n_cols = table.ncols    # 获取该sheet中的有效列数
# row_list = table.row(rowx=0)    # 返回某行中所有的单元格对象组成的列表
# cols_list = table.col(colx=0)    # 返回某列中所有的单元格对象组成的列表

    
# # 返回某行中所有单元格的数据组成的列表
# row_data=table.row_values(0,start_colx=0,end_colx=None)
 
 
# # 返回某列中所有单元格的数据组成的列表
cols_name=table.col_values(0,start_rowx=0,end_rowx=None)
# 获取第二列的内容
cols_numb=table.col_values(1,start_rowx=0,end_rowx=None)

cols_number = list(map(int, cols_numb[:]))
print(cols_name)
print(cols_number)
# row_lenth=table.row_len(0)   # 返回某行的有效单元格长度
b = dict(Counter(cols_name))
# 挑选出重复的名字形成list
c = [key for key,value in b.items()if value > 1]
print(type(c))
print (c)

a = 0;
# 循环输出重复的名字
for ic , j in zip(c,cols_name):
    # Python 判断列表cols_name(list)内是重复名字的位置
    d = [i for i,v in enumerate(cols_name) if v==ic]
    for xx in d:
      a = a+1;
    # 0占位占位到5位
      e = str(a).zfill(5)
      print('UPDATE a01 SET number = '+str(e)+' WHERE name = \''+str(j)+'\' and number = '+str(cols_number[xx]))
# 循环输出不重复的名字
for i , j in zip(cols_name,cols_number):
    # 判断重复的人名不输出
    if any(word if word in i else False for word in c) :
        continue;
    a = a+1;
    # 0占位占位到5位
    e = str(a).zfill(5)
    print('UPDATE a01 SET number = '+str(e)+' WHERE name = \''+str(i)+'\' and number = '+str(j))  
 
# # excel工作表的单元格操作
# row_col=table.cell(rowx=0,colx=0) # 返回单元格对象
# row_col_data=table.cell_value(rowx=0,colx=0) # 返回单元格中的数据