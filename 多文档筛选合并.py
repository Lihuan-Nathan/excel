import glob
import pdb
import os
import datetime
import xlrd
import numpy as np
import xlsxwriter
import pandas as pd

name_list = ['你要筛选的串1','你要筛选的串2','你要筛选的串3']

# 获取当前文件
dir_name = 'C:/Users/Administrator/Desktop/st_report/'
path_name = 'C:/Users/Administrator/Desktop/st_report/*.*'
path_files = glob.glob(pathname=path_name)
# 打印文件个数
print('文件数量:',len(path_files))

# 按照设置的类型解析为字符串
datestr = datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
# 设置文件的名称 按照项目名称加时间
suffix = ".xlsx"
filename = datestr + suffix
# 绝对输出的路径
abspath = os.path.join(dir_name, filename)
# 创建一个Workbook对象，这就相当于创建了一个Excel文件
book = xlsxwriter.Workbook(abspath)
# 创建sheet
sheet = book.add_worksheet()
# 新的excel从第0行开始写数据
row = 0
# 循环遍历文件
for sing_file in path_files:
    # 读取文件是第几个文件（从0开始）
    file_index = path_files.index(sing_file)
    # 获得文件的名字不加后缀
    filename = os.path.basename(sing_file)
    shotname,extend = os.path.splitext(filename)
    # 读取文件
    if file_index == 0:
        # 如果是第一个文件的话，要读取表头
        df= pd.read_excel(sing_file,header=None)
    else:
        # 不用读取表头
        df= pd.read_excel(sing_file)
    # dataframe转list
    values = np.array(df).tolist()
    pdb.set_trace()
    index = 4 #Campaign Name的下标（E）
    for value in values:
        msg_index = 0
        Campaign_Name = value[index]
        # 子串取前两个
        Sub_Campaign_Name = Campaign_Name[:]
        # 筛选
        if row == 0:
            for msg in value:
                if msg_index==0:
                    # 首行首列表头
                    sheet.write(row, msg_index,"Country")
                else:
                    sheet.write(row, msg_index,msg)
                msg_index += 1
            row = row + 1
        
        elif Sub_Campaign_Name in name_list:
            for msg in value:
                try:
                    # 开始时间换为文件名字
                    if msg_index == 0:
                        sheet.write(row,msg_index,shotname)
                    elif msg_index == 1:
                        time_str = msg.strftime('%Y-%m-%d')
                        sheet.write(row, msg_index,time_str)
                    else:
                        sheet.write(row, msg_index,msg)

                except:
                    sheet.write(row, msg_index,'')
                msg_index += 1
            row = row + 1
        else:
            # 跳出
            continue
   # 关闭文件        
book.close()
print('合并文档完成')