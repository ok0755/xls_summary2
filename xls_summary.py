#coding=utf-8
#author:Zhoubin
import os
import csv
import xlrd
import arrow
from multiprocessing import Pool,Process,freeze_support,cpu_count
""" 程序功能：获取本程序当前目录下所有excel文档，文档内容汇总至同一表格results.csv """
def get_xlsfiles():
    """ 获取当前文件夹下excel文件 """
    current_dir = os.getcwd()
    for files in os.listdir(current_dir):
        path = os.path.join(current_dir,files)
        if os.path.splitext(files)[1]=='.xls' or os.path.splitext(files)[1]=='.xlsx':
            yield path

def get_result(xls):
    """ 获取excel内容 """
    print 'Collecting:>> ',os.path.basename(xls).split('.')[0]
    th = []
    xls_path = xlrd.open_workbook(xls,encoding_override="utf-8",formatting_info=False)
    for sh in xls_path.sheet_names():                                       ## 遍历工作表
        sh_object = xls_path.sheet_by_name(sh)
        rows = sh_object.nrows
        for k in range(0,rows):
            ar = sh_object.row(k)
            ar_ctype = [a.ctype for a in ar]
            ar_value = [a.value for a in ar]
            if any(ar_value):                                                ## 忽略excel空白行
                for index,key in enumerate(ar_ctype):
                    if key==1:                                               ## 中文编码
                        ar_value[index]=ar_value[index].replace(' ','').encode('cp936','ignore')
                    elif key==3:                                             ## 日期格式化
                        date_cell_value = ar_value[index]
                        cell_as_datetime = arrow.get(*xlrd.xldate_as_tuple(date_cell_value,xls_path.datemode))
                        date_format = cell_as_datetime.format("YYYY-MM-DD")
                        ar_value[index] = date_format

                base_name = os.path.basename(xls).split('.')[0]
                ar_value.insert(0,base_name)                                  ## 文件名
                ar_value.insert(1,sh.encode('gb18030'))                       ## 工作表名
                th.append(ar_value)
    write_csv(th)

def write_csv(th):
    """ 写入csv表格 """
    with open('summary.csv','ab+') as f:
        wf =csv.writer(f,dialect ='excel')
        wf.writerows(th)

def write_first():
    """ 表格第一列插入文件名，第二列插入工作表名 """
    first_row = ['File name','Table name']
    with open('summary.csv','ab+') as f:
        wf =csv.writer(f,dialect ='excel')
        wf.writerow(first_row)

def del_exists():
    """ 清除已存在的summary.csv """
    file_path = os.getcwd()
    path = os.path.join(file_path,'summary.csv')
    if os.path.exists(path):
        os.remove(path)

if __name__=='__main__':
    """ 主程序 """
    freeze_support()
    print u'正在启动...'
    print u'汇总与本程序同一文件夹内excel文档，清除空白行，同文件夹下生成summary.csv，第一列为文件名，第二列为工作表名'
    del_exists()
    write_first()
    cpucount = cpu_count()
    p = Pool(cpucount)
    files = get_xlsfiles()
    for f in files:
        p.apply(get_result,(f,))
    p.close()
    p.join()
    print 'summary.csv has been created'
    print 'Copyright by Zhoub'
    os.system('pause')
