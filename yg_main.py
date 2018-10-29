# -*- coding: utf-8 -*-



import pandas as pd
import numpy as np
import time
from package_methods.root_file import file_name,change_name,mymovefile
from package_methods.office import sheet_index
from package_methods.sql_orm import ALMSSQL
import datetime



if __name__=='__main__':

    ms = ALMSSQL()
    directory = r"D:\需入库文件夹" + "\\"  #目录路径
    start = time.clock()
    filename_list = file_name(directory)
    for filename in filename_list:
        if "xlsx" in filename and "~$" and "已入库-" not in filename:
                print(directory+filename)
                a = sheet_index(directory + filename)
                sheet_rn=a[0]
                sheet_name=a[1]
                if sheet_rn==1:
                    print("正在执行----->sheet"+str(sheet_name),'\n')
                    df = pd.read_excel(directory + filename)  # excel
                    new_columns = []
                    for col in range(len(df.columns)):
                        new_columns.append(str(df.columns[col]).replace("(", "").replace(")", ""))
                    df.columns = new_columns
                    df.to_sql(sheet_name[0], ms.engine, if_exists='append', index=False, chunksize=4000000)
                elif sheet_rn>1:
                    for i in range(sheet_rn):
                        print("正在执行----->sheet  "+str(sheet_name[i]),"\n")
                        df = pd.read_excel(directory + filename, sheetname=sheet_name[i])  # excel
                        new_columns=[]
                        for col in range(len(df.columns)):
                            new_columns.append(str(df.columns[col]).replace("(","").replace(")",""))
                        df.columns=new_columns
                        df.to_sql(sheet_name[i], ms.engine, if_exists='append', index=False, chunksize=150000)
                # change_name(directory,filename,"已入库-"+time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))+"-"+filename)
                mymovefile(directory+filename,"D:\历史数据源文件"+"\\"+filename)

    print("------------开始清洗数据-------------------")

    ms.ExecNonQuery("insert into [dbo].[信审信息] (门店,进件编号,进件状态,申请产品,进件时间,拒绝原因终审) select 门店,进件编号,进件状态,申请产品,进件时间,拒绝原因终审 from [pcfk_refuse_info_sql]")
    ms.ExecNonQuery("truncate  table pcfk_refuse_info_sql")
    print("自动拒单-->清洗完成\n")
    ms.ExecNonQuery("[Data_Clean] '贷后','进件编号'")
    print("贷后-->清洗完成\n")
    ms.ExecNonQuery("[Data_Clean] '运营','进件编号'")
    print("运营-->清洗完成\n")
    ms.ExecNonQuery("[Data_Clean] '客户经理和团队经理进件数据','order_num'")
    print("客户经理和团队经理进件数据-->清洗完成\n")
    ms.ExecNonQuery("[Data_Clean] '信审信息','进件编号'")
    print("信审信息-->清洗完成\n")
    ms.ExecNonQuery("[Data_Clean] '业务操作流程','进件编号,环节,开始时间'")
    print("业务操作流程-->清洗完成\n")
    ms.ExecNonQuery("[daily_paper_increment]")
    print("信审数据源_日报-->更新完成\n")



    end = time.clock()
    print('运行时间：%.4f秒 完成工作 退出！\n ' % (end - start))
    print("--------------------------分割线-------------------------------")




















