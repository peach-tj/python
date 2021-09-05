import openpyxl
from openpyxl import Workbook
import cx_Oracle
import datetime
import time
import xlrd
import datetime
from time import strftime


stop_time = "2021-07-01 00:00:00"
start_time = "2021-06-01 00:00:00"
end_time = "2021-06-01 23:59:59"

CSheet = xlrd.open_workbook(r'C:\Users\sbspf\Desktop\工控机程序更新\python\下城区ip.xlsx',encoding_override='utf-8')   # 需要查找的号牌表，将此处修改为表格所在的路径及文件名
data = CSheet.sheets()[0]
nrows = data.nrows

def initData():
    start_time = "2021-06-01 00:00:00"
    end_time = "2021-06-01 23:59:59"
    return start_time,end_time


for i in range(1,nrows):
    ip = data.row_values(i)[0]
    conInfo = 'ERIS/ERIS@%s:1521/ERIS'%(ip)

    try:
        con = cx_Oracle.connect(conInfo)       # 连接数据库
        cur = con.cursor() # 获取游标

        sql1 = "update t_offline_upload_config set start_time = 0,end_time = 4"
        cur.execute(sql1)
        con.commit()
    
        print("正在删除 %s 数据……"%(ip))

        while(start_time <= stop_time):
            sql2 = "delete from t_non_vehicle_illegal where cap_time between to_date('%s','yyyy-mm-dd hh24:mi:ss') and to_date('%s','yyyy-mm-dd hh24:mi:ss')"%(start_time, end_time)
            sql3 = "delete from t_tag_pic_record where cap_date between to_date('%s','yyyy-mm-dd hh24:mi:ss') and to_date('%s','yyyy-mm-dd hh24:mi:ss')"%(start_time, end_time)

            cur.execute(sql2)
            con.commit()

            cur.execute(sql3)
            con.commit()
            
            start_time = (datetime.datetime.strptime(start_time, "%Y-%m-%d %H:%M:%S") + datetime.timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")
            end_time = (datetime.datetime.strptime(end_time, "%Y-%m-%d %H:%M:%S") + datetime.timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")

        start_time,end_time = initData()

    except Exception as err:
        print(err)
