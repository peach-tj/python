import cx_Oracle
import pandas
import time
import datetime
from time import strftime

tag_list = []       # 全部生产数据
exit_tag = []       # 有识读记录的生产数据
nexit_tag = []      # 无识读的生产数据

excelPath = r"C:\Users\Administrator\Desktop\测试文件\python脚本\生产管控\tag_produce_hangzhou.xlsx"

# 获取所有生产数据
def get_tag_produce_list(path):
    df = pandas.read_excel(path)
    data = df.values
    tag_list = []

    for j in range(len(data)):
        tagHphm = data[j][1]
        tag_list.append(tagHphm)
    
    return tag_list


start_time = "2021-01-01 00:00:00"
end_time = "2021-01-01 23:59:59"
cur_time = strftime("%Y-%m-%d %H:%M:%S",time.localtime())

tag_list = get_tag_produce_list(excelPath)

con = cx_Oracle.connect('ERIS/ERIS@172.17.2.31:1521/ERIS')
cur = con.cursor()

while(end_time <= cur_time):

    print("正在查询 %s 的过车数据"%(start_time.split(' ')[0]))

    sql = "select distinct hphm,TID from t_tag_pic_record where record_type != 1  and (cap_date between to_date('%s','yyyy-mm-dd hh24:mi:ss') and to_date('%s','yyyy-mm-dd hh24:mi:ss'))"%(start_time, end_time)
    
    # 执行sql查询
    cur.execute(sql)
    results = cur.fetchall()
    rows = len(results)

    for i in range(rows):
        rfid_hphm = results[i][0]

        if (rfid_hphm in tag_list) and (rfid_hphm not in exit_tag) :
            exit_tag.append(rfid_hphm)
        else :
            continue
    
    print("%s 过车查询完成！"%(start_time.split(' ')[0]))

    start_time = (datetime.datetime.strptime(start_time, "%Y-%m-%d %H:%M:%S") + datetime.timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")
    end_time = (datetime.datetime.strptime(end_time, "%Y-%m-%d %H:%M:%S") + datetime.timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")


# 遍历所有生产数据，找出未识读的数据
for k in tag_list:
    if k not in exit_tag:
        nexit_tag.append(k)

# 将识读过的数据写入到Excel
if len(exit_tag) > 0 :
    data_df = pandas.DataFrame(exit_tag, columns=['识读号牌号码'])
    writer = pandas.ExcelWriter('exit_tag.xlsx')
    data_df.to_excel(writer, index=None)
    writer.save()

# 将未识读过的数据写入到Excel
if len(nexit_tag) > 0 :
    data_df2 = pandas.DataFrame(nexit_tag, columns=['未识读号牌号码'])
    writer2 = pandas.ExcelWriter('nexit_tag.xlsx')
    data_df2.to_excel(writer2, index=None)
    writer2.save()

