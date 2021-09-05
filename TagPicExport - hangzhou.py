import xlrd
import xlwt
import urllib3
import time
from time import strftime
import os
import cx_Oracle

con = cx_Oracle.connect('ERIS/ERISRP@33.83.247.42:1521/ERIS')
cur = con.cursor()

CSheet = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\hphm.xlsx',encoding_override='utf-8')   # 需要导出图片的表格路径
data = CSheet.sheets()[0]
nrows = data.nrows
print(nrows)

SqlList = []

for i in range(1,nrows):
    hphm = data.row_values(i)[1]
    hphm = ''.join(filter(lambda c:ord(c)<256, hphm))   # 去除汉字

    sql = "select distinct(hphm),dev_name,cap_date,car_img1_url,pts_ip,pts_port from t_tag_pic_record where record_type = 1 and hphm = '%s'"%(hphm)
    cur.execute(sql)
    SqlRes = cur.fetchall()
    print(len(SqlRes))
    for j in range(len(SqlRes)) :
        print(SqlRes[j])
        SqlList.append(SqlRes[j])

a = os.path.exists('C:\\Users\\Administrator\\Desktop\\tag_pic')    # 导出过车图片路径，首先检测，如果不存在则创建
if a == False :
    os.mkdir('C:\\Users\\Administrator\\Desktop\\tag_pic')
else :
    print(a)

for k in SqlList:
    print(k)
    url = "http://%s:%s/%s"%(k[4], k[5], k[3])              # 图片请求的url
    fileName = "%s_%s_%s.jpg"%(k[0], k[1], k[2])    # 图片保存文件名
    pathFile = "C:\\Users\\Administrator\\Desktop\\tag_pic\\%s"%(fileName)

    # print(url)
    # print(fileName)

    http = urllib3.PoolManager()
    res = http.request('GET', url).data
    status = http.request('GET',url).status

    if status == 200 :
        with open (pathFile, "wb") as f :
            f.write(res)
    else :
        print("图片不存在")
   


