import xlrd
import xlwt
import cx_Oracle

Sheet = xlrd.open_workbook(r"C:\Users\Administrator\Desktop\Python\IPList.xlsx")
data = Sheet.sheets()[0]
nrow = data.nrows

FalledList = []

for i in range(1,nrow):
    ip = data.row_values(i)[0]

    conInfo = "ERIS/ERIS@%s:1521/ERIS"%(ip)
    print(conInfo)

    try :
        conn = cx_Oracle.connect(conInfo)
        cur = conn.cursor()
        cur.close()  # 关闭游标
        conn.close() # 关闭链接    
    except Exception as ret :
        print("%s 连接超时："%(ip), ret)
        FalledList.append(ip)   

print('连接失败的工控机ip有：')
for j in FalledList:
    print(j)

