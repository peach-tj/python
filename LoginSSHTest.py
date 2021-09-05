import xlrd
import xlwt
import paramiko

Sheet = xlrd.open_workbook(r"C:\Users\sbspf\Desktop\工控机程序更新\升级.xlsx")
data = Sheet.sheets()[0]
nrow = data.nrows

FalledList = []

for i in range(1,nrow):
    ip = data.row_values(i)[0]

    try :
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())   # 自动添加策略，保存服务器的主机名和秘钥信息，不添加，不在本地hnow_hosts文件中的记录将无法连接
        ssh.connect(hostname=ip, port=22, username="root", password="dhERIS@2018*#")
            
    except Exception as ret :
        print("%s 连接超时："%(ip), ret)
        FalledList.append(ip)   

print('连接失败的工控机ip有：')
for j in FalledList:
    print(j)

