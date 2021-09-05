from paramiko.sftp_client import SFTP
import xlrd
import xlwt
import paramiko
import os
import time

CSheet = xlrd.open_workbook(r'C:\Users\sbspf\Desktop\工控机程序更新\杭州工控机.xlsx')
data = CSheet.sheets()[0]
nrow = data.nrows

FailedList = []


for i in range(1, nrow):
    ip = data.row_values(i)[1]

    try:
        transport = paramiko.Transport((ip, 22))
        transport.connect(username='root', password='dhERIS@2018*#')

        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())   # 自动添加策略，保存服务器的主机名和秘钥信息，不添加，不在本地hnow_hosts文件中的记录将无法连接
        ssh._transport = transport

        print("%s 连接成功" % (ip))

        sftp = paramiko.SFTPClient.from_transport(transport)
        
        ssh.exec_command("kill -9 $(ps aux|grep 'rinetd$'|awk '{print$2}')")
        time.sleep(1)
        sftp.put('rinetd', '/mnt/soft/softv/rinetd')
        time.sleep(1)
        ssh.exec_command("/mnt/soft/softv/rinetd &")

    except Exception as reason:
        print(reason)
        FailedList.append(ip)

print("连接失败工控机：")
for j in FailedList:
    print(j)

