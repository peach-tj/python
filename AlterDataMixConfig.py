import cx_Oracle
import xlrd
import paramiko
import time

Sheet = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\PythonTest\IPList.xlsx')
data = Sheet.sheets()[0]
nrow = data.nrows

TimeoutList = []

for i in range(1, nrow):
    ip = data.row_values(i)[0]
    conInfo = "ERIS/ERIS@%s:1521/ERIS"%(ip)
    
    sql = 'update T_DATA_MIX_CONFIG set offset_digit = 2'

    try:
        con = cx_Oracle.connect(conInfo)
        cur = con.cursor()
        cur.execute(sql)
        con.commit()
        print("%s 提交成功"%(ip))

        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(hostname=ip, port=22, username='root', password='dhERIS@2021*#')

        ssh.exec_command("kill -9 $(ps aux|grep 'CenterServer.exe$'|awk '{print$2}') ")
        # 启动看门狗
        chan = ssh.invoke_shell()
        chan.send('nohup sh /mnt/soft/softv/dogdaemon.sh & \n')
        time.sleep(2)
 
    except Exception as err:
        print("%s 连接超时： "%(ip),err)
        TimeoutList.append(ip)

if len(TimeoutList) > 0 :
    print("连接超时的工控机有：")
    for j in TimeoutList :
        print(j)

