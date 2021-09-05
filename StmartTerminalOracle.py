import cx_Oracle
import xlrd
import xlwt
import paramiko

Sheet = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\Python\工控机程序更新\IPList.xlsx')
data = Sheet.sheets()[0]
nrow = data.nrows

TimeoutList = []

for i in range(1, nrow):
    ip = data.row_values(i)[0]
    conInfo = "ERIS/ERIS@%s:1521/ERIS"%(ip)
    
    sqls = ['update T_DATA_MIX_CONFIG set offset_digit = 0',
            'update t_offline_upload_config set start_time=0,end_time=4',
            'update T_ILLEGAL_ANALYSIS_CONFIG set is_analysis_tag_pic=0, is_analysis_tag_video=0, is_analysis_camera_pic=0, is_analysis_camera_video=0, is_analysis_mixed_pic=0, is_analysis_mixed_video=1, is_analysis_upload_data=0, is_analysis_offline_data=1 where id =1',
            'truncate table t_flux',
            'truncate table t_flux_runtime']

    try:
        con = cx_Oracle.connect(conInfo)
        cur = con.cursor()

        for sql in sqls:
            cur.execute(sql)
            con.commit()
        print("%s 提交成功"%(ip))

        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(hostname=ip, port=22, username='root', password='dhERIS@2018*#')

        ssh.exec_command("kill -9 $(ps aux|grep 'CenterServer.exe$'|awk '{print$2}') ")
 
    except Exception as err:
        print("%s 连接超时： "%(ip),err)
        TimeoutList.append(ip)

if len(TimeoutList) > 0 :
    print("连接超时的工控机有：")
    for j in TimeoutList :
        print(j)



