import cx_Oracle
import xlrd
import paramiko
import os
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8' 


Sheet = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\Python\工控机程序更新\IPList.xlsx')
data = Sheet.sheets()[0]
nrow = data.nrows

TimeoutList = []

localFile = "C:/Users/Administrator/Desktop/批量分区/OracleFile/rman_incr.pl"
remoteFile = "/u01/app/oracle_scripts/rman_incr.pl"


# 获取创建违法和过车Job的SQL
CJobSqls = []
rootdir = 'C:\\Users\\Administrator\\Desktop\\批量分区\\OracleFile\\JobSql'
list = os.listdir(rootdir) #列出文件夹下所有的目录与文件
for i in range(0,len(list)):
       path = os.path.join(rootdir,list[i])
       if os.path.isfile(path):
            with open(path,"r",encoding="utf-8") as f:
                CJobSql = f.read().replace("\n","")
                CJobSqls.append(CJobSql)


# 获取连接信息
def _create_ssh(ip):
    transport = paramiko.Transport((ip, 22))
    transport.connect(username='root', password='dhERIS@2018*#')
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())   # 自动添加策略，保存服务器的主机名和秘钥信息，不添加，不在本地hnow_hosts文件中的记录将无法连接
    ssh._transport = transport
    sftp = paramiko.SFTPClient.from_transport(transport)

    return ssh,sftp


for i in range(1, nrow):
    ip = data.row_values(i)[0]
    conInfo = "ERIS/ERIS@%s:1521/ERIS"%(ip)

    try:
        ssh,sftp = _create_ssh(ip)
        print("1. %s 连接成功！"%(ip))

        # 上传 rman_incr.pl 文件，去掉备份
        sftp.put(localFile,remoteFile)
        print("2. %s 备份文件已上传至 %s"%(localFile,remoteFile))

        # 连接数据库
        con = cx_Oracle.connect(conInfo)
        cur = con.cursor()

        # 删除redo日志
        redoLogSql = "select group#, status from v$log"
        cur.execute(redoLogSql)
        results = cur.fetchall()
        inactiveList = []
        if (len(results) > 2) :
            for i in range(len(results)):
                if (results[i][1] == "INACTIVE"):
                    groupid = results[i][0]
                    inactiveList.append(groupid)
            for j in range(len(inactiveList)-1):
                # 删除redo log文件
                redoLogFileSql = "select member from v$logfile where group# = %s"%(inactiveList[j])
                cur.execute(redoLogFileSql)
                results1 = cur.fetchall()
                logFile = results1[0][0]
                ssh.exec_command("rm -rf %s"%(logFile))
                # 删除redo记录
                cur.execute("ALTER DATABASE drop logfile group %s"%(inactiveList[j]))
                con.commit()
            print("3. redo日志与记录已删除")
        else:
            print("3. redo日志与记录无需删除")
        
        sqls = ['ALTER DATABASE FLASHBACK OFF',
                'alter database drop supplemental log data']

        # 去掉闪回和补充日志
        for sql in sqls:
            cur.execute(sql)
            con.commit()
        print("4. 闪回和补充日志已关闭")


        JobSqls = ["drop procedure P_CLEAN_ILLEGAL_RECORD",
                    "drop procedure P_CLEAN_TAG_RECORD",
                    "declare vv_job_name varchar2(400); begin vv_job_name := upper('JOB_T_P_DLOG_FLUX'); for job in (select job_name from user_scheduler_jobs where upper(job_name) = vv_job_name) loop dbms_scheduler.drop_job(job_name => vv_job_name, force => true); end loop; end;",
                    "declare vv_job_name varchar2(400); begin vv_job_name := upper('JOB_T_P_CLN_SURVEY_ALARM'); for job in (select job_name from user_scheduler_jobs where upper(job_name) = vv_job_name) loop dbms_scheduler.drop_job(job_name => vv_job_name, force => true); end loop; end;",
                    "declare vv_job_name varchar2(400); begin vv_job_name := upper('JOB_T_P_CLN_FLUX_RUNTIME'); for job in (select job_name from user_scheduler_jobs where upper(job_name) = vv_job_name) loop dbms_scheduler.drop_job(job_name => vv_job_name, force => true); end loop; end;",
                    "declare vv_job_name varchar2(400); begin vv_job_name := upper('JOB_T_P_CLN_FLUX'); for job in (select job_name from user_scheduler_jobs where upper(job_name) = vv_job_name) loop dbms_scheduler.drop_job(job_name => vv_job_name, force => true); end loop; end;",
                    "declare vv_job_name varchar2(400); begin vv_job_name := upper('JOB_T_P_CLN_READ_ERROR_STATUS'); for job in (select job_name from user_scheduler_jobs where upper(job_name) = vv_job_name) loop dbms_scheduler.drop_job(job_name => vv_job_name, force => true); end loop; end;",
                    "declare vv_job_name varchar2(400); begin vv_job_name := upper('JOB_T_P_CLN_READ_STATISTICS'); for job in (select job_name from user_scheduler_jobs where upper(job_name) = vv_job_name) loop dbms_scheduler.drop_job(job_name => vv_job_name, force => true); end loop; end;"]

        # 删除及优化Job
        for JobSql in JobSqls:
            cur.execute(JobSql)
            con.commit()
        print("Job计划已删除")
        for CJobSql in CJobSqls:
            cur.execute(CJobSql)
        print("违法和过车Job已重新创建")
        print("5. 删除及优化Job已完成")

    except Exception as err:
        print("%s 连接超时： "%(ip),err)
        TimeoutList.append(ip)

if len(TimeoutList) > 0 :
    print("连接超时的工控机有：")
    for j in TimeoutList :
        print(j)

