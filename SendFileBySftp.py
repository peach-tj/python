import xlrd
import paramiko
import os
import time
import hashlib

CSheet = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\Python\工控机程序更新\IPList.xlsx')
data = CSheet.sheets()[0]
nrow = data.nrows

FailedList = []
notCompleteList = []

# 检查路径是否存在
def _is_exists(path, function):
    path = path.replace('\\', '/')
    try:
        function(path)
    except Exception as error:
        return False
    else:
        return True


# 获取文件的MD5值
def _get_filemd5(filename):
    if not os.path.isfile(filename):
        return
    myhash = hashlib.md5()
    f = open(filename,'rb')
    while True:
        b = f.read(8096)
        if not b :
            break
        myhash.update(b)
    f.close()
    return myhash.hexdigest()


# 文件上传至临时目录，并判断文件MD5值是否一致
def _upload_tempfile(ssh, sftp, local, filename):
    TempFilename = TempFile + filename
    try:
        sftp.stat(TempFile)      # 检查临时目录是否存在
        ssh.exec_command("rm -f %s*"%(TempFile))
        ssh.exec_command("mkdir -p %s"%(TempFile))        
    except IOError as error:
        ssh.exec_command("mkdir -p %s"%(TempFile))
    finally:
        sftp.put(local, TempFilename)

        # 本地文件的MD5值
        localFileMd5 = _get_filemd5(local)
        print("本地文件MD5值为：%s"%(localFileMd5))

        # 上传文件的MD5值
        stdin, stdout, stderr = ssh.exec_command("md5sum %s |cut -d ' ' -f1"%(TempFilename))
        # print(stdout.read().decode().strip())
        TempFileMd5 = stdout.read().decode().strip()
        print("服务文件MD5值为：%s"%(TempFileMd5))

        if (localFileMd5 == TempFileMd5) :
            return True
        else:
            return False


# 拷贝文件
def _copy(ssh, sftp, local, remote):
    errCount = 0
    # 判断remote是否是目录
    if _is_exists(remote, function=sftp.chdir):
        # 是，获取local路径中的最后一个文件名拼接到remote中
        filename = os.path.basename(os.path.normpath(local))
        remote = os.path.join(remote, filename).replace('\\', '/')
    # 如果local为目录
    if os.path.isdir(local):
        # 在远程创建相应的目录
        _is_exists(remote, function=sftp.mkdir)
        # 遍历local
        for file in os.listdir(local):
            # 取得file的全路径
            localfile = os.path.join(local, file).replace('\\', '/')
            # 深度递归_copy()
            _copy(ssh=ssh, sftp=sftp, local=localfile, remote=remote)
    # 如果local为文件
    if os.path.isfile(local):
        uploadResult = _upload_tempfile(ssh, sftp, local, filename)
        if (uploadResult == True) :
            try:
                TempFilename = TempFile + filename

                ssh.exec_command("rm -rf %s"%(remote))
                ssh.exec_command("cp -rf %s %s"%(TempFilename, remote))

                ssh.exec_command("chmod 777 %s"%(remote))

            except Exception as error:
                print(error)
                print('[put]', local, '==>', remote, 'FAILED')
            else:
                print('[put]', local, '==>', remote, 'SUCCESSED')
        else :
            print("%s文件传输不完整"%(remote))
            errCount = errCount + 1
    return errCount


for i in range(1, nrow):
    ip = data.row_values(i)[0]
    TempFile = "/Temp/"
    errCounts = 0

    try:
        transport = paramiko.Transport((ip, 22))
        transport.connect(username='root', password='dhERIS@2018*#')

        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())   # 自动添加策略，保存服务器的主机名和秘钥信息，不添加，不在本地hnow_hosts文件中的记录将无法连接
        ssh._transport = transport

        print("%s 连接成功" % (ip))

        sftp = paramiko.SFTPClient.from_transport(transport)

        errCounts = _copy(ssh=ssh, sftp=sftp, local='C:\\Users\\Administrator\\Desktop\\Python\\工控机程序更新\\server-packages\\ERI-Server', remote='/mnt/soft/')

        if (errCounts > 0):
            notCompleteList.append(ip)

    except Exception as reason:
        print(reason)
        FailedList.append(ip)

print("连接失败工控机：")
for j in FailedList:
    print(j)

print("有传输文件失败的工控机：")
for j in notCompleteList:
    print(j)
