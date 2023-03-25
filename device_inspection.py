from openpyxl import load_workbook
from netmiko import ConnectHandler
from time import strftime
from  threading import Thread
from time import time
from os import getcwd


def netmiko_dev_connect(ip,username,passwd,check_comm,dev_type,back_file,en_passwd=''):
    dev = {'device_type':dev_type,'ip':ip,'username':username,'password':passwd,'secret':en_passwd,'session_log': 'netdevops.log'}
    # try:
    connect = ConnectHandler(**dev)
    if not dev['secret'] == None:
        connect.enable()
    dev_name = connect.find_prompt().strip('<>') #匹配主机名
    now_time = strftime('-%Y-%m-%d-%H-%M-%S') #记录时间
    for comm in check_comm: #发送命令，保存输出
        #strip_prompt用于在输出的文本中会显示设备名
        #strip_command用于在输出的文本中显示输入的命令
        content = connect.send_command(comm,strip_prompt=False,strip_command=False)
        backup_dev_name  = back_file + dev_name + now_time
        with open( backup_dev_name + '.txt','a',encoding='utf-8') as backup_conf: #以设备名+时间命名的txt输出文件
            backup_conf.write(content + '\n')
    
        # success_list.append(dev_name + '-' + str(ip))
    # except:
    #     # fail_list.append(ip)
    #     pass


#judge函数，根据设备类型获取不同厂商对应的设备命令，返回巡检文件保存的目录路径
def dev_judge(DeviceBrand):
    dev_value = DeviceBrand.lower().strip()
    pwd_path=getcwd().replace('\\','/')
    #Huawei
    if dev_value == 'huawei':
        hw_comm = pwd_path + 'devInfo_file/hw_comm.txt'
        hw_bac_path = pwd_path + '/backup_file/hw_bac_file/' #backup路径最后要加/，否则文件无法创建成功
        return ['huawei',hw_comm,hw_bac_path]
    #New H3C
    elif dev_value == 'h3c':
        h3c_comm = pwd_path + '/devInfo_file/h3c_comm.txt'
        h3c_bac_path = pwd_path + '/backup_file/h3c_bac_file/'
        return ['hp_comware',h3c_comm,h3c_bac_path]
    #Fortinet
    elif dev_value == 'forigate':
        ft_comm = pwd_path + '/devInfo_file/ft_comm.txt'
        ft_bac_path = pwd_path + '/backup_file/ft_bac_file/'
        return ['fortinet',ft_comm,ft_bac_path]
    #Cisco
    elif dev_value == 'cisco':
        cisco_comm = pwd_path + '/devInfo_file/cisco_comm.txt'
        cisco_bac_path = pwd_path + '/backup_file/cisco_bac_file/'
        return ['cisco_ios',cisco_comm,cisco_bac_path]
    #Hillstone
    elif dev_value == 'hillstone':
        sg_comm = pwd_path + '/devInfo_file/sg_comm.txt'
        sg_bac_path = pwd_path + '/backup_file/sg_bac_file/'
        return ['SG',sg_comm,sg_bac_path]
    else: 
        return 'None' #如果没判断出来返回一个none

def main():
    start_time = time() #记录时间
    wb = load_workbook('device_info.xlsx')
    ws =wb.active
    row_range = ws.max_row #行数
    therad_list = []
    for row_list in range(2,row_range+1):
        dev_info = ws[row_list]
        dev_temp = []
        for i in dev_info: #遍历后添加到列表
            dev_temp.append(i.value)
        print(dev_temp)
        # dev_name = dev_temp[1]
        # dev_type = dev_temp[2]
        # dev_connect = dev_temp[4]
        # dev_user = dev_temp[5]
        # dev_pwd = dev_temp[6]
        # dev_ip = dev_temp[7]

        judge_result = dev_judge(dev_temp[2]) #判断设备厂商类型
        if not judge_result == 'None':
            check_command = open(judge_result[1],'r')
            thread_netmiko = Thread(target=netmiko_dev_connect,
            args=(dev_temp[7],
                  dev_temp[5],
                  dev_temp[6],
                  check_command,
                  judge_result[0],
                  judge_result[2],
                  dev_temp[8]
                  )
                  )
            thread_netmiko.start()
            therad_list.append(thread_netmiko)
        else:
            print('不支持此设备:',dev_temp[1])
    for y in therad_list:
        y.join()

    end_time = time()
    total_time = end_time - start_time
    print(total_time) 

main()
        