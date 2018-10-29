# -*- encoding: utf-8 -*-

import imaplib,time,os,sys,gc,datetime
from email.parser import Parser
from email.header import decode_header


#1、处理文件

def get_email_cntent(message, base_save_path,file):
    j = 0
    content = ''
    attachment_files = []
    for part in message.walk():
        j = j + 1
        file_name = part.get_filename()
        contentType = part.get_content_type()
        # 保存附件
        if file_name:  # Attachment
            dh = decode_header(file_name)
            if dh[0][1]:  # 如果包含编码的格式，则按照该格式解码
                try:
                    filename = str(dh[0][0], encoding="gb2312",errors='strict')
                except:
                    filename = str(dh[0][0], encoding="UTF8", errors='strict')

                if filename[:len(filename)- 20]  in file:
                    print("获取到附件:" + filename,"\n")
                    data = part.get_payload(decode=True)
                    att_file = open(base_save_path + filename, 'wb')  # base_save_path 保存当前路径
                    attachment_files.append(filename)
                    att_file.write(data)
                    att_file.close()
                    print("保存成功!路径----->" + base_save_path,"\n")
                    print("--------------数据开始入库------------------")
                    os.system("python D:/py脚本/yg_main.py")

            elif file_name[:len(file_name)- 20] in file:
                print("获取到附件:" + file_name,"\n")
                data = part.get_payload(decode=True)
                att_file = open(base_save_path + file_name, 'wb') # base_save_path 保存当前路径
                attachment_files.append(file_name)
                att_file.write(data)
                att_file.close()
                print("保存成功!路径----->"+base_save_path,"\n")
                print("--------------数据开始入库------------------")
                os.system("python D:/py脚本/yg_main.py")
    # return  attachment_files


#2、登陆邮箱
def log_main(account_number,password):

    while True:
        try:
            server_side = imaplib.IMAP4_SSL("imap.qiye.163.com", 993)  # 连接smtp服务，建立登陆对象
            server_side.login(account_number, password)     #登录个人帐号
            print("用户 ：%s---> 登陆成功\n"%account_number)
            server_side.select("INBOX")             #进入收件箱
            typ, data = server_side.search(None, 'ALL')     #查看邮件
            msgList = data[0].split()                       #获取邮件编号列表
            oldlast_num =len(msgList)
            print('现在数量:%s\n等待接收新的邮件.......' % oldlast_num)
            mail_txt_list=[]
            while True: #判断收件箱的邮件数量与之前的是否相同，如果不同则向下执行，如果相同则继续循环做判断
                server_side.select("INBOX")  # 进入收件箱
                typ, data = server_side.search(None, 'ALL')
                msgList = data[0].split()  # 获取邮件编号列表continue)
                newlast_num = len(msgList)

                if newlast_num == oldlast_num:
                    # print('无新的邮件----现在数量:%s'% oldlast_num)
                    time.sleep(2)
                    continue
                else:
                    print('收到新的邮件----现在数量:%s\n'% newlast_num)
                    break

            mail_num_list=msgList[len(msgList) - (newlast_num-oldlast_num):]
            for main_num in mail_num_list:
                # last_num = msgList[len(msgList) - 1]
                typ_mail, datas = server_side.fetch(main_num, '(RFC822)')       #将邮件取来
                mail_txt = Parser().parsestr(datas[0][1].decode('utf-8'))
                mail_txt_list.append(mail_txt)
            break
        except:
            Nowtime = str(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            exc_type, exc_value, exc_tb = sys.exc_info()
            erro = "\n出现错误\n错误类型 : %s\n错误内容 : %s\n时间 : %s\n" % (str(exc_type)[1:-1], exc_value, Nowtime)
            print(erro)
            time.sleep(60)

    return   mail_txt_list


#3、主程序
def main(base_save_path,file):
    while True:
        mail_txt_list=log_main("zhaotianyu@yingu.com", "ZTYzty753951")
        for mail_txt in mail_txt_list:
            get_email_cntent(mail_txt, base_save_path,file)
        gc.collect()


if __name__ == '__main__':
    base_save_path = r'D:\需入库文件夹\\'
    file = ["loan_rm_pro1", "incoming_data","overdue_data","pcfk_refuse_info"]
    main(base_save_path, file)









