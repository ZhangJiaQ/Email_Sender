import smtplib
import xlrd
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header


def data_reader():
    data = xlrd.open_workbook(r'Excel.xls')
    table = data.sheets()[0]

    #读取登陆第三方STMP邮箱登陆信息
    mail_data = {
        'user': table.cell(1, 3).value,
        'password': table.cell(1, 4).value,
        'attachment': table.cell(1, 5).value,
    }

    #不同的发件人邮箱不同的第三方STMP地址设置
    if re.search('@foxmail.com', mail_data['user']) or re.search('@qq.com', mail_data['user']):
        mail_data['mail_host'] = 'stmp.qq.com'
        mail_data['port'] = 465
    if re.search('@163.com', mail_data['user']):
        mail_data['mail_host'] = 'smtp.163.com'
        mail_data['port'] = 25
    if re.search('@126.com', mail_data['user']):
        mail_data['mail_host'] = 'smtp.126.com'
        mail_data['port'] = 25
    if re.search('@sina.com', mail_data['user']):
        mail_data['mail_host'] = 'smtp.sina.com'
        mail_data['port'] = 25

    #遍历公司信息，将应聘邮件发给EXCEL中所有记录的公司
    for i in range(1,table.nrows):
        # print (table.row_values(i))
        company = table.row_values(i)
        email_sender(mail_data, company)
    return True


def email_sender(mail_data, company):

    # 构造MIMEMultipart对象做为根容器
    message = MIMEMultipart()

    message['From'] = Header(mail_data['user'])
    message['To'] = Header(company[2])
    subject = '应聘{0}，***，182*******5'.format(company[1])
    message['Subject'] = Header(subject, 'utf-8')

    # 构造MIMEText对象做为邮件显示内容并附加到根容器
    content = '''
先生/女士您好，
    我在拉钩网上看到贵公司的招聘信息，我对{0}的职位非常有兴趣，特来应聘。
    对照公司及职位的要求，我的情况简述如下：
    1、
    2、
    3、
    综上，我认为自己能够胜任，也非常喜欢这份工作。希望得到面试机会，谢谢!
    更多我的信息，详见附件简历。
    祝工作愉快!
    ***
    182*****715
    '''.format(company[1])
    message.attach(MIMEText(content))

    #构造附件1，传送当前目录下的 attachment 文件
    att1 = MIMEText(open(mail_data['attachment'], 'rb').read(), 'base64', 'utf-8')
    att1["Content-Type"] = 'application/octet-stream'
    att1.add_header('Content-Disposition', 'attachment', filename=mail_data['attachment'])
    message.attach(att1)

    try:
        if re.search('@foxmail.com', mail_data['user']) or re.search('@qq.com', mail_data['user']):
            smtpObj = smtplib.SMTP_SSL()
            smtpObj.connect(mail_data['mail_host'], mail_data['port'])
        else:
            smtpObj = smtplib.SMTP()
            smtpObj.connect(mail_data['mail_host'], mail_data['port'])
        smtpObj.login(mail_data['user'], mail_data['password'])
        smtpObj.sendmail(mail_data['user'], company[2], message.as_string())
        print('{0},OK'.format(company[2]))
        print ("邮件发送成功")

    except smtplib.SMTPException:
        print ("Error: 无法发送邮件")


data_reader()