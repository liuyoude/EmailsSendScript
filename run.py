# coding: utf-8
# author: liuyoude

import sys
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import smtplib
from email.mime.text import MIMEText
import pathlib
import pandas as pd
import yaml

def load_config(config_path):
    with open(config_path, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)
    return config

def get_contact_info_from_xlsx(file_path, keys, filte_dict=None):
    contact_df = pd.read_excel(file_path, engine='openpyxl')
    print(contact_df.keys())
    if filte_dict:
        for filte_key in filte_dict.keys():
            if filte_key in contact_df.keys():
                for value in filte_dict[filte_key]:
                    contact_df = contact_df[contact_df[filte_key].str.contains(value)==False]
    contact_dict = {}
    for key in keys:
        if key not in contact_df.keys():
            raise ValueError(f'{key} is not in your file.')
        contact_dict[key] = contact_df[key].values
    return contact_dict

def save_contact_info(file_path, contact_dict):
    info_list = []
    for key in contact_dict.keys():
        info_list.append(contact_dict[key])
    for info_tuple_i in zip(*info_list):
        with open(file_path, 'a', encoding='utf-8') as f:
            info_str_i = ""
            for info_item_i in info_tuple_i:
                info_str_i += str(info_item_i)
                info_str_i += '\t'
            info_str_i += '\n'
            f.write(info_str_i)

def content_template(name, duty):
    head = f'尊敬的{name}{duty},\n'
    body = '您好！第十一届环化大会将于2021年12月24-28日在哈尔滨召开，本次大会首次安排了“环境科学与工程教育论坛”，以环境教育为核心，聚焦环境科学与工程专业、学科在人才培养、科学研究、社会服务、国际交流等方面的改革与创新；对标国际上环境类学科发展，梳理我国环境类专业学科发展历史、总结经验、探讨未来发展方向。分会将安排新时期人才培养、课程教材（建设）、课程思政、新工科、学科交叉、教学保障条件与机制等内容。该分会是环境化学大会首次安排的环境教育分会，对于环境科学与工程教育具有重要指导价值。欢迎各位老师莅临参会，摘要网上投稿截止10月20日，http://www.ncec2021.com，详见本论坛宣传海报，谢谢！'
    tail = '十一届环化大会58分会“环境科学与工程教育论坛”'
    return head + '\n' + ' '*8 + body + '\n' + ' '*144 + tail

def get_contents(names, dutys):
    contents = []
    for name, duty in zip(names, dutys):
        contents.append(content_template(name, duty))
    return contents


def send_mail(username, passwd, recviers, title, contents, mail_host='smtp.163.com', port=25, files=None):
    '''
    发送邮件函数，默认使用163smtp
    :param username: 邮箱账号 xx@163.com
    :param passwd: 邮箱密码
    :param recv: 邮箱接收人地址，多个账号以逗号隔开
    :param title: 邮件标题
    :param content: 邮件内容
    :param mail_host: 邮箱服务器
    :param port: 端口号
    :param files: 邮件附件
    :param files_name: 邮件附件名称
    :return:
    '''
    # load success info and rewrite error info
    with open('./success.txt', 'r') as f:
        have_send_set = set(f.read().split())
    with open('./error.txt', 'w') as f:
        f.write('')

    # send personal email
    for recvier, content in zip(recviers, contents):
        try:
            # avoid send repeatedly
            if recvier in have_send_set:
                print(f'{recvier} have send!')
                continue
            # attach multiple additional files
            if files:
                msg = MIMEMultipart()
                # 构建正文
                part_text = MIMEText(content)
                msg.attach(part_text)  # 把正文加到邮件体里面去

                # 构建邮件附件
                for file in files:
                    with open(file, 'rb') as f:
                        part_attach = MIMEApplication(f.read())  # 打开附件
                    part_attach.add_header('Content-Disposition', 'attachment', filename=pathlib.Path(file).name)
                    # part_attach.add_header('Content-Disposition', 'attachment', filename=files_name)  # 为附件命名
                    msg.attach(part_attach)  # 添加附件
            else:
                msg = MIMEText(content)  # 邮件内容
            msg['Subject'] = title  # 邮件主题
            msg['From'] = username  # 发送者账号
            msg['To'] = recvier  # 接收者账号列表
            # login email
            smtp = smtplib.SMTP(mail_host, port=port)
            smtp.login(username, passwd)
            smtp.sendmail(username, recvier, msg.as_string())
            #  record success info
            with open('./success.txt', 'a', encoding='utf-8') as f:
                f.write(f'send {recvier} success!' + '\n')
            print(f'send {recvier} success!')
        except Exception:
            # record fail info
            with open('./error.txt', 'a', encoding='utf-8') as f:
                f.write(str(recvier)+'\t'+str(sys.exc_info()) + '\n'*3)
            print(f'send {recvier} fail!')

    smtp.quit()

if __name__ == '__main__':
    # ===================
    # get config file
    # ===================
    config_path = './config.yaml'
    config = load_config(config_path)
    sender = config['sender']
    pwd = config['pwd']
    mail_host = config['mail_host']
    port = config['port']
    title = config['title']
    add_files = config['addition_files']
    contact_path = config['contact_file']

    # ================================
    # get needed contact info from xlsx
    # ================================
    # keys
    contact_keys = ['高校', '联系人姓名', '职务', '邮箱']
    # key, value that need be filtered out
    filte_dict = {
        '高校': ['清华大学', '哈尔滨工业大学', '南开大学', '同济大学'],
    }
    contact_dict = get_contact_info_from_xlsx(contact_path, contact_keys, filte_dict=filte_dict)
    # save info need to send
    save_contact_info('./info.txt', contact_dict)

    # ===================
    # get special email content for every one from defined template
    # ===================
    names = contact_dict[contact_keys[1]]
    dutys = contact_dict[contact_keys[2]]
    recivers = contact_dict[contact_keys[3]]
    contents = get_contents(names, dutys)

    # ===================
    # send mail
    # ===================
    send_mail(username=sender,
              passwd=pwd,
              recviers=recivers,
              title=title,
              contents=contents,
              mail_host=mail_host,
              files=add_files,
              port=port)