import pandas as pd
import glob
import fnmatch
from selenium import webdriver
import time

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
import os
import schedule


# 导入xls，xlsx，csv文件
class DataImporter:
    def __init__(self):
        self.init = 0
        self.data = pd.DataFrame()

    def login(self, http, username, password):
        self.driver = webdriver.Chrome('./chromedriver.exe')  # 替换为您的Chrome驱动程序路径

        # 打开登录页面
        self.driver.get(http)  # 替换为登录页面的URL

        # 找到用户名和密码输入框，输入登录信息
        username_input = self.driver.find_element_by_xpath('//*[@id="uc-common-account"]')  # 替换为用户名输入框的XPath
        password_input = self.driver.find_element_by_xpath('//*[@id="ucsl-password-edit"]')  # 替换为密码输入框的XPath

        username_input.send_keys(username)
        password_input.send_keys(password)
        time.sleep(2)
        # 提交登录表单
        password_input.submit()

    def close(self):
        # 关闭浏览器
        self.driver.quit()

    def import_data(self, folder_path, file_pattern):
        # 模糊匹配文件
        file_paths = glob.glob(folder_path + '/' + file_pattern)
        if len(file_paths) == 0:
            raise ValueError('No matching files found')

        # 导入数据
        for file_path in file_paths:
            if file_path.endswith('.xlsx'):
                xls = pd.ExcelFile(file_path)
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)
                    df['消费类型'] = sheet_name  # 添加消费类型列
                    df['工作表名称'] = xls.sheet_names  # 添加工作表名称列
                    self.data = self.data.append(df)
            elif file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
                df['消费类型'] = 'CSV'  # 添加消费类型列
                df['工作表名称'] = 'N/A'  # 添加工作表名称列
                self.data = self.data.append(df)
            else:
                raise ValueError('Unsupported file format')

    def export_to_csv(self, output_folder):
        for sheet_df in self.data:
            sheet_name = sheet_df['Sheet'].iloc[0]
            output_path = output_folder + '/' + sheet_name + '.csv'
            sheet_df.to_csv(output_path, index=False, encoding='utf-8-sig')

    def get_sheet_data(self, sheet_name):
        return self.data[self.data['Sheet'] == sheet_name]

class EmailSender:
    def __init__(self, smtp_server, port, sender_email, password):
        self.smtp_server = smtp_server
        self.port = port
        self.sender_email = sender_email
        self.password = password

    def send_email(self, recipients, cc, subject, body_texts, images, attachments, html_signature):
        msg = MIMEMultipart('related')
        msg['From'] = self.sender_email
        msg['To'] = ', '.join(recipients)
        msg['Cc'] = ', '.join(cc)
        msg['Subject'] = subject

        msg_alternative = MIMEMultipart('alternative')
        msg.attach(msg_alternative)

        for text in body_texts:
            msg_text = MIMEText(text, 'plain')
            msg_alternative.attach(msg_text)

        for image_path in images:
            with open(image_path, 'rb') as f:
                msg_image = MIMEImage(f.read())
                msg_image.add_header('Content-ID', '<{}>'.format(os.path.basename(image_path)))
                msg.attach(msg_image)

        msg_html = MIMEText(html_signature, 'html')
        msg_alternative.attach(msg_html)

        for attachment_path in attachments:
            with open(attachment_path, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename="{}"'.format(os.path.basename(attachment_path)))
                msg.attach(part)

        with smtplib.SMTP(self.smtp_server, self.port) as server:
            server.login(self.sender_email, self.password)
            server.sendmail(self.sender_email, recipients + cc, msg.as_string())




if __name__ == '__main__':


    email_sender = EmailSender('mail.xm12t.com', 25, 'fanglongsheng@xm12t.com', '008759')
    recipients = ['fanglongsheng@xm12t.com']
    cc = ['fanglongsheng@xm12t.com']
    subject = 'Subject'
    body_texts = ['Text1', 'Text2']
    images = []
    attachments = ['匹配结果.xlsx']
    html_signature = "<html><body><p>Best regards,<br>John</p></body></html>"

    # 定义定时任务，每天的特定时间发送邮件
    schedule.every().day.at('14:16').do(email_sender.send_email(recipients, cc, subject, body_texts, images, attachments, html_signature))  # 修改为您希望的发送时间

    # while True:
    #     schedule.run_pending()
    #     time.sleep(1)



