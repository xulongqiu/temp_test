#!/usr/bin/python
# -*- coding: utf-8 -*-

# Ahter: James.Li
# Type:  Private

import os
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart

# 邮件发送类
class mail_sender:
    # ssl表示是否使用加密方式, 默认: 是
    def __init__(self, ssl=True):
        self.use_ssl = ssl
        # 邮件接受者
        self.receiver = ""
        # 发送方的邮件服务器登录账号
        self.usr_pwd = {'james.li@libratone.com.cn':'<PASSWD>', 
                        'liyangzmx@163.com':'<PASS_WORD>'}
        # SMTP邮件发送服务器地址配置
        self.stmp_servers = {'163.com':'smtp.163.com', 
                            'libratone.com.cn':'smtp.exmail.qq.com'}
        # SMTP邮件发送服务器端口配置
        self.smtp_ports = {'163.com':25, 
                          'libratone.com.cn':25}
        # SMTP邮件发送服务器SSL端口设置
        self.smtp_ssl_ports = {'163.com':465, 
                          'libratone.com.cn':465}        
        # 创建多媒体对象
        self.msg = MIMEMultipart('related')
        
    # 添加HTML到邮件正文
    def add_html(self, html):
        text_msg = MIMEText(html, 'html', _charset='utf-8')
        self.msg.attach(text_msg)        
        pass
    
    # 添加图片到邮件附件
    def add_image(self, src, img_id):
        if not isinstance(self.msg, MIMEMultipart):
            print 'MIME message not a \'MIMEMultipart\' instance'
            return
        dst = src
        # 支持目录中含有'~'的情况
        if src.find('~') == 0:
            dst = os.environ['HOME'] + src[1:]
            print 'Add picture: ', dst
        fp = open(dst, 'rb')
        # 创建MIME图片对象
        msg_img = MIMEImage(fp.read())
        fp.close()
        # 创建附件信息头
        msg_img.add_header('Content-ID', '<' + img_id.split('.')[0] + '>')
        print 'Content-ID: ' + '<' + img_id.split('.')[0] + '>'
        # 添加附件属性字段
        msg_img.add_header('Content-Type', 'image/png')
        # 关联附件到邮件
        self.msg.attach(msg_img)
        
    # 设置邮件发送这
    def set_sender(self, sender):
        self.sender = sender
        self.pwd = self.usr_pwd[sender]
        # 解析发送方的邮件地址, 去除服务器域名
        server = sender.split('@')[1]
        # 通过服务器的域名自动选区合适的配置
        self.stmp_server = self.stmp_servers[server]
        if self.use_ssl:
            self.smtp_port = self.smtp_ssl_ports[server]
        else:
            self.smtp_port = self.smtp_ports[server]
            
    # 添加抄送, 该部分目前NG, 没有事件找原因, 建议先发送给自己
    def set_cc(self, cc):
        self.msg['CC'] = cc
    
    # 添加密送, 该部分目前NG, 没有事件找原因, 建议先发送给自己
    def set_bcc(self, bcc):
        self.msg['BCC'] = bcc
    
    # 发送邮件, 
    # receiver: 接受者邮箱
    # subject: 标题
    def send(self, receiver, subject):
        self.receiver = receiver
        self.msg['subject'] = subject
        self.msg['from'] = self.sender
        self.msg['to'] = self.receiver
        if self.use_ssl:
            server = smtplib.SMTP_SSL(self.stmp_server, self.smtp_port)
            server.set_debuglevel(1)
            server.ehlo()
            # server.starttls()             
        else:
            server = smtplib.SMTP_SSL(self.stmp_server, self.smtp_port)
        server.login(self.sender, self.pwd)
        server.sendmail(self.sender, self.receiver, self.msg.as_string())
        server.quit()
        
if __name__ == '__main__':
    servers = {'163':'send_163', 'libratone':'send_libratone'}
    mail = mail_sender()
    reciever = u'james.li@libratone.com.cn'
    subject = u'下班提醒'
    mail.add_image('~/pic/lcd_driver.png', 'image1')
    # 邮件HTML主体
    body = u'''
    <h1>
	<p>Hi, all:</p>
    </h1>
    <p>结温图表:</p>
    <img src="cid:%s">
    ''' % 'image1'
    mail.add_html(body)
    mail.set_sender('james.li@libratone.com.cn')
    mail.send(reciever, subject)
