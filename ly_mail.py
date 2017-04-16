#!/usr/bin/python
# -*- coding: utf-8 -*-

# Ahter: James.Li
# Type:  Private

import os
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart

class mail_sender:
    def __init__(self, ssl=True):
        self.use_ssl = ssl
        self.receiver = ""
        self.usr_pwd = {'james.li@libratone.com.cn':'Zmx13224548916', 
                        'liyangzmx@163.com':'13224548916,,,'}
        self.stmp_servers = {'163.com':'smtp.163.com', 
                            'libratone.com.cn':'smtp.exmail.qq.com'}
        self.smtp_ports = {'163.com':25, 
                          'libratone.com.cn':25}
        self.smtp_ssl_ports = {'163.com':465, 
                          'libratone.com.cn':465}        
        self.msg = MIMEMultipart('related')
        
    def add_html(self, html):
        text_msg = MIMEText(html, 'html', _charset='utf-8')
        self.msg.attach(text_msg)        
        pass
    
    def add_image(self, src, img_id):
        if not isinstance(self.msg, MIMEMultipart):
            print 'MIME message not a \'MIMEMultipart\' instance'
            return
        dst = src
        if src.find('~') == 0:
            dst = os.environ['HOME'] + src[1:]
            print 'Add picture: ', dst
        fp = open(dst, 'rb')
        msg_img = MIMEImage(fp.read())
        fp.close()
        msg_img.add_header('Content-ID', '<' + img_id.split('.')[0] + '>')
        print 'Content-ID: ' + '<' + img_id.split('.')[0] + '>'
        msg_img.add_header('Content-Type', 'image/png')
        self.msg.attach(msg_img)
        
    def set_sender(self, sender):
        self.sender = sender
        self.pwd = self.usr_pwd[sender]
        server = sender.split('@')[1]
        self.stmp_server = self.stmp_servers[server]
        if self.use_ssl:
            self.smtp_port = self.smtp_ssl_ports[server]
        else:
            self.smtp_port = self.smtp_ports[server]
            
    def set_cc(self, cc):
        self.msg['CC'] = cc
        
    def set_bcc(self, bcc):
        self.msg['BCC'] = bcc
        
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