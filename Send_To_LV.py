from Global import grabobjs
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate

import smtplib
import ssl
import os

CurrDir = os.path.dirname(os.path.abspath(__file__))
Global_Objs = grabobjs(CurrDir)


class SendToLV:
    server = None
    message = MIMEMultipart()

    def __init__(self):
        self.email_server = Global_Objs['Settings'].grab_item('email_server')
        self.email_port = Global_Objs['Settings'].grab_item('email_port')
        self.email_user = Global_Objs['Settings'].grab_item('email_user')
        self.email_pass = Global_Objs['Settings'].grab_item('email_pass')
        self.email_from = Global_Objs['Local_Settings'].grab_item('email_from')
        self.email_to = Global_Objs['Local_Settings'].grab_item('email_to')
        self.email_cc = Global_Objs['Local_Settings'].grab_item('email_cc')

    def email_connect(self):
        self.server = smtplib.SMTP_SSL(self.email_server, self.email_port)
        self.server.login(self.email_user, self.email_pass)

    def email_close(self):
        self.server.close()

    def package_email(self):
        self.message['From'] = self.email_from
        self.message['To'] = self.email_to
        self.message['Cc'] = self.email_cc
        self.message['Date'] = formatdate(localtime=True)
        self.message['Subject'] = self.email_cc


if __name__ == '__main__':
    if not Global_Objs['Settings'].grab_item('email_server'):
        Global_Objs['Settings'].add_item(key='email_server',
                                         inputmsg='Please provide the email server you would like to connect to:')

    if not Global_Objs['Settings'].grab_item('email_port'):
        Global_Objs['Settings'].add_item(key='email_port',
                                         inputmsg='Please provide the email port for the server. (Default is 465):')

    if not Global_Objs['Settings'].grab_item('email_user'):
        Global_Objs['Settings'].add_item(key='email_user', inputmsg='Please provide the user name for the email login:')

    if not Global_Objs['Settings'].grab_item('email_pass'):
        Global_Objs['Settings'].add_item(key='email_pass', inputmsg='Please provide the user pass for the email login:')



'''
# C:\Workspace\python\Bin
email_server = 'imail.granitenet.com'
email_port = 465
email_user = ''
email_pass = ''

context = ssl.create_defaul_context()

with smtplib.SMTP_SSL(email_server, email_port) as server:
    server.login(email_user, email_pass)
    server.sendmail()
'''