from Global import grabobjs
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

import smtplib
import zipfile
import os
import pandas as pd
import datetime

CurrDir = os.path.dirname(os.path.abspath(__file__))
BatchedDir = os.path.join(CurrDir, '02_Batched')
Global_Objs = grabobjs(CurrDir)
email_subject = 'Send To LV Batch'
email_message = 'Hello LV,\n\nPlease see the attached Send To LV Batch.\n\nIf you have any questions, please reach-out to {0} for more information.'
email_message2 = 'Hello LV,\n\nThere is no Send To LV Batch.\n\nIf you have any questions, please reach-out to {0} for more information.'


class EmailLV:
    server = None
    message = MIMEMultipart()

    def __init__(self, file):
        self.file = file
        self.email_server = Global_Objs['Settings'].grab_item('email_server')
        self.email_port = Global_Objs['Settings'].grab_item('email_port')
        self.email_user = Global_Objs['Settings'].grab_item('email_user')
        self.email_pass = Global_Objs['Settings'].grab_item('email_pass')
        self.email_from = Global_Objs['Local_Settings'].grab_item('email_from')
        self.email_to = Global_Objs['Local_Settings'].grab_item('email_to')
        self.email_cc = Global_Objs['Local_Settings'].grab_item('email_cc')

    def email_connect(self):
        Global_Objs['Event_Log'].write_log('Connecting to Server {0} port {1}'.format(self.email_server,
                                                                                      self.email_port))

        self.server = smtplib.SMTP('imail.granitenet.com', 587)

        self.server.ehlo()
        self.server.starttls()
        self.server.ehlo()
        self.server.login(self.email_user, self.email_pass)

    def email_send(self):
        self.server.sendmail(self.email_from, self.email_to, str(self.message))

        Global_Objs['Event_Log'].write_log('Batch to LV has been sent')

    def email_close(self):
        self.server.close()

    def package_email(self):
        Global_Objs['Event_Log'].write_log('Packaging e-mail to be sent to LV')

        self.message['From'] = self.email_from
        self.message['To'] = self.email_to
        if self.email_cc:
            self.message['Cc'] = self.email_cc
        self.message['Date'] = formatdate(localtime=True)
        self.message['Subject'] = email_subject

        if self.file:
            self.message.attach(MIMEText(email_message.format(self.email_cc)))

            part = MIMEBase('application', "octet-stream")
            zip_filepath = self.zip_file()
            zf = open(zip_filepath, 'rb')
            part.set_payload(zf.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="{0}"'.format(os.path.basename(zip_filepath)))
            self.message.attach(part)
        else:
            self.message.attach(MIMEText(email_message2.format(self.email_cc)))

    def zip_file(self):
        zip_filepath = os.path.join(os.path.dirname(self.file),
                                    '{0}.zip'.format(os.path.splitext(os.path.basename(self.file))[0]))

        zip_file = zipfile.ZipFile(zip_filepath, mode='w')

        try:
            zip_file.write(self.file, os.path.basename(self.file))
        finally:
            zip_file.close()

        os.remove(self.file)
        return zip_filepath


class LVBatch:
    file = None
    data = pd.DataFrame()

    def __init__(self):
        self.asql = Global_Objs['SQL']
        self.asql.connect('alch')
        self.table = Global_Objs['Local_Settings'].grab_item('STLV_TBL')
        self.columns = Global_Objs['Local_Settings'].grab_item('STLV_TBL_Cols')

    def validate(self):
        data = self.asql.query('''
            SELECT
                1

            FROM {0}

            WHERE
                Batch is null
                    AND
                Is_Rejected is null
            '''.format(self.table))

        if data.empty:
            return True
        else:
            Global_Objs['Event_Log'].write_log(
                'There are {0} items that hasnt been reviewed by DART. Unable to batch to LV'.format(len(data)))

            return False

    def grab_batch(self):
        Global_Objs['Event_Log'].write_log('Grabbing items from {0} to batch to LV'.format(self.table))

        self.data = self.asql.query('''
            SELECT
                {0}
            
            FROM {1}
            
            WHERE
                Batch is null
                    AND
                Is_Rejected = 'N'
            '''.format(self.columns, self.table))

    def write_batch(self):
        if self.data.empty:
            Global_Objs['Event_Log'].write_log('No items were found to batch', 'warning')
            myinput = None

            while myinput:
                print('Would you like to send LV a no batch e-mail? (yes, no)')
                myinput = input()

                if myinput.lower() not in ['yes', 'no']:
                    myinput = None

            if myinput.lower() == 'yes':
                return True
            else:
                return False
        else:
            Global_Objs['Event_Log'].write_log('Found {0} items to batch. Proceeding to batch to excel'.format(
                len(self.data)
            ))
            self.file = os.path.join(BatchedDir, '{0}_LV-Batch.xlsx'.format(
                datetime.datetime.now().__format__("%Y%m%d")))

            with pd.ExcelWriter(self.file) as writer:
                self.data.to_excel(writer, index=False, sheet_name=datetime.datetime.now().__format__("%Y%m%d"))

            return True

    def send_batch(self):
        obj = EmailLV(self.file)
        obj.email_connect()

        try:
            obj.package_email()
            obj.email_send()

        finally:
            obj.email_close()

    def close_conn(self):
        self.asql.close()


def check_settings():
    if not Global_Objs['Settings'].grab_item('email_server'):
        Global_Objs['Settings'].add_item(key='email_server',
                                         inputmsg='Please provide the email server you would like to connect to:')

    if not Global_Objs['Settings'].grab_item('email_port'):
        Global_Objs['Settings'].add_item(key='email_port',
                                         inputmsg='Please provide the email port for the server. (Default is 465):')

    if not Global_Objs['Settings'].grab_item('email_user'):
        Global_Objs['Settings'].add_item(key='email_user',
                                         inputmsg='Please provide the user name for the email login:')

    if not Global_Objs['Settings'].grab_item('email_pass'):
        Global_Objs['Settings'].add_item(key='email_pass',
                                         inputmsg='Please provide the user pass for the email login:')

    if not Global_Objs['Local_Settings'].grab_item('email_from'):
        Global_Objs['Local_Settings'].add_item(key='email_from',
                                               inputmsg='Please provide your e-mail address:')

    if not Global_Objs['Local_Settings'].grab_item('email_to'):
        Global_Objs['Local_Settings'].add_item(key='email_to',
                                         inputmsg='Please provide the e-mail address to send e-mail:')

    if not Global_Objs['Local_Settings'].grab_item('email_cc'):
        Global_Objs['Local_Settings'].add_item(key='email_cc',
                                               inputmsg='Please provide a cc to include in the e-mail:')

    if not Global_Objs['Local_Settings'].grab_item('STLV_TBL'):
        Global_Objs['Local_Settings'].add_item(key='STLV_TBL',
                                               inputmsg='Please provide the Send To LV SQL table:')

    if not Global_Objs['Local_Settings'].grab_item('STLV_TBL_Cols'):
        Global_Objs['Local_Settings'].add_item(key='STLV_TBL_Cols',
                                               inputmsg='Please provide the columns for the Send To LV SQL table:')


if __name__ == '__main__':
    check_settings()

    if not os.path.exists(BatchedDir):
        os.makedirs(BatchedDir)

    myobj = LVBatch()

    try:
        if myobj.validate():
            myobj.grab_batch()

            if myobj.write_batch():
                myobj.send_batch()

    finally:
        myobj.close_conn()