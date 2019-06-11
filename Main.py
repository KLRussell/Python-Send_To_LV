from Global import grabobjs
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from Settings import SettingsGUI

import smtplib
import zipfile
import os
import pandas as pd
import datetime

curr_dir = os.path.dirname(os.path.abspath(__file__))
main_dir = os.path.dirname(curr_dir)
batcheddir = os.path.join(main_dir, '02_Batched')
global_objs = grabobjs(main_dir, 'Send_To_LV')
email_subject = 'Send To LV Batch'
email_message = 'Hello LV,\n\nPlease see the attached Send To LV Batch.\n\nIf you have any questions, please reach-out to {0} for more information.'
email_message2 = 'Hello LV,\n\nThere is no Send To LV Batch.\n\nIf you have any questions, please reach-out to {0} for more information.'


class EmailLV:
    zip_filepath = None
    server = None
    message = MIMEMultipart()

    def __init__(self, file):
        self.file = file
        self.email_server = global_objs['Settings'].grab_item('Email_Server')
        self.email_port = global_objs['Settings'].grab_item('Email_Port')
        self.email_user = global_objs['Settings'].grab_item('Email_User')
        self.email_pass = global_objs['Settings'].grab_item('Email_Pass')
        self.email_from = global_objs['Local_Settings'].grab_item('Email_From')
        self.email_to = global_objs['Local_Settings'].grab_item('Email_To')
        self.email_cc = global_objs['Local_Settings'].grab_item('Email_Cc')

    def email_connect(self):
        global_objs['Event_Log'].write_log('Connecting to Server {0} port {1}'.format(self.email_server.decrypt_text(),
                                                                                      self.email_port.decrypt_text()))

        self.server = smtplib.SMTP(str(self.email_server.decrypt_text()), int(self.email_port.decrypt_text()))

        self.server.ehlo()
        self.server.starttls()
        self.server.ehlo()
        self.server.login(self.email_user.decrypt_text(), self.email_pass.decrypt_text())

    def email_send(self):
        self.server.sendmail(self.email_from.decrypt_text(), self.email_to.decrypt_text(), str(self.message))

        global_objs['Event_Log'].write_log('Batch to LV has been sent')

    def email_close(self):
        self.server.close()

    def package_email(self):
        global_objs['Event_Log'].write_log('Packaging e-mail to be sent to LV')

        self.message['From'] = self.email_from.decrypt_text()
        self.message['To'] = self.email_to.decrypt_text()
        self.message['Cc'] = self.email_cc.decrypt_text()
        self.message['Date'] = formatdate(localtime=True)
        self.message['Subject'] = email_subject

        if self.file:
            self.message.attach(MIMEText(email_message.format(self.email_cc.decrypt_text())))

            part = MIMEBase('application', "octet-stream")
            self.zip_file()
            zf = open(self.zip_filepath, 'rb')

            try:
                part.set_payload(zf.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename="{0}"'.format(
                    os.path.basename(self.zip_filepath)))
                self.message.attach(part)
            finally:
                zf.close()
        else:
            self.message.attach(MIMEText(email_message2.format(self.email_cc.decrypt_text())))

    def zip_file(self):
        i = 1

        while not self.zip_filepath:
            if i > 1:
                self.zip_filepath = os.path.join(os.path.dirname(self.file),
                                                 '{0}{1}.zip'.format(
                                                     os.path.splitext(os.path.basename(self.file))[0], i))
            else:
                self.zip_filepath = os.path.join(os.path.dirname(self.file),
                                                 '{0}.zip'.format(os.path.splitext(os.path.basename(self.file))[0]))

            if os.path.exists(self.zip_filepath):
                i += 1
                self.zip_filepath = None

        zip_file = zipfile.ZipFile(self.zip_filepath, mode='w')

        try:
            zip_file.write(self.file, os.path.basename(self.file))
        finally:
            zip_file.close()
            os.remove(self.file)


class LVBatch:
    file = None
    data = pd.DataFrame()

    def __init__(self):
        self.asql = global_objs['SQL']
        self.asql.connect('alch')
        self.table = global_objs['Local_Settings'].grab_item('Stlv_Tbl')
        self.batch_table = global_objs['Local_Settings'].grab_item('Stlv_Batch_Tbl')
        self.columns = self.sql_columns()
        self.source_table = global_objs['Local_Settings'].grab_item('Source_Tbl')
        self.cat_table = global_objs['Local_Settings'].grab_item('Cat_Tbl')

    @staticmethod
    def sql_columns():
        cols = list(global_objs['Local_Settings'].grab_item('Stlv_Tbl_Cols'))

        for index, col in enumerate(cols):
            if col == 'Assigned_To':
                cols[index] = 'CAT.Full_Name CAT_Rep'
            elif col == 'Logged_By':
                cols[index] = 'CAT.Full_Name Logged_By_Rep'
            elif col == 'ST_ID':
                cols[index] = 'ST.Source_TBL'
            elif col == 'BD_ID':
                cols[index] = 'CAST(GETDATE() AS DATE) Batch_DT'
            elif col == 'Back_Bill_Info':
                cols[index] = "CASE WHEN Back_Bill_Info = 1 THEN 'Yes' ELSE 'No' END Back_Bill_Info"

        return cols

    def validate(self):
        data = self.asql.query('''
            SELECT
                1

            FROM {0}

            WHERE
                BD_ID is null
                    AND
                Is_Rejected is null
            '''.format(self.table.decrypt_text()))

        if data.empty:
            return True
        else:
            global_objs['Event_Log'].write_log(
                'There are {0} items that hasnt been reviewed by DART. Unable to batch to LV'.format(len(data)))
            return False

    def grab_batch(self):
        global_objs['Event_Log'].write_log('Grabbing items from {0} to batch to LV'.format(self.table.decrypt_text()))

        self.data = self.asql.query('''
            SELECT
                {0}
            
            FROM {1} HR
            INNER JOIN {2} ST
            ON
                HR.ST_ID = ST.ST_ID
            INNER JOIN {3} CAT
            ON
                HR.Assigned_To = CAT.CAT_ID
            
            WHERE
                HR.BD_ID is null
                    AND
                HR.Is_Rejected = 0
            '''.format(', '.join(self.columns), self.table.decrypt_text(),
                       self.source_table.decrypt_text(), self.cat_table.decrypt_text()))

    def write_batch(self):
        if self.data.empty:
            global_objs['Event_Log'].write_log('No items were found to batch', 'warning')
            myinput = None

            while not myinput:
                print('Would you like to send LV a no batch e-mail? (yes, no)')
                myinput = input()

                if myinput.lower() not in ['yes', 'no']:
                    myinput = None

            if myinput.lower() == 'yes':
                return True
            else:
                return False
        else:
            global_objs['Event_Log'].write_log('Found {0} items to batch. Proceeding to batch to excel'.format(
                len(self.data)
            ))
            i = 1

            while not self.file:
                if i > 1:
                    self.file = os.path.join(batcheddir, '{0}_LV-Batch{1}.xlsx'.format(
                        datetime.datetime.now().__format__("%Y%m%d"), i))
                else:
                    self.file = os.path.join(batcheddir, '{0}_LV-Batch.xlsx'.format(
                        datetime.datetime.now().__format__("%Y%m%d")))

                if os.path.exists(self.file):
                    self.file = None
                    i += 1

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

    def update_tbl(self):
        if not self.data.empty:
            self.asql.upload(self.data['STLV_ID'], 'LV_Tmp')

            self.asql.execute('''
                INSERT INTO {0}
                (
                    Batch_DT
                )
                SELECT
                    GETDATE()
                WHERE
                    NOT EXISTS
                    (
                        SELECT
                            1
                            
                        FROM {0}
                        
                        WHERE
                            BATCH_DT = GETDATE()
                    )
            '''.format(self.batch_table.decrypt_table()))

            batch_date = self.asql.query('''
                SELECT
                    BD_ID
                    
                FROM {0}
                
                WHERE
                    Batch_DT = GETDATE()
            '''.format(self.batch_table.decrypt_table()))

            if not batch_date.empty:
                self.asql.execute('''
                UPDATE A
                    SET
                        A.BD_ID = 
                
                FROM {0} As A
                INNER JOIN LV_Tmp As B
                ON
                    A.HR_ID = B.HR_ID
                '''.format(self.table.decrypt_text(), batch_date['BD_ID'][0]))

                self.asql.execute('DROP TABLE LV_Tmp')

    def close_conn(self):
        self.asql.close()


def check_settings():
    my_return = False
    obj = SettingsGUI()

    if not os.path.exists(batcheddir):
        os.makedirs(batcheddir)

    if not global_objs['Settings'].grab_item('Server') \
            or not global_objs['Settings'].grab_item('Database') \
            or not global_objs['Settings'].grab_item('Email_Server') \
            or not global_objs['Settings'].grab_item('Email_Port') \
            or not global_objs['Settings'].grab_item('Email_User') \
            or not global_objs['Settings'].grab_item('Email_Pass') \
            or not global_objs['Local_Settings'].grab_item('Email_From') \
            or not global_objs['Local_Settings'].grab_item('Email_To') \
            or not global_objs['Local_Settings'].grab_item('Email_Cc') \
            or not global_objs['Local_Settings'].grab_item('Stlv_Tbl') \
            or not global_objs['Local_Settings'].grab_item('Stlv_Batch_Tbl') \
            or not global_objs['Local_Settings'].grab_item('Source_Tbl') \
            or not global_objs['Local_Settings'].grab_item('Cat_Tbl'):
        header_text = 'Welcome to Send to LV!\nSettings haven''t been established.\nPlease fill out all empty fields below:'
        obj.build_gui(header_text)
    else:
        try:
            if not obj.sql_connect():
                header_text = 'Welcome to Send to LV!\nNetwork settings are invalid.\nPlease fix the network settings below:'
                obj.build_gui(header_text)
            else:
                email_err = 0

                try:
                    server = smtplib.SMTP(str(global_objs['Settings'].grab_item('Email_Server').decrypt_text()),
                                          int(global_objs['Settings'].grab_item('Email_Port').decrypt_text()))

                    try:
                        server.ehlo()
                        server.starttls()
                        server.ehlo()
                        server.login(global_objs['Settings'].grab_item('Email_User').decrypt_text(),
                                     global_objs['Settings'].grab_item('Email_Pass').decrypt_text())
                    except:
                        email_err = 2
                        pass
                    finally:
                        server.close()
                except:
                    email_err = 1
                    pass

                if email_err == 1:
                    header_text = 'Welcome to Send to LV!\nServer and/or Port does not exist.\nPlease specify the correct information:'
                    obj.build_gui(header_text)
                elif email_err == 2:
                    header_text = 'Welcome to Send to LV!\nUser name and/or user pass is invalid.\nPlease specify the correct information:'
                    obj.build_gui(header_text)
                else:
                    my_return = True
        finally:
            obj.sql_close()

    obj.cancel()
    del obj

    return my_return


if __name__ == '__main__':
    if check_settings():
        myobj = LVBatch()
        try:
            if myobj.validate():
                myobj.grab_batch()

                if myobj.write_batch():
                    myobj.send_batch()
                    myobj.update_tbl()

        finally:
            myobj.close_conn()
    else:
        global_objs['Event_Log'].write_log('Settings Mode was established. Need to re-run script', 'warning')

    os.system('pause')
