# Global Module import
from tkinter import *
from tkinter import messagebox
from Global import grabobjs
from Global import CryptHandle

import smtplib
import pandas as pd
import os
import sys

if getattr(sys, 'frozen', False):
    application_path = sys.executable
else:
    application_path = __file__

# Global Variable declaration
curr_dir = os.path.dirname(os.path.abspath(application_path))
main_dir = os.path.dirname(curr_dir)
global_objs = grabobjs(main_dir, 'SendToLV')


class SettingsGUI:
    save_settings_button = None
    stlv_list_box = None
    stlvs_list_box = None
    move_right_all_button = None
    move_right_button = None
    move_left_all_button = None
    move_left_button = None
    stlv_sqltbl_txtbox = None
    ecc_txtbox = None
    eto_txtbox = None
    efrom_txtbox = None
    eupass_txtbox = None
    euname_txtbox = None
    eport_txtbox = None
    eserver_txtbox = None
    stlv_sqltbl2_txtbox = None
    cnr_source_table_txtbox = None
    cat_table_txtbox = None
    sql_tables = pd.DataFrame()
    stlv_list_sel = 0
    stlvs_list_sel = 0

    # Function that is executed upon creation of SettingsGUI class
    def __init__(self):
        self.header_text = 'Welcome to Send to LV Settings!\nSettings can be changed below.\nPress save when finished'

        self.email_upass_obj = global_objs['Settings'].grab_item('Email_Pass')
        self.asql = global_objs['SQL']
        self.main = Tk()

        # GUI Variables
        self.server = StringVar()
        self.database = StringVar()
        self.email_server = StringVar()
        self.email_port = StringVar()
        self.email_user_name = StringVar()
        self.email_user_pass = StringVar()
        self.email_from = StringVar()
        self.email_to = StringVar()
        self.email_cc = StringVar()
        self.sql_table = StringVar()
        self.sql_batch = StringVar()
        self.source_table = StringVar()
        self.cat_table = StringVar()

        self.main.bind('<Destroy>', self.gui_cleanup)

    def store_sql_tables(self):
        self.sql_tables = self.asql.query('''
            select
                concat(table_schema, '.', table_name) SQL_TBL

            from information_schema.tables''')

    def gui_cleanup(self, event):
        self.asql.close()

    # Static function to fill textbox in GUI
    @staticmethod
    def fill_textbox(setting_list, val, key):
        assert(key and val and setting_list)
        item = global_objs[setting_list].grab_item(key)

        if isinstance(item, CryptHandle):
            val.set(item.decrypt_text())

    # static function to add setting to Local_Settings shelf files
    @staticmethod
    def add_setting(setting_list, val, key, encrypt=True):
        assert (key and setting_list)

        global_objs[setting_list].del_item(key)

        if val:
            global_objs[setting_list].add_item(key=key, val=val, encrypt=encrypt)

        global_objs[setting_list].write_shelf()

    # Function to build GUI for settings
    def build_gui(self, header=None):
        # Change to custom header title if specified
        if header:
            self.header_text = header

        # Set GUI Geometry and GUI Title
        self.main.geometry('509x755+500+50')
        self.main.title('Send to LV Settings')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        network_frame = LabelFrame(self.main, text='Network Settings', width=508, height=70)
        email_frame = LabelFrame(self.main, text='E-mail', width=508, height=170)
        econn_frame = LabelFrame(email_frame, text='Settings', width=252, height=170)
        emsg_frame = LabelFrame(email_frame, text='Message', width=253, height=170)
        stlv_frame = LabelFrame(self.main, text='Send To LV', width=508, height=200)
        stlv_list_frame = LabelFrame(stlv_frame, text='SQL Table Columns', width=227, height=200)
        stlv_list_frame2 = Frame(stlv_frame, width=50, height=200)
        stlv_list_frame3 = LabelFrame(stlv_frame, text='Table Columns Selected', width=228, height=200)
        cat_table_frame = LabelFrame(stlv_frame, text='SQL CAT Emp Table', width=505, height=70)
        cnr_source_table_frame = LabelFrame(stlv_frame, text='SQL CNR Source Table', width=505, height=70)
        stlv_batch_frame = LabelFrame(stlv_frame, text='SQL Batch Table', width=505, height=70)
        stlv_settings_frame = LabelFrame(stlv_frame, text='SQL History Record Table', width=505, height=70)
        buttons_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        network_frame.pack(fill="both")
        email_frame.pack(fill="both")
        econn_frame.grid(row=0, column=0, ipady=5)
        emsg_frame.grid(row=0, column=1, ipady=20)
        stlv_frame.pack(fill="both")
        cat_table_frame.grid(row=0, columnspan=3, ipady=2)
        cnr_source_table_frame.grid(row=1, columnspan=3, ipady=2)
        stlv_batch_frame.grid(row=2, columnspan=3, ipady=2)
        stlv_settings_frame.grid(row=3, columnspan=3, ipady=2)
        stlv_list_frame.grid(row=4, column=0, rowspan=4, ipady=2, sticky='e')
        stlv_list_frame2.grid(row=4, column=1, rowspan=4, ipady=2)
        stlv_list_frame3.grid(row=4, column=2, rowspan=4, ipady=2, sticky='w')
        buttons_frame.pack(fill='both')

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=self.header_text, width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Network Labels & Input boxes to the Network_Frame
        #     SQL Server Input Box
        server_label = Label(self.main, text='Server:', padx=15, pady=7)
        server_txtbox = Entry(self.main, textvariable=self.server)
        server_label.pack(in_=network_frame, side=LEFT)
        server_txtbox.pack(in_=network_frame, side=LEFT)
        server_txtbox.bind('<FocusOut>', self.check_network)

        #     Server Database Input Box
        database_label = Label(self.main, text='Database:')
        database_txtbox = Entry(self.main, textvariable=self.database)
        database_txtbox.pack(in_=network_frame, side=RIGHT, pady=7, padx=15)
        database_label.pack(in_=network_frame, side=RIGHT)
        database_txtbox.bind('<KeyRelease>', self.check_network)

        # Apply Email Labels & Input boxes to the EConn_Frame
        #     Email Server Input Box
        eserver_label = Label(econn_frame, text='Email Server:')
        self.eserver_txtbox = Entry(econn_frame, textvariable=self.email_server)
        eserver_label.grid(row=0, column=0, padx=8, pady=5, sticky='w')
        self.eserver_txtbox.grid(row=0, column=1, padx=13, pady=5, sticky='e')

        #     Email Port Input Box
        eport_label = Label(econn_frame, text='Email Port:')
        self.eport_txtbox = Entry(econn_frame, textvariable=self.email_port)
        eport_label.grid(row=1, column=0, padx=8, pady=5, sticky='w')
        self.eport_txtbox.grid(row=1, column=1, padx=13, pady=5, sticky='e')

        #     Email User Name Input Box
        euname_label = Label(econn_frame, text='Email User Name:')
        self.euname_txtbox = Entry(econn_frame, textvariable=self.email_user_name)
        euname_label.grid(row=2, column=0, padx=8, pady=5, sticky='w')
        self.euname_txtbox.grid(row=2, column=1, padx=13, pady=5, sticky='e')

        #     Email User Pass Input Box
        eupass_label = Label(econn_frame, text='Email User Pass:')
        self.eupass_txtbox = Entry(econn_frame, textvariable=self.email_user_pass)
        eupass_label.grid(row=3, column=0, padx=8, pady=5, sticky='w')
        self.eupass_txtbox.grid(row=3, column=1, padx=13, pady=5, sticky='e')
        self.eupass_txtbox.bind('<KeyRelease>', self.hide_pass)

        # Apply Email Labels & Input boxes to the EMsg_Frame
        #     From Email Address Input Box
        efrom_label = Label(emsg_frame, text='From Addr:')
        self.efrom_txtbox = Entry(emsg_frame, textvariable=self.email_from)
        efrom_label.grid(row=0, column=0, padx=8, pady=5, sticky='w')
        self.efrom_txtbox.grid(row=0, column=1, padx=13, pady=5, sticky='e')

        #     To Email Address Input Box
        eto_label = Label(emsg_frame, text='To Addr:')
        self.eto_txtbox = Entry(emsg_frame, textvariable=self.email_to)
        eto_label.grid(row=1, column=0, padx=8, pady=5, sticky='w')
        self.eto_txtbox.grid(row=1, column=1, padx=13, pady=5, sticky='e')

        #     CC Email Address Input Box
        ecc_label = Label(emsg_frame, text='CC Addr:')
        self.ecc_txtbox = Entry(emsg_frame, textvariable=self.email_cc)
        ecc_label.grid(row=2, column=0, padx=8, pady=5, sticky='w')
        self.ecc_txtbox.grid(row=2, column=1, padx=13, pady=5, sticky='e')

        # Apply Line Verification SQL Table Label & Input boxes to the EMsg_Frame
        self.cat_table_txtbox = Entry(cat_table_frame, textvariable=self.cat_table, width=76)
        self.cat_table_txtbox.grid(row=0, columnspan=3, padx=20, pady=5)

        # Apply Line Verification SQL Table Label & Input boxes to the EMsg_Frame
        self.cnr_source_table_txtbox = Entry(cnr_source_table_frame, textvariable=self.source_table, width=76)
        self.cnr_source_table_txtbox.grid(row=0, columnspan=3, padx=20, pady=5)

        #     Send to LV Batch SQL TBL Input Box
        self.stlv_sqltbl2_txtbox = Entry(stlv_batch_frame, textvariable=self.sql_batch, width=76)
        self.stlv_sqltbl2_txtbox.grid(row=0, columnspan=3, padx=20, pady=5)

        #     Send to LV History Records SQL TBL Input Box
        self.stlv_sqltbl_txtbox = Entry(stlv_settings_frame, textvariable=self.sql_table, width=76)
        self.stlv_sqltbl_txtbox.grid(row=0, columnspan=3, padx=20, pady=5)
        self.stlv_sqltbl_txtbox.bind('<KeyRelease>', self.check_table_event)

        # Send to LV SQL TBL Columns to the Stlv_List_Frame
        stlv_yscrollbar = Scrollbar(stlv_list_frame, orient="vertical")
        stlv_xscrollbar = Scrollbar(stlv_list_frame, orient="horizontal")
        self.stlv_list_box = Listbox(stlv_list_frame, selectmode=SINGLE, width=25, yscrollcommand=stlv_yscrollbar.set,
                                     xscrollcommand=stlv_xscrollbar.set)
        stlv_yscrollbar.config(command=self.stlv_list_box.yview)
        stlv_xscrollbar.config(command=self.stlv_list_box.xview)
        self.stlv_list_box.grid(row=0, column=0, rowspan=4, padx=8, pady=5)
        stlv_yscrollbar.grid(row=0, column=1, rowspan=4, sticky=N + S)
        stlv_xscrollbar.grid(row=4, column=0, columnspan=2, sticky=E + W)
        self.stlv_list_box.bind("<Down>", self.stlv_list_down)
        self.stlv_list_box.bind("<Up>", self.stlv_list_up)
        self.stlv_list_box.bind('<<ListboxSelect>>', self.stlv_list_select)

        # Apply Migration buttons for both list boxes for Frame STLV List Frame
        self.move_right_all_button = Button(stlv_list_frame2, text='>>', width=4, command=self.move_right_all)
        self.move_right_all_button.grid(row=0, column=1, pady=10)
        self.move_right_button = Button(stlv_list_frame2, text='>', width=4, command=self.move_right)
        self.move_right_button.grid(row=1, column=1, pady=5)
        self.move_left_all_button = Button(stlv_list_frame2, text='<', width=4, command=self.move_left)
        self.move_left_all_button.grid(row=2, column=1, pady=5)
        self.move_left_button = Button(stlv_list_frame2, text='<<', width=4, command=self.move_left_all)
        self.move_left_button.grid(row=3, column=1, pady=5)

        # Send to LV SQL TBL Selection Columns to the Stlv_List_Frame
        stlvs_yscrollbar = Scrollbar(stlv_list_frame3, orient="vertical")
        stlvs_xscrollbar = Scrollbar(stlv_list_frame3, orient="horizontal")
        self.stlvs_list_box = Listbox(stlv_list_frame3, selectmode=SINGLE, width=25,
                                      yscrollcommand=stlvs_yscrollbar.set, xscrollcommand=stlvs_xscrollbar.set)
        stlvs_yscrollbar.config(command=self.stlvs_list_box.yview)
        stlvs_xscrollbar.config(command=self.stlvs_list_box.xview)
        self.stlvs_list_box.grid(row=0, column=0, rowspan=4, padx=8, pady=5)
        stlvs_yscrollbar.grid(row=0, column=1, rowspan=4, sticky=N + S)
        stlvs_xscrollbar.grid(row=4, column=0, columnspan=2, sticky=E + W)
        self.stlvs_list_box.bind("<Down>", self.stlvs_list_down)
        self.stlvs_list_box.bind("<Up>", self.stlvs_list_up)
        self.stlvs_list_box.bind('<<ListboxSelect>>', self.stlvs_list_select)

        # Apply Buttons to the Buttons Frame
        #     Save Button
        self.save_settings_button = Button(buttons_frame, text='Save Settings', width=20, command=self.save_settings)
        self.save_settings_button.grid(row=0, column=0, pady=6, padx=15)

        #     Cancel Button
        cancel_button = Button(buttons_frame, text='Cancel', width=20, command=self.cancel)
        cancel_button.grid(row=0, column=1, pady=6, padx=165)

        self.fill_gui()

        # Show dialog
        self.main.mainloop()

    # Function to fill GUI textbox fields
    def fill_gui(self):
        self.fill_textbox('Settings', self.server, 'Server')
        self.fill_textbox('Settings', self.database, 'Database')

        if not self.server.get() or not self.database.get() or not self.asql.test_conn('alch'):
            self.save_settings_button.configure(state=DISABLED)
            self.stlv_list_box.configure(state=DISABLED)
            self.stlvs_list_box.configure(state=DISABLED)
            self.move_right_all_button.configure(state=DISABLED)
            self.move_right_button.configure(state=DISABLED)
            self.move_left_all_button.configure(state=DISABLED)
            self.move_left_button.configure(state=DISABLED)
            self.stlv_sqltbl_txtbox.configure(state=DISABLED)
            self.cnr_source_table_txtbox.config(state=DISABLED)
            self.cat_table_txtbox.config(state=DISABLED)
            self.ecc_txtbox.configure(state=DISABLED)
            self.eto_txtbox.configure(state=DISABLED)
            self.efrom_txtbox.configure(state=DISABLED)
            self.eupass_txtbox.configure(state=DISABLED)
            self.euname_txtbox.configure(state=DISABLED)
            self.eport_txtbox.configure(state=DISABLED)
            self.eserver_txtbox.configure(state=DISABLED)
        else:
            self.asql.connect('alch')
            self.store_sql_tables()
            self.fill_textbox('Settings', self.email_server, 'Email_Server')
            self.fill_textbox('Settings', self.email_port, 'Email_Port')
            self.fill_textbox('Settings', self.email_user_name, 'Email_User')
            self.fill_textbox('Local_Settings', self.email_from, 'Email_From')
            self.fill_textbox('Local_Settings', self.email_to, 'Email_To')
            self.fill_textbox('Local_Settings', self.email_cc, 'Email_Cc')
            self.fill_textbox('Local_Settings', self.sql_table, 'Stlv_Tbl')
            self.fill_textbox('Local_Settings', self.sql_batch, 'Stlv_Batch_Tbl')
            self.fill_textbox('Local_Settings', self.source_table, 'Source_Tbl')
            self.fill_textbox('Local_Settings', self.cat_table, 'Cat_Tbl')

            if not self.email_server.get():
                self.email_server.set('imail.granitenet.com')

            if not self.email_port.get():
                self.email_port.set('587')

            if self.email_upass_obj and isinstance(self.email_upass_obj, CryptHandle):
                self.email_user_pass.set('*' * len(self.email_upass_obj.decrypt_text()))

            if self.check_table(self.sql_table.get()):
                self.populate_lists(self.sql_table.get(), global_objs['Local_Settings'].grab_item('Stlv_Tbl_Cols'))
            else:
                self.sql_table.set('')
                self.stlv_list_box.configure(state=DISABLED)
                self.stlvs_list_box.configure(state=DISABLED)
                self.move_right_all_button.configure(state=DISABLED)
                self.move_right_button.configure(state=DISABLED)
                self.move_left_all_button.configure(state=DISABLED)
                self.move_left_button.configure(state=DISABLED)

    def check_table(self, table):
        return not self.sql_tables.empty and table and\
               len(self.sql_tables[self.sql_tables['SQL_TBL'].str.lower() == table.lower()]) > 0

    def populate_lists(self, table, cols=None):
        mytbl = table.split('.')
        true_cols = self.asql.query('''
                    select
                        Column_Name
        
                from INFORMATION_SCHEMA.COLUMNS
        
                where
                    TABLE_SCHEMA = '{0}'
                        and
                    TABLE_NAME = '{1}'
                '''.format(mytbl[0], mytbl[1]))

        if not true_cols.empty:
            if not true_cols.empty and cols and (isinstance(cols, list) or isinstance(cols, tuple)):
                for col in true_cols['Column_Name'].tolist():
                    found = False

                    for col2 in cols:
                        if col.lower() == col2.lower():
                            found = True
                            break

                    if found:
                        self.stlvs_list_box.insert('end', col)
                    else:
                        self.stlv_list_box.insert('end', col)
            elif not true_cols.empty:
                for col in true_cols['Column_Name'].tolist():
                    self.stlv_list_box.insert('end', col)
        else:
            messagebox.showerror('No Columns Error!', 'Table has no columns in SQL Server')

    def hide_pass(self, event):
        if self.email_upass_obj:
            currpass = self.email_upass_obj.decrypt_text()

            if len(self.email_user_pass.get()) > len(currpass):
                i = 0

                for letter in self.email_user_pass.get():
                    if letter != '*':
                        if i > len(currpass) - 1:
                            currpass += letter
                        else:
                            mytext = list(currpass)
                            mytext.insert(i, letter)
                            currpass = ''.join(mytext)
                    i += 1
            elif len(self.email_user_pass.get()) > 0:
                i = 0

                for letter in self.email_user_pass.get():
                    if letter != '*':
                        mytext = list(currpass)
                        mytext[i] = letter
                        currpass = ''.join(mytext)
                    i += 1

                if len(currpass) - i > 0:
                    currpass = currpass[:i]
            else:
                currpass = None

            if currpass:
                self.email_upass_obj.encrypt_text(currpass)
                self.email_user_pass.set('*' * len(self.email_upass_obj.decrypt_text()))
            else:
                self.email_upass_obj = None
                self.email_user_pass.set("")
        else:
            self.email_upass_obj = CryptHandle()
            self.email_upass_obj.encrypt_text(self.email_user_pass.get())
            self.email_user_pass.set('*' * len(self.email_upass_obj.decrypt_text()))

    def check_table_event(self, event):
        if self.stlv_list_box.size() > 0:
            self.stlv_list_sel = 0
            self.stlv_list_box.delete(0, self.stlv_list_box.size() - 1)

        if self.stlvs_list_box.size() > 0:
            self.stlvs_list_sel = 0
            self.stlvs_list_box.delete(0, self.stlvs_list_box.size() - 1)

        if self.check_table(self.sql_table.get()):
            if str(self.stlv_list_box['state']) != 'normal':
                self.stlv_list_box.configure(state=NORMAL)
                self.stlvs_list_box.configure(state=NORMAL)
                self.move_right_all_button.configure(state=NORMAL)
                self.move_right_button.configure(state=NORMAL)
                self.move_left_all_button.configure(state=NORMAL)
                self.move_left_button.configure(state=NORMAL)

            self.populate_lists(self.sql_table.get())
        elif str(self.stlv_list_box['state']) != 'disabled':
            self.stlv_list_box.configure(state=DISABLED)
            self.stlvs_list_box.configure(state=DISABLED)
            self.move_right_all_button.configure(state=DISABLED)
            self.move_right_button.configure(state=DISABLED)
            self.move_left_all_button.configure(state=DISABLED)
            self.move_left_button.configure(state=DISABLED)

    # Function to check network settings if populated
    def check_network(self, event):
        if self.server.get() and self.database.get() and \
                (global_objs['Settings'].grab_item('Server') != self.server.get() or
                 global_objs['Settings'].grab_item('Database') != self.database.get()):
            self.asql.change_config(server=self.server.get(), database=self.database.get())

            if self.asql.test_conn('alch'):
                self.save_settings_button.configure(state=NORMAL)
                self.stlv_list_box.configure(state=NORMAL)
                self.stlvs_list_box.configure(state=NORMAL)
                self.move_right_all_button.configure(state=NORMAL)
                self.move_right_button.configure(state=NORMAL)
                self.move_left_all_button.configure(state=NORMAL)
                self.move_left_button.configure(state=NORMAL)
                self.stlv_sqltbl_txtbox.configure(state=NORMAL)
                self.cnr_source_table_txtbox.config(state=NORMAL)
                self.cat_table_txtbox.config(state=NORMAL)
                self.ecc_txtbox.configure(state=NORMAL)
                self.eto_txtbox.configure(state=NORMAL)
                self.efrom_txtbox.configure(state=NORMAL)
                self.eupass_txtbox.configure(state=NORMAL)
                self.euname_txtbox.configure(state=NORMAL)
                self.eport_txtbox.configure(state=NORMAL)
                self.eserver_txtbox.configure(state=NORMAL)
                self.add_setting('Settings', self.server.get(), 'Server')
                self.add_setting('Settings', self.database.get(), 'Database')
                self.asql.connect('alch')
                self.store_sql_tables()
            else:
                self.save_settings_button.configure(state=DISABLED)
                self.stlv_list_box.configure(state=DISABLED)
                self.stlvs_list_box.configure(state=DISABLED)
                self.move_right_all_button.configure(state=DISABLED)
                self.move_right_button.configure(state=DISABLED)
                self.move_left_all_button.configure(state=DISABLED)
                self.move_left_button.configure(state=DISABLED)
                self.stlv_sqltbl_txtbox.configure(state=DISABLED)
                self.cnr_source_table_txtbox.config(state=DISABLED)
                self.cat_table_txtbox.config(state=DISABLED)
                self.ecc_txtbox.configure(state=DISABLED)
                self.eto_txtbox.configure(state=DISABLED)
                self.efrom_txtbox.configure(state=DISABLED)
                self.eupass_txtbox.configure(state=DISABLED)
                self.euname_txtbox.configure(state=DISABLED)
                self.eport_txtbox.configure(state=DISABLED)
                self.eserver_txtbox.configure(state=DISABLED)
        else:
            self.save_settings_button.configure(state=DISABLED)
            self.stlv_list_box.configure(state=DISABLED)
            self.stlvs_list_box.configure(state=DISABLED)
            self.move_right_all_button.configure(state=DISABLED)
            self.move_right_button.configure(state=DISABLED)
            self.move_left_all_button.configure(state=DISABLED)
            self.move_left_button.configure(state=DISABLED)
            self.stlv_sqltbl_txtbox.configure(state=DISABLED)
            self.cnr_source_table_txtbox.config(state=DISABLED)
            self.cat_table_txtbox.config(state=DISABLED)
            self.ecc_txtbox.configure(state=DISABLED)
            self.eto_txtbox.configure(state=DISABLED)
            self.efrom_txtbox.configure(state=DISABLED)
            self.eupass_txtbox.configure(state=DISABLED)
            self.euname_txtbox.configure(state=DISABLED)
            self.eport_txtbox.configure(state=DISABLED)
            self.eserver_txtbox.configure(state=DISABLED)

    # Function adjusts selection of item when user clicks item (STLV List)
    def stlv_list_select(self, event):
        if self.stlv_list_box and self.stlv_list_box.curselection() \
                and -1 < self.stlv_list_sel < self.stlv_list_box.size() - 1:
            self.stlv_list_sel = self.stlv_list_box.curselection()[0]

    # Function adjusts selection of item when user presses down key (STLV List)
    def stlv_list_down(self, event):
        if self.stlv_list_sel < self.stlv_list_box.size() - 1:
            self.stlv_list_box.select_clear(self.stlv_list_sel)
            self.stlv_list_sel += 1
            self.stlv_list_box.select_set(self.stlv_list_sel)

    # Function adjusts selection of item when user presses up key (STLV List)
    def stlv_list_up(self, event):
        if self.stlv_list_sel > 0:
            self.stlv_list_box.select_clear(self.stlv_list_sel)
            self.stlv_list_sel -= 1
            self.stlv_list_box.select_set(self.stlv_list_sel)

    # Function adjusts selection of item when user clicks item (STLVS List)
    def stlvs_list_select(self, event):
        if self.stlvs_list_box and self.stlvs_list_box.curselection() \
                and -1 < self.stlvs_list_sel < self.stlvs_list_box.size() - 1:
            self.stlvs_list_sel = self.stlvs_list_box.curselection()[0]

    # Function adjusts selection of item when user presses down key (STLVS List)
    def stlvs_list_down(self, event):
        if self.stlvs_list_sel < self.stlvs_list_box.size() - 1:
            self.stlvs_list_box.select_clear(self.stlvs_list_sel)
            self.stlvs_list_sel += 1
            self.stlvs_list_box.select_set(self.stlvs_list_sel)

    # Function adjusts selection of item when user presses up key (STLVS List)
    def stlvs_list_up(self, event):
        if self.stlvs_list_sel > 0:
            self.stlvs_list_box.select_clear(self.stlvs_list_sel)
            self.stlvs_list_sel -= 1
            self.stlvs_list_box.select_set(self.stlvs_list_sel)

    # Button to migrate single record to right list (SQL TBL Section)
    def move_right(self):
        if self.stlv_list_box.curselection():
            self.stlvs_list_box.insert('end', self.stlv_list_box.get(self.stlv_list_box.curselection()))
            self.stlv_list_box.delete(self.stlv_list_box.curselection(), self.stlv_list_box.curselection())

            if self.stlv_list_box.size() > 0:
                if self.stlv_list_sel > 0:
                    self.stlv_list_sel -= 1
                self.stlv_list_box.select_set(self.stlv_list_sel)
            elif self.stlv_list_sel > 0:
                self.stlv_list_sel = -1
            else:
                self.stlv_list_sel = 0
                self.stlvs_list_sel = 0
                self.stlvs_list_box.select_set(self.stlvs_list_sel)

    # Button to migrate single record to right list (SQL TBL Section)
    def move_right_all(self):
        if self.stlv_list_box.size() > 0:
            for i in range(self.stlv_list_box.size()):
                self.stlvs_list_box.insert('end', self.stlv_list_box.get(i))

            self.stlv_list_box.delete(0, self.stlv_list_box.size() - 1)
            self.stlv_list_sel = 0
            self.stlvs_list_sel = 0
            self.stlvs_list_box.select_set(self.stlvs_list_sel)

    # Button to migrate single record to right list (SQL TBL Section)
    def move_left(self):
        if self.stlvs_list_box.curselection():
            self.stlv_list_box.insert('end', self.stlvs_list_box.get(self.stlvs_list_box.curselection()))
            self.stlvs_list_box.delete(self.stlvs_list_box.curselection(), self.stlvs_list_box.curselection())

            if self.stlvs_list_box.size() > 0:
                if self.stlvs_list_sel > 0:
                    self.stlvs_list_sel -= 1
                self.stlvs_list_box.select_set(self.stlvs_list_sel)
            elif self.stlvs_list_sel > 0:
                self.stlvs_list_sel = -1
            else:
                self.stlvs_list_sel = 0
                self.stlv_list_sel = 0
                self.stlv_list_box.select_set(self.stlv_list_sel)

    # Button to migrate single record to right list (SQL TBL Section)
    def move_left_all(self):
        if self.stlvs_list_box.size() > 0:
            for i in range(self.stlvs_list_box.size()):
                self.stlv_list_box.insert('end', self.stlvs_list_box.get(i))

            self.stlvs_list_box.delete(0, self.stlvs_list_box.size() - 1)
            self.stlvs_list_sel = 0
            self.stlv_list_sel = 0
            self.stlv_list_box.select_set(self.stlv_list_sel)

    # Function to connect to SQL connection for this class
    def sql_connect(self):
        if self.asql.test_conn('alch'):
            self.asql.connect('alch')
            return True
        else:
            return False

    # Function to close SQL connection for this class
    def sql_close(self):
        self.asql.close()

    # Function to save settings when the Save Settings button is pressed
    def save_settings(self):
        if not self.email_server.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Email Server',
                                 parent=self.main)
        elif not self.email_port.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Email Port',
                                 parent=self.main)
        elif not self.email_user_name.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Email User Name',
                                 parent=self.main)
        elif not self.email_user_pass.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Email User Pass',
                                 parent=self.main)
        elif not self.email_from.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Email From',
                                 parent=self.main)
        elif not self.email_to.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Email To',
                                 parent=self.main)
        elif not self.email_cc.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Email CC',
                                 parent=self.main)
        elif not self.sql_batch.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for SQL Batch Table',
                                 parent=self.main)
        elif not self.sql_table.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for SQL History Record Table',
                                 parent=self.main)
        elif not self.source_table.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for SQL Source Table',
                                 parent=self.main)
        elif not self.cat_table.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for SQL CAT Table',
                                 parent=self.main)
        elif self.stlvs_list_box.size() < 1:
            messagebox.showerror('Field Empty Error!', 'No table columns have been selected from SQL History Record Table',
                                 parent=self.main)
        elif not str(self.email_port.get()).isnumeric():
            messagebox.showerror('Field Error!', 'Email Port is a non-numeric port',
                                 parent=self.main)
        elif not self.check_table(self.sql_batch.get()):
            messagebox.showerror('Field Error!', 'SQL Batch Table does not exist in SQL Server',
                                 parent=self.main)
        else:
            email_err = 0

            try:
                server = smtplib.SMTP(str(self.email_server.get()), int(self.email_port.get()))

                try:
                    server.ehlo()
                    server.starttls()
                    server.ehlo()
                    server.login(self.email_user_name.get(), self.email_upass_obj.decrypt_text())
                except:
                    email_err = 2
                    pass
                finally:
                    server.close()
            except:
                email_err = 1
                pass

            if email_err == 1:
                messagebox.showerror('Invalid Email Settings!',
                                     'Server and/or Port does not exist. Please specify the correct information',
                                     parent=self.main)
            elif email_err == 2:
                messagebox.showerror('Invalid Email Settings!',
                                     'User name and/or user pass is invalid. Please specify the correct information',
                                     parent=self.main)
            else:
                self.add_setting('Settings', self.email_server.get(), 'Email_Server')
                self.add_setting('Settings', self.email_port.get(), 'Email_Port')
                self.add_setting('Settings', self.email_user_name.get(), 'Email_User')
                self.add_setting('Settings', self.email_upass_obj.decrypt_text(), 'Email_Pass')
                self.add_setting('Local_Settings', self.email_from.get(), 'Email_From')
                self.add_setting('Local_Settings', self.email_to.get(), 'Email_To')
                self.add_setting('Local_Settings', self.email_cc.get(), 'Email_Cc')
                self.add_setting('Local_Settings', self.sql_table.get(), 'Stlv_Tbl')
                self.add_setting('Local_Settings', self.sql_batch.get(), 'Stlv_Batch_Tbl')
                self.add_setting('Local_Settings', self.source_table.get(), 'Source_Tbl')
                self.add_setting('Local_Settings', self.cat_table.get(), 'Cat_Tbl')
                self.add_setting('Local_Settings', self.stlvs_list_box.get(0, self.stlvs_list_box.size() - 1),
                                 'Stlv_Tbl_Cols', False)

                self.main.destroy()

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        self.main.destroy()


# Main loop routine to create GUI Settings
if __name__ == '__main__':
    obj = SettingsGUI()
    obj.build_gui()
