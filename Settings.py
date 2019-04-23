from tkinter import *
from tkinter import messagebox
from Global import grabobjs
from Global import CryptHandle

import os

CurrDir = os.path.dirname(os.path.abspath(__file__))
Global_Objs = grabobjs(CurrDir)


class SettingsGUI:
    email_user_pass_obj = None
    curr_sql_table = None
    listbox = None
    listbox2 = None
    selection = 0
    selection2 = 0

    def __init__(self, header=None):
        if header:
            self.header_text = header
        else:
            self.header_text = 'Welcome to Send to LV Settings!\nSettings can be changed below.\nPress save when finished'

        self.asql = Global_Objs['SQL']
        self.asql.connect('alch')
        self.main = Tk()
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
        self.sql_tables = self.store_sql_tables()
        assert not self.sql_tables.empty

    def store_sql_tables(self):
        return self.asql.query('''
            select
                lower(concat(table_schema, '.', table_name)) SQL_TBL,
                table_schema,
                table_name
            
            from information_schema.tables''')

    def gui_cleanup(self, event):
        self.asql.close()

    def build_gui(self):
        # Set GUI Geometry and GUI Title
        self.main.geometry('509x600+500+100')
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
        stlv_settings_frame = LabelFrame(stlv_frame, text='SQL Table', width=505, height=70)
        buttons_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        network_frame.pack(fill="both")
        email_frame.pack(fill="both")
        econn_frame.grid(row=0, column=0, ipady=5)
        emsg_frame.grid(row=0, column=1, ipady=20)
        stlv_frame.pack(fill="both")
        stlv_settings_frame.grid(row=0, columnspan=3, ipady=2)
        stlv_list_frame.grid(row=1, column=0, rowspan=4, ipady=2, sticky='e')
        stlv_list_frame2.grid(row=1, column=1, rowspan=4, ipady=2)
        stlv_list_frame3.grid(row=1, column=2, rowspan=4, ipady=2, sticky='w')
        buttons_frame.pack(fill='both')

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=self.header_text, width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Network Labels & Input boxes to the Network_Frame
        #     SQL Server Input Box
        server_label = Label(self.main, text='SQL Server:', padx=15, pady=7)
        server_txtbox = Entry(self.main, textvariable=self.server)
        server_label.pack(in_=network_frame, side=LEFT)
        server_txtbox.pack(in_=network_frame, side=LEFT)

        #     Server Database Input Box
        database_label = Label(self.main, text='Server Database:')
        database_txtbox = Entry(self.main, textvariable=self.database)
        database_txtbox.pack(in_=network_frame, side=RIGHT, pady=7, padx=15)
        database_label.pack(in_=network_frame, side=RIGHT)

        # Apply Email Labels & Input boxes to the EConn_Frame
        #     Email Server Input Box
        eserver_label = Label(econn_frame, text='Email Server:')
        eserver_txtbox = Entry(econn_frame, textvariable=self.email_server)
        eserver_label.grid(row=0, column=0, padx=8, pady=5, sticky='w')
        eserver_txtbox.grid(row=0, column=1, padx=13, pady=5, sticky='e')

        #     Email Port Input Box
        eport_label = Label(econn_frame, text='Email Port:')
        eport_txtbox = Entry(econn_frame, textvariable=self.email_port)
        eport_label.grid(row=1, column=0, padx=8, pady=5, sticky='w')
        eport_txtbox.grid(row=1, column=1, padx=13, pady=5, sticky='e')

        #     Email User Name Input Box
        euname_label = Label(econn_frame, text='Email User Name:')
        euname_txtbox = Entry(econn_frame, textvariable=self.email_user_name)
        euname_label.grid(row=2, column=0, padx=8, pady=5, sticky='w')
        euname_txtbox.grid(row=2, column=1, padx=13, pady=5, sticky='e')

        #     Email User Pass Input Box
        eupass_label = Label(econn_frame, text='Email User Pass:')
        eupass_txtbox = Entry(econn_frame, textvariable=self.email_user_pass)
        eupass_label.grid(row=3, column=0, padx=8, pady=5, sticky='w')
        eupass_txtbox.grid(row=3, column=1, padx=13, pady=5, sticky='e')

        # Apply Email Labels & Input boxes to the EMsg_Frame
        #     From Email Address Input Box
        efrom_label = Label(emsg_frame, text='From Addr:')
        efrom_txtbox = Entry(emsg_frame, textvariable=self.email_from)
        efrom_label.grid(row=0, column=0, padx=8, pady=5, sticky='w')
        efrom_txtbox.grid(row=0, column=1, padx=13, pady=5, sticky='e')

        #     To Email Address Input Box
        eto_label = Label(emsg_frame, text='To Addr:')
        eto_txtbox = Entry(emsg_frame, textvariable=self.email_to)
        eto_label.grid(row=1, column=0, padx=8, pady=5, sticky='w')
        eto_txtbox.grid(row=1, column=1, padx=13, pady=5, sticky='e')

        #     CC Email Address Input Box
        ecc_label = Label(emsg_frame, text='CC Addr:')
        ecc_txtbox = Entry(emsg_frame, textvariable=self.email_cc)
        ecc_label.grid(row=2, column=0, padx=8, pady=5, sticky='w')
        ecc_txtbox.grid(row=2, column=1, padx=13, pady=5, sticky='e')

        # Apply Line Verification SQL Table Label & Input boxes to the EMsg_Frame
        #     From Email Address Input Box
        lv_sqltbl_txtbox = Entry(stlv_settings_frame, textvariable=self.sql_table, width=76)
        lv_sqltbl_txtbox.grid(row=0, columnspan=3, padx=20, pady=5)

        lv_scrollbar = Scrollbar(stlv_list_frame, orient="vertical")
        lv_scrollbar2 = Scrollbar(stlv_list_frame, orient="horizontal")
        self.listbox = Listbox(stlv_list_frame, selectmode=SINGLE, width=25, yscrollcommand=lv_scrollbar.set,
                               xscrollcommand=lv_scrollbar2.set)
        lv_scrollbar.config(command=self.listbox.yview)
        lv_scrollbar2.config(command=self.listbox.xview)
        self.listbox.grid(row=0, column=0, rowspan=4, padx=8, pady=5)
        lv_scrollbar.grid(row=0, column=1, rowspan=4, sticky=N + S)
        lv_scrollbar2.grid(row=4, column=0, columnspan=2, sticky=E + W)

        lv_scrollbar3 = Scrollbar(stlv_list_frame3, orient="vertical")
        lv_scrollbar4 = Scrollbar(stlv_list_frame3, orient="horizontal")
        self.listbox2 = Listbox(stlv_list_frame3, selectmode=SINGLE, width=25, yscrollcommand=lv_scrollbar3.set,
                                xscrollcommand=lv_scrollbar4.set)
        lv_scrollbar3.config(command=self.listbox2.yview)
        lv_scrollbar4.config(command=self.listbox2.xview)
        self.listbox2.grid(row=0, column=0, rowspan=4, padx=8, pady=5)
        lv_scrollbar3.grid(row=0, column=1, rowspan=4, sticky=N + S)
        lv_scrollbar4.grid(row=4, column=0, columnspan=2, sticky=E + W)

        move_all_button = Button(stlv_list_frame2, text='>>', width=4, command=self.move_right_all)
        move_all_button.grid(row=0, column=1, pady=10)
        move_all_button = Button(stlv_list_frame2, text='>', width=4, command=self.move_right)
        move_all_button.grid(row=1, column=1, pady=5)
        move_all_button = Button(stlv_list_frame2, text='<', width=4, command=self.move_left)
        move_all_button.grid(row=2, column=1, pady=5)
        move_all_button = Button(stlv_list_frame2, text='<<', width=4, command=self.move_left_all)
        move_all_button.grid(row=3, column=1, pady=5)

        save_settings_button = Button(buttons_frame, text='Save Settings', width=20, command=self.save_settings)
        save_settings_button.grid(row=0, column=0, pady=6, padx=15)
        cancel_button = Button(buttons_frame, text='Cancel', width=20, command=self.cancel)
        cancel_button.grid(row=0, column=1, pady=6, padx=165)

        self.listbox.bind("<Down>", self.on_list_down)
        self.listbox.bind("<Up>", self.on_list_up)
        self.listbox.bind('<<ListboxSelect>>', self.on_select)
        self.listbox2.bind("<Down>", self.on_list_down2)
        self.listbox2.bind("<Up>", self.on_list_up2)
        self.listbox2.bind('<<ListboxSelect>>', self.on_select2)
        eupass_txtbox.bind('<KeyRelease>', self.hide_pass)
        lv_sqltbl_txtbox.bind('<KeyRelease>', self.populate_columns)
        self.main.bind('<Destroy>', self.gui_cleanup)

        self.fill_gui()
        # Show dialog
        self.main.mainloop()

    @staticmethod
    def fill_textbox(setting_name, set_val, val):
        if val:
            set_val.set(Global_Objs[setting_name].grab_item(val).decrypt_text())

    def fill_gui(self):
        self.fill_textbox('Settings', self.server, 'Server')
        self.fill_textbox('Settings', self.database, 'Database')
        self.fill_textbox('Settings', self.email_server, 'email_server')
        self.fill_textbox('Settings', self.email_port, 'email_port')
        self.fill_textbox('Settings', self.email_user_name, 'email_user')
        self.fill_textbox('Local_Settings', self.email_from, 'email_from')
        self.fill_textbox('Local_Settings', self.email_to, 'email_to')
        self.fill_textbox('Local_Settings', self.email_cc, 'email_cc')
        self.fill_textbox('Local_Settings', self.sql_table, 'STLV_TBL')

        self.email_user_pass_obj = Global_Objs['Settings'].grab_item('email_pass')
        myitems = Global_Objs['Local_Settings'].grab_item('STLV_TBL_Cols').decrypt_text()

        if self.email_user_pass_obj:
            self.email_user_pass.set('*' * len(self.email_user_pass_obj.decrypt_text()))

        if myitems:
            myitems = myitems.replace(', ', ',').split(',')

            for col in myitems:
                self.listbox2.insert('end', col)

        if self.sql_table:
            self.curr_sql_table = self.sql_table.get()
            myrow = self.sql_tables[self.sql_tables['SQL_TBL'] == self.sql_table.get().lower()].reset_index()

            if not myrow.empty:
                myresults = self.asql.query('''
                    select
                        Column_Name
        
                from INFORMATION_SCHEMA.COLUMNS
        
                where
                    TABLE_SCHEMA = '{0}'
                        and
                    TABLE_NAME = '{1}'
                '''.format(myrow['table_schema'][0], myrow['table_name'][0]))

                for col in myresults['Column_Name']:
                    if not myitems or col not in myitems:
                        self.listbox.insert('end', col)

    def hide_pass(self, event):
        if self.email_user_pass_obj:
            currpass = self.email_user_pass_obj.decrypt_text()
            i = 0

            for letter in self.email_user_pass.get():
                if letter != '*':
                    if i > len(currpass) - 1:
                        currpass += letter
                    else:
                        currpass[i] = letter
                i += 1

            if len(currpass) - i > 0:
                currpass = currpass[:i]

            if currpass:
                self.email_user_pass_obj.encrypt_text(currpass)
                self.email_user_pass.set('*' * len(self.email_user_pass_obj.decrypt_text()))
            else:
                self.email_user_pass_obj = None
                self.email_user_pass.set("")
        else:
            self.email_user_pass_obj = CryptHandle()
            self.email_user_pass_obj.encrypt_text(self.email_user_pass.get())
            self.email_user_pass.set('*' * len(self.email_user_pass_obj.decrypt_text()))

    def populate_columns(self, event):
        if self.sql_table.get().lower() in self.sql_tables['SQL_TBL'].tolist() and self.curr_sql_table != self.sql_table.get():
            myresponse = messagebox.askokcancel('Change Notice!',
                                                'Changing existing SQL table will erase the column lists below. Would you like to proceed?',
                                                parent=self.main)
            if myresponse:
                self.curr_sql_table = self.sql_table.get()

                if self.listbox.size() > 0:
                    self.listbox.delete(0, self.listbox.size() - 1)

                if self.listbox2.size() > 0:
                    self.listbox2.delete(0, self.listbox2.size() - 1)

                myrow = self.sql_tables[self.sql_tables['SQL_TBL'] == self.sql_table.get().lower()].reset_index()

                if not myrow.empty:
                    myresults = self.asql.query('''
                        select
                            Column_Name
        
                    from INFORMATION_SCHEMA.COLUMNS
        
                    where
                        TABLE_SCHEMA = '{0}'
                            and
                        TABLE_NAME = '{1}'
                    '''.format(myrow['table_schema'][0], myrow['table_name'][0]))

                    for col in myresults['Column_Name']:
                        self.listbox.insert('end', col)
            else:
                self.sql_table.set(self.curr_sql_table)

    def on_select(self, event):
        if self.listbox and self.listbox.curselection() and -1 < self.selection < self.listbox.size() - 1:
            self.selection = self.listbox.curselection()[0]

    def on_list_down(self, event):
        if self.selection < self.listbox.size() - 1:
            self.listbox.select_clear(self.selection)
            self.selection += 1
            self.listbox.select_set(self.selection)

    def on_list_up(self, event):
        if self.selection > 0:
            self.listbox.select_clear(self.selection)
            self.selection -= 1
            self.listbox.select_set(self.selection)

    def on_select2(self, event):
        if self.listbox2 and self.listbox2.curselection() and -1 < self.selection2 < self.listbox2.size() - 1:
            self.selection2 = self.listbox2.curselection()[0]

    def on_list_down2(self, event):
        if self.selection2 < self.listbox2.size() - 1:
            self.listbox2.select_clear(self.selection2)
            self.selection2 += 1
            self.listbox2.select_set(self.selection2)

    def on_list_up2(self, event):
        if self.selection2 > 0:
            self.listbox2.select_clear(self.selection2)
            self.selection2 -= 1
            self.listbox2.select_set(self.selection2)

    def save_settings(self):
        print('save settings')

    def cancel(self):
        self.main.destroy()

    def move_right_all(self):
        print('all right')

    def move_right(self):
        print('right')

    def move_left_all(self):
        print('all left')

    def move_left(self):
        print('left')


if __name__ == '__main__':
    header_text = 'Welcome to Send to LV Settings!\nSettings haven''t been established.\nPlease fill out all empty fields below:'
    myobj = SettingsGUI()
    myobj.build_gui()
