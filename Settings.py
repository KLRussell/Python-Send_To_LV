from tkinter import *
from tkinter import messagebox


class SettingsGUI:
    listbox = None
    listbox2 = None

    def __init__(self, header=None):
        if header:
            self.header_text = header
        else:
            self.header_text = 'Welcome to Send to LV Settings!\nSettings can be changed below.\nPress save when finished'

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

    def build_gui(self):
        # Set GUI Geometry and GUI Title
        self.main.geometry('509x595+500+100')
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
        buttons_frame.pack()

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
        self.listbox.grid(row=0, column=0, rowspan=4, padx=8, pady=5)

        lv_scrollbar3 = Scrollbar(stlv_list_frame3, orient="vertical")
        lv_scrollbar4 = Scrollbar(stlv_list_frame3, orient="horizontal")
        self.listbox2 = Listbox(stlv_list_frame3, selectmode=SINGLE, width=25, yscrollcommand=lv_scrollbar3.set,
                                xscrollcommand=lv_scrollbar4.set)
        self.listbox2.grid(row=0, column=0, rowspan=4, padx=8, pady=5)

        move_all_button = Button(stlv_list_frame2, text='>>', width=4, command=self.move_right_all)
        move_all_button.grid(row=0, column=1, pady=10)
        move_all_button = Button(stlv_list_frame2, text='>', width=4, command=self.move_right)
        move_all_button.grid(row=1, column=1, pady=5)
        move_all_button = Button(stlv_list_frame2, text='<', width=4, command=self.move_left)
        move_all_button.grid(row=2, column=1, pady=5)
        move_all_button = Button(stlv_list_frame2, text='<<', width=4, command=self.move_left_all)
        move_all_button.grid(row=3, column=1, pady=5)

        # Show dialog
        self.main.mainloop()

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
