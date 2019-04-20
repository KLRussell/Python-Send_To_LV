from tkinter import *
from tkinter import messagebox


class SettingsGUI:
    def __init__(self):
        self.main = Tk()
        self.server = StringVar()
        self.database = StringVar()
        self.email_server = StringVar()
        self.email_port = StringVar()
        self.email_user_name = StringVar()
        self.email_user_pass = StringVar()

    def build_gui(self):
        # Set GUI Geometry and GUI Title
        self.main.geometry('500x565+500+300')
        self.main.title('Send to LV Add Settings')

        # Set GUI Frames
        header_frame = Frame(self.main)
        network_frame = LabelFrame(self.main, text='Network Settings', width=500, height=70)
        email_frame = LabelFrame(self.main, text='E-mail', width=500, height=170)
        econn_frame = LabelFrame(email_frame, text='Settings', width=250, height=170)
        emsg_frame = LabelFrame(email_frame, text='Message', width=250, height=170)
        stlv_frame = LabelFrame(self.main, text='Send To LV', width=500, height=200)
        stlv_list_frame = LabelFrame(self.main, text='Table Columns', width=300, height=200)
        stlv_settings_frame = LabelFrame(self.main, text='Settings', width=200, height=200)
        buttons_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        network_frame.pack(fill="both")
        email_frame.pack(fill="both")
        econn_frame.grid(row=3, column=0, ipady=14)
        emsg_frame.grid(row=3, column=1)
        stlv_frame.pack(fill="both")
        stlv_list_frame.pack(in_=stlv_frame, fill="both", side=RIGHT)
        stlv_settings_frame.pack(in_=stlv_frame, fill="both", side=LEFT)
        buttons_frame.pack()

        # Apply Header text to Header_Frame that describes purpose of GUI
        header_text = 'Welcome to Send to LV! Settings haven''t been established.\nPlease fill out all empty fields below:'
        header = Message(self.main, text=header_text, width=375, justify=CENTER)
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
        eserver_label = Label(econn_frame, text='Email Server:')
        eserver_txtbox = Entry(econn_frame, textvariable=self.email_server)
        eserver_label.grid(row=0, column=0, padx=8, pady=5, sticky='w')
        eserver_txtbox.grid(row=0, column=1, padx=13, pady=5, sticky='e')

        eport_label = Label(econn_frame, text='Email Port:')
        eport_txtbox = Entry(econn_frame, textvariable=self.email_port)
        eport_label.grid(row=1, column=0, padx=8, pady=5, sticky='w')
        eport_txtbox.grid(row=1, column=1, padx=13, pady=5, sticky='e')

        euname_label = Label(econn_frame, text='Email User Name:')
        euname_txtbox = Entry(econn_frame, textvariable=self.email_user_name)
        euname_label.grid(row=2, column=0, padx=8, pady=5, sticky='w')
        euname_txtbox.grid(row=2, column=1, padx=13, pady=5, sticky='e')

        eupass_label = Label(econn_frame, text='Email User Pass:')
        eupass_txtbox = Entry(econn_frame, textvariable=self.email_user_pass)
        eupass_label.grid(row=3, column=0, padx=8, pady=5, sticky='w')
        eupass_txtbox.grid(row=3, column=1, padx=13, pady=5, sticky='e')

        # Show dialog
        self.main.mainloop()


if __name__ == '__main__':
    myobj = SettingsGUI()
    myobj.build_gui()
