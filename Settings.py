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
        econn_frame = LabelFrame(self.main, text='Settings', width=250, height=170)
        econn_item_frame = Frame(self.main)
        econn_item_frame2 = Frame(self.main)
        econn_item_frame3 = Frame(self.main)
        econn_item_frame4 = Frame(self.main)
        emsg_frame = LabelFrame(self.main, text='Message', width=250, height=170)
        stlv_frame = LabelFrame(self.main, text='Send To LV', width=500, height=200)
        stlv_list_frame = LabelFrame(self.main, text='Table Columns', width=300, height=200)
        stlv_settings_frame = LabelFrame(self.main, text='Settings', width=200, height=200)
        buttons_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        network_frame.pack(fill="both")
        email_frame.pack(fill="both")
        econn_frame.pack(in_=email_frame, fill="both", side=LEFT)
        econn_item_frame.pack(in_=econn_frame)
        econn_item_frame2.pack(in_=econn_frame)
        econn_item_frame3.pack(in_=econn_frame)
        econn_item_frame4.pack(in_=econn_frame)
        emsg_frame.pack(in_=email_frame, fill="both", side=RIGHT)
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
        eserver_label = Label(self.main, text='Email Server:')
        eserver_txtbox = Entry(self.main, textvariable=self.email_server)
        eserver_label.pack(in_=econn_item_frame, side=LEFT, pady=7)
        eserver_txtbox.pack(in_=econn_item_frame, side=LEFT)

        eport_label = Label(self.main, text='Email Port:')
        eport_txtbox = Entry(self.main, textvariable=self.email_port, width=8)
        eport_label.pack(in_=econn_item_frame2, side=LEFT, pady=7)
        eport_txtbox.pack(in_=econn_item_frame2, side=LEFT)

        euname_label = Label(self.main, text='Email User Name:')
        euname_txtbox = Entry(self.main, textvariable=self.email_user_name, width=14)
        euname_label.pack(in_=econn_item_frame3, side=LEFT, pady=7)
        euname_txtbox.pack(in_=econn_item_frame3, side=LEFT)

        eupass_label = Label(self.main, text='Email User Pass:')
        eupass_txtbox = Entry(self.main, textvariable=self.email_user_pass, width=14)
        eupass_label.pack(in_=econn_item_frame4, side=LEFT, pady=7)
        eupass_txtbox.pack(in_=econn_item_frame4, side=LEFT)


        # Show dialog
        self.main.mainloop()


if __name__ == '__main__':
    myobj = SettingsGUI()
    myobj.build_gui()
