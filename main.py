import tkinter as tk
from PIL import ImageTk, Image
from decimal import Decimal
import csv
import smtplib
import os.path
from os.path import isfile
from os.path import basename
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import sqlite3
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
from ttkbootstrap.tableview import Tableview
from ttkbootstrap.scrolled import ScrolledFrame
from ttkbootstrap.toast import ToastNotification
from copy import deepcopy
from datetime import date
import datetime
from time import localtime, strftime
import threading
from ttkbootstrap.icons import Icon
import pyglet
import pprint
from tkinter import ttk
import os
from dotenv import load_dotenv
import httpx
import webbrowser
import msal
from dotenv import load_dotenv
from datetime import timedelta

def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS2
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

class Application(ttkb.Window):
    def __init__(self, *args, themename=None, iconphoto=None, **kwargs):
        super().__init__(*args, themename=themename, iconphoto=iconphoto, **kwargs)
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()

        c.execute('''CREATE TABLE IF NOT EXISTS Emails(
            email_id INTEGER PRIMARY KEY AUTOINCREMENT,
            email_name TEXT NOT NULL,
            email_active INTEGER 
        )''')
        conn.commit()

        c.execute('''CREATE TABLE IF NOT EXISTS Customers(
            customer_id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_name TEXT NOT NULL
        )''')
        conn.commit()
        c.execute('''CREATE TABLE IF NOT EXISTS Customer_Items(
            item_id INTEGER PRIMARY KEY AUTOINCREMENT,
            real_item_id TEXT NOT NULL,
            customer_id INTEGER,
            FOREIGN KEY (customer_id) REFERENCES Customers(customer_id)
        )''')
        conn.commit()

        c.execute('''CREATE TABLE IF NOT EXISTS Shipments(
            shipment_id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER,
            total_weight INTEGER,
            total_trucks REAL,
            date TEXT,
            FOREIGN KEY (customer_id) REFERENCES Customers(customer_id)
            FOREIGN KEY (date) REFERENCES Dates(date)
        )''')
        conn.commit()
        conn.close()
        self.iconphoto = iconphoto
        self.themename = themename
        self.geometry("1200x600")
        self.title("Linamar Planner")
        self.place_window_center()
        self.iconbitmap(resource_path("Images\\linamarLogo256.ico"))

        pyglet.font.add_file("Fonts\\Ubuntu-Bold.ttf")
        pyglet.font.add_file("Fonts\\Inter-VariableFont_opsz,wght.ttf")
        
        self.home_style_buttons = ttkb.Style()
        self.home_style_buttons.configure(
            'info.TButton', font=("Inter 11 bold")

            )
        self.back_button_styles = ttkb.Style()
        self.back_button_styles.configure(
            'danger.TButton', font=("Inter 11 bold")
        )
        self.warning_button_styles = ttkb.Style()
        self.warning_button_styles.configure(
            'warning.TButton', font=("Inter 11 bold")
        )
        self.success_button_styles = ttkb.Style()
        self.success_button_styles.configure(
            'success.TButton', font=("Inter 11 bold")
        )
        self.warning_button_outline_styles = ttkb.Style()
        self.warning_button_outline_styles.configure(
            'warning.Outline.TButton', font=("Inter 10 bold")
        )
        self.danger_button_outline_styles = ttkb.Style()
        self.danger_button_outline_styles.configure(
            'danger.Outline.TButton', font=("Inter 10 bold")
        )
        self.success_button_outline_styles = ttkb.Style()
        self.success_button_outline_styles.configure(
            'success.Outline.TButton', font=("Inter 10 bold")
        )
        self.info_button_outline_styles = ttkb.Style()
        self.info_button_outline_styles.configure(
            'info.Outline.TButton', font=("Inter 10 bold")
        )

        self.columnconfigure(0, weight=1, uniform="Main")
        self.columnconfigure(1, weight=1, uniform="Main")
        self.columnconfigure(2, weight=1, uniform="Main")
        
        self.rowconfigure(0, weight=1, uniform="Main")
        self.rowconfigure(1, weight=1, uniform="Main")

        self.main_frame = ttkb.Frame(self)
        self.main_frame.grid(row=0, column=0, columnspan=3, rowspan=2, sticky='nsew')
        self.main_frame.columnconfigure(0, weight=1, uniform="main")
        self.main_frame.columnconfigure(1, weight=1, uniform="main")
        self.main_frame.columnconfigure(2, weight=1, uniform="main")
        self.main_frame.columnconfigure(3, weight=1, uniform="main")
        self.main_frame.rowconfigure(0, weight=1, uniform="main")
        self.main_frame.rowconfigure(1, weight=1, uniform="main")

        self.btn_1 = ttkb.Button(self.main_frame, text="New Customer", bootstyle="info", style="info.TButton", command=lambda: [self.hide_root(), self.show_customer_frame()])
        self.btn_1.grid(row=0, column=0, sticky='new', padx=10, pady=10)

        self.btn_2 = ttkb.Button(self.main_frame, text="New Appointment", bootstyle="info", command=lambda: [self.hide_root(), self.show_appointment_frame()])
        self.btn_2.grid(row=0, column=1,sticky='new', padx=10, pady=10)

        self.btn_3 = ttkb.Button(self.main_frame, text="Generate Report", bootstyle="info", command=lambda: [self.hide_root(), self.show_status_frame()])
        self.btn_3.grid(row=0, column=2,sticky='new', padx=10, pady=10)

        self.btn_4 = ttkb.Button(self.main_frame, text="Options", bootstyle="info", command=self.show_options_frame)
        self.btn_4.grid(row=0, column=3, sticky='new', padx=10, pady=10)


        

        self.appointment_frame = NewAppointment(self)
        #self.home_frame.grid(row=0, column=0, columnspan=2, sticky='nsew')

        self.appointment_frame.access_token_main()

        self.customer_frame = NewCustomer(self)
        #self.customer_frame.grid(row=0, column=0, columnspan=2, sticky='nsew')

        self.status_frame = CheckStatus(self)
        #self.customer_frame.grid(row=0, column=0, columnspan=2, sticky='nsew')

        self.options_frame = OptionsFrame(self)

        #Appointment Setting
        self.appointment_frame.set_customer_frame(self.customer_frame)
        self.appointment_frame.set_status_frame(self.status_frame)
        
        #Customer Frame Setting
        self.customer_frame.set_appointment_frame(self.appointment_frame)
        self.customer_frame.set_status_frame(self.status_frame)

        #Status Frame Setting
        self.status_frame.set_appointment_frame(self.appointment_frame)
        self.status_frame.set_customer_frame(self.customer_frame)

        #Notice Frame Setting
        
    def hide_root(self):
        self.main_frame.grid_remove()
        
    def show_customer_frame(self):
        self.customer_frame.grid(row=0, column=0, columnspan=6, rowspan=2, sticky='nsew')

    def show_appointment_frame(self):
        self.appointment_frame.grid(row=0, column=0, columnspan=3, rowspan=2, sticky='nsew')
        self.appointment_frame.receieve_customer_names()
        self.appointment_frame.validate_entry()

    def show_options_frame(self):
        self.options_frame.grid(row=0, column=0, columnspan=3, rowspan=2, sticky='nsew')

    def show_status_frame(self):
        self.status_frame.grid(row=0, column=0, columnspan=3, rowspan=2, sticky='nsew')
        self.status_frame.retrieve_customers()

    def return_to_home(self):
        self.customer_frame.grid_remove()
        self.appointment_frame.grid_remove()
        self.status_frame.grid_remove()
        self.options_frame.grid_remove()
        self.main_frame.grid(row=0, column=0, rowspan=2, columnspan=3, sticky='nsew')

def main():
    app = Application(themename="darkly")
    app.mainloop()

class OptionsFrame(ttkb.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        #Main Column & Row Config
        self.option_add('*TCombobox*Listbox.font', "Inter 12")

        #self.validate_new_item_entry = self.register(self.validate_new_item_entry)

        self.columnconfigure(0, weight=1, uniform="options")
        self.columnconfigure(1, weight=1, uniform="options")
        self.columnconfigure(2, weight=1, uniform="options")
        self.columnconfigure(3, weight=1, uniform="options")
        
        self.rowconfigure(0, weight=1, uniform="options")
        self.rowconfigure(1, weight=1, uniform="options")
        self.rowconfigure(2, weight=1, uniform="options")
        self.rowconfigure(3, weight=1, uniform="options")

        self.main_frame = ttkb.Frame(self)
        self.main_frame.grid(row=0, column=0, rowspan=4, columnspan=4, sticky='nsew')
        self.main_frame.rowconfigure(0, weight=1, uniform="options")
        self.main_frame.rowconfigure(1, weight=1, uniform="options")
        self.main_frame.rowconfigure(2, weight=1, uniform="options")
        self.main_frame.columnconfigure(0, weight=1, uniform="options")
        self.main_frame.columnconfigure(1, weight=1, uniform="options")
        self.main_frame.columnconfigure(2, weight=1, uniform="options")
        self.main_frame.columnconfigure(3, weight=1, uniform="options")

        #Show_Cust Frame
        self.cust_opt_frame = ttkb.Frame(self, bootstyle="")
        #self.cust_opt_frame.grid(row=0, column=0, rowspan=4, columnspan=4, sticky='nsew')
        self.cust_opt_frame.rowconfigure(0, weight=1, uniform="options")
        self.cust_opt_frame.rowconfigure(1, weight=1, uniform="options")
        self.cust_opt_frame.rowconfigure(2, weight=1, uniform="options")
        self.cust_opt_frame.rowconfigure(3, weight=2, uniform="options")
        self.cust_opt_frame.columnconfigure(0, weight=1, uniform="options")
        self.cust_opt_frame.columnconfigure(1, weight=1, uniform="options")
        self.cust_opt_frame.columnconfigure(2, weight=1, uniform="options")
        self.cust_opt_frame.columnconfigure(3, weight=1, uniform="options")

        #Left Frame
        self.left_frame = ttkb.Frame(self.cust_opt_frame)
        self.left_frame.grid(row=0, column=0, columnspan=2, rowspan=3, sticky='nsew')
        self.left_frame.columnconfigure(0, weight=1, uniform="options")
        self.left_frame.rowconfigure(0, weight=2, uniform="options")

        #Scrolled Frame
        self.scrolled_frame = ScrolledFrame(self.left_frame, width=400)
        self.scrolled_frame.grid(row=0, column=0, columnspan=2, rowspan=3, sticky='nsew', pady=10)
        self.scrolled_frame.columnconfigure(0, weight=4, uniform="options")
        self.scrolled_frame.columnconfigure(1, weight=1, uniform="options")
        self.scrolled_frame.columnconfigure(2, weight=1, uniform="options")

        self.scrolled_frame.rowconfigure(0, weight=1, uniform="options")
        self.scrolled_frame.rowconfigure(1, weight=1, uniform="options")

        self.scrolled_frame.hide_scrollbars()
        
        #Right Frame
        self.right_frame = ttkb.Frame(self.cust_opt_frame, bootstyle="")
        self.right_frame.grid(row=0, column=2, columnspan=4, rowspan=3, sticky='nsew')
        self.right_frame.columnconfigure(0, weight=1, uniform="options")
        self.right_frame.columnconfigure(1, weight=1, uniform="options")
        self.right_frame.rowconfigure(0, weight=2, uniform="options")
        self.right_frame.rowconfigure(1, weight=1, uniform="options")
        self.right_frame.rowconfigure(2, weight=1, uniform="options")
        self.right_frame.rowconfigure(3, weight=1, uniform="options")
        self.right_frame.rowconfigure(4, weight=1, uniform="options")
        self.right_frame.rowconfigure(5, weight=1, uniform="options")

        #Top Right Frame
        self.top_right_frame = ttkb.Frame(self.right_frame, bootstyle="")
        self.top_right_frame.grid(row=0, rowspan=2, column=0, columnspan=4, sticky='nsew')
        self.top_right_frame.columnconfigure(0, weight=1, uniform="options")
        self.top_right_frame.columnconfigure(1, weight=1, uniform="options")

        self.top_right_frame.rowconfigure(0, weight=1, uniform="options")
        self.top_right_frame.rowconfigure(1, weight=1, uniform="options")
        self.top_right_frame.rowconfigure(2, weight=1, uniform="options")
        self.top_right_frame.rowconfigure(3, weight=1, uniform="options")

        self.top_right_frame_top = ttkb.Frame(self.top_right_frame, bootstyle="")
        self.top_right_frame_top.grid(row=0,rowspan=2, column=0, columnspan=4, sticky='nsew', pady=10)
        self.top_right_frame_top.columnconfigure(0, weight=1, uniform="options")
        self.top_right_frame_top.columnconfigure(1, weight=1, uniform="options")
        self.top_right_frame_top.columnconfigure(2, weight=1, uniform="options")
        self.top_right_frame_top.columnconfigure(3, weight=1, uniform="options")

        self.top_right_frame_top.rowconfigure(0, weight=1, uniform="options")
        self.top_right_frame_top.rowconfigure(1, weight=1, uniform="options")
        #self.top_right_frame_top.rowconfigure(2, weight=2, uniform="options")
        

        self.top_right_frame_bottom = ttkb.Frame(self.top_right_frame, bootstyle="")
        self.top_right_frame_bottom.grid(row=2,rowspan=4, column=0, columnspan=4, sticky='nsew', pady=5)
        self.top_right_frame_bottom.columnconfigure(0, weight=1, uniform="options")
        self.top_right_frame_bottom.columnconfigure(1, weight=1, uniform="options")
        self.top_right_frame_bottom.columnconfigure(2, weight=1, uniform="options")
        self.top_right_frame_bottom.columnconfigure(3, weight=1, uniform="options")

        self.top_right_frame_bottom.rowconfigure(0, weight=1, uniform="options")
        self.top_right_frame_bottom.rowconfigure(1, weight=1, uniform="options")
        

        #Middle Right Frame
        self.middle_right_frame = ttkb.Frame(self.right_frame, bootstyle="")
        self.middle_right_frame.grid(row=2, rowspan=2, column=0, columnspan=2, sticky='nsew')

        self.middle_right_frame.columnconfigure(0, weight=1, uniform="options")
        self.middle_right_frame.columnconfigure(1, weight=1, uniform="options")

        self.middle_right_frame.rowconfigure(0, weight=1, uniform="options")
        self.middle_right_frame.rowconfigure(1, weight=1, uniform="options")

        self.middle_right_top_frame = ttkb.Frame(self.middle_right_frame, bootstyle="")
        self.middle_right_top_frame.grid(row=0, column=0, columnspan=4, sticky='nsew', pady=10)
        self.middle_right_top_frame.columnconfigure(0, weight=1, uniform="options")
        self.middle_right_top_frame.columnconfigure(1, weight=1, uniform="options")
        self.middle_right_top_frame.columnconfigure(2, weight=1, uniform="options")
        self.middle_right_top_frame.columnconfigure(3, weight=1, uniform="options")

        self.middle_right_top_frame.rowconfigure(0, weight=1, uniform="options")
        self.middle_right_top_frame.rowconfigure(1, weight=1, uniform="options")

        self.middle_right_bottom_frame = ttkb.Frame(self.middle_right_frame, bootstyle="")
        self.middle_right_bottom_frame.grid(row=1, column=0, columnspan=4, sticky='nsew', pady=10)
        self.middle_right_bottom_frame.columnconfigure(0, weight=1, uniform="options")
        self.middle_right_bottom_frame.columnconfigure(1, weight=1, uniform="options")
        self.middle_right_bottom_frame.columnconfigure(2, weight=1, uniform="options")
        self.middle_right_bottom_frame.columnconfigure(3, weight=1, uniform="options")

        self.middle_right_bottom_frame.rowconfigure(0, weight=1, uniform="options")
        self.middle_right_bottom_frame.rowconfigure(1, weight=1, uniform="options")

        #Bottom Right Frame
        self.bottom_right_frame = ttkb.Frame(self.right_frame, bootstyle="")
        self.bottom_right_frame.grid(row=4, rowspan=5, column=0, columnspan=2, sticky='nsew')
        self.bottom_right_frame.columnconfigure(0, weight=1, uniform="options")
        self.bottom_right_frame.columnconfigure(1, weight=1, uniform="options")
        self.bottom_right_frame.columnconfigure(2, weight=1, uniform="options")
        self.bottom_right_frame.columnconfigure(3, weight=1, uniform="options")

        self.bottom_right_frame.rowconfigure(0, weight=1, uniform="options")
        self.bottom_right_frame.rowconfigure(1, weight=1, uniform="options")
        
        self.bottom_right_frame_top = ttkb.Frame(self.bottom_right_frame, bootstyle="")
        self.bottom_right_frame_top.grid(row=0, column=0, columnspan=4, sticky='nsew', pady=10)
        self.bottom_right_frame_top.columnconfigure(0, weight=1, uniform="options")
        self.bottom_right_frame_top.columnconfigure(1, weight=1, uniform="options")
        self.bottom_right_frame_top.columnconfigure(2, weight=1, uniform="options")
        self.bottom_right_frame_top.columnconfigure(3, weight=1, uniform="options")
        
        self.bottom_right_frame_top.rowconfigure(0, weight=1, uniform="options")
        self.bottom_right_frame_top.rowconfigure(1, weight=1, uniform="options")

        self.bottom_right_frame_bottom = ttkb.Frame(self.bottom_right_frame, bootstyle="")
        self.bottom_right_frame_bottom.grid(row=1, column=0, columnspan=4, sticky='nsew', pady=10)
        self.bottom_right_frame_bottom.columnconfigure(0, weight=1, uniform="options")
        self.bottom_right_frame_bottom.columnconfigure(1, weight=1, uniform="options")
        self.bottom_right_frame_bottom.columnconfigure(2, weight=1, uniform="options")
        self.bottom_right_frame_bottom.columnconfigure(3, weight=1, uniform="options")
        
        self.bottom_right_frame_bottom.rowconfigure(0, weight=1, uniform="options")
        self.bottom_right_frame_bottom.rowconfigure(1, weight=1, uniform="options")

        #Bottom Status Frame
        self.bottom_status_frame = ttkb.Frame(self.cust_opt_frame)
        self.bottom_status_frame.grid(row=5, column=0, columnspan=4, sticky='nsew')
        self.bottom_status_frame.columnconfigure(0, weight=1, uniform="options")
        self.bottom_status_frame.columnconfigure(1, weight=1, uniform="options")
        self.bottom_status_frame.columnconfigure(2, weight=1, uniform="options")
        self.bottom_status_frame.columnconfigure(3, weight=1, uniform="options")

        self.bottom_status_frame.rowconfigure(0, weight=1, uniform="options")

        self.cust_label = ttkb.Label(self.top_right_frame_top, text="Existing Name:", font=("Inter", 13))
        self.cust_label.grid(row=0, column=0, columnspan=2, sticky='ew', padx=10)

        self.new_cust_lbl = ttkb.Label(self.top_right_frame_top, text="New Name", font=("Inter", 13))
        self.new_cust_lbl.grid(row=0, column=2, sticky='ew', padx=10)

        self.cust_entry = ttkb.Entry(self.top_right_frame_top, font=("Inter", 13), width=0, state="disabled")
        self.cust_entry.grid(row=1, column=0, columnspan=2, sticky='new', padx=10)

        self.new_cust_name_entry = ttkb.Entry(self.top_right_frame_top, font=("Inter", 13), width=0, state="disabled")
        self.new_cust_name_entry.grid(row=1, column=2, columnspan=4, sticky='new', padx=10)

        #Email Frame Config --------------

        self.email_opt_frame = ttkb.Frame(self, bootstyle="")
        #self.email_opt_frame.grid(row=0, column=0, rowspan=4, columnspan=4, sticky='nsew')
        self.email_opt_frame.rowconfigure(0, weight=1, uniform="email")
        self.email_opt_frame.rowconfigure(1, weight=1, uniform="email")
        self.email_opt_frame.rowconfigure(2, weight=1, uniform="email")
        self.email_opt_frame.rowconfigure(3, weight=2, uniform="email")
        self.email_opt_frame.columnconfigure(0, weight=1, uniform="email")
        self.email_opt_frame.columnconfigure(1, weight=1, uniform="email")
        self.email_opt_frame.columnconfigure(2, weight=1, uniform="email")
        self.email_opt_frame.columnconfigure(3, weight=1, uniform="email")

        #LEFT FRAME AND RIGHT FRAME ARE SWAPPED BECUASE LAZY


        #Left Frame
        self.email_left_frame = ttkb.Frame(self.email_opt_frame)
        self.email_left_frame.grid(row=0, column=2, columnspan=4, rowspan=3, sticky='nsew')
        self.email_left_frame.columnconfigure(0, weight=1, uniform="email")
        self.email_left_frame.rowconfigure(0, weight=2, uniform="email")

        #Scrolled Frame
        self.email_scrolled_frame = ScrolledFrame(master=self.email_left_frame)
        self.email_scrolled_frame.grid(row=0, column=0, columnspan=2, rowspan=3, sticky='nsew', pady=10)
        self.email_scrolled_frame.columnconfigure(0, weight=4, uniform="email")
        self.email_scrolled_frame.columnconfigure(1, weight=1, uniform="email")
        self.email_scrolled_frame.columnconfigure(2, weight=1, uniform="email")

        self.email_scrolled_frame.rowconfigure(0, weight=1, uniform="email")
        self.email_scrolled_frame.rowconfigure(1, weight=1, uniform="email")

        self.email_scrolled_frame.hide_scrollbars()
        
        #Right Frame
        self.email_right_frame = ttkb.Frame(self.email_opt_frame, bootstyle="")
        self.email_right_frame.grid(row=0, column=0, columnspan=2, rowspan=3, sticky='nsew')
        self.email_right_frame.columnconfigure(0, weight=1, uniform="email")
        self.email_right_frame.columnconfigure(1, weight=1, uniform="email")
        self.email_right_frame.rowconfigure(0, weight=1, uniform="email")
        self.email_right_frame.rowconfigure(1, weight=1, uniform="email")
        self.email_right_frame.rowconfigure(2, weight=1, uniform="email")

        #Top Right Frame
        self.email_top_right_frame = ttkb.Frame(self.email_right_frame)
        self.email_top_right_frame.grid(row=0, column=0, columnspan=4, sticky='nsew')
        self.email_top_right_frame.columnconfigure(0, weight=1, uniform="email")
        self.email_top_right_frame.columnconfigure(1, weight=1, uniform="email")

        self.email_top_right_frame.rowconfigure(0, weight=1, uniform="email")
        self.email_top_right_frame.rowconfigure(1, weight=1, uniform="email")

        self.email_top_right_frame_top = ttkb.Frame(self.email_top_right_frame, bootstyle="")
        self.email_top_right_frame_top.grid(row=0, column=0, columnspan=4, sticky='nsew', pady=10)
        self.email_top_right_frame_top.columnconfigure(0, weight=1, uniform="email")
        self.email_top_right_frame_top.columnconfigure(1, weight=1, uniform="email")
        self.email_top_right_frame_top.columnconfigure(2, weight=1, uniform="email")
        self.email_top_right_frame_top.columnconfigure(3, weight=1, uniform="email")
        
        self.email_top_right_frame_top.rowconfigure(0, weight=1, uniform="email")
        self.email_top_right_frame_top.rowconfigure(1, weight=1, uniform="email")

        self.email_top_right_frame_bottom = ttkb.Frame(self.email_top_right_frame, bootstyle="")
        self.email_top_right_frame_bottom.grid(row=1, column=0, columnspan=4, sticky='nsew', pady=10)
        self.email_top_right_frame_bottom.columnconfigure(0, weight=1, uniform="email")
        self.email_top_right_frame_bottom.columnconfigure(1, weight=1, uniform="email")
        self.email_top_right_frame_bottom.columnconfigure(2, weight=1, uniform="email")
        self.email_top_right_frame_bottom.columnconfigure(3, weight=1, uniform="email")

        self.email_top_right_frame_bottom.rowconfigure(0, weight=1, uniform="email")
        self.email_top_right_frame_bottom.rowconfigure(1, weight=1, uniform="email")

        #Middle Right Frame
        self.email_middle_right_frame = ttkb.Frame(self.email_right_frame)
        self.email_middle_right_frame.grid(row=1, column=0, columnspan=2, sticky='nsew')

        self.email_middle_right_frame.columnconfigure(0, weight=1, uniform="email")
        self.email_middle_right_frame.columnconfigure(1, weight=1, uniform="email")

        self.email_middle_right_frame.rowconfigure(0, weight=1, uniform="email")
        self.email_middle_right_frame.rowconfigure(1, weight=1, uniform="email")

        self.email_middle_right_top_frame = ttkb.Frame(self.email_middle_right_frame, bootstyle="")
        self.email_middle_right_top_frame.grid(row=0, column=0, columnspan=4, sticky='nsew', pady=10)
        self.email_middle_right_top_frame.columnconfigure(0, weight=1, uniform="email")
        self.email_middle_right_top_frame.columnconfigure(1, weight=1, uniform="email")
        self.email_middle_right_top_frame.columnconfigure(2, weight=1, uniform="email")
        self.email_middle_right_top_frame.columnconfigure(3, weight=1, uniform="email")

        self.email_middle_right_top_frame.rowconfigure(0, weight=1, uniform="email")
        self.email_middle_right_top_frame.rowconfigure(1, weight=1, uniform="email")

        self.email_middle_right_bottom_frame = ttkb.Frame(self.email_middle_right_frame, bootstyle="")
        self.email_middle_right_bottom_frame.grid(row=1, column=0, columnspan=4, sticky='nsew', pady=10)
        self.email_middle_right_bottom_frame.columnconfigure(0, weight=1, uniform="email")
        self.email_middle_right_bottom_frame.columnconfigure(1, weight=1, uniform="email")
        self.email_middle_right_bottom_frame.columnconfigure(2, weight=1, uniform="email")
        self.email_middle_right_bottom_frame.columnconfigure(3, weight=1, uniform="email")

        self.email_middle_right_bottom_frame.rowconfigure(0, weight=1, uniform="email")
        self.email_middle_right_bottom_frame.rowconfigure(1, weight=1, uniform="email")

        #Bottom Right Frame
        self.email_bottom_right_frame = ttkb.Frame(self.email_right_frame)
        self.email_bottom_right_frame.grid(row=2, column=0, columnspan=2, sticky='nsew')
        self.email_bottom_right_frame.columnconfigure(0, weight=1, uniform="email")
        self.email_bottom_right_frame.columnconfigure(1, weight=1, uniform="email")
        self.email_bottom_right_frame.columnconfigure(2, weight=1, uniform="email")
        self.email_bottom_right_frame.columnconfigure(3, weight=1, uniform="email")

        self.email_bottom_right_frame.rowconfigure(0, weight=1, uniform="email")
        self.email_bottom_right_frame.rowconfigure(1, weight=1, uniform="email")
        
        self.email_bottom_right_frame_top = ttkb.Frame(self.email_bottom_right_frame)
        self.email_bottom_right_frame_top.grid(row=0, column=0, columnspan=4, sticky='nsew', pady=10)
        self.email_bottom_right_frame_top.columnconfigure(0, weight=1, uniform="email")
        self.email_bottom_right_frame_top.columnconfigure(1, weight=1, uniform="email")
        self.email_bottom_right_frame_top.columnconfigure(2, weight=1, uniform="email")
        self.email_bottom_right_frame_top.columnconfigure(3, weight=1, uniform="email")
        
        self.email_bottom_right_frame_top.rowconfigure(0, weight=1, uniform="email")
        self.email_bottom_right_frame_top.rowconfigure(1, weight=1, uniform="email")

        self.email_bottom_right_frame_bottom = ttkb.Frame(self.email_bottom_right_frame)
        self.email_bottom_right_frame_bottom.grid(row=1, column=0, columnspan=4, sticky='nsew', pady=10)
        self.email_bottom_right_frame_bottom.columnconfigure(0, weight=1, uniform="email")
        self.email_bottom_right_frame_bottom.columnconfigure(1, weight=1, uniform="email")
        self.email_bottom_right_frame_bottom.columnconfigure(2, weight=1, uniform="email")
        self.email_bottom_right_frame_bottom.columnconfigure(3, weight=1, uniform="email")

        self.email_bottom_right_frame_bottom.rowconfigure(0, weight=1, uniform="email")
        self.email_bottom_right_frame_bottom.rowconfigure(1, weight=1, uniform="email")
        
        #Bottom Status Frame
        self.email_bottom_status_frame = ttkb.Frame(self.email_opt_frame)
        self.email_bottom_status_frame.grid(row=5, column=0, columnspan=4, sticky='nsew')
        self.email_bottom_status_frame.columnconfigure(0, weight=1, uniform="email")
        self.email_bottom_status_frame.columnconfigure(1, weight=1, uniform="email")
        self.email_bottom_status_frame.columnconfigure(2, weight=1, uniform="email")
        self.email_bottom_status_frame.columnconfigure(3, weight=1, uniform="email")

        self.email_bottom_status_frame.rowconfigure(0, weight=1, uniform="email")

        #Email Opt Widgets
        self.email_label = ttkb.Label(self.email_top_right_frame_top, text="Email Address:", font=("Inter", 13))
        self.email_label.grid(row=0, column=0, columnspan=2, sticky='ew', padx=10)

        self.email_addr_var = tk.StringVar()

        self.new_email_entry = ttkb.Entry(self.email_top_right_frame_top, textvariable=self.email_addr_var, font=("Inter", 12), width=0, state="active")
        self.new_email_entry.grid(row=1, column=0, columnspan=2, sticky='new', padx=10)

        self.email_addr_var.trace('w', self.email_entry_checker)

        self.add_email_button = ttkb.Button(self.email_top_right_frame_top, text="Add", bootstyle="info-outline", command=self.add_email_address, state="disabled")
        self.add_email_button.grid(row=1, column=2, sticky='ew', padx=10)

        self.email_error_var = tk.StringVar()

        self.email_error_label = ttkb.Label(self.email_top_right_frame_bottom, textvariable=self.email_error_var, foreground="#FF0000", font=("Inter", 13))
        self.email_error_label.grid(row=0, column=0, columnspan=2, sticky='ew', padx=10)
        #Cust Opt Widgets
        self.cust_label = ttkb.Label(self.top_right_frame_top, text="Existing Name:", font=("Inter", 13))
        self.cust_label.grid(row=0, column=0, columnspan=2, sticky='ew', padx=10)

        self.new_cust_lbl = ttkb.Label(self.top_right_frame_top, text="New Name", font=("Inter", 13))
        self.new_cust_lbl.grid(row=0, column=2, sticky='ew', padx=10)

        self.cust_entry = ttkb.Entry(self.top_right_frame_top, font=("Inter", 13), width=0, state="disabled")
        self.cust_entry.grid(row=1, column=0, columnspan=2, sticky='new', padx=10)

        self.new_cust_name_var = tk.StringVar()

        self.new_cust_name_entry = ttkb.Entry(self.top_right_frame_top, textvariable=self.new_cust_name_var, font=("Inter", 13), width=0, state="disabled")
        self.new_cust_name_entry.grid(row=1, column=2, columnspan=4, sticky='new', padx=10)

        self.email_options_return_btn = ttkb.Button(self.email_bottom_status_frame, text="Back", bootstyle="danger", command=self.return_to_options)
        self.email_options_return_btn.grid(row=2, column=0, sticky='sew', padx=10, pady=10)

        self.item_label = ttkb.Label(self.top_right_frame_bottom, text="Existing Item #", font=("Inter", 13))
        self.item_label.grid(row=0, column=0, sticky='ew', columnspan=2, padx=10)

        self.new_item_name_label = ttkb.Label(self.middle_right_top_frame, text="Update Item Name:", font=("Inter", 13))
        #self.new_item_name_label.grid(row=0, column=0, sticky='new', padx=10)
        
        self.add_item_label = ttkb.Label(self.middle_right_top_frame, text="Add New Item:", font=("Inter", 13))
        self.add_item_label.grid(row=0, column=0, sticky='ew', columnspan=2, padx=10)

        self.new_item_checker_var = tk.StringVar()
        self.existing_item_checker_var = tk.StringVar()


        #This is when adding a brand new item
        self.add_item_entry = ttkb.Entry(self.middle_right_top_frame, textvariable=self.new_item_checker_var, font=("Inter", 13), width=0, state="disabled")
        self.add_item_entry.grid(row=1, column=0, columnspan=2, sticky='new', padx=10)

        self.new_item_checker_var.trace('w', self.check_new_item_var)
        
        #This is when editing an existing item
        self.new_item_entry = ttkb.Entry(self.middle_right_top_frame, font=("Inter", 13), textvariable=self.existing_item_checker_var, width=0)
        #self.new_item_entry.grid(row=1, column=0, sticky='new', columnspan=2, padx=10)

        self.existing_item_checker_var.trace('w', self.check_existing_item_var)
        
        self.item_var = tk.StringVar()
        self.error_var = tk.StringVar()

        self.item_list = []   

        self.item_combo_box = ttkb.Combobox(
            self.top_right_frame_bottom, 
            textvariable=self.item_var, 
            values=self.item_list, 
            state="disabled", 
            font=('Inter', 12)
            )
        
        self.item_var.trace('w', self.check_item_selected)
        
        self.item_combo_box.grid(row=1, column=0, columnspan=2, sticky='ew', padx=10)

        self.edit_cust_btn = ttkb.Button(self.main_frame, text="Edit Customers", bootstyle="info", command=self.show_edit_cust_frame)
        self.edit_cust_btn.grid(row=0, column=0, sticky='new', padx=10, pady=10)

        self.email_addr_btn = ttkb.Button(self.main_frame, text="Email Destination", bootstyle="info", command=self.show_edit_email_frame)
        self.email_addr_btn.grid(row=0, column=1, sticky='new', padx=10, pady=10)

        self.back_btn = ttkb.Button(self.main_frame, text="Back", bootstyle="danger", command=parent.return_to_home)
        self.back_btn.grid(row=2, column=0, sticky='sew', padx=10, pady=10)

        self.options_return_btn = ttkb.Button(self.bottom_status_frame, text="Back", bootstyle="danger", command=self.return_to_options)
        self.options_return_btn.grid(row=2, column=0, sticky='sew', padx=10, pady=10)

        #This is for existing item
        self.confirm_button = ttkb.Button(self.middle_right_top_frame, text="Confirm", bootstyle="success-outline", command=self.change_item_db_name)
        #self.save_button.grid(row=1, column=2, sticky='ew', padx=10)
        
        #This is for existing item
        self.cancel_button = ttkb.Button(self.middle_right_top_frame, text="Cancel", bootstyle="danger-outline", command=self.cancel_edit)
        #self.cancel_button.grid(row=1, column=3, sticky='ew', padx=10)


        self.edit_item_btn = ttkb.Button(self.top_right_frame_bottom, text="Edit", bootstyle="outline-warning", command=self.edit_item, state="disabled")
        self.edit_item_btn.grid(row=1, column=2, sticky='we', padx=10)

        self.del_item_btn = ttkb.Button(self.top_right_frame_bottom, text="Remove", bootstyle="danger-outline", state="disabled", command=self.delete_item_notice)
        self.del_item_btn.grid(row=1, column=3, sticky='we', padx=10)


        #This is for NEW item
        self.add_button = ttkb.Button(self.middle_right_top_frame, text="Add", bootstyle="info-outline", state="disabled", command=self.add_item)
        self.add_button.grid(row=1, column=2, sticky='ew', padx=10)
        #This is for NEW item
        self.cancel_add_button = ttkb.Button(self.middle_right_top_frame, text="Cancel", bootstyle="danger-outline", state="disabled", command=self.cancel_add_item)
        self.cancel_add_button.grid(row=1, column=3, sticky='ew', padx=10)

        self.selected_label = ttkb.Label(self.scrolled_frame, text="Selected", foreground='#00FF00', font=("Inter", 12))

        self.full_save_btn = ttkb.Button(self.top_right_frame_top, text="Save", bootstyle="success-outline", command=self.update_cust_name, state="disabled")
        self.full_save_btn.grid(row=2, column=0, sticky='new', padx=10, pady= 10)

        self.full_cancel_btn = ttkb.Button(self.bottom_right_frame_bottom, text="Reset", bootstyle="warning-outline", command=self.reset_button_func, state="disabled")
        self.full_cancel_btn.grid(row=0, column=0, rowspan=2, sticky='ew', padx=10)

        self.submitted_label = ttkb.Label(self.middle_right_top_frame, text="Submitted", foreground='#00FF00', font=("Inter", 14))

        self.error_label = ttkb.Label(self.bottom_right_frame_top, foreground="#FF0000", textvariable=self.error_var, font=("Inter", 13))
        self.error_label.grid(row=0, column=0, columnspan=4, sticky='ew', padx=10)

        self.new_cust_name_var.trace('w', self.check_new_cust_name)

    #Email Opt Functions:
    def add_email_address(self):
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()
        c.execute('''
                    INSERT INTO Emails (email_name, email_active) 
                    VALUES (?, ?)
                    ''', (self.new_email_entry.get().lower(), 0))
        conn.commit()
        conn.close()

        self.new_email_entry.delete(0, tk.END)

        for child in self.email_scrolled_frame.winfo_children():
            child.destroy()
        self.fetch_email_addresses()

    def fetch_email_addresses(self):
        self.email_label_dict = {}
        self.email_checkbutton_dict = {}
        self.email_deletebutton_dict = {}
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()
        
        c.execute('''SELECT * FROM Emails''')
        email_results = c.fetchall()
        if email_results:
                for i in email_results: 
                    if i not in self.email_label_dict and i not in self.email_checkbutton_dict:
                        self.email_label_dict[i] = {}
                        self.email_checkbutton_dict[i] = {}
                        self.email_deletebutton_dict[i] = {}
        conn.close()
        self.update_email_scrolled(self.email_label_dict, self.email_checkbutton_dict)

    def update_email_scrolled(self, label_dict, check_button_dict):
        index = 0
        for i in label_dict.items():
            checkbutton_var = tk.IntVar()
            email_label = ttkb.Label(self.email_scrolled_frame, text=i[0][1], font=("Inter", 13))
            #print(i[0][0])
            email_label.grid(row=index, column=0, sticky='nw', padx=10, pady=10)
            checkbutton = ttkb.Checkbutton(self.email_scrolled_frame, variable=checkbutton_var, onvalue=1, offvalue=0, bootstyle="info-round-toggle",
                        command=lambda i = i: self.email_checkbuttons_check(i[0]))
            delete_button = ttkb.Button(self.email_scrolled_frame, text="Delete", bootstyle="outline-danger", command=lambda i=i: self.delete_email(i[0]))
            #print(checkbutton.winfo_class())
            checkbutton_var.set(i[0][2])
            checkbutton.grid(row=index, column=1, sticky='w', padx=10, pady=10)
            delete_button.grid(row=index, column=2, sticky='w', padx=10, pady=10)
            index += 1

    def email_entry_checker(self, *args):
        if self.email_addr_var.get():
            
            self.email_error_var.set('')
            if '@' in self.email_addr_var.get():
                
                self.email_error_var.set('')
                if '.' in self.email_addr_var.get():
                    
                    self.email_error_var.set('')
                    self.add_email_button.configure(state="active")
                else:
                    self.email_error_var.set('Email address is not valid')
                    self.add_email_button.configure(state="disabled")
            else:
                self.email_error_var.set('Email address is not valid')
                self.add_email_button.configure(state="disabled")
        else:
            self.email_error_var.set('')
                

    def delete_email(self, key):
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()
        #print('key[1]:', key[1])
        c.execute('''DELETE FROM Emails
                  WHERE email_name = ?''', (key[1],))
        conn.commit()
        conn.close()

        for child in self.email_scrolled_frame.winfo_children():
            child.destroy()
        self.fetch_email_addresses()
        
 
    def email_checkbuttons_check(self, key):
        new_value = 0
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()
        c.execute('''SELECT * FROM Emails
                  WHERE email_name = ?
                  ''', (key[1],))
        results = c.fetchall()
        if key[2] == 0:
            new_value = 1
        else: 
            new_value = 0
        if results:
            #print('this is results:', results)
            c.execute('''UPDATE Emails
                      SET email_active = ?
                      WHERE email_name = ?
                      ''', (new_value, key[1]))
            conn.commit()
        conn.close()
        
    #Cust Opt Functions:

    def check_new_cust_name(self, *args):
        if self.new_cust_name_var.get():
            self.full_save_btn.configure(state="active")
        else:
            self.full_save_btn.configure(state="disabled")
    def update_cust_name(self):
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()
        c.execute('''SELECT Customers.customer_id
                  FROM Customers
                  WHERE Customers.customer_name = ?''', (self.cust_entry.get().lower(),))
        cust_id = c.fetchall()
        #print('this is the entry:', self.new_cust_name_entry.get().lower())
        if cust_id:
            for i in cust_id:
                #print('i[0] is:',i[0])
                c.execute('''UPDATE Customers
                          SET customer_name = ?
                          WHERE customer_id = ?
                          ''', (self.new_cust_name_entry.get().lower(), i[0]))
                conn.commit()
                self.new_cust_name_entry.delete(0, tk.END)
        conn.close()
        self.reset_button_func()
        self.error_label.configure(foreground="#00FF00")
        self.error_var.set('Customer name successfully updated')

    def delete_cust(self, cust):
        self.delete_cust_window.destroy()
        #print('deleting cust')
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()
        c.execute('''DELETE FROM Customer_Items
        WHERE real_item_id IN (
                  SELECT Customer_Items.real_item_id
                  FROM Customer_Items
                  JOIN Customers ON Customer_Items.customer_id = Customers.customer_id
                  WHERE Customers.customer_name = ?
                )
        ''', (cust.lower(),))
        conn.commit()
        c.execute('''DELETE FROM Customers
                  WHERE customer_name = ?
                  
                  ''', (cust.lower(),))
        conn.commit()
        conn.close()
        self.reset_button_func()

    def cancel_delete_cust(self):
        #print('cancel delete cust')
        self.delete_cust_window.destroy()

    def delete_cust_notice(self, cust_name):
        #print('delete cust notice')
        self.delete_cust_window = ttkb.Toplevel()
        self.delete_cust_window.place_window_center()
        self.delete_cust_window.geometry('600x220')
        self.delete_cust_window.title("Remove Customer")
        self.delete_cust_window.iconbitmap(resource_path("Images\\linamarLogo256.ico"))

        self.delete_cust_window.rowconfigure(0, weight=1, uniform="weight")
        self.delete_cust_window.rowconfigure(1, weight=1, uniform="weight")
        self.delete_cust_window.rowconfigure(2, weight=1, uniform="weight")

        self.delete_cust_window.columnconfigure(0, weight=1, uniform="weight")
        self.delete_cust_window.columnconfigure(1, weight=1, uniform="weight")
        self.delete_cust_window.columnconfigure(2, weight=1, uniform="weight")
        self.delete_cust_window.columnconfigure(3, weight=1, uniform="weight")

        self.delete_warn_label = ttkb.Label(self.delete_cust_window, text="Warning!", foreground="#FFA500", font=("Inter", 15))
        self.delete_warn_label.grid(row=0, column=1, columnspan=2, sticky='')

        self.delete_warn_text = ttkb.Label(self.delete_cust_window, justify="center", text=f"You are about to remove {cust_name.capitalize()} and all {cust_name.capitalize()} Items.\nThis will NOT remove existing {cust_name.capitalize()} orders.\nAre you sure you want to proceed?", font=("Inter", 13))
        self.delete_warn_text.grid(row=1, column=1, columnspan=2, sticky='n')

        self.accept_cust_delete_btn = ttkb.Button(self.delete_cust_window, text="Confirm", bootstyle="success", command=lambda: self.delete_cust(cust_name))
        self.accept_cust_delete_btn.grid(row=2, column=2, columnspan=3, sticky='sew')

        self.cancel_cust_delete_btn = ttkb.Button(self.delete_cust_window, text="Cancel", bootstyle="danger", command=self.cancel_delete_cust)
        self.cancel_cust_delete_btn.grid(row=2, column=0, columnspan=2, sticky='sew')

    def delete_item(self):
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()

        #print('This is cust and item:', self.cust_entry.get(), self.item_var.get())

        c.execute('''DELETE FROM Customer_Items
        WHERE real_item_id IN (
                  SELECT Customer_Items.real_item_id
                  FROM Customer_Items
                  JOIN Customers ON Customer_Items.customer_id = Customers.customer_id
                  WHERE Customers.customer_name = ? AND Customer_Items.real_item_id = ?
                )
        ''', (self.cust_entry.get().lower(), self.item_var.get()))
        conn.commit()
        

        conn.close()
        #print('delete item ran')
        self.fetch_db_items(self.cust_entry.get().lower())
        self.item_var.set('')
        self.delete_item_window.destroy()

        
        self.error_label.configure(foreground='#00FF00')
        self.error_var.set('Item successfully removed')
    
    def cancel_delete_item(self):
        self.delete_item_window.destroy()

    def cancel_add_item(self):
        self.add_item_entry.delete(0, tk.END)

    def add_item(self):
        
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()

        #c.execute('''SELECT item_id FROM Customer_Items WHERE real_item_id = ?''', (self.add_item_entry.get(),))
        #print('in add item func')
        c.execute('''SELECT Customer_Items.item_id, Customer_Items.real_item_id, Customers.customer_id
                  FROM Customer_Items
                  JOIN Customers ON Customer_Items.customer_id = Customers.customer_id
                  WHERE Customer_Items.real_item_id = ? AND Customers.customer_name = ?''', (self.add_item_entry.get(), self.cust_entry.get().lower()))

        results = c.fetchall()
        if results:
            #print('Cust id:', results[0][2])
            #print('item already found', results)

            self.error_label.configure(foreground="#FF0000")
            self.error_var.set(f'Item #{self.add_item_entry.get()} already exists for {self.cust_entry.get()}')
        else:
            #print('item not found, going to search for customer_id and customer name')
            c.execute('''SELECT customer_id FROM Customers WHERE customer_name = ?''', (self.cust_entry.get().lower(),))
            cust_id_results = c.fetchall()
            if cust_id_results:
                #print('cust_id_results', cust_id_results)
                c.execute('''INSERT INTO Customer_Items (real_item_id, customer_id) VALUES (?, ?)''', (self.add_item_entry.get(), cust_id_results[0][0]))
                conn.commit()

                item = self.add_item_entry.get()
                #print('captured item before deletion:', item)
                self.error_label.configure(foreground='#00FF00')
                self.add_item_entry.delete(0, tk.END)
                self.error_var.set(f'Item #{item} successfully added')
                #print('captured item after deletion:', item)
                
                
            self.fetch_db_items(self.cust_entry.get().lower())

        conn.close()

    def delete_item_notice(self):
        #print('delete item notice')
        self.delete_item_window = ttkb.Toplevel()
        self.delete_item_window.place_window_center()
        self.delete_item_window.geometry('600x200')
        self.delete_item_window.title("Remove Item")
        self.delete_item_window.iconbitmap(resource_path("Images\\linamarLogo256.ico"))

        self.delete_item_window.rowconfigure(0, weight=1, uniform="weight")
        self.delete_item_window.rowconfigure(1, weight=1, uniform="weight")
        self.delete_item_window.rowconfigure(2, weight=1, uniform="weight")

        self.delete_item_window.columnconfigure(0, weight=1, uniform="weight")
        self.delete_item_window.columnconfigure(1, weight=1, uniform="weight")
        self.delete_item_window.columnconfigure(2, weight=1, uniform="weight")
        self.delete_item_window.columnconfigure(3, weight=1, uniform="weight")

        self.weight_warn_label = ttkb.Label(self.delete_item_window, text="Warning!", foreground="#FFA500", font=("Inter", 13))
        self.weight_warn_label.grid(row=0, column=1, columnspan=2, sticky='n', pady=10)

        self.weight_warn_text = ttkb.Label(self.delete_item_window, text=f"You are about to remove Item #{self.item_var.get()} from {self.cust_entry.get().capitalize()}.\nAre you sure you want to proceed?", font=("Inter", 13))
        self.weight_warn_text.grid(row=1, column=1, columnspan=2, sticky='n')

        self.accept_btn = ttkb.Button(self.delete_item_window, text="Accept", bootstyle="success", command=self.delete_item)
        self.accept_btn.grid(row=2, column=2, columnspan=3, sticky='sew')

        self.cancel_btn = ttkb.Button(self.delete_item_window, text="Cancel", bootstyle="danger", command=self.cancel_delete_item)
        self.cancel_btn.grid(row=2, column=0, columnspan=2, sticky='sew')

    def reset_cust_options_page(self):
        #self.cust_edit_buttons_dict.clear()
        #self.cust_delete_buttons_dict.clear()
        self.clear_scrolled_frame()
        self.return_to_default()
        
        self.item_var.set('')
        self.error_var.set('')
        self.new_item_checker_var.set('')
        self.existing_item_checker_var.set('')
        self.cust_entry.configure(state="active")
        self.cust_entry.delete(0, tk.END)
        self.cust_entry.configure(state="disabled")
        
        self.add_item_entry.delete(0, tk.END)
        self.new_item_entry.delete(0, tk.END)
        self.new_cust_name_entry.delete(0, tk.END)
        self.new_cust_name_entry.configure(state="disabled")

        self.item_combo_box.configure(state="disabled")
        self.add_item_entry.configure(state="disabled")
        self.submitted_label.grid_remove()
        self.full_save_btn.configure(state="disabled")
        self.full_cancel_btn.configure(state="disabled")
    
    def clear_scrolled_frame(self):
        for child in self.scrolled_frame.winfo_children():
            child.destroy()

        #self.fetch_db_custs()

    def reset_button_func(self):
        self.item_var.set('')
        self.error_var.set('')
        self.new_item_checker_var.set('')
        self.existing_item_checker_var.set('')
        self.cust_entry.configure(state="active")
        self.cust_entry.delete(0, tk.END)
        self.cust_entry.configure(state="disabled")
        
        self.add_item_entry.delete(0, tk.END)
        self.new_item_entry.delete(0, tk.END)
        self.new_cust_name_entry.delete(0, tk.END)
        self.new_cust_name_entry.configure(state="disabled")

        self.item_combo_box.configure(state="disabled")
        self.add_item_entry.configure(state="disabled")
        #self.submitted_label.grid_remove()
        self.full_save_btn.configure(state="disabled")
        self.full_cancel_btn.configure(state="disabled")

        for child in self.scrolled_frame.winfo_children():
            child.destroy()

        self.fetch_db_custs()

        
    def check_item_selected(self, *args):
        if self.item_var.get():
            self.return_to_default()
            
            self.error_label.configure(foreground="#FF0000")
            self.error_var.set('')
            self.edit_item_btn.configure(state="active")
            self.del_item_btn.configure(state="active")

            self.confirm_button.configure(state="disabled")
            #self.cancel_button.configure(state="disabled")
            #print('Item is selected')
            self.submitted_label.grid_remove()
            
        else:
            self.edit_item_btn.configure(state="disabled")
            self.del_item_btn.configure(state="disabled")

    
    def check_new_item_var(self, *args):
        
        #print(self.new_item_checker_var.get())
        if self.new_item_checker_var.get():
            self.cancel_edit()
            self.item_var.set('')
            self.submitted_label.grid_remove()
            if self.new_item_checker_var.get().isdigit():
                if len(self.new_item_checker_var.get()) == 6:
                    self.error_var.set('')
                    self.add_button.configure(state="active")
                    self.cancel_add_button.configure(state="active")
                elif len(self.new_item_checker_var.get()) > 6:
                    self.error_var.set('ERROR: Item number must be 6 digits in total.')
                    self.add_button.configure(state="disabled")
                    self.cancel_add_button.configure(state="disabled")
                else:
                    self.error_var.set('')
                    self.add_button.configure(state="disabled")
                    self.cancel_add_button.configure(state="disabled")
            else:
                self.error_var.set('ERROR: Item must only contain digits.')
                self.add_button.configure(state="disabled")
                self.cancel_add_button.configure(state="disabled")
        else:
            self.error_var.set('')
            self.add_button.configure(state="disabled")
            self.cancel_add_button.configure(state="disabled")

    def check_existing_item_var(self, *args):
        #print(self.existing_item_checker_var.get())
        if self.existing_item_checker_var.get():
            if self.existing_item_checker_var.get().isdigit():
                if len(self.existing_item_checker_var.get()) == 6:
                    self.error_var.set('')
                    self.confirm_button.configure(state="active")
                    self.cancel_button.configure(state="active")

                elif len(self.existing_item_checker_var.get()) > 6:
                    self.error_var.set('ERROR: Item number must be 6 digits in total.')
                    self.confirm_button.configure(state="disabled")
                    #self.cancel_button.configure(state="disabled")
                else:
                    self.error_var.set('')
                    self.confirm_button.configure(state="disabled")
                    #self.cancel_button.configure(state="disabled")
            else:
                self.error_var.set('ERROR: Item must only contain digits.')
                self.confirm_button.configure(state="disabled")
                #self.cancel_button.configure(state="disabled")
        else:
            self.error_var.set('')
            self.confirm_button.configure(state="disabled")
            #self.cancel_button.configure(state="disabled")
  

    def cancel_edit(self):
        self.edit_item_btn.configure(state="active")
        self.del_item_btn.configure(state="active")

        self.new_item_name_label.grid_remove()
        self.confirm_button.grid_remove()
        self.cancel_button.grid_remove()
        self.new_item_entry.grid_remove()

        self.add_button.grid_remove()
        self.cancel_add_button.grid_remove()
        self.add_item_entry.grid_remove()
        self.add_item_label.grid_remove()

        self.add_button = ttkb.Button(self.middle_right_top_frame, text="Add", bootstyle="info-outline", state="disabled", command=self.add_item)
        self.add_button.grid(row=1, column=2, sticky='ew', padx=10)

        self.cancel_add_button = ttkb.Button(self.middle_right_top_frame, text="Cancel", bootstyle="danger-outline", state="disabled", command=self.cancel_add_item)
        self.cancel_add_button.grid(row=1, column=3, sticky='ew', padx=10)

        self.add_item_entry = ttkb.Entry(self.middle_right_top_frame, font=("Inter", 13), textvariable=self.new_item_checker_var, width=0)
        self.add_item_entry.grid(row=1, column=0, columnspan=2, sticky='new', padx=10)

        self.add_item_label = ttkb.Label(self.middle_right_top_frame, text="Add New Item:", font=("Inter", 13))
        self.add_item_label.grid(row=0, column=0, sticky='ew', columnspan=2, padx=10)

        self.new_item_entry.delete(0, tk.END)
 
    def change_item_db_name(self):
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()
        
        c.execute('''SELECT Customers.customer_id, Customer_Items.item_id, Customer_Items.real_item_id
                  FROM Customers
                  JOIN Customer_Items ON Customers.customer_id = Customer_Items.customer_id
                  WHERE Customers.customer_name = ? AND real_item_id = ?''',(self.cust_entry.get().lower(), self.item_var.get()))
        customer_items_results = c.fetchall()
        if customer_items_results:
            for i in customer_items_results:
                #print('this is i012', i[0], i[1], i[2])
                c.execute('''UPDATE Customer_Items
                              SET real_item_id = ?
                              WHERE item_id = ? AND customer_id = ?
                              ''', (self.new_item_entry.get(), i[1], i[0]))
                conn.commit()
        
        c.execute('''SELECT Customers.customer_name, Order_Items.item_id, Order_Items.real_item_id, Customers.customer_id
                  FROM Order_Items
                  JOIN Shipments ON Shipments.shipment_id = Order_Items.shipment_id
                  JOIN Customers ON Shipments.customer_id = Customers.customer_id
                  WHERE Customers.customer_name = ? AND Order_Items.real_item_id = ?
                  ''', (self.cust_entry.get().lower(), self.item_var.get()))
        
        results = c.fetchall()
        #print('this is results:', results)
        if results:
            for i in results:
                #print('this is i[1]', i[1])
                c.execute('''UPDATE Order_Items
                          SET real_item_id = ?
                          WHERE item_id = ?
                          ''', (self.new_item_entry.get(), i[1]))
                conn.commit()
                #print('this is item_var.get()', self.item_var.get())
                
        self.new_item_entry.delete(0, tk.END)
        self.fetch_db_items(self.cust_entry.get().lower())

        self.new_item_name_label.grid_remove()
        self.confirm_button.grid_remove()
        self.cancel_button.grid_remove()
        self.new_item_entry.grid_remove()
        self.submitted_label.grid(row=0, column=0, columnspan=4, padx=10, rowspan=2)
        self.item_var.set('')

        conn.close()

    def activate_edit_btns(self, *args):
        #print(self.item_var.get())
        self.edit_item_btn.configure(state="active")
        self.del_item_btn.configure(state="active")

    def return_to_default(self):
        self.new_item_name_label.grid_remove()
        self.confirm_button.grid_remove()
        self.cancel_button.grid_remove()
        self.new_item_entry.grid_remove()

        self.add_button.grid_remove()
        self.cancel_add_button.grid_remove()
        self.add_item_entry.grid_remove()
        self.add_item_label.grid_remove()

        self.add_button = ttkb.Button(self.middle_right_top_frame, text="Add", bootstyle="info-outline", state="disabled", command=self.add_item)
        self.add_button.grid(row=1, column=2, sticky='ew', padx=10)

        self.cancel_add_button = ttkb.Button(self.middle_right_top_frame, text="Cancel", bootstyle="danger-outline", state="disabled", command=self.cancel_add_item)
        self.cancel_add_button.grid(row=1, column=3, sticky='ew', padx=10)

        self.add_item_entry = ttkb.Entry(self.middle_right_top_frame, font=("Inter", 13), width=0, textvariable=self.new_item_checker_var)
        self.add_item_entry.grid(row=1, column=0, columnspan=2, sticky='new', padx=10)

        self.add_item_label = ttkb.Label(self.middle_right_top_frame, text="Add New Item:", font=("Inter", 13))
        self.add_item_label.grid(row=0, column=0, sticky='ew', columnspan=2, padx=10)

    def edit_item(self):
        self.edit_item_btn.configure(state="disabled")
        self.del_item_btn.configure(state="disabled")

        self.add_button.grid_remove()
        self.cancel_add_button.grid_remove()
        self.add_item_entry.grid_remove()
        self.add_item_label.grid_remove()

        self.new_item_name_label.grid_remove()
        self.confirm_button.grid_remove()
        self.cancel_button.grid_remove()
        self.new_item_entry.grid_remove()

        self.add_button = ttkb.Button(self.middle_right_bottom_frame, text="Add", bootstyle="info-outline", state="disabled", command=self.add_item)
        self.add_button.grid(row=1, column=2, sticky='ew', padx=10)

        self.cancel_add_button = ttkb.Button(self.middle_right_bottom_frame, text="Cancel", bootstyle="danger-outline", state="disabled", command=self.cancel_add_item)
        self.cancel_add_button.grid(row=1, column=3, sticky='ew', padx=10)

        self.add_item_entry = ttkb.Entry(self.middle_right_bottom_frame, font=("Inter", 13), width=0, textvariable=self.new_item_checker_var)
        self.add_item_entry.grid(row=1, column=0, columnspan=2, sticky='new', padx=10)

        self.add_item_label = ttkb.Label(self.middle_right_bottom_frame, text="Add New Item:", font=("Inter", 13))
        self.add_item_label.grid(row=0, column=0, sticky='ew', columnspan=2, padx=10)

        self.new_item_name_label.grid(row=0, column=0, columnspan=2, sticky='ew', padx=10)
        self.confirm_button.grid(row=1, column=2, sticky='ew', padx=10)
        self.cancel_button.grid(row=1, column=3, sticky='ew', padx=10)
        self.new_item_entry.grid(row=1, column=0, sticky='new', columnspan=2, padx=10)

    def fetch_db_items(self, cust):
        self.item_list.clear()
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()
        
        c.execute('''SELECT Customer_Items.real_item_id
                  FROM Customer_Items
                  JOIN Customers ON Customer_Items.customer_id = Customers.customer_id
                  WHERE Customers.customer_name = ?
                  ''', (self.cust_entry.get().lower(),))
        results = c.fetchall()
        for i in results:
            if i not in self.item_list:
                self.item_list.append(i[0])

        #print(self.item_list)
        self.item_combo_box.configure(values=self.item_list)
        conn.close()


    def on_edit_button_click(self, key):
        found = [False]
        #print('edit buttons dict:', self.cust_edit_buttons_dict)
        # Retrieve the exact button using the key
        #print('sql cust list in on edit:', self.sql_cust_list)
        #print('key[1]', key[1])
        for i in self.sql_cust_list:
            if key[1] in i:
                #print('key found', i, 'row:', i[1])
                row_index = i[1]
                edit_btn = self.cust_edit_buttons_dict[i[0]]
                del_btn = self.cust_delete_buttons_dict[i[0]]
                found = [True, i]
                edit_btn.grid_remove()
                del_btn.grid_remove()

                selected_label = ttkb.Label(self.scrolled_frame, text="Selected", foreground='#00FF00', font=("Inter", 12))
                selected_label.grid(row=row_index, column=1, columnspan=2, padx=5)

                self.error_var.set('')
                self.error_label.configure(foreground="#FF0000")

                self.cust_entry.configure(state="active")
                self.cust_entry.insert(tk.END, i[0].capitalize())
                self.cust_entry.configure(state="readonly")

                self.new_cust_name_entry.configure(state="active")
                self.add_item_entry.configure(state="active")
                self.full_cancel_btn.configure(state="active")
                continue
            else:
                #print(self.sql_cust_list)
                edit_btn_to_remove = self.cust_edit_buttons_dict[i[0]]
                edit_btn_to_remove.configure(state="disabled")
                del_btn_to_remove = self.cust_delete_buttons_dict[i[0]]
                del_btn_to_remove.configure(state="disabled")
        if found[0] == True:
            self.item_combo_box.configure(state="readonly")
            self.fetch_db_items(found[1])

        #print(f"Button {key} was clicked! Button text is: {button.cget('text')}")
            

    def on_delete_button_click(self, key):
        button = self.cust_delete_buttons_dict[key[1]]
        #print(f"Button {key} was clicked! Button text is: {button.cget('text')}")
        #print('key1', key[1])
        self.delete_cust_notice(key[1])

    def fetch_db_custs(self):
        self.cust_labels_dict = {}
        self.cust_edit_buttons_dict = {}
        self.cust_delete_buttons_dict = {}
        index = 0
        self.sql_cust_list = []
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()

        c.execute('''SELECT * FROM Customers''')
        results = c.fetchall()
        if results:
            index = 0
            for i in results:
                if i[1] not in self.sql_cust_list:
                    self.sql_cust_list.append((i[1], index))
                if i[1] not in self.cust_labels_dict and i[1] not in self.cust_edit_buttons_dict:
                    self.cust_labels_dict[i[1]] = ttkb.Label(self.scrolled_frame, text=f'{i[1].capitalize()}', font=("Inter", 13))
                    self.cust_edit_buttons_dict[i[1]] = ttkb.Button(self.scrolled_frame, text="Edit", command=lambda i=i: self.on_edit_button_click(i), width=6, bootstyle="warning-outline")
                    self.cust_delete_buttons_dict[i[1]] = ttkb.Button(self.scrolled_frame, text="Delete", command=lambda i=i: self.on_delete_button_click(i), width=6, bootstyle="danger-outline")
                    
                    self.cust_labels_dict[i[1]].grid(row=index, column=0, sticky='nw', padx=10, pady=10)
                    self.cust_edit_buttons_dict[i[1]].grid(row=index, column=1, sticky='n', padx=5, pady=10)
                    self.cust_delete_buttons_dict[i[1]].grid(row=index, column=2, sticky='n', padx=5, pady=10)
                    index += 1  
        conn.close()

    def show_edit_cust_frame(self):
        self.main_frame.grid_remove()
        
        self.fetch_db_custs()
        self.cust_opt_frame.grid(row=0, column=0, rowspan=4, columnspan=4, sticky='nsew')

    def return_to_options(self):
        self.reset_cust_options_page()
        self.cust_opt_frame.grid_remove()
        self.email_opt_frame.grid_remove()
        self.main_frame.grid(row=0, column=0, rowspan=4, columnspan=4, sticky='nsew')

    def show_edit_email_frame(self):
        self.main_frame.grid_remove()
        self.fetch_email_addresses()
        self.email_opt_frame.grid(row=0, column=0, rowspan=4, columnspan=4, sticky='nsew')

class NewAppointment(ttkb.Frame):
    def __init__(self, parent, *args):
        super().__init__(parent, *args)
        self.customer_frame = None
        self.status_frame = None
        
        self.columnconfigure(0, weight=1, uniform="Appointment")
        self.columnconfigure(1, weight=1, uniform="Appointment")
        self.columnconfigure(2, weight=1, uniform="Appointment")
        self.columnconfigure(3, weight=1, uniform="Appointment")
        self.rowconfigure(0, weight=1, uniform="Appointment")
        self.rowconfigure(1, weight=1, uniform="Appointment")
        self.rowconfigure(2, weight=1, uniform="Appointment")
        self.rowconfigure(3, weight=1, uniform="Appointment")
        
        #Main Frame
        self.main_frame = ttkb.Frame(self)
        self.main_frame.grid(row=0, column=0, columnspan=4, rowspan=4, sticky='nsew')
        self.main_frame.columnconfigure(0, weight=1, uniform="Appointment")
        self.main_frame.columnconfigure(1, weight=1, uniform="Appointment")

        self.main_frame.rowconfigure(0, weight=3, uniform="Appointment")
        self.main_frame.rowconfigure(1, weight=3, uniform="Appointment")
        self.main_frame.rowconfigure(2, weight=3, uniform="Appointment")
        self.main_frame.rowconfigure(3, weight=1, uniform="Appointment")
        self.main_frame.rowconfigure(4, weight=1, uniform="Appointment")
        #Left Frame
        self.left_frame = ttkb.Frame(self.main_frame)
        self.left_frame.grid(row=0, column=0, rowspan=3, sticky='nsew')
        self.left_frame.columnconfigure(0, weight=1, uniform="Appointment")
        self.left_frame.columnconfigure(1, weight=1, uniform="Appointment")

        self.left_frame.rowconfigure(0, weight=2, uniform="Appointment")
        self.left_frame.rowconfigure(1, weight=1, uniform="Appointment")
        self.left_frame.rowconfigure(2, weight=1, uniform="Appointment")
        #Right Frame
        self.right_frame = ttkb.Frame(self.main_frame)
        self.right_frame.grid(row=0, column=1, rowspan=3, sticky='nsew')
        self.right_frame.columnconfigure(0, weight=1, uniform="Appointment")
        self.right_frame.rowconfigure(0, weight=1, uniform="Appointment")
        
        #Bottom Status Frame
        self.bottom_status_frame = ttkb.Frame(self.main_frame)
        self.bottom_status_frame.grid(row=3, column=0, columnspan=4, sticky='nsew')
        self.bottom_status_frame.columnconfigure(0, weight=1, uniform="Appointment")
        self.bottom_status_frame.columnconfigure(1, weight=1, uniform="Appointment")
        self.bottom_status_frame.columnconfigure(2, weight=1, uniform="Appointment")
        self.bottom_status_frame.columnconfigure(3, weight=1, uniform="Appointment")
        self.bottom_status_frame.rowconfigure(0, weight=1, uniform="Appointment")

        #Bottom Button Frame
        self.bottom_button_frame = ttkb.Frame(self.main_frame)
        self.bottom_button_frame.grid(row=4, column=0, columnspan=4, sticky='nsew')
        self.bottom_button_frame.columnconfigure(0, weight=1, uniform="Appointment")
        self.bottom_button_frame.columnconfigure(1, weight=1, uniform="Appointment")
        self.bottom_button_frame.columnconfigure(2, weight=1, uniform="Appointment")
        self.bottom_button_frame.columnconfigure(3, weight=1, uniform="Appointment")
        self.bottom_button_frame.rowconfigure(0, weight=1, uniform="Appointment")

        #Top Left Frame
        self.top_left_frame = ttkb.Frame(self.left_frame)
        self.top_left_frame.grid(row=0, column=0, columnspan=2, rowspan=2, sticky='nsew')
        self.top_left_frame.columnconfigure(0, weight=1, uniform="Appointment")
        self.top_left_frame.columnconfigure(1, weight=1, uniform="Appointment")
        self.top_left_frame.rowconfigure(0, weight=1, uniform="Appointment")
        self.top_left_frame.rowconfigure(1, weight=1, uniform="Appointment")
        self.top_left_frame.rowconfigure(2, weight=1, uniform="Appointment")
        self.top_left_frame.rowconfigure(3, weight=1, uniform="Appointment")
        self.top_left_frame.rowconfigure(4, weight=1, uniform="Appointment")
        
        #Bottom Left Frame
        self.bottom_left_Frame = ttkb.Frame(self.left_frame)
        self.bottom_left_Frame.grid(row=2, column=0, columnspan=2, sticky='nsew')
        self.bottom_left_Frame.columnconfigure(0, weight=1, uniform="Appointment")
        self.bottom_left_Frame.columnconfigure(1, weight=1, uniform="Appointment")
        self.bottom_left_Frame.rowconfigure(0, weight=1, uniform="Appointment")

        self.customer_var = tk.StringVar()
        self.item_var = tk.StringVar()
        self.error_var = tk.StringVar()
        self.pre_check_var = tk.StringVar()

        
       

        self.current_cust_list = []
        self.order_list = []
        self.temp_new_order = {}
        
        self.weight_list = []
        self.final_truck_usage = 0
        self.items_list = []
        self.new_weight_list = []
        self.new_cust_dict = {}

        self.validate_weight = self.register(self.validate_weight)
        self.validate_cust_box = self.register(self.validate_cust_box)
        self.validate_item_box = self.register(self.validate_item_box)

        self.cust_check = False
        self.item_check = False

        self.date_row_index = {}

        self.item_list = []
        self.date_row_index_save = {}
        self.dict_for_submit = {}
        self.cust_for_submit = ''

        self.date_row_index_post_check = {}
        self.date_row_index_post_check_save = {}
        self.existing_data_db = {}
        self.data_that_will_overwrite = {}
        self.warning_var = tk.StringVar()
        self.updating_match_list = []
        self.existing_match_list = []
        self.confirmed_matches = []
        self.confirmed_existing_matches = []

        self.all_ship_dates = []
        self.total_weight_dict = {}
        self.weight_decision = "None"
        self.need_to_show_notice = False
        self.weight_dates = []
        self.accepted_weight_dates = []

        self.refresh_token = None
        self.auth_error = tk.StringVar()

        self.option_add('*TCombobox*Listbox.font', "Inter 13")
        

        self.customer_name_label = ttkb.Label(self.top_left_frame, text="Choose Customer:", font=('Inter', 13))
        self.customer_name_label.grid(row=0, column=0, sticky='nw', padx=10, pady=10)

        self.add_item_coldata = [
            #{"text": "Shipment #", "stretch": True},
            {"text": "Date", "stretch": True, },
            {"text": "Item Name", "stretch": True},
            {"text": "Item Weight", "stretch": True},
            {"text": "Total Weight", "stretch": True},
            {"text": "Total Trucks", "stretch": True}
        ]

        style = ttkb.Style()
        style.configure('Treeview.Heading', font=("Inter", 10))
        style.configure('Treeview', font=("Inter", 10))

        self.add_item_table = Tableview(
            master=self.right_frame,
            paginated=False,
            searchable=False,
            coldata=self.add_item_coldata,
            bootstyle="darkly",
            autofit=True,
        )

        #self.add_item_table.bind("<Button-3>", self.disable_right_click)

        self.add_item_table.grid(row=0, column=0, sticky='news', padx=10, pady=10)

        self.weight_var = tk.StringVar()


        self.item_name_label = ttkb.Label(self.top_left_frame, text="Choose Item:", font=('Inter', 13))
        self.item_name_label.grid(row=1, column=0, sticky='nw', padx=10, pady=10)

        self.customer_menu = ttkb.Combobox(self.top_left_frame, textvariable=self.customer_var, values=self.current_cust_list, state="readonly", validate="focus", validatecommand=(self.validate_cust_box, '%P'), font=('Inter', 13))
        self.customer_menu.grid(row=0, column=1, sticky='new', padx=10, pady=10)
        
        self.item_dropdown = ttkb.Combobox(self.top_left_frame, textvariable=self.item_var, values=self.items_list, state="readonly", validate="focus", validatecommand=(self.validate_item_box, '%P'), font=('Inter', 13))
        self.item_dropdown.grid(row=1, column=1, sticky='new', padx=10, pady=10)

        self.add_btn = ttkb.Button(self.bottom_left_Frame, text="Add", bootstyle="info", command=self.email_check)
        self.add_btn.grid(row=0, column=0, sticky='sew', padx=10, pady=10)

        self.reset_btn = ttkb.Button(self.bottom_left_Frame, text="Reset", bootstyle="warning", command=self.reset_frame)
        self.reset_btn.grid(row=0, column=1, sticky='sew', padx=10, pady=10)

        self.submit_btn = ttkb.Button(self.bottom_button_frame, text="Submit", bootstyle="success", command=lambda: self.submite_order2(self.dict_for_submit, self.cust_for_submit))
        self.submit_btn.grid(row=0, column=3, sticky='sew', padx=10, pady=10)

        self.weight_label = ttkb.Label(self.top_left_frame, text="Weight (lbs):", font=('Inter', 13))
        self.weight_label.grid(row=2, column=0, sticky='nw', padx=10, pady=10)

        self.weight_entry = ttkb.Entry(self.top_left_frame, validate="focus", textvariable=self.weight_var, validatecommand=(self.validate_weight, '%P'), font=('Inter', 13))
        self.weight_entry.grid(row=2, column=1, sticky='new', padx=10, pady=10)

        self.date_label = ttkb.Label(self.top_left_frame, text="Enter Date:", font=('Inter', 13))
        self.date_label.grid(row=3, column=0, sticky='nw', padx=10, pady=10)

        calander_style = ttkb.Style()
        calander_style.configure('TCalendar', font=("Inter", 12))

        self.appointment_calander = ttkb.DateEntry(self.top_left_frame, style="TCalendar", bootstyle="info")
        self.appointment_calander.grid(row=3, column=1, sticky='new', padx=10, pady=10)

        self.appointment_calander.entry.configure(font=("Inter", 12))

        self.error_label = ttkb.Label(self.top_left_frame, textvariable=self.error_var, foreground="#FF0000", font=('Inter', 13))
        #self.error_label.grid(row=4, column=0, columnspan=2)

        self.status_label = ttkb.Label(self.bottom_status_frame, textvariable=self.pre_check_var, foreground='#00FF00', font=('Inter', 13))
        self.status_label.grid(row=0, column=0,columnspan=4, sticky='n', pady=10)
        
        self.customer_var.trace('w', self.receive_items)
        self.item_var.trace('w', self.check_item_change)
        #self.weight_var.trace('w', self.check_weight_change)

        self.placeholder_cleared = False

        self.back_btn = ttkb.Button(self.bottom_button_frame, text="Back", bootstyle="danger",command=lambda:[self.reset_frame(), parent.return_to_home()])
        self.back_btn.grid(row=0, column=0, sticky='sew', padx=10, pady=10)

    def set_customer_frame(self, customer_frame):
        self.customer_frame = customer_frame

    def set_status_frame(self, status_frame):
        self.status_frame = status_frame

    def set_notice_frame(self, notice_frame):
        self.notice_frame = notice_frame

    def update_add_items_var(self, *args):
        self.test_add_item_tree.set()

    def disable_right_click(self):
        return "break"

    def email_check(self):
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()

        c.execute('''SELECT * FROM Emails''')
        results = c.fetchall()
        conn.close()
        if results:
            self.dict_input_checker()
        else:
            self.show_email_notice()

    def check_item_change(self, *args):
        self.submit_btn.configure(state="disabled")

    """def check_weight_change(self, *args):
        self.submit_btn.configure(state="disabled")"""
            

    def show_email_notice(self):
        #print('in email_notice')
        self.email_window = ttkb.Toplevel()
        self.email_window.geometry('600x220')
        self.email_window.title("No Email Set")
        self.email_window.place_window_center()
        self.email_window.iconbitmap(resource_path("Images\\linamarLogo256.ico"))

        self.email_window.rowconfigure(0, weight=1, uniform="email")
        self.email_window.rowconfigure(1, weight=1, uniform="email")
        self.email_window.rowconfigure(2, weight=1, uniform="email")

        self.email_window.columnconfigure(0, weight=1, uniform="email")
        self.email_window.columnconfigure(1, weight=1, uniform="email")
        self.email_window.columnconfigure(2, weight=1, uniform="email")
        self.email_window.columnconfigure(3, weight=1, uniform="email")

        self.email_warn_label = ttkb.Label(self.email_window, text="Warning!", foreground="#FFA500", font=("Inter", 15))
        self.email_warn_label.grid(row=0, column=1, columnspan=2, sticky='', pady=10)

        self.email_warn_text = ttkb.Label(self.email_window, justify="center", text="You have not yet set an Email address to receive CSV files.\nNavigate to Options to configure your email recipients.\nPress Continue to ignore.", font=("Inter", 13))
        self.email_warn_text.grid(row=1, column=1, columnspan=2, sticky='n')

        self.email_accept = ttkb.Button(self.email_window, text="Continue", bootstyle="success", command=self.email_confirm_func)
        self.email_accept.grid(row=2, column=2, columnspan=3, sticky='sew')

        self.email_cancel = ttkb.Button(self.email_window, text="Cancel", bootstyle="danger", command=self.email_cancel_func)
        self.email_cancel.grid(row=2, column=0, columnspan=2, sticky='sew')

    def email_confirm_func(self):
        self.email_window.destroy()
        self.dict_input_checker()
    
    def email_cancel_func(self):
        self.email_window.destroy()

    def confirm_changes(self, customer_param, new, existing):
        #print('Confirmed choices')
        self.new_window.destroy()
        self.status_label.configure(foreground="#00FF00")
        self.customer_menu.configure(state="disabled")
        #print('this is new within confirm_changes:', new)
        #print('this is existing within confirm_changes:', existing)
     
        for options in new:
            #print('thisis options:', options)
            if options not in self.confirmed_matches:
                self.confirmed_matches.append(options)
                #print('confirmed matches in confirm_changes:', self.confirmed_matches)
            else:
                continue

        for options2 in existing:
            #print('thisis options2:', options2)
            if options2 not in self.confirmed_existing_matches:
                self.confirmed_existing_matches.append(options2)
                #print('confirmed matches in confirm_changes:', self.confirmed_existing_matches)
            else:
                continue

        self.updating_match_list.clear()
        self.existing_match_list.clear()
        self.pre_check_var.set('Confirmed updates. Press Submit to finalize.')
        self.analyze_list(customer_param)

    def cancel_changes(self):

        self.date_row_index = self.date_row_index_save
        self.date_row_index_post_check = self.date_row_index_post_check_save

        #print('Canceled choices')
        self.new_window.destroy()
        self.status_label.configure(foreground="#FFA500")
        self.pre_check_var.set('Cancelled update.')

        #print('clearing data thatwill overwrite and existing data_db hmm')
        self.data_that_will_overwrite.clear()
        self.existing_data_db.clear()

          
    def fetch_current_db_weight(self, date, cust_param):
        #print('this is date_row_index inside FETCH_CURRENT_DB_WEIGHT:', self.date_row_index)
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()

        chosen_date = date
        chosen_cust_name = cust_param
        c.execute('''SELECT Shipments.date, Shipments.shipment_id, Shipments.customer_id, Customers.customer_id, Customers.customer_name, Shipments.total_weight, Shipments.total_trucks
        FROM Shipments
        JOIN Customers ON Shipments.customer_id = Customers.customer_id
        WHERE Customers.customer_name = ? AND Shipments.date = ?
        ''', (chosen_cust_name, chosen_date))
        
        results = c.fetchall()
        if results:
            rows = results[0]
            total_db_weight = rows[5]
            #print('total_db_weight in fetch:', total_db_weight)
            return total_db_weight
        conn.close()

    def fetch_existing_items_and_weights(self, date, cust_param):
        #print('this is date_row_index inside FETCH_EXISTING_ITEMS_WEIGHTS:', self.date_row_index)
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()
        c.execute('''SELECT Order_Items.real_item_id, Order_Items.weight, Order_Items.shipment_id FROM
        Order_Items
        JOIN Shipments on Order_Items.shipment_id = Shipments.shipment_id
        JOIN Customers on Shipments.customer_id = Customers.customer_id
        WHERE Shipments.date = ? AND Customers.customer_name = ?
        ''', (date, cust_param))

        all_items = c.fetchall()

        #print('this is fetching_existing_items from db. This is what is found:', all_items)
        
        for item, weight, ship in all_items:
            if date not in self.existing_data_db:
                self.existing_data_db[date] = {}
            if date not in self.data_that_will_overwrite:
                self.data_that_will_overwrite[date] = {}
            self.existing_data_db[date][item] = weight
        #print('here is fresh dict:', self.existing_data_db)
        conn.close()
            
            

    def parse_dictionaries(self, cust_param):
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()
        self.date_row_index_post_check_save = deepcopy(self.date_row_index_post_check_save)
        
        for date, date_dict in self.date_row_index.items():
            if date not in self.date_row_index_post_check:
                self.date_row_index_post_check[date] = {}
                #print('these are all the dates in index_post_check:', self.date_row_index_post_check)
                
            total_weight = self.fetch_current_db_weight(date, cust_param)
            no_db_weight = 0
            if total_weight:
                #print(f'was able to find date in database, added "Total_DB_Weight" to {date}')
                self.date_row_index_post_check[date]["Total_DB_Weight"] = total_weight
            else:
                #print(f'was NOT able to find date in database, added "Total_Weight_NoDB" to {date}')
                if "Total_Weight_NoDB" not in self.date_row_index_post_check[date]:

                    self.date_row_index_post_check[date]["Total_Weight_NoDB"] = no_db_weight
                else:
                    #print('updated no_db_weight?')
                    self.date_row_index_post_check[date]["Total_Weight_NoDB"] = no_db_weight
            #Date and Total_db_weight for those dates are inserted
        
        for date, date_dict in self.date_row_index.items():
            self.fetch_existing_items_and_weights(date, cust_param)
            for key, value in date_dict.items():
                c.execute('''SELECT Shipments.date, Shipments.shipment_id, Shipments.customer_id, Customers.customer_id, Customers.customer_name, Shipments.total_weight, Shipments.total_trucks
                FROM Shipments
                JOIN Customers ON Shipments.customer_id = Customers.customer_id
                WHERE Customers.customer_name = ? AND Shipments.date = ?
                ''', (cust_param, date))
                ship_results = c.fetchall()
                if ship_results:
                    ship_rows = ship_results[0]
                    ship_id = ship_rows[1]
                    c.execute('''SELECT Order_Items.real_item_id, Order_Items.weight, Shipments.shipment_id, Order_Items.item_id
                    FROM Order_Items
                    JOIN Shipments ON Order_Items.shipment_id = Shipments.shipment_id
                    WHERE Order_Items.real_item_id = ? AND Shipments.shipment_id = ?
                    ''', (key, ship_id))
                    item_results = c.fetchall()
                    if item_results:
                        item_rows = item_results[0]
                        order_item_name = item_rows[0]
                        order_item_weight = item_rows[1]
                        if key == order_item_name:
                            self.date_row_index_post_check[date][key] = value
                            #print(f'Notice Box: {key} matches with {order_item_name}, {value} will override {order_item_weight}')
                    else:
                        #If item doesnt exist for date
                        self.date_row_index_post_check[date][key] = value
                else:
                    #If shipment doesn't exist for cust
                    if key not in date:
                        #print('Added key to ship date that is not in database')
                        self.date_row_index_post_check[date][key] = value
                    else:
                        #print('Updated item that is in dict but not in database')
                        self.date_row_index_post_check[date][key] = value 
        #print('date_row_index_post_check:', self.date_row_index_post_check)
        self.fetch_total_date_weight(cust_param)
        conn.close()
        
    def check_if_overwriting(self, item, date_param):
        for date, date_dict in self.existing_data_db.items():
            for key, value in date_dict.items():
                if item == key:
                    #print('this is item:', item, 'this is key:', key, 'See if they are matching?')
                    #print('found this weight', value, 'for the item', key)
                    return value
                    
    def fetch_total_date_weight(self, cust_param):
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()
        weight_sum = 0
        for date, date_dict in self.date_row_index_post_check.items():
            weight_sum = 0
            #print('this is date in the dict:', date)
            c.execute('''SELECT Shipments.total_weight FROM Shipments WHERE Shipments.date = ?''', (date,))
            date_results = c.fetchall()
            if date_results:
                #print('this is date_results IMPORTANT', date_results)
                if date not in self.total_weight_dict:
                    self.total_weight_dict[date] = {}
                for i in date_results:
                    weight_sum += i[0]
                    self.total_weight_dict[date]["Total_weight"] = weight_sum
        #print('this is weight dict:', self.total_weight_dict)
        conn.close()
        self.weight_check_full(cust_param)
        
    def weight_check_full(self, cust_param):
        #print('In weight_check_full:')
        sum_with_db = 0
        sum_with_no_db = 0
        stored_db_weight = 0
        dict_date_values = 0
        
        is_overwriting = False
        self.weight_notice = False

        for date, date_dict in self.total_weight_dict.items():
            #print('this is date1', date)
            for date2, date_dict2 in self.date_row_index_post_check.items():
                #print('this is date2', date2)
                if date == date2:
                    self.date_row_index_post_check[date]["Total_date_weight"] = self.total_weight_dict[date]["Total_weight"]
        #print('this is after adding total_date_weight:', self.date_row_index_post_check)

        for date, date_dict in self.date_row_index_post_check.items():
            everything_good = False
            if date in self.total_weight_dict:
                weight_totality = self.total_weight_dict[date]["Total_weight"]
                #print('this is weight totality:', weight_totality)
            sum_with_db = 0
            sum_with_no_db = 0
            if "Total_DB_Weight" in date_dict:
                for key, values in date_dict.items():
                    if key == "Total_DB_Weight":
                        stored_db_weight = values
                        sum_with_db += stored_db_weight
                """if "Total_date_weight" in date_dict:
                    sum_with_db += weight_totality"""

                #print('this is sum_with_db after adding stored and date', sum_with_db)  
                
                for key, values in date_dict.items():  
                    if key != "Total_DB_Weight" and key != "Total_date_weight":
                        over_writing_results = self.check_if_overwriting(key, date)
                        if over_writing_results:
                            is_overwriting = True
                            self.need_to_show_notice = True
                            #print(f'found conflicting weight within the database: {over_writing_results}, going to subtract it from sum, check below')
                            sum_with_db -= over_writing_results
                            #print('after subtracting:', over_writing_results, 'from', sum_with_db)

                        if sum_with_db + values > 0:
                            if sum_with_db + values <= 360000:
                                #print('sum is less than 360k')
                                sum_with_db += values
                                if is_overwriting == True:
                                    #print('overwriting is true')
                                    self.data_that_will_overwrite.setdefault(date, {})[key] = values
  
                            else:
                                #Weight is greater than 360k
                                if date not in self.weight_dates:
                                    self.weight_dates.append(date)
                                self.weight_notice = True
                                sum_with_db += values
                                if is_overwriting == True:
                                    self.data_that_will_overwrite.setdefault(date, {})[key] = values
                                break
                        else:
                            self.error_var.set(f'ERROR: Shipment cannot be below 0 lbs.')
                            break

            elif "Total_Weight_NoDB" in date_dict:
                for key, values in date_dict.items():
                    if key == "Total_Weight_NoDB":
                        #print(f'Found {key} in date_dict')
                        dict_date_values = values
                #print('this is sumnodb before IMPOR:', sum_with_no_db)
                sum_with_no_db += dict_date_values
                #print('this is sumnodb after IMPOR:', sum_with_no_db)
                if "Total_date_weight" in date_dict:
                    sum_with_no_db += weight_totality
                    #print('found Total_date_weight in NoDB, added to sum_with_no_db:', sum_with_no_db)
                for key, values in date_dict.items():  
                    if key != "Total_Weight_NoDB" and key != "Total_date_weight":
                        #print('these are the keys:', key, 'and values:', values, 'IMPORTANT')
                        #print('THIS IS sum_with_no_db:', sum_with_no_db)
                        if sum_with_no_db + values <= 360000:
                            sum_with_no_db += values
                        else:
                            if date not in self.weight_dates:
                                self.weight_dates.append(date)
                            #print(f'ERROR: Shipment {date} current weight is {sum_with_no_db}. Adding {values} exceeds 360k')
                            self.weight_notice = True
                            break

        if self.weight_notice == True:
            #print('THIS IS IS WEIGHT NOTICE IF STATMENT')
            for i in self.weight_dates:
                #print('this is i in weight_dates', i)
                #print('this is acepted_weight_dates:', self.accepted_weight_dates)
                if i not in self.accepted_weight_dates:
                    #print('i is not in accepted_weight-dates')
                    everything_good = False
                    self.weight_notice_decision(cust_param)
                    break
                else:
                    everything_good = True
                    continue
        elif is_overwriting == True:
            #print('is_overwriting elif ran')
            self.handle_notice_data(cust_param)
        else:
            everything_good = True
        if everything_good == True:
            everything_good = False
            self.customer_menu.configure(state="disabled")
            #self.submit_btn.configure(state="active")
            self.analyze_list(cust_param)

    def weight_notice_decision(self, cust_param):
        #print('inside weight_notice decision')
        if self.weight_notice == True:
            self.weight_notice = False
            self.show_weight_notice(cust_param)
        else:
            #print('no weight_notice')            
            #print('Run analysis:')
            if self.need_to_show_notice == True:
                #print('need to show notice is set to TRUE')
                self.handle_notice_data(cust_param)
            else:
                self.customer_menu.configure(state="disabled")
                self.analyze_list(cust_param)
            
    def show_weight_notice(self, cust_param):
        #print('in show_weight_notice')
        self.weight_window = ttkb.Toplevel()
        self.weight_window.geometry('600x200')
        self.weight_window.title("Weight Limit Exceeded")
        self.weight_window.place_window_center()
        self.weight_window.iconbitmap(resource_path("Images\\linamarLogo256.ico"))

        self.weight_window.rowconfigure(0, weight=1, uniform="weight")
        self.weight_window.rowconfigure(1, weight=1, uniform="weight")
        self.weight_window.rowconfigure(2, weight=1, uniform="weight")

        self.weight_window.columnconfigure(0, weight=1, uniform="weight")
        self.weight_window.columnconfigure(1, weight=1, uniform="weight")
        self.weight_window.columnconfigure(2, weight=1, uniform="weight")
        self.weight_window.columnconfigure(3, weight=1, uniform="weight")

        self.weight_warn_label = ttkb.Label(self.weight_window, text="Warning!", foreground="#FFA500", font=("Inter", 13))
        self.weight_warn_label.grid(row=0, column=1, columnspan=2, sticky='n', pady=10)

        self.weight_warn_text = ttkb.Label(self.weight_window, justify="center", text="You are about to set a weight that is over the daily limit of 8 trucks.\nAre you sure you want to proceed?", font=("Inter", 13))
        self.weight_warn_text.grid(row=1, column=1, columnspan=2, sticky='n')

        self.accept_btn = ttkb.Button(self.weight_window, text="Accept", bootstyle="success", command=lambda: self.accept_weight_limit(cust_param))
        self.accept_btn.grid(row=2, column=2, columnspan=3, sticky='sew')

        self.cancel_btn = ttkb.Button(self.weight_window, text="Cancel", bootstyle="danger", command=lambda: self.decline_weight_limit(cust_param))
        self.cancel_btn.grid(row=2, column=0, columnspan=2, sticky='sew')

    def accept_weight_limit(self, cust_param):
        self.weight_decision = 'Accepted'
        self.weight_window.destroy()
        self.accepted_weight_dates = self.weight_dates.copy()
        self.weight_limit_deliberation(cust_param)
        
    def decline_weight_limit(self, cust_param):
        self.weight_decision = 'Declined'
        self.weight_window.destroy()
        self.weight_limit_deliberation(cust_param)

    def weight_limit_deliberation(self, cust_param):
        #print('In weight limit deliberation')
        if self.weight_decision == "Declined":
            #print('weight decision declined within weight_check')
            self.date_row_index = self.date_row_index_save
            self.date_row_index_post_check = self.date_row_index_post_check_save
            self.weight_decision = "None"
            self.weight_entry.delete(0, tk.END)
            self.submit_btn.configure(state="active")
        elif self.weight_decision == "Accepted":
            #print('weight decision accepted within weight_check')
            #print('Run analysis:')
            if self.need_to_show_notice == True:
                self.need_to_show_notice = False
                self.weight_decision = "None"
                self.handle_notice_data(cust_param)
            else:
                self.weight_decision = "None"
                self.customer_menu.configure(state="disabled")
                self.analyze_list(cust_param)
        else:
            print('None got through')

    def handle_notice_data(self, cust_param):
        #print('hand notice ran')
        #print('overwrite dict:', self.data_that_will_overwrite)
        #print('existig data db dict:', self.existing_data_db)
        override_list = []
        existinglist = []

        for date, date_dict in self.data_that_will_overwrite.items():
            for key, value in date_dict.items():
                if (cust_param, date, key, value) not in self.confirmed_matches:
                    #print(f'{cust_param},{date}, {key}, {value} not found in confirmed_matches')
                    override_list.append((cust_param, date, key, value))

        for date2, date_dict2 in self.existing_data_db.items():
            for key2, value2 in date_dict2.items():
                if (cust_param, date2, key2, value2) not in self.confirmed_existing_matches:
                    existinglist.append((cust_param, date2, key2, value2))

        updating_matches = [(o_cust, o_date, o_key, o_value) for (o_cust, o_date, o_key, o_value) in override_list
           for (e_cust, e_date, e_key, e_value) in existinglist
           if (o_cust, o_date, o_key) == (e_cust, e_date, e_key)]

        existing_matches = [(e_cust, e_date, e_key, e_value) for (o_cust, o_date, o_key, o_value) in override_list
            for (e_cust, e_date, e_key, e_value) in existinglist
            if (o_cust, o_date, o_key) == (e_cust, e_date, e_key)]

        if updating_matches and existing_matches:
            #print('ran show notice')
            self.show_notice(updating_matches, existing_matches, cust_param)
        else:
            #print('ran analyze list')
            self.analyze_list(cust_param)

    def show_notice(self, newrowdata, existingrowdata, customer_param):
        #print('inside show notice')
        self.new_window = ttkb.Toplevel()
        self.new_window.geometry("1000x600")
        self.new_window.place_window_center()
        self.new_window.title("Overwrite Existing Item")
        self.new_window.iconbitmap(resource_path("Images\\linamarLogo256.ico"))
    
        self.new_window.columnconfigure(0, weight=1, uniform="window")
        self.new_window.columnconfigure(1, weight=1, uniform="window")
        self.new_window.columnconfigure(2, weight=1, uniform="window")
        self.new_window.columnconfigure(3, weight=1, uniform="window")
        self.new_window.rowconfigure(0, weight=1, uniform="window")
        self.new_window.rowconfigure(1, weight=1, uniform="window")
        self.new_window.rowconfigure(2, weight=1, uniform="window")
        self.new_window.rowconfigure(3, weight=1, uniform="window")
        self.new_window.rowconfigure(4, weight=1, uniform="window")
        self.new_window.rowconfigure(5, weight=1, uniform="window")
    
        self.current_label = ttkb.Label(self.new_window, text="New Values:", font=('Inter', 14))
        self.current_label.grid(row=0, column=0, columnspan=2, sticky='')

        self.existing_label = ttkb.Label(self.new_window, text="Existing Database Values:", font=('Inter', 14))
        self.existing_label.grid(row=0, column=2, columnspan=4, sticky='')

        self.warning_text = self.warning_var.get()
        #print('this is self.warning_var', self.warning_var.get())
        #print('this is self_warning_text:', self.warning_text)
        
        coldata = [
            {"text": "Company", "stretch": True},
            {"text": "Date", "stretch": True},
            {"text": "Item", "stretch": True},
            {"text": "Weight", "stretch": True}
        ]
        coldata2 = [
            {"text": "Company", "stretch": True},
            {"text": "Date", "stretch": True},
            {"text": "Item", "stretch": True},
            {"text": "Weight", "stretch": True}
        ]

        self.the_table1 = Tableview(
            master=self.new_window,
            coldata=coldata,
            rowdata=newrowdata,
            bootstyle="darkly",
            autofit=True,
            autoalign=True

        )

        self.the_table1.autofit_columns()
        self.the_table2 = Tableview(
            master=self.new_window,
            coldata=coldata2,
            rowdata=existingrowdata,
            #paginated=True,
            #searchable=True,
            bootstyle="darkly",
            autofit=True,
            autoalign=True 
        )
        self.the_table2.autofit_columns()

        self.the_table1.grid(row=1, column=0, rowspan=4, columnspan=2, sticky='nsew', padx=2)
        self.the_table2.grid(row=1, column=2, rowspan=4, columnspan=4, sticky='nsew', padx=2)
        
        self.warning_label = ttkb.Label(self.new_window, justify="center", text="Warning: You will overwrite the following weight. Press Confirm to update item with new weight.", foreground="#FFBC11", font=('Inter', 13))
        self.warning_label.grid(row=5, column=0, columnspan=4, padx=2)

        self.confirm_button = ttkb.Button(self.new_window, text="Confirm", bootstyle="success", command=lambda: self.confirm_changes(customer_param, newrowdata, existingrowdata))
        self.confirm_button.grid(row=6, column=3, sticky='sew', padx=2)

        self.cancel_button = ttkb.Button(self.new_window, text="Cancel", bootstyle="danger", command=self.cancel_changes)
        self.cancel_button.grid(row=6, column=0, sticky='sew', padx=2)

        self.new_window.mainloop()

    def validate_cust_box(self, val)-> bool:
        if val.isalpha():
            self.cust_check = True
            self.pre_check_var.set('')
            self.error_var.set('')
            return True
        else:
            return False
    
    def validate_item_box(self, val)-> bool:
        if val.isnumeric():
            self.pre_check_var.set('')
            self.error_var.set('')
            self.item_check = True
            return True
        else:
            return False

    def validate_entry(self):
        if self.weight_entry.get() == "":
            self.pre_check_var.set('')
            self.error_var.set('')
            self.submit_btn.configure(state="disabled")

    def validate_weight(self, val) -> bool:
        if val.isdigit():  
            self.pre_check_var.set('')
            self.error_var.set('')
            return True   
        else: 
            return False
        
    

    #Searches sql db for existing customer sets
    def receieve_customer_names(self):
        self.options_menu_width = 30
        path = resource_path('work.db')
        check_file = os.path.isfile(path)
        if check_file == True:
            conn = sqlite3.connect(resource_path('work.db'))
            c = conn.cursor()
            c.execute('''SELECT customer_name FROM Customers''')
            customers = c.fetchall()

            self.error_var.set('')
            if not self.placeholder_cleared:
                self.placeholder_cleared = True
            for (customer_name,) in customers:
                if customer_name not in self.current_cust_list:
                    self.current_cust_list.append(customer_name)
                    self.customer_menu.config(values=self.current_cust_list)
                    self.new_cust_dict[customer_name] = {}
                    #print(f"This is the current dict in receive customers: {self.new_cust_dict}")
                    #print(self.current_cust_list)
            conn.close()
        else:
            print('File not existing')
    #A trace on customer_var is then ran in a sql query to correspond the appropriate items
    def receive_items(self, *args):
        self.submit_btn.configure(state="disabled")
        self.error_var.set('')
        selected_customer = self.customer_var.get().lower()
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()

        c.execute('''SELECT * FROM Customers''')
        customers = c.fetchall()
        items_ok = []
        for customer in customers:
            customer_id, customer_name = customer
            if customer_name == selected_customer:
                c.execute('''SELECT * FROM Customer_Items WHERE customer_id = ?''', (customer_id,))
                items = c.fetchall()
                for item in items:
                    item_id, real_item_id, customer_id = item[0], item[1], item[2]
                    self.items_list.append(item[1])
                    self.order_list.append(item[1])
                    items_ok.append(item[1])
                    self.item_dropdown.config(values=items_ok)
        conn.close()

    def reset_frame(self):
        path = resource_path('work.db')
        check_file = os.path.isfile(path)
        if check_file:
            #print(f'in reset frame, selected var: {self.item_var.get()}')
            self.pre_check_var.set('Cancelled update.')

            self.all_ship_dates.clear()
            self.total_weight_dict.clear()
            self.weight_decision = "None"
            self.need_to_show_notice = False
            self.weight_dates.clear()
            self.accepted_weight_dates.clear()

            self.date_row_index_post_check_save.clear()
            self.existing_data_db.clear()
            self.data_that_will_overwrite.clear()
            self.warning_var.set('')
            self.pre_check_var.set('')
            self.error_var.set('')
            self.updating_match_list.clear()
            self.existing_match_list.clear()
            self.confirmed_matches.clear()
            self.confirmed_existing_matches.clear()
            self.all_ship_dates.clear()
            self.item_list.clear()
            self.date_row_index_save.clear()
            self.dict_for_submit.clear()
            self.cust_for_submit = ''

            self.date_row_index.clear()
            self.date_row_index_post_check.clear()

            self.add_item_table.delete_rows()
            self.add_item_table.load_table_data()
            self.customer_var.set('')
            self.item_var.set('')
            self.weight_entry.delete(0, tk.END)

            self.customer_menu.configure(state="readonly")
            self.cust_check = False
            self.item_check = False
            self.submit_btn.configure(state="disabled")
        else:
            print("nothing to clear")

    
    def analyze_list(self, customer_param):
        self.consolidated = {}
        current_index = 0  # Start with a unique index
        
        for date, entries in self.date_row_index.items():
            if date not in self.consolidated:
                self.consolidated[date] = {}
            for key, value in entries.items():
                self.consolidated[date][current_index] = {key: value} # Add current entry
                current_index += 1
            # Insert a break after the last entry of the date
            self.consolidated[date][current_index] = {'break'}
            current_index += 1
        total_sum = 0

        for date, line_id_dict in self.consolidated.items():
            total_sum = 0
            if "Total_weight" not in line_id_dict.items():
                self.consolidated[date]["Total_weight"] = 0
                self.consolidated[date]["Total_trucks"] = 0
            for line, line_data in line_id_dict.items():
                if isinstance(line_data, dict):
                    total_sum += list(line_data.values())[0]
                    self.consolidated[date]["Total_weight"] = total_sum
                    self.consolidated[date]["Total_trucks"] = float(Decimal(self.consolidated[date]["Total_weight"] / 45000).quantize(Decimal("0.00")))
        self.update_tree(self.consolidated, customer_param)  

    def update_tree(self, the_dict, customer_param):
        self.add_item_table.delete_rows()
        self.add_item_table.autofit_columns()
        self.add_item_table.load_table_data()
        update_tree_dict = the_dict

        for date, line_id_dict in the_dict.items(): 
            for line_id, line_dict in line_id_dict.items():
                if isinstance(line_dict, dict):
                    self.add_item_table.insert_row(line_id, (date, list(line_dict.keys())[0], list(line_dict.values())[0], '---', '---'))
                    self.add_item_table.autofit_columns()
                else: 
                    if line_id != "Total_weight" and line_id != "Total_trucks":
                        if "break" in line_dict:
                            self.add_item_table.insert_row(line_id, ('---','---', '---',the_dict[date]["Total_weight"], the_dict[date]["Total_trucks"]))
                            self.add_item_table.autofit_columns()
            
        self.add_item_table.load_table_data()
        self.submit_btn.configure(state='active')
        #.set(self.item_var.get())

        #self.item_var.set('')
        self.weight_entry.delete(0, tk.END)

        self.dict_for_submit = update_tree_dict
        self.cust_for_submit = customer_param

    def dict_pre_check(self, customer_param, item_param, date_param, weight_param):
        self.date_row_index_save = deepcopy(self.date_row_index)
        # Check if the date already exists
        if date_param not in self.date_row_index:
            # If not, create a new entry for the date
            self.date_row_index[date_param] = {item_param: weight_param}
        else:
            # If the date exists, check if the item exists
            if item_param in self.date_row_index[date_param]:
                # If the item exists, update its weight
                self.date_row_index[date_param][item_param] = weight_param
            else:
                # If the item does not exist, add it to the date
                self.date_row_index[date_param][item_param] = weight_param

        self.parse_dictionaries(customer_param)

    def dict_input_checker(self):
        selected_item = self.item_var.get()
        selected_date = tk.StringVar()
        selected_date = self.convert_date(self.appointment_calander.entry.get())
        selected_customer = self.customer_var.get()
        selected_weight = self.weight_entry.get()

        if selected_customer and self.cust_check == True:
            if selected_item and self.item_check == True:
                if selected_weight:
                    if selected_weight.isdigit():
                        selected_weight = int(self.weight_entry.get())
                        #if selected_weight <= 360000:
                        if selected_date:
                            self.dict_pre_check(selected_customer, selected_item, selected_date, selected_weight)
 
    def convert_date(self, date):
        date_parts = date.split('/')
        if len(date_parts) == 3:
            month = date_parts[0] 
            day = date_parts[1]
            year = date_parts[2]
        if len(year) == 2:
            year = '20' + year
        new_date = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
        return new_date
    
    def create_csv_send_email(self, cust_param):
        multi_ship = False
        if len(self.all_ship_dates) > 1:
            multi_ship = True
        ship_ids = []
        #print('this is self.all_ship_dates within create_csv:', self.all_ship_dates)
        #print('csv func ran')

        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()

        today = strftime("%Y-%m-%d-%I-%M-%S%p", localtime())
        file_name = f"{today}_email_copy.csv"
        email_reports_dir = os.path.join(resource_path("Email_Reports"))
        file_path = os.path.join(resource_path(email_reports_dir), file_name)

        with open(file_path, mode='w', newline='') as csvfile:
            fieldnames = ['Shipment_ID', 'Customer_Name', 'Item_ID', 'Weight', 'Date', 'Total_Weight', 'Total_Trucks']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    
            for i in self.all_ship_dates:
                #print('i in csv:', i)
                c.execute('''SELECT Customers.customer_id, Shipments.date, Shipments.customer_id, Shipments.shipment_id, Order_Items.shipment_id, Shipments.total_weight, Shipments.total_trucks, Order_Items.real_item_id, Order_Items.weight
                FROM Customers
                JOIN Shipments ON Customers.customer_id = Shipments.customer_id
                JOIN Order_Items ON Shipments.shipment_id = Order_Items.shipment_id
                WHERE Customers.customer_name = ? AND Shipments.date = ?
                ''', (cust_param, i))
                results = c.fetchall()

                if results:
                    #print('found results in csv')
                    row = results[0]
                    cust_id = row[0]
                    ship_date = row[1]
                    ship_id = row[3]
                    ship_total_weight = row[5]
                    ship_total_trucks = row[6]
                    
                    if multi_ship == True:
                        ship_ids.append(ship_id)
                    writer.writeheader()

                    for j in results:
                        real_item_name = j[7]
                        item_weight = j[8]
                        #print('this is j', j)
                        writer.writerow({"Shipment_ID": ship_id, "Customer_Name": cust_param, "Item_ID": real_item_name, "Weight": item_weight, "Date": i})
                    
                    writer.writerow({"Total_Weight": ship_total_weight, "Total_Trucks": ship_total_trucks})
                    writer.writerow({})

        

        toast = ToastNotification(title="File Saved", message=f"File: {today}_email_copy.csv has been saved in Linamar Planner/dist/Email_Reports", position=(), duration=15000, bootstyle="info")
        toast.show_toast()

        c.execute('''SELECT * FROM Emails''')

        results = c.fetchall()
        if results:
            self.status_label.configure(foreground="#FFA500")
            self.pre_check_var.set('Sending email...')
            for i in results:
                if i[2] == 1:
                    #print('Is the email active:', i[1], i[2])
                    threading.Thread(target=self.send_email, args=(cust_param, file_path, ship_ids, i[1])).start()
                
        else:
            #Clear everything
            self.all_ship_dates.clear()
            self.total_weight_dict.clear()
            self.weight_decision = "None"
            self.need_to_show_notice = False
            self.weight_dates.clear()
            self.accepted_weight_dates.clear()

            self.date_row_index_post_check_save.clear()
            self.date_row_index_post_check.clear()
            self.existing_data_db.clear()
            self.data_that_will_overwrite.clear()
            self.warning_var.set('')
            self.pre_check_var.set('')
            self.error_var.set('')
            self.updating_match_list.clear()
            self.existing_match_list.clear()
            self.confirmed_matches.clear()
            self.confirmed_existing_matches.clear()

            self.all_ship_dates.clear()
            self.item_list.clear()
            self.date_row_index_save.clear()
            self.dict_for_submit.clear()
            self.cust_for_submit = ''
            self.date_row_index.clear()
            
            self.add_item_table.delete_rows()
            self.add_item_table.autofit_columns()
            self.add_item_table.load_table_data()

            self.customer_var.set('')
            self.item_var.set('')
            self.weight_entry.delete(0, tk.END)

            self.new_cust_dict.clear()
            self.customer_menu.configure(state="readonly")
            self.cust_check = False
            
            self.item_check = False

            self.current_cust_list.clear()
            self.receieve_customer_names()
            self.receive_items()
            self.status_label.configure(foreground='#00FF00')
            self.pre_check_var.set('Successfully submitted.')
        
        conn.close()

    def send_email(self, cust_param, file_path, ship_list, recipient):
        load_dotenv()
        EMAIL_CODE = os.getenv("EMAIL_CODE")
        today = date.today()
        def ship_ids_list(the_ship_list):
            my_string = ''
            for i in the_ship_list:
                my_string += f'#{i}\n'
                #print('this is my_string:', my_string)
            return my_string
        
        def ship_ids_subject(the_ship_list):
            my_string2 = ''
            for i, data in enumerate(the_ship_list):
                if data != the_ship_list[-1]:
                    my_string2 += f'#{i}, '
                else:
                    my_string2 += f'#{i}'
            return my_string2

        subject_for_ships = ship_ids_subject(ship_list)
        email = 'edgar.apache1@gmail.com'
        receiver_email = recipient
        subject = f'Report: {today} | Shipment(s): {subject_for_ships}'
        message_for_ships = ship_ids_list(ship_list)
        message = f'Customer: {cust_param.capitalize()}'
        
        msg = MIMEMultipart()
        msg['From'] = email
        msg['To'] = receiver_email
        msg['Subject'] = subject
        body = MIMEText(message, 'plain')
        msg.attach(body)

        with open(file_path, 'r') as f:
            attachment = MIMEApplication(f.read(), Name=basename(file_path))
            attachment['Content-Disposition'] = 'attachment; filename="{}"'.format(basename(file_path))

        msg.attach(attachment)
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(email, EMAIL_CODE)
        server.sendmail(email, receiver_email, msg.as_string())
        #print("email has been sent")

        #Clear everything
        self.all_ship_dates.clear()
        self.total_weight_dict.clear()
        self.weight_decision = "None"
        self.need_to_show_notice = False
        self.weight_dates.clear()
        self.accepted_weight_dates.clear()

        self.date_row_index_post_check_save.clear()
        self.existing_data_db.clear()
        self.data_that_will_overwrite.clear()
        self.warning_var.set('')
        self.pre_check_var.set('')
        self.error_var.set('')
        self.updating_match_list.clear()
        self.existing_match_list.clear()
        self.confirmed_matches.clear()
        self.confirmed_existing_matches.clear()

        self.all_ship_dates.clear()
        self.item_list.clear()
        self.date_row_index_save.clear()
        self.dict_for_submit.clear()
        self.cust_for_submit = ''
        self.date_row_index.clear()
        self.date_row_index_post_check.clear()
        
        self.add_item_table.delete_rows()
        self.add_item_table.autofit_columns()
        self.add_item_table.load_table_data()

        self.customer_var.set('')
        self.item_var.set('')
        self.weight_entry.delete(0, tk.END)

        self.new_cust_dict.clear()
        self.customer_menu.configure(state="readonly")
        self.cust_check = False
        
        self.item_check = False

        self.current_cust_list.clear()
        self.receieve_customer_names()
        self.receive_items()

        self.status_label.configure(foreground='#00FF00')
        self.pre_check_var.set('Successfully submitted.')


    def submite_order2(self, the_dict, customer_param):
        #print('cust param in submitorder2:', customer_param)
        working_dict = the_dict
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()

        total_weight = 0
        total_trucks = 0

        c.execute('''SELECT customer_id FROM Customers WHERE customer_name = ?''', (customer_param,))
        result = c.fetchone()

        if result:
            #print('databse found existing customer')
            for date, date_dict in working_dict.items():
                #print(date, date_dict)
                the_date = date
                self.all_ship_dates.append(date)
                for key, value in date_dict.items():
                    if key == "Total_weight":
                        total_weight = value
                        
                        #print('total_weight:', total_weight)
                    if key == "Total_trucks":
                        total_trucks = value

                c.execute('''SELECT Shipments.date, Shipments.shipment_id, Shipments.customer_id, Customers.customer_id, Customers.customer_name, Shipments.total_weight, Shipments.total_trucks
                FROM Shipments
                JOIN Customers ON Shipments.customer_id = Customers.customer_id
                WHERE Customers.customer_name = ? AND Shipments.date = ?
                ''', (customer_param, date))

                results = c.fetchall()
                #If customer already has the shipdate associated to them
                if results:
                    #print('was able to find the customer based on the ship date')
                    row = results[0]
                    shipment_id = row[1]
                    custtomer_id = row[3]
                    ship_total_weight = row[5]
                    ship_total_trucks = row[6] 
                    
                    #Searches if the item - ship_id already exists
                    for key, value in date_dict.items():
                        if isinstance(value, dict):
                            c.execute('''SELECT Order_Items.real_item_id, Order_Items.weight, Shipments.shipment_id, Order_Items.item_id
                            FROM Order_Items
                            JOIN Shipments ON Order_Items.shipment_id = Shipments.shipment_id
                            WHERE Order_Items.real_item_id = ? AND Shipments.shipment_id = ?
                            ''', (list(value.keys())[0], shipment_id))

                            item_found = c.fetchall()
                            #Update the associated items
                            if item_found:
                                existing_item_dict = {}
                                #print(f'found {list(value.keys())[0]} in {shipment_id}')
                                row = item_found[0]
                                item_weight = row[1]
                                item_id = row[3]

                                c.execute('''SELECT Shipments.date, Shipments.shipment_id, Shipments.customer_id, Customers.customer_id, Customers.customer_name, Shipments.total_weight, Shipments.total_trucks
                                FROM Shipments
                                JOIN Customers ON Shipments.customer_id = Customers.customer_id
                                WHERE Customers.customer_name = ? AND Shipments.date = ?
                                ''', (customer_param, date))

                                ifship_results = c.fetchall()
                                ifrows = ifship_results[0]
                                if_updated_total_weight = ifrows[5]
                                
                                new_correct_total_weight = 0
                                new_correct_total_weight = if_updated_total_weight
                                new_correct_total_weight -= item_weight
                                new_correct_total_weight += list(value.values())[0]

                                new_correct_total_trucks = 0
                                new_correct_total_trucks = float(Decimal(new_correct_total_weight / 45000).quantize(Decimal("0.00")))

                                if list(value.keys())[0] not in existing_item_dict:
                                    existing_item_dict[list(value.keys())[0]] = {'Customer': customer_param, 'Date': date, 'Weight': list(value.values())[0], 'Ship_id': ifrows[1]}

                                c.execute('''UPDATE Order_Items
                                SET weight = ?
                                WHERE item_id = ?
                                ''', (list(value.values())[0], item_id))
                                conn.commit()

                                c.execute('''UPDATE Shipments
                                SET total_weight = ?, total_trucks = ?
                                WHERE shipment_id = ?
                                ''', (new_correct_total_weight, new_correct_total_trucks, shipment_id))
                                conn.commit()
                                
                                self.status_label.configure(foreground="#FFA500")
                                self.pre_check_var.set('Creating calander event...')
                                
                                threading.Thread(target=self.create_outlook_event, args=(customer_param, existing_item_dict)).start()
                            #If the ship_id exists but item not in ship, adds item
                            else:
                                existing_ship_dict = {}
                                
                                #print('This ran. Ship_id exists, but item not in ship')
                                c.execute('''SELECT Shipments.date, Shipments.shipment_id, Shipments.customer_id, Customers.customer_id, Customers.customer_name, Shipments.total_weight, Shipments.total_trucks
                                FROM Shipments
                                JOIN Customers ON Shipments.customer_id = Customers.customer_id
                                WHERE Customers.customer_name = ? AND Shipments.date = ?
                                ''', (customer_param, date))
                                elseship_results = c.fetchall()
                                rows = elseship_results[0]
                                updated_total_weight = rows[5]
                                #print('supposedly updated total_weight:', updated_total_weight)

                                new_correct_total_weight2 = updated_total_weight

                                if list(value.keys())[0] not in existing_ship_dict:
                                    existing_ship_dict[list(value.keys())[0]] = {'Customer': customer_param, 'Date': date, 'Weight': list(value.values())[0], 'Ship_id': rows[1]}
                                
                                #print('this is the retrieved total_weight after first item update:', new_correct_total_weight2)
                                new_correct_total_weight2 += list(value.values())[0]
                                #print('this is the retrieved total_weight after adding the new item weight:', list(value.values())[0], new_correct_total_weight2)
                                

                                new_correct_total_trucks2 = 0
                                new_correct_total_trucks2 = float(Decimal(new_correct_total_weight2 / 45000).quantize(Decimal("0.00")))
                                #print('Ran insert into order items if item not in shipment. Inserted:', list(value.keys())[0], shipment_id, list(value.values())[0])
                                c.execute('''INSERT INTO Order_Items (real_item_id, shipment_id, weight) VALUES (?, ?, ?)''',(
                                    list(value.keys())[0], shipment_id, list(value.values())[0]
                                ))

                                c.execute('''UPDATE Shipments
                                SET total_weight = ?, total_trucks = ?
                                WHERE shipment_id = ?
                                ''', (new_correct_total_weight2, new_correct_total_trucks2, shipment_id))
                                conn.commit() 

                                #self.create_outlook_event(customer_param, existing_ship_dict)
                                self.status_label.configure(foreground="#FFA500")
                                self.pre_check_var.set('Creating calander event...')
                                
                                threading.Thread(target=self.create_outlook_event, args=(customer_param, existing_ship_dict)).start()
                else:
                    item_dict = {}
                    #If customer exists, but Item not associated with selected ship_date
                    c.execute('''INSERT INTO Shipments (customer_id, total_weight, total_trucks, date) VALUES (?, ?, ?, ?)''',(
                        result[0], total_weight, total_trucks, date
                    ))
                    new_ship_id = c.lastrowid

                    for key, value in date_dict.items():
                        if isinstance(value, dict):
                            c.execute('''INSERT INTO Order_Items (real_item_id, shipment_id, weight) VALUES (?, ?, ?)''',(
                                list(value.keys())[0], new_ship_id, list(value.values())[0]
                            ))
                            if list(value.keys())[0] not in item_dict:
                                item_dict[list(value.keys())[0]] = {'Customer': customer_param, 'Date': date, 'Weight': list(value.values())[0], 'Ship_id': new_ship_id}
                                
                    conn.commit()
                    #print('this is the dict im working with:', item_dict)
                    """for item, item_dict_values in item_dict.items():
                        print('this is item:', item)
                        print('this is cust name:', item_dict[item]['Customer'])"""
                    #self.create_outlook_event(customer_param, item_dict)

                    self.status_label.configure(foreground="#FFA500")
                    self.pre_check_var.set('Creating calander event...')
                    
                    threading.Thread(target=self.create_outlook_event, args=(customer_param, item_dict)).start()

            conn.commit()
        else:
            #If new customer
            #print('database added new customer. psst: this should NEVER run')
            c.execute('''INSERT INTO Customers (customer_name) VALUES (?)''', (customer_param,))
            cust_id = c.lastrowid
            conn.commit()

            for date, date_dict in working_dict.items():
                self.all_ship_dates.append(date)

                the_date = date
                for key, value in date_dict.items():
                    if key == "Total_weight":
                        total_weight = value
                        #print('total_weight:', total_weight)
                    if key == "Total_trucks":
                        total_trucks = value
    
                c.execute('''INSERT INTO Shipments (customer_id, total_weight, total_trucks, date) VALUES (?, ?, ?, ?)''',(
                    cust_id, total_weight, total_trucks, date
                ))
                ship_id = c.lastrowid

                for key, value in date_dict.items():
                    if isinstance(value, dict):
                        c.execute('''INSERT INTO Order_Items (real_item_id, shipment_id, weight) VALUES (?, ?, ?)''',(
                            list(value.keys())[0], ship_id, list(value.values())[0]
                        ))
                conn.commit()
            conn.commit()

        conn.close()
        self.create_csv_send_email(customer_param)
        self.submit_btn.configure(state="disabled")

    def submit_auth_code(self, application_id, client_secret, scopes):
        client = msal.ConfidentialClientApplication(
            client_id=application_id,
            client_credential=client_secret,
            authority='https://login.microsoftonline.com/consumers'
        )
        authorization_code = self.auth_code_entry.get()

        if not authorization_code:
            raise ValueError('Auth code is empty')
        
        token_response = client.acquire_token_by_authorization_code(
            code = authorization_code,
            scopes=scopes
        )

        if 'access_token' in token_response:
            if 'refresh_token' in token_response:
                with open('refresh_token.txt', 'w') as file:
                    """self.error_text.configure(foreground='#00FF00')
                    self.auth_error.set('Success')"""
                    #print('success!')
                    file.write(token_response['refresh_token'])
            self.auth_code_window.destroy()
            return token_response['access_token']
        else:
            self.error_text.configure(foreground="#FF0000")
            self.auth_error.set("ERROR: The code you have entered is incorrect.")
            
            #raise Exception('Failed to acquire acccess token:', + str(token_response))
         

    def no_access_token_yet(self, application_id, client_secret, scopes):
        
        
        self.auth_code_window = ttkb.Toplevel()
        self.auth_code_window.place_window_center()
        self.auth_code_window.geometry('620x300')
        self.auth_code_window.title("Authorization Code")
        self.auth_code_window.iconbitmap(resource_path("Images\\linamarLogo256.ico"))

        self.auth_code_window.rowconfigure(0, weight=1, uniform="auth")
        self.auth_code_window.rowconfigure(1, weight=1, uniform="auth")
        self.auth_code_window.rowconfigure(2, weight=1, uniform="auth")
        self.auth_code_window.rowconfigure(3, weight=1, uniform="auth")

        self.auth_code_window.columnconfigure(0, weight=1, uniform="auth")
        self.auth_code_window.columnconfigure(1, weight=1, uniform="auth")
        self.auth_code_window.columnconfigure(2, weight=1, uniform="auth")
        self.auth_code_window.columnconfigure(3, weight=1, uniform="auth")

        self.label_entry_frame = ttkb.Frame(self.auth_code_window)
        self.label_entry_frame.grid(row=1, column=0, columnspan=4, sticky='nsew')
        self.label_entry_frame.columnconfigure(0, weight=1, uniform="auth")
        

        self.auth_warn_label = ttkb.Label(self.auth_code_window, text="Action Needed", foreground="#FFA500", font=("Inter", 15))
        self.auth_warn_label.grid(row=0, column=1, columnspan=2, sticky='', padx=10)

        self.auth_warn_text = ttkb.Label(self.label_entry_frame, text=f"Enter in the Authorization Code located in the URL of the opened browser:", font=("Inter", 13))
        self.auth_warn_text.grid(row=1, column=0, columnspan=2, sticky='ew', padx=10)

        self.auth_code_entry = ttkb.Entry(self.label_entry_frame, font=("Inter", 13))
        self.auth_code_entry.grid(row=2, column=0, sticky='ew', columnspan=3, padx=10, pady=5)

        self.error_text = ttkb.Label(self.auth_code_window, justify="center", foreground="#FF0000",textvariable=self.auth_error, font=("Inter", 13))
        self.error_text.grid(row=2, column=0, columnspan=4)

        self.submit_auth_btn = ttkb.Button(self.auth_code_window, text="Confirm", bootstyle="success", command=lambda: self.submit_auth_code(application_id, client_secret, scopes))
        self.submit_auth_btn.grid(row=3, column=2, columnspan=3, sticky='sew')

        self.cancel_auth_btn = ttkb.Button(self.auth_code_window, text="Cancel", bootstyle="danger")
        self.cancel_auth_btn.grid(row=3, column=0, columnspan=2, sticky='sew')

        client = msal.ConfidentialClientApplication(
            client_id=application_id,
            client_credential=client_secret,
            authority='https://login.microsoftonline.com/consumers'
        )

        auth_request_url = client.get_authorization_request_url(scopes)
        webbrowser.open(auth_request_url)

    def generate_access_token(self, application_id, client_secret, scopes):
        client = msal.ConfidentialClientApplication(
            client_id=application_id,
            client_credential=client_secret,
            authority='https://login.microsoftonline.com/consumers'
        )
        
        if os.path.exists(resource_path('refresh_token.txt')):
            #print('found refreshtoken')
            with open('refresh_token.txt', 'r') as file:
                self.refresh_token = file.read().strip()
                #print('this is refresh token since the file exists:', self.refresh_token)
        if self.refresh_token:
            token_response = client.acquire_token_by_refresh_token(self.refresh_token, scopes=scopes)
        else:
            self.no_access_token_yet(application_id, client_secret, scopes)

        if 'access_token' in token_response:
            if 'refresh_token' in token_response:
                with open('refresh_token.txt', 'w') as file:
                    file.write(token_response['refresh_token'])
            return token_response['access_token']
        else:
            #print('failed to acquire access token')
            raise Exception('Failed to acquire acccess token:', + str(token_response))
        
    def access_token_main(self):
        load_dotenv()
        APPLICATION_ID = os.getenv('APPLICATION_ID')
        CLIENT_SECRET = os.getenv('CLIENT_SECRET')
        SCOPES = ['User.Read', 'Calendars.ReadWrite']

        #print('app_id:', APPLICATION_ID)
        #print('client_secret:', CLIENT_SECRET)
        #print('scopes:', SCOPES)

        try:
            access_token = self.generate_access_token(application_id=APPLICATION_ID, client_secret=CLIENT_SECRET, scopes=SCOPES)
            #print('this is access token within main:', access_token)
            headers = {
                'Authorization': 'Bearer ' + access_token
            }
            #print(headers)
        except Exception as e:
            print('Error on 66:', e)

    def create_outlook_event(self, cust_name, the_item_dict):
        MS_GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0'
        load_dotenv()
        APPLICATION_ID = os.getenv('APPLICATION_ID')
        CLIENT_SECRET = os.getenv('CLIENT_SECRET')
        SCOPES = ['Calendars.ReadWrite']

        endpoint= f'{MS_GRAPH_BASE_URL}/me/calendar/events'
        #print('this is item_dict in outlook event:', the_item_dict)
        for item, item_values in the_item_dict.items():
            item_name = item
            #Customer
            cust = the_item_dict[item]['Customer']
            weight = the_item_dict[item]['Weight']
            item_date = the_item_dict[item]['Date']
            ship_id = the_item_dict[item]['Ship_id']
            trucks = float(Decimal(weight / 45000).quantize(Decimal("0.00")))
            
            try:
                access_token = self.generate_access_token(
                    application_id=APPLICATION_ID,
                    client_secret=CLIENT_SECRET,
                    scopes=SCOPES
                )
                headers = {
                    'Authorization': 'Bearer ' + access_token
                }

                def construct_event_detail(event_name, **event_details):
                    request_body = {
                        'subject': f'{event_name.capitalize()}-{item_name}-{weight}-({trucks})-ID:{ship_id}',
                        'isAllDay': True
                    }
                    for key, val in event_details.items():
                        request_body[key] = val
                    return request_body

                event_name = cust

                now = datetime.datetime.now()
                military_time = now.strftime("%H:%M:%S")

                start_day_obj = datetime.datetime.strptime(item_date, "%Y-%m-%d")
                end_day_obj = start_day_obj + timedelta(days=1)
                start = {
                    'dateTime': f'{start_day_obj.strftime("%Y-%m-%d")}T00:00:00',
                    'timeZone': 'UTC'
                }

                end = {
                    'dateTime': f'{end_day_obj.strftime("%Y-%m-%d")}T00:00:00',
                    'timeZone': 'UTC'
                }

                payload = construct_event_detail(event_name, start=start, end=end)
                #print('Event payload:', payload)
                location = 'Tokyo, Japan'

                create_event = httpx.post(endpoint, headers=headers, json=construct_event_detail(event_name, start=start, end=end))
                if create_event.status_code != 202:
                    print('Code:', create_event.status_code, 'Error:', create_event.text)
            except httpx.HTTPStatusError as e:
                print('HTTP Error:', e)
            except Exception as e:
                print('Error:', e)
    
class NewCustomer(ttkb.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        self.home_frame = None
        self.status_frame = None

        self.columnconfigure(0, weight=1, uniform="Customer")
        self.columnconfigure(1, weight=1, uniform="Customer")
        self.columnconfigure(2, weight=1, uniform="Customer")
        self.columnconfigure(3, weight=1, uniform="Customer")
        self.rowconfigure(0, weight=1, uniform="Customer")
        self.rowconfigure(1, weight=1, uniform="Customer")
        self.rowconfigure(2, weight=1, uniform="Customer")
        self.rowconfigure(3, weight=1, uniform="Customer")

        #Main Frame
        self.main_frame = ttkb.Frame(self)
        self.main_frame.grid(row=0, column=0, columnspan=4, rowspan=4, sticky='nsew')
        self.main_frame.columnconfigure(0, weight=1, uniform="Customer")
        self.main_frame.columnconfigure(1, weight=1, uniform="Customer")
        self.main_frame.rowconfigure(0, weight=3, uniform="Customer")
        self.main_frame.rowconfigure(1, weight=3, uniform="Customer")
        self.main_frame.rowconfigure(2, weight=1, uniform="Customer")
        self.main_frame.rowconfigure(3, weight=1, uniform="Customer")
        self.main_frame.rowconfigure(4, weight=1, uniform="Customer")
        
        #Left Frame
        self.left_frame = ttkb.Frame(self.main_frame)
        self.left_frame.grid(row=0, column=0, rowspan=3, sticky='nsew')
        self.left_frame.columnconfigure(0, weight=1, uniform="Customer")
        self.left_frame.columnconfigure(1, weight=1, uniform="Customer")
        self.left_frame.rowconfigure(0, weight=2, uniform="Customer")
        self.left_frame.rowconfigure(1, weight=1, uniform="Customer")
        self.left_frame.rowconfigure(2, weight=1, uniform="Customer")
        
        #Right Frame
        self.right_frame = ttkb.Frame(self.main_frame)
        self.right_frame.grid(row=0, column=1, rowspan=3, sticky='nsew')
        self.right_frame.columnconfigure(0, weight=1, uniform="Customer")
        self.right_frame.rowconfigure(0, weight=1, uniform="Customer")
        
        #Bottom Status Frame
        self.bottom_status_frame = ttkb.Frame(self.main_frame)
        self.bottom_status_frame.grid(row=3, column=0, columnspan=4, sticky='nsew')
        self.bottom_status_frame.columnconfigure(0, weight=1, uniform="Customer")
        self.bottom_status_frame.columnconfigure(1, weight=1, uniform="Customer")
        self.bottom_status_frame.columnconfigure(2, weight=1, uniform="Customer")
        self.bottom_status_frame.columnconfigure(3, weight=1, uniform="Customer")
        self.bottom_status_frame.rowconfigure(0, weight=1, uniform="Customer")

        #Bottom Button Frame
        self.bottom_button_frame = ttkb.Frame(self.main_frame)
        self.bottom_button_frame.grid(row=4, column=0, columnspan=4, sticky='nsew')
        self.bottom_button_frame.columnconfigure(0, weight=1, uniform="Customer")
        self.bottom_button_frame.columnconfigure(1, weight=1, uniform="Customer")
        self.bottom_button_frame.columnconfigure(2, weight=1, uniform="Customer")
        self.bottom_button_frame.columnconfigure(3, weight=1, uniform="Customer")
        self.bottom_button_frame.rowconfigure(0, weight=1, uniform="Customer")

        #Top Left Frame
        self.top_left_frame = ttkb.Frame(self.left_frame)
        self.top_left_frame.grid(row=0, column=0, columnspan=2, sticky='nsew')
        self.top_left_frame.columnconfigure(0, weight=1, uniform="Customer")
        self.top_left_frame.columnconfigure(1, weight=1, uniform="Customer")
        self.top_left_frame.rowconfigure(0, weight=1, uniform="Customer")
        self.top_left_frame.rowconfigure(1, weight=1, uniform="Customer")
        self.top_left_frame.rowconfigure(2, weight=1, uniform="Customer")

        #Top Middle Frame
        self.top_middle_frame = ttkb.Frame(self.left_frame)
        self.top_middle_frame.grid(row=1, column=0, columnspan=2, sticky='nsew')
        self.top_middle_frame.columnconfigure(0, weight=1, uniform="Customer")
        self.top_middle_frame.columnconfigure(1, weight=1, uniform="Customer")
        self.top_middle_frame.rowconfigure(0, weight=1, uniform="Customer")

        #Bottom Left Frame
        self.bottom_left_Frame = ttkb.Frame(self.left_frame)
        self.bottom_left_Frame.grid(row=2, column=0, columnspan=2, sticky='nsew')
        self.bottom_left_Frame.columnconfigure(0, weight=1, uniform="Customer")
        self.bottom_left_Frame.columnconfigure(1, weight=1, uniform="Customer")
        self.bottom_left_Frame.rowconfigure(0, weight=1, uniform="Customer")

        #Variables
        self.error_var = tk.StringVar()
        self.show_label_var = tk.StringVar()
        self.show_label_var.set("")
        self.success_label_var = tk.StringVar()
        self.existing_names = []
        self.items = []
        self.item_entry_var = tk.StringVar()
        self.customer_var = tk.StringVar()

        self.database_customers = []
        self.has_created_db = False

        self.item_list = []
        self.customer_name = ""


        self.customer_name_label2 = ttkb.Label(self.top_left_frame, text="Customer:", font=('Inter', 13))
        self.customer_name_label2.grid(row=0, column=0, sticky='nw', padx=10, pady=10)
        
        self.customer_entry = ttkb.Entry(self.top_left_frame, font=("Inter", 12), textvariable=self.customer_var)
        self.customer_entry.grid(row=0, column=1, sticky='new', padx=10, pady=10)

        self.item_name_label = ttkb.Label(self.top_left_frame, text="Item:", font=('Inter', 13))
        self.item_name_label.grid(row=1, column=0, sticky='nw', padx=10, pady=10)

        self.item_entry = ttkb.Entry(self.top_left_frame, textvariable=self.item_entry_var, font=("Inter", 12))
        self.item_entry.grid(row=1, column=1, sticky='new', padx=10, pady=10)

        self.listbox_object = tk.Listbox(self.right_frame, height=10, bd=1, font=('Inter', 13))
        self.listbox_object.grid(row=0, column=0, sticky='news', padx=10, pady=10)

        self.error_label = ttkb.Label(self.top_middle_frame, textvariable=self.error_var, foreground="#FF0000", font=('Inter', 13))
        self.error_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        self.add_item_btn = ttkb.Button(self.bottom_left_Frame, text="Add", bootstyle="info", command=self.send_customer_and_items_and_weight)
        self.add_item_btn.grid(row=0, column=0, sticky='sew', padx=10, pady=10)

        self.new_customer_btn = ttkb.Button(self.bottom_left_Frame, text="Reset", bootstyle="warning", command=self.reset_page)
        self.new_customer_btn.grid(row=0, column=1, sticky='sew', padx=10, pady=10)

        self.success_label = ttkb.Label(self.bottom_status_frame, textvariable=self.success_label_var, foreground='#00FF00', font=('Inter', 13))
        self.success_label.grid(row=0, column=0, columnspan=4, padx=10, pady=20)

        self.submit_btn = ttkb.Button(self.bottom_button_frame, text="Submit",bootstyle="success", command=lambda: self.submit_final_info(), state="disabled")
        self.submit_btn.grid(row=0, column=3, sticky='sew', padx=10, pady=10)

        self.back_btn = ttkb.Button(self.bottom_button_frame, text="Back", bootstyle="danger", command=lambda: (self.reset_page(), parent.return_to_home()))
        self.back_btn.grid(row=0, column=0, sticky='swe', padx=10, pady=10)

        self.item_entry_var.trace('w', self.new_item_check)
        self.customer_var.trace('w', self.check_cust_entry)

    def set_status_frame(self, status_frame):
        self.status_frame = status_frame

    def set_appointment_frame(self, appointment_frame):
        self.appointment_frame = appointment_frame

    def check_cust_entry(self, *args):
        self.success_label_var.set('')

    def new_item_check(self, *args):
        self.submit_btn.configure(state="disabled")

    def submit_final_info(self): 
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()
        
        c.execute('''CREATE TABLE IF NOT EXISTS Order_Items(
            item_id INTEGER PRIMARY KEY AUTOINCREMENT,
            real_item_id TEXT NOT NULL,
            shipment_id INTEGER,
            weight INTEGER,
            FOREIGN KEY (shipment_id) REFERENCES Shipments(shipment_id)
        )''')
        conn.commit()

        c.execute('CREATE INDEX IF NOT EXISTS idx_shipments_date ON Shipments(date)')
        conn.commit()

        the_customer_name = self.existing_names[0]
        
        c.execute('''INSERT INTO Customers (customer_name) VALUES (?)''', (the_customer_name,))
        conn.commit()

        cust_id = c.lastrowid
        
        for item in self.item_list:
            cust_items = item
            c.execute('''INSERT INTO Customer_Items (real_item_id, customer_id) VALUES (?, ?)''',(cust_items,cust_id))

            conn.commit()
        conn.close()
        
        self.appointment_frame.receieve_customer_names()
        self.has_created_db = True
        self.existing_names.clear()
        
        self.show_label_var.set('Customer Name:')
        self.listbox_object.delete(0, tk.END)
        self.customer_entry.configure(state="active")
        self.customer_entry.delete(0, tk.END)
        self.item_entry.delete(0, tk.END)
        self.error_var.set('')
        self.item_list.clear()
        self.existing_names.clear()
        self.submit_btn.configure(state="disabled")
        self.items.clear()
        self.success_label_var.set('Successfully Submitted')
    def check_customer_name_db(self):
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()
        c.execute('''SELECT customer_name FROM Customers''')
        customer_names = c.fetchall()
        if customer_names:
            for (customer_name,) in customer_names:
                if customer_name not in self.database_customers:
                    self.database_customers.append(customer_name)
                    #print(f"Customers received from db: {self.database_customers}")
        else:
            conn.close()
        conn.close()


    def check_cust_exists_main_db(self, name):
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()

        c.execute('''SELECT name FROM sqlite_master WHERE type="table" AND name="Customers"''')
        customers_exists = c.fetchone()

        if customers_exists:
            c.execute('''SELECT customer_name FROM Customers WHERE customer_name = ?''', (name.lower(),))
            exist_result = c.fetchone()
            conn.close()
            if exist_result:
                #print(f'Found {name} in Customers table')
                return True
            else:
                #print(f'Did not find {name} in Customers table')
                return False
        else:
            #print('Customers table does not exist')
            return False

    def check_existing_items(self, name, item):
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()
        the_name = name.lower()
        
        if self.check_cust_exists_main_db(the_name) == True:
            #print('Confirmed customer is in legit database, need to check for existing items')
            c.execute('''SELECT name FROM sqlite_master WHERE type="table" AND name="Order_Items"''')
            table_exists = c.fetchone()

            c.execute('''SELECT * FROM Customers
                      WHERE customer_name = ?''', (the_name,))
            cust_found = c.fetchall()
            if cust_found:
                return "Found"
            else:
                return "Not_Found"

        conn.close()

    def reset_page(self):
        self.customer_entry.configure(state="active")
        self.existing_names.clear()
        self.items.clear()
        self.error_var.set('')
        self.success_label_var.set('')
        self.show_label_var.set('')
        self.listbox_object.delete(0, tk.END)
        self.customer_entry.delete(0, tk.END)
        self.item_entry.delete(0, tk.END)
        self.item_list.clear()
        self.submit_btn.configure(state="disabled")
        self.check_customer_name_db()

    def listbox_frame_retrieve_item(self, name, item):
        self.itemlist_option = f"Item: #{item}"
        self.listbox_object.insert(tk.END, self.itemlist_option)
        self.show_label_var.set(name.capitalize())
        self.is_sent_to_dropdown = False
        self.customer_name = name
        self.item_list.append(item)

    def send_customer_and_items_and_weight(self):
        name = self.customer_entry.get()
        item = self.item_entry.get()
        
        if self.has_created_db == True:
            self.check_customer_name_db()
        self.success_label_var.set('')

        if self.check_existing_items(name, item) == 'Found':
            #print('Did not submit because item found in either temp or legit')
            self.error_var.set(f'ERROR: {name.capitalize()} already exists in the database.\nTo add new items to {name.capitalize()} navigate to the Options menu')
        else:
            if name:
                name = str(name.lower())
                if name.isalnum():
                    if name not in self.database_customers:
                        if item:
                            if item.isalnum():
                                if len(item) == 6:
                                    if item not in self.items:
                                        self.existing_names.append(name)
                                        if name == self.existing_names[0]:
                                            self.error_var.set('')
                                            self.items.append(item)
                                            #print(self.items)
                                            self.listbox_frame_retrieve_item(name, item)
                                            self.customer_entry.configure(state="readonly")
                                            self.item_entry.delete(0, tk.END)
                                            self.submit_btn.configure(state="active")
                                        else:
                                            self.error_var.set(f'ERROR: Can\'t change name of current customer {self.existing_names[0]}. \nTo startover press \'Reset\'')
                                            #print(self.existing_names)
                                            #print(name)
                                    else:
                                        self.error_var.set(f'ERROR: Item {item} already added')
                                        #print('Item already added')
                                        #print(self.items)
                                        self.item_entry.delete(0, tk.END)
                                else:
                                    self.error_var.set('ERROR: Items must be 6 digits long')
                            else:
                                self.error_var.set('ERROR: Item cannot contain symbols')
                        else:
                            self.error_var.set('ERROR: You must provide an item')
                    else:
                        self.error_var.set(f"ERROR: {name.capitalize()} is already an existing customer.\nTo add new items to {name.capitalize()}, go to Options")
                        self.customer_entry.delete(0, tk.END)
                        self.item_entry.delete(0, tk.END)
                else:
                    self.error_var.set("ERROR: Customer name cannot contain symbols")
            else:
                self.error_var.set("ERROR: No customer name provided")
                self.item_entry.delete(0, tk.END)


class CheckStatus(ttkb.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        self.customer_frame = None
        self.appointment_frame = None
       
        #Main Column & Row Config
        self.columnconfigure(0, weight=1, uniform="main")
        self.columnconfigure(1, weight=1, uniform="main")
        self.columnconfigure(2, weight=1, uniform="main")
        self.columnconfigure(3, weight=1, uniform="main")
        self.rowconfigure(0, weight=1, uniform="main")
        self.rowconfigure(1, weight=1, uniform="main")
        self.rowconfigure(2, weight=1, uniform="main")
        self.rowconfigure(3, weight=1, uniform="main")

        self.main_frame = ttkb.Frame(self)
        self.main_frame.grid(row=0, column=0, rowspan=4, columnspan=4, sticky='nsew')
        self.main_frame.rowconfigure(0, weight=1, uniform="main")
        self.main_frame.rowconfigure(1, weight=1, uniform="main")
        self.main_frame.rowconfigure(2, weight=1, uniform="main")
        self.main_frame.columnconfigure(0, weight=1, uniform="main")
        self.main_frame.columnconfigure(1, weight=1, uniform="main")

        self.left_frame = ttkb.Frame(self.main_frame)
        self.left_frame.grid(row=0, column=0, rowspan=2, sticky='nsew')
        self.left_frame.rowconfigure(0, weight=2, uniform="main")
        self.left_frame.rowconfigure(1, weight=1, uniform="main")
        self.left_frame.columnconfigure(0, weight=1, uniform="main")

        self.right_frame = ttkb.Frame(self.main_frame)
        self.right_frame.grid(row=0, column=1, rowspan=2, sticky='nsew')
        self.right_frame.rowconfigure(0, weight=1, uniform="main")
        self.right_frame.columnconfigure(0, weight=1, uniform="main")

        self.bottom_frame = ttkb.Frame(self.main_frame)
        self.bottom_frame.grid(row=2, column=0, columnspan=2, sticky='nsew')
        self.bottom_frame.rowconfigure(0, weight=1, uniform="main")
        self.bottom_frame.columnconfigure(0, weight=1, uniform="main")
        self.bottom_frame.columnconfigure(1, weight=1, uniform="main")
        self.bottom_frame.columnconfigure(2, weight=1, uniform="main")
        self.bottom_frame.columnconfigure(3, weight=1, uniform="main")
       
        #Left_Frame Column & Row Config
        self.left_top_frame = ttkb.Frame(self.left_frame)
        self.left_top_frame.grid(row=0, column=0, sticky='nsew')
        self.left_top_frame.columnconfigure(0, weight=1, uniform="main")
        self.left_top_frame.columnconfigure(1, weight=1, uniform="main")
        self.left_top_frame.rowconfigure(0, weight=1, uniform="main")
        self.left_top_frame.rowconfigure(1, weight=1, uniform="main")
        self.left_top_frame.rowconfigure(2, weight=1, uniform="main")
        self.left_top_frame.rowconfigure(3, weight=1, uniform="main")
        #Left_bottom_Frame
        self.left_bottom = ttkb.Frame(self.left_frame)
        self.left_bottom.grid(row=1, column=0, sticky='nsew')
        self.left_bottom.columnconfigure(0, weight=1, uniform="main")
        self.left_bottom.columnconfigure(1, weight=1, uniform="main")
        self.left_bottom.columnconfigure(2, weight=1, uniform="main")
        self.left_bottom.rowconfigure(0, weight=1, uniform="main")
        self.left_bottom.rowconfigure(1, weight=1, uniform="main")
        self.left_bottom.rowconfigure(2, weight=1, uniform="main")
        self.left_bottom.rowconfigure(3, weight=1, uniform="main")

        self.cust_var = tk.StringVar()
        self.item_var = tk.StringVar()
        self.error_var = tk.StringVar()

        self.option_add('*TCombobox*Listbox.font', "Inter 13")
        self.validate_item = self.register(self.validate_item)
        self.validate_customer = self.register(self.validate_customer)

        self.cust_list = []
        self.item_list = []
        self.item_search_var = tk.IntVar()
        self.item_exists = False
        self.cust_exists = False
        self.customer_search_var = tk.IntVar()
        self.date_search_var = tk.IntVar()

        self.customer_label = ttkb.Label(self.left_top_frame, text="Customer:", font=("Inter", 13))
        self.customer_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')

        self.customer_combo = ttkb.Combobox(
            self.left_top_frame, 
            textvariable=self.cust_var, 
            values=self.cust_list, state="readonly", 
            validate="focus", 
            validatecommand=(self.validate_customer, '%P'), 
            font=('Inter', 13))
        self.customer_combo.grid(row=0, column=1, sticky='new', padx=10, pady=10)

        self.from_label = ttkb.Label(self.left_top_frame, text="From:", font=("Inter", 13))
        self.from_label.grid(row=1, column=0, padx=10, pady=10, sticky='nw')

        self.to_label = ttkb.Label(self.left_top_frame, text="To:", font=("Inter", 13))
        self.to_label.grid(row=2, column=0, padx=10, pady=10, sticky='nw')

        self.item_label = ttkb.Label(self.left_top_frame, text="Item:", font=("Inter", 13))
        self.error_label = ttkb.Label(self.left_bottom, textvariable=self.error_var, foreground="#FF0000", font=("Inter", 13))
        self.error_label.grid(row=2, column=0, columnspan=4, sticky='n')
        coldata = [
            {"text": "Company", "stretch": True},
            {"text": "Shipment ID", "stretch": True},
            {"text": "Date", "stretch": True},
            {"text": "Item", "stretch": True},
            {"text": "Weight", "stretch": True},
            {"text": "Shipment Weight", "stretch": True},
            {"text": "Total Trucks", "stretch": True}
        ]
        
        self.generate_table = Tableview(
            master=self.right_frame,
            coldata=coldata,
            bootstyle="dark",
            autofit=True,
            autoalign=True  
        )
        self.generate_table.grid(row=0, rowspan=5, column=0, columnspan=5, sticky='nsew', padx=10, pady=10)

        self.item_combo_box = ttkb.Combobox(
            self.left_top_frame, 
            textvariable=self.item_var, 
            values=self.item_list, 
            state="readonly", 
            validate="focus", 
            validatecommand=(self.validate_item, '%P'), 
            font=('Inter', 13))

        self.from_calander = ttkb.DateEntry(self.left_top_frame, bootstyle="info")
        self.from_calander.grid(row=1, column=1, padx=10, pady=10, sticky='new')

        self.from_calander.entry.configure(font=("Inter", 12))
        
        self.to_calander = ttkb.DateEntry(self.left_top_frame, bootstyle="info")
        self.to_calander.entry.configure(font=("Inter", 12))
        self.to_calander.grid(row=2, column=1, padx=10, pady=10, sticky='new')

        style = ttkb.Style()
        style.configure('TCheckbutton', font=("Inter 12"))

        self.item_search = ttkb.Checkbutton(
            self.left_bottom,
            bootstyle="info",
            text="Item Search", 
            variable=self.item_search_var, 
            onvalue=1, 
            offvalue=0,
            command=self.checkbuttons_check)
            

        self.customer_search = ttkb.Checkbutton(
            self.left_bottom, 
            bootstyle="info",
            text="Customer Search", 
            variable=self.customer_search_var, 
            onvalue=1,
            offvalue=0,
            command=self.checkbuttons_check)

        self.customer_search.grid(row=0, column=0, sticky='n')

        self.date_search = ttkb.Checkbutton(
            self.left_bottom, 
            bootstyle="info",
            text="Date Search", 
            variable=self.date_search_var, 
            onvalue=1,
            offvalue=0,
            command=self.date_search_check)

        self.date_search.grid(row=0, column=1, sticky='n')
        
        self.item_search.grid(row=0, column=2, sticky='n')

        self.back_btn = ttkb.Button(self.bottom_frame, text="Back", bootstyle='danger', command=lambda: [self.reset_generate_frame(), parent.return_to_home()])
        self.back_btn.grid(row=0, column=0, sticky='sew', padx=10, pady=10)

        self.generate_btn = ttkb.Button(self.bottom_frame, text="Generate", bootstyle='success', command=self.pre_check)
        self.generate_btn.grid(row=0, column=3, sticky='sew', padx=10, pady=10)

        self.save_btn = ttkb.Button(self.bottom_frame, text="Save", bootstyle='info', command=self.save_file)
        #self.save_btn.grid(row=0, column=2, sticky='sew', padx=10, pady=10)

        self.customer_search_var.set(1)
        #self.customer_search.configure(state="disabled")
        self.generate_btn.configure(state="disabled")
        self.cust_var.trace('w', self.retrieve_items)
        

    def set_customer_frame(self, customer_frame):
        self.customer_frame = customer_frame

    def set_appointment_frame(self, appointment_frame):
        self.appointment_frame = appointment_frame

    def clear_error_var(self):
        self.error_var.set('')

    def save_file(self):
        
        customer_only = False
        column_list = []
        row_list = []
        rows = self.generate_table.tablerows_visible
        columns = self.generate_table.tablecolumns_visible
        for index, data in enumerate(columns):
            column_list.append(columns[index]._headertext)

        #print(column_list)
        for index1, data1 in enumerate(rows):
            row_list.append(rows[index1].values)

        for i in row_list:
            if i[3] == 0 and i[4] == 0:
                i.pop(3)
                i.pop(3)
                customer_only = True
            else:
                customer_only = False

        today = strftime("%Y-%m-%d-%I-%M-%S%p", localtime())
        #print('csv func ran')

        file_name = f"{today}_local_copy.csv"
        
        generated_reports_dir = os.path.join(resource_path("Generated_Reports"))
        file_path = os.path.join(resource_path(generated_reports_dir), file_name)
        
        if customer_only == True:
            with open(file_path, mode='w', newline='') as csvfile:
                fieldnames = column_list
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                
                writer.writeheader()
            
                for i in row_list:
                    customer_name = i[0]
                    shipment_id = i[1]
                    ship_date = i[2]
                    total_weight = i[3]
                    total_trucks = i[4]
                    writer.writerow({"Company": customer_name, "Shipment ID": shipment_id, "Date": ship_date, "Shipment Weight": total_weight, "Total Trucks": total_trucks})
                    toast = ToastNotification(title="File Saved", message=f"File: {today}_local_copy.csv has been saved in Linamar Planner/dist/Generated_Reports", position=(), duration=15000, bootstyle="info")
                    toast.show_toast()
                    self.error_label.configure(foreground="#00FF00")
                    self.error_var.set('Succesfully saved')
        else:
            with open(file_path, mode='w', newline='') as csvfile:
                fieldnames = column_list
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
            
                for i in row_list:
                    customer_name = i[0]
                    shipment_id = i[1]
                    ship_date = i[2]
                    item_name = i[3]
                    item_weight = i[4]

                    writer.writerow({"Company": customer_name, "Shipment ID": shipment_id, "Date": ship_date, "Item": item_name, "Weight": item_weight})
                    toast = ToastNotification(title="File Saved", message=f"File: {today}_local_copy.csv has been saved in /Generated_Reports", position=(), duration=15000, bootstyle="info")
                    toast.show_toast()
                    self.error_label.configure(foreground="#00FF00")
                    self.error_var.set('Succesfully saved')

    def date_search_check(self):
        self.error_var.set('')
        if self.date_search_var.get() == 1:
            self.customer_search_var.set(0)
            self.item_search_var.set(0)
            self.item_search.configure(state="disabled")
            self.customer_search.configure(state="disabled")
            self.generate_btn.configure(state="active")
            self.error_var.set('')

            self.customer_label.grid_remove()
            self.customer_combo.grid_remove()
            self.item_label.grid_remove()
            self.item_combo_box.grid_remove()
            self.from_label.grid_remove()
            self.to_label.grid_remove()
            self.from_calander.grid_remove()
            self.to_calander.grid_remove()

            #self.customer_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
            #self.customer_combo.grid(row=0, column=1, padx=10, pady=10, sticky='new')
            self.from_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
            self.to_label.grid(row=1, column=0, padx=10, pady=10, sticky='nw')
            self.from_calander.grid(row=0, column=1, padx=10, pady=10, sticky='new')
            self.to_calander.grid(row=1, column=1, padx=10, pady=10, sticky='new')


        else:
            self.error_var.set('')
            self.generate_btn.configure(state="disabled")
            self.customer_search_var.set(1)
            self.item_search.configure(state="active")
            self.customer_search.configure(state="active")

            self.customer_label.grid_remove()
            self.customer_combo.grid_remove()
            self.item_label.grid_remove()
            self.item_combo_box.grid_remove()
            self.from_label.grid_remove()
            self.to_label.grid_remove()
            self.from_calander.grid_remove()
            self.to_calander.grid_remove()

            self.customer_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
            self.customer_combo.grid(row=0, column=1, padx=10, pady=10, sticky='new')
            self.from_label.grid(row=1, column=0, padx=10, pady=10, sticky='nw')
            self.to_label.grid(row=2, column=0, padx=10, pady=10, sticky='nw')
            self.from_calander.grid(row=1, column=1, padx=10, pady=10, sticky='new')
            self.to_calander.grid(row=2, column=1, padx=10, pady=10, sticky='new')
    def checkbuttons_check(self):
        self.error_var.set('')
        if self.customer_search_var.get() == 1 and self.item_search_var.get() == 0:
            self.customer_search.configure(state="active")
            self.date_search.configure(state="active")
            self.date_search_var.set(0)
            self.error_var.set('')
            
            self.customer_label.grid_remove()
            self.customer_combo.grid_remove()
            self.item_label.grid_remove()
            self.item_combo_box.grid_remove()
            self.from_label.grid_remove()
            self.to_label.grid_remove()
            self.from_calander.grid_remove()
            self.to_calander.grid_remove()

            self.customer_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
            self.customer_combo.grid(row=0, column=1, padx=10, pady=10, sticky='new')
            self.from_label.grid(row=1, column=0, padx=10, pady=10, sticky='nw')
            self.to_label.grid(row=2, column=0, padx=10, pady=10, sticky='nw')
            self.from_calander.grid(row=1, column=1, padx=10, pady=10, sticky='new')
            self.to_calander.grid(row=2, column=1, padx=10, pady=10, sticky='new')

            #self.customer_search.configure(state="disabled")
        elif self.item_search_var.get() == 1:
            self.error_var.set('')
            self.customer_search_var.set(1)
            self.date_search_var.set(0)
            self.date_search.configure(state="disabled")
            
            
            self.generate_btn.configure(state="disabled")
            self.customer_search.configure(state="disabled")

            self.item_label.grid_remove()
            self.item_combo_box.grid_remove()
            self.customer_label.grid_remove()
            self.customer_combo.grid_remove()
            self.from_label.grid_remove()
            self.to_label.grid_remove()
            self.from_calander.grid_remove()
            self.to_calander.grid_remove()

            self.customer_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
            self.customer_combo.grid(row=0, column=1, padx=10, pady=10, sticky='new')
            self.item_label.grid(row=1, column=0, padx=10, pady=10, sticky='nw')
            self.item_combo_box.grid(row=1, column=1, padx=10, pady=10, sticky='new')
            self.from_label.grid(row=2, column=0, padx=10, pady=10, sticky='nw')
            self.to_label.grid(row=3, column=0, padx=10, pady=10, sticky='nw')
            self.from_calander.grid(row=2, column=1, padx=10, pady=10, sticky='new')
            self.to_calander.grid(row=3, column=1, padx=10, pady=10, sticky='new')

    
        elif self.item_search_var.get() == 0 and self.customer_search_var.get() == 0:
            self.error_var.set('')
            self.customer_label.grid_remove()
            self.customer_combo.grid_remove()
            self.item_label.grid_remove()
            self.item_combo_box.grid_remove()
            self.from_label.grid_remove()
            self.to_label.grid_remove()
            self.from_calander.grid_remove()
            self.to_calander.grid_remove()

            self.customer_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
            self.customer_combo.grid(row=0, column=1, padx=10, pady=10, sticky='new')
            self.from_label.grid(row=1, column=0, padx=10, pady=10, sticky='nw')
            self.to_label.grid(row=2, column=0, padx=10, pady=10, sticky='nw')
            self.from_calander.grid(row=1, column=1, padx=10, pady=10, sticky='new')
            self.to_calander.grid(row=2, column=1, padx=10, pady=10, sticky='new')

    def generate_date_only_report(self):
        self.generate_table.unhide_selected_column(cid=5)
        self.generate_table.unhide_selected_column(cid=6)
        self.generate_table.unhide_selected_column(cid=4)
        self.generate_table.unhide_selected_column(cid=3)
        self.generate_table.hide_selected_column(cid=4)
        self.generate_table.hide_selected_column(cid=3)
        self.generate_table.unhide_selected_column(cid=5)
        self.generate_table.unhide_selected_column(cid=6)
        self.generate_table.autofit_columns()
        self.generate_table.autoalign_columns()
        self.generate_table.load_table_data()

        ship_id_list = []
        temp_list = []

        from_date = self.convert_date(self.from_calander.entry.get())
        to_date = self.convert_date(self.to_calander.entry.get())
        #print('from date:', from_date)
        #print('to date:', to_date)
        conn = sqlite3.connect('work.db')
        c = conn.cursor()

        c.execute('''SELECT * FROM Shipments
                  JOIN Customers ON Shipments.customer_id = Customers.customer_id
                  WHERE Shipments.date >= ? AND Shipments.date <= ?''', (from_date, to_date))
        results = c.fetchall()
        

        if results:
            for i in results:
                if i[0] not in ship_id_list:
                    ship_id_list.append(i[1])
                    #print(i[3], i[4])
                    temp_list.append((i[6], i[0], i[4], 0, 0, i[2], i[3]))
                else:
                    continue
            #print(temp_list) 
            conn.close()

            for j in temp_list:
                self.generate_table.insert_row(tk.END, j)
                self.generate_table.load_table_data()

            self.error_var.set('')   
            self.cust_var.set('')
            self.item_var.set('')
            self.item_search_var.set(0)

            self.to_calander.entry.delete(0, tk.END)
            self.from_calander.entry.delete(0, tk.END)

            the_date = str(date.today())

            self.to_calander.entry.insert(0, self.convert_date2(the_date))
            self.from_calander.entry.insert(0, self.convert_date2(the_date))

            self.generate_btn.grid_remove()
            self.generate_btn.grid(row=0, column=2, sticky='sew')
            self.generate_btn.configure(state="disabled")
            
            self.save_btn.grid(row=0, column=3, sticky='sew', padx=10, pady=10)
            self.save_btn.configure(state="active")

            self.customer_label.grid_remove()
            self.customer_combo.grid_remove()
            self.item_label.grid_remove()
            self.item_combo_box.grid_remove()
            self.from_label.grid_remove()
            self.to_label.grid_remove()
            self.from_calander.grid_remove()
            self.to_calander.grid_remove()

            self.customer_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
            self.customer_combo.grid(row=0, column=1, padx=10, pady=10, sticky='new')
            self.from_label.grid(row=1, column=0, padx=10, pady=10, sticky='nw')
            self.to_label.grid(row=2, column=0, padx=10, pady=10, sticky='nw')
            self.from_calander.grid(row=1, column=1, padx=10, pady=10, sticky='new')
            self.to_calander.grid(row=2, column=1, padx=10, pady=10, sticky='new')

            self.customer_search.configure(state="active")
            self.customer_search_var.set(1)
            self.item_search.configure(state="active")
            self.item_search_var.set(0)
            self.date_search_var.set(0)
        else:
            self.error_var.set('ERROR: Could not find any shipments within your selected range.')


        conn.close()

    def pre_check(self):
        self.error_label.configure(foreground="#FF0000")
        self.error_var.set('')
        self.generate_table.delete_rows()
        self.generate_table.unhide_selected_column(cid=5)
        self.generate_table.unhide_selected_column(cid=4)
        self.generate_table.unhide_selected_column(cid=3)
        self.error_var.set('')

        item_selected = False
        cust_given = False
        item_given = False

        if self.date_search_var.get() == 1:
            self.generate_date_only_report()

        if self.cust_var.get():
            cust_given = True 
        else:
            cust_given = False      

        if self.item_search_var.get() == 1:
            item_selected = True
            if cust_given == True:
                if self.item_var.get():
                    item_given = True
                else:
                    item_given = False
        else:
            item_selected = False

        if cust_given == True and item_selected == True:
            if item_given == True:
                if self.item_var.get() == "- All Items -":
                    #print('if statement for "- All Items -" ran')
                    self.generate_report_all_items()
                else:
                    #print('if statement of cust_given and item_selected ran')
                    self.generate_report_item_and_cust()
            else:
                self.error_var.set('ERROR: Expected an item but receieved nothing.')
        elif cust_given == True and item_selected == False:
            #print('if statemnt of cust_only ran')
            self.generate_report_cust_only()
        elif self.date_search_var.get() == 1:
            print('date search ran')
        
        
    def validate_item(self, val):
        if val:
            self.item_exists = True
            if self.cust_exists == True:
                self.generate_btn.configure(state="active")
                return True
            else:
                self.generate_btn.configure(state="disabled")
                return False
        else:
            return False
        
    def validate_customer(self, val):
        item_search_state = self.item_search_var.get()
        #print('This is item_search_state, which is item_search_var.get()', item_search_state)
        if val:
            #print('val ran')
            self.cust_exists = True
            if item_search_state == 0:
                #print('thinks item var is 0')
                self.save_btn.configure(state="disabled")
                self.save_btn.grid_remove()
                self.generate_btn.configure(state="active")
                self.generate_btn.grid_remove()
                self.generate_btn.grid(row=0, column=3, sticky='sew', padx=10, pady=10)
                return True
            else:
                #print('thinks item var is 1')
                if self.item_exists == True:
                    #print('thinks item exists')
                    self.save_btn.configure(state="disabled")
                    self.save_btn.grid_remove()
                    
                    self.generate_btn.grid_remove()
                    self.generate_btn.configure(state="active")
                    self.generate_btn.grid(row=0, column=3, sticky='sew', padx=10, pady=10)
                    return True
                else:
                    self.generate_btn.configure(state="disabled")
                    return False
        else:
            #print('thinks nothing in cust val')
            self.save_btn.configure(state="disabled")
            self.save_btn.grid_remove()
            self.generate_btn.configure(state="disabled")
            self.generate_btn.grid_remove()
            self.generate_btn.grid(row=0, column=3, sticky='sew', padx=10, pady=10)
            return False

    def convert_date(self, date):
        date_parts = date.split('/')
        if len(date_parts) == 3:
            month = date_parts[0] 
            day = date_parts[1]
            year = date_parts[2]
        if len(year) == 2:
            year = '20' + year
        new_date = f"{year}-{month.zfill(2)}-{day.zfill(2)}"

        return new_date

    def convert_date2(self, date):
        date_parts = date.split('-')
        year = date_parts[0]
        month = date_parts[1]
        day = date_parts[2]
        #print(f"this is what is being inserted: {month}/{day}/{year}")
        return f"{month.zfill(2)}/{day.zfill(2)}/{year}"

    def retrieve_items(self, *args):
        self.error_var.set('')
        cust = self.cust_var.get().lower()
        self.item_list.clear()

        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()

        c.execute('''SELECT Customers.customer_id, Order_Items.real_item_id, Shipments.shipment_id
        FROM Customers
        JOIN Shipments ON Customers.customer_id = Shipments.customer_id
        JOIN Order_Items ON Shipments.shipment_id = Order_Items.shipment_id
        WHERE Customers.customer_name = ?
        ''', (cust,))

        results = c.fetchall()
        if results:
            if "- All Items -" not in self.item_list:
                self.item_list.append("- All Items -")
            for i in results:
                item = i[1]
                if item not in self.item_list:
                    self.item_list.append(item)
            self.item_combo_box.config(values=self.item_list)
        conn.close()

    def retrieve_customers(self):
        #self.reset_generate_frame()
        #print('retrieve customer func ran')
        conn = sqlite3.connect(resource_path("work.db"))
        c = conn.cursor()
        
        c.execute('''SELECT name FROM sqlite_master WHERE type="table" AND name="Customers"''')
        table_exists = c.fetchone()

        if table_exists:
            #print('found Customers table')
            c.execute('''SELECT * FROM Customers''')
            results = c.fetchall()
            if results:
                #print('found customers within the Customers table')
                rows = results[0]
                cust_ids = rows[0]
                
                for i in results:
                    if i[1] not in self.cust_list:
                        self.cust_list.append(i[1])
                        #print(i[1])
                self.customer_combo.config(values=self.cust_list)
            conn.close()
        else:
            conn.close()

    def reset_generate_frame(self):

        self.generate_table.delete_rows()
        self.generate_table.unhide_selected_column(cid=5)
        self.generate_table.unhide_selected_column(cid=4)
        self.generate_table.unhide_selected_column(cid=3)

        self.error_var.set('')   
        self.cust_var.set('')
        self.item_var.set('')
        self.item_search_var.set(0)
        self.customer_search.configure(state="active")
        self.item_search.configure(state="active")

        self.to_calander.entry.delete(0, tk.END)
        self.from_calander.entry.delete(0, tk.END)

        the_date = str(date.today())

        self.to_calander.entry.insert(0, self.convert_date2(the_date))
        self.from_calander.entry.insert(0, self.convert_date2(the_date))

        self.save_btn.grid_remove()
        #self.generate_btn.grid(row=0, column=2, sticky='sew')
        self.generate_btn.configure(state="disabled")
        
        self.generate_btn.grid(row=0, column=3, sticky='sew', padx=10, pady=10)
        #self.generate_btn.configure(state="active")

        self.customer_label.grid_remove()
        self.customer_combo.grid_remove()
        self.item_label.grid_remove()
        self.item_combo_box.grid_remove()
        self.from_label.grid_remove()
        self.to_label.grid_remove()
        self.from_calander.grid_remove()
        self.to_calander.grid_remove()

        self.customer_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
        self.customer_combo.grid(row=0, column=1, padx=10, pady=10, sticky='new')
        self.from_label.grid(row=1, column=0, padx=10, pady=10, sticky='nw')
        self.to_label.grid(row=2, column=0, padx=10, pady=10, sticky='nw')
        self.from_calander.grid(row=1, column=1, padx=10, pady=10, sticky='new')
        self.to_calander.grid(row=2, column=1, padx=10, pady=10, sticky='new')


    def generate_report_all_items(self):
        self.generate_table.unhide_selected_column(cid=5)
        self.generate_table.unhide_selected_column(cid=6)
        self.generate_table.unhide_selected_column(cid=4)
        self.generate_table.unhide_selected_column(cid=3)
        self.generate_table.delete_rows()
        self.generate_table.hide_selected_column(cid=5)
        self.generate_table.hide_selected_column(cid=6)
        self.generate_table.unhide_selected_column(cid=4)
        self.generate_table.unhide_selected_column(cid=3)
        self.generate_table.autofit_columns()
        self.generate_table.autoalign_columns()

        self.generate_table.load_table_data()
        self.customer_search.configure(state="active")
        self.date_search.configure(state="active")
        #print('report for BOTH item and cust ran')
        temp_list = []
        cust_name = self.cust_var.get()
        item_name = self.item_var.get()
        from_date = self.convert_date(self.from_calander.entry.get())
        to_date = self.convert_date(self.to_calander.entry.get())
        #print('from date:', from_date)
        #print('to date:', to_date)
         
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()

        c.execute('''SELECT Customers.customer_id, Shipments.shipment_id, Order_Items.shipment_id, Order_Items.real_item_id, Shipments.date, Order_Items.weight
        FROM Customers
        JOIN Shipments ON Shipments.customer_id = Customers.customer_id
        JOIN Order_Items ON Shipments.shipment_id = Order_Items.shipment_id
        WHERE Customers.customer_name = ? AND Shipments.date >= ? AND Shipments.date <= ?''', (self.cust_var.get().lower(), from_date, to_date))

        results = c.fetchall()
        if results:
            for i in results:
                #print(i[3], i[4])
                temp_list.append((cust_name, i[1], i[4], i[3], i[5], 0, 0))
            #print(temp_list) 
            conn.close()

            for j in temp_list:
                self.generate_table.insert_row(tk.END, j)
                self.generate_table.align_column_left()
                self.generate_table.load_table_data() 

            self.error_var.set('')   
            self.cust_var.set('')
            self.item_var.set('')
            self.item_search_var.set(0)

            self.to_calander.entry.delete(0, tk.END)
            self.from_calander.entry.delete(0, tk.END)
            the_date = str(date.today())
            self.to_calander.entry.insert(0, self.convert_date2(the_date))
            self.from_calander.entry.insert(0, self.convert_date2(the_date))

            self.generate_btn.grid_remove()
            self.generate_btn.grid(row=0, column=2, sticky='sew')
            self.generate_btn.configure(state="disabled")

            self.save_btn.grid(row=0, column=3, sticky='sew', padx=10, pady=10)
            self.save_btn.configure(state="active")

            self.customer_label.grid_remove()
            self.customer_combo.grid_remove()
            self.item_label.grid_remove()
            self.item_combo_box.grid_remove()
            self.from_label.grid_remove()
            self.to_label.grid_remove()
            self.from_calander.grid_remove()
            self.to_calander.grid_remove()

            self.customer_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
            self.customer_combo.grid(row=0, column=1, padx=10, pady=10, sticky='new')
            self.from_label.grid(row=1, column=0, padx=10, pady=10, sticky='nw')
            self.to_label.grid(row=2, column=0, padx=10, pady=10, sticky='nw')
            self.from_calander.grid(row=1, column=1, padx=10, pady=10, sticky='new')
            self.to_calander.grid(row=2, column=1, padx=10, pady=10, sticky='new')
        else:
            self.error_var.set('ERROR: Item does not exist for customer within your selected date range.')

    def generate_report_item_and_cust(self):
        self.generate_table.unhide_selected_column(cid=5)
        self.generate_table.unhide_selected_column(cid=6)
        self.generate_table.unhide_selected_column(cid=4)
        self.generate_table.unhide_selected_column(cid=3)
        self.generate_table.delete_rows()
        self.generate_table.hide_selected_column(cid=5)
        self.generate_table.hide_selected_column(cid=6)
        self.generate_table.unhide_selected_column(cid=4)
        self.generate_table.unhide_selected_column(cid=3)
        self.generate_table.autofit_columns()
        self.generate_table.autoalign_columns()
        self.generate_table.load_table_data()

        #print('report for BOTH item and cust ran')
        temp_list = []
        cust_name = self.cust_var.get()
        item_name = self.item_var.get()
        from_date = self.convert_date(self.from_calander.entry.get())
        to_date = self.convert_date(self.to_calander.entry.get())
        #print('from date:', from_date)
        #print('to date:', to_date)
         
        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()

        c.execute('''SELECT Customers.customer_id, Shipments.shipment_id, Order_Items.shipment_id, Order_Items.real_item_id, Shipments.date, Order_Items.weight
        FROM Customers
        JOIN Shipments ON Shipments.customer_id = Customers.customer_id
        JOIN Order_Items ON Shipments.shipment_id = Order_Items.shipment_id
        WHERE Customers.customer_name = ? AND Shipments.date >= ? AND Shipments.date <= ? AND Order_Items.real_item_id = ?''', (self.cust_var.get().lower(), from_date, to_date, item_name))
        results = c.fetchall()
        
        if results:
            for i in results:
                #print(i[3], i[4])
                temp_list.append((cust_name, i[1], i[4], i[3], i[5], 0, 0))
            #print(temp_list) 
            conn.close()
            for j in temp_list:
                self.generate_table.insert_row(tk.END, j)
                self.generate_table.align_column_left()
                self.generate_table.load_table_data() 

            self.error_var.set('')   
            self.cust_var.set('')
            self.item_var.set('')
            self.item_search_var.set(0)
            self.customer_search.configure(state="active")
            self.date_search.configure(state="active")
            self.to_calander.entry.delete(0, tk.END)
            self.from_calander.entry.delete(0, tk.END)
            the_date = str(date.today())
            self.to_calander.entry.insert(0, self.convert_date2(the_date))
            self.from_calander.entry.insert(0, self.convert_date2(the_date))

            self.generate_btn.grid_remove()
            self.generate_btn.grid(row=0, column=2, sticky='sew')
            self.generate_btn.configure(state="disabled")

            self.save_btn.grid(row=0, column=3, sticky='sew', padx=10, pady=10)
            self.save_btn.configure(state="active")

            self.customer_label.grid_remove()
            self.customer_combo.grid_remove()
            self.item_label.grid_remove()
            self.item_combo_box.grid_remove()
            self.from_label.grid_remove()
            self.to_label.grid_remove()
            self.from_calander.grid_remove()
            self.to_calander.grid_remove()

            self.customer_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
            self.customer_combo.grid(row=0, column=1, padx=10, pady=10, sticky='new')
            self.from_label.grid(row=1, column=0, padx=10, pady=10, sticky='nw')
            self.to_label.grid(row=2, column=0, padx=10, pady=10, sticky='nw')
            self.from_calander.grid(row=1, column=1, padx=10, pady=10, sticky='new')
            self.to_calander.grid(row=2, column=1, padx=10, pady=10, sticky='new')
        else:
            self.error_var.set('ERROR: Item does not exist for customer within your selected date range.')
          
    def generate_report_cust_only(self):
        self.generate_table.unhide_selected_column(cid=5)
        self.generate_table.unhide_selected_column(cid=6)
        self.generate_table.unhide_selected_column(cid=4)
        self.generate_table.unhide_selected_column(cid=3)
        self.generate_table.hide_selected_column(cid=4)
        self.generate_table.hide_selected_column(cid=3)
        self.generate_table.unhide_selected_column(cid=5)
        self.generate_table.unhide_selected_column(cid=6)
        self.generate_table.autofit_columns()
        self.generate_table.autoalign_columns()
        self.generate_table.load_table_data()

        #print('Customer ONLY report generated:')
        temp_list = []
        ship_id_list = []
        cust_name = self.cust_var.get()
        from_date = self.convert_date(self.from_calander.entry.get())
        to_date = self.convert_date(self.to_calander.entry.get())
        #print('from date:', from_date)
        #print('to date:', to_date)

        conn = sqlite3.connect(resource_path('work.db'))
        c = conn.cursor()

        c.execute('''SELECT Customers.customer_id, Shipments.shipment_id, Order_Items.shipment_id, Order_Items.real_item_id, Shipments.date, Order_Items.weight, Shipments.total_weight, Shipments.total_trucks
        FROM Customers
        JOIN Shipments ON Shipments.customer_id = Customers.customer_id
        JOIN Order_Items ON Shipments.shipment_id = Order_Items.shipment_id
        WHERE Customers.customer_name = ? AND Shipments.date >= ? AND Shipments.date <= ?''', (self.cust_var.get().lower(), from_date, to_date))
        results = c.fetchall()
        if results:
            for i in results:
                if i[1] not in ship_id_list:
                    ship_id_list.append(i[1])
                    #print(i[3], i[4])
                    temp_list.append((cust_name, i[1], i[4], 0, 0, i[6], i[7]))
                else:
                    continue
            #print(temp_list) 
            conn.close()

            for j in temp_list:
                self.generate_table.insert_row(tk.END, j)
                self.generate_table.load_table_data()

            self.error_var.set('')   
            self.cust_var.set('')
            self.item_var.set('')
            self.item_search_var.set(0)

            self.to_calander.entry.delete(0, tk.END)
            self.from_calander.entry.delete(0, tk.END)

            the_date = str(date.today())

            self.to_calander.entry.insert(0, self.convert_date2(the_date))
            self.from_calander.entry.insert(0, self.convert_date2(the_date))

            self.generate_btn.grid_remove()
            self.generate_btn.grid(row=0, column=2, sticky='sew')
            self.generate_btn.configure(state="disabled")
            
            self.save_btn.grid(row=0, column=3, sticky='sew', padx=10, pady=10)
            self.save_btn.configure(state="active")

            self.customer_label.grid_remove()
            self.customer_combo.grid_remove()
            self.item_label.grid_remove()
            self.item_combo_box.grid_remove()
            self.from_label.grid_remove()
            self.to_label.grid_remove()
            self.from_calander.grid_remove()
            self.to_calander.grid_remove()

            self.customer_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
            self.customer_combo.grid(row=0, column=1, padx=10, pady=10, sticky='new')
            self.from_label.grid(row=1, column=0, padx=10, pady=10, sticky='nw')
            self.to_label.grid(row=2, column=0, padx=10, pady=10, sticky='nw')
            self.from_calander.grid(row=1, column=1, padx=10, pady=10, sticky='new')
            self.to_calander.grid(row=2, column=1, padx=10, pady=10, sticky='new')
        else:
            self.error_var.set('ERROR: Customer does not have shipments within your selected range.')


if __name__ == "__main__":
    main()