import pandas as pd
import tkinter as tk
from tkinter import filedialog
import tkinter.font as font
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from secrets import home_email, password
from email import encoders
import os
#Get the email data



class email_client(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.email_list = []
        self.fail_list = []
        
        self.attach_file = None
        self.create_widgets()

    def create_widgets(self):
        self.import_Frame = tk.Frame(self)
        self.import_Frame.pack()
        self.import_label = tk.Label(self.import_Frame, text="Select File", width=25)
        self.import_label.pack(side="left")
        self.import_button = tk.Button(self.import_Frame, text="Import Email List", width=25, command=self.import_file)
        self.import_button.pack(side="right")
        self.attach_Frame = tk.Frame(self)
        self.attach_Frame.pack()
        self.attach_label = tk.Label(self.attach_Frame, text="Attach File", width=25)
        self.attach_label.pack(side="left")
        self.attach_button = tk.Button(self.attach_Frame, text="Attach File", width=25, command=self.select_attach)
        self.attach_button.pack(side="right")
        self.subject_Frame = tk.Frame(self)
        self.subject_Frame.pack()
        self.subject_label = tk.Label(self.subject_Frame, text="Subject", width=25)
        self.subject_label.pack(side="left")
        self.subject_text = tk.Entry(self.subject_Frame, width=30)
        self.subject_text.pack(side="right")
        self.body_Frame = tk.Frame(self)
        self.body_Frame.pack()
        self.body_label = tk.Label(self.body_Frame, text="Body", width=25)
        self.body_label.pack(side="left")
        self.body_text = tk.Text(self.body_Frame, height=6, width=23)
        self.body_text.pack(side="right")
        self.button_Frame = tk.Frame(self)
        self.button_Frame.pack()
        self.send_button = tk.Button(self.button_Frame, text="Send Emails", width=25, command=self.send_emails)
        self.send_button.pack(side="left")
        self.quit_button = tk.Button(self.button_Frame, text="Quit", width=25, command=self.master.destroy)
        self.quit_button.pack(side="right")

    def import_file(self):
        self.filename = filedialog.askopenfilename(initialdir="/Documents", title="Select A file", filetypes= (("xlsx files", "*.xlsx"), ("All Files", "*.*") ))
        try:
            data = pd.read_excel(self.filename)
            data = pd.DataFrame(data)
            try:
                self.email_list = data['Email'].tolist()
                self.filename = os.path.split(self.filename)[1]
                self.import_label['text'] = self.filename
                print(self.email_list)
            except:
                print("There is no Email option in this file")
        except:
            print("I cannot find the file")

    def select_attach(self):
        self.attach_file = filedialog.askopenfilename( initialdir="/Documents", title="Select Attachment", filetypes=(("PDF files", "*.pdf"), ("PNG files", "*.png"), ("JPEG Files", "*.jpg"), ("All Files", "*.*")))
        self.attach_filename = os.path.split(self.attach_file)
        self.attach_filename = self.attach_filename[1]
        self.attach_label['text'] = self.attach_filename 

    def send_emails(self):
        if len(self.email_list) == 0:
            print("There are no emails avaliable")
        else:
            subject = self.subject_text.get()
            body = self.body_text.get("1.0", "end")
            msg = MIMEMultipart()
            msg["Subject"] = subject
            msg.attach(MIMEText(body, "plain"))
            try:
                attachment = open(self.attach_file, "rb")
                part = MIMEBase("application", "octet-stream")
                part.set_payload((attachment).read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", "attachment; filename= "+self.attach_filename)
                msg.attach(part)
            except:
                print("No attachment")
            text = msg.as_string()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(home_email, password)
                print(text)
                for email in self.email_list:
                    try:
                        smtp.sendmail(home_email, email, text)
                        print("Email sent to {0}".format(email))
                    except:
                        print("Failed to send an email to {0}".format(email))
                        self.fail_list.append(email)
                if len(self.fail_list) != 0:
                    print("The bot couldn't load on these emails")
                    for email in self.fail_list:
                        print(email)
            
        
    
        
root = tk.Tk()
the_client = email_client(master=root)




