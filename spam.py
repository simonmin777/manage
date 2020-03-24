"""
GUI send email to selective tenants
Construct tenant list using excel.py
"""

# from datetime import datetime, timedelta
from typing import List
# import logging
import sys
import time
import smtplib
import excel
import tkinter as tk

SPAM_TITLE = "Reminder: all common area will be cleaned "
SPAM_BODY = """Dear tenant:
Please note all will be cleaned, blah, 
blah, blah, 
Thank you, 

Best regards,
Paul
"""


# only first 3 letter for now
def get_group_name(text: str) -> str:
    return text[0:3]


class SpamGroup:
    def __init__(self, group_name, name, email):
        self.group_name = group_name
        self.name = [name, ]
        self.email = [email, ]

    def add_target(self, name, email):
        self.name.append(name)
        self.email.append(email)


class SpamCenter:
    def __init__(self, tenant: List[excel.Tenant]):
        self.group = {}  # type: dict
        for item in tenant:
            if item.sendemail:
                group_name = get_group_name(item.room)
                if group_name in self.group:
                    self.group[group_name].add_target(item.name, item.email)
                else:
                    self.group[group_name] = SpamGroup(group_name, item.name, item.email)


class MainFrame:
    def __init__(self, service_type):
        self.root = tk.Tk()
        self.root.title(service_type)
        left = tk.Frame(self.root)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=20, pady=20)
        right = tk.Frame(self.root)
        right.pack(side=tk.RIGHT, fill=tk.Y, padx=20, pady=20)
        xlsx = excel.Excel('OasisTenates.xlsx', 'next')
        xlsx.process()
        self.spam = SpamCenter(xlsx.tenant)

        # add checkbox and name/email list to the left frame
        self.var_checkbox = []      # type List[dict]
        for key in self.spam.group.keys():
            # add checkbox
            temp_int_var = tk.IntVar()
            self.var_checkbox.append({key: temp_int_var})
            temp_checkbox = tk.Checkbutton(left, text=key, variable=temp_int_var, onvalue=1, offvalue=0)
            temp_checkbox.pack(side=tk.TOP)
            # add label with name and email
            temp_frame = tk.Frame(left)
            temp_frame.pack(side=tk.TOP)
            temp_name = ''
            for item in self.spam.group[key].name:
                temp_name += item + '\n'
            temp_email = ''
            for item in self.spam.group[key].email:
                temp_email += item + '\n'
            temp_lb_name = tk.Label(temp_frame, text=temp_name)
            temp_lb_name.pack(side=tk.LEFT)
            temp_lb_email = tk.Label(temp_frame, text=temp_email)
            temp_lb_email.pack(side=tk.RIGHT)

        # add email title and body to the right frame
        self.var_title = tk.StringVar()
        self.var_title.set(SPAM_TITLE)
        tk.Label(right, text="Email title:").pack(anchor=tk.NW)
        tk.Entry(right, width=50, textvariable=self.var_title).pack(side=tk.TOP)
        tk.Label(right, text="Email body:").pack(anchor=tk.NW)
        self.text_widget = tk.Text(right, width=50, height=30)
        self.text_widget.insert(tk.END, SPAM_BODY)
        self.text_widget.pack(side=tk.TOP)
        # add confirm button to the right frame
        self.bt_confirm = tk.Button(right)
        self.bt_confirm.config(text="SEND SPAM", command=self.send_spam)
        self.bt_confirm.pack(anchor=tk.SE, pady=10)

    def send_spam(self):
        print("Yep, sending spams now " + self.var_title.get())


# start top level
if __name__ == '__main__':
    # unittest.main()
    frame = MainFrame('cleaning')
    frame.root.mainloop()
# end of file
