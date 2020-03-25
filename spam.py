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
from manage import send_gmail
import tkinter as tk

SPAM_TITLE = "Reminder: all common area will be cleaned "
SPAM_BODY = """Dear tenant:
Please note all will be cleaned, blah, 
blah, blah, 
Thank you, 

Best regards,
Paul
"""


class SpamGroup:
    def __init__(self, group_name: str, name: str, email: str) -> None:
        self.group_name = group_name
        self.is_select = False
        self.name = [name, ]
        self.email = [email, ]

    def add_target(self, name: str, email: str) -> None:
        self.name.append(name)
        self.email.append(email)

    def select_group(self) -> None:
        self.is_select = True

    def deselect_group(self) -> None:
        self.is_select = False

    def get_name_string(self, div='\n') -> str:
        string = ''
        for item in self.name:
            string += item + div
        return string

    def get_email_string(self, div='\n') -> str:
        string = ''
        for item in self.email:
            string += item + div
        return string

    # only send spam if is selected
    def send_spam(self, myaddress, mypassword, title, body) -> bool:
        if not self.is_select:
            return False
        for item in self.email:
            if send_gmail(myaddress, mypassword, item, title, body):
                print(item + " sent successfully\n")
            else:
                return False
        return True

    @ staticmethod
    # only first 3 letter for now
    def get_group_name(text: str) -> str:
        return text[0:3]


class SpamCenter:
    def __init__(self, tenant: List[excel.Tenant]):
        self.group = {}  # type: dict
        for item in tenant:
            if item.sendemail:
                group_name = SpamGroup.get_group_name(item.room)
                if group_name in self.group:
                    self.group[group_name].add_target(item.name, item.email)
                else:
                    self.group[group_name] = SpamGroup(group_name, item.name, item.email)


class MainFrame:
    def __init__(self, service_type, myaddress, mypassword):
        self.root = tk.Tk()
        self.root.title(service_type)
        self.myaddress = myaddress
        self.mypassword = mypassword
        left = tk.Frame(self.root)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=20, pady=20)
        right = tk.Frame(self.root)
        right.pack(side=tk.RIGHT, fill=tk.Y, padx=20, pady=20)
        xlsx = excel.Excel('OasisTenates.xlsx', 'next')
        xlsx.process()
        self.spam = SpamCenter(xlsx.tenant)

        # add checkbox and name/email list to the left frame
        self.var_checkbox = {}      # type dict
        for key in self.spam.group.keys():
            # add checkbox
            temp_int_var = tk.IntVar()
            self.var_checkbox[key] = temp_int_var
            temp_checkbox = tk.Checkbutton(left, text=key, variable=temp_int_var, onvalue=1, offvalue=0, command=self.update_status)
            temp_checkbox.pack(side=tk.TOP)
            # add label with name and email
            temp_frame = tk.Frame(left)
            temp_frame.pack(side=tk.TOP)
            temp_lb_name = tk.Label(temp_frame, text=self.spam.group[key].get_name_string())
            temp_lb_name.pack(side=tk.LEFT)
            temp_lb_email = tk.Label(temp_frame, text=self.spam.group[key].get_email_string())
            temp_lb_email.pack(side=tk.RIGHT)

        # add email title and body to the right frame
        self.var_title = tk.StringVar()
        self.var_title.set(SPAM_TITLE)
        tk.Label(right, text="Email title:").pack(anchor=tk.NW)
        tk.Entry(right, width=50, textvariable=self.var_title).pack(side=tk.TOP)
        tk.Label(right, text="Email body:").pack(anchor=tk.NW)
        self.text_widget = tk.Text(right, width=50, height=20)
        self.text_widget.insert(tk.END, SPAM_BODY)
        self.text_widget.pack(side=tk.TOP)
        # add status frame to the bottom of right frame
        status = tk.Frame(right)
        status.pack(side=tk.BOTTOM, fill=tk.X, pady=20)
        # add list of selected emails to the left of status frame
        # tk.Label(status, text="Spam list:").pack(anchor=tk.NW)
        self.lb_status = tk.Label(status, text="No group selected")
        self.lb_status.pack(side=tk.LEFT)
        # add confirm button to the right of status frame
        self.bt_confirm = tk.Button(status)
        self.bt_confirm.config(text="SEND SPAM", fg="red", width=20, height=5, command=self.send_spam)
        self.bt_confirm.pack(anchor=tk.NE)

    def update_status(self):
        for key in self.var_checkbox:
            if self.var_checkbox[key].get() == 1:
                self.spam.group[key].select_group()
            else:
                self.spam.group[key].deselect_group()
        string = ''
        for key in self.spam.group:
            if self.spam.group[key].is_select:
                string += self.spam.group[key].get_email_string()
        string = 'no group selected' if string == '' else string
        self.lb_status.config(text=string)

    def send_spam(self):
        # checkbox should be updated already, so is self.spam
        # load title and body text
        title = self.var_title.get()
        body = self.text_widget.get("1.0", "end-1c")
        string = 'Spam to '
        for key in self.var_checkbox:
            string += key + ' ' + str(self.var_checkbox[key].get()) + ' '
        string += '\n'
        string += 'my: ' + self.myaddress + ' ' + self.mypassword + '\n'
        string += 'title: ' + title + '\n'
        string += 'body' + body
        print(string)


# start top level
if __name__ == '__main__':
    with open('config.txt') as fin:
        line1 = fin.readline().strip()
        line2 = fin.readline().strip()
    # check on line1 and line2 to make sure cross platform line definition is correct
    if line1.endswith('@gmail.com') and line2.startswith('Hsi'):
        frame = MainFrame('cleaning', line1, line2)
        frame.root.mainloop()
        print("End of script\n")
    else:
        print("Check config.txt\n")
# end of file
