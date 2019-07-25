"""
top level to use excel module

python3 manage.py -i
rbiz07.645sj@gmail.com

vervion 1.0 released on 2018.01.01
version 1.2 released on 2019.02.15
version 1.3 released on 2019.07.24
"""

import logging
import sys
import time
import openpyxl
import smtplib
import excel
import datetime

# gloable error code here
ERROR_WRONG_CMD = 101
ERROR_WRONG_XLSX = 102
ERROR_WRONG_SMTP = 103

def send_gmail(myaddress, mypassword, toaddress, subject, bodytxt, flag=True):
    """ myaddress should be gmail, send to toaddress with subject and bodytxt """
    rest = 'From: %s\n' % myaddress
    rest += 'To: %s\n' % toaddress
    rest += 'Subject: %s\n\n' % subject
    rest += bodytxt
    # connect to gmail server
    if flag:
        try:
            server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            server.ehlo()
            server.login(myaddress, mypassword)
            server.sendmail(myaddress, toaddress, rest)
            server.quit()
        except smtplib.SMTPException as error:
            excel.Excel.error_exit('SMTP e-mail got exception' + str(error.__class__), ERROR_WRONG_SMTP)
        time.sleep(3)  # wait 3 seconds to refresh server

def send_gmail_all(myaddress, mypassword, tenant_list, flag=True):
    """ send to all tenant in a tenant list """
    for tenant in tenant_list:
        if tenant.sendemail:
            send_gmail(myaddress, mypassword, tenant.email,
                       'Utility bill due %s' % tenant.service_cycle.get_billday_string(), tenant.get_email_txt(), flag)
            print('E-maind sent successfully ==> %s [%s %s]' % (tenant.email, tenant.room, tenant.name))

def send_gmail_m_all(myaddress, mypassword, tenant_list, target_date, flag=True):
    """ send maintaince email to all tenant in a tenant list """
    for tenant in tenant_list:
        if tenant.sendemail:
            # TODO: format email + 2day for room not start with '645' & '632', logic here is not correct
            monday_list = ['645', '632']
            for _ in monday_list:
                if _ == tenant.room[:3]:
                    target_date += datetime.timedelta(days=2)
                    break

            send_gmail(myaddress, mypassword, tenant.email,
                       'Utility bill due %s' % tenant.service_cycle.get_billday_string(), tenant.get_email_txt(), flag)
            print('E-maind sent successfully ==> %s [%s %s]' % (tenant.email, tenant.room, tenant.name))

def ask_confirm(action):
    strin = ' '
    while strin != 'yes':
        strin = input('\nEnter Yes to %s, No to quit ==> ' % action).lower().lstrip().rstrip()
        if strin == 'no':
            return False
    return True

def test_open(filename, sheetname):
    """" pass in xlsx file must contain [room, tenants, e-mail, service dates] in first row """
    try:
        wb = openpyxl.load_workbook(filename, read_only=True)
        ws = wb[sheetname]
    except KeyError:
        return False

    # test if first row contains 4 key words and in first 26 colomn
    keydict = {}
    index = 0
    for cell in ws[1]:
        keydict[str(cell.value).lower().lstrip().rstrip()] = 1
        if index == 25:
            break
        index += 1
    wb.close()
    return {'room', 'tenants', 'e-mail', 'service dates'} < keydict.keys()


# start top level
if __name__ == '__main__':
    def main():
        print(' ')
        filename = 'nosuchfile'
        if len(sys.argv) < 2 or (sys.argv[1].lower() != '-i' and sys.argv[1].lower() != '-m'):
            excel.Excel.error_exit('Usage: python manage.py -i [file.xlsx] [sheetname]', ERROR_WRONG_CMD)
        elif len(sys.argv) == 2:
            filename = 'OasisTenates.xlsx'
        else:
            filename = sys.argv[2] if sys.argv[2].endswith('.xlsx') else sys.argv[2] + '.xlsx'
        if sys.argv[1] == '-I':
            logging.disable(logging.INFO)
        sheetname = 'next' if len(sys.argv) <= 3 else sys.argv[3]
        # test open file and find sheet using openpyxl
        if not test_open(filename, sheetname):
            excel.Excel.error_exit('%s [%s] has wrong format or not accessible' % (filename, sheetname), ERROR_WRONG_XLSX)

        if sys.argv[1].lower() == '-i':
            xlsx = excel.Excel(filename, sheetname)
            xlsx.process()
            print('\nCheck all tenants')
            xlsx.tenant_check()
            xlsx.write_all_tenant_to_file()
            if ask_confirm('send e-mail'):
                with open('config.txt') as fin:
                    email = fin.readline().strip()
                    password = fin.readline().strip()
                send_gmail_all(email, password, xlsx.tenant, True)
            if ask_confirm('backup xlsx'):
                xlsx.backup(filename[:-5]+' backup.xlsx')
            if ask_confirm('clean input xlsx'):
                xlsx.cleanup()
            xlsx.close()
        elif sys.argv[1].lower() == '-m':
            xlsx = excel.Excel(filename, sheetname)
            xlsx.process()
            target_date = datetime.datetime(sys.argv[4], sys.argv[2], sys.argv[3])
            print('Target date is %s' % str(target_date)[:10])
            if ask_confirm('send e-mail'):
                with open('config.txt') as fin:
                    email = fin.readline().strip()
                    password = fin.readline().strip()
                send_gmail_m_all(email, password, xlsx.tenant, target_date, True)
            xlsx.close()

    logging.basicConfig(level=logging.WARNING, format='%(levelname)s - %(message)s')
    main()

# end of top level
