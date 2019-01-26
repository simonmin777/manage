===============================================================================
Changes:

1. use run-time search to key words so colomn CAN be added. for example, insert a new column between tenants and e-mail would not cause any issue

2. a single xlsx file CAN take mulitply service dates, that is, one file for all tenants

3. add a 'skip e-mail' column, with yes value it means calculate fees for that tenant, but do not send e-mail

4. the SMTP code in send_gamil() is disabled, now it is a fake send, do all the check, but do not send, so actual e-mail address would not cause an issue

5. library change in saving backups, no need to install pypiwin32, thus mac air should work. well, theoretically.

6. backup method is modified to create new xlsx file (if not found) with the string 'backup' appended to input filename

7. add a feature to cleanup input xlsx file, so service dates are cleared, force new run to enter new service dates

===============================================================================

1. install python-3.7.1
	==> go to https://www.python.org/downloads/windows/
	==> scroll down a little bit to find
		Python 3.7.1 - 2018-10-20
			Download Windows x86 web-based installer
			Download Windows x86 executable installer
			Download Windows x86 embeddable zip file
			Download Windows x86-64 web-based installer
			...
	==> click and download the 2nd one: x86 executable installer
	==> download file name should be python-3.7.1.exe
	==> double click to run it

2. make sure python (py) in system PATH
	==> open cmd, type: py -h
	==> if see help message of python, it works

3. still in cmd, install python libray, type each of the following and hit enter
	==> pip install datetime
	==> pip install openpyxl

4. copy manage.py excel.py SJ645.xlsx to a clean directory
	IMPORTANT: make a new folder under that directory as "housebill", all email txt will be stored in that folder

5. open cmd, cd to the directory contain manage.py, type
	==> py manage.py -i SJ645

6. if need to test send e-mail, follow these steps:
	==> turn on less secure app access, google instruction at https://support.google.com/accounts/answer/6010255
	==> also make sure 2-step verification is turned off for this version
	==> open manage.py with a txt editor, notepad ok, but not word, find this line at very bottom
            send_gmail_all('youremail@gmail.com', 'password', xlsx.tenant, False)
    ==> change the gmail account and passowrd, and change last argument to True
	==> the enclosing ' ' must be there, and no space inside ' ', for example, simonmin777@gmail with password 12345 will be
		==> 	send_gmail_all('simonmin777@gmail.com', '12345', xlsx.tenant, True)


