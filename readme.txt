Release V2.0    2020-03-23

1. Finalize manage.py and excel.py
2. TODO: start DEV branch on spam.py

===============================================================================
Release V1.2    2019-02-15

1. add an addon field for extra charges, parking, keys, etc. (May not need)
2. use log instead of printf, ignore 0 payment tenants
3. add backup system to the program, now backup 5 runs
4. fix error in counting last day 

===============================================================================
Release V1.1

1. use a config file to input account name and password

2. better format of backup file, column width, sum, etc.

3. Auto create e-mail txt directory

4. TODO: change error check so 0 service date is OK (no need to remove old tenants)


===============================================================================
Release V1.0

1. use run-time search to key words so colomn CAN be added. for example, insert a new column between tenants and e-mail would not cause any issue

2. a single xlsx file CAN take mulitply service dates, that is, one file for all tenants

3. add a 'skip e-mail' column, with yes value it means calculate fees for that tenant, but do not send e-mail

4. the SMTP code in send_gamil() is disabled, now it is a fake send, do all the check, but do not send, so actual e-mail address would not cause an issue

5. library change in saving backups, no need to install pypiwin32, thus mac air should work. well, theoretically.

6. backup method is modified to create new xlsx file (if not found) with the string 'backup' appended to input filename

7. add a feature to cleanup input xlsx file, so service dates are cleared, force new run to enter new service dates

===============================================================================
