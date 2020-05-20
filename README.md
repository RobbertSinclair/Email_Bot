# Email_Bot
A simple program that takes in an excel file and then sends out emails to all the email addresses in the excel file.

# Instructions of Use
This program only works with xlsx files. Ensure that xlrd is installed before using this software which can be installed using the command "pip install xlrd". The xlsx file must have a column which is labelled "Email". This program can only send around 100 emails at a time so if you are sending more than that number then consider having multiple xlsx files with 100 emails at a time.

Your login details should be in a secrets.py file. It should be formatted as shown in the secrets.py file in this reposity.
