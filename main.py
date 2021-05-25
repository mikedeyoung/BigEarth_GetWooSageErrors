# GetWooSageErrors
# main.py v1.0
# mike@phonecompany.cloud

import datetime
import os.path
from shutil import copyfile
from time import sleep
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

html_report = """
<html>
    <head>
        <title>Webstore ErrorLog Report</title>
        <h1>Webstore ErrorLog Report</h1>
    </head>
    <body>
        <large>There was an error importing from the Webstore into Sage. 
        </br>
        </br>
        See attached.</large>
        </br>
        </br>
        <small>This is an automated message.</small>
    </body>
</html>
"""

html_offline = """ """

def main():

    # Check for the existence of ErrorLog.txt
    sleep(2)
    if os.path.isfile('C:/Sage/Sage 300 ERP/Custom/macros/WooCommerceSync/Logs/Error/ErrorLog.txt'):
        print ('ErrorLog Exists. Processing Started ' + str(datetime.datetime.today()) + '.')

        global FileName1
        global FileName2
        FileName1 = "C:\Sage\Sage 300 ERP\Custom\macros\WooCommerceSync\Logs\Error\ErrorLog.txt"
        FileName2 = "C:\Sage\Sage 300 ERP\Custom\macros\WooCommerceSync\Logs\Error\Log_Backup\ErrorLog.txt"

        # Connect to SMTP server
        session = smtplib.SMTP(host='smtp.office365.com', port=587)
        print('Connected to SMTP server...')
        sleep(2)
        session.starttls()
        print('Started TLS...')
        sleep(2)
        session.login('alerts@bigearthsupply.com', 'password')
        print('Logged in...')
        sleep(2)

        # Define SMTP Session
        msg = MIMEMultipart('alternative')
        email_body = html_report
        msg['Subject'] = 'Webstore Import Error - ' + str(datetime.datetime.today())[:10]
        msg['From'] = 'alerts@bigearthsupply.com'
        msg['To'] = 'email1@nodomain.com'
        msg['CC'] = 'email2@nodomain.com'
        # Insert Message Body
        msg.attach(MIMEText(email_body, 'html'))

        # Open the file ErrorLog.txt
        filename = "ErrorLog.txt"
        attachment = open("C:/Sage/Sage 300 ERP/Custom/macros/WooCommerceSync/Logs/Error/ErrorLog.txt", "rb")
        # instance of MIMEBase and named as p
        p = MIMEBase('application', 'octet-stream')
        # To change the payload into encoded form
        p.set_payload((attachment).read())
        # encode into base64
        encoders.encode_base64(p)
        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        # attach the instance 'p' to instance 'msg'
        msg.attach(p)
        attachment.close()

        # Send the message
        session.send_message(msg)
        session.quit()
        print('Email sent...')
        sleep(2)
        print('Completed')

        # Backup ErrorLog.txt and delete
        os.remove(FileName2)
        copyfile(FileName1, FileName2)
        os.remove(FileName1)

    else:
        print('ErrorLog.txt does not exist. Quitting...')
        sleep(2)

main()