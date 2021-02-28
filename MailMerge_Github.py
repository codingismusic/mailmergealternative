import xlrd
import email, ssl
import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.message import EmailMessage

loc = ("TestBook.xls")      #Store the file in the same directory as your python script
 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)
 
for i in range(sheet.nrows):    # looping through the excel sheet
    message = EmailMessage()
    sender = "xxxxxxxxx@gmail.com"
    recipient = sheet.cell_value(i, 2) # Change it to the cell where your recepients are located in the sheet
    message['From'] = sender
    message['To'] = recipient
    message['Subject'] = f'Your email subject {sheet.cell_value(i, 0)}' # use the F string for dynamic content in the subject field. Change the cell number accordingly
    body = f'''Your email body "{sheet.cell_value(i, 0)}"? ''' # use the F string for dynamic content. Change the cell number accordingly

    message.set_content(body)
    mime_type, _ = mimetypes.guess_type('LOA.txt') # File location, i would recommend keep the file in the same directory as python script
    mime_type, mime_subtype = mime_type.split('/')
    with open('LOA.txt', 'rb') as file:
     message.add_attachment(file.read(),
     maintype=mime_type,
     subtype=mime_subtype,
     filename='LOA.txt')
    print(message)

    
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.set_debuglevel(1)
    mail_server.login("xxxxxxxxx@gmail.com", "Put your password here")
    mail_server.send_message(message)
    mail_server.quit()
