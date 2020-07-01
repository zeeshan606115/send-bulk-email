import pandas as p
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

data = p.read_excel('mail.xlsx')
email_col = data.get("email")
list_of_email = list(email_col)
print(list_of_email)

try:
    server = sm.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login("Email_id","your_password")
    from_ ="Email_id"
    to_ = list_of_email
    message = MIMEMultipart('alternative')
    message['Subject'] = "This is testing message"
    message['from'] = 'Email_id'
    html='''
        <html>
        <head></head>
        <body>
        <h1>Message from zeeshan</h1>
        <h2> Just testing message</h2>
        <button style = "padding:20px;background:green; color:white;">
        Verify</button>
        </body>
        </html>
    
    '''
    text = MIMEText(html, 'html')
    message.attach(text)
    server.sendmail(from_, to_, message.as_string())
    print('Message has been send to email ')
except Exception as e:
    print (e)