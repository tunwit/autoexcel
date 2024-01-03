from email.message import EmailMessage
import ssl
import smtplib

sender = 'haristhailand.payment@gmail.com'
password = 'oejoaiafvwkujlah'
reciver = ['tunwit2458@gmail.com','kinmour2020@gmail.com']

subject = 'testsystem3'

body = """
Body tester for email system

"""


context = ssl.create_default_context()

with smtplib.SMTP_SSL('smtp.gmail.com',465,context = context) as smtp:
    smtp.login(sender,password)
    for m in reciver:
        em = EmailMessage()
        em['From'] = sender
        em['To'] = m
        em['subject'] = f'message to {m}'
        em.set_content(f"Here is salary for {m}")
        smtp.sendmail(sender,m,em.as_string())
        print(f'Mail has senfd to {m}')