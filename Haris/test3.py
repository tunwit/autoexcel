import smtplib

sender = 'haristhailand.payment@gmail.com'
receivers = ['tunwit2458@gmail.com']

message = """From: From Person <from@fromdomain.com>
To: To Person <to@todomain.com>
Subject: SMTP e-mail test

This is a test e-mail message.
"""

smtpObj = smtplib.SMTP('smtp.sendgrid.net')
smtpObj.login(user='apikey',password='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')
smtpObj.sendmail(sender, receivers, message)         
'oejoaiafvwkujlah'