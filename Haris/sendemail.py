from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import ssl
import smtplib
import json
import os
import time 
from email.utils import formatdate
from email.mime.application import MIMEApplication
import pypdfium2 as pdfium

check = str(input("Do you sure to activate send email program (y/n) : "))

if check.lower() == "y":
    mouth = {
        "January":"มกราคม", 
        "February":"กุมภาพันธ์", 
        "March":"มีนาคม", 
        "April":"เมษายน", 
        "May":"พฤษภาคม", 
        "June":"มิถุนายน", 
        "July":"กรกฎาคม", 
        "August":"สิงหาคม", 
        "September":"กันยายน",
        "October":"ตุลาคม", 
        "November":"พฤศจิกายน", 
        "December":"ธันวาคม"
    }

    with open("config.json","r",encoding="utf8") as config:
        data = json.load(config)
        
    sender = data["sender_email"]
    password = data["email_password"]
    context = ssl.create_default_context()

    # text = """
    # เรียน {name},\n\n
    # \tใบสลิปเงินเดือนของ {name} ประจำเดือน {DY} ของบริษัท ไอแอมฟู้ด จำกัด\n
    # หากมีข้อผิดพลาดประการใดขออภัยไว้ ณ ที่นี้\n\n
    # \t\t\t\t\tจึงเรียนมาเพื่อทราบ\n
    # \t\t\t\t\tบริษัท ไอแอมฟู้ด จำกัด
    # """
    with smtplib.SMTP_SSL('smtp.gmail.com',465,context = context) as smtp:
        smtp.login(sender,password)
        directory = [f for f in os.listdir() if '.' not in f and f != "__pycache__" and f != "languages"]
        p = []
        for d in directory:
            for person in os.listdir(d):
             p.append(person)

        all_person = len(p)
        done = 0
        for d in directory:
            for person in os.listdir(d):
                if person.split(',')[1][0:-4] == "-":
                    continue
                msg = MIMEMultipart()
                msg['From'] = "Haris premium buffet"
                msg['To'] = person.split(',')[1][0:-4]
                msg['subject'] = f'สลิปเงินเดือนของ {person.split(",")[0]}'
                msg['Date'] = formatdate(localtime=True)

                # msg.attach(MIMEText(text.format(name=person.split(",")[0],DY=f"{mouth[time.strftime('%B')]} {int(time.strftime('%Y'))+543}")))

                msg.attach(MIMEText('<img src="cid:image1" width="1000" height="772">', 'html'))
                
                pdf = pdfium.PdfDocument(f"{d}/{person}")
                page = pdf.get_page(0)
                pil_image = page.render_topil(scale = 300/72)
                pil_image.save(f"{person}.png")
                with open(f"{person}.png", "rb") as f:
                    data = f.read()
                img = MIMEImage(data,_subtype="png")
                img.add_header('Content-ID', '<image1>')
                msg.attach(img)

                with open(f"{d}/{person}", "rb") as f:
                    attach = MIMEApplication(f.read(),_subtype="pdf")
                attach.add_header('Content-Disposition','attachment',filename=f"{mouth[time.strftime('%B')]} {int(time.strftime('%Y'))+543}.pdf")
                msg.attach(attach)
                os.remove(f"{person}.png")
                smtp.sendmail(sender,person.split(',')[1][0:-4],msg.as_string())
                done += 1
                print(f'Mail has send to {person.split(",")[0]} | {person.split(",")[1][0:-4]} {done}/{all_person}')

input("press ENTER to close program :")