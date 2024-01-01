import smtplib
import imaplib
import ssl
import time
from email.mime.text import MIMEText
from email.message import EmailMessage
import text
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd

Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
zalacznik = askopenfilename(
    title='Wybierz plik zalacznika')  # show an "Open" dialog box and return the path to the selected file
print(zalacznik)
sender_email = 'zdzislawolczyk@poczta.onet.pl'
nadajnik = smtplib.SMTP_SSL('smtp.poczta.onet.pl', 465)  # 587 #465


file = 'C:\Roboczy\zeszyt1.xlsx'
df_excel = pd.read_excel(io=file, sheet_name=0, header=0, names=['email', 'imie', 'nazwisko', 'firma'])

haslo = '@mesi1000Olczyk'
CC = 'test@gmail.com'



# ------------------------------------------------------------------------------------------------------
def send_email(tytul, body, sender, adresat, CC, zalacznik):
    subject = tytul

    msg = EmailMessage()
    body_mail = MIMEText(body, 'html')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = adresat
    msg['CC'] = CC
    msg.set_content(body, subtype='html')
    with open(zalacznik, "rb") as file:
        msg.add_attachment(
            file.read(),
            filename=zalacznik.split('/')[-1],
            maintype="application",
            subtype=zalacznik.split('\\')[-1].rsplit('.', maxsplit=-1)[-1]
        )
    # Start the TLS connection
    # server.starttls()
    nadajnik = smtplib.SMTP_SSL('smtp.poczta.onet.pl', 465)
    nadajnik.login(sender_email, "@mesi1000Olczyk")
    nadajnik.ehlo_or_helo_if_needed()
    nadajnik.sendmail(sender, adresat, msg.as_string())
    print(adresat)
    nadajnik.quit
    odbiornik = imaplib.IMAP4_SSL('imap.poczta.onet.pl', 993)
    odbiornik.login(sender_email,haslo)
    #odbiornik.select('SENT')
    odbiornik.append('SENT', '', imaplib.Time2Internaldate(time.time()), str(msg).encode('utf-8'))
    odbiornik.logout()

for x in df_excel.values:
    mail = x[0]
    imie = x[1]
    nazwisko = x[2]
    firma = x[3]
    tytul = f'mail TEST do {imie} {nazwisko}'
    body = text.body(imie, nazwisko, firma)
    send_email(tytul=tytul, body=body, sender=sender_email, adresat=mail, CC=CC, zalacznik=zalacznik)

nadajnik.close()

