import win32com.client as win32
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import csv
import pandas as pd
import openpyxl
import text_html
from termcolor import colored

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
zalacznik = askopenfilename(title='Wybierz plik zalacznika') # show an "Open" dialog box and return the path to the selected file

# instancja outlooka
outlook = win32.Dispatch('outlook.application')
outlook.GetNameSpace('MAPI')

#file = askopenfilename(title='Wybierz plik z adresatami')
file='C:\Roboczy\zeszyt1.xlsx'
df_excel=pd.read_excel(io=file,sheet_name=0,header=0,names=['email','imie','nazwisko','firma'])
#print(df_excel)

sender='zdzislawolczyk@poczta.onet.pl'

# obiekt emaila
def send_mail(tytul, body,sender, adresat,zalacznik):
    mail = outlook.CreateItem(0)
    mail.subject = tytul
    mail.BodyFormat=3
    mail.body = body
    mail.sender= sender
    mail.to=adresat
    mail.Attachments.Add(zalacznik)
    mail.importance=2
    #mail.DeferredDeliveryTime='23/12/2023 20:56:00 PM'
    mail.Send()


for x in df_excel.values:
    mail=x[0]
    imie=x[1]
    nazwisko=x[2]
    firma=x[3]
    tytul=f'mail testowy do {imie}'
    body=text_html.body(imie,nazwisko,firma)
    send_mail(tytul=tytul,body=body,sender=sender,adresat=mail,zalacznik=zalacznik)

'''
for x in df_excel.values:
    mail=x[0]
    imie=x[1]
    nazwisko=x[2]
    firma=x[3]
    print(mail,imie,nazwisko,firma)
'''
