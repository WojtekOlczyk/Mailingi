import win32com.client as win32
import os


# instancja outlooka
outlook = win32.Dispatch('outlook.application')
#outlook.GetNameSpace('MAPI')


# obiekt emaila
imie='Wojtek'
zap = outlook.CreateItem(1)
zap.subject = 'Zaproszenie na Teams'
zap.body = f'Zapraszam na spotkanie {imie} '
zap.location='Teams application'
zap.start='18/12/2023 10:00:00 AM'
zap.duration=60
#zap.Attachments.Add('')
zap.importance=2
zap.MeetingStatus=1
required=zap.Recipients.add('woyciecholczyk@gmail.com')
required.Type=1
optional=zap.Recipients.add('zdzislawolczyk@poczta.onet.pl')
optional.Type=2
zap.display()

outlook.GetNameSpace('MAPI')
zap.Send()




