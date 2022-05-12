# -*- coding: utf-8 -*-
"""
Created on 23 Aug 2019
@author: Swapnil Jiaswal
"""
import imaplib
import os
import email
from datetime import date, timedelta
import datetime
# import send_mail
import pandas as pd
import numpy as np
#import leela_validation
email_user = 'revseed@revnomix.com'
email_pass = 'Revenue@123'
imap_url = 'imap.gmail.com'
day =1
attachment_path=r'E:\Yadnesh\Revseed_HNF\email_data'
last_day_date= (date.today() - timedelta(days=day)).__format__('%Y-%m-%d')
today = date.today()
# dd/mm/YY
d1 = today.strftime("%Y-%m-%d")
path = os.path.join(attachment_path, last_day_date)
if(os.path.exists(path)):
    print("Allready Exist")
else:
    os.mkdir(path)
map = pd.read_excel(r'E:\Yadnesh\Revseed_HNF\Mapping\MailMapping.xlsx')

def emailAttachment():
    def auth(email_user,email_pass,imap_url):
        con = imaplib.IMAP4_SSL(imap_url)
        con.login(email_user,email_pass)
        return con
    def get_body(msg):
        if msg.is_multipart():
            return get_body(msg.get_payload(0))
        else:
            return msg.get_payload(None,True)
    def get_attachment(msg,sub):
        test = (msg['Date'])
        print(test)
        try:
            d = datetime.datetime.strptime(test, '%a, %d %b %Y %H:%M:%S +%f (IST)').strftime('%Y-%m-%d')
        except:
            try:
                d = datetime.datetime.strptime(test, '%d %b %Y %H:%M:%S +%f').strftime('%Y-%m-%d')
            except:
                d = datetime.datetime.strptime(test, '%a, %d %b %Y %H:%M:%S +%f').strftime('%Y-%m-%d')
        d1 = datetime.datetime.strftime(date.today(),'%Y-%m-%d')
        print(d)
        print(d1)
        if (d == d1):
            for part in msg.walk():
                if part.get_content_maintype()=="multipart":
                    continue
                fileName = part.get_filename()
                if bool(fileName):
                   ext= fileName.split(".")[-1]
                   if ext in ['txt','xls','xlsx']:
                        filePath = os.path.join(path, fileName )
                        with open(filePath,'wb') as f:
                            f.write(part.get_payload(decode=True))
                            fl_name = sub
                            map['Status'] = np.where(map['Hotel Name'] == fl_name, 'Success', map['Status'])

        else:
            fl_name = sub
            print(fl_name)
            map['Status'] = np.where(map['Hotel Name'] == fl_name, 'Failed', map['Status'])
            ema_name = dict(zip(map['Hotel Name'], map['Email Id']))
            email_id = ema_name[fl_name]
            # send_mail.send_alert_msg(fl_name, "missing on", d1, email_id)
    #

    con = auth(email_user,email_pass,imap_url)
    # con = auth(email_user,email_pass,imap_url)
    con.select('INBOX')


    sub_list= ['GRR-JW CCU','STR Report-JW','GRR - CY IXB','Reservation Reports-JW CCU',
               'Reservation Reports- CY IXB']
    for sub in sub_list:
        result, data =con.uid('search',None,'Subject "{}"'.format(sub))
        inbox_item_list=data[0].split()
        for item in inbox_item_list:
            result2, email_data = con.uid('fetch',item,'(RFC822)')
            raw_email = email_data[0][1].decode('utf-8')

        email_message = email.message_from_string(raw_email)
        # if email_message['Subject']==sub:
        t =(email_message['Date'])
        print(t)
        get_attachment(email_message,sub)
    map.to_excel(r'D:\Swapnil All Program\alembic_report\Mapping/ExpMailMapping.xlsx', index=False)
    con.logout()

if __name__ =='__main__':
    emailAttachment()
