
"""
Created on 23 Aug 2019
@author: Swapnil Jiaswal

Upgraded on 16 Oct 2021
@author: Jigar Bhatt
"""

import imaplib
import os
import email
from datetime import date
import datetime
#import send_mail
import pandas as pd
import numpy as np
# import database
import re


email_user = 'revseed@revnomix.com'
email_pass = 'Revenue@123'
imap_url = 'imap.gmail.com'


#attachment_path='E:/Jigar/ftp/email_data'
path='C:\Revseed_HNF\Input'
# today = datetime.date.today()
# d1 = today.strftime("%Y-%m-%d")
# path = os.path.join(attachment_path, d1)
# if(os.path.exists(path)):
#     print("Allready Exist")
# else:
#     os.mkdir(path)


#map =pd.read_excel(r'C:/ftp/Email_Common_Mapping/Oceanic_mapping_file.xlsx')

def get_sub_file_name( ):
    df = pd.read_excel("C:\Revseed_HNF\HNF_code\masters/autodownlaod_mail.xlsx")
    # sub = df["Subject_name"].to_list()
    # fl_name = df["File_name"].to_list()
    file_dict = dict(zip(df["Subject_name"], df["File_name"]))
    sub_list = list(file_dict.keys())
    file_list = list(file_dict.values())

    return sub_list, file_list, file_dict, df


def auth(email_user, email_pass, imap_url):
    con = imaplib.IMAP4_SSL(imap_url)
    con.login(email_user, email_pass)
    return con
def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None,True)

def get_attachment(msg,sub,file_list, file_dict,map_df):
    date_tuple = email.utils.parsedate_tz(msg['Date'])
    if date_tuple:
        test = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))

    d = test
    # d = test.strftime("%Y-%m-%d %H:%M:%S")

    d1 = datetime.datetime.today()
    d_today = d1.strftime("%Y-%m-%d %H:%M:%S")
    # d_today = d1

    last_updated_date_ = map_df["Updated_Date"][map_df["Subject_name"] == sub].reset_index(drop=True)[0]
    last_updated_date = pd.to_datetime(last_updated_date_)
    if last_updated_date.date() < d1.date():
        map_df['Status'] = np.where(map_df['Subject_name'] == sub, 0, map_df['Status'])

    status_in_df = map_df["Status"][map_df["Subject_name"] == sub].reset_index(drop=True)[0]


    # application
    if (d.date() == d1.date()):
        if status_in_df == 0:
            for part in msg.walk():
                if part.get_content_maintype() in ["multipart","text","image"]:
                    continue
                fileName = part.get_filename()
                fileType = part.get_content_type()
                if bool(fileName):
                    if fileName in file_list:
                        filePath = os.path.join(path, fileName)
                        with open(filePath,'wb') as f:
                            f.write(part.get_payload(decode=True))
                            print('{} file downloaded successfully'.format(fileName))
                            # fl_name = sub
                            map_df['Status'] = np.where(map_df['Subject_name'] == sub, 1, map_df['Status'])
                            map_df['Updated_Date'] = np.where(map_df['Subject_name'] == sub,d_today,map_df['Updated_Date'])
                    elif fileType in ['file/txt','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','application/octet-stream']:
                        file_name_reqd = file_dict.get(sub)
                        fileName = fileName.replace(fileName, file_name_reqd)
                        filePath = os.path.join(path, fileName)
                        with open(filePath, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                            print('{} file downloaded successfully'.format(fileName))
                            # fl_name = sub
                            map_df['Status'] = np.where(map_df['Subject_name'] == sub, 1, map_df['Status'])
                            map_df['Updated_Date'] = np.where(map_df['Subject_name'] == sub,d_today,map_df['Updated_Date'])

        else:
            map_df['Updated_Date'] = np.where(map_df['Subject_name'] == sub, map_df['Updated_Date'],
                                              map_df['Updated_Date'])
            print("Already Downloaded")
            pass
    else:
        map_df['Updated_Date'] = np.where(map_df['Subject_name'] == sub, map_df['Updated_Date'],map_df['Updated_Date'])
        pass
        # fl_name = sub
        # map['Status'] = np.where(map['Hotel Name'] == fl_name, 'Failed', map['Status'])
        # ema_name = dict(zip(map['Hotel Name'], map['Email Id']))
        # email_id = ema_name[fl_name]
        # send_mail.send_alert_msg(fl_name, "missing on", d1, email_id)

    map_df.to_excel(r'C:\Revseed_HNF\HNF_code\masters/autodownlaod_mail.xlsx', index=False)


def emailAttachment():
    sub_list, file_list, file_dict,map_df = get_sub_file_name()
    con = auth(email_user,email_pass,imap_url)

    #con.select('INBOX')
    con.select('"[Gmail]/All Mail"')

    for sub in sub_list:
        result, data =con.uid('search',None,'SUBJECT "{}"'.format(sub))
        inbox_item_list=data[0].split()
        item = sorted(inbox_item_list, reverse=True)
        if len(item) == 0:
            continue
        else:
            item = item[0]

        result2, email_data = con.uid('fetch',item,'(RFC822)')
        raw_email = email_data[0][1].decode('utf-8')

        # for item in inbox_item_list:                                 # commented and moved out of loop J.B. 29Sept'2021
        #     result2, email_data = con.uid('fetch',item,'(RFC822)')
        #     raw_email = email_data[0][1].decode('utf-8')

        email_message = email.message_from_string(raw_email)
        # t =(email_message['Date'])
        # print(t)
        get_attachment(email_message,sub,file_list, file_dict, map_df)
    #map.to_excel(r'C:/ftp/Email_Common_Mapping/Oceanic_mapping_file.xlsx', index=False)
    con.logout()

if __name__ =='__main__':
    emailAttachment()
