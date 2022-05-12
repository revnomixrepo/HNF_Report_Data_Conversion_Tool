import os
import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd

def send_Hnf(fl_name, today):
    # creates SMTP session
    # text, pickup_df, segment_pickup_df_updated = summary(hf_df, ms_df, date)
    eid_data = pd.read_excel("C:\Revseed_HNF\HNF_code\masters/send_email.xlsx")
    # eid_h = eid_data['email_id'][eid_data['hotelcode'] == fl_name]
    eid_h_to = eid_data['email_id'][(eid_data['hotelcode'] == fl_name) & (eid_data['Type'] == 1)]
    eid_h_cc = eid_data['email_id'][(eid_data['hotelcode'] == fl_name) & (eid_data['Type'] == 2)]


    msg = MIMEMultipart()

    text_body = "Dear Team," \
                "\r\n\r\n" \
                "Greetings of the Day!" \
                "\r\n\r\n" \
                "Please find the attached HnF Report: {}.".format(fl_name)

  #   html = """\
  #   <html>
  #   <head>
  #   <style>
  #   th {{height: 40px; background-color: #606d99 !important; font-weight: bold;}}
  #   td {{vertical-align: center;}}
  #   table {{width: 75%; border-collapse: collapse; height: 120px; font-size: 20px;margin-left: auto;
  # margin-right: auto;}}
  #   table, th, td {{border:1px solid black; text-align: center  !important;}}
  #   </style>
  #   </head>
  #   <body style="padding-left: 80px;
  #               padding-right: 80px;
  #               padding-top: 50px;
  #               padding-bottom: 50px;
  #               font-size: 20px;
  #               text-align: center !important;">
  #   <p><br>
  #   {0}
  #  <br>
  #   {1}
  #   <br>
  #   {2}
  #   <br>
  #   {3}
  #   </p>
  #   </body>
  #   </html>
  #   """
  #
  #

        # .format(text,  month_df.to_html(index=False),pickup_df.to_html(index=False), segment_pickup_df.to_html(index=False))

    msg.attach(MIMEText(text_body, 'plain'))
    # msg.attach(MIMEText(html, 'html'))


    part = MIMEBase('application', "octet-stream")
    path = "C:\Revseed_HNF\Output"
    # rdate = datetime.datetime.strftime(run_date.date, '%d-%m-%Y')
    rdate = today
    # path = os.path.join(path + "/" + rdate)
    path = os.path.join(path + "/" + str(rdate))
    # files_contains = msgs+"_"+fl_name.lower()
    files = os.listdir(path)
    # for f in reversed(files):
    #     if f.lower().__contains__("F_" + fl_name.lower()) == True:
    #         path = os.path.join(path + "/" + str(f))
    #         break
    # print(path)
    for f in reversed(files):
        if f.lower().__contains__(fl_name.lower()) == True:
            path = os.path.join(path + "/" + str(f))
            break
    part.set_payload(open(path, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename={}'.format(f))
    msg.attach(part)

    s = smtplib.SMTP('smtp.gmail.com', 587)

    # start TLS for security
    s.starttls()
    # sender = "revseed@revnomix.com"
    # sender = 'Revnomix Revenue Management Services <revnomix.RMS@revnomix.com>'
    sender = "yadnesh.kolhe@revnomix.com"

    # recipients = ['abhijeet.rode@revnomix.com','paritosh.palkar@revnomix.com','jigar.bhatt@revnomix.com']
    # recipients = eid_h.to_list()
    recipients_to = eid_h_to.to_list()
    recipients_cc = eid_h_cc.to_list()
    # rdate1 = datetime.datetime.strftime(run_date.date, '%d %b %Y')
    rdate1 = today
    # msg['Subject'] = fl_name.upper() + ' ' + msgs + ' ' + rdate
    # msg['Subject'] = msgs + " | " + fl_name.upper() + " | " + rdate1
    msg['Subject'] = "HNF" + " | " + fl_name + " | " + str(rdate1)
    msg['From'] = sender
    # msg['To'] = ' ,'.join([str(elem) for elem in recipients.split(',')[:1]])
    # msg['To'] = ", ".join(recipients)
    msg['To'] = ", ".join(recipients_to)
    msg['Cc'] = ", ".join(recipients_cc)
    # msg['reply-to'] = ", ".join(recipients)

    # msg['Cc'] = ' ,'.join([str(elem) for elem in recipients.split(',')[1:]])
    # Authentication
    # s.login("revseed@revnomix.com", "Revenue@123")
    s.login("yadnesh.kolhe@revnomix.com", "yadnesh@15")
    s.sendmail(sender, recipients_to+recipients_cc, msg.as_string())
    # s.sendmail(sender,  recipients.split(','), msg.as_string())
    s.quit()

#
# if __name__ == '__main__':
#     fl_name = 'TheResort'
# #     # msgs = 'isell'
#     today = '2022-03-10'
#     # email_id = ['abhijeet.rode@revnomix.com','paritosh.palkar@revnomix.com']
# #     # send_Hnf(fl_name, msgs, d_file)
#     send_Hnf(fl_name, today)