import pandas as pd
from pandas import DataFrame
#from matplotlib import pyplot as plt
from pandas.api.types import CategoricalDtype
import datetime
import numpy as np
import xlrd
import xlsxwriter


import smtplib
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#Data connection...

#data = pd.read_excel("cardnew.xlsx")


data =pd.read_csv("bkapp.csv",encoding='latin1')
# data .drop(data .columns[[0]], axis=1, inplace=True)





# data = data[(data['RESPONSE_MEANING'] == 'Issuer or switch inoperative')]
fx = data[['Year_Month','Transactions','Amount_Transacted','Channel']]



#dd1 = fx['TERMINAL_ID'].tolist()
dd1 = ['d1','d2','d3']
dd = 8


writer = pd.ExcelWriter(r'Failedtrnx.xlsx', engine='xlsxwriter')
sheets_in_writer=['FailedTransactions']
data_frame_for_writer=[fx]

for i,j in zip(data_frame_for_writer,sheets_in_writer):
    i.to_excel(writer,j,index=False,startrow=5 )
workbook=writer.book
# Add a header and amount format--------------------------------------------------
header_template = workbook.add_format({'font_color': '#ffffff','bg_color': '#0D4DA2','bold': True,'border': 0 })
Amount_template = workbook.add_format({'num_format': '#,##0;(#,##0)','border': 0})
Amount_ID = workbook.add_format({'border': 0 }  )
# ---------------------------------------------------------------------------------
for i,j in zip(data_frame_for_writer,sheets_in_writer):
    for col_num, value in enumerate(i.columns.values):
        writer.sheets[j].write(5, col_num, value, header_template)
        #writer.sheets[j].hide_gridlines(2)
        writer.sheets[j].freeze_panes(6, 4)
        writer.sheets[j].set_column('A:A', 25)
        writer.sheets[j].set_column('D:Y', 17)
        writer.sheets[j].insert_image('A1', 'BK.png',{'x_scale': 0.6, 'y_scale': 0.2})
        #writer.sheets[j].autofilter(0,0,0,i.shape[1]-1)  #(up, left)
        writer.sheets[j].autofilter(5,0,5,i.shape[1]-1) 
        writer.sheets[j].conditional_format('F26:D7',{ 'type': 'cell', 'criteria': '<>', 'value':    '"None"','format':   Amount_template} )
        writer.sheets[j].conditional_format('A7:C26',{'type': 'cell',  'criteria': '<>','value':    '"None"', 'format':   Amount_ID  } )
    title = writer.book.add_format({
        'bold': True,
        'font_size': 14,
        'align': 'left',
        'valign': 'vcenter',
        'font_color': '#ffffff',
        'bg_color': '#0D4DA2',
        'border': 0
        })
         
    title_text1 = u"{0} {1}".format(("Failed transactions upto"),max(data['DATE']))
    writer.sheets[j].merge_range('A3:D4', title_text1, title)

        
        
writer.save()






MY_ADDRESS = 'elhabimana'
PASSWORD = 'Greece16@'


def get_contacts(filename):
    """
    Return two lists names, emails containing names and email addresses
    read from a file specified by filename.
    """
    
    names = []
    emails = []
    with open(filename, mode='r', encoding='utf-8') as contacts_file:
        for a_contact in contacts_file:
            names.append(a_contact.split()[0])
            emails.append(a_contact.split()[1])
    return names, emails

def read_template(filename):
    """
    Returns a Template object comprising the contents of the 
    file specified by filename.
    """
    
    with open(filename, 'r', encoding='latin-1') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

def main():
    if dd > 5:
         print("Sorry we can't send that email")
        
    else:
        names, emails = get_contacts(r'mycontacts.txt') # read contacts
        message_template = read_template(r'message.txt')

        # set up the SMTP server
        s = smtplib.SMTP(host='smtp.office365.com', port=587)
        s.starttls()
        s.login(MY_ADDRESS, PASSWORD)

        # For each contact, send the email:
        for name, email in zip(names, emails):
            msg = MIMEMultipart()       # create a message
            # add in the actual person name to the message template
            message = message_template.substitute(PERSON_NAME=name.title(), dd=dd, dd1=dd1)

            # Prints out the message body for our sake
            print('Done')

            # setup the parameters of the message
            msg['From']=MY_ADDRESS
            msg['To']=email
            msg['Subject']="Card Monitoring Alert"

            # add in the message body
            msg.attach(MIMEText(message, 'plain'))

            # send the message via the server set up earlier.
            filename12='Failedtrnx.xlsx'
            attachment  =open(filename12,'rb')
            part = MIMEBase('application','octet-stream')
            part.set_payload((attachment).read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',"attachment; filename= "+filename12)
            msg.attach(part)
            s.send_message(msg)
            #del msg


        # Terminate the SMTP session and close the connection
        s.quit()

if __name__ == '__main__':
    main()

    
