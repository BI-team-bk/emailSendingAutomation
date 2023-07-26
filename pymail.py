import smtplib
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import datetime
import pandas as pd
import numpy as np
import locale


MY_ADDRESS = 'elhabimana@bk.rw'
PASSWORD = 'Greece16@'


def get_contacts(filename):
    """
    Return two lists names, emails containing names and email addresses
    read from a file specified by filename.
    """
    
    names = []
    emails = []
    with open(filename, mode='r', encoding='utf-8') as contacts_file:
    #    lqq for a_contact in contacts_file:
    #         names.append(a_contact.split()[0])
    #         emails.append(a_contact.split()[1])
        for a_contact in contacts_file:
            # Split the line and strip any leading/trailing whitespace
            contact_info = a_contact.strip().split()
            # Check if there are at least two elements (name and email)
            if len(contact_info) >= 2:
                names.append(contact_info[0])
                emails.append(contact_info[1])
            else:
                # Handle the case when the line format is not as expected
                print(f"Error: Invalid contact information in line '{a_contact}'")
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
    data = pd.read_csv("bkapp.csv", encoding='latin1')

    # Clean the 'Transactions' and 'Amount_Transacted' columns by removing commas and whitespace
    # data['Transactions'] = data['Transactions'].str.replace(',', '').str.strip()
    # data['Amount_Transacted'] = data['Amount_Transacted'].str.replace(',', '').str.strip()

    # Convert the 'Transactions' and 'Amount_Transacted' columns to np.int64
    data['Transactions'] = data['Transactions'].astype(np.int64)
    data['Amount_Transacted'] = data['Amount_Transacted'].astype(np.int64)

    data['AUTH_DATE'] = pd.to_datetime(data['AUTH_DATE'], format='%Y%m%d')
    min_date = data['AUTH_DATE'].min().strftime("%Y-%m-%d")
    max_date = data['AUTH_DATE'].max().strftime("%Y-%m-%d")

    # Create a summary table by filtering data in the "channel" column and calculating the total values
    summary_table = data.groupby('Channel').agg({
    'Transactions': 'sum',
    'Amount_Transacted': 'sum',
    'DEBIT_CUSTOMER': 'nunique'
    }).reset_index()

    # Rename the columns
    summary_table = summary_table.rename(columns={
        'Transactions': 'TOTAL TRANSACTIONS',
        'Amount_Transacted': 'AMOUNT TRANSACTED',
        'DEBIT_CUSTOMER': 'USERS'
    })
    
     # Set the locale to use the comma as the thousands separator
    locale.setlocale(locale.LC_ALL, '')

    # Apply formatting with commas to the 'Transactions' and 'Amount_Transacted' columns
    summary_table['TOTAL TRANSACTIONS'] = summary_table['TOTAL TRANSACTIONS'].apply(lambda x: locale.format_string("%d", x, grouping=True))
    summary_table['AMOUNT TRANSACTED'] = summary_table['AMOUNT TRANSACTED'].apply(lambda x: locale.format_string("%d", x, grouping=True))

    names, emails = get_contacts('mycontacts.txt') # read contacts
    message_template = read_template('message.txt')
    # set up the SMTP server
    s = smtplib.SMTP(host='smtp.office365.com', port=587)
    s.starttls()
    s.login(MY_ADDRESS, PASSWORD)

    # For each contact, send the email:
    for name, email in zip(names, emails):
        msg = MIMEMultipart()       # create a message
        # summary_table_html = summary_table.to_html()
        # Convert the summary table to an HTML table with aligned headers and commas between values
        summary_table_html = summary_table.to_html(index=False,justify='center', classes='dataframe')
        summary_table_html = summary_table_html.replace('<table', '<table style="border-collapse: collapse; width: 100%;"')
        summary_table_html = summary_table_html.replace('<thead>', '<thead style="background-color: #f2f2f2;">')
        summary_table_html = summary_table_html.replace('<th>', '<th style="text-align: center; padding: 5px; border: 1px solid #ddd;background-color:#03529c;color:#f2f2f2">')
        summary_table_html = summary_table_html.replace('<td>', '<td style="text-align: center; padding: 5px; border: 1px solid #ddd;">')
        message = message_template.substitute(
            PERSON_NAME=name.title(),
            MIN_DATE=min_date,
            MAX_DATE=max_date,
            SUMMARY_TABLE=summary_table_html)

        # Prints out the message body for our sake
        print('Done')

        # setup the parameters of the message
        msg['From']=MY_ADDRESS
        msg['To']=email
        msg['Subject']="Trend between old bkApp and new bkApp"
        
        # add in the message body
        msg.attach(MIMEText(message, 'html'))
        # send the message via the server set up earlier.
        filename12='bkapp.csv'
        attachment  =open(filename12,'rb')
        part = MIMEBase('application','octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',"attachment; filename= "+filename12)
        msg.attach(part)
        s.send_message(msg)
        #del msg

        
        # send the message via the server set up earlier.
        s.send_message(msg)
        #del msg
        
        
    # Terminate the SMTP session and close the connection
    s.quit()
    
if __name__ == '__main__':
    main()
