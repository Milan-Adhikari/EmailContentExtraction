"""
Author: Milan Adhikari
Date: 2023-01-17
github: https://github.com/Milan-Adhikari
"""


from datetime import datetime
import imaplib
import email
import os


DATE_F = '%d-%b-%Y'


def get_new_attachments(email_address, password, white_list, file_path, start_time, end_time, delete_read=False):
    # make a directory if it doesnt exist
    if not os.path.exists(file_path):
        os.makedirs(file_path)
    # connect to the gmail 
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(email_address, password)
    mail.select("inbox")  # selecting the inbox folder
    # searching for emails from start time to the end time
    result, data = mail.uid('search', None, '(SINCE "{}" BEFORE "{}")'.format(start_time, end_time))  
    inbox_item_list = data[0].split()
    attachments_saved = 0  # number of attachments saved
    for item in inbox_item_list:
        result2, email_data = mail.uid('fetch', item, '(RFC822)')
        raw_email = email_data[0][1].decode("utf-8")
        # print(type(raw_email))
        email_message = email.message_from_string(raw_email)
        print(type(email_message['From']))
        print(email.utils.parseaddr(email_message['From'])[1])
        if email.utils.parseaddr(email_message['From'])[1] in white_list:
            for part in email_message.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                file_name = part.get_filename()
                if bool(file_name):
                    attachments_saved += 1
                    save_path = os.path.join(file_path, email_message['From']+"_"+str(email_message['Date'])+"_"+str(attachments_saved))
                    with open(save_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))
    # if the flag is set to true, then delete the email from which the attacfhments are saved               
    if delete_read:
        for item in inbox_item_list:
            mail.uid('store', item, '+FLAGS', '\\Deleted')
        mail.expunge()
    mail.close()
    mail.logout()
    
# enter the app password and not the gmail password
email_id = ''
pwd = ''

# the list of emails to accept the mail from
email_white_list = ['adwinastra@gmail.com', 'whilelisted.email2.com']

file_path = 'attachments/'

# set the appropriate start time and end time
start_time = '2023-01-17 00:08:33'
end_time = '2023-01-18 19:00:00'

# converting to the gmail acceptable form of the date i.e ISO format
start_time= datetime.fromisoformat(start_time)
end_time = datetime.fromisoformat(end_time)

# initializing the main function
get_new_attachments(email_id, pwd, email_white_list, file_path, start_time.strftime(DATE_F), end_time.strftime(DATE_F))

