from datetime import datetime
import imaplib
import email
import os

DATE_F = '%d-%b-%Y'

def get_new_attachments(email_address, password, white_list, file_path, start_time, end_time, delete_read=False):
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(email_address, password)
    mail.select("inbox")
    
    result, data = mail.uid('search', None, '(SINCE "{}" BEFORE "{}")'.format(start_time, end_time))
    # argum = bytes(f'SINCE {start_time} BEFORE {end_time}', 'utf-8')
    # result, data = mail.search(bytes(f'SINCE {start_time} BEFORE {end_time}', 'utf-8'))  #['SINCE' , start_time, 'BEFORE', end_time]
    print(data[0])
    inbox_item_list = data[0].split()
    
    attachments_saved = 0
    for item in inbox_item_list:
        result2, email_data = mail.uid('fetch', item, '(RFC822)')
        raw_email = email_data[0][1].decode("utf-8")
        # print(type(raw_email))
        email_message = email.message_from_string(raw_email)
        # email_message = email.message_from_bytes(raw_email)
        print(type(email_message['From']))
        # email_message_from = 
        # check if sender is in white 9list
        # if email_message['From'] in white_list:
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
                        
    if delete_read:
        for item in inbox_item_list:
            mail.uid('store', item, '+FLAGS', '\\Deleted')
        mail.expunge()
    mail.close()
    mail.logout()


email_id = ''
pwd = ''

email_white_list = ['maxacem.edu.np']

file_path = 'attachments/'

start_time = '2023-01-12 12:08:33'
end_time = '2023-01-13 13:00:00'

start_time= datetime.fromisoformat(start_time)
end_time = datetime.fromisoformat(end_time)


get_new_attachments(email_id, pwd, email_white_list, file_path, start_time.strftime(DATE_F), end_time.strftime(DATE_F))

