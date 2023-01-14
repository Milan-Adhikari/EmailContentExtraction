from datetime import datetime
import unicodedata
import imaplib
import string
import email
import os

DATE_F = '%d-%b-%Y'

def get_new_attachments(email_address: str, password:str , white_list: list[str], file_path: str, start_time:str, end_time:str, delete_read: bool) -> None:
    """Saves the attachment from the mail of given email from the whitelisted accounts

    Args:
        email_address (str): email address of the account
        password (str): password of the account
        white_list (list[str]): list of emails to look for attachments and save from
        file_path (str): name of directory to save the attachments
        start_time (str): beginning time to look email from
        end_time (str): final time to look email till
        delete_read (bool): if to delete the mail or not
    """
    #Check if the attachments folder exists or not, if not then create the folder to store the downloaded attachments
    if not os.path.isdir(f"{os.getcwd()}/{file_path}"):
        try:
            os.mkdir(f"{os.getcwd()}/{file_path}")
        except Exception as e:
            #Couldnt make folder for some reason
            print(f"Couldnt make the attachments folder in the current directory. Error: {e}")
            return
    
    # connect to the gmail account
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(email_address, password)
    mail.select("inbox")  # selecting the inbox folder
    
    _, data = mail.uid('search', None, '(SINCE "{}" BEFORE "{}")'.format(start_time, end_time))  # searching for the emails arrived between the start and end time
    inbox_item_list = data[0].split()
    
    attachments_saved = 0  # number of attachments saved
    for item in inbox_item_list:
        _, email_data = mail.uid('fetch', item, '(RFC822)')
        raw_email = email_data[0][1].decode("utf-8")
        email_message = email.message_from_string(raw_email)
        print(email.utils.parseaddr(email_message['From'])[1])
        if email.utils.parseaddr(email_message['From'])[1] in white_list:  # executed only if the email is in white list
            for part in email_message.walk():   
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                file_name = part.get_filename()
                if bool(file_name):
                    attachments_saved += 1
                    # save_path = os.path.join(file_path, email_message['From'], "_", str(email_message['Date']), "_",str(attachments_saved))
                    safe_file_string = clean_filename(f"{email_message['From']}_{str(email_message['Date'])}")
                    try:
                        os.mkdir(f"{os.getcwd()}/{file_path}/{safe_file_string}")
                    except: pass
                    save_path = os.path.join(f"{os.getcwd()}/{file_path}/{safe_file_string}", f"{attachments_saved}.{file_name.split('.')[-1]}")
                    with open(save_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))
    # to delete the messages after downloading the attachments        
    if delete_read:
        for item in inbox_item_list:
            mail.uid('store', item, '+FLAGS', '\\Deleted')
        mail.expunge()
    mail.close()
    mail.logout()


valid_filename_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
char_limit = 255
def clean_filename(filename, whitelist=valid_filename_chars, replace=' '):
    for r in replace:
        filename = filename.replace(r,'_')
    # keep only valid ascii chars
    cleaned_filename = unicodedata.normalize('NFKD', filename).encode('ASCII', 'ignore').decode()
    # keep only whitelisted chars
    cleaned_filename = ''.join(c for c in cleaned_filename if c in whitelist)
    if len(cleaned_filename)>char_limit:
        print("Warning, filename truncated because it was over {}. Filenames may no longer than".format(char_limit))
    return cleaned_filename[:char_limit]   

email_id = '<email-address>'
pwd = '<app-password>'

email_white_list = ['<whitelisted.email1.com>', 'whilelisted.email2.com']

file_path = 'attachments'

start_time = '2023-01-15 00:00:33'
end_time = '2023-01-16 18:00:00'

start_time= datetime.fromisoformat(start_time)
end_time = datetime.fromisoformat(end_time)


get_new_attachments(email_id, pwd, email_white_list, file_path, start_time.strftime(DATE_F), end_time.strftime(DATE_F), False)

