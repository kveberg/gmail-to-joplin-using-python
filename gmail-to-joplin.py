""""
Retrieves unread e-mails (from, subject, main text and attachments) from a Gmail account, imports them to a notebook Joplin and initiates Joplin synchronization.

 Requirements
 - Python 3
 - Joplin CLI https://joplinapp.org/terminal/
 - Gmail account with API enabled https://console.cloud.google.com/home/
 - OAuth client credentials from said Gmail account stored as credentials.json the same folder as this script
 - A target notebook in Joplin CLI with the same name as in the "JOPLIN_INBOX" variable in the CONFIG-section below.
 
 Consider adding pre-approved addresses to the "APPROVED_SENDERS" variable in the CONFIG-section below, to avoid importing unwanted e-mails.
"""

# IMPORTS
import os.path
import shutil
from datetime import datetime
import base64
import email

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# CONFIG:
# Target Joplin notebook
JOPLIN_INBOX = "_inbox"
# List of addresses you want to import mails from (leave blank if you want everything)
APPROVED_SENDERS = []
# A log-file (log.txt) is created in the script directory. You can specify the number of entries it should hold here.
LOG_SIZE = 100
# A download folder to temporarily hold e-mails and attachments
DOWNLOAD_PATH = f"{os.getcwd()}\\gmtj_downloads"
# Debug boolean. If True, Joplin will not be called to sync and unread mails will not be moved to trash.
DEBUG = False


# PREPARE FOR DOWNLOADS
if os.path.exists(DOWNLOAD_PATH):
    shutil.rmtree(DOWNLOAD_PATH)
    os.mkdir(DOWNLOAD_PATH)
else:
    os.mkdir(DOWNLOAD_PATH)


"""
TALK TO JOPLIN
"""

def import_to_joplin(id, subject, text, attachments):
    """ Imports text and attachments to Joplin-notebook, with subject as title """
    
    # Store text in file using its unique id
    path = f"gmtj_downloads\\{id}\\"
    if not os.path.exists(path):
        os.makedirs(path)
    
    text = text.encode()
    text_file = str(path + id + ".md")
    with open(text_file, "wb") as f:
        f.write(text)

    # Import text as new note in joplin_inbox
    path_to_text = os.getcwd() + "\\" + text_file
    os.system("joplin import \"" + path_to_text + "\" " + JOPLIN_INBOX)
    
    # Append attachments using list of attachment paths
    if attachments:
        for i in attachments:
            os.system("joplin attach " + "\"" + id + "\" \"" + i + "\"")
    
    # Use GMail subject to set title
    os.system("joplin set \"" + id +"\"" + " title \"" + subject + "\"")

    return


""" 
TALK TO GMAIL
"""

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]
creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())


def check_gmail():
    """ Returns list of Gmail ids for any unread mail. """
    try:
        service = build("gmail", "v1", credentials=creds)
        
        result = service.users().messages().list(userId="me", q="is:unread").execute()
        
        # Print number of e-mails
        n_new_mails = result["resultSizeEstimate"]
        print(str(n_new_mails) + " unread mail(s) found!")
        if n_new_mails == 0:
            quit()
        
        # Get id's
        messages = result["messages"]
        ids = []
        for i in messages:
            ids.append(i["id"])
        
        # Return id's
        return ids
    
    except Exception as e:
        write_to_log(type="ERROR", entry=f"Unknown exception when calling check_gmail(). {e}")
        quit()


def import_gmail(id):
    """ Downloads Gmail, and passes and any attachments to import_to_joplin() """

    # Get GMail
    service = build("gmail", "v1", credentials=creds)
    msg = service.users().messages().get(userId="me", id=str(id), format="raw").execute()

    # Convert to MIME
    msg_bytes = base64.urlsafe_b64decode(msg['raw'])
    mime_msg = email.message_from_bytes(msg_bytes)

    # Get subject
    subject = mime_msg["subject"]
    if email.header.decode_header(subject)[0][1] != None:
        subject = email.header.decode_header(subject)[0][0]
        subject = subject.decode()

    # Get sender
    sender = mime_msg["from"]
    sender = sender[sender.index("<")+1:sender.index(">")]
    sender = sender.lower()

    # Trash if not from approved senders
    if APPROVED_SENDERS:
        if sender not in APPROVED_SENDERS:
            if not DEBUG:
                service.users().messages().trash(userId="me", id=str(id)).execute()
            write_to_log(type="ERROR", entry=f"{sender} not in list of approved senders. E-mail {id} put in trash.")
            print(f"{sender} not in list of approved senders. E-mail {id} put in trash.")
            
            return False

    # Get text
    for part in mime_msg.walk():
        try:
            if part.get_content_type()=="text/plain":
                text = part.get_payload(decode=True)
                text = text.decode()
        except Exception as e:
            text_decode_error = f" #Decode-Error!\n\"{e}\"\nThings may not look exactly right."
            problem_text = part.get_payload(decode=True)
            text = text_decode_error + problem_text.decode("utf-8", "replace")
            write_to_log(type="ERROR", entry=f"DECODE ERROR {id}\t FROM  {sender}m \t\t\t {subject}. {e}")
            

    # Get attachments
    att_filenames = []
    for part in mime_msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        if part.get("Content-Disposition") is None:
            continue

        filename = part.get_filename()
        if email.header.decode_header(filename)[0][1] != None:
            filename = email.header.decode_header(filename)[0][0]
            filename = filename.decode()
        data = part.get_payload(decode=True)
        
        file_path = f"{DOWNLOAD_PATH}\\{filename}"
        att_filenames.append(file_path)
        with open(file_path, "wb") as f:
                f.write(data)

    print(f"Passing to Joplin: {id} with {len(att_filenames)} attachment(s) from {sender}.")

    # Import to Joplin
    import_to_joplin(id, subject, text, att_filenames)
    
    # Thrash Gmail
    if not DEBUG:
        service.users().messages().trash(userId="me", id=str(id)).execute()

    # Log
    write_to_log(type="IMPORT", entry=f"{id} with {len(att_filenames)} attachments.\tFROM {sender}.\t\tSUBJECT \"{subject}\".")

    return True

"""
WRITE LOG
"""
def write_to_log(type, entry):
    """ Writes entry to log.txt, and deletes old entries """
    timestamp = datetime.now()
    timestamp = timestamp.strftime("%d-%m-%Y %H:%M:%S")

    with open("log.txt", "r") as f:
        data = f.read().splitlines(True)
        if len(data) >= LOG_SIZE:
            with open("log.txt", "w") as f:
                f.writelines(data[len(data)-LOG_SIZE+1:])
    log_text = f"{type}\t{timestamp}\t{entry}\n"

    with open("log.txt", "a") as f:
        f.write(log_text)
    
    return
  
   
"""
RUN THE SCRIPT
"""
# Lower case to avoid false negatives when get_gmail() checks for them
for i in range(len(APPROVED_SENDERS)):
    APPROVED_SENDERS[i] = APPROVED_SENDERS[i].lower()

# "gmtj_downloads"-folder to temporarily hold data should not exist, but things can happen

if not os.path.exists("credentials.json"):
    print("Could not find OAUTH file \"credenials.json\".")
    quit()

# Check for new e-mails
new_gmails = check_gmail()

# Import e-mails, if any
counter = 0

if new_gmails:
    for i in new_gmails:
        get_it = import_gmail(i)
        if get_it:
            counter +=1
            
print(f"Imported {counter} of {len(new_gmails)} unread mails.")

# Joplin sync
if new_gmails and not DEBUG:
    os.system("joplin sync")

# Clean up download folder
print("Cleaning temporary download folder ...")
shutil.rmtree(DOWNLOAD_PATH)
