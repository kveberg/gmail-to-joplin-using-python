""""
Retrieves unread e-mails (from, subject, main text and attachments) from a Gmail account, imports them to a notebook Joplin and initiates Joplin synchronization.

 Requirements
 - Python 3
 - Joplin CLI https://joplinapp.org/terminal/
 - Gmail account with API enabled https://console.cloud.google.com/home/
 - OAuth client credentials from said Gmail account stored as credentials.json the same folder as this script
"""
# IMPORTS
import os.path
import subprocess
import shutil
import base64
import email
import logging
from datetime import datetime
from logging.handlers import RotatingFileHandler

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# CONFIG:
JOPLIN_INBOX = "_inbox"                              # Target Joplin notebook. Script creates it if it does not already exist
APPROVED_SENDERS = []                                # List of addresses you want to import mails from. Leave blank if you want everything
LOG_SIZE = 5000                                      # Max size of log file in bytes
DOWNLOAD_PATH = f"{os.getcwd()}\\gmtj_downloads"     # A download folder to temporarily hold e-mails and attachments

DEBUG = False                                        # If True:  1) Joplin will not be called to sync 
#                                                                2) Unread mails will not be moved to trash.
#                                                                3) JOPLIN_INBOX is set to _debug         

"""
TALK TO JOPLIN
"""

def import_to_joplin(id, subject, text, attachments):
    """ Imports text and attachments to Joplin-notebook, with subject as title """
    
    try: 
        # Store text in file using its unique id
        path = f"gmtj_downloads\\{id}\\"
        if not os.path.exists(path):
            os.makedirs(path)
        
        text = text.encode()
        text_file = f"{path}{id}.md"
        with open(text_file, "wb") as f:
            f.write(text)

        # Import text as new note in joplin_inbox. Create inbox if it does not exist.
        path_to_text = f"{os.getcwd()}\\{text_file}"
        result = subprocess.run(f"joplin import \"{path_to_text}\" {JOPLIN_INBOX}", shell=True, capture_output=True)
        
        if str(result.stdout.decode()) == f"Cannot find \"{JOPLIN_INBOX}\".\n":
            logger.warning(f"Created target notebook \"{JOPLIN_INBOX}\". Restart of Joplin CLI may be required.")
            subprocess.run(f"joplin mkbook {JOPLIN_INBOX}", shell=True)
            subprocess.run(f"joplin import \"{path_to_text}\" {JOPLIN_INBOX}", shell=True)
        
        # Append attachments using list of attachment paths
        if attachments:
            for i in attachments:
                subprocess.run(f"joplin attach \"{id}\" \"{i}\"", shell=True)
        
        # Use GMail subject to set title
        subprocess.run(f"joplin set \"{id}\" title \"{subject}\"", shell=True)

        return True
    
    except Exception as e:
        logger.error(f"Something went wrong during import of {id}: {e}.")

        return False


""" 
TALK TO GMAIL
"""

def get_gmail_service():

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

    service = build("gmail", "v1", credentials=creds)

    return service


def check_gmail(service):
    """ Returns list of Gmail ids for any unread mail. """

    try:
        #service = build("gmail", "v1", credentials=creds)
        result = service.users().messages().list(userId="me", q="is:unread").execute()
        
        # Print number of e-mails
        n_new_mails = result["resultSizeEstimate"]
        print(f"{n_new_mails} unread mail(s)!")
        if n_new_mails == 0:
            return []

        # Get id's
        messages = result["messages"]
        ids = []
        for i in messages:
            ids.append(i["id"])
        
        logger.info(f"Found {len(ids)} unread mail(s)")

        # Return id's
        return ids
    
    except Exception as e:
        logger.error(f"Failed to retrieve Gmail. {e}")
        

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
            
            logger.warning(f"{sender} not in list of approved senders. E-mail {id} put in trash.")
            
            return False

    # Get text
    for part in mime_msg.walk():
        try:
            if part.get_content_type()=="text/plain":
                text = part.get_payload(decode=True)
                text = text.decode()
        
        except Exception as e:
            text_decode_error = f" # Decode-Error!\n\"{e}\".\n"
            problem_text = part.get_payload(decode=True)
            text = text_decode_error + problem_text.decode("utf-8", "replace")

            logger.error(f"Trouble decoding {id} from {sender} on {subject}. {e}")

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

    # Import to Joplin
    result = import_to_joplin(id, subject, text, att_filenames)
    
    # Thrash Gmail on successful import
    if not DEBUG:
        if result:
            service.users().messages().trash(userId="me", id=str(id)).execute()

    if result:
        logger.info(f"IMPORTED: {id} with {len(att_filenames)} attachments.\tFROM: {sender}.\t\tSUBJECT: \"{subject}\".")
        
        return True

    elif not result:
        logger.info(f"{id} was NOT imported, and remains unread in Gmail inbox.")
        
        return False


"""
RUN THE SCRIPT
"""
logging.basicConfig(
    handlers=[RotatingFileHandler("log.txt", maxBytes=LOG_SIZE, backupCount=0)],
    format="%(asctime)s - %(levelname)s -\t %(message)s", 
    datefmt="%m/%d/%Y %I:%M:%S %p")

logger = logging.getLogger("gmail-to-joplin")
logger.setLevel(logging.INFO)

if not DEBUG:
    logger.info("Running script ...")
 
if DEBUG:
    JOPLIN_INBOX = "_debug"
    logger.info("Script is running in DEBUG MODE ...")

# Check for credentials
if not os.path.exists("credentials.json"):
    logger.Error("Could not find OATH file \"credentials.json\" in script path.")
    quit()

# Prepare download path
if os.path.exists(DOWNLOAD_PATH):
    shutil.rmtree(DOWNLOAD_PATH)
    os.mkdir(DOWNLOAD_PATH)
else:
    os.mkdir(DOWNLOAD_PATH)

# Avoid false negatives when get_gmail() checks for approved senders
for i in range(len(APPROVED_SENDERS)):
    APPROVED_SENDERS[i] = APPROVED_SENDERS[i].lower()

# Check for new e-mails
new_gmails = check_gmail(get_gmail_service())

# Import e-mails
counter = 0
if new_gmails:
    for i in new_gmails:
        print(f"Importing {i} of {len(new_gmails)} mail(s).")
        get_it = import_gmail(i)
        if get_it:
            counter +=1

# Joplin sync
if new_gmails and not DEBUG:
    subprocess.run("joplin sync", shell=True)

if new_gmails:
    logger.info(f"Imported {counter} of {len(new_gmails)} unread mails.\n")
elif not new_gmails:
    logger.info(f"No new mail.")

# Clean up
shutil.rmtree(DOWNLOAD_PATH)
logger.info(f"Script finished.\n")
