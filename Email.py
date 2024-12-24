import imaplib
import email
import os
import pymongo
from email.header import decode_header
from email.utils import parseaddr
from bson import Binary

IMAP_SERVER = "outlook.office365.com"
IMAP_PORT = 993  
EMAIL_ACCOUNT = input('Enter your Email : ')
PASSWORD = input('Enter Your Password :  ')

MONGO_URI = "mongodb://localhost:27017/" 
DB_NAME = "email_db"
COLLECTION_NAME = "emails"

client = pymongo.MongoClient(MONGO_URI)
db = client[DB_NAME]
collection = db[COLLECTION_NAME]

ATTACHMENT_FOLDER = "attachments"

if not os.path.exists(ATTACHMENT_FOLDER):
    os.makedirs(ATTACHMENT_FOLDER)

def connect_to_mail():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(EMAIL_ACCOUNT, PASSWORD)
    return mail

def get_inbox(mail):
    mail.select("inbox")
    
    status, messages = mail.search(None, "ALL")
    
    email_ids = messages[0].split()
    return email_ids

def decode_subject(subject):
    decoded_subject, encoding = decode_header(subject)[0]
    if isinstance(decoded_subject, bytes):
        decoded_subject = decoded_subject.decode(encoding if encoding else 'utf-8')
    return decoded_subject

def save_attachment(part):
    filename = part.get_filename()
    if filename:
        file_content = part.get_payload(decode=True)
        return {"filename": filename, "file_content": Binary(file_content)}
    return None

def read_email(mail, email_id):
    status, msg_data = mail.fetch(email_id, "(RFC822)")
    
    email_data = {
        "subject": "",
        "from": "",
        "body": "",
        "attachments": []
    }

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            
            email_data["subject"] = decode_subject(msg["Subject"])
            email_data["from"] = parseaddr(msg.get("From"))[1]

            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    
                    if "attachment" in content_disposition:
                        attachment = save_attachment(part)
                        if attachment:
                            email_data["attachments"].append(attachment)  # Save attachment
                    elif content_type == "text/plain":
                        email_data["body"] = part.get_payload(decode=True).decode()
            else:
                email_data["body"] = msg.get_payload(decode=True).decode()


collection.insert_one(email_data)
print(f"Email with subject '{email_data['subject']}' stored in MongoDB.")

mail = connect_to_mail()
email_ids = get_inbox(mail)

for email_id in email_ids[-5:]:
    read_email(mail, email_id)

mail.logout()
