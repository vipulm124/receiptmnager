import base64
import re
import os
import io
from email import message_from_bytes

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

from PyPDF2 import PdfReader, PdfWriter
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()


SENDER_EMAIL = "service@iciciprulife.com"
SUBJECT = "Your renewal premium receipt"
PDF_PASSWORD = os.getenv("PDF_PASSWORD")
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID")

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive.file"]


def authenticate_google_services():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            # creds = flow.run_local_server(port=0)
            creds = flow.run_local_server(port=8080, open_browser=False)
        with open("token.json", "w") as token_file:
            token_file.write(creds.to_json())
    
    return creds


def get_matching_emails(service):
    results = service.users().messages().list(userId="me", q=f"from:{SENDER_EMAIL} has:attachment").execute()
    messages = results.get("messages", [])

    for msg in messages:
        msg_data = service.users().messages().get(userId='me', id=msg['id'], format='raw').execute()
        raw_msg = base64.urlsafe_b64decode(msg_data['raw'].encode('UTF-8'))
        email_msg = message_from_bytes(raw_msg)

        subject = str(email_msg['Subject'])
        if subject == SUBJECT:
            for part in email_msg.walk():
                if part.get_content_maintype() == 'application' and part.get_filename().endswith('.pdf'):
                    filename = part.get_filename()
                    file_data = part.get_payload(decode=True)
                    with open(filename, 'wb') as f:
                        f.write(file_data)
                    print(f"Downloaded: {filename}")
                    return filename
    print("No matching emails found.")
    return None


def remove_pdf_password(input_path, output_path, password):
    reader = PdfReader(input_path)
    if reader.is_encrypted:
        reader.decrypt(password)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        
        with open(output_path, 'wb') as out_file:
            writer.write(out_file)
        print(f"Password removed from: {output_path}")

    
def upload_to_drive(drive_service, file_path, folder_id):
    file_metadata = {
        'name': os.path.basename(file_path),
        'parents': [folder_id]
    }
    media = MediaIoBaseUpload(io.FileIO(file_path, 'rb'), mimetype='application/pdf')
    uploaded_file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f"Uploaded to Drive: {uploaded_file.get('id')}")

def main(input_file_path, is_file_from_email: bool = True):
    creds = authenticate_google_services()
    drive_service = build('drive', 'v3', credentials=creds)
    
    if is_file_from_email:
        gmail_service = build('gmail', 'v1', credentials=creds)

        downloaded_pdf_file = get_matching_emails(gmail_service)
        if not downloaded_pdf_file:
            print("No PDF file downloaded.")
            return
    else:
        downloaded_pdf_file = input_file_path

    decrypted_pdf_file = f"{datetime.now().strftime('%B%Y')}.pdf"
    remove_pdf_password(downloaded_pdf_file, decrypted_pdf_file, PDF_PASSWORD)
    upload_to_drive(drive_service, decrypted_pdf_file, DRIVE_FOLDER_ID)

if __name__ == "__main__":
    downloaded_pdf_file = "F5879300_T6324138_228083596.pdf"
    main(downloaded_pdf_file, False)
    print("Script completed successfully.")
