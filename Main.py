from __future__ import print_function
import os.path
import base64
import pandas as pd
from email.mime.text import MIMEText
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Google API Scopes
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/spreadsheets"
]

# Spreadsheet & Email Configs
SPREADSHEET_ID = "your_google_sheet_id"  # Replace with your actual Google Sheet ID
RANGE_NAME = "Sheet1!A2:C"  # Adjust based on your sheet structure
EMAIL_SENDER = "your_email@gmail.com"

# Authenticate and Create Services
def authenticate_services():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    sheets_service = build("sheets", "v4", credentials=creds)
    gmail_service = build("gmail", "v1", credentials=creds)
    return sheets_service, gmail_service

# Fetch contacts from Google Sheets
def get_contacts_from_sheets(service):
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get("values", [])
    contacts = []
    
    for row in values:
        if len(row) >= 2:
            name, email = row[:2]
            status = row[2] if len(row) > 2 else ""  # Check if an email was already sent
            contacts.append({"name": name, "email": email, "status": status})
    
    return contacts

# Create email content
def create_email(to, name):
    subject = f"Exciting Sponsorship Opportunity - {name}"
    body = f"""
    Dear {name},

    I hope you're doing well! I'm reaching out to explore potential sponsorship opportunities. 
    We believe that your company would be a great fit for our event.

    Let’s connect and discuss further!

    Best regards,
    Andy
    """
    message = MIMEText(body)
    message["to"] = to
    message["subject"] = subject
    return {"raw": base64.urlsafe_b64encode(message.as_bytes()).decode()}

# Send email using Gmail API
def send_email(service, to_email, name):
    message = create_email(to_email, name)
    service.users().messages().send(userId="me", body=message).execute()
    print(f"Email sent to {to_email}")

# Update Google Sheets with email sent status
def update_sheet(service, contacts):
    sheet = service.spreadsheets()
    updated_values = [[contact["status"]] for contact in contacts]
    update_range = f"Sheet1!C2:C{len(updated_values) + 1}"
    body = {"values": updated_values}
    
    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID, range=update_range,
        valueInputOption="RAW", body=body
    ).execute()
    print("Google Sheet updated.")

# Main function
def main():
    sheets_service, gmail_service = authenticate_services()
    
    contacts = get_contacts_from_sheets(sheets_service)
    for contact in contacts:
        if contact["status"] != "Sent":  # Check if the email was already sent
            send_email(gmail_service, contact["email"], contact["name"])
            contact["status"] = "Sent"

    update_sheet(sheets_service, contacts)

if __name__ == "__main__":
    main()
