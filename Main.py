from __future__ import print_function
import os.path
import base64
import pandas as pd
import concurrent.futures
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import openai
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

# Google API Scopes
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/spreadsheets"
]

# Spreadsheet & Email Configs
SPREADSHEET_ID = "1H106oroMpzSnFPaKOp1m5jiCWP9C1vfSs14STOG1G3o"  # Replace with your actual Google Sheet ID
RANGE_NAME = "Sheet1!A:F"  # Updated to include all columns
EMAIL_SENDER = "info@neoleague.dev"

HTML_EMAIL_SIGNATURE = """
<table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
    <tbody>
        <tr>
            <td>
                <table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
                    <tbody>
                        <tr>
                            <td style="vertical-align: top;">
                                <table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
                                    <tbody>
                                        <tr>
                                            <td class="template1__ImageContainer-sc-nmby7a-0 byHTIl" style="text-align: center;">
                                                <img src="https://i.imgur.com/ennLYRF.png" role="presentation" width="130" class="image__StyledImage-sc-hupvqm-0 fOIYAq" style="display: block; max-width: 130px;">
                                            </td>
                                        </tr>
                                        <tr>
                                            <td height="30"></td>
                                        </tr>
                                        <tr>
                                            <td style="text-align: center;">
                                                <table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="display: inline-block; vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
                                                    <tbody>
                                                        <tr style="text-align: center;">
                                                            <td>
                                                                <a href="https://x.com/NeoDevLeague" color="#065f46" class="social-links__LinkAnchor-sc-py8uhj-2 iWFMIm" style="display: inline-block; padding: 0px; background-color: rgb(6, 95, 70);">
                                                                    <img src="https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/twitter-icon-2x.png" alt="twitter" color="#065f46" width="24" class="social-links__LinkImage-sc-py8uhj-1 cgBQhD" style="background-color: rgb(6, 95, 70); max-width: 135px; display: block;">
                                                                </a>
                                                            </td>
                                                            <td width="5">
                                                                <div></div>
                                                            </td>
                                                            <td>
                                                                <a href="https://www.linkedin.com/company/neo-developer-league" color="#065f46" class="social-links__LinkAnchor-sc-py8uhj-2 iWFMIm" style="display: inline-block; padding: 0px; background-color: rgb(6, 95, 70);">
                                                                    <img src="https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/linkedin-icon-2x.png" alt="linkedin" color="#065f46" width="24" class="social-links__LinkImage-sc-py8uhj-1 cgBQhD" style="background-color: rgb(6, 95, 70); max-width: 135px; display: block;">
                                                                </a>
                                                            </td>
                                                            <td width="5">
                                                                <div></div>
                                                            </td>
                                                            <td>
                                                                <a href="https://www.instagram.com/neodevleague" color="#065f46" class="social-links__LinkAnchor-sc-py8uhj-2 iWFMIm" style="display: inline-block; padding: 0px; background-color: rgb(6, 95, 70);">
                                                                    <img src="https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/instagram-icon-2x.png" alt="instagram" color="#065f46" width="24" class="social-links__LinkImage-sc-py8uhj-1 cgBQhD" style="background-color: rgb(6, 95, 70); max-width: 135px; display: block;">
                                                                </a>
                                                            </td>
                                                            <td width="5">
                                                                <div></div>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>
                            </td>
                            <td width="46">
                                <div></div>
                            </td>
                            <td style="padding: 0px; vertical-align: middle;">
                                <h2 color="#000000" class="name__NameContainer-sc-1m457h3-0 csBPEs" style="margin: 0px; font-size: 18px; color: rgb(0, 0, 0); font-weight: 600;">
                                    <span>{name}</span>
                                </h2>
                                <p color="#000000" font-size="medium" class="company-details__CompanyContainer-sc-j5pyy8-0 cSOAsl" style="margin: 0px; font-weight: 500; color: rgb(0, 0, 0); font-size: 14px; line-height: 22px;">
                                    <span>Neo Developer League</span>
                                </p>
                                <table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="width: 100%; vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
                                    <tbody>
                                        <tr>
                                            <td height="30"></td>
                                        </tr>
                                        <tr>
                                            <td color="#065f46" direction="horizontal" width="auto" height="1" class="color-divider__Divider-sc-1h38qjv-0 bofWVx" style="width: 100%; border-bottom: 1px solid rgb(6, 95, 70); border-left: none; display: block;"></td>
                                        </tr>
                                        <tr>
                                            <td height="30"></td>
                                        </tr>
                                    </tbody>
                                </table>
                                <table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
                                    <tbody>
                                        <tr height="25" style="vertical-align: middle;">
                                            <td width="30" style="vertical-align: middle;">
                                                <table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
                                                    <tbody>
                                                        <tr>
                                                            <td style="vertical-align: bottom;">
                                                                <span color="#065f46" width="11" class="contact-info__IconWrapper-sc-mmkjr6-1 ldYaqt" style="display: inline-block; background-color: rgb(6, 95, 70);">
                                                                    <img src="https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/email-icon-2x.png" color="#065f46" alt="emailAddress" width="13" class="contact-info__ContactLabelIcon-sc-mmkjr6-0 gxFfYp" style="display: block; background-color: rgb(6, 95, 70);">
                                                                </span>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </td>
                                            <td style="padding: 0px;">
                                                <a href="mailto:info@neoleague.dev" color="#000000" class="contact-info__ExternalLink-sc-mmkjr6-2 jOTYAn" style="text-decoration: none; color: rgb(0, 0, 0); font-size: 12px;">
                                                    <span>info@neoleague.dev</span>
                                                </a>
                                            </td>
                                        </tr>
                                        <tr height="25" style="vertical-align: middle;">
                                            <td width="30" style="vertical-align: middle;">
                                                <table cellpadding="0" cellspacing="0" border="0" globalstyles="[object Object]" class="table__StyledTable-sc-1avdl6r-0 lmHSv" style="vertical-align: -webkit-baseline-middle; font-size: medium; font-family: Arial;">
                                                    <tbody>
                                                        <tr>
                                                            <td style="vertical-align: bottom;">
                                                                <span color="#065f46" width="11" class="contact-info__IconWrapper-sc-mmkjr6-1 ldYaqt" style="display: inline-block; background-color: rgb(6, 95, 70);">
                                                                    <img src="https://cdn2.hubspot.net/hubfs/53/tools/email-signature-generator/icons/link-icon-2x.png" color="#065f46" alt="website" width="13" class="contact-info__ContactLabelIcon-sc-mmkjr6-0 gxFfYp" style="display: block; background-color: rgb(6, 95, 70);">
                                                                </span>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </td>
                                            <td style="padding: 0px;">
                                                <a href="https://neoleague.dev" color="#000000" class="contact-info__ExternalLink-sc-mmkjr6-2 jOTYAn" style="text-decoration: none; color: rgb(0, 0, 0); font-size: 12px;">
                                                    <span>neoleague.dev</span>
                                                </a>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </td>
        </tr>
    </tbody>
</table>
"""

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
    values = result.get("values", [])[1:]  # Skip header row
    contacts = []
    
    for row in values:
        # Adjust indexing based on your sheet structure
        selected = row[0] if len(row) > 0 else ""
        company = row[1] if len(row) > 1 else "NeoDev League"
        name = row[2] if len(row) > 2 else ""
        email = row[3] if len(row) > 3 else ""
        
        if selected == "Selected":
            contacts.append({
                "name": name, 
                "email": email, 
                "company": company
            })
    
    return contacts

# Generate company-specific content using ChatGPT
def generate_text(company, model="gpt-4o-mini", max_tokens=150):
    messages = [{
        "role": "system", 
        "content": ("You are a member of a tech startup called the Neo Developer League and are explaining "
                    "to a company why your company's values align with theirs and be specific. The general idea "
                    "of the Neo Developer League is: the Neo Developer League is a student-led organization that "
                    "hosts competitive events created to inspire high school students to pursue engineering and "
                    "build connections in a fun and competitive way. Also limit answer to TWO sentences and say it "
                    "like you just finished explaining the company. Also speak as a team (use 'we)")
    }]
    messages.append({"role": "user", "content": f"Tell me why your company aligns with the values at {company}"})

    response = openai.chat.completions.create(
        model=model,
        messages=messages,
        max_tokens=max_tokens,
        temperature=0.7,
        n=1,
        stop=None,
        stream=False,
        frequency_penalty=2.0
    )
    return response.choices[0].message.content

# Create email content with optional PDF attachment
def create_email(to, name, company, sender_name, pdf_path=None):
    subject = f"Exciting Sponsorship Opportunity - {name}"
    company_statement = generate_text(company)
    formatted_signature = HTML_EMAIL_SIGNATURE.format(name=sender_name)
    
    # Plain text version remains the same
    body = f"""
Dear {name},

Hello! I'm {sender_name}, a co-founder of NeoDev League, and we're excited to connect with you for an upcoming partnership opportunity.

The NeoDev League is a non-profit dedicated to building engineering passion and creativity in high school students through dynamic, team-based competitions. We recently hosted our debut event in Waterloo, where 100 students from various schools collaborated, innovated, and competed at the Den 1880, leading to an amazing event.

Now, as we prepare to bring the NeoDev League to Toronto this May, we're gearing up to accommodate 300 talented participants and expand our impact. With this growth, we're seeking partners like you who share our vision of inspiring the next generation of engineers. Your support would play a pivotal role in empowering these young minds and enable us to provide an even more impactful and lasting experience.

Our events unite teams of 8-10 students for a full day of collaboration, creativity, and competition, culminating in project presentations to a panel of judges. {company_statement} We're confident that our collaboration will not only provide exceptional visibility and engagement opportunities for your brand but also strengthen the tech community surrounding us.

We offer several sponsorship tiers outlined in the attached sponsorship package. Sponsoring the NeoDev League will allow your company to enjoy brand exposure, interact with emerging talent, and display leadership in supporting youth initiatives.

Thank you for considering this opportunity. We would love to discuss this further and explore how we can make a difference together. Please feel free to reach out with any questions.

Best regards, {sender_name}
    """

    # Create HTML version with proper formatting
    html_body = f"""
    <div style="font-family: Arial, sans-serif; line-height: 1.5;">
        <p>Dear {name},</p>
        
        <p>Hello! I'm {sender_name}, a co-founder of NeoDev League, and we're excited to connect with you for an upcoming partnership opportunity.</p>
        
        <p>The NeoDev League is a non-profit dedicated to building engineering passion and creativity in high school students through dynamic, team-based competitions. We recently hosted our debut event in Waterloo, where 100 students from various schools collaborated, innovated, and competed at the Den 1880, leading to an amazing event.</p>
        
        <p>Now, as we prepare to bring the NeoDev League to Toronto this May, we're gearing up to accommodate 300 talented participants and expand our impact. With this growth, we're seeking partners like you who share our vision of inspiring the next generation of engineers. Your support would play a pivotal role in empowering these young minds and enable us to provide an even more impactful and lasting experience.</p>
        
        <p>Our events unite teams of 8-10 students for a full day of collaboration, creativity, and competition, culminating in project presentations to a panel of judges. {company_statement} We're confident that our collaboration will not only provide exceptional visibility and engagement opportunities for your brand but also strengthen the tech community surrounding us.</p>
        
        <p>We offer several sponsorship tiers outlined in the attached sponsorship package. Sponsoring the NeoDev League will allow your company to enjoy brand exposure, interact with emerging talent, and display leadership in supporting youth initiatives.</p>
        
        <p>Thank you for considering this opportunity. We would love to discuss this further and explore how we can make a difference together. Please feel free to reach out with any questions.</p>
        
        <p>Best regards,<br>{sender_name}</p>
    </div>
    """

    message = MIMEMultipart('mixed')
    message["to"] = to
    message["cc"] = "admin@neoleague.dev"
    message["subject"] = subject

    # Plain text version
    #plain_text_body = MIMEText(body, 'plain')
    #message.attach(plain_text_body)

    # HTML version with signature
    full_html = f"{html_body}{formatted_signature}"
    html_content = MIMEText(full_html, 'html')
    message.attach(html_content)

    # Attach PDF if provided and exists
    if pdf_path and os.path.exists(pdf_path):
        with open(pdf_path, "rb") as pdf_file:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(pdf_file.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(pdf_path)}"')
        message.attach(part)
        return {"raw": base64.urlsafe_b64encode(message.as_bytes()).decode()}
    # Send email using Gmail API with optional PDF attachment

def send_email(service, to_email, name, company, sender_name, pdf_path=None):
    message = create_email(to_email, name, company, sender_name, pdf_path)
    service.users().messages().send(userId="me", body=message).execute()
    print(f"Email sent to {to_email} with attachment: {pdf_path if pdf_path else 'No attachment'}")

# Update Google Sheets with email sent status
def update_sheet(service, sheet_id, sender_name):
    sheet = service.spreadsheets()
    
    # Get the data again to find rows to update
    result = sheet.values().get(spreadsheetId=sheet_id, range=RANGE_NAME).execute()
    values = result.get("values", [])
    
    # Prepare batch update
    update_requests = []
    for i, row in enumerate(values[1:], start=2):  # Start from 2nd row (index 1) because of header
        if row[0] == "Selected":
            # Create an update request for the 'Selected' column, 'Date sent' column, and 'Person' column
            update_requests.append({
                "range": f"Sheet1!A{i}:F{i}",
                "values": [["Sent", row[1], row[2], row[3], pd.Timestamp.now().strftime('%Y-%m-%d'), sender_name]]
            })

    # Perform batch update
    if update_requests:
        batch_update_values_request_body = {
            'valueInputOption': 'USER_ENTERED',
            'data': update_requests
        }
        
        sheet.values().batchUpdate(
            spreadsheetId=sheet_id, 
            body=batch_update_values_request_body
        ).execute()
        print("Google Sheet updated.")

# Main function using concurrency to send emails in batches of 8
def main():
    # Prompt for sender's name
    sender_name = input("Enter your name (First Name & Last Name) (to be used in the email): ").strip()
    
    sheets_service, gmail_service = authenticate_services()
    pdf_path = "NeoDevProspectus.pdf"  # Change to your actual PDF path if needed
    contacts = get_contacts_from_sheets(sheets_service)
    
    # Use ThreadPoolExecutor to send up to 8 emails concurrently
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
        futures = []
        for contact in contacts:
            future = executor.submit(
                send_email,
                gmail_service,
                contact["email"],
                contact["name"],
                contact["company"],
                sender_name,
                pdf_path
            )
            futures.append(future)
        
        # Wait for all the email sending tasks to complete
        concurrent.futures.wait(futures)
    
    # Update sheet with sent status and current date
    update_sheet(sheets_service, SPREADSHEET_ID, sender_name)

if __name__ == "__main__":
    main()