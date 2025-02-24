import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import random
import string
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from xhtml2pdf import pisa
import io
import os
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from PIL import Image

import logging


# Set Streamlit Page Config
st.set_page_config(page_title="Sustaining Sponsor Benefits", page_icon=":clipboard:")

# Add a loader while initializing the app
with st.spinner("Loading application..."):
    submission_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    secret_config = st.secrets["google_sheets"]

    try:
        service_account_info = {
            "type": secret_config["type"],
            "project_id": secret_config["project_id"],
            "private_key_id": secret_config["private_key_id"],
            "private_key": secret_config["private_key"].replace("\\n", "\n"),
            "client_email": secret_config["client_email"],
            "client_id": secret_config["client_id"],
            "auth_uri": secret_config["auth_uri"],
            "token_uri": secret_config["token_uri"],
            "auth_provider_x509_cert_url": secret_config["auth_provider_x509_cert_url"],
            "client_x509_cert_url": secret_config["client_x509_cert_url"]
        }
        UploadImagefolder = secret_config["UploadImagefolder"]
    except KeyError as e:
        st.error(f"Missing key in secrets file: {e}")
        st.stop()

    @st.cache_resource
    def get_google_client():
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        gauth = GoogleAuth()
        gauth.credentials = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
        client = gspread.authorize(gauth.credentials)
        drive = GoogleDrive(gauth)
        return client, drive

    client, drive = get_google_client()
    sheet_id = secret_config["sheet_id"]
    sheet = client.open_by_key(sheet_id)
    worksheetSubmitted = sheet.worksheet("Submitted")
    worksheetConfig = sheet.worksheet("Config")

    @st.cache_data
    def fetch_options(tab_name):
        try:
            worksheet = sheet.worksheet(tab_name)
            return worksheet.get_all_records()
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"Worksheet {tab_name} not found in Google Sheets.")
            st.stop()

    options_data = fetch_options("Config")
    sections_data = fetch_options("Config")

def generate_random_uid():
    return "UID-" + ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))

st.image("logo.jpg", width=200)
st.title("Sustaining Sponsor Benefits")

company = st.text_input("Organization Name", help="Enter your organization's name.")
contact_name = st.text_input("Your Name", help="Main contact for marketing & event updates.")
email = st.text_input("Email", help="Enter a valid email address.")
phone_number = st.text_input("Phone Number", help="Enter your contact number.")

total_points = st.number_input("SSP Points", min_value=0, max_value=100, value=0, help="Enter a value between 0 and 100.")

uploadedfile = st.file_uploader("Upload Logo", type=['png', 'jpg'])

def upload_image_to_drive(uploadedfile):
    fname = generate_random_uid() + "-" + uploadedfile.name
    bytes_data = uploadedfile.read()
    gfile = drive.CreateFile({"parents": [{'id': UploadImagefolder}], "title": fname, 'mimeType': "image/jpeg"})
    
    with io.BytesIO() as f:
        Image.open(io.BytesIO(bytes_data)).convert('RGB').save(f, format="JPEG")
        with open(f'uploads/{fname}', "wb") as binary_file:
            binary_file.write(f.read())
    
    gfile.SetContentFile(f'uploads/{fname}')
    gfile.Upload()
    os.remove(f'uploads/{fname}')
    
    return gfile['id']

if uploadedfile and "uploadedfile" not in st.session_state:
    with st.spinner("Uploading logo..."):
        st.session_state["uploadedfile"] = upload_image_to_drive(uploadedfile)
        st.success("Logo uploaded successfully!")

st.sidebar.title("Submission Details")
st.sidebar.write(f"**Name:** {contact_name}")
st.sidebar.write(f"**Email:** {email}")
st.sidebar.write(f"**Organization:** {company}")
st.sidebar.write(f"**Total Points:** {total_points}")

if st.button("Submit"):
    if not all([company, contact_name, email, phone_number, total_points]):
        st.warning("All fields are required!")
    else:
        with st.spinner("Submitting form..."):
            submission_uid = generate_random_uid()
            data = {
                "Name": contact_name,
                "Company": company,
                "Email": email,
                "Phone": phone_number,
                "Total Points": total_points,
                "Submission Date": submission_date
            }

            worksheetSubmitted.append_row([
                data["Name"], data["Company"], data["Email"], data["Phone"],
                data["Total Points"], submission_uid, submission_date
            ])
            
            output = io.BytesIO()
            pdf_content = f"""
            Name: {contact_name}
            Company: {company}
            Email: {email}
            Phone: {phone_number}
            Total Points: {total_points}
            Submission Date: {submission_date}
            """
            pisa.CreatePDF(pdf_content, dest=output)
            pdf_data = output.getvalue()

            def send_email(recipient_email, subject, body, pdf_data):
                try:
                    message = MIMEMultipart()
                    message['From'] = 'Your Name <your-email@example.com>'
                    message['To'] = recipient_email
                    message['Subject'] = subject
                    message.attach(MIMEText(body, 'html'))
                    
                    pdf_part = MIMEText(pdf_data, 'base64')
                    pdf_part.add_header("Content-Disposition", "attachment", filename="Sustaining Sponsor Benefits.pdf")
                    message.attach(pdf_part)

                    with smtplib.SMTP('smtp.gmail.com', 587) as server:
                        server.starttls()
                        server.login(secret_config["EmailSender"], secret_config["EmailPass"])
                        server.sendmail(secret_config["EmailSender"], recipient_email, message.as_string())
                    
                    st.success("Form submitted & email sent successfully!")
                except Exception as e:
                    st.error(f"Email failed: {e}")

            send_email(email, "Form Submitted", "Your submission was successful!", pdf_data)
            st.download_button('Download PDF', pdf_data, file_name='Sustaining_Sponsor_Benefits.pdf', mime='application/pdf')


# Configure logging
LOG_FILENAME = "streamlit_app.log"
logging.basicConfig(
    filename=LOG_FILENAME,
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
    filemode="w",  # Overwrites logs each run; change to "a" to append
)

logging.info("Streamlit App Started.")

# Function to fetch data from Google Sheets
def fetch_options(sheet, tab_name):
    try:
        logging.info(f"Fetching data from Google Sheet: {tab_name}")
        worksheet = sheet.worksheet(tab_name)
        data = worksheet.get_all_records()

        if not data:
            logging.warning(f"Sheet {tab_name} is empty!")

        logging.info(f"Successfully fetched {len(data)} records from {tab_name}")
        return data, worksheet
    except gspread.exceptions.WorksheetNotFound as e:
        logging.error(f"Worksheet {tab_name} not found: {e}")
        st.error(f"Worksheet {tab_name} not found.")
        st.stop()
    except Exception as e:
        logging.error(f"Unexpected error fetching {tab_name}: {e}")
        st.error(f"Error fetching data from {tab_name}: {e}")
        st.stop()

# Session State Debugging
logging.info("Checking session state variables...")
if "gdrivesetup" in st.session_state:
    logging.info("Session state found, using cached setup.")
    scope, gauth, client, drive, sheet_id, sheet, options_data, sections_data, columns, worksheetSubmitted, worksheetConfig = st.session_state["gdrivesetup"]
else:
    logging.info("No session state found, initializing Google Sheets setup.")

    # Google Sheets Setup
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    gauth = GoogleAuth()
    gauth.credentials = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
    client = gspread.authorize(gauth.credentials)

    sheet_id = secret_config["sheet_id"]
    sheet = client.open_by_key(sheet_id)
    drive = GoogleDrive(gauth)

    # Fetch Data with Logging
    options_data, worksheetConfig = fetch_options(sheet, "Config")
    sections_data, worksheetConfig = fetch_options(sheet, "Config")

    # Log the number of entries fetched
    logging.info(f"Options data fetched: {len(options_data)} records")
    logging.info(f"Sections data fetched: {len(sections_data)} records")

    worksheetSubmitted = sheet.worksheet("Submitted")
    columns = sheet.worksheet("Submitted").row_values(1)
    logging.info(f"Column headers retrieved: {columns}")

    st.session_state["gdrivesetup"] = [scope, gauth, client, drive, sheet_id, sheet, options_data, sections_data, columns, worksheetSubmitted, worksheetConfig]

# Display logs in Streamlit UI (for debugging)
if st.button("Show Logs"):
    with open(LOG_FILENAME, "r") as log_file:
        st.text(log_file.read())

# Debugging Google Sheets Data
if not options_data:
    logging.error("options_data is empty!")
    st.error("Options data is empty. Check the Google Sheet.")
if not sections_data:
    logging.error("sections_data is empty!")
    st.error("Sections data is empty. Check the Google Sheet.")

logging.info("Google Sheets data successfully loaded into session state.")