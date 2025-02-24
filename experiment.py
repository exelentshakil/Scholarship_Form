import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import random
import string
import toml
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from streamlit.elements import image
from xhtml2pdf import pisa
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.utils import formataddr
from email import encoders
from xhtml2pdf.files import getFile, pisaFileObject
import io
import os
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from PIL import Image
import time

pc=st.set_page_config(page_title="Sustaining Sponsor Benefits",page_icon= ":clipboard:")#,page_icon= "logo.jpg")


# Add current date and time to the data
submission_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# Placeholder for loading animation
loading_placeholder = st.empty()

# Show a professional loading screen **only when data is not cached**
if 'gdrivesetup' not in st.session_state:
    with loading_placeholder.container():
        # Centering the logo using Streamlit columns
        col1, col2, col3 = st.columns([1, 2, 1])  # Adjust column width for centering
        with col2:
            st.image("logo.jpg", width=150)  # Ensure logo loads properly

        # Use Markdown to center text
        st.markdown(
            """
            <div style="text-align: center;">
                <h2>🔄 Preparing Your Sponsorship Experience</h2>
                <h4>Loading event sponsorship opportunities & exclusive benefits...</h4>
                <p style="color: gray;">This won't take long – making sure everything is ready for you! ⏳</p>
            </div>
            """,
            unsafe_allow_html=True
        )

        # Centering the spinner using columns
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            with st.spinner("Fetching sponsorship data, please wait..."):
                time.sleep(1.5)  # Small delay to show the spinner effect

# Load the TOML configuration from Streamlit secrets
secret_config = st.secrets["google_sheets"]
#g_secret_config = toml.load("secret.toml")
# Extract service account information
try:
    #secret_config = g_secret_config['google_sheets']
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
    UploadImagefolder =  secret_config["UploadImagefolder"]
except KeyError as e:
    st.error(f"Missing key in TOML file: {e}")
    st.stop()
# Add custom CSS to hide the GitHub icon
hide_github_icon = """ <style>
#GithubIcon, .stAppToolbar, [class^=_profileContainer_]  {
  visibility: hidden;
}
 </style>
"""
st.markdown(hide_github_icon, unsafe_allow_html=True)
def fetch_options(sheet, tab_name):
    try:
        worksheet = sheet.worksheet(tab_name)
        data = worksheet.get_all_records()
        return data,worksheet
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"Worksheet {tab_name} not found in Google Sheets.")
        st.stop()


if all( map(lambda l: l in list(st.session_state.keys()),['gdrivesetup']) ):
    scope,gauth,client,drive,sheet_id,sheet,options_data,sections_data,columns,worksheetSubmitted,worksheetConfig = st.session_state['gdrivesetup']
else:
    # Google Sheets setup
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    gauth = GoogleAuth()
    gauth.credentials = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
    client = gspread.authorize(gauth.credentials)
    sheet_id = secret_config["sheet_id"]#"16Ln8V-XTaSKDm1ycu5CNUkki-x2STgVvPHxSnOPKOwM"  # Replace with your actual Google Sheet ID
    sheet = client.open_by_key(sheet_id)
    drive = GoogleDrive(gauth)
    options_data,worksheetConfig = fetch_options(sheet, "Config")
    sections_data,worksheetConfig = fetch_options(sheet, "Config")
    worksheetSubmitted= sheet.worksheet("Submitted")
    columns = sheet.worksheet("Submitted").row_values(1)
    st.session_state['gdrivesetup'] = [scope,gauth,client,drive,sheet_id,sheet,options_data,sections_data,columns,worksheetSubmitted,worksheetConfig]

# **INSERT IT HERE**: Remove loading screen after setup is complete
loading_placeholder.empty()

def send_email(sender_email, sender_password, recipient_email, subject, body,file):
    try:
        for re in recipient_email.split(','):
            if len(re)==0:
                continue
            # Create a MIME object
            message = MIMEMultipart('alternative')
            message['From'] = 'Laurie Moher <'+sender_email+'>'
            message['To'] = re
            message['Subject'] = subject


            im = MIMEImage(open("logo.jpg", 'rb').read(),  name=os.path.basename("logo.jpg"))
            im.add_header('Content-ID', '<logo.jpg>')
            message.attach(im)

            uploadedsign = st.session_state['uploadedfile']
            if uploadedsign:
                file6 = drive.CreateFile({'id': uploadedsign})
                file6.GetContentFile('uploads/'+uploadedsign+'.jpg')
                im2 = MIMEImage(open('uploads/'+uploadedsign+'.jpg', 'rb').read(), name=os.path.basename("uploadedfile.jpg"))
                im2.add_header('Content-ID', '<uploadedfile.jpg>')
                message.attach(im2)
                os.remove('uploads/'+uploadedsign+'.jpg')

            # Attach the body to the message
            message.attach(MIMEText(body, 'html'))
            part = MIMEBase("application", "octet-stream")
            part.set_payload(file)
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition", 'attachment', filename=os.path.basename("Sustaining Sponsor Benefits.pdf")
            )
            message.attach(part)

            # Establish a secure session with Gmail's outgoing SMTP server
            with smtplib.SMTP('smtp.gmail.com', 587) as server:
                server.starttls()  # Secure the connection
                server.login(sender_email, sender_password)  # Log in with app password
                text = message.as_string()
                server.sendmail(sender_email, re, text)  # Send the email
                print("Email sent successfully!")

    except Exception as e:
        print(f"Failed to send email: {e}")


def getEmail(submittedData,tmp,isAll,isPdf):
    tmp2=tmp.replace("{"+str(45)+"}",str(datetime.now().year))
    template=''
    for i,t,v in submittedData:
        if i<=9:
            tmp2 =tmp2.replace("{"+str(i)+"}",str(v))
        else:
            t1=' - '.join(t.split(' - ')[:-1])
            t2=t.split(' - ')[-1]
            global options

            isSelectedstyle = "style='background:#eeffcc;'" if v.lower()!='no' else ''
            if len(isSelectedstyle)>0 or isAll:
                if isAll or isPdf:
                    template+="<tr "+isSelectedstyle+" ><td>"+t1+"</td><td>"+t2+"</td><td>"+', '.join(options[t]['description'])+"</td><td>"+str(options[t]['points'])+"</td><td>"+str(options[t]['max'] or '')+"</td><td>"+str(v)+"</td></tr>"
                else:
                    template+=f'<p  ><strong style="color:red;">{t1}: - <span style="color:black;"> {t2}</span></strong><br /> '+(', '.join(options[t]['description']))+f' '+str(options[t]['points'])+'/Points</p></br>'

    tmp2 =tmp2.replace("{table}",str(template))
    return tmp2

# Function to generate a random UID for each submission
def generate_random_uid():
    return "UID-" + ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))



if not options_data or not sections_data:
    st.error("Unable to fetch data from Google Sheets.")
    st.stop()

# Convert options data into a usable format
options = {}
for entry in options_data:
    option_name = entry['Computed Column']
    if entry['Status']!='Ongoing':
        continue

    try:
        points = int(entry['Points'])
    except ValueError:
        points = 0
    try:
        if str(entry['Max']).lower() == "n/a":
            max_range = "n/a"
        else:
            max_range = int(entry['Max'])
    except ValueError:
        max_range = 0
    try:
        max_month_selection = int(entry.get('Max Month Selection', 0))
    except ValueError:
        max_month_selection = 0

    options[option_name] = {
        "points": points,
        "max": max_range,
        "uid": entry['UID'],
        "description": entry.get('Details', '').split(','),
        "max_month_selection": max_month_selection,
        "computed_months_options": entry.get('Computed Months Options', '').split(','),
        'extra':entry
    }

# Convert sections data into a usable format
event_sections = {}
for entry in sections_data:
    section = entry['Event Name']
    option = entry['Sponsorship Type']
    if section not in event_sections:
        event_sections[section] = []
    event_sections[section].append(option)

# Function to handle months selection without modifying session_state directly
def handle_month_selection(unique_key, max_months, available_months):
    selected_months_key = f"{unique_key}_months"

    # Initialize the selected months in session state if not present
    if selected_months_key not in st.session_state:
        st.session_state[selected_months_key] = []
    ulabel=unique_key.replace("_"," ")
    # Create the multiselect widget
    selected_months = st.multiselect(
        f"Select the months you choose to sponsor for {ulabel}",
        available_months,
        #st.session_state[selected_months_key],
        key=selected_months_key,
        max_selections=max_months
    )

    # Enforce maximum selection by Streamlit (not manually)
    #if len(selected_months) > max_months:
    #    st.warning(f"You can select up to {max_months} months.")
    #    # Trim the selection to max_months
    #    selected_months = selected_months[:max_months]
    #    # Update the session state accordingly
    #    st.session_state[selected_months_key] = selected_months

    # No direct modification of st.session_state.remaining_points here
    # Points deduction will be handled separately

# Function to handle months selection without modifying session_state directly
def handle_month_selection2(unique_key,  available_months):
    max_months=1
    selected_months_key = f"{unique_key}_months"

    # Initialize the selected months in session state if not present
    if selected_months_key not in st.session_state:
        st.session_state[selected_months_key] = []
    ulabel=unique_key.replace("_"," ")
    # Create the multiselect widget
    selected_months = st.multiselect(
        f"Select the months you choose to sponsor for {ulabel}",
        available_months,
        #st.session_state[selected_months_key],
        key=selected_months_key,
    #    max_selections=max_months
    )


# Function to calculate remaining points
def calculate_remaining_points():
    deducted_points = 0
    for key, selected in st.session_state.selected_options.items():
        if selected:
            #option_name = key.split("_")[1]
            optionKey=key.replace('_',' - ')
            months_key = f"{key}_months"
            multiplier=1
            if months_key in st.session_state and (options[optionKey]['extra']['PointsDeductionMultiple'] or '').lower()!='no':#options[optionKey]['extra']['Multiples']=='Yes' and
                multiplier=len(st.session_state[months_key])

            deducted_points += (options[optionKey]['points']*multiplier)
            # Check if this option has associated months

            #if :
            #    deducted_points += len(st.session_state[months_key]) * 3  # Deduct 3 points per selected month
    st.session_state.remaining_points = st.session_state.total_points - deducted_points

# Streamlit form
st.image("logo.jpg", width=200)
st.title("Sustaining Sponsor Benefits")


#datainfo=[(i,c,st.session_state['temp_newpdf_data'][i] if 'temp_newpdf_data' in st.session_state else '') for i,c in enumerate(columns)]
#st.session_state.selected_options["checkboxlabels"][i]
#datainfo=[(i,c,st.session_state['temp_newpdf_data'][i] if 'temp_newpdf_data' in st.session_state and i<len(st.session_state['temp_newpdf_data']) else '' ) for i,c in enumerate(columns)]
datainfo=[(i,c,str(st.session_state['temp_newpdf_data'][i]) if 'temp_newpdf_data' in st.session_state and i<len(st.session_state['temp_newpdf_data']) else ''  ) for i,c in enumerate(columns)]#,options[c] if c in options else None
# Update Max value in Config sheet based on selected UIDs

config_data =     sections_data
pdfinfoEmpty = getEmail(datainfo, open("pdftemplate.tmp", "r").read().replace('{-1}',submission_date),True,True)
outputEmpty = io.BytesIO()
pisa.CreatePDF(pdfinfoEmpty,debug=1,
                     # page data
                    dest=outputEmpty, encoding='UTF-8'                                              # destination "file"
                )
docEmpty =outputEmpty.getbuffer().tobytes()
c1, c2 = st.columns([5,3])
c2.download_button('Print Form',docEmpty , file_name='Sustaining Sponsor Benefits.pdf', mime='application/pdf')
# Basic information inputs with clear labels
company = c1.text_input("Organization Name", help="Enter your organization's name.")
contact_name = st.text_input('Your Name :grey[(Who should be our regular contact when we need names for events, marketing materials, etc.)]')
#st.caption("Who should be our regular contact when we need names for events, marketing materials, etc.")

line_seperator=st.divider()
Contact_Name=st.text_input("Contact Name")
#Contact_Company=st.text_input("Contact Company ")
#Contact_Email=st.text_input("Contact Email")

email = st.text_input("Email", help="Enter a valid email address.")
phone_number = st.text_input('Phone Number', help="Enter your contact number.")

# Total points input
total_points = st.number_input(
    "SSP Points (Please use the number in the email)",
    min_value=0,
    max_value=100,
    value=0,
    help="Please enter a value between 0 and 100."
)

def init_session():
    # Initialize session state for total_points and remaining_points
    if 'total_points' not in st.session_state or st.session_state.total_points != total_points:
        st.session_state.total_points = total_points
        st.session_state.remaining_points = total_points

    # Initialize session state for selected options and months
    if 'selected_options' not in st.session_state:
        st.session_state.selected_options = {}
    if 'selected_months' not in st.session_state:
        st.session_state.selected_months = {}

    if 'Submited' not in st.session_state:
        st.session_state.Submited = False

    if 'uploadedfile' not in st.session_state:
        st.session_state['uploadedfile'] =""
        st.session_state['uploadedfiledetail'] ={ 'gid':'', 'gname':'','uname':''}

    #if 'checkboxlabels' not in st.session_state:
    #    st.session_state.selected_options["checkboxlabels"]={}

init_session()
# **New:** Calculate remaining points before rendering options
calculate_remaining_points()

uploadedfile  = st.file_uploader("please upload Logo image", accept_multiple_files=False, type= ['png', 'jpg'] )

# Displaying sections and options
for section, section_options in event_sections.items():
    st.subheader(section)
    for option in section_options:
        optionKey=(section+' - '+option)
        if  optionKey in options:
            unique_key = f"{section}_{option}"
            option_info = options[optionKey]
            points = option_info['points']
            max_range = option_info['max']
            description = option_info["description"]
            uid = option_info['uid']

            #if "WEBSITE:  2025 - Member Spotlight-Quarterly" in option_info['extra']['Computed Column']:
            #    ab=2
            # Clean and format the description into bullet points
            formatted_description = "\n".join([f"- {desc.strip()}" for desc in description if desc.strip()])

            # Initialize the session state for the unique key if it doesn't exist
            if unique_key not in st.session_state.selected_options:
                st.session_state.selected_options[unique_key] = False

            # Determine if the checkbox should be disabled
            disabled = False

            # Check for max selection constraints
            if isinstance(max_range, int) :
                current_selection_count = sum(
                    1 for key, selected in st.session_state.selected_options.items()
                    if selected and key.startswith(section)
                )
                if current_selection_count >= max_range and not st.session_state.selected_options[unique_key]:
                    disabled = True

            # **New Logic:** Disable if selecting this option would exceed remaining points
            if not st.session_state.selected_options[unique_key] and points > st.session_state.remaining_points:
                disabled = True

            # Display the checkbox and immediately update the session state based on the checkbox value
            checkbox_label = f"**{option}** - Points: {points}"
            if not (not option_info['extra']['Associated Subtitle']):
                # Modify the label for Luncheon sponsors
                checkbox_label += f" ({option_info['extra']['Associated Subtitle']})"

            if max_range!="n/a":
                checkbox_label += f", Max: {option_info['extra']['Default Max']} ({max_range} remaining)"

            selected = st.checkbox(
                checkbox_label,
                value=st.session_state.selected_options[unique_key],
                key=unique_key,
                disabled=disabled
            )
            #st.session_state.selected_options["checkboxlabels"][unique_key]=checkbox_label

            # Update session state based on the checkbox selection
            if selected != st.session_state.selected_options[unique_key]:
                st.session_state.selected_options[unique_key] = selected
                calculate_remaining_points()

            # Show the dropdown for months if the option is selected and it's a Luncheons option
            isSelectedOpt=unique_key in [so for so in st.session_state.selected_options if st.session_state.selected_options[so]]
            if isSelectedOpt and option_info['max_month_selection']>0 and len([a for a in option_info['computed_months_options'] if a])>0:#and "Luncheon" in option:
                max_months = option_info['max_month_selection']
                available_months = option_info['computed_months_options']
                handle_month_selection(unique_key, max_months, available_months)

            if isSelectedOpt and option_info['extra']['Multiples']=='Yes':
                available_months = option_info['extra']['Multiple Options'].split(',')
                handle_month_selection2(unique_key,  available_months)

            # Display the formatted description
            st.markdown(formatted_description)

    st.write("---")

# **Removed:** Existing call to calculate_remaining_points() after the loop
# calculate_remaining_points()

# Display remaining points in the sidebar only when total_points > 0
if st.session_state.total_points > 0:
    st.sidebar.title("Submission Details")
    st.sidebar.write(f"**Name:** {contact_name}")
    st.sidebar.write(f"**Email:** {email}")
    st.sidebar.write(f"**Organization:** {company}")
    st.sidebar.write(f"**Total Points Allotted:** {st.session_state.total_points}")
    st.sidebar.write(f"### Remaining Points: {st.session_state.remaining_points}")

pisaFileObject.getNamedFile = lambda self: self.uri

if st.session_state.Submited:
    st.success("Wait for Form submission!")


col1, col2, col3 = st.columns([1,3,4])
#if col1.button("PDF"):
#    pdfbtn = """ <a onclick="window.open('Test.pdf')" href="javascript:void(0)">Printable Version</a>
#    """
#    col1.markdown(pdfbtn, unsafe_allow_html=True)
st.session_state['temp_newpdf_data'] =['', contact_name, company, email, phone_number,st.session_state.total_points, st.session_state.remaining_points,Contact_Name,"",'']
# Submit button
if col1.button("Submit",disabled=(st.session_state.Submited)):
    if uploadedfile is not None and uploadedfile!=st.session_state['uploadedfiledetail']['uname']:
        fname = generate_random_uid()+"-"+uploadedfile.name
        bytes_data = uploadedfile.read()
        gfile = drive.CreateFile({"parents": [{'id': UploadImagefolder}], "title": fname, 'mimeType':"image/jpeg"})
        if uploadedfile.name.lower().endswith(".png"):
            with io.BytesIO() as f:

                Image.open(io.BytesIO(bytes_data), mode='r', formats=None).convert('RGB').save(f,format="JPEG")
                with open('uploads/'+fname, "wb") as binary_file:
                    binary_file.write(f.read())
            gfile.SetContentFile('uploads/'+fname)
        else:
            with open('uploads/'+fname, "wb") as binary_file:
                binary_file.write(bytes_data)
            gfile.SetContentFile('uploads/'+fname)
        try:
            gfile.Upload()
        finally:
            gfile.content.close()
        if gfile.uploaded:
            st.session_state['uploadedfile'] = gfile['id']#{ 'gid':gfile['id'], 'gname':fname,'uname':uploadedfile.name}
            st.session_state['uploadedfiledetail'] ={ 'gid':gfile['id'], 'gname':fname,'uname':uploadedfile.name}
            os.remove('uploads/'+fname)
            st.toast(f"Logo Upload Success.")

    selected_options = [key.replace("_"," - ") for key, selected in st.session_state.selected_options.items() if selected]
    selected_options_full = [key for key, selected in st.session_state.selected_options.items() if selected]
    selected_uids = [options[option]['uid'] for option in selected_options]

    # Collect selected months data
    selected_months_data = {
        key: " - ".join(months) for key, months in st.session_state.items() if months and key.endswith("_months")
    }


    # Generate a random UID for the submission
    submission_uid = generate_random_uid()

    # Prepare data to store in Google Sheets
    data = {
        "Name": contact_name,
        "Company": company,
        "Email": email,
        "phoneNumber": phone_number,
        "Total Points": st.session_state.total_points,
        "Contact Name":Contact_Name,
       # "Contact Company":Contact_Company,
       # "Contact Email":Contact_Email,
        "Remaining Points": st.session_state.remaining_points,
        "Selected Options": "; ".join([f"{option} (UID: {options[option]['uid']})" for option in selected_options]),
        "UID": ", ".join([f" {options[option]['uid']}" for option in selected_options]),
        "CustomerUID": submission_uid,
        "Selected Months": " | ".join([f"{key.split('_')[1]}: {months}" for key, months in selected_months_data.items()]),
        "Submission Date": submission_date
    }
    warnings = []
    for key,selected in st.session_state.selected_options.items():
        mms = options[key.replace("_"," - ")]["max_month_selection"]
        isMultiple= options[key.replace("_"," - ")]['extra']['PointsDeductionMultiple'].lower()!='no'
        curSel  = len((st.session_state[key+'_months'] if (key+'_months' in st.session_state) else []))
        if selected and ((curSel!=mms and not isMultiple) or (curSel<=0 and isMultiple)) and mms>0:
            warnings.append(key)
    #warnings = [key  ]#and len(st.session_state[key+'_months'])==0
    reqfields = { "Organization Name":len(company)>0,"Your Name":len(contact_name)>0,"Contact Name":len(Contact_Name)>0, "Email":len(email)>0,"Phone Number":len(phone_number )>0,"Total points":st.session_state.total_points>0}
    reqfieldserror = [f for f,fv in reqfields.items() if not fv]

    if not(len(reqfieldserror)==0  and st.session_state.remaining_points>-1):
        if len(reqfieldserror)>0:
            st.warning("Form Not Submited: " +", ".join(reqfieldserror)+" are Required")
        elif st.session_state.remaining_points<0:
            st.warning("Form Not Submited: Remaining points incorrect, please check your selection")
    elif len(warnings)>0:
        st.warning(', '.join([str(k.replace('_',' - ')) for k in warnings])+" field Not filled properly.")

    else:
        with st.spinner('Submitting...'):
            st.session_state.Submited =True

            # Store data in 'raw info' sheet
            sheet.worksheet("raw info").append_row([
                data["Name"], data["Company"], data["Email"], data["phoneNumber"], data["Total Points"],data["Contact Name"],
                data["Remaining Points"], data["Selected Options"], data["UID"], submission_uid, data["Selected Months"], data["Submission Date"],st.session_state['uploadedfile']
            ])

            # Store data in 'Submitted' sheet
            submission_data = [
                submission_uid, data["Name"], data["Company"], data["Email"], data["phoneNumber"],
                data["Total Points"], data["Remaining Points"],data["Contact Name"],"",st.session_state['uploadedfile'],#,data["Contact Company"],data["Contact Email"]
            ]

            # Dynamically add columns for each selected option with the formatted event and sponsorship details
            for section, section_options in event_sections.items():
                for option in section_options:
                    computed_column = f"{section} - {option}"
                    unique_key = f"{section}_{option}"
                    col_value = ""
                    if unique_key in selected_options_full:
                        col_value = "YES"
                        # Append selected months if applicable
                        selected_months = st.session_state.get(f"{unique_key}_months", [])
                        if selected_months:
                            col_value += " (" + ", ".join(selected_months) + ")"
                    else:
                        col_value = "NO"
                    submission_data.append(col_value)

            # Fetch email credentials from Streamlit secrets for security
            # sender_email = st.secrets["email"]["sender"]
            # app_password = st.secrets["email"]["app_password"]

            recipient_email = email
            subject = "Form Submitted Successfully"

            a=f"""
UID: {submission_uid}
Name: {contact_name}
Company: {company}
Email: {email}
Total Points: {st.session_state.total_points}
Remaining Points: {st.session_state.remaining_points}

Selected Options:
{", ".join(selected_options)}

Selected Months:
{", ".join([f"{key.split('_')[1]}: {months}" for key, months in selected_months_data.items()])}

Submission Date: {submission_date}
"""


            worksheetSubmitted.append_row(submission_data)

            datainfo=[(i,c,submission_data[i]) for i,c in enumerate(columns)]
            # Update Max value in Config sheet based on selected UIDs


            for i, entry in enumerate(sections_data):
                if entry['UID'] in selected_uids:
                    try:
                        if isinstance(entry['Max'], int):
                            new_max = entry['Max'] - 1
                            worksheetConfig.update_cell(i + 2, worksheetConfig.find('Max').col, new_max)
                    except ValueError:
                        pass

            pdfinfo = getEmail(datainfo, open("pdftemplate.tmp", "r").read().replace('{-1}',submission_date),False,True)
            output = io.BytesIO()
            pisa.CreatePDF(pdfinfo,debug=1,
                     # page data
                    dest=output, encoding='UTF-8'                                              # destination "file"
                )
            doc =output.getbuffer().tobytes()
            body = getEmail(datainfo, open("htmltemplate.tmp", "r").read().replace('{-1}',submission_date),False,False)
            send_email(secret_config["EmailSender"],secret_config["EmailPass"],secret_config["EmailRecieve"]+","+email,subject,body,doc)
            col2.download_button('Download PDF',doc , file_name='Sustaining Sponsor Benefits.pdf', mime='application/pdf')

        st.session_state.Submited =False
        st.success("Form submitted successfully!")