import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

# Google Sheets setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

google_service_account_info = st.secrets['google_service_account']
creds = ServiceAccountCredentials.from_json_keyfile_dict(google_service_account_info, scope)

# creds = ServiceAccountCredentials.from_json_keyfile_name("m2.json", scope)

client = gspread.authorize(creds)
sheet = client.open_by_key("16Ln8V-XTaSKDm1ycu5CNUkki-x2STgVvPHxSnOPKOwM")

# Fetching options and descriptions from the sheet
def fetch_options(sheet, tab_name):
    worksheet = sheet.worksheet(tab_name)
    data = worksheet.get_all_records()
    return data

# Update tab names based on your sheet
options_data = fetch_options(sheet, "Config")  # Ensure this tab exists
sections_data = fetch_options(sheet, "Config")  # Ensure this tab exists

if not options_data or not sections_data:
    st.error("Unable to fetch data from Google Sheets.")
    st.stop()

# Convert options data into a usable format
options = {}
for entry in options_data:
    # Use 'Sponsorship Type' as the key for options
    option_name = entry['Sponsorship Type']
    options[option_name] = {
        "points": entry['Points'],
        "max": entry['Max'],  # Assuming 'Max' value is available
        "uid": entry['UID'],  # Assuming 'UID' value is available
        "description": entry.get('Details', '').split(',')  # Use 'Details' instead of 'Description'
    }

# Convert sections data into a usable format
event_sections = {}
for entry in sections_data:
    section = entry['Event Name']
    option = entry['Sponsorship Type']
    if section not in event_sections:
        event_sections[section] = []
    event_sections[section].append(option)

# Streamlit form
st.title("Sponsorship Selection Form")

# Basic information inputs
name = st.text_input("Name")
company = st.text_input("Company")
email = st.text_input("Email")

# Total points input
total_points = st.number_input("Enter Total Points Allotted", min_value=0, value=0)

# Initialize session state for total_points and remaining_points
if 'total_points' not in st.session_state:
    st.session_state.total_points = total_points

if 'remaining_points' not in st.session_state:
    st.session_state.remaining_points = total_points

# Initialize session state for selected options
if 'selected_options' not in st.session_state:
    st.session_state.selected_options = {option: False for option in options.keys()}

# Function to update remaining points
def update_remaining_points():
    # Calculate total deducted points based on selected options
    deducted_points = sum(option_info['points'] for option, option_info in options.items() if st.session_state.selected_options[option])
    st.session_state.remaining_points = st.session_state.total_points - deducted_points

# Update session state values
st.session_state.total_points = total_points
update_remaining_points()

# Displaying sections and options
for section, section_options in event_sections.items():
    st.subheader(section)
    for option in section_options:
        if option in options:
            # Generate a unique key for each checkbox
            unique_key = f"{section}_{option}"
            option_info = options[option]
            points = option_info['points']
            max_range = option_info['max']
            description = option_info["description"]
            uid = option_info['uid']

            # Format the description into bullet points
            formatted_description = "\n".join([f"- {desc.strip()}" for desc in description if desc.strip()])

            # Display option checkbox and points deduction with max range and description
            st.session_state.selected_options[option] = st.checkbox(
                f"**{option}** - Points: {points}, Max: {max_range}",  # UID removed from the label
                value=st.session_state.selected_options[option],
                key=unique_key  # Ensure the key is unique
            )
            st.markdown(formatted_description)

    st.write("---")  # Divider line between sections
update_remaining_points()
# Display remaining points
st.write(f"### Remaining Points: {st.session_state.remaining_points}")

# Submit button
if st.button("Submit"):
    # Collect selected options and their UIDs
    selected_options = [option for option, selected in st.session_state.selected_options.items() if selected]
    selected_uids = [options[option]['uid'] for option in selected_options]

    # Prepare data to store in Google Sheets
    data = {
        "Name": name,
        "Company": company,
        "Email": email,
        "Total Points": st.session_state.total_points,
        "Remaining Points": st.session_state.remaining_points,
        "Selected Options": ", ".join([f"{option} (UID: {options[option]['uid']})" for option in selected_options]),
        "UID": ", ".join(selected_uids)
    }

    # Store data in Google Sheets
    sheet.worksheet("Sheet1").append_row([data["Name"], data["Company"], data["Email"], data["Total Points"], data["Remaining Points"], data["Selected Options"], data["UID"]])

    # Update Max value in Config sheet based on selected UIDs
    config_worksheet = sheet.worksheet("Config")
    config_data = config_worksheet.get_all_records()

    # Loop through config_data and find the matching UIDs
    for i, entry in enumerate(config_data):
        if entry['UID'] in selected_uids:
            # Subtract 1 from the Max field for the selected UID
            new_max = entry['Max'] - 1
            # Update the Max value in the sheet
            config_worksheet.update_cell(i + 2, config_worksheet.find('Max').col, new_max)

    # Reset form inputs by clearing session state values
    st.session_state.selected_options = {option: False for option in options.keys()}
    st.session_state.total_points = 0
    st.session_state.remaining_points = 0
    st.success("Form submitted successfully!")
