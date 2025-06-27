import streamlit as st
import pandas as pd
import json
import smtplib
from email.message import EmailMessage
from datetime import datetime, timedelta
import uuid
#import pywhatkit as kit

# Set page to wide mode
st.set_page_config(layout="wide")

# Load data
candidates_df = pd.read_excel("candidates.xlsx")
panels_df = pd.read_excel("panel.xlsx")

# Initialize session state
if 'groups' not in st.session_state:
    st.session_state.groups = {}
if 'panels' not in st.session_state:
    st.session_state.panels = {}

# Function to generate Google Meet link
def generate_meet_link():
    # Generate a unique meeting ID
    meeting_id = str(uuid.uuid4())[:8]
    return f"https://meet.google.com/{meeting_id}"

# Configuration section
st.sidebar.title("Configuration Settings")

# Load settings from JSON file
def load_settings():
    try:
        with open("settings.json", "r") as file:
            return json.load(file)
    except FileNotFoundError:
        return {}

# Save settings to JSON file
def save_settings(settings):
    with open("settings.json", "w") as file:
        json.dump(settings, file)

# Load stored settings
settings = load_settings()

# Gmail Configuration in an expander
with st.sidebar.expander("📧 Gmail Configuration", expanded=False):
    st.markdown("""
    ### Gmail Setup Instructions
    1. Enable 2-Step Verification in your Google Account
    2. Generate an App Password for this application
    3. Use the App Password below instead of your regular password
    """)
    gmail_email = st.text_input("Your Gmail Address", settings.get("gmail_email", ""))
    gmail_password = st.text_input("Your Gmail App Password", settings.get("gmail_password", ""), type="password")
    smtp_server = st.text_input("SMTP Server", settings.get("smtp_server", "smtp.gmail.com"))
    smtp_port = st.number_input("SMTP Port", settings.get("smtp_port", 587))

    # Save Configuration
    if st.button("Save Gmail Settings"):
        settings.update({
            "gmail_email": gmail_email,
            "gmail_password": gmail_password,
            "smtp_server": smtp_server,
            "smtp_port": smtp_port
        })
        save_settings(settings)
        st.success("Gmail settings saved successfully!")

# WhatsApp Configuration in an expander
with st.sidebar.expander("📱 WhatsApp Configuration", expanded=False):
    whatsapp_number = st.text_input("Your WhatsApp Number", settings.get("whatsapp_number", ""))
    whatsapp_message = st.text_area("Default Message Template", settings.get("whatsapp_message", "Your interview details have been sent via email."))
    
    if st.button("Save WhatsApp Settings"):
        settings.update({
            "whatsapp_number": whatsapp_number,
            "whatsapp_message": whatsapp_message
        })
        save_settings(settings)
        st.success("WhatsApp settings saved successfully!")

# Function to send email with enhanced formatting
def send_email(to, subject, body, is_html=False):
    if not gmail_email or not gmail_password:
        return "Error: Please configure your Gmail settings in the sidebar first."

    email = EmailMessage()
    email["From"] = gmail_email
    email["To"] = to
    email["Subject"] = subject
    
    if is_html:
        email.add_alternative(body, subtype='html')
    else:
        email.set_content(body)

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(gmail_email, gmail_password)
            server.send_message(email)
        return "Email sent successfully!"
    except smtplib.SMTPAuthenticationError:
        return """Gmail Authentication Error: Please follow these steps:
        1. Enable 2-Step Verification in your Google Account
        2. Generate an App Password:
           - Go to Google Account Settings
           - Security → App Passwords
           - Select 'Mail' and 'Other'
           - Copy the generated 16-character password
        3. Use this App Password in the sidebar settings
        
        For detailed instructions, visit: https://support.google.com/accounts/answer/185833"""
    except Exception as e:
        return f"Error sending email: {str(e)}"

# Function to send WhatsApp message
def send_whatsapp(phone, message):
    try:
        kit.sendwhatmsg_instantly(phone, message)
        return "WhatsApp message sent successfully!"
    except Exception as e:
        return f"Error sending WhatsApp message: {e}"

# Streamlit Main UI
st.title("Interview Management System")

# Tabs for different functionalities
tab1, tab2, tab3, tab4 = st.tabs(["Candidate Groups", "Panel Management", "Schedule Interviews", "Send Custom Message"])

# Candidate Groups Tab
with tab1:
    st.header("Candidate Groups Management")
    
    # Create new group
    col1, col2 = st.columns([2, 1])
    with col1:
        new_group_name = st.text_input("Enter New Group Name")
    with col2:
        if st.button("Create Group"):
            if new_group_name:
                if new_group_name not in st.session_state.groups:
                    st.session_state.groups[new_group_name] = []
                    st.success(f"Group '{new_group_name}' created successfully!")
            else:
                st.error("Please enter a group name")

    # Select group to manage
    if st.session_state.groups:
        selected_group = st.selectbox("Select Group to Manage", list(st.session_state.groups.keys()))
        
        # Enhanced candidate selection
        st.subheader("Add Candidates to Group")
        
        # Initialize selection state if not exists
        if 'candidate_selection' not in st.session_state:
            st.session_state.candidate_selection = pd.DataFrame(candidates_df)
            st.session_state.candidate_selection['Selected'] = False
        
        # Selection options in columns
        col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
        with col1:
            range_input = st.text_input("Enter range (e.g., 1-5) or specific numbers (e.g., 1,3,5)")
        with col2:
            if st.button("Select Range"):
                if range_input:
                    try:
                        if '-' in range_input:
                            start, end = map(int, range_input.split('-'))
                            indices = range(start-1, end)
                        else:
                            indices = [int(i)-1 for i in range_input.split(',')]
                        st.session_state.candidate_selection.loc[indices, 'Selected'] = True
                    except:
                        st.error("Invalid range format")
        with col3:
            if st.button("Select All"):
                st.session_state.candidate_selection['Selected'] = True
        with col4:
            if st.button("Clear Selection"):
                st.session_state.candidate_selection['Selected'] = False
        
        # Display candidate table with individual checkboxes
        st.dataframe(
            st.session_state.candidate_selection,
            column_config={
                "Selected": st.column_config.CheckboxColumn(
                    "Select",
                    help="Select candidate",
                    default=False
                ),
                "Name": st.column_config.TextColumn("Name"),
                "Email": st.column_config.TextColumn("Email"),
                "Skills": st.column_config.TextColumn("Skills"),
                "Experience": st.column_config.TextColumn("Experience")
            },
            hide_index=True
        )
        
        # Add selected candidates to group
        if st.button("Add Selected to Group"):
            selected_candidates = candidates_df[st.session_state.candidate_selection['Selected']]['Email'].tolist()
            if selected_candidates:
                # Remove duplicates by converting to set
                current_group = set(st.session_state.groups[selected_group])
                current_group.update(selected_candidates)
                st.session_state.groups[selected_group] = list(current_group)
                st.success(f"Added {len(selected_candidates)} candidates to group!")
                # Clear selection after adding
                st.session_state.candidate_selection['Selected'] = False
            else:
                st.warning("No candidates selected")
        
        # Display group members
        if st.session_state.groups[selected_group]:
            st.subheader("Current Group Members")
            group_members = candidates_df[candidates_df['Email'].isin(st.session_state.groups[selected_group])]
            
            # Display group members with remove option
            st.dataframe(group_members)
            
            # Option to remove members
            members_to_remove = st.multiselect(
                "Select members to remove from group",
                group_members['Email'].tolist()
            )
            if st.button("Remove Selected Members"):
                if members_to_remove:
                    st.session_state.groups[selected_group] = [
                        member for member in st.session_state.groups[selected_group] 
                        if member not in members_to_remove
                    ]
                    st.success(f"Removed {len(members_to_remove)} members from the group")
                    st.rerun()

# Panel Management Tab
with tab2:
    st.header("Panel Management")
    
    # Display available panel members
    st.subheader("Available Panel Members")
    st.dataframe(panels_df)
    
    # Create interview panel
    st.subheader("Create Interview Panel")
    selected_panel_members = st.multiselect(
        "Select Panel Members",
        panels_df['Email'].tolist()
    )
    
    panel_name = st.text_input("Enter Panel Name")
    if st.button("Create Panel") and panel_name and selected_panel_members:
        st.session_state.panels[panel_name] = selected_panel_members
        st.success(f"Panel '{panel_name}' created with {len(selected_panel_members)} members!")
    
    # Display existing panels
    if st.session_state.panels:
        st.subheader("Existing Panels")
        for panel, members in st.session_state.panels.items():
            with st.expander(f"Panel: {panel}"):
                st.write("Members:", ", ".join(members))

# Schedule Interviews Tab
with tab3:
    st.header("Schedule Interviews")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Select group
        interview_group = st.selectbox("Select Candidate Group", 
            list(st.session_state.groups.keys()) if st.session_state.groups else ["No groups created"])
        
        # Select panel
        interview_panel = st.selectbox("Select Interview Panel", 
            list(st.session_state.panels.keys()) if st.session_state.panels else ["No panels created"])
    
    with col2:
        # Schedule details
        interview_date = st.date_input("Select Interview Date")
        start_time = st.time_input("Select Start Time")
        duration = st.number_input("Interview Duration (minutes)", min_value=15, value=30)
    
    # Google Meet Link
    col1, col2 = st.columns(2)
    with col1:
        meet_link = st.text_input(
            "Meeting Link (Optional)",
            placeholder="Enter your meeting link (Google Meet, Zoom, etc.)",
            help="Enter any meeting link (Google Meet, Zoom, etc.)"
        )
    # Message input moved outside the column as per instructions
    message = st.text_area(
        "Message (Optional)",
        placeholder="Enter your message here",
        height=200,
        help="Enter any custom message for the interview. This box is large for longer messages."
    )
    if st.button("Schedule Interviews"):
        if interview_group in st.session_state.groups and interview_panel in st.session_state.panels:
            candidates = st.session_state.groups[interview_group]
            panel_members = st.session_state.panels[interview_panel]
            current_time = datetime.combine(interview_date, start_time)
            
            # Prepare panel member email
            panel_email_html = f"""
            <html>
            <body>
            <h2>Interview Schedule - {interview_date.strftime('%B %d, %Y')}</h2>
            <p>Respecet Sir/Madam, <br> You have been assigned to conduct interviews for the following candidates:</p>
            <table border="1" style="border-collapse: collapse; width: 100%;">
                <tr style="background-color: #f2f2f2;">
                    <th style="padding: 8px;">Time</th>
                    <th style="padding: 8px;">Candidate</th>
                    <th style="padding: 8px;">Details</th>
                </tr>
            """
            
            # Send emails to candidates
            for candidate in candidates:
                candidate_details = candidates_df[candidates_df['Email'] == candidate].iloc[0]
                
                # Prepare candidate email
                subject = "Interview Schedule Notification"
                body = f"""
                Dear {candidate_details.get('Name', 'Candidate')},
                
                Your interview has been scheduled for {current_time.strftime('%Y-%m-%d %H:%M')}.
                Duration: {duration} minutes
                Panel: {interview_panel}
                
                Meeting Link: {meet_link}
                
                """
                if message:
                    body += f"{message}\n\n"
                body += """Please be prepared and join on time.
                
                Best regards,
                Interview Team
                """
                
                # Add to panel email
                panel_email_html += f"""
                <tr>
                    <td style="padding: 8px;">{current_time.strftime('%H:%M')}</td>
                    <td style="padding: 8px;">{candidate_details.get('Name', candidate)}</td>
                    <td style="padding: 8px;">{candidate_details.get('Experience', '')} - {candidate_details.get('Skills', '')}</td>
                </tr>
                """
                
                # Send email to candidate
                result = send_email(candidate, subject, body)
                st.write(f"Notification to {candidate}: {result}")
                
                # Update time for next candidate
                current_time += timedelta(minutes=duration)
            
            # Complete and send panel email
            panel_email_html += """
            </table>
            <p>Meeting Link: {meet_link}</p>
            <p>Please be prepared for the interviews.</p>
            </body>
            </html>
            """
            
            # Send email to panel members
            for panel_member in panel_members:
                panel_result = send_email(
                    panel_member,
                    f"Interview Schedule - {interview_date.strftime('%B %d, %Y')}",
                    panel_email_html,
                    is_html=True
                )
                st.write(f"Notification to panel member {panel_member}: {panel_result}")
            
            st.success("All interviews scheduled and notifications sent!")
        else:
            st.error("Please create both groups and panels before scheduling interviews.")

# Custom Message Tab
with tab4:
    st.header("Send Custom Message")
    
    # Use columns for recipient type and sending mode selection
    col1, col2 = st.columns(2)
    with col1:
        recipient_type = st.radio("Select Recipient Type", ["Candidate", "Panel Member"])
    with col2:
        sending_mode = st.radio("Select Sending Mode", ["Single Recipient", "Multiple Recipients", "All Recipients"])
    
    recipients = []
    if recipient_type == "Candidate":
        if sending_mode == "Single Recipient":
            # Single candidate selection
            selected_candidate = st.selectbox(
                "Select Candidate",
                candidates_df['Email'].tolist(),
                format_func=lambda x: f"{candidates_df[candidates_df['Email'] == x]['Name'].iloc[0] if len(candidates_df[candidates_df['Email'] == x]) > 0 and pd.notna(candidates_df[candidates_df['Email'] == x]['Name'].iloc[0]) else 'Unknown'} ({x})"
            )
            if selected_candidate:
                recipients = [selected_candidate]
                candidate_details = candidates_df[candidates_df['Email'] == selected_candidate].iloc[0]
                with st.expander("View Candidate Details"):
                    st.write(f"Name: {candidate_details.get('Name', 'N/A')}")
                    st.write(f"Email: {candidate_details.get('Email', 'N/A')}")
                    st.write(f"Skills: {candidate_details.get('Skills', 'N/A')}")
                    st.write(f"Experience: {candidate_details.get('Experience', 'N/A')}")
        
        elif sending_mode == "Multiple Recipients":
            # Multiple candidates selection
            selected_candidates = st.multiselect(
                "Select Candidates",
                candidates_df['Email'].tolist(),
                format_func=lambda x: f"{candidates_df[candidates_df['Email'] == x]['Name'].iloc[0] if len(candidates_df[candidates_df['Email'] == x]) > 0 and pd.notna(candidates_df[candidates_df['Email'] == x]['Name'].iloc[0]) else 'Unknown'} ({x})"
            )
            recipients = selected_candidates
            if selected_candidates:
                with st.expander("View Selected Candidates Details"):
                    for candidate in selected_candidates:
                        candidate_details = candidates_df[candidates_df['Email'] == candidate].iloc[0]
                        st.write("---")
                        st.write(f"Name: {candidate_details.get('Name', 'N/A')}")
                        st.write(f"Email: {candidate_details.get('Email', 'N/A')}")
                        st.write(f"Skills: {candidate_details.get('Skills', 'N/A')}")
                        st.write(f"Experience: {candidate_details.get('Experience', 'N/A')}")
        
        else:  # All Recipients
            recipients = candidates_df['Email'].tolist()
            with st.expander(f"View All Candidates ({len(recipients)})"):
                st.dataframe(candidates_df)
    
    else:  # Panel Member
        if sending_mode == "Single Recipient":
            # Single panel member selection
            selected_panel_member = st.selectbox(
                "Select Panel Member",
                panels_df['Email'].tolist(),
                format_func=lambda x: f"{panels_df[panels_df['Email'] == x]['Name'].iloc[0] if len(panels_df[panels_df['Email'] == x]) > 0 and pd.notna(panels_df[panels_df['Email'] == x]['Name'].iloc[0]) else 'Unknown'} ({x})"
            )
            if selected_panel_member:
                recipients = [selected_panel_member]
                panel_details = panels_df[panels_df['Email'] == selected_panel_member].iloc[0]
                with st.expander("View Panel Member Details"):
                    st.write(f"Name: {panel_details.get('Name', 'N/A')}")
                    st.write(f"Email: {panel_details.get('Email', 'N/A')}")
                    st.write(f"Expertise: {panel_details.get('Expertise', 'N/A')}")
        
        elif sending_mode == "Multiple Recipients":
            # Multiple panel members selection
            selected_panel_members = st.multiselect(
                "Select Panel Members",
                panels_df['Email'].tolist(),
                format_func=lambda x: f"{panels_df[panels_df['Email'] == x]['Name'].iloc[0] if len(panels_df[panels_df['Email'] == x]) > 0 and pd.notna(panels_df[panels_df['Email'] == x]['Name'].iloc[0]) else 'Unknown'} ({x})"
            )
            recipients = selected_panel_members
            if selected_panel_members:
                with st.expander("View Selected Panel Members Details"):
                    for member in selected_panel_members:
                        panel_details = panels_df[panels_df['Email'] == member].iloc[0]
                        st.write("---")
                        st.write(f"Name: {panel_details.get('Name', 'N/A')}")
                        st.write(f"Email: {panel_details.get('Email', 'N/A')}")
                        st.write(f"Expertise: {panel_details.get('Expertise', 'N/A')}")
        
        else:  # All Recipients
            recipients = panels_df['Email'].tolist()
            with st.expander(f"View All Panel Members ({len(recipients)})"):
                st.dataframe(panels_df)

    # Show number of selected recipients
    if recipients:
        st.info(f"Selected Recipients: {len(recipients)}")

    # Message composition
    st.subheader("Compose Message")
    message_subject = st.text_input("Subject")
    
    # Add personalized greeting option for candidates
    if recipient_type == "Candidate":
        use_personalized_greeting = st.checkbox("Use personalized greeting (Dear [Name])", value=True)
    else:
        use_personalized_greeting = False
    
    message_body = st.text_area("Message Body", height=200, 
                               placeholder="Enter your message here. Use [NAME] as a placeholder for the recipient's name if you want personalized greetings.")
    
    # Add meeting link option
    include_meeting_link = st.checkbox("Include Meeting Link")
    if include_meeting_link:
        meeting_link = st.text_input("Meeting Link")
        if meeting_link:
            message_body = f"{message_body}\n\nMeeting Link: {meeting_link}"
    
    # Send button with progress tracking
    if st.button("Send Message"):
        if not message_subject or not message_body:
            st.error("Please enter both subject and message body")
        elif not recipients:
            st.error("Please select at least one recipient")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            successful_sends = 0
            failed_sends = 0
            
            for idx, recipient in enumerate(recipients):
                status_text.text(f"Sending email to {recipient}...")
                
                # Personalize message for each recipient
                personalized_body = message_body
                
                if recipient_type == "Candidate" and use_personalized_greeting:
                    # Get candidate details
                    candidate_details = candidates_df[candidates_df['Email'] == recipient]
                    if len(candidate_details) > 0:
                        candidate_name = candidate_details.iloc[0].get('Name', 'Candidate')
                        if pd.notna(candidate_name):
                            # Replace [NAME] placeholder with actual name
                            personalized_body = message_body.replace('[NAME]', str(candidate_name))
                            # If no [NAME] placeholder found, add "Dear [Name]" at the beginning
                            if '[NAME]' not in message_body:
                                personalized_body = f"Dear {candidate_name},\n\n{message_body}"
                
                elif recipient_type == "Panel Member" and use_personalized_greeting:
                    # Get panel member details
                    panel_details = panels_df[panels_df['Email'] == recipient]
                    if len(panel_details) > 0:
                        panel_name = panel_details.iloc[0].get('Name', 'Panel Member')
                        if pd.notna(panel_name):
                            # Replace [NAME] placeholder with actual name
                            personalized_body = message_body.replace('[NAME]', str(panel_name))
                            # If no [NAME] placeholder found, add "Dear [Name]" at the beginning
                            if '[NAME]' not in message_body:
                                personalized_body = f"Dear {panel_name},\n\n{message_body}"
                
                result = send_email(recipient, message_subject, personalized_body)
                
                if result == "Email sent successfully!":
                    successful_sends += 1
                else:
                    failed_sends += 1
                    st.error(f"Failed to send to {recipient}: {result}")
                
                # Update progress
                progress = (idx + 1) / len(recipients)
                progress_bar.progress(progress)
            
            # Final status
            if failed_sends == 0:
                st.success(f"Successfully sent messages to all {len(recipients)} recipients!")
            else:
                st.warning(f"Sent messages to {successful_sends} recipients. Failed to send to {failed_sends} recipients.")
            
            status_text.text("Completed sending messages.")

# Show warning if Gmail not configured
if not gmail_email or not gmail_password:
    st.sidebar.warning("⚠️ Please configure your Gmail settings to enable email notifications")

st.write("Note: Make sure your Gmail settings are configured correctly in the sidebar.")