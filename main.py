import streamlit as st
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import pandas as pd

def send_individual_email(sender_email, sender_password, recipient_email, email_content, email_subject, attachments, cc_emails):
    try:
    # Extract email addresses from the input
        recipients = [email.strip() for email in recipient_email.split(",")]

        # Splitting CC email addresses
        cc_recipients = [cc.strip() for cc in cc_emails.split(",")] if cc_emails else []

        # SMTP setup
        server = smtplib.SMTP("smtp-mail.outlook.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)

        # Send emails to main recipients
        for recipient in recipients:
            # Email content
            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = recipient.strip()
            msg["Subject"] = email_subject  # Use the user-provided subject
            body = email_content
            msg.attach(MIMEText(body, "plain"))

            # Attach uploaded files
            if attachments:
                for attachment in attachments:
                    file_name = attachment.name
                    attachment_content = attachment.read()
                    file_part = MIMEApplication(attachment_content)
                    file_part.add_header('Content-Disposition', f'attachment; filename="{file_name}"')
                    msg.attach(file_part)

            # Adding CC recipients
            if cc_recipients:
                msg["CC"] = ", ".join(cc_recipients)

            # Print email content and recipient email ID
            print(f"Email Content: {body}")
            print(f"Recipient Email ID: {recipient.strip()}")

            # Send email to main recipient and CC recipients
            recipients_all = [recipient.strip()] + cc_recipients
            server.sendmail(sender_email, recipients_all, msg.as_string())

        st.success("Emails sent successfully!")

    except Exception as e:
        st.error(f"Error: {e}")
    finally:
        server.quit()



# Function to send emails
def send_emails(sender_email, sender_password, email_content, email_subject, attachments, cc_emails, xlsx_file):
    try:
        # Read XLSX file
        df = pd.read_excel(xlsx_file, engine='openpyxl')

        # Extract father's email addresses from the "Father EmailId" column
        recipients = df["Father EmailId"].astype(str).tolist()

        # Extract email addresses from the "Email Id" column
        email_recipients = df["Email Id"].astype(str).tolist()

        # Combine both lists of recipients
        recipients += email_recipients

        # Splitting CC email addresses
        cc_recipients = [cc.strip() for cc in cc_emails.split(",")]

        # SMTP setup
        server = smtplib.SMTP("smtp-mail.outlook.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)

        # Send emails to main recipients
        for recipient in recipients:
            # Email content
            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = recipient.strip()
            msg["Subject"] = email_subject  # Use the user-provided subject
            body = email_content
            msg.attach(MIMEText(body, "plain"))

            # Attach uploaded files
            if attachments:
                for attachment in attachments:
                    file_name = attachment.name
                    attachment_content = attachment.read()
                    file_part = MIMEApplication(attachment_content)
                    file_part.add_header('Content-Disposition', f'attachment; filename="{file_name}"')
                    msg.attach(file_part)

            # Adding CC recipients
            if cc_recipients:
                msg["CC"] = ", ".join(cc_recipients)

            # Print email content and recipient email ID
            print(f"Email Content: {body}")
            print(f"Recipient Email ID: {recipient.strip()}")

            # Send email to main recipient and CC recipients
            recipients_all = [recipient.strip()] + cc_recipients
            server.sendmail(sender_email, recipients_all, msg.as_string())

        st.success("Emails sent successfully!")

    except Exception as e:
        st.error(f"Error: {e}")
    finally:
        server.quit()

# Streamlit app


# Sidebar navigation
selected_option = st.sidebar.selectbox("Select an option", ["Home", "Parent-Student Update", "Send Individual Email"])

# Welcome pa9ge
if selected_option == "Home":
    st.title("Welcome to WoxMail")
    st.write("Navigate effortlessly between two distinct realms – 'Parent-Student Update' and 'Send Individual Emails' – tailored to cater to your diverse messaging needs.")

# Email Sender App
elif selected_option == "Parent-Student Update":
    st.title("Parent-Student Update")
    st.write("In the 'Parent-Student Update' section, experience the seamless flow of announcements and personalized communication.")

    # Email configuration
    # Email configuration
    sender_email = "Woxsenuniversity@outlook.com"  # replace with your email
    sender_password = "Woxsen@123"  # replace with your password

    # Input box for email content
    email_content = st.text_area("Enter your email content here:", "")

    # Input box for subject
    email_subject = st.text_input("Enter the subject of the email:", "")

    # File uploader for attachments
    attachments = st.file_uploader("Upload Attachments (Optional)", type=["pdf", "csv", "xlsx", "jpg", "png", "txt"], accept_multiple_files=True)

    # Input box for CC email addresses (comma-separated)
    cc_emails = st.text_input("Enter CC email addresses (comma-separated):", "")

    xlsx_file = st.sidebar.file_uploader("Upload XLSX File", type=["xlsx"])

    # Button to send emails
    if st.button("Send Emails"):
        send_emails(sender_email, sender_password, email_content, email_subject, attachments, cc_emails, xlsx_file)
elif selected_option == "Send Individual Email":
    st.title("Send Individual Email")
    st.write("Switch gears to the 'Send Individual Emails' section, Craft tailored messages with ease, ensuring your individual communications stand out.")

    # Email configuration
    sender_email = "Woxsenuniversity@outlook.com"  # replace with your email
    sender_password = "Woxsen@123"  # replace with your password

    # Input box for email content
    email_content_individual = st.text_area("Enter your email content here:", "")

    # Input box for subject
    email_subject_individual = st.text_input("Enter the subject of the email:", "")

    # File uploader for attachments
    attachments_individual = st.file_uploader("Upload Attachments (Optional)", type=["pdf", "csv", "xlsx", "jpg", "png", "txt"], accept_multiple_files=True)

    # Input box for CC email addresses (comma-separated)
    cc_emails_individual = st.text_input("Enter CC email addresses (comma-separated):", "")

    # Input box for recipients' email addresses (comma-separated)
    recipient_email = st.text_area("Enter recipients' email addresses (comma-separated):", "")

    # Button to send individual email
    if st.button("Send Individual Email"):
        send_individual_email(sender_email, sender_password, recipient_email, email_content_individual, email_subject_individual, attachments_individual, cc_emails_individual)
