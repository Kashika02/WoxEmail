import streamlit as st
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import pandas as pd

# Streamlit app
st.title("Email Sender App")

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
