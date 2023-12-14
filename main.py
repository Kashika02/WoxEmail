import streamlit as st
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Streamlit app
st.title("Email Sender App")

# Email configuration
sender_email = "kashika.parmar@woxsen.edu.in"  # replace with your email
sender_password = "Fah32413"  # replace with your password

# Input box for email content
email_content = st.text_area("Enter your email content here:", "")

# Input box for subject
email_subject = st.text_input("Enter the subject of the email:", "")

# Input box for recipient email addresses (comma-separated)
recipient_emails = st.text_input("Enter recipient email addresses (comma-separated):", "")

# Button to send emails
if st.button("Send Emails"):

    # Splitting email addresses
    recipients = recipient_emails.split(",")

    # SMTP setup
    try:
        # Use Outlook SMTP server and port
        server = smtplib.SMTP("smtp-mail.outlook.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)

        for recipient in recipients:
            # Email content
            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = recipient.strip()
            msg["Subject"] = email_subject  # Use the user-provided subject
            body = email_content
            msg.attach(MIMEText(body, "plain"))

            # Send email
            server.sendmail(sender_email, recipient.strip(), msg.as_string())

        st.success("Emails sent successfully!")

    except Exception as e:
        st.error(f"Error: {e}")
    finally:
        server.quit()
