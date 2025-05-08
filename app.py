import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time

st.set_page_config(page_title="Cotton Industry Mailer", layout="centered")
st.title("📨 Hariom Cotton Industries Email Sender")
st.write("Upload an Excel file containing customer emails and send a promotional message.")

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("📋 Preview of Uploaded Contacts")
    st.dataframe(df)

    st.subheader("✉️ Gmail Credentials")
    sender_email = st.text_input("Sender Gmail", placeholder="example@gmail.com")
    app_password = st.text_input("App Password", type="password")

    default_message = (
        "Dear Sir/Madam,"
        "Greetings from Hariom Cotton Industries."
        "We are a well-established cotton ginning enterprise based in Gujarat, India, specializing in premium-quality cotton products. We are reaching out to explore potential business opportunities and would be honored to collaborate with you."
        "If our offerings align with your needs, we would be glad to discuss further details at your convenience. Please feel free to reply to this email for any inquiries or to initiate a conversation."
        "Looking forward to the possibility of working together."
        "Warm regards, 
        "Hariom Cotton Industries  
        "Gujarat, India"
    )

    subject = st.text_input("Email Subject", "Supply of Premium Cotton Products – Hariom Cotton Industries")
    message_body = st.text_area("Email Body", default_message)

    if st.button("📬 Send Emails"):
        if not sender_email or not app_password:
            st.error("Please enter both sender email and app password.")
        else:
            sent_count = 0
            for index, row in df.iterrows():
                name = row.get("Name", "")
                recipient = row.get("Email", "")
                if not recipient:
                    continue

                final_message = f"Hi {name},\n\n{message_body}\n\nRegards,\nHariom Cotton Industries"

                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = recipient
                msg['Subject'] = subject
                msg.attach(MIMEText(final_message, 'plain'))

                try:
                    with smtplib.SMTP("smtp.gmail.com", 587) as server:
                        server.starttls()
                        server.login(sender_email, app_password)
                        server.sendmail(sender_email, recipient, msg.as_string())
                        sent_count += 1
                        st.success(f"✅ Sent to {name} <{recipient}>")
                        time.sleep(1)  # Small delay to reduce risk of throttling
                except Exception as e:
                    st.error(f"❌ Failed to send to {recipient}: {e}")
                    time.sleep(1)

            st.info(f"📧 Emails sent: {sent_count}")
