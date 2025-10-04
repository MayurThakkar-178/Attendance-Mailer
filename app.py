import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time

# Streamlit setup
st.set_page_config(page_title="Attendance Mailer", layout="centered")
st.title("📨 Attendance Alert Mailer")
st.write("Upload an Excel sheet with student attendance. Emails will be sent automatically to students with <85% attendance.")

# Upload Excel file
uploaded_file = st.file_uploader("Upload Attendance Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    st.subheader("📋 Preview of Uploaded Attendance")
    st.dataframe(df)

    st.subheader("✉️ Faculty Gmail Credentials")
    sender_email = st.text_input("Sender Gmail", placeholder="example@gmail.com")
    app_password = st.text_input("App Password", type="password", help="Use Gmail App Password for security.")

    subject = st.text_input("Email Subject", "Attendance Review Notification")

    if st.button("📬 Send Attendance Emails"):
        if not sender_email or not app_password:
            st.error("Please enter both sender email and app password.")
        else:
            sent_count = 0
            for index, row in df.iterrows():
                name = row.get("Name", "")
                email = row.get("Email", "")
                attendance = row.get("Attendance (%)", 100)

                if pd.isnull(email) or pd.isnull(attendance):
                    continue

                if attendance < 85:  # threshold
                    message_body = f"""Dear {name},

Our records show that your attendance is currently {attendance}%, which is below the minimum requirement of 85%.

Please review your attendance and take necessary steps to improve it.

Regards,
Faculty / Administration
"""

                    # Construct email
                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = email
                    msg['Subject'] = subject
                    msg.attach(MIMEText(message_body, 'plain'))

                    try:
                        with smtplib.SMTP("smtp.gmail.com", 587) as server:
                            server.starttls()
                            server.login(sender_email, app_password)
                            server.sendmail(sender_email, email, msg.as_string())
                            sent_count += 1
                            st.success(f"✅ Email sent to {name} ({email})")
                            time.sleep(1)  # avoid throttling
                    except Exception as e:
                        st.error(f"❌ Failed to send to {email}: {e}")
                        time.sleep(1)

            st.info(f"📧 Total Emails Sent: {sent_count}")
