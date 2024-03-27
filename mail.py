import streamlit as st
import pandas as pd
import plotly.express as px
import imaplib
import email
from email.header import decode_header
import yaml
import re
from nltk.tokenize import sent_tokenize

st.set_page_config(
    page_title="Email Extractor",
    page_icon="📧",
    layout="wide",
)

def read_credentials(filename):
    with open(filename, 'r') as file:
        credentials = yaml.safe_load(file)
    return credentials['user'], credentials['password']

def extract_phone_numbers(text):
    phone_regex = re.compile(r'\+?91[ -]?\d{10}|\b\d{5}[-.\s]?\d{5}\b')
    phone_numbers = re.findall(phone_regex, text)
    return phone_numbers

def get_email_summary(body):
    sentences = sent_tokenize(body)
    summary = " ".join(sentences[:3])
    return summary

def get_email_body(msg):
    parts = []
    for part in msg.walk():
        content_type = part.get_content_type()
        if content_type == "text/plain":
            charset = part.get_content_charset() or "utf-8"
            body = part.get_payload(decode=True).decode(charset, errors="ignore")
            parts.append(body)
        elif content_type == "text/html":
            pass
        elif content_type.startswith("image/"):
            alt_text = part.get("Content-Description")
            parts.append(f"Image Alt Text: {alt_text}" if alt_text else "Image (No Alt Text)")
    if parts:
        return "\n".join(parts)
    else:
        return "No text content available."

def fetch_emails_with_filters(email_address, password, keyword, from_filter, to_filter, date_filter, read_status, mailbox="inbox", server="imap.gmail.com", port=993, num_emails=10):

    try:
        mail = imaplib.IMAP4_SSL(server, port)
        mail.login(email_address, password)
        mail.select(mailbox)
        
        _, total_messages_data = mail.search(None, 'ALL')
        email_ids = total_messages_data[0].split() if total_messages_data and total_messages_data[0] else []

        filtered_email_count = len(email_ids)

        if email_ids and read_status != "All":
            _, unread_messages_data = mail.search(None, 'UNSEEN')
            unread_email_ids = unread_messages_data[0].split() if unread_messages_data and unread_messages_data[0] else []
            unread_email_count = len(unread_email_ids)
            read_email_count = filtered_email_count - unread_email_count
        else:
            unread_email_count = 0
            read_email_count = 0

        email_details_list = []

        for email_id in email_ids[:num_emails]:
            _, msg_data = mail.fetch(email_id, "(RFC822)")
            if msg_data and msg_data[0]:
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)
            else:
                continue

            sender_name, sender_email = email.utils.parseaddr(msg["From"])
            subject, encoding = decode_header(msg["Subject"])[0]
            subject = subject.decode(encoding or "utf-8") if isinstance(subject, bytes) else subject

            phone_numbers = extract_phone_numbers(get_email_body(msg))
            summary = get_email_summary(get_email_body(msg))
            keyword_match = keyword.lower() in subject.lower() or keyword.lower() in get_email_body(msg).lower()
            unread_status = email_id in unread_email_ids if read_status != "All" else False

            if (read_status == "Read" and not unread_status) or (read_status == "Unread" and unread_status) or (read_status == "All"):
                if keyword == "" or keyword_match:
                    date_time_received = email.utils.parsedate_to_datetime(msg["Date"])
                    date_received = date_time_received.strftime("%Y-%m-%d")

                    email_details_list.append({
                        "Sender Name": sender_name,
                        "Sender Email": sender_email,
                        "Subject": subject,
                        "Date": date_received,
                        "Summary": summary,
                        "Phone Numbers": phone_numbers,
                        "Read Status": "Read" if not unread_status else "Unread"
                    })

        if mailbox == "inbox":
            total_emails_inbox = len(email_details_list)  # Update the variable here
        elif mailbox == "sent":
            total_emails_sent = len(email_details_list)  # Update the variable here

        mail.logout()

        df = pd.DataFrame(email_details_list)

        if mailbox == "inbox":
            return df, filtered_email_count, unread_email_count, read_email_count, total_emails_inbox
        elif mailbox == "sent":
            return df, filtered_email_count, unread_email_count, read_email_count, total_emails_sent

    except Exception as e:
        print(f"Error: {e}")

def main():
    st.title("Email Extractor")
    st.sidebar.header("FILTERS")

    email_address, password = read_credentials('credentials.yaml')

    with st.sidebar.expander("Inbox"):
        search_keyword_inbox = st.text_input("Enter the keyword to filter emails:", key="search_inbox")
        from_filter_inbox = st.text_input("From:", key="from_inbox")
        date_filter_inbox = st.text_input("Date:", key="date_inbox")
        read_status_inbox = st.selectbox("Filter by:", ["All", "Read", "Unread"], index=0, key="read_status_inbox")
        num_emails_inbox = st.selectbox("Number of emails to display:", ["All", 10, 20, 50], index=0, key="num_emails_inbox")

    with st.sidebar.expander("Sent"):
        to_filter_sent = st.text_input("To:", key="to_sent")
        search_keyword_sent = st.text_input("Enter the keyword to filter emails:", key="search_sent")

    if st.sidebar.button("Fetch Emails"):
        if num_emails_inbox == "All":
            num_emails_inbox = None
        else:
            num_emails_inbox = int(num_emails_inbox)

        email_details_df_inbox, filtered_emails_inbox, unread_emails_inbox, read_emails_inbox, total_emails_inbox = fetch_emails_with_filters(email_address, password, search_keyword_inbox, from_filter_inbox, to_filter_sent, date_filter_inbox, read_status_inbox, num_emails=num_emails_inbox, mailbox="inbox")

        if email_details_df_inbox is not None and not email_details_df_inbox.empty:
            st.success("Inbox Emails fetched successfully!")

            st.sidebar.markdown("---")

            try:
                excel_filename_inbox = f"email_details_streamlit_inbox.xlsx"
                email_details_df_inbox.to_excel(excel_filename_inbox, index=False)
                st.info(f"Inbox Email details saved to {excel_filename_inbox}")
            except PermissionError:
                st.warning("Permission denied to save the file. Please check your permissions.")

        else:
            st.warning("No Inbox emails found matching the criteria.")
    
    # Displaying the dashboard on the main screen
    st.subheader("Email Analysis Dashboard")

    if 'total_emails_inbox' in locals():
        st.write(f"Total Emails in Inbox: {total_emails_inbox}")
        if read_status_inbox == "Unread":
            st.write(f"Unread Emails: {unread_emails_inbox}")
        elif read_status_inbox == "Read":
            st.write(f"Read Emails: {read_emails_inbox}")

        st.subheader("Bar Chart: Number of Emails by Sender")
        sender_counts = email_details_df_inbox['Sender Name'].value_counts()
        bar_chart_sender = px.bar(sender_counts, x=sender_counts.index, y=sender_counts.values, labels={'x':'Sender Name', 'y':'Number of Emails'}, title='Number of Emails by Sender')
        st.plotly_chart(bar_chart_sender, use_container_width=True)

        st.subheader("Pie Chart: Read/Unread Emails Distribution")
        read_status_distribution = email_details_df_inbox['Read Status'].value_counts()
        pie_chart_read_status = px.pie(read_status_distribution, values=read_status_distribution.values, names=read_status_distribution.index, title='Read/Unread Emails Distribution')
        st.plotly_chart(pie_chart_read_status, use_container_width=True)

if __name__ == "__main__":
    main()
