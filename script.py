import email.parser
import imaplib
import sys
import re
import ssl
import csv


SEARCH_FOLDER = ['"[Gmail]/&BigGMQZKBi8- &BkUGMQYzBkQ-"', '"INBOX"']  # select the folder that you want to retrieve
DEFAULT_MAIL_SERVER = 'imap.gmail.com' # connect to gmail
USERNAME = 'al.dweni001@gmail.com'
PASSWORD = 'ifagqbkpnvnwfgdz'
excluded_words = ['support', 'reply', 'service', 'notification', 'notify', 'lamah', 'facebook', 'adobe']

# No user parameters below this line
ADDR_PATTERN = re.compile("<(.+)>")  # Finds email as <nospam@nospam.com>

def connect(user, pwd, server=DEFAULT_MAIL_SERVER):
    """Connect to [the specified] mail server. Return an open connection"""
    conn = imaplib.IMAP4_SSL(host=server,
            ssl_context=ssl.create_default_context())
    try:
        conn.login(user, pwd)
    except imaplib.IMAP4.error:
        print("Failed to login")
        sys.exit(1)
    return conn


def print_folders(conn):
    """Print a list of open mailbox folders"""
    for f in conn.list():
        for i in f:
            print("\t", i)


# def filter_by_words(word_list, substrings):
#     filtered_list = [word for word in word_list if all(substring not in word for substring in substrings)]
#     return filtered_list


def get_mails_from_folder(conn, folder_name):
    """Fetch a specific folder (or label) from server"""
    typ, data = conn.select(mailbox=folder_name, readonly=True)
    if typ != 'OK':
        print("Could not open specified folder. Known labels:")
        print_folders(conn)
        return

    typ, data = conn.search(None, 'ALL')
    if typ != 'OK':
        print("Could not get mail list of folder: ", folder_name)
        return

    return data[0].split()

def fetch_message(conn, msg_uid):
    """
    Fetch a specific message uid (not sequential id!) from the given folder;
    return the parsed message. User must ensure that specified
    message ID exists in that folder.
    """
    # TODO: Could we fetch just the envelope of the response to save bandwidth?
    typ, data = conn.fetch(msg_uid, '(RFC822)')
    if typ != 'OK':
        print("ERROR fetching message #", msg_uid)
        return

    return email.parser.BytesParser().parsebytes(data[0][1], headersonly=True)


def get_recipients(msg):
    """Given a parsed message, extract and return recipient list"""
    recipients = []
    addr_fields = ['From', 'To', 'Cc', 'Bcc', 'Reply-To', 'Sender']

    for f in addr_fields:
        if msg[f] is None:
            continue

        # str conversion is needed for non-ascii chars
        rlist = ADDR_PATTERN.findall(str(msg[f]))
        if not any(word in rlist[0] for word in excluded_words):
            all_emails.extend(rlist)
        # recipients.extend(filter_list)

    return recipients


if __name__ == "__main__":
    # Connect
    mail_conn = connect(USERNAME, PASSWORD)

    # show folders of mail account
    #print_folders(mail_conn)

    # Open folders and get list of email message uids
    all_emails = []
    for folder in SEARCH_FOLDER:
        # switch to folder
        for mail_id in get_mails_from_folder(mail_conn, folder):
            data = fetch_message(mail_conn, mail_id)
            recip_list = get_recipients(data)
            all_emails.extend(recip_list)

        mail_conn.close()

    mail_conn.logout()

    with open('emails.csv', 'w', newline='') as emails_file:
        writer = csv.writer(emails_file)
        writer.writerow(['Emails'])
        writer.writerows([[email] for email in set(all_emails)])