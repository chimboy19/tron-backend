import os
import time
import requests
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_URL = "https://graph.microsoft.com/v1.0"

# --------------------------------------------------------------------
# AUTHENTICATION
# --------------------------------------------------------------------
def get_access_token():
    app = ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}"
    )
    token = app.acquire_token_silent(GRAPH_SCOPE, account=None)
    if not token:
        print("[AUTH] Fetching new token…")
        token = app.acquire_token_for_client(scopes=GRAPH_SCOPE)

    if "access_token" not in token:
        print("[AUTH ERROR]", token)
        return None

    print("[AUTH] Token acquired")
    return token["access_token"]


# --------------------------------------------------------------------
# SEND EMAIL
# --------------------------------------------------------------------
def send_email(subject, body, recipient):
    access_token = get_access_token()
    url = f"{GRAPH_URL}/users/{EMAIL_ADDRESS}/sendMail"

    payload = {
        "message": {
            "subject": subject,
            "body": { "contentType": "HTML", "content": body },
            "toRecipients": [{"emailAddress": {"address": recipient}}]
        }
    }

    r = requests.post(url, json=payload, headers={"Authorization": f"Bearer {access_token}"})
    print("[SEND MAIL]", r.status_code, r.text)


# --------------------------------------------------------------------
# CREATE DRAFT
# --------------------------------------------------------------------
def create_draft(subject, body, recipient):
    access_token = get_access_token()
    url = f"{GRAPH_URL}/users/{EMAIL_ADDRESS}/messages"

    payload = {
        "subject": subject,
        "body": { "contentType": "HTML", "content": body },
        "toRecipients": [{"emailAddress": {"address": recipient}}]
    }

    r = requests.post(url, json=payload, headers={"Authorization": f"Bearer {access_token}"})
    print("[CREATE DRAFT]", r.status_code, r.text)


# --------------------------------------------------------------------
# MONITOR INBOX (POLLING)
# --------------------------------------------------------------------
def monitor_inbox():
    access_token = get_access_token()
    url = f"{GRAPH_URL}/users/{EMAIL_ADDRESS}/mailFolders/Inbox/messages?$top=5&$filter=isRead eq false"

    r = requests.get(url, headers={"Authorization": f"Bearer {access_token}"})

    if r.status_code != 200:
        print("[INBOX ERROR]", r.status_code, r.text)
        return

    messages = r.json().get("value", [])

    if not messages:
        print("[INBOX] No unread emails")
    else:
        print(f"[INBOX] {len(messages)} unread email(s):")
        for m in messages:
            print("----")
            print("From:", m["from"]["emailAddress"]["address"])
            print("Subject:", m["subject"])
            print("Received:", m["receivedDateTime"])
            print("----")


# --------------------------------------------------------------------
# MAIN LOOP
# --------------------------------------------------------------------
if __name__ == "__main__":
    print("---- Graph API Email Monitor Started ----")

    while True:
        try:
            monitor_inbox()
            time.sleep(10)  # check inbox every 10 seconds
        except KeyboardInterrupt:
            print("\nStopped.")
            break
        except Exception as e:
            print("[ERROR]", e)
            time.sleep(5)
