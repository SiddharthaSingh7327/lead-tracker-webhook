import requests
from datetime import datetime, timedelta
from get_emails import AuthManager
from config import Config
from get_emails import FileManager

def renew_subscription():
    config = Config()
    file_manager = FileManager()
    auth_manager = AuthManager(config, file_manager)

    access_token = auth_manager.get_access_token()
    if not access_token:
        print("❌ Failed to authenticate")
        return

    expiration = (datetime.utcnow() + timedelta(days=3)).isoformat() + "Z"
    url = "https://graph.microsoft.com/v1.0/subscriptions"
    payload = {
        "changeType": "created",
        "notificationUrl": "https://your-public-endpoint.com/notifications",  # replace with real
        "resource": "me/mailFolders('Inbox')/messages",
        "expirationDateTime": expiration,
        "clientState": "leadtracker123"
    }

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.post(url, headers=headers, json=payload)
    if response.status_code == 201:
        print(f"✅ Subscription renewed until: {expiration}")
    else:
        print(f"⚠️ Renewal failed: {response.status_code} - {response.text}")

if __name__ == "__main__":
    renew_subscription()
