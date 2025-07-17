import requests
from datetime import datetime, timedelta

def subscribe_to_emails(access_token):
    expiration = (datetime.utcnow() + timedelta(days=3)).isoformat() + "Z"
    url = "https://graph.microsoft.com/v1.0/subscriptions"
    payload = {
        "changeType": "created",
        "notificationUrl": "https://your-ngrok-or-cloud-endpoint/notifications",
        "resource": "me/mailFolders('Inbox')/messages",
        "expirationDateTime": expiration,
        "clientState": "secret123"
    }

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.post(url, headers=headers, json=payload)
    print("Subscription response:", response.json())
