import requests
from datetime import datetime

def process_email_by_id(message_id, access_token, calendar_manager, gemini_parser):
    """Fetch email by ID, parse with Gemini, and create calendar event if meeting detected"""
    try:
        # Step 1: Fetch email details
        url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}"
        headers = {"Authorization": f"Bearer {access_token}"}
        response = requests.get(url, headers=headers)

        if response.status_code != 200:
            print(f"❌ Failed to fetch email: {response.status_code}")
            return

        email = response.json()
        subject = email.get("subject", "")
        body_preview = email.get("bodyPreview", "")
        sender_info = email["sender"]["emailAddress"]
        sender_name = sender_info.get("name")
        sender_email = sender_info.get("address")

        # Step 2: Parse the email using Gemini
        parsed = gemini_parser.parse_email_with_retry(body_preview, subject)

        if parsed.get("has_meeting"):
            print(f"📅 Meeting detected: {parsed.get('subject')} | Creating calendar event...")
            calendar_manager.create_event(parsed, sender_email, sender_name)
        else:
            print(f"💬 No meeting in: {subject}")

    except Exception as e:
        print(f"🚨 Error in processing email {message_id}: {str(e)}")
