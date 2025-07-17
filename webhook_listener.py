from flask import Flask, request, jsonify
import logging

from get_emails import AuthManager
from get_emails import CalendarManager
from get_emails import GeminiParser
from email_processor import process_email_by_id
from config import Config
from get_emails import FileManager

# ---------------- Setup ----------------
app = Flask(__name__)

# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize core services
config = Config()
file_manager = FileManager()
auth_manager = AuthManager(config, file_manager)
calendar_manager = CalendarManager(headers={}, file_manager=file_manager)
gemini_parser = GeminiParser(config.GEMINI_API_KEY)

# ---------------- Route ----------------
@app.route("/notifications", methods=["POST"])
def receive_notification():
    # Microsoft validation handshake on first subscription
    token = request.args.get("validationToken")
    if token:
        return token, 200, {'Content-Type': 'text/plain'}

    # Process actual notifications
    data = request.get_json()
    logger.info(f"🔔 Received notification: {data}")

    for notification in data.get("value", []):
        message_id = notification.get("resourceData", {}).get("id")
        if message_id:
            access_token = auth_manager.get_access_token()
            calendar_manager.headers = {"Authorization": f"Bearer {access_token}"}
            process_email_by_id(message_id, access_token, calendar_manager, gemini_parser)
        else:
            logger.warning("No message ID found in notification payload")

    return jsonify({"status": "processed"}), 202

# ---------------- Run Server ----------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
