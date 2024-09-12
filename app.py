from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage
from dotenv import load_dotenv
import os
import re
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Load environment variables
load_dotenv()

app = Flask(__name__)

# Set up LINE API
Channel_secret = os.getenv('LINE_CHANNEL_SECRET')
Channel_access_token = os.getenv('LINE_CHANNEL_ACCESS_TOKEN')

if not Channel_secret or not Channel_access_token:
    raise ValueError("LINE_CHANNEL_SECRET and LINE_CHANNEL_ACCESS_TOKEN must be set in the environment.")

line_bot_api = LineBotApi(Channel_access_token)
handler = WebhookHandler(Channel_secret)

# Set up Google Sheets connection
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Load JSON credentials from environment variable
json_creds = os.getenv("GOOGLE_CREDENTIALS_JSON", "{}")
creds_dict = json.loads(json_creds)
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)

try:
    client = gspread.authorize(creds)
    sheet = client.open_by_key("14gzrAAgOd8I3NnyXQtd4J9gHkuyrDL_uD3sV7K4S0qo").sheet1
except gspread.SpreadsheetNotFound:
    print("Error: Google Sheet not found.")
except Exception as e:
    print(f"Error opening Google Sheet: {e}")

@app.route("/")
def home():
    return "Welcome to the Flask server!"

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature')
    if not signature:
        abort(400, description="Missing X-Line-Signature header")

    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400, description="Invalid signature")

    return 'OK', 200

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    message_text = event.message.text

    # Patterns to extract data
    patterns = {
        'Date': r":: Created\s*>>\s*(\d{2}-\w{3}-\d{2})",
        'Job ID': r"JOB\s*:\s*(\S+)",
        'CM Team': r"CM-(\w+)",
        'Tel': r"Tel\s*:\s*(\d{10,13})",
        'No. CM': r":: ใบงานที่\s*(\d+)",
        'Priority Job': r":: priority\s*:\s*(\w+)",
        'NOC time assign': r":: Call Center C-Fiber Assign\s*>>\s*(\d{2}-\w{3}-\d{2} \d{2}[:.]\d{2})",
        'Accept time': r":: Accept\s*:\s*(\d{2}[.:]\d{2})",
        'No Alarm time': r":: No Alarm\s*:\s*(\d{2}[:.]\d{2})",
        'Detail down': r":: สาเหตุเสีย\s*:\s*(.+)",
        'CM Zone': r"Zone\s*:\s*(\w+)",
        'Material used': r":: รายการ Material ที่ใช้งาน\s*(.*?)(?=\n\s*\n|$)"
    }

    job_data = {key: re.search(pattern, message_text).group(1) if re.search(pattern, message_text) else "N/A" for key, pattern in patterns.items()}

    if job_data['Job ID'] != "N/A":
        try:
            all_values = sheet.get_all_values()
            job_id_to_row = {row[0]: idx + 2 for idx, row in enumerate(all_values[1:])}

            if job_data['Job ID'] in job_id_to_row:
                row_number = job_id_to_row[job_data['Job ID']]
                sheet.update(f'A{row_number}:L{row_number}', [[job_data[key] for key in patterns]])
            else:
                sheet.append_row([job_data[key] for key in patterns])
        except Exception as e:
            print(f"Error updating Google Sheet: {e}")

if __name__ == "__main__":
    app.run(port=8888, debug=True)

