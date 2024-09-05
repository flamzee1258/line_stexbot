from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError, LineBotApiError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
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
Channel_secret = os.getenv('LINE_CHANNEL_SECRET', 'default_secret')
Channel_access_token = os.getenv('LINE_CHANNEL_ACCESS_TOKEN', 'default_access_token')

line_bot_api = LineBotApi(Channel_access_token)
handler = WebhookHandler(Channel_secret)

# Set up Google Sheets connection
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Load JSON credentials from environment variable
json_creds = os.getenv('GOOGLE_CREDENTIALS')
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
    date_pattern = r":: Created\s*>>\s*(\d{2}-\w{3}-\d{2})"
    job_id_pattern = r"JOB\s*:\s*(\S+)"
    cm_team_pattern = r"CM-(\w+)"
    tel_pattern = r"Tel\s*:\s*(\d{10,13})"
    No_cm_pattern = r":: ใบงานที่\s*(\d+)"
    priority_job_pattern = r":: priority\s*:\s*(\w+)"
    time_assign_pattern = r":: Call Center C-Fiber Assign\s*>>\s*(\d{2}-\w{3}-\d{2} \d{2}[:.]\d{2})"
    accept_time_pattern = r":: Accept\s*:\s*(\d{2}[.:]\d{2})"
    noalarm_time_pattern = r":: No Alarm\s*:\s*(\d{2}[:.]\d{2})"
    detail_down_pattern = r":: สาเหตุเสีย\s*:\s*(.+)"
    zone_list_pattern = r"Zone\s*:\s*(\w+)"
    material_used_pattern = r":: รายการ Material ที่ใช้งาน\s*(.*?)(?=\n\s*\n|$)"

    date_match = re.search(date_pattern, message_text)
    job_id_match = re.search(job_id_pattern, message_text)
    cm_team_match = re.search(cm_team_pattern, message_text)
    tel_match = re.search(tel_pattern, message_text)
    No_cm_match = re.search(No_cm_pattern, message_text)
    priority_job_match = re.search(priority_job_pattern, message_text)
    time_assign_match = re.search(time_assign_pattern, message_text)
    accept_time_match = re.search(accept_time_pattern, message_text)
    noalarm_time_match = re.search(noalarm_time_pattern, message_text)
    detail_down_match = re.search(detail_down_pattern, message_text)
    zone_list_match = re.search(zone_list_pattern, message_text)
    material_used_match = re.search(material_used_pattern, message_text, re.DOTALL)

    if job_id_match:
        job_id = job_id_match.group(1)
        job_data = {
            'Date': date_match.group(1) if date_match else "N/A",
            'CM Team': cm_team_match.group(1) if cm_team_match else "N/A",
            'Tel': tel_match.group(1) if tel_match else "N/A",
            'No. CM': No_cm_match.group(1) if No_cm_match else "N/A",
            'Priority Job': priority_job_match.group(1) if priority_job_match else "N/A",
            'NOC time assign': time_assign_match.group(1) if time_assign_match else "N/A",
            'Accept time': accept_time_match.group(1) if accept_time_match else "N/A",
            'No Alarm time': noalarm_time_match.group(1) if noalarm_time_match else "N/A",
            'Detail down': detail_down_match.group(1) if detail_down_match else "N/A",
            'CM Zone': zone_list_match.group(1) if zone_list_match else "N/A",
            'Material used': material_used_match.group(1) if material_used_match else "N/A"
        }

        try:
            all_values = sheet.get_all_values()
            header = all_values[0]
            rows = all_values[1:]

            # Create a map of job_id to row number
            job_id_to_row = {row[0]: idx + 2 for idx, row in enumerate(rows)}

            if job_id in job_id_to_row:
                row_number = job_id_to_row[job_id]
                existing_row = sheet.row_values(row_number)

                # Prepare updated row data
                row_data = [
                    job_id,
                    job_data.get('Date', existing_row[1]),
                    job_data.get('CM Team', existing_row[2]),
                    job_data.get('Tel', existing_row[3]),
                    job_data.get('No. CM', existing_row[4]),
                    job_data.get('Priority Job', existing_row[5]),
                    job_data.get('NOC time assign', existing_row[6]),
                    job_data.get('Accept time', existing_row[7]),
                    job_data.get('No Alarm time', existing_row[8]),
                    job_data.get('Detail down', existing_row[9]),
                    job_data.get('CM Zone', existing_row[10]),
                    job_data.get('Material used', existing_row[11])
                ]

                # Update row in Google Sheet
                sheet.update(f'A{row_number}:L{row_number}', [row_data])
            else:
                # Append new row if job_id not found
                last_row = len(sheet.get_all_values()) + 1
                row_data = [
                    job_id,
                    job_data['Date'],
                    job_data['CM Team'],
                    job_data['Tel'],
                    job_data['No. CM'],
                    job_data['Priority Job'],
                    job_data['NOC time assign'],
                    job_data['Accept time'],
                    job_data['No Alarm time'],
                    job_data['Detail down'],
                    job_data['CM Zone'],
                    job_data['Material used']
                ]
                sheet.insert_row(row_data, last_row)

            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=f"Job data for {job_id} has been processed and saved.")
            )
        except gspread.exceptions.APIError as e:
            print(f"Google Sheets API error: {e}")
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="เกิดข้อผิดพลาดในการเชื่อมต่อกับ Google Sheets API.")
            )
        except LineBotApiError as e:
            print(f"LINE Bot API error: {e}")
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="เกิดข้อผิดพลาดในการตอบกลับไปที่ LINE.")
            )
        except Exception as e:
            print(f"Unexpected error: {e}")
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="เกิดข้อผิดพลาดที่ไม่คาดคิด โปรดลองอีกครั้ง.")
            )

if __name__ == "__main__":
    app.run(port=8888, debug=True)
