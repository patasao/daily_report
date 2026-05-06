import time
import requests
import logging
from services.ms_auth import TokenManager
import config

logger = logging.getLogger(__name__)

class Email:
    def __init__(self):
        self.access_token = TokenManager(config.credentials['client_id'],
                                         config.credentials['authority'],
                                         config.credentials['client_secret'],
                                         config.credentials['scope'])
        self.endpoint = config.credentials['graph_endpoint']

    def send(self, sender_email:str, message:dict):
        """Method to send email via Graph API"""
        max_retries = 3
        delay_seconds = 5

        for attempt in range(1, max_retries + 1):
            try:
                token = self.access_token.get_token()
                headers = {
                    "Authorization": f"Bearer {token}", 
                    "Content-Type": "application/json"
                    }
                url = f"{self.endpoint}/users/{sender_email}/sendMail"

                try:
                    response = requests.post(url, headers=headers, json=message)

                    if response.status_code == 202:
                        logging.info(f"Email sent successfully")
                        return True
                    else:
                        raise Exception(f"Graph API error {response.status_code}: {response.text}")
                    
                except Exception as send_error:
                    logging.warning(f"---Attempt {attempt}/{max_retries}: Error sending email: {send_error}")
                    if attempt < max_retries:
                        time.sleep(delay_seconds)
                
                logging.error(f"---Failed to send email after {max_retries} attempts")
                return False

            except Exception as e:
                logging.error(f"---Error drafting email: {e}", exc_info=True)
                return False

# # --- Example message template (can add or remove elements) ---
# raw_recipients = "pataraporn.s@compasshospitality.com"
# recipients = [
#     {"emailAddress": {"address": email.strip()}}
#     for email in raw_recipients.split(";")
#     if email.strip()
# ]

# message = {
#     "message": {
#         "subject": "subject",
#         "body": {"contentType": "HTML", "content": "email_body"},
#         "toRecipients": recipients,
        # "attachments": [
        #     {
        #         "@odata.type": "#microsoft.graph.fileAttachment",
        #         "name": "table.png",
        #         "contentBytes": img_base64,
        #         "contentId": "table1",
        #         "isInline": True,
        #     }
        # ],
#     },
#     "saveToSentItems": "true",
# }

