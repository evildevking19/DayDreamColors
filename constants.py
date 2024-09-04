import os
from selenium import webdriver

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

NEW_POST_URL = "https://daydreamcolors.com/wp-admin/post-new.php"
USERNAME = "Daydream Designer"
PASSWORD = ""

OPENAI_KEY = "<Your Api Key>"
CHATGPT_PROMPT = "write an opening paragraph of 100 words for a page with 4 coloring pages from the title"

def getGoogleService(service_name, version):
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    
    credential = None
    if os.path.exists('token.json'):
        credential = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not credential or not credential.valid:
        if credential and credential.expired and credential.refresh_token:
            credential.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            credential = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(credential.to_json())
            
    try:
        service = build(service_name, version, credentials=credential)
        return service
    except HttpError as err:
        print(err)
        return None
    
def getGoogleDriver():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless")
    # chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--log-level=3")
    driver = webdriver.Chrome(options=chrome_options)

    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source":
            "const newProto = navigator.__proto__;"
            "delete newProto.webdriver;"
            "navigator.__proto__ = newProto;"
    })
    return driver
