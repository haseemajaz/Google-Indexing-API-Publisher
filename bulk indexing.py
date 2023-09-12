from oauth2client.service_account import ServiceAccountCredentials
import httplib2
import json
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# https://developers.google.com/search/apis/indexing-api/v3/prereqs#header_2
SCOPES = ["https://www.googleapis.com/auth/indexing"]

def indexURL(urls, http):
    ENDPOINT = "https://indexing.googleapis.com/v3/urlNotifications:publish"
    
    for u in urls:
        content = {}
        content['url'] = u.strip()
        content['type'] = "URL_UPDATED"
        json_ctn = json.dumps(content)
    
        response, content = http.request(ENDPOINT, method="POST", body=json_ctn)

        result = json.loads(content.decode())

        if "error" in result:
            print("Error({} - {}): {}".format(result["error"]["code"], result["error"]["status"], result["error"]["message"]))
        else:
            print("urlNotificationMetadata.url: {}".format(result["urlNotificationMetadata"]["url"]))
            print("urlNotificationMetadata.latestUpdate.url: {}".format(result["urlNotificationMetadata"]["latestUpdate"]["url"]))
            print("urlNotificationMetadata.latestUpdate.type: {}".format(result["urlNotificationMetadata"]["latestUpdate"]["type"]))
            print("urlNotificationMetadata.latestUpdate.notifyTime: {}".format(result["urlNotificationMetadata"]["latestUpdate"]["notifyTime"]))

def browse_files():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Ask for the JSON key file
    json_file_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])

    if json_file_path:
        # Ask for the XLSX file containing URLs
        xlsx_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

        if xlsx_file_path:
            JSON_KEY_FILE = json_file_path
            df = pd.read_excel(xlsx_file_path)
            urls = df["URL"].tolist()
            credentials = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scopes=SCOPES)
            http = credentials.authorize(httplib2.Http())
            indexURL(urls, http)

if __name__ == "__main__":
    browse_files()
