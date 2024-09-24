import datetime
import os.path
import htmdec_formats
import os
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"]


def main(wavenote_file):
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    count = 0
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json", SCOPES
        )
        creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    file_folder = os.listdir(os.path.dirname(wavenote_file)) 
    for name in file_folder:
        if not '.pxt' in name:
            file_folder.remove(name)
    sorted_files = sorted(file_folder, key=lambda x: int(x.split('_')[1].split('.')[0]))
    length = len(sorted_files)
    for file in sorted_files:
        if count == 0:
            data_file = os.path.join(os.path.dirname(wavenote_file), "data_holder.txt")
            dataset = htmdec_formats.ARPESDataset.from_file(wavenote_file)
            with open(data_file, "w") as f:
                l = f.write(dataset._metadata)
            with open(data_file, 'r') as file:
                lines = file.readlines()
            os.remove(data_file)
            start = [lines[28].split('=')[1],lines[29].split('=')[1]]
        elif count == length - 1:
            ti_m = os.path.getmtime(wavenote_file)
            m_ti = time.ctime(ti_m) 
            if ':' in m_ti.split(' ')[3]:
                end_time = m_ti.split(' ')[3]
            else:
                end_time = m_ti.split(' ')[4]
            if ':' in m_ti.split(' ')[3]:
                end_date = m_ti.split(' ')[0] + '-' + m_ti.split(' ')[1] + '-' + m_ti.split(' ')[2]
            else:
                end_date = m_ti.split(' ')[4]
            end = []
    try:
        service = build("calendar", "v3", credentials=creds)

        # Call the Calendar API
        now = datetime.datetime.utcnow().isoformat() + "Z"  # 'Z' indicates UTC time
        print("Getting the upcoming 10 events")
        events_result = (
            service.events()
            .list(
                calendarId="primary",
                timeMin=now,
                maxResults=10,
                singleEvents=True,
                orderBy="startTime",
            )
            .execute()
        )
        events = events_result.get("items", [])

        if not events:
            print("No upcoming events found.")
            return

        # Prints the start and name of the next 10 events
        for event in events:
            start = event["start"].get("dateTime", event["start"].get("date"))
        print(start, event["summary"])

    except HttpError as error:
        print(f"An error occurred: {error}")


if __name__ == "__main__":
  main()