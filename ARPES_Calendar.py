import datetime
import os.path
import htmdec_formats
import os
import time
import pytz

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar"]


def main(wavenote_file = None):
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

    if wavenote_file:
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
                    end_date = m_ti.split(' ')[0] + '-' + m_ti.split(' ')[1] + '-' + m_ti.split(' ')[3]
                end = [end_date, end_time]

    timezone = pytz.timezone('America/New_York')
    # Current time in the given timezone
    current_time = datetime.datetime.now(timezone)
    # Check if daylight saving time is in effect
    is_dst = bool(current_time.dst())  # .dst() returns a timedelta, so we convert to boolean
    if is_dst:
        offset = '05:00'
    else:
        offset = '04:00'
    
    final_start = start[0] + 'T' + start[1] + '-' + offset
    final_end = end[0] + 'T' + end[1] + '-' + offset

    try:
        service = build("calendar", "v3", credentials=creds)

        event = {
        'summary': 'ARPES User',
        'location': '343 Campus Road, Ithaca, NY 14853',
        'description': 'PARADIM user used ARPES during this time.',
        'start': {
            'dateTime': final_start,
            'timeZone': 'America/New_York',
        },
        'end': {
            'dateTime': final_end,
            'timeZone': 'America/New_York',
        }
        }

        event = service.events().insert(calendarId='primary', body=event).execute()
        print('Event created: %s' % (event.get('htmlLink')))

    except HttpError as error:
        print(f"An error occurred: {error}")


if __name__ == "__main__":
  main()