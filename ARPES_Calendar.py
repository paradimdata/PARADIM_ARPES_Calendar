import datetime
import os.path
import htmdec_formats
import os
import time
import pytz
import argparse
import re

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar"]


def month_to_num(month_abbr):
    # Dictionary mapping month abbreviations to their corresponding numbers
    month_map = {
        "Jan": '01',
        "Feb": '02',
        "Mar": '03',
        "Apr": '04',
        "May": '05',
        "Jun": '06',
        "Jul": '07',
        "Aug": '08',
        "Sep": '09',
        "Oct": '10',
        "Nov": '11',
        "Dec": '12'
    }
    month_abbr = month_abbr.capitalize()
    return month_map.get(month_abbr, None)

def get_iso8601(year, month, day, hour, minute, second, timezone_str):
    timezone = pytz.timezone(timezone_str)
    dt = datetime.datetime(year, month, day, hour, minute, second)
    localized_dt = timezone.localize(dt)
    return localized_dt.isoformat()

def main(wavenote_file = None):
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    count = 0
    final_end = ''
    final_start = ''
    
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        try:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                raise ValueError()
        except:
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
        sorted_files = sorted(
        file_folder,
        key=lambda x: float(re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", x)[-1]) if re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", x) else 0
        ) 
        length = len(sorted_files)
        for file in sorted_files:
            if count == 0:
                dataset = htmdec_formats.ARPESDataset.from_file(wavenote_file)
                lines = dataset._metadata.split("\n")
                folder = wavenote_file + '/../../../../..'
                path = os.path.normpath(folder)
                directory_name = os.path.basename(path.rstrip('/\\'))
                name_array = directory_name.split()
                for element in name_array:
                    if element.isdigit():
                        project_number = element
                start = [lines[28].split('=')[1].replace('\n',''),lines[29].split('=')[1].replace('\n','')]
                if project_number:
                    username = lines[25].split('=')[1] + ' (#' + str(project_number) + ') ARPES'
                else:
                    username = lines[25].split('=')[1] + ' (#) ARPES'
                instrument = lines[23].split('=')[1] + 'was used over this time period'
            elif count == length - 1:
                end_file = os.path.join(os.path.dirname(wavenote_file), file)
                ti_m = os.path.getmtime(end_file)
                m_ti = time.ctime(ti_m) 
                if ':' in m_ti.split(' ')[3]:
                    end_time = m_ti.split(' ')[3]
                else:
                    end_time = m_ti.split(' ')[4]
                if ':' in m_ti.split(' ')[3]:
                    end_date = m_ti.split(' ')[4] + '-' + month_to_num(m_ti.split(' ')[1]) + '-' + m_ti.split(' ')[2]
                else:
                    end_date = m_ti.split(' ')[5] + '-' + month_to_num(m_ti.split(' ')[1]) + '-' + m_ti.split(' ')[3]
                end = [end_date, end_time]
            count += 1
        
        final_start = get_iso8601(int(start[0].split('-')[0]), int(start[0].split('-')[1]), int(start[0].split('-')[2]), 
                                  int(start[1].split(':')[0]), int(start[1].split(':')[1]), int(start[1].split(':')[2]), 'America/New_York')
        final_end = get_iso8601(int(end[0].split('-')[0]), int(end[0].split('-')[1]), int(end[0].split('-')[2]), 
                                int(end[1].split(':')[0]), int(end[1].split(':')[1]), int(end[1].split(':')[2]), 'America/New_York')

    try:
        service = build("calendar", "v3", credentials=creds)

        event = {
        'summary': username,
        'location': '343 Campus Road, Ithaca, NY 14853',
        'description': instrument,
        'start': {
            'dateTime': final_start,
            'timeZone': 'America/New_York',
        },
        'end': {
            'dateTime': final_end,
            'timeZone': 'America/New_York',
        }
        }
        event = service.events().insert(calendarId='7262038e2634deb88fae6c4900df5cc42df1f06f06522f3e3fd43a8bc7e4c10a@group.calendar.google.com', body=event).execute()
        print('Event created: %s' % (event.get('htmlLink')))

    except HttpError as error:
        print(f"An error occurred: {error}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="ARPES Calendar Event Creator")
    # Adding command-line arguments
    parser.add_argument('--wavenote', type=str, required=True, help="Path to the wavenote file")
    args = parser.parse_args()
    # Call the main function with parsed arguments
    main(args.wavenote)