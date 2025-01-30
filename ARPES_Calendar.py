import datetime
import traceback
import os.path
import htmdec_formats
import os
import time
import pytz
import argparse
import re
import sys
from dateutil import parser

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

def get_calendar_values(wavenote_file):
    project_number = None
    index = 0
    full_path = os.path.abspath(wavenote_file)
    file_folder = os.listdir(os.path.dirname(full_path))

    for name in file_folder:
        if not '.pxt' in name:
            file_folder.remove(name)
    sorted_files = sorted(
    file_folder,
    key=lambda x: float(re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", x)[-1]) if re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", x) else 0
    ) 

    start = []
    end = []
    for file in sorted_files:
        dataset = htmdec_formats.ARPESDataset.from_file(os.path.dirname(full_path) + '/' + file)
        lines = dataset._metadata.split("\n")

        if index == 0:
            folder = wavenote_file + '/../../../../..'
            path = os.path.normpath(folder)
            directory_name = os.path.basename(path.rstrip('/\\'))
            name_array = directory_name.split()

            for element in name_array:
                if element.isdigit():
                    project_number = element
            if project_number:
                username = lines[25].split('=')[1] + ' (#' + str(project_number) + ') ARPES'
            else:
                username = lines[25].split('=')[1] + ' (#) ARPES'
            instrument = lines[23].split('=')[1] + 'was used over this time period'

        start.append([lines[28].split('=')[1].replace('\n',''),lines[29].split('=')[1].replace('\n','')])
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
        end.append([end_date, end_time])
        index += 1
    
    final_start = []
    final_end = []
    count = 0
    while count < len(start):
        final_start.append(get_iso8601(int(start[count][0].split('-')[0]), int(start[count][0].split('-')[1]), int(start[count][0].split('-')[2]), 
                            int(start[count][1].split(':')[0]), int(start[count][1].split(':')[1]), int(start[count][1].split(':')[2]), 'America/New_York'))
        final_end.append(get_iso8601(int(end[count][0].split('-')[0]), int(end[count][0].split('-')[1]), int(end[count][0].split('-')[2]), 
                            int(end[count][1].split(':')[0]), int(end[count][1].split(':')[1]), int(end[count][1].split(':')[2]), 'America/New_York'))
        count += 1
    
    return username, instrument, final_start, final_end
    
def input_arpes_event(username, instrument, final_start, final_end, creds):
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

def get_calendar_events(creds, past_time):
    service = build("calendar", "v3", credentials=creds)

    # Convert `past_time` to ISO 8601 format
    if past_time.tzinfo is not None:  # Check if datetime is timezone-aware
        now = past_time.astimezone(datetime.timezone.utc).isoformat().replace('+00:00', 'Z')
    else:
        now = past_time.isoformat() + 'Z'

    events_result = service.events().list(
        calendarId='7262038e2634deb88fae6c4900df5cc42df1f06f06522f3e3fd43a8bc7e4c10a@group.calendar.google.com',
        timeMin=now,
        maxResults=2500,
        singleEvents=True,
        orderBy='startTime'
    ).execute()

    events = events_result.get('items', [])

    if not events:
        return []

    event_times = []

    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        end = event['end'].get('dateTime', event['end'].get('date'))
        summary = event.get('summary', 'No Title')
        event_times.append({
            'summary': summary,
            'start': start,
            'end': end
        })

    return event_times

def parse_datetime(dt_str):
    try:
        # Use dateutil.parser for robust ISO 8601 parsing
        dt = parser.isoparse(dt_str)
    except (ValueError, ImportError):
        # Handle basic date format (e.g., 'YYYY-MM-DD') if `isoparse` fails
        try:
            dt = datetime.strptime(dt_str, '%Y-%m-%d')
        except ValueError:
            raise ValueError(f"Unsupported datetime format: {dt_str}")
    
    # Normalize to offset-naive if timezone info is present
    return dt.replace(tzinfo=None)

def gather_and_insert_arpes_event(username, instrument, final_start, final_end, creds):
    start_holder = None
    end_holder = None
    for i in range(len(final_start)):

        if not start_holder:
            start_holder = final_start[i]
        if not end_holder:
            end_holder = final_end[i]
        
        dt1 = datetime.datetime.fromisoformat(final_start[i])
        dt2 = datetime.datetime.fromisoformat(end_holder)
        time_difference = abs(dt1 - dt2)
        difference_seconds = time_difference.total_seconds()

        if difference_seconds < 86400 and i != len(final_start) - 1:
            end_holder = final_end[i]
        elif difference_seconds < 86400 and i == len(final_start) - 1:
            input_arpes_event(username, instrument, start_holder, final_end[i], creds)
        elif difference_seconds > 86400 and i != len(final_end) - 1:
            input_arpes_event(username, instrument, start_holder, end_holder, creds)
            start_holder = final_start[i]
            end_holder = final_end[i]
        else:
            input_arpes_event(username, instrument, start_holder, end_holder, creds)
            input_arpes_event(username, instrument, final_start[i], final_end[i], creds)

def duplicate_check(data, time1,time2):
    # Parse the times to check
    time1_dt = parse_datetime(time1)
    time2_dt = parse_datetime(time2)

    # Variable to track if both times are in the same event
    common_event = None

    # Check each event
    for event in data:
        # Parse start and end times
        start_dt = parse_datetime(event['start'])
        end_dt = parse_datetime(event['end'])

        # Check if both times fall within the same event
        if start_dt == time1_dt and time2_dt == end_dt:
            return False
    return True

def main(wavenote_file=None, wavenote_folder=None):
    creds = None
    final_end = ''
    final_start = ''
    
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
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
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    if wavenote_file:
        username, instrument, final_start, final_end = get_calendar_values(wavenote_file)
        gather_and_insert_arpes_event(username, instrument, final_start, final_end, creds)
            

    elif wavenote_folder:
        # Correct initialization of `past_time` as a datetime object
        past_time = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=365)
        
        # No need for `.endswith()`, handle string conversion in `get_calendar_events`
        past_events = get_calendar_events(creds, past_time)
        filetypes = ['.pxt']
        filtered_files = []
        dir = wavenote_folder

        for current_folder, subfolders, files in os.walk(dir):
            # Get the full paths of the matching files
            matching_files = [os.path.join(current_folder, f) for f in files if any(f.endswith(filetype) for filetype in filetypes)]
            if matching_files:  # Only add non-empty lists
                filtered_files.append((current_folder, matching_files))

        for file in filtered_files:
            try:
                username, instrument, final_start, final_end = get_calendar_values(os.path.abspath(file[1][0]))
                if duplicate_check(past_events, final_start[0], final_end[-1]):
                    gather_and_insert_arpes_event(username, instrument, final_start, final_end, creds)
            except Exception as e:
                print(f"Error processing file {file[1][0]}: {e}")
                traceback.print_exc()  # Print the full traceback

if __name__ == "__main__":
    # Rename the ArgumentParser object to avoid conflict
    arg_parser = argparse.ArgumentParser(description="ARPES Calendar Event Creator")
    
    # Adding command-line arguments
    group = arg_parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--wavenote', type=str, help="Path to the wavenote file")
    group.add_argument('--folder', type=str, help="Path to the directory holding wavenote files")
    
    args = arg_parser.parse_args()

    # Call the main function with parsed arguments
    main(wavenote_file=args.wavenote, wavenote_folder=args.folder)