"""
Created on April 19, 2022
@author: Muntasib Muhib Chowdhury (muntasibchowdhury@gmail.com)
"""

def install_missing_modules():
    """
    Purpose: This function will check all the installed modules
    If anything is missing, it will install them only
    """
    try:
        import platform
        import traceback
        import subprocess
        import sys
        try:
            from pip._internal.operations import freeze
        except ImportError:
            # pip < 10.0
            from pip.operations import freeze
        # NULL output device for disabling print output of pip installs
        try:
            from subprocess import DEVNULL  # py3k
        except ImportError:
            import os
            DEVNULL = open(os.devnull, 'wb')

        req_list = ["requests", "colorama", "python-dateutil", "pytz"]
        freeze_list = freeze.freeze()
        alredy_installed_list = []
        for p in freeze_list:
            name = p.split("==")[0]
            if "@" not in name:
                # '@' symbol appears in some python modules in Windows
                alredy_installed_list.append(str(name).lower())

        # installing any missing modules
        installed = False
        error = False
        for module_name in req_list:
            if module_name.lower() not in alredy_installed_list:
                try:
                    print("module_installer: Installing module: %s" % module_name)
                    subprocess.check_call([sys.executable, "-m", "pip", "install", "--trusted-host=pypi.org", "--trusted-host=files.pythonhosted.org", module_name], stderr=DEVNULL, stdout=DEVNULL, )
                    print("module_installer: Installed missing module: %s" % module_name)
                    installed = True
                except:
                    print("module_installer: Failed to install module: %s" % module_name)
                    error = True

        if not installed and not error:
            print("All required modules are already installed. Continuing...\n")

        try:
            from google.auth.transport.requests import Request
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import InstalledAppFlow
            from googleapiclient.discovery import build
            from googleapiclient.errors import HttpError
        except:
            module_name = "google-api-python-client google-auth-httplib2 google-auth-oauthlib"
            print("module_installer: Installing module: Google API")
            subprocess.check_output("pip install --trusted-host=pypi.org --trusted-host=files.pythonhosted.org --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib")
            print("module_installer: Installed missing module: Google API")
    except:
        traceback.print_exc()
        print("Failed to install missing modules...\n")


install_missing_modules()

color_info = """
colorId: 1
  Background: #a4bdfc   #ash blue
  Foreground: #1d1d1d
colorId: 2
  Background: #7ae7bf   #bottle green
  Foreground: #1d1d1d
colorId: 3
  Background: #dbadff   #light purple
  Foreground: #1d1d1d
colorId: 4
  Background: #ff887c   #pink
  Foreground: #1d1d1d
colorId: 5
  Background: #fbd75b   #yellow
  Foreground: #1d1d1d
colorId: 6
  Background: #ffb878   #orange
  Foreground: #1d1d1d
colorId: 7
  Background: #46d6db   #cyan
  Foreground: #1d1d1d
colorId: 8
  Background: #e1e1e1   #ash
  Foreground: #1d1d1d
colorId: 9
  Background: #5484ed   #blue
  Foreground: #1d1d1d
colorId: 10
  Background: #51b749   #green
  Foreground: #1d1d1d
colorId: 11
  Background: #dc2127   #red
  Foreground: #1d1d1d
"""


def decide_color(status):
    colorId = {
        "ash_blue":  {"Background": "#a4bdfc","colorId": "1", "Foreground": "#1d1d1d"},
        "bottle_green":  {"Background": "#7ae7bf","colorId": "2", "Foreground": "#1d1d1d"},
        "light_purple":  {"Background": "#dbadff","colorId": "3", "Foreground": "#1d1d1d"},
        "pink":  {"Background": "#ff887c","colorId": "pink", "4": "#1d1d1d"},
        "yellow":  {"Background": "#fbd75b","colorId": "5", "Foreground": "#1d1d1d"},
        "orange":  {"Background": "#ffb878","colorId": "6", "Foreground": "#1d1d1d"},
        "cyan":  {"Background": "#46d6db","colorId": "7", "Foreground": "#1d1d1d"},
        "ash":  {"Background": "#e1e1e1","colorId": "8", "Foreground": "#1d1d1d"},
        "blue":  {"Background": "#5484ed","colorId": "9", "Foreground": "#1d1d1d"},
        "green":  {"Background": "#51b749","colorId": "10", "Foreground": "#1d1d1d"},
        "red":  {"Background": "#dc2127","colorId": "11", "Foreground": "#1d1d1d"},
    }
    if status.strip().lower() == "confirmed": return colorId["yellow"]["colorId"]
    if status.strip().lower() == "visited": return colorId["orange"]["colorId"]
    if status.strip().lower() == "completed": return colorId["red"]["colorId"]
    else: return colorId["blue"]["colorId"]


import os, sys, json, time, signal
from datetime import datetime, timedelta
from dateutil.tz import tzlocal

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from colorama import init as colorama_init
from colorama import Fore

# Initialize colorama for the current platform
colorama_init(autoreset=True)

utc = str(datetime.now(tzlocal()).utcoffset())
utc = utc[:utc.find(":", 2)]
utc = "+" + utc if utc.startswith("0") else "+0" + utc

def Exception_Handler(exec_info):
    try:
        exc_type, exc_obj, exc_tb = exec_info
        Error_Type = (
            (str(exc_type).replace("type ", ""))
            .replace("<", "")
            .replace(">", "")
            .replace(";", ":")
        )
        Error_Message = str(exc_obj)
        File_Name = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        Function_Name = os.path.split(exc_tb.tb_frame.f_code.co_name)[1]
        Line_Number = str(exc_tb.tb_lineno)
        Error_Detail = (
            "Error Type ~ %s: Error Message ~ %s: File Name ~ %s: Function Name ~ %s: Line ~ %s"
            % (Error_Type, Error_Message, File_Name, Function_Name, Line_Number)
        )
        print(Fore.RED + "Following exception occurred: %s" % (Error_Detail))
    except:
        return


def ExecLog(sDetails, iLogLevel=4):
    if iLogLevel == 1:
        status = "Created"
        line_color = Fore.GREEN
    elif iLogLevel == 2:
        status = "Updated"
        line_color = Fore.YELLOW
    elif iLogLevel == 3:
        status = "Deleted"
        line_color = Fore.RED
    else:
        status = "Info"
        line_color = Fore.CYAN
    print(line_color + f"[{status.upper()}] {sDetails}")


service_sheet = None
service_calendar = None
creds = None


def Authenticate():
    try:
        SCOPES = [
            "https://www.googleapis.com/auth/calendar.events",
            "https://www.googleapis.com/auth/spreadsheets.readonly",
        ]
        global creds
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'client_secret.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        if creds is None:
            raise Exception("Could not authenticate")
        global service_sheet, service_calendar
        service_calendar = build("calendar", "v3", credentials=creds)
        service_sheet = build('sheets', 'v4', credentials=creds)

    except:
        return Exception_Handler(sys.exc_info())


def read_sheet():
    try:
        SAMPLE_SPREADSHEET_ID = '11JgRdq_pF16YHcdHFO1llrkttsx4hnUGINo71WY153I'
        SAMPLE_RANGE_NAME = 'Sheet1!A:L'
        # Call the Sheets API
        sheet = service_sheet.spreadsheets()
        results = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute().get('values', [])
        if not results:
            raise Exception("No data found in Sheet")

        max_len = max(*[len(i) for i in results])
        results = [row + [""] * (max_len - len(row)) for row in results]
        list_of_dict = []
        for row in results[1:]:
            list_of_dict.append({key: val for key, val in list(zip(*[results[0], row]))})

        i = 0
        while i < len(list_of_dict):
            if not list_of_dict[i]["unique site number"].strip():
                del list_of_dict[i]
                continue
            if not list_of_dict[i]["date of visit"].strip():
                list_of_dict[i]["date of visit"] = "01/01/2020"
            if not list_of_dict[i]["time of visit"].strip():
                list_of_dict[i]["time of visit"] = "12.00 am"
            if not list_of_dict[i]["visit duration"].strip():
                list_of_dict[i]["visit duration"] = "1 hour"
            i += 1
        results = [[key for key in list_of_dict[0]]]
        for row in list_of_dict:
            results.append([row[key] for key in row])


        ExecLog(f"Found {len(results)-1} rows in sheet")
        return results, list_of_dict
    except:
        Exception_Handler(sys.exc_info())
        raise Exception("Could not parse Sheet Data")


def read_calendar(results):
    try:
        dictionary = {key: val for key, *val in list(zip(*results))}
        datetimes = sorted([datetime.strptime(" ".join(i), '%d/%m/%Y %I.%M %p') for i in list(zip(dictionary["date of visit"], dictionary["time of visit"]))])
        maxtime = (datetimes[-1] + timedelta(seconds=1)).strftime("%Y-%m-%dT%H:%M:%S") + utc
        mintime = datetimes[0].strftime("%Y-%m-%dT%H:%M:%S") + utc
        resultc = service_calendar.events().list(calendarId='primary', timeMax=maxtime, timeMin=mintime).execute().get('items', [])
        ExecLog(f"Found {len(resultc)} events from {mintime} to {maxtime}")
        return resultc
    except:
        Exception_Handler(sys.exc_info())
        return []


def create_event(row):
    try:
        if row["Status"].strip().lower() == "cancelled":
            return
        start_time = datetime.strptime(row["date of visit"] + " " + row["time of visit"], '%d/%m/%Y %I.%M %p')
        duration = int(row["visit duration"][:row["visit duration"].find(" ")])
        end_time = (start_time + timedelta(hours=duration)).strftime("%Y-%m-%dT%H:%M:%S") + utc
        start_time = start_time.strftime("%Y-%m-%dT%H:%M:%S") + utc
        sheet_event = {
            'summary': row["unique site number"],
            'location': row["site address"],
            'start': {
                'dateTime': start_time,
                # 'timeZone': 'America/Los_Angeles',
            },
            'end': {
                'dateTime': end_time,
                # 'timeZone': 'America/Los_Angeles',
            },
            "colorId": decide_color(row["Status"]),
        }

        event = service_calendar.events().insert(calendarId='primary', body=sheet_event).execute()
        ExecLog(f"Event {sheet_event}", 1)
        row["eventId"] = event["id"]
    except:
        return Exception_Handler(sys.exc_info())


def update_calendar_event(row, ev):
    try:
        if row["Status"].strip().lower() == "cancelled":
            return delete_event_from_row(row)
        start_time = datetime.strptime(row["date of visit"] + " " + row["time of visit"], '%d/%m/%Y %I.%M %p')
        duration = int(row["visit duration"][:row["visit duration"].find(" ")])
        end_time = (start_time + timedelta(hours=duration)).strftime("%Y-%m-%dT%H:%M:%S") + utc
        start_time = start_time.strftime("%Y-%m-%dT%H:%M:%S") + utc

        sheet_event = {
            'summary': row["unique site number"],
            'location': row["site address"],
            'start': {
                'dateTime': start_time,
                # 'timeZone': 'America/Los_Angeles',
            },
            'end': {
                'dateTime': end_time,
                # 'timeZone': 'America/Los_Angeles',
            },
            "colorId": decide_color(row["Status"]),
        }
        cal_event = {
            'summary': ev["summary"],
            'location': ev["location"],
            'start': {
                'dateTime': ev["start"]["dateTime"],
                # 'timeZone': 'America/Los_Angeles',
            },
            'end': {
                'dateTime': ev["end"]["dateTime"],
                # 'timeZone': 'America/Los_Angeles',
            },
            "colorId": ev["colorId"] if "colorId" in ev else "9",
        }

        if sheet_event != cal_event:
            updated_event = service_calendar.events().update(calendarId='primary', eventId=ev["id"], body=sheet_event).execute()
            ExecLog(f"Event {sheet_event}", 2)
    except:
        return Exception_Handler(sys.exc_info())


def delete_event_from_row(row):
    try:
        service_calendar.events().delete(calendarId='primary', eventId=row["eventId"]).execute()
        ExecLog(f'unique site number = {row["unique site number"]} eventId = {row["eventId"]}', 3)
    except:
        return Exception_Handler(sys.exc_info())


def sanitize_data(results):
    try:
        list_of_dict = []
        for row in results[1:]:
            list_of_dict.append({key: val for key, val in list(zip(*[results[0], row]))})
    
        c = 0
        for row in list_of_dict:
            if not row["unique site number"].strip():
                del list_of_dict[c]
            if not row["date of visit"].strip():
                row["date of visit"] = "01/01/2020"
            if not row["time of visit"].strip():
                row["time of visit"] = "12.00 am"
            if not row["visit duration"].strip():
                row["visit duration"] = "1 hour"
            c += 1
        return list_of_dict
    except:
        Exception_Handler(sys.exc_info())
        raise Exception("Could not parse sheet data")
    

def main():
    try:
        Authenticate()
        while True:
            try:
                results, list_of_dict = read_sheet()
                resultc = read_calendar(results)

                for row in list_of_dict:
                    for ev in resultc:
                        if "summary" in ev and row["unique site number"].strip() == ev["summary"].strip():
                            row["eventId"] = ev["id"]  # If there is already an event is found against a row save the event id
                            break
                    else:
                        create_event(row)  # If no event found against a row create an event

                if os.path.isfile("calendar_sheet_map.json"):      # If there is previously data saved
                    with open("calendar_sheet_map.json", "r") as f:
                        json_data = json.loads(f.read())
                    for j_row in json_data:
                        for row in list_of_dict:
                            if row["unique site number"].strip() == j_row["unique site number"].strip():
                                break
                        else:       # If there is a sheet row does not match with previous json_data that means it should be deleted
                            delete_event_from_row(j_row)

                # If any update found in sheet row update it to calendar
                for row in list_of_dict:
                    for ev in resultc:
                        if "summary" in ev and row["unique site number"].strip() == ev["summary"].strip():
                            update_calendar_event(row, ev)
                    else:
                        pass

                with open("calendar_sheet_map.json", "w") as f:
                    f.write(json.dumps(list_of_dict, indent=2))
                print()
                time.sleep(2)
            except:
                Exception_Handler(sys.exc_info())
    except:
        return Exception_Handler(sys.exc_info())


def signal_handler(sig, frame):
    print("Disconnecting from server...")
    os._exit(1)


if __name__ == '__main__':
    signal.signal(signal.SIGINT, signal_handler)
    main()
    pass