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


def display_result(SEARCH_PARAMETERS, list_of_dict, Return="site address"):
    result = [row for row in list_of_dict]
    for param in SEARCH_PARAMETERS:
        result = [row for row in result if row[param].strip().lower() == SEARCH_PARAMETERS[param].strip().lower() and row["Status"].strip().lower() != "cancelled"]
    result = set([row[Return] for row in result])
    ExecLog("Following are the search results")
    for val in result:
        print(Fore.YELLOW + val)

def main():
    try:
        Authenticate()
        results, list_of_dict = read_sheet()
        SEARCH_PARAMETERS = {
            # "date of visit": "01/01/2022",
            "Region": "Sydney",
        }
        display_result(SEARCH_PARAMETERS, list_of_dict, Return="site address")

    except:
        return Exception_Handler(sys.exc_info())


def signal_handler(sig, frame):
    print("Disconnecting from server...")
    os._exit(1)


if __name__ == '__main__':
    signal.signal(signal.SIGINT, signal_handler)
    main()
    pass