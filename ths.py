import os.path
import subprocess
from pprint import pformat
from datetime import datetime, timedelta
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import json
from screeninfo import get_monitors
import cachetools.func
import psutil
import logging
import time
import signal
from typing import *
logger = logging.getLogger('ths')



def init_logging():
    
    logger.setLevel(logging.DEBUG)
    # create console handler and set level to debug
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)

    # create formatter
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    # add formatter to ch
    ch.setFormatter(formatter)
    # add ch to logger
    logger.addHandler(ch)

def kill_program(name: str):
    for proc in psutil.process_iter():
        if name in proc.name():
            logger.info(f"{name} is running, killing it")
            proc.kill()

def start_firefox():
    kill_program("firefox-bin")
    logger.info("Opening Firefox")
    subprocess.Popen(["firefox", 
        "https://www.notion.so/a939aebe271a4ac3b0c5b0a58f92ed9a"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
    
def calc_duration(duration: str)->int:
    duration =  (datetime.strptime(duration, "%H:%M:%S") - datetime.strptime("0:00:00", "%H:%M:%S")).seconds
    return duration

class HubStaff:
    CLIENT = 'HubstaffClient.bin.x86_64'
    CLI = 'HubstaffCLI.bin.x86_64'
    ERR_MSG = "Could not connect to timer"
    def open(self):
        subprocess.Popen([self.CLIENT],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)

    @cachetools.func.ttl_cache(maxsize=None, ttl=10 * 60)
    def status(self)->tuple:
        while True:
            response = subprocess.run([self.CLI, 'status'], stdout=subprocess.PIPE)
            response = json.loads(response.stdout)
            logger.info("HubstaffCLI.bin.x86_64 status response %s", response)
            if "error" in response and response['error'] == self.ERR_MSG:
                self.open()
                time.sleep(60)
            else:
                return (response["active_project"]["tracked_today"], response["tracking"])

    def stop(self):
        response = subprocess.run([self.CLI, 'stop'], stdout=subprocess.PIPE)
        response = json.loads(response.stdout)
        logger.info("HubstaffCLI.bin.x86_64 stop response %s", response)
        if "error" in response and response['error'] == self.ERR_MSG:
            self.open()

    def resume(self):
        response = subprocess.run(['HubstaffCLI.bin.x86_64', 'resume'], stdout=subprocess.PIPE)  
        response = json.loads(response.stdout)
        logger.info("HubstaffCLI.bin.x86_64 resume response %s", response)
        if "error" in response and response['error'] == "Could not connect to timer":
            self.open()

    def kill(self):
        kill_program(self.CLIENT)                     

class Googlesheet:
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SPREADSHEET_ID = '1NTeEHNwggEbVthDbJib5UatM_KPPAiTMCv_lJ4fWK7U'
    RANGE_TEMPLATE = 'Tracking!B{row}:D'
    RANGE_NAME = RANGE_TEMPLATE.format(row=2)
    TOKEN_JSON = 'google_token.json'
    CREDENTIALS_JSON = 'credentials.json'

    @cachetools.func.ttl_cache(maxsize=None, ttl=60 * 60)
    def get_cred(self):
        creds = None
        if os.path.exists(self.TOKEN_JSON):
            creds = Credentials.from_authorized_user_file(self.TOKEN_JSON, self.SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.CREDENTIALS_JSON, self.SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(self.TOKEN_JSON, 'w') as token:
                token.write(creds.to_json())
        self.creds = creds
                
    def __init__(self):
        self.get_cred()
        self.service = build('sheets', 'v4', credentials=self.creds)
        self.sheet = self.service.spreadsheets()

    def read_sheet(self)->List:
        try:
            result = self.sheet.values().get(
                spreadsheetId=self.SPREADSHEET_ID,
                range=self.RANGE_NAME).execute()
            return result.get('values', [])
        except HttpError as err:
            logger.error("Updating sheet failed, response %s", err)
            return None

    def update_sheet(self, tracked_today: str, date = datetime.today()):
        values = self.read_sheet()
        last_row = values[-1]
        row = next(last_row for last_row in reversed(values) if last_row[0] != datetime.today().strftime("%Y/%m/%d"))
        tracked_weekly = float(row[-1] if date.weekday() > 0 else 0)
        tracked_today_hours = (datetime.strptime(tracked_today, "%H:%M:%S") - datetime(1900,1,1)).seconds/3600
        tracked_weekly +=  tracked_today_hours
        value_input_option = 'USER_ENTERED' 

        # How the input data should be inserted.
        insert_data_option = 'INSERT_ROWS'
        date = date.strftime("%Y/%m/%d")
        value_range_body = {
            "values": [
                [
                date,
                tracked_today,
                round(tracked_weekly, 2)
                ]
            ]
        }
        try:
            if date == last_row[0]:
                last_row_ix = len(values) + 1
                last_row_range = self.RANGE_TEMPLATE.format(row=last_row_ix)
                request = self.service.spreadsheets().values().update(
                    spreadsheetId=self.SPREADSHEET_ID, 
                    range=last_row_range, 
                    valueInputOption=value_input_option, 
                    body=value_range_body)
            else:
                request = self.service.spreadsheets().values().append(
                    spreadsheetId=self.SPREADSHEET_ID, 
                    range=self.RANGE_NAME, 
                    valueInputOption=value_input_option, 
                    insertDataOption=insert_data_option, 
                    body=value_range_body)
            response = request.execute()

            logger.info("Response from updating sheet  %s", pformat(response))
        except HttpError as err:
            logger.error("Updating sheet failed, response %s", err)

def exit_gracefully(handle_hubstaff: HubStaff):
    def helper(signum, frame):
        logger.info("Stopping Hubstaff before exiting")
        handle_hubstaff.stop()
        handle_hubstaff.kill()
        os.kill(os.getpid(), signal.SIGKILL)
    return helper

def main():
    SHIFT_LENGTH = 8.1
    
    hubstaff = HubStaff()
    gsheet = Googlesheet()
    
    
    now:datetime = datetime.today() + timedelta(days=1)
    pending = 0
    last_date = datetime.strptime(next(reversed(gsheet.read_sheet()))[0], "%Y/%m/%d")
    pending_duration = (datetime.today() - last_date).days
    for ix in range(1, pending_duration):
        missing_date = last_date + timedelta(days=ix)
        gsheet.update_sheet("00:00:00", missing_date)
    hubstaff.open()
    exit_routine = exit_gracefully(hubstaff)
    signal.signal(signal.SIGINT, exit_routine)
    signal.signal(signal.SIGTERM, exit_routine)
    start_firefox()
    today = datetime.today()
    while True:
        if (_now := datetime.today()) < now and (values := gsheet.read_sheet()):
            row = next(last_row for last_row in reversed(values) if last_row[0] != today.strftime("%Y/%m/%d"))
            tracked_weekly = 0 if today.weekday() == 0 else float(row[-1])
            projected_track = SHIFT_LENGTH*(min(now.weekday(), 5) )
            pending = projected_track- tracked_weekly
            pending *= 60 * 60
            logger.info(f"Weekday = {now.weekday()}")
            logger.info(f"tracked for the week = {tracked_weekly}")
            logger.info(f"Projected to track for today = {projected_track}")
            logger.info(f"Pending weekly tracked time = {pending}")
        now = _now
        
        tracked_today, tracking  = hubstaff.status()
        duration = calc_duration(tracked_today)
        remaining = max(SHIFT_LENGTH*60*60 - duration, 0)
        remaining_delta:timedelta = timedelta(seconds=remaining)
        logger.info(f"Tracking Info [Tracked Today:{tracked_today}, Status:{tracking}, Cumm duration:{duration}, Pending:{pending}, Delta:{remaining_delta.seconds}, Delta Day:{(now + remaining_delta).day}, Now Day:{now.day}")
        if tracking is True and remaining_delta.seconds < 60*5 and (now + remaining_delta).day != now.day:
            logger.info("New Day begining")
            gsheet.update_sheet(tracked_today)
        if tracking is True and duration > pending:
            logger.info(f"Resuming Hubstaff as duration={duration} > pending={pending}")
            gsheet.update_sheet(tracked_today)
            hubstaff.stop()
            hubstaff.status.cache_clear()
        if tracking is False and duration <= pending:
            logger.info(f"Resuming Hubstaff as duration={duration} < pending={pending}")
            hubstaff.resume() 
            hubstaff.status.cache_clear() 
        time.sleep(60)    

if __name__ == '__main__':
    init_logging()
    main()
