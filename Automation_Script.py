#Import the required Libraries
import re
import uuid
import time
import os.path
import datetime as dt
from dotenv import load_dotenv
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from playwright.sync_api import Playwright, sync_playwright, Expect
import os

load_dotenv()

username = os.getenv("username")
password = os.getenv("password")
calendarid = os.getenv('calendar_id') # personal tutoring calendar ID

SCOPES = ["https://www.googleapis.com/auth/calendar"]

# define the function that books the lesson on the
# Lanterna portal
def lanterna_function(name,dd,mm,yyyy,hour,min,lesson_number):
    def run_playwright(playwright):
        
        # launch the browser
        browser = playwright.chromium.launch(headless=False, slow_mo = 300)
        #browser = playwright.chromium.launch(headless=True)

        context = browser.new_context()
        page = context.new_page()

        # load the Lanterna tutoring portal
        page.goto("https://portal.lanterna.com/login")

        # log in
        page.locator("#username").fill(username)
        page.locator("#username").press("Tab")
        page.get_by_label("Password").fill(password) 
        page.get_by_label("Password").press("Enter")

        # Navigate to "My Students" tab
        try:
            frame = page.locator('[data-test-id="interactive-frame"]').content_frame
            frame.locator("#interactive-close-button-container").click()
            #frame.get_by_role("button", name="Close").click()
            page.get_by_role("button", name="Decline").click()
        except Exception:
            pass   # ignore failures and continue
        
        page.get_by_role("link", name="MY STUDENTS MY STUDENTS").click()

        # find the student
        page.get_by_text(f'{name}').click()

        # book the lesson
        page.get_by_role("button", name="BOOK LESSON").click()
        
        page.get_by_label("Date").fill(f"{yyyy}-{mm}-{dd}")
        page.get_by_role("combobox").first.select_option(f"{hour}")  # time for hours
        page.get_by_role("combobox").nth(1).select_option(f"{min}")  # time for minutes
        page.get_by_role("textbox", name="Remarks:").click()
        page.get_by_role("textbox", name="Remarks:").fill(f"Lesson {lesson_number}")
        page.get_by_role("button", name="Book Lesson", exact=True).click()
        
        # ---------------------
        context.close()
        browser.close()

    with sync_playwright() as playwright:
        run_playwright(playwright)

# define the function that schedules a Google Meet call
def meet_function(name, dd,mm,yyyy, hour, minute, email, lesson_number):

    creds = None

    # Load existing credentials from token.json if available
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # If no valid credentials, prompt the user to log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("calendar", "v3", credentials=creds)

        # create the event
        event = {
            "summary": f"{name} Lesson {lesson_number}",

            # need to change the way time is input
            "start": {
            "dateTime": f"{yyyy}-{mm}-{dd}T{hour}:{minute}:00+00:00", # make 00+01:00 for summer time
            "timeZone": "Europe/London"
            },

            "end": {
            "dateTime": f"{yyyy}-{mm}-{dd}T{hour + 1}:{minute}:00+00:00", # make 00+01:00 for summer time
            "timeZone": "Europe/London"
            },

            "attendees": [      
            {"email": f"{email}"}                       
            ],

                    "conferenceData": {
                    "createRequest": {
                        "requestId": str(uuid.uuid4()),  # A unique ID for the request  
                        "conferenceSolutionKey": {
                            "type": "hangoutsMeet"       # Meeting link
                        },
                        "status": {
                            "statusCode": "success"
                }
            }
            }
        }

        event = service.events().insert(calendarId=calendarid, body=event, conferenceDataVersion=1).execute()

    except HttpError as error:
        print(f"An error occurred: {error}")

if __name__ == "__main__":
    #meet_function('ritwik', '28', '08', '2025', 5, 45, 'ritwik.anand2001@gmail.com', 3)
    print(calendarid)