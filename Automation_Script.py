#Import the required Libraries
import re
import uuid
import time
import os.path
import datetime as dt
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from playwright.sync_api import Playwright, sync_playwright, Expect

SCOPES = ["https://www.googleapis.com/auth/calendar"]

# define the function that books the lesson on the
# Lanterna portal
def lanterna_function(name,dd,mm,yyyy,hour,min,lesson_number):
    def run_playwright(playwright):
        # launch the browser
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        # load the Lanterna tutoring portal
        page.goto("https://portal.lanterna.com/login")

        # log in
        page.locator("#username").fill("amritwik@gmail.com")
        page.locator("#username").press("Tab")
        page.get_by_label("Password").fill("PASSWORD") 
        page.get_by_label("Password").press("Enter")

        # Navigate to "My Students" tab
        page.frame_locator("[data-test-id=\"interactive-frame\"]").locator("#interactive-close-button").click()
        page.get_by_role("link", name="MY STUDENTS MY STUDENTS").click()

        # find the student
        page.get_by_text(f'{name}').click()

        # book the lesson
        page.get_by_role("button", name="BOOK LESSON").click()
        
        page.get_by_label("Date").fill(f"{yyyy}-{mm}-{dd}")
        page.get_by_role("combobox").first.select_option(f"{hour}")  # time for hours
        page.get_by_role("combobox").nth(1).select_option(f"{min}")  # time for minutes
        page.get_by_role("button", name="Decline").click()  # some misc decline cookies thing
        page.get_by_label("Remarks:").click()
        page.get_by_label("Remarks:").fill(f"Lesson {lesson_number}")
        page.get_by_role("button", name="Book Lesson", exact=True).click()

        # ---------------------
        context.close()
        browser.close()

    with sync_playwright() as playwright:
        run_playwright(playwright)

# define the function that schedules a Google Meet call
def meet_function(name, dd,mm,yyyy, hour, minute, email, lesson_number):

    # set calendar ID to personal tutoring calendar
    calendarid = "f1114c8d9df3717baf451aa5a3c@group.calendar.google.com"

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
            
            #"location": "Online",
            #"description": "Tutoring thing example",
            #"colorId": '6',


            # need to change the way time is input
            "start": {
            "dateTime": f"{yyyy}-{mm}-{dd}T{hour}:{minute}:00+01:00",
            "timeZone": "Europe/London"
            },

            "end": {
            "dateTime": f"{yyyy}-{mm}-{dd}T{hour + 1}:{minute}:00+01:00",
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

    

