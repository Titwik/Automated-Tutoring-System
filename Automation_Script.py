#Import the required Libraries
import pyautogui
import time
import pandas as pd 
import re
from playwright.sync_api import Playwright, sync_playwright, Expect

# import tutee details
#tutees = pd.read_excel("Tutee Details.xlsx")     # necessary to extract their name and email

# define the function that books the lesson on the
# Lanterna portal
def lanterna_function(name,dd,mm,yyyy,lesson_time, lesson_number):
    def run_playwright(playwright):
        # launch the browser
        browser = playwright.chromium.launch(headless=False, slow_mo=100)
        context = browser.new_context()
        page = context.new_page()

        # load the Lanterna tutoring portal
        page.goto("https://portal.lanterna.com/login")

        # log in
        page.locator("#username").fill("amritwik@gmail.com")
        page.locator("#username").press("Tab")
        page.get_by_label("Password").fill("dpha")      
        page.get_by_label("Password").press("CapsLock")
        page.get_by_label("Password").fill("dphaSA")
        page.get_by_label("Password").press("CapsLock")
        page.get_by_label("Password").fill("dphaSAfzxc@1854269")
        page.get_by_label("Password").press("Enter")

        # Navigate to "My Students" tab
        page.get_by_role("link", name="MY STUDENTS MY STUDENTS").click()
        page.frame_locator("[data-test-id=\"interactive-frame\"]").locator("#interactive-close-button").click()

        # find the student
        page.get_by_text(f'{name}').click()
        
        # book the lesson
        page.get_by_role("button", name="BOOK LESSON").click()
        
        page.get_by_label("Date").fill(f"{yyyy}-{mm}-{dd}")
        page.get_by_role("combobox").first.select_option("15")  # time for hours
        page.get_by_role("combobox").nth(1).select_option("00")  # time for minutes
        page.get_by_role("button", name="Decline").click()  # some misc decline cookies thing
        page.get_by_label("Remarks:").click()
        page.get_by_label("Remarks:").fill(f"Lesson {lesson_number}")
        #page.get_by_role("button", name="Book Lesson", exact=True).click()

        # ---------------------
        context.close()
        browser.close()

    with sync_playwright() as playwright:
        run_playwright(playwright)

# define the function that schedules a Google Meet call
def meet_function(name,date,lesson_time,email, lesson_number):
    # open MS Edge
    pyautogui.hotkey('winleft', 's')
    pyautogui.typewrite('microsoft edge')
    pyautogui.press('enter')
    time.sleep(2)
    
    # Open google meets
    pyautogui.typewrite('meet.google.com') 
    pyautogui.press('enter')
    time.sleep(3)
    
    # schedule a meeting for later
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.click(x=700, y = 530)
    
    # fill in details for the meeting
    # name
    time.sleep(3)
    pyautogui.typewrite(f'{name} Lesson {lesson_number}')
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('tab')
    
    # date
    pyautogui.typewrite(f'{date}')
    
    pyautogui.press('tab')
    
    # time
    pyautogui.typewrite(f'{lesson_time}')
    
    # select email line
    time.sleep(1)
    pyautogui.click(x = 860, y = 380)
    
    # enter email for invite
    time.sleep(0.5)
    pyautogui.typewrite(f'{email}')
    
    # select the 'type' of meeting
    pyautogui.click(x = 200, y = 600)
        
    # make the meeting 'type' tutoring
    pyautogui.click(x = 200, y = 780)
    
    # send the invite
    pyautogui.click(x = 840, y = 170)
    time.sleep(0.5)
    pyautogui.moveTo(x= 970, y = 530)   # change to click for final version
    pyautogui.press('esc')
    pyautogui.hotkey('ctrl', 'shift', 'w')




    