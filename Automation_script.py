#Import the required Libraries
import pandas as pd 
import pyautogui
import time

# import tutee details
tutees = pd.read_excel("Tutee Details.xlsx")

name = 'DaireKenny'
import re
from playwright.sync_api import Playwright, sync_playwright, expect

def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False, slow_mo=100    )
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://portal.lanterna.com/login")
    page.locator("#username").fill("amritwik@gmail.com")
    page.locator("#username").press("Tab")
    page.get_by_label("Password").fill("dpha")      
    page.get_by_label("Password").press("CapsLock")
    page.get_by_label("Password").fill("dphaSA")
    page.get_by_label("Password").press("CapsLock")
    page.get_by_label("Password").fill("dphaSAfzxc@1854269")
    page.get_by_label("Password").press("Enter")
    page.get_by_role("link", name="MY STUDENTS MY STUDENTS").click()
    page.frame_locator("[data-test-id=\"interactive-frame\"]").locator("#interactive-close-button").click()
    page.get_by_text(name).click()
    page.get_by_role("button", name="BOOK LESSON").click()
    page.get_by_label("Date").fill("2024-07-25")
    page.get_by_role("combobox").nth(1).select_option("00")
    page.get_by_role("button", name="Decline").click()
    page.get_by_label("Remarks:").click()
    page.get_by_label("Remarks:").fill("Lesson 1")
    #page.get_by_role("button", name="Book Lesson", exact=True).click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)



def meet_function(name,date,timee,email):
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
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('enter')
    
    # fill in details for the meeting
    # name
    time.sleep(2)
    pyautogui.typewrite(f'{name} Lesson 1')
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('tab')
    
    # date
    pyautogui.typewrite(f'{date}')
    
    pyautogui.press('tab')
    
    # time
    pyautogui.typewrite(f'{timee}')
    
    for i in range(7):
        pyautogui.press('tab')
    
    # enter email for invite
    time.sleep(2)
    pyautogui.typewrite(f'{email}')
    for i in range(16):
        pyautogui.press('tab')
        
    # make the meeting 'type' tutoring
    pyautogui.press('enter')
    time.sleep(1)
    for i in range(4):    
        pyautogui.press('down')
    pyautogui.press('enter')
    
    for i in range(49):    
        pyautogui.press('tab')
    
    # send the invite
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(1)
    for i in range(4):
        pyautogui.press('tab')