# Automated Tutoring Booking

Created a script to automate logging into and finding a relevant student's profile on a tutoring platform using websrcaping in Playwright. Additionally, made use of the Google calendar API to generate Google Meet events and invite the specified student to these meetings.

The `Automated_Script.py` file contains two functions, one pertaining to navigating and booking a lesson on the tutoring platform, while the other sets up a Google Meet event in my calendar and sends the invite out. 

The `Application_script.py` file contains the `tkinter` code used to create the GUI for the data entry form used in making the booking.

## The booking data entry form

![image](https://github.com/user-attachments/assets/0c096843-1791-440a-a0e5-4b30687d09b5)

Names are updated in an excel spreadsheet along with their corresponding emails and this is fed into the GUI where the name can be selected using a dropdown list, and the email updates automatically when the name is selected. Lesson date, time and number (Lesson 30 with the student Ritwik) is then input and submitted.

This then navigates to my personal Lanterna tutoring portal where my student's details are shown:

## Booking it on tutoring platform

![image](https://github.com/user-attachments/assets/2d18292d-6afb-40ac-a7f0-2be701e3a38b)

The script then books the lesson and uses the data entered in the form to create a lesson on the tutoring platform. 

Finally, it creates a Google Meet event on a specified calendar (calendarId is specified in the file `Application_script.py`) and sends the invite out to the student's email used in the data entry form.



