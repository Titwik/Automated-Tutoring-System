# import relevant modules
import re
import time
import numpy as np
import pandas as pd
import datetime as dt
import tkinter as tk
from tkinter import ttk
import Automation_Script as au
from tkcalendar import DateEntry
from playwright.sync_api import Playwright, sync_playwright, Expect

#------------------------------------------------------------------------------------------
# import the student details
details = pd.read_excel('/home/titwik/Projects/Tutoring Automation/Tutee Details.xlsx')

# create a function that links names to emails
def on_name_selected(event):
    selected_name = name_var.get()
    email_var.set(email_dict.get(selected_name, ""))

# create a function to close the window
def close_window():
    window.destroy()

# create a function that loads automation functions 
def submit():   

    start = time.time()
    # get the entry values
    name = name_var.get()
    name_code = lant_dict[f'{name}']        # find the entry on the Lanterna website
    email = email_var.get()
    selected_date = date_var.get()
    dd = selected_date[0:2]
    mm = selected_date[3:5]     
    yyyy = selected_date[6:10]
    hour_value = int(hour_variable.get())
    minute_value = min_variable.get()
    lesson_no = lesson_entry.get()

    au.meet_function(name, dd,mm,yyyy, hour_value, minute_value, email, lesson_no)
    print('Google Meet Set up!')
    au.lanterna_function(name_code, dd, mm, yyyy,hour_value, minute_value, lesson_no)
    print('Lesson booked on Lanterna!')
    end = time.time()
    elapsed = end - start
    print(elapsed)

    window.destroy()

# create a clear all function
def clear():
    name_var.set('')
    email_var.set('')
    date_var.set('')
    hour_variable.set('')
    min_variable.set('')
    lesson_entry.delete(0, tk.END)

#------------------------------------------------------------------------------------------
# data for names, emails, dates, and times

names = details['Name'].tolist()                  # Student names
email = details['Email'].tolist()                 # Student emails
lant_code = details['Lanterna Code'].tolist()     # Playwright seeks these entries
email_dict = dict(zip(names, email))              # Linking names and emails
lant_dict = dict(zip(names, lant_code))           # Linking names and codes
time_h = list(np.arange(9, 20, 1))                # Time hour
time_m = ['00','15',"30",'45']                    # Time minute

#------------------------------------------------------------------------------------------
# Create the main window

window = tk.Tk()
window.title("Booking Process!")
#window.geometry("500x500")
 
title = tk.Label(text="Booking details!", font=("Arial", 24))
title.pack()

main_frame = tk.Frame(window)
main_frame.pack(fill=tk.BOTH, expand=True)
main_frame.pack_propagate(False)
main_frame.rowconfigure([0, 1, 2, 3, 4, 5], weight=1)
main_frame.columnconfigure([0, 1], weight=1)

#------------------------------------------------------------------------------------------
# Name

name_label = tk.Label(main_frame, text="Name", font=("Arial", 18))
name_label.grid(row=0, column=0, padx=10, pady=5)

name_var = tk.StringVar()   
name_dropdown = ttk.Combobox(main_frame, textvariable=name_var, font=('Arial', 18))
name_dropdown['values'] = names
name_dropdown.bind("<<ComboboxSelected>>", on_name_selected)
name_dropdown.grid(row=0, column=1, padx=10, pady=5)

#------------------------------------------------------------------------------------------
# Email

email_label = tk.Label(main_frame, text="Email", font=("Arial", 18))
email_label.grid(row=1, column=0, padx=10, pady=5)

email_var = tk.StringVar()
email_entry = tk.Entry(main_frame, textvariable=email_var, font=('Arial', 18), state='readonly')
email_entry.grid(row=1, column=1, padx=10, pady=5)

#------------------------------------------------------------------------------------------
# Lesson Date

date_label = tk.Label(main_frame, text="Lesson Date", font=("Arial", 18))
date_label.grid(row=2, column=0, padx=10, pady=5)

date_var = tk.StringVar()
date_entry = DateEntry(main_frame, textvariable=date_var, font=('Arial', 18), date_pattern='dd-mm-yyyy')
date_entry.grid(row = 2, column = 1, padx=10, pady=5)

#------------------------------------------------------------------------------------------
# Lesson Time

time_label = tk.Label(main_frame, text="Lesson Time", font=("Arial", 18))
time_label.grid(row=3, column=0, padx=10, pady=5)

# create a frame with 1 row and two columns
time_frame = tk.Frame(main_frame)
time_frame.rowconfigure(0, weight=1)
time_frame.columnconfigure([0,1], weight=1)

# create a combobox for the hours in column 0
hour_variable = tk.StringVar()   
hour_dropdown = ttk.Combobox(time_frame, textvariable=hour_variable, font=('Arial', 18))
hour_dropdown['values'] = time_h
hour_dropdown.grid(row=0, column=0, padx=10, pady=5)

# create a combobox for the minutes in column 1
min_variable = tk.StringVar()
min_dropdown = ttk.Combobox(time_frame, textvariable=min_variable, font=('Arial', 18))
min_dropdown['values'] = time_m
min_dropdown.grid(row=0, column=1, padx=10, pady=5)

# place the timeframe row in the main_frame row 3, column 1
time_frame.grid(row=3,column=1, padx=10, pady=10)

#------------------------------------------------------------------------------------------
# Lesson Number

lesson_label = tk.Label(main_frame, text="Lesson Number", font=("Arial", 18))
lesson_label.grid(row=4, column=0, padx=10, pady=5)

lesson_entry = tk.Entry(main_frame, font=('Arial', 18))
lesson_entry.grid(row=4, column=1, padx=10, pady=5)

#------------------------------------------------------------------------------------------
# Button frame

buttonframe = tk.Frame(main_frame)
buttonframe.rowconfigure(0, weight=1)
buttonframe.columnconfigure([0, 1, 2], weight=1)

# Clear and Submit buttons
clear_button = tk.Button(buttonframe, text="Clear", font=('Arial', 18), command = clear)
clear_button.grid(row=0, column=0, padx=5, pady=5)

submit_button = tk.Button(buttonframe, text="Submit", font=('Arial', 18), command = submit)
submit_button.grid(row=0, column=1, padx=5, pady=5)

close_button = tk.Button(buttonframe, text="Close", font=('Arial', 18), command = close_window)
close_button.grid(row=0, column=2, padx=5, pady=5)

# Place buttonframe in the 5th row of the main frame
buttonframe.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

main_frame.pack(fill=tk.BOTH, expand=True)

# Run the application
window.mainloop()

