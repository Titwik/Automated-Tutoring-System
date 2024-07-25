# import relevant modules
from tkcalendar import DateEntry
import tkinter as tk
from tkinter import ttk
import pandas as pd
import numpy as np
import Automation_Script as au

#------------------------------------------------------------------------------------------

# import the student details
details = pd.read_excel('/home/titwik/Tutoring Automation/Tutee Details.xlsx')

# create a function that links names to emails
def on_name_selected(event):
    selected_name = name_var.get()
    email_var.set(email_dict.get(selected_name, ""))

# create a function to close the window
def close_window():
    window.destroy()

# create a function that loads automation functions 
def submit():

    # get the entry values
    name = name_var.get()
    email = email_var.get()
    selected_date = date_var.get()
    dd = selected_date[0:2]
    mm = selected_date[3:5]     
    yyyy = selected_date[6:10]
    time = time_entry.get()
    lesson_no = lesson_entry.get()

    au.meet_function(name, selected_date, time, email, lesson_no)

# create a clear all function
def clear():
    name_var.set('')
    email_var.set('')
    date_var.set('')
    time_entry.delete(0, tk.END)
    lesson_entry.delete(0, tk.END)

#------------------------------------------------------------------------------------------

# data for names, emails, dates, and times
names = details['Name'].tolist()            # Student names
email = details['Email'].tolist()           # Student emails
email_dict = dict(zip(names, email))        # Linking names and emails
time_h = np.arange(9, 20, 1)                # Time hour
time_m = np.arange(0, 46, 15)               # Time minute

#------------------------------------------------------------------------------------------

# Create the main window
window = tk.Tk()
window.title("Booking Process!  ")
window.geometry("500x500")  

title = tk.Label(text="Booking details!", font=("Arial", 24))
title.pack()

frame = tk.Frame(window)
frame.rowconfigure([0, 1, 2, 3, 4, 5], weight=1)
frame.columnconfigure([0, 1], weight=1)

#------------------------------------------------------------------------------------------
# Name

name_label = tk.Label(frame, text="Name", font=("Arial", 18))
name_label.grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)

name_var = tk.StringVar()   
name_dropdown = ttk.Combobox(frame, textvariable=name_var, font=('Arial', 18))
name_dropdown['values'] = names
name_dropdown.bind("<<ComboboxSelected>>", on_name_selected)
name_dropdown.grid(row=0, column=1, padx=10, pady=5, sticky=tk.W+tk.E)

#------------------------------------------------------------------------------------------
# Email

email_label = tk.Label(frame, text="Email", font=("Arial", 18))
email_label.grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)

email_var = tk.StringVar()
email_entry = tk.Entry(frame, textvariable=email_var, font=('Arial', 18), state='readonly')
email_entry.grid(row=1, column=1, padx=10, pady=5, sticky=tk.W+tk.E)

#------------------------------------------------------------------------------------------
# Lesson Date

date_label = tk.Label(frame, text="Lesson Date", font=("Arial", 18))
date_label.grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)

date_var = tk.StringVar()
date_entry = DateEntry(frame, textvariable=date_var, font=('Arial', 18), date_pattern='dd-mm-yyyy')
date_entry.grid(row = 2, column = 1, padx=10, pady=5, sticky=tk.W+tk.E)

#------------------------------------------------------------------------------------------
# Lesson Time

time_label = tk.Label(frame, text="Lesson Time", font=("Arial", 18))
time_label.grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)

time_entry = tk.Entry(frame, font=('Arial', 18))
time_entry.grid(row=3, column=1, padx=10, pady=5, sticky=tk.W+tk.E)

#------------------------------------------------------------------------------------------
# Lesson Number

lesson_label = tk.Label(frame, text="Lesson Number", font=("Arial", 18))
lesson_label.grid(row=4, column=0, padx=10, pady=5, sticky=tk.W)

lesson_entry = tk.Entry(frame, font=('Arial', 18))
lesson_entry.grid(row=4, column=1, padx=10, pady=5, sticky=tk.W+tk.E)

#------------------------------------------------------------------------------------------
# Button frame

buttonframe = tk.Frame(frame)
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
buttonframe.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky=tk.W+tk.E)

frame.pack(fill=tk.BOTH, expand=True)

# Run the application
window.mainloop()
