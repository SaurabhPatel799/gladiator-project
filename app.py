# import tkinter as tk
# from tkinter import messagebox
# from tkcalendar import DateEntry
# import gspread
# from oauth2client.service_account import ServiceAccountCredentials
# import random
# from PIL import Image, ImageTk
# import os
# import sys
# from pathlib import Path

# # Use sys._MEIPASS to get the directory of the bundled executable
# if getattr(sys, 'frozen', False):
#     # If the script is frozen (i.e., running as an executable)
#     base_dir = Path(sys._MEIPASS)
# else:
#     # If running as a script
#     base_dir = Path(__file__).resolve().parent

# image_path = base_dir / 'gladiator1.jpeg'
# credential_path = base_dir / 'credential.json'

# # Google Sheets setup
# scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
# creds = ServiceAccountCredentials.from_json_keyfile_name(str(credential_path), scope)
# client = gspread.authorize(creds)

# try:
#     # Open your Google Sheet
#     sheet = client.open('Gym Attendance').sheet1
# except gspread.SpreadsheetNotFound:
#     print("Error: Spreadsheet not found. Ensure the spreadsheet name is correct and it is shared with the service account.")
#     exit()

# # Tkinter GUI
# def generate_user_id(existing_ids):
#     while True:
#         user_id = str(random.randint(1000, 9999))
#         if user_id not in existing_ids:
#             return user_id
        
# def submit_user():
#     name = name_entry.get()
#     mobile = mobile_entry.get()
#     date = date_entry.get()
#     if name and mobile and date:
#         # user_id = str(uuid.uuid4())[:8]
#         existing_ids = sheet.col_values(3)
#         user_id = generate_user_id(existing_ids)
#         sheet.append_row([name, mobile, user_id, date, 0])
#         messagebox.showinfo("Success", f"User ID generated: {user_id}")
#         name_entry.delete(0, tk.END)
#         mobile_entry.delete(0, tk.END)
#     else:
#         messagebox.showerror("Error", "All fields are required")

# def submit_attendance():
#     user_id = user_id_entry.get()
#     if user_id:
#         # Get all user IDs from the sheet
#         user_ids = sheet.col_values(3)  # Assuming the User ID is in the 3rd column
#         # print(f"User IDs in sheet: {user_ids}")
#         if user_id in user_ids:
#             # Find the row number of the user ID
#             user_index = user_ids.index(user_id) + 1
#             # Get the current attendance count from column E (index 5)
#             current_attendance = sheet.cell(user_index, 5).value
#             # Increment the attendance count
#             new_attendance = int(current_attendance) + 1
#             sheet.update_cell(user_index, 5, new_attendance)
#             messagebox.showinfo("Success", "Attendance marked successfully")
#             user_id_entry.delete(0, tk.END)
#         else:
#             messagebox.showerror("Error", "User ID not found")
#     else:
#         messagebox.showerror("Error", "User ID is required")

# # Main Window
# root = tk.Tk()
# root.title("Gym Attendance System")

# # Load the background image
# background_image = Image.open(image_path)
# new_width = 1550  # Set your desired width
# new_height = int(background_image.height * new_width / background_image.width)
# background_image = background_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
# background_photo = ImageTk.PhotoImage(background_image)

# # Create a Canvas widget and set the background image
# canvas = tk.Canvas(root, width=new_width, height=new_height)
# canvas.pack(fill="both", expand=True)
# canvas.create_image(0, 0, image=background_photo, anchor="nw")
# canvas.create_text(background_image.width // 2, 50, text="GLADIATOR FITNESS GYM MUNGER", 
#                 font=('Rockwell', 40), fill="white", anchor="n")

# canvas.create_text(200, 150, text="Generate User ID", font=('Helvetica', 16, 'bold'), fill="white")

# canvas.create_text(140, 200, text="Name:", font=('Helvetica', 12, 'bold'), fill="white")
# name_entry = tk.Entry(canvas)
# canvas.create_window(250, 200, window=name_entry)

# canvas.create_text(140, 250, text="Mobile:", font=('Helvetica', 12, 'bold'), fill="white")
# mobile_entry = tk.Entry(canvas)
# canvas.create_window(250, 250, window=mobile_entry)

# canvas.create_text(135, 300, text="Date:", font=('Helvetica', 12, 'bold'), fill="white")
# date_entry = DateEntry(canvas, width=12, background='darkblue', foreground='white', borderwidth=2)
# canvas.create_window(240, 300, window=date_entry)

# generate_button = tk.Button(canvas, text="Generate", command=submit_user)
# canvas.create_window(240, 350, window=generate_button)

# # Right side for marking attendance
# canvas.create_text(1250, 150, text="Mark Attendance", font=('Helvetica', 16, 'bold'), fill="white")

# canvas.create_text(1190, 200, text="User ID:", font=('Helvetica', 12, 'bold'), fill="white")
# user_id_entry = tk.Entry(canvas)
# canvas.create_window(1300, 200, window=user_id_entry)

# submit_button = tk.Button(canvas, text="Submit", command=submit_attendance)
# canvas.create_window(1250, 250, window=submit_button)

# root.mainloop()

# import tkinter as tk
# from tkinter import messagebox
# from tkcalendar import DateEntry
# import gspread
# from oauth2client.service_account import ServiceAccountCredentials
# import random
# from PIL import Image, ImageTk
# import os
# import sys
# from pathlib import Path

# # Use sys._MEIPASS to get the directory of the bundled executable
# if getattr(sys, 'frozen', False):
#     # If the script is frozen (i.e., running as an executable)
#     base_dir = Path(sys._MEIPASS)
# else:
#     # If running as a script
#     base_dir = Path(__file__).resolve().parent

# image_path = base_dir / 'gladiator1.jpeg'
# login_image_path = base_dir / 'gladiatorlogin.jpeg'
# credential_path = base_dir / 'credential.json'

# # Google Sheets setup
# scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
# creds = ServiceAccountCredentials.from_json_keyfile_name(str(credential_path), scope)
# client = gspread.authorize(creds)

# try:
#     sheet = client.open('Gym Attendance').sheet1
#     login_sheet = client.open('Login Details').sheet1
# except gspread.SpreadsheetNotFound:
#     print("Error: Spreadsheet not found. Ensure the spreadsheet name is correct and it is shared with the service account.")
#     exit()

# def generate_user_id(existing_ids):
#     while True:
#         user_id = str(random.randint(1000, 9999))
#         if user_id not in existing_ids:
#             return user_id
        
# def submit_user():
#     name = name_entry.get()
#     mobile = mobile_entry.get()
#     date = date_entry.get()
#     if name and mobile and date:
#         existing_ids = sheet.col_values(3)
#         user_id = generate_user_id(existing_ids)
#         sheet.append_row([name, mobile, user_id, date, 0])
#         messagebox.showinfo("Success", f"User ID generated: {user_id}")
#         name_entry.delete(0, tk.END)
#         mobile_entry.delete(0, tk.END)
#     else:
#         messagebox.showerror("Error", "All fields are required")

# def submit_attendance():
#     user_id = user_id_entry.get()
#     if user_id:
#         user_ids = sheet.col_values(3)
#         if user_id in user_ids:
#             user_index = user_ids.index(user_id) + 1
#             current_attendance = sheet.cell(user_index, 5).value
#             new_attendance = int(current_attendance) + 1
#             sheet.update_cell(user_index, 5, new_attendance)
#             messagebox.showinfo("Success", "Attendance marked successfully")
#             user_id_entry.delete(0, tk.END)
#         else:
#             messagebox.showerror("Error", "User ID not found")
#     else:
#         messagebox.showerror("Error", "User ID is required")

# def login():
#     username = username_entry.get()
#     password = password_entry.get()
#     if username and password:
#         login_data = login_sheet.get_all_records()
#         for record in login_data:
#             if record['Username'] == username and record['Password'] == password:
#                 messagebox.showinfo("Success", "Login successful")
#                 login_window.destroy()
#                 open_main_app()
#                 return
#         messagebox.showerror("Error", "Invalid username or password")
#     else:
#         messagebox.showerror("Error", "All fields are required")

# def open_main_app():
#     global name_entry, mobile_entry, date_entry, user_id_entry

#     root = tk.Tk()
#     root.title("Gym Attendance System")

#     background_image = Image.open(image_path)
#     new_width = 1550
#     new_height = int(background_image.height * new_width / background_image.width)
#     background_image = background_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
#     background_photo = ImageTk.PhotoImage(background_image)

#     canvas = tk.Canvas(root, width=new_width, height=new_height)
#     canvas.pack(fill="both", expand=True)
#     canvas.create_image(0, 0, image=background_photo, anchor="nw")
#     canvas.create_text(background_image.width // 2, 50, text="GLADIATOR FITNESS GYM MUNGER", 
#                     font=('Rockwell', 40), fill="white", anchor="n")

#     canvas.create_text(200, 150, text="Generate User ID", font=('Helvetica', 16, 'bold'), fill="white")

#     canvas.create_text(140, 200, text="Name:", font=('Helvetica', 12, 'bold'), fill="white")
#     name_entry = tk.Entry(canvas)
#     canvas.create_window(250, 200, window=name_entry)

#     canvas.create_text(140, 250, text="Mobile:", font=('Helvetica', 12, 'bold'), fill="white")
#     mobile_entry = tk.Entry(canvas)
#     canvas.create_window(250, 250, window=mobile_entry)

#     canvas.create_text(135, 300, text="Date:", font=('Helvetica', 12, 'bold'), fill="white")
#     date_entry = DateEntry(canvas, width=12, background='darkblue', foreground='white', borderwidth=2)
#     canvas.create_window(240, 300, window=date_entry)

#     generate_button = tk.Button(canvas, text="Generate", command=submit_user)
#     canvas.create_window(240, 350, window=generate_button)

#     canvas.create_text(1250, 150, text="Mark Attendance", font=('Helvetica', 16, 'bold'), fill="white")

#     canvas.create_text(1190, 200, text="User ID:", font=('Helvetica', 12, 'bold'), fill="white")
#     user_id_entry = tk.Entry(canvas)
#     canvas.create_window(1300, 200, window=user_id_entry)

#     submit_button = tk.Button(canvas, text="Submit", command=submit_attendance)
#     canvas.create_window(1250, 250, window=submit_button)

#     root.mainloop()

# # Main Login Window
# login_window = tk.Tk()
# login_window.title("Login")
# login_window.geometry("800x600")
# login_window.resizable(False, False)

# login_background_image = Image.open(login_image_path)
# login_background_photo = ImageTk.PhotoImage(login_background_image)

# canvas = tk.Canvas(login_window, width=login_background_photo.width(), height=login_background_photo.height())
# canvas.pack(fill="both", expand=True)
# canvas.create_image(0, 0, image=login_background_photo, anchor="nw")

# canvas.create_text(400, 150, text="Login", font=('Helvetica', 24, 'bold'), fill="white")

# canvas.create_text(400, 200, text="Username:", font=('Helvetica', 12, 'bold'), fill="white")
# username_entry = tk.Entry(login_window, font=('Helvetica', 12, 'bold'))
# canvas.create_window(400, 230, window=username_entry)

# canvas.create_text(400, 270, text="Password:", font=('Helvetica', 12, 'bold'), fill="white")
# password_entry = tk.Entry(login_window, show='*', font=('Helvetica', 12, 'bold'))
# canvas.create_window(400, 300, window=password_entry)

# login_button = tk.Button(login_window, text="Login", font=('Helvetica', 12, 'bold'), command=login)
# canvas.create_window(400, 340, window=login_button)

# canvas.create_text(400, 580, text="© Copyright Vajre India Technologies Pvt. Ltd. All Rights Reserved 2024 || Version 1.0", 
#                    font=('Helvetica', 10, 'bold'), fill="white")

# login_window.mainloop()

import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import random
from PIL import Image, ImageTk
import os
import sys
from pathlib import Path
from datetime import datetime
import pyttsx3

# Use sys._MEIPASS to get the directory of the bundled executable
if getattr(sys, 'frozen', False):
    # If the script is frozen (i.e., running as an executable)
    base_dir = Path(sys._MEIPASS)
else:
    # If running as a script
    base_dir = Path(__file__).resolve().parent

image_path = base_dir / 'gladiator1.jpeg'
login_image_path = base_dir / 'gladiatorlogin.jpeg'
credential_path = base_dir / 'credential.json'

# Google Sheets setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(str(credential_path), scope)
client = gspread.authorize(creds)

try:
    sheet = client.open('Gym Attendance').sheet1
    login_sheet = client.open('Login Details').sheet1
except gspread.SpreadsheetNotFound:
    print("Error: Spreadsheet not found. Ensure the spreadsheet name is correct and it is shared with the service account.")
    exit()

def generate_user_id(existing_ids):
    while True:
        user_id = str(random.randint(1000, 9999))
        if user_id not in existing_ids:
            return user_id
        
def submit_user():
    name = name_entry.get()
    mobile = mobile_entry.get()
    date = date_entry.get()
    subscription_end_date = subscription_end_date_entry.get()
    if name and mobile and date and subscription_end_date:
        existing_ids = sheet.col_values(3)
        user_id = generate_user_id(existing_ids)
        sheet.append_row([name, mobile, user_id, date, subscription_end_date, 0])
        messagebox.showinfo("Success", f"User ID generated: {user_id}")
        name_entry.delete(0, tk.END)
        mobile_entry.delete(0, tk.END)
        subscription_end_date_entry.delete(0, tk.END)
    else:
        messagebox.showerror("Error", "All fields are required")

def speak_message(message):
    engine = pyttsx3.init()
    engine.say(message)
    engine.runAndWait()

def submit_attendance():
    user_id = user_id_entry.get()
    if user_id:
        user_ids = sheet.col_values(3)
        if user_id in user_ids:
            user_index = user_ids.index(user_id) + 1
            current_attendance = int(sheet.cell(user_index, 6).value)
            new_attendance = current_attendance + 1
            sheet.update_cell(user_index, 6, new_attendance)

            # Check if the current date matches the subscription end date
            subscription_end_date = sheet.cell(user_index, 5).value
            current_date = datetime.now().strftime("%m/%d/%y")
            subscription_end_date = datetime.strptime(subscription_end_date, "%m/%d/%y").strftime("%m/%d/%y")
            if current_date == subscription_end_date:
                message = "Your Subscription Ends Today. Please Pay Your Fee to Enjoy Our Services."
                messagebox.showwarning("Subscription End Alert", message)
                speak_message(message)

            # Find the first empty column from column 7 onward to store the date and time
            for col in range(7, sheet.col_count + 1):
                if not sheet.cell(user_index, col).value:
                    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    sheet.update_cell(user_index, col, current_time)
                    break
                    
            messagebox.showinfo("Success", "Attendance marked successfully")
            user_id_entry.delete(0, tk.END)
        else:
            messagebox.showerror("Error", "User ID not found")
    else:
        messagebox.showerror("Error", "User ID is required")

def login():
    username = username_entry.get()
    password = password_entry.get()
    if username and password:
        login_data = login_sheet.get_all_records()
        for record in login_data:
            if record['Username'] == username and record['Password'] == password:
                messagebox.showinfo("Success", "Login successful")
                login_window.destroy()
                open_main_app()
                return
        messagebox.showerror("Error", "Invalid username or password")
    else:
        messagebox.showerror("Error", "All fields are required")

def open_main_app():
    global name_entry, mobile_entry, date_entry, subscription_end_date_entry, user_id_entry

    root = tk.Tk()
    root.title("Gym Attendance System")

    background_image = Image.open(image_path)
    new_width = 1550
    new_height = int(background_image.height * new_width / background_image.width)
    background_image = background_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
    background_photo = ImageTk.PhotoImage(background_image)

    canvas = tk.Canvas(root, width=new_width, height=new_height)
    canvas.pack(fill="both", expand=True)
    canvas.create_image(0, 0, image=background_photo, anchor="nw")
    canvas.create_text(new_width // 2, 50, text="GLADIATOR FITNESS GYM MUNGER", 
                    font=('Rockwell', 40), fill="white", anchor="n")

    # Adjusted placement for widgets
    canvas.create_text(200, 150, text="Generate User ID", font=('Helvetica', 16, 'bold'), fill="white")

    canvas.create_text(140, 200, text="Name:", font=('Helvetica', 12, 'bold'), fill="white")
    name_entry = tk.Entry(canvas)
    canvas.create_window(250, 200, window=name_entry)

    canvas.create_text(140, 250, text="Mobile:", font=('Helvetica', 12, 'bold'), fill="white")
    mobile_entry = tk.Entry(canvas)
    canvas.create_window(250, 250, window=mobile_entry)

    canvas.create_text(135, 300, text="Date:", font=('Helvetica', 12, 'bold'), fill="white")
    date_entry = DateEntry(canvas, width=12, background='darkblue', foreground='white', borderwidth=2)
    canvas.create_window(240, 300, window=date_entry)
    
    canvas.create_text(125, 350, text="Subscription\nEnd Date:", font=('Helvetica', 12, 'bold'), fill="white")
    subscription_end_date_entry = DateEntry(canvas, width=12, background='darkblue', foreground='white', borderwidth=2)
    canvas.create_window(240, 350, window=subscription_end_date_entry)

    generate_button = tk.Button(canvas, text="Generate", command=submit_user)
    canvas.create_window(240, 400, window=generate_button)

    # Adjusted placement for Mark Attendance section
    canvas.create_text(1250, 150, text="Mark Attendance", font=('Helvetica', 16, 'bold'), fill="white")

    canvas.create_text(1190, 200, text="User ID:", font=('Helvetica', 12, 'bold'), fill="white")
    user_id_entry = tk.Entry(canvas)
    canvas.create_window(1300, 200, window=user_id_entry)

    submit_button = tk.Button(canvas, text="Submit", command=submit_attendance)
    canvas.create_window(1250, 250, window=submit_button)

    canvas.create_text(750, 750, text="© Copyright Vajre India Technologies Pvt. Ltd. All Rights Reserved 2024 || Version 1.0", 
                   font=('Helvetica', 14, 'bold'), fill="white")
    root.mainloop()


# Main Login Window
login_window = tk.Tk()
login_window.title("Login")
login_window.geometry("800x600")
login_window.resizable(False, False)

login_background_image = Image.open(login_image_path)
login_background_photo = ImageTk.PhotoImage(login_background_image)

canvas = tk.Canvas(login_window, width=login_background_photo.width(), height=login_background_photo.height())
canvas.pack(fill="both", expand=True)
canvas.create_image(0, 0, image=login_background_photo, anchor="nw")

canvas.create_text(400, 150, text="Login", font=('Helvetica', 24, 'bold'), fill="white")

canvas.create_text(400, 200, text="Username:", font=('Helvetica', 12, 'bold'), fill="white")
username_entry = tk.Entry(login_window, font=('Helvetica', 12, 'bold'))
canvas.create_window(400, 230, window=username_entry)

canvas.create_text(400, 270, text="Password:", font=('Helvetica', 12, 'bold'), fill="white")
password_entry = tk.Entry(login_window, show='*', font=('Helvetica', 12, 'bold'))
canvas.create_window(400, 300, window=password_entry)

login_button = tk.Button(login_window, text="Login", font=('Helvetica', 12, 'bold'), command=login)
canvas.create_window(400, 340, window=login_button)

canvas.create_text(400, 580, text="© Copyright Vajre India Technologies Pvt. Ltd. All Rights Reserved 2024 || Version 1.0", 
                   font=('Helvetica', 10, 'bold'), fill="white")

login_window.mainloop()