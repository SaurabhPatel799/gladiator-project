import tkinter as tk
from tkinter import messagebox, ttk
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
    base_dir = Path(sys._MEIPASS)
else:
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
        update_table()  # Update table after adding new user
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
            subscription_end_date = sheet.cell(user_index, 5).value
            current_date = datetime.now().strftime("%m/%d/%y")
            subscription_end_date = datetime.strptime(subscription_end_date, "%m/%d/%y").strftime("%m/%d/%y")
            if current_date == subscription_end_date:
                message = "Your Subscription Ends Today. Please Pay Your Fee to Enjoy Our Services."
                messagebox.showwarning("Subscription End Alert", message)
                speak_message(message)

            for col in range(7, sheet.col_count + 1):
                if not sheet.cell(user_index, col).value:
                    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    sheet.update_cell(user_index, col, current_time)
                    break
                    
            messagebox.showinfo("Success", "Attendance marked successfully")
            user_id_entry.delete(0, tk.END)
            update_table()  # Update table after marking attendance
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

# Define start_index globally to track pagination
start_index = 1

def delete_row_in_google_sheet(sheet_row_index):
    try:
        sheet.delete_rows(sheet_row_index)  # Use delete_rows() method to delete the specific row
        messagebox.showinfo("Success", f"Row {sheet_row_index} deleted successfully")
        update_table()  # Refresh the table after deletion
    except Exception as e:
        messagebox.showerror("Error", f"Error during deletion process: {str(e)}")

def on_delete_click(event):
    item = table.identify('item', event.x, event.y)
    if table.identify_column(event.x) == '#7':  # Check if the 'Delete' button was clicked
        sheet_row_index = int(table.item(item, "tags")[0])
        confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this record?")
        if confirm:
            try:
                sheet.delete_rows(sheet_row_index)  # Delete the correct row
                messagebox.showinfo("Success", "Record deleted successfully")
                update_table()  # Refresh the table after deletion
            except Exception as e:
                messagebox.showerror("Error during deletion", str(e))


def update_table(filtered_data=None):
    # Clear the table first
    for row in table.get_children():
        table.delete(row)
    # Fetch data from Google Sheets
    all_data = sheet.get_all_values()

    # Determine if we're displaying filtered data or all data
    data_to_display = filtered_data if filtered_data else all_data[2:]
    # Insert data into the table (displaying 10 rows at a time)
    for i, row_data in enumerate(data_to_display, start=1):  # Starting from row 2 in Google Sheets
        sheet_row_index = all_data.index(row_data) + 1  # Get the correct row index from the original sheet data
        row_id = table.insert('', 'end', values=row_data[:6] + ["Delete"])  # Add "Delete" text in 7th column
        table.item(row_id, tags=(str(sheet_row_index),))
    table.bind("<Button-1>", on_delete_click)  # Bind left-click to detect delete clicks

def search_table():
    query = search_entry.get().lower()
    if not query:
        update_table()  # Refresh the table without filtering
        return
    # Fetch data from Google Sheets
    all_data = sheet.get_all_values()
    # Filter rows based on the query and include the sheet row index
    filtered_data = [row for row in all_data[2:] if query in row[0].lower()]
    # Update the table with filtered data
    update_table(filtered_data)

def open_main_app():
    global name_entry, mobile_entry, date_entry, subscription_end_date_entry, user_id_entry, search_entry, table, root

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

    canvas.create_text(1250, 150, text="Mark Attendance", font=('Helvetica', 16, 'bold'), fill="white")

    canvas.create_text(1190, 200, text="User ID:", font=('Helvetica', 12, 'bold'), fill="white")
    user_id_entry = tk.Entry(canvas)
    canvas.create_window(1300, 200, window=user_id_entry)

    submit_button = tk.Button(canvas, text="Submit", command=submit_attendance)
    canvas.create_window(1250, 250, window=submit_button)

    # Search entry
    search_entry = tk.Entry(canvas, font=('Helvetica', 12, 'bold'))
    search_entry.bind("<KeyRelease>", lambda event: search_table())
    canvas.create_window(1150, 450, window=search_entry)
    search_button = tk.Button(canvas, text="Search", command=search_table)
    canvas.create_window(1270, 450, window=search_button)

    # Table setup
    columns = ("1", "2", "3", "4", "5", "6", "7")
    table = ttk.Treeview(root, columns=columns, show='headings', height=10)

    headings = ["Name", "Mobile", "User ID", "Date", "Subscription End", "Attendance","Action"]
    for i, heading in enumerate(headings, start=1):
        table.heading(str(i), text=heading)
        table.column(str(i), width=150)  # Adjust the width as needed
    table.place(x=250, y=480)  # Place the table below the Generate User ID and Mark Attendance sections

    update_table()  # Initialize table with data

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
