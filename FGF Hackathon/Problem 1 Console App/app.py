import win32com.client as win32
import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
import pytz

# Local timezone
local_tz = pytz.timezone('America/New_York')

# Check for room availability using Outlook Calendar
def book_room():
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar = namespace.GetDefaultFolder(9)
    appointments = calendar.Items
    appointments.IncludeRecurrences = True

    start_datetime = datetime.strptime(date_entry.get() + ' ' + time_entry.get(), "%Y-%m-%d %H:%M")
    end_datetime = start_datetime + timedelta(hours=int(num_hours_entry.get()))
    
    appointments.Sort("[Start]")
    appointments = appointments.Restrict("[Start] >= '" + start_datetime.strftime("%m/%d/%Y %I:%M %p") + "'")

    rooms = {"Room 1": 10, "Room 2": 15, "Room 3": 20}  # Example rooms and capacities
    num_attendees = int(num_attendees_entry.get())
    available_room = None

    for room, capacity in rooms.items():
        if capacity >= num_attendees:
            room_available = True
            for appointment in appointments:
                appointment_start = appointment.Start
                appointment_end = appointment.End
                if not appointment_start.tzinfo:
                    appointment_start = local_tz.localize(appointment_start)
                if not appointment_end.tzinfo:
                    appointment_end = local_tz.localize(appointment_end)
                
                if appointment.Location == room and appointment_start < end_datetime and start_datetime < appointment_end:
                    room_available = False
                    break
            
            if room_available:
                available_room = room
                break

    if not available_room:
        messagebox.showerror("Error", "No available room can accommodate the number of attendees at the requested time.")
        return

    # Create appointment in Outlook Calendar
    attendee_name = attendee_name_entry.get()
    email = email_entry.get()
    new_appointment = outlook.CreateItem(1)
    new_appointment.Start = start_datetime
    new_appointment.Duration = int(num_hours_entry.get()) * 60
    new_appointment.Subject = f"Meeting booked by {attendee_name}"
    new_appointment.Location = available_room
    new_appointment.MeetingStatus = 1  # Set the appointment as a meeting
    new_appointment.Recipients.Add(email)

    new_appointment.Save()
    new_appointment.Send()

    messagebox.showinfo("Success", f"Room '{available_room}' booked successfully for {attendee_name} from {start_datetime} to {end_datetime}")

def load_calendar():
    # Clear the listbox
    calendar_listbox.delete(0, tk.END)

    # Load Outlook calendar
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar = namespace.GetDefaultFolder(9)
    appointments = calendar.Items
    appointments.IncludeRecurrences = True

    # Get all appointments
    appointments.Sort("[Start]")
    now = datetime.now()
    start = now - timedelta(days=1)  # Start date for fetching appointments
    appointments = appointments.Restrict("[Start] >= '" + start.strftime("%m/%d/%Y %I:%M %p") + "'")

    for appointment in appointments:
        start_time = appointment.Start
        end_time = appointment.End
        subject = appointment.Subject
        location = appointment.Location
        
        calendar_listbox.insert(tk.END, f"{start_time} - {end_time}: {subject} at {location}")

# GUI setup
root = tk.Tk()
root.title("Room Booking System")

tk.Label(root, text="Attendee's Name:").grid(row=0, column=0)
attendee_name_entry = tk.Entry(root)
attendee_name_entry.grid(row=0, column=1)

tk.Label(root, text="Email:").grid(row=1, column=0)
email_entry = tk.Entry(root)
email_entry.grid(row=1, column=1)

tk.Label(root, text="No. of Attendees:").grid(row=2, column=0)
num_attendees_entry = tk.Entry(root)
num_attendees_entry.grid(row=2, column=1)

tk.Label(root, text="Date (YYYY-MM-DD):").grid(row=3, column=0)
date_entry = tk.Entry(root)
date_entry.grid(row=3, column=1)

tk.Label(root, text="Time (HH:MM):").grid(row=4, column=0)
time_entry = tk.Entry(root)
time_entry.grid(row=4, column=1)

tk.Label(root, text="No. of Hours:").grid(row=5, column=0)
num_hours_entry = tk.Entry(root)
num_hours_entry.grid(row=5, column=1)

tk.Button(root, text="Book Room", command=book_room).grid(row=6, column=0, columnspan=2)

# Calendar display
tk.Label(root, text="Outlook Calendar:").grid(row=7, column=0, columnspan=2)
calendar_listbox = tk.Listbox(root, width=100, height=10)
calendar_listbox.grid(row=8, column=0, columnspan=2)

tk.Button(root, text="Load Calendar", command=load_calendar).grid(row=9, column=0, columnspan=2)

root.mainloop()
