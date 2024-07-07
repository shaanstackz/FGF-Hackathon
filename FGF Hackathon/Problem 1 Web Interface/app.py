from flask import Flask, render_template, request
import pythoncom
import win32com.client
from datetime import datetime, timedelta
import pytz

# Initialize COM at the start of the script
pythoncom.CoInitialize()

app = Flask(__name__)

# Define the EST timezone
est_tz = pytz.timezone('America/New_York')

# Function to convert naive datetime to EST
def convert_to_est(naive_datetime):
    return naive_datetime.astimezone(est_tz)

# Function to check room availability in Outlook Calendar
def check_room_availability(start_datetime, end_datetime, num_attendees, rooms):
    outlook = win32com.client.Dispatch("Outlook.Application", pythoncom.CoInitialize())
    namespace = outlook.GetNamespace("MAPI")
    calendar = namespace.GetDefaultFolder(9)  # 9 represents the calendar folder in Outlook
    appointments = calendar.Items
    appointments.IncludeRecurrences = True

    appointments.Sort("[Start]")
    # Convert start_datetime to EST before querying Outlook
    est_start_datetime = convert_to_est(start_datetime)
    appointments = appointments.Restrict(f"[Start] >= '{est_start_datetime.strftime('%m/%d/%Y %I:%M %p')}'")

    for room, capacity in rooms.items():
        if capacity >= num_attendees:
            room_available = True
            for appointment in appointments:
                appointment_start = convert_to_est(appointment.Start)
                appointment_end = convert_to_est(appointment.End)

                if appointment.Location == room and appointment_start < end_datetime and start_datetime < appointment_end:
                    room_available = False
                    break
            
            if room_available:
                return room

    return None

# Function to create an appointment in Outlook Calendar
def create_appointment(start_datetime, num_hours, attendee_name, email, location):
    outlook = win32com.client.Dispatch("Outlook.Application", pythoncom.CoInitialize())
    new_appointment = outlook.CreateItem(1)  # 1 represents the Outlook AppointmentItem
    # Convert start_datetime to EST before creating the appointment
    est_start_datetime = convert_to_est(start_datetime)
    new_appointment.Start = est_start_datetime
    new_appointment.Duration = num_hours * 60  # Duration in minutes
    new_appointment.Subject = f"Meeting booked by {attendee_name}"
    new_appointment.Location = location
    new_appointment.MeetingStatus = 1  # Set the appointment as a meeting
    new_appointment.Recipients.Add(email)

    new_appointment.Save()
    new_appointment.Send()

# Route to render the booking form
@app.route('/')
def index():
    return render_template('index.html')

# Route to handle room booking
@app.route('/book_room', methods=['POST'])
def book_room():
    attendee_name = request.form['attendee_name']
    email = request.form['email']
    num_attendees = int(request.form['num_attendees'])
    date = request.form['date']
    time = request.form['time']
    num_hours = int(request.form['num_hours'])

    start_datetime = datetime.strptime(f"{date} {time}", "%Y-%m-%d %H:%M")
    end_datetime = start_datetime + timedelta(hours=num_hours)

    rooms = {
        "1235 Ormont Blueberry": 6,
        "1235 Ormont Cinnamon": 6,
        "1235 Ormont Chocolate Chip": 6,
        "1235 Ormont Coffee Cake": 10,
        "1235 Ormont Maple": 10,
        "1235 Ormont Strawberry": 10,
        "IDC Learning Studio 1": 25,
        "IDC Learning Studio 2": 15,
        "Production Meeting Room": 15,
        "T&D Meeting Room": 15
    }

    available_room = check_room_availability(start_datetime, end_datetime, num_attendees, rooms)

    if not available_room:
        error_message = "No available room can accommodate the number of attendees at the requested time."
        return render_template('error.html', error_message=error_message)

    create_appointment(start_datetime, num_hours, attendee_name, email, available_room)

    return render_template('success.html', 
                           available_room=available_room,
                           attendee_name=attendee_name,
                           start_datetime=start_datetime,
                           end_datetime=end_datetime)

if __name__ == '__main__':
    app.run(debug=True)
