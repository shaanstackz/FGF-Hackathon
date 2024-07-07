function loadCalendar() {
    fetch('/load_calendar')
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.json();
        })
        .then(data => {
            const calendarList = document.getElementById('calendar_list');
            calendarList.innerHTML = '';
            data.forEach(appointment => {
                const li = document.createElement('li');
                li.textContent = `${appointment.start} - ${appointment.end}: ${appointment.subject} at ${appointment.location}`;
                calendarList.appendChild(li);
            });
        })
        .catch(error => {
            console.error('Error fetching or parsing data:', error);
        });
}
