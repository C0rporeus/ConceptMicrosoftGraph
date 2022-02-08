async function displayUI() {
    await signIn();

    // Display info from user profile
    const user = await getUser();
    var userName = document.getElementById('userName');
    userName.innerText = user.displayName;

    // Hide login button and initial UI
    var signInButton = document.getElementById('signin');
    signInButton.style = "display: none";
    var content = document.getElementById('content');
    content.style = "display: block";

    var btnShowEvents = document.getElementById('btnShowEvents');
    btnShowEvents.style = "display: block";
}


var nextLink;
async function displayEmail() {
    const emails = await getEmails(nextLink);
    if (!emails || emails.value.length < 1) {
        return;
    }
    nextLink = emails['@odata.nextLink'];
    document.getElementById('displayEmail').style = 'display: none';

    var emailsUl = document.getElementById('emails');
    emails.value.forEach(email => {
        var emailLi = document.createElement('li');
        emailLi.innerText = `${email.subject} (${new Date(email.receivedDateTime).toLocaleString()})`;
        emailsUl.appendChild(emailLi);
    });
    if (nextLink) {
        document.getElementById('loadMoreContainer').style = 'display: block';
    }
    window.scrollTo({ top: emailsUl.scrollHeight, behavior: 'smooth' });
}

async function displayEvents() {
    const events = await getEvents();
    if (!events || events.value.length < 1) {
        const content = document.getElementById('content');
        const noItemsMessage = document.createElement('p');
        noItemsMessage.innerHTML = `No events for the coming week!`;
        content.appendChild(noItemsMessage);
    } else {
        const wrapperShowEvents = document.getElementById('eventWrapper');
        wrapperShowEvents.style = 'display: block';
        const eventsElement = document.getElementById('events');
        eventsElement.innerHTML = '';
        events.value.forEach(event => {
            const eventList = document.createElement('li');
            eventList.innerText = `${event.subject} - Desde ${new Date(event.start.dateTime).toLocaleString()} hasta ${new Date(event.end.dateTime).toLocaleString()} - Clasificación: ${event.categories}`;
            eventsElement.appendChild(eventList);
        });
    }
    const btnShowEvents = document.getElementById('btnShowEvents');
    btnShowEvents.style = 'display: none';
}

document.querySelector('mgt-agenda').templateContext = {
    dayFromDateTime: dateTimeString => {
        let date = new Date(dateTimeString);
        date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
        let monthNames = [
            'Enero',
            'Febrero',
            'March',
            'April',
            'May',
            'June',
            'July',
            'August',
            'September',
            'October',
            'November',
            'December'
        ];

        let monthIndex = date.getMonth();
        let day = date.getDate();
        let year = date.getFullYear();

        return monthNames[monthIndex] + ' ' + day + ' ' + year;
    },

    timeRangeFromEvent: (event, timeFormat) => {
        if (event.isAllDay) {
            return 'TODO EL DIA';
        }

        let prettyPrintTimeFromDateTime = date => {
            date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
            let hours = date.getHours();
            let minutes = date.getMinutes();
            let minutesStr = minutes < 10 ? '0' + minutes : minutes;
            let timeString = hours + ':' + minutesStr;
            if (timeFormat === '12') {
                let ampm = hours >= 12 ? 'PM' : 'AM';
                hours = hours % 12;
                hours = hours ? hours : 12;

                timeString = hours + ':' + minutesStr + ' ' + ampm;
            }

            return timeString;
        };

        let start = prettyPrintTimeFromDateTime(new Date(event.start.dateTime));
        let end = prettyPrintTimeFromDateTime(new Date(event.end.dateTime));

        return start + ' - ' + end;
    }
};
