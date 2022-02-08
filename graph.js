const authProvider = {
    getAccessToken: async () => {
        return getToken();
    }
};

const graphClient = MicrosoftGraph.Client.initWithMiddleware({ authProvider });

async function getUser() {
    ensureScope('user.read');
    return graphClient
        .api('/me')
        .select('id,displayName')
        .get();
}

async function getEmails(nextLink) {
    ensureScope('mail.read');
    if (nextLink) {
        return graphClient
            .api(nextLink)
            .get();
    }
    else {
        return graphClient
            .api('/me/messages')
            .select('subject,receivedDateTime')
            .orderby('receivedDateTime desc')
            .top(10)
            .get();
    }
}

async function getEvents() {
    ensureScope('calendars.read');
    const dateNow = new Date();
    const dateNextWeek = new Date();
    dateNextWeek.setDate(dateNow.getDate() + 30);
    
    const query = `startDateTime=${dateNow.toISOString()}&endDateTime=${dateNextWeek.toISOString()}`;
    console.log(query)
    return graphClient
        .api('/me/calendarview')
        .query(query)
        .select('subject,start,end,categories')
        .orderby('Start/DateTime')
        .get();
}
