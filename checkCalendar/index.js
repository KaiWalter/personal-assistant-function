// expects a list of events in the body coming from Microsoft Graph me/calendar/{calendarId}/events

module.exports = function(context, req) {
    let totals = [];
    let creates = [];
    let deletes = [];

    const threshold = 300; // 5 hours
    const markerSubject = 'PA-BLOCKER';

    req.body.forEach(e => {
        var start = new Date(e.start);
        var end = new Date(e.end);

        if (start.getDay() > 0 && start.getDay() < 6) { // only count weekdays

            // create aggregation entry
            var day = start.toISOString().substring(0, 10);
            var found = totals.find(e => e.day === day);
            if (!found) {
                totals.push({
                    day: day,
                    total: 0,
                    isBlocked: false,
                    id: ''
                })
                found = totals.find(e => e.day === day);
            }

            // capture already blocked days
            if (e.isAllDay && e.subject === markerSubject) {
                if (found) {
                    found.isBlocked = true;
                    found.id = e.id;
                }
            } else if (e.isAllDay && e.showAs !== 'free') { // ignore completely blocked days
            } else if (e.showAs !== 'free' && !e.isAllDay) { // only count busy days
                var duration = (end.getTime() - start.getTime()) / 60000; // minutes

                if (found) {
                    found.total += duration;
                }
            }
        }

    });

    totals.forEach(e => {
        if (e.total > threshold && !e.isBlocked) {
            creates.push({ day: e.day, event: markerSubject });
        } else if (e.isBlocked) {
            deletes.push({ day: e.day, id: e.id });
        }
    });

    context.res = {
        body: {
            creates: creates,
            deletes: deletes,
            totals: totals
        }
    };

    context.done();
};