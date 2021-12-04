// expects a list of events in the body coming from Microsoft Graph me/calendar/{calendarId}/events

module.exports = function(context, req) {
    let totals = [];
    let creates = [];
    let deletes = [];

    const threshold = 300; // 5 hours
    const markerSubject = 'PA-BLOCKER';

    // go through each event
    req.body.forEach(e => {
        var start = new Date(e.start);
        var end = new Date(e.end);

        if (start.getDay() > 0 && start.getDay() < 6) { // only count weekdays

            // create aggregation entry
            var day = start.toISOString().substring(0, 10);
            var entry = totals.find(e => e.day === day);
            if (!entry) {
                totals.push({
                    day: day,
                    total: 0,
                    isBlocked: false,
                    id: ''
                })
                entry = totals.find(e => e.day === day);
            }

            // capture already blocked days
            if (e.isAllDay && e.subject === markerSubject) {
                if (entry) {
                    entry.isBlocked = true;
                    entry.id = e.id;
                }
            } else if (e.isAllDay && e.showAs !== 'free') { // ignore completely blocked days
            } else if (e.showAs !== 'free' && !e.isAllDay) { // only count busy days
                var duration = (end.getTime() - start.getTime()) / 60000; // minutes

                if (entry) {
                    entry.total += duration;
                }
            }
        }

    });

    totals.forEach(e => {
        if (e.total > threshold && !e.isBlocked) { // create block when above threshold
            creates.push({ day: e.day, event: markerSubject });
        } else if (e.total < threshold && e.isBlocked) { // delete block when below threshold
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