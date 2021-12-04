Date.prototype.addDays = function(days) {
    const date = new Date(this.valueOf());
    date.setDate(date.getDate() + days);
    return date;
};

// expects a list of events in the body coming from Microsoft Graph me/calendar/{calendarId}/events
module.exports = function(context, req) {
    const threshold = 300; // 5 hours
    const markerSubject = 'PA-BLOCKER';

    // explode events spanning multiple days
    let events = [];
    req.body.forEach(e => {
        if (e.isAllDay && e.start.substring(0, 10) < e.end.substring(0, 10)) {
            var start = new Date(e.start);
            var end = new Date(e.end);
            var day = start;
            while (day < end) {
                eCopy = e;
                eCopy.start = day.toISOString();
                eCopy.end = day.toISOString();
                day = day.addDays(1);
                events.push(eCopy);
            }
        } else {
            events.push(e);
        }
    });

    // go through each event and total duration
    let totals = [];

    events.forEach(e => {
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
                    totalBlocks: 0,
                    isBlocked: false,
                    id: ''
                })
                entry = totals.find(e => e.day === day);
            }

            // capture already blocked days
            if (e.isAllDay && e.subject === markerSubject) {
                if (entry) {
                    entry.isBlocked = true;
                    entry.totalBlocks++;
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

    // determine when to create a blocker and when to remove it
    let creates = [];
    let deletes = [];

    totals.forEach(e => {
        if (e.isBlocked) {
            if (e.total < threshold || e.totalBlocks > 1) { // delete block when below threshold or when redundant blocks
                deletes.push({ day: e.day, id: e.id });
            }
        } else {
            if (e.total > threshold) { // create block when above threshold
                creates.push({ day: e.day, event: markerSubject });
            }
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