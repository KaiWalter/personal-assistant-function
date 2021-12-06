// expects a list of events in the body coming from Microsoft Graph me/calendar/{calendarId}/events
module.exports = function(context, req) {
    const threshold = 300; // 5 hours
    const markerSubject = 'AUTO-BLOCKER';

    // explode events spanning multiple days
    let events = explodeMultipleDays(req);

    // go through each event and total duration
    let totals = calculateTotals(events, markerSubject);

    // determine when to create a blocker and when to remove it
    let creates = [];
    let deletes = [];

    totals.forEach(e => {
        if (e.isBlocked) {
            if (e.total < threshold || e.totalBlocks > 1) { // delete block when below threshold or when redundant blocks
                if (e.id) {
                    deletes.push({ day: e.day, id: e.id });
                }
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

Date.prototype.addDays = function(days) {
    const date = new Date(this.valueOf());
    date.setDate(date.getDate() + days);
    return date;
};

function explodeMultipleDays(req) {
    let events = [];
    req.body.forEach(e => {
        if (e.isAllDay && e.start.substring(0, 10) < e.end.substring(0, 10)) {
            var start = new Date(e.start);
            var end = new Date(e.end);
            var day = start;
            while (day < end) {

                var eClone = Object.assign({}, e);;

                eClone.start = day.toISOString();
                eClone.end = day.toISOString();
                events.push(eClone);

                day = day.addDays(1);
            }
        } else {
            events.push(e);
        }
    });
    return events;
}

function calculateTotals(events, markerSubject) {
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
                    slots: new Array(24 * 4).fill(0) // 4 slots per hour
                });
                entry = totals.find(e => e.day === day);
            }

            // capture already blocked days
            if (entry) {
                if (e.isAllDay && e.subject === markerSubject) {
                    entry.isBlocked = true;
                    entry.totalBlocks++;
                    entry.id = e.id;
                } else if (e.isAllDay && e.showAs !== 'free') { // count complete blocks
                    entry.isBlocked = true;
                    entry.totalBlocks++;
                } else if (!e.isAllDay && e.showAs !== 'free') { // only count busy days
                    calculateTotalMinutes(start, end, entry);
                }
            }
        }

    });

    return totals;
}

function calculateTotalMinutes(start, end, entry) {

    // fill slots
    var startSlot = (start.getHours() * 4 + Math.floor(start.getMinutes() / 15)) - 1;
    var endSlot = (end.getHours() * 4 + Math.floor(end.getMinutes() / 15)) - 1;

    for (var i = startSlot; i < endSlot; i++) {
        entry.slots[i] = 1;
    }

    // count slots with value 1 and convert to minutes
    entry.total = entry.slots.reduce((a, b) => a + b, 0) * 15;
}