// Operator rule 2026-11-23: automated Calendar changes never notify attendees ("Don't send").
// Static check over apps-script/: Calendar writes may only use CalendarApp setters that send no
// guest updates; Calendar API writes must pass sendUpdates=none; no CalendarApp guest add/remove or
// invites. (The one allowed guest change, adding CF Preinspects to proven rep-booked pre-inspections,
// goes through the Calendar API with sendUpdates=none: operator decision 2026-11-23.)
// Run: node tests/calendar_no_notifications.test.js
'use strict';
const fs = require('fs'), path = require('path'), assert = require('assert');
const dir = path.join(__dirname, '..', 'apps-script'), problems = [];
fs.readdirSync(dir).filter(f => f.endsWith('.gs') && !/^9\d_/.test(f)).forEach(f => {
  const text = fs.readFileSync(path.join(dir, f), 'utf8'), calendarFile = /googleapis\.com\/calendar\/v3/.test(text);
  text.split('\n').forEach((line, i) => {
    const at = f + ':' + (i + 1) + '  ' + line.trim().slice(0, 120);
    if (/\.(addGuest|removeGuest|setGuestsCanSeeGuests|setGuestsCanModify|setGuestsCanInviteOthers)\(/.test(line)) problems.push('guest change: ' + at);
    if (/sendInvites\s*:\s*true|sendNotifications\s*[:=]\s*true/.test(line)) problems.push('invites on: ' + at);
    if (/sendUpdates/.test(line) && !/sendUpdates\s*[:=]\s*['"]?none/.test(line)) problems.push('sendUpdates not none: ' + at);
    if (/Calendar\.Events\.(patch|update|insert|move|remove)\(/.test(line) && !/sendUpdates\s*:\s*['"]none['"]/.test(line))
      problems.push('Calendar API write without sendUpdates:"none": ' + at);
    if (/googleapis\.com\/calendar\/v3/.test(line) && /method\s*:\s*['"](post|put|patch|delete)/i.test(line) && !/sendUpdates=none/.test(line))
      problems.push('REST Calendar write without sendUpdates=none: ' + at);
    // Any write fetch in a file that talks to the Calendar API must carry sendUpdates=none on the same line.
    if (calendarFile && /UrlFetchApp\.fetch\(/.test(line) && /method\s*:\s*['"](post|put|patch|delete)/i.test(line) && !/sendUpdates=none/.test(line))
      problems.push('Calendar-file write fetch without sendUpdates=none: ' + at);
  });
});
assert.deepStrictEqual(problems, [], 'Calendar writes that could notify attendees:\n' + problems.join('\n'));
console.log('calendar_no_notifications: no attendee-notifying Calendar writes');
