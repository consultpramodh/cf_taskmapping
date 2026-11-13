'use strict';
// Daily digest: events in all four workflows with no Striven Task link, sent ONLY to Pramodh.
const assert = require('node:assert/strict');
const path = require('node:path');
const h = require('../preinspection/harness.cjs');
const env = h.load(path.resolve(__dirname, '..', '..', 'apps-script'), {fixedNow: Date.parse('2026-11-13T17:00:00Z')});
const c = env.ctx, mails = [];
const T = s => new Date(s);
const ev = (id, title, start, desc, loc) => ({getId: () => id, getTitle: () => title, getStartTime: () => T(start),
  getEndTime: () => new Date(T(start).getTime() + 3600000), getDescription: () => desc, getLocation: () => loc, isAllDayEvent: () => false});
const byCal = {
  c_3088a3989f3eb809957ed5c40137a7111a0ac97f68c29b40c144028cb14320dc: [
    ev('pi1', '63032 - Angela - (416) 791-8355', '2026-11-16T15:00:00Z', 'notes<br><a href="x?TaskID=1">Task</a>', '184 Bothwell'),
    ev('pi2', '63300 - New Customer - (416) 555-0100', '2026-11-17T15:00:00Z', 'no links', '1 Main St')],
  'classicfireplace.ca_vpg05s56lp9f7kaegsiqask34o': [ev('in1', 'Confirmed SO#586303 - Cooper', '2026-11-18T14:00:00Z', '', '9 Oak Ave'),
    ev('in2', 'Team meeting', '2026-11-18T14:00:00Z', '', 'Office')],
  'classicfireplace.ca_2h90kflud3re121hga9bvsi33o': [],
  'classicfireplace.ca_a7v0u8dna3egshtknhqofp5at0': [ev('sv1', 'Al Konstantinou service', '2026-11-12T14:00:00Z', '', '5 Bay St')]
};
c.CalendarApp = {getCalendarById: id => { const k = Object.keys(byCal).find(x => id.indexOf(x) === 0);
  return k ? {getEvents: () => byCal[k]} : {getEvents: () => []}; }};
c.MailApp = {getRemainingDailyQuota: () => 100, sendEmail: m => mails.push(m)};
c.tmv3_audit_ = () => {};
const scan = c.tmv3_missingLinksScan_();
assert.deepEqual(scan.rows.map(r => r.vertical + ':' + r.title).join('|'),
  'PreInspection:63300 - New Customer - (416) 555-0100|Install:Confirmed SO#586303 - Cooper');   // linked, meeting, past excluded
const r = c.tmv3_sendMissingLinksDigest_(false);
assert.equal(r.status, 'SENT'); assert.equal(mails.length, 1);
assert.equal(mails[0].to, 'pramodh@classicfireplace.ca'); assert.ok(!mails[0].cc && !mails[0].bcc);
assert.match(mails[0].subject, /^\[ALL WORKFLOWS\] 2 event\(s\) missing Striven links \| PRE-INSPECTION 1 · INSTALL 1$/);
assert.match(mails[0].htmlBody, />PRE-INSPECTION<\/span>[\s\S]*>INSTALL<\/span>/);
assert.equal(c.tmv3_sendMissingLinksDigest_(false).status, 'ALREADY_SENT_TODAY');
console.log('PASS: missing-links digest covers all four workflows, operator only, once a day');
