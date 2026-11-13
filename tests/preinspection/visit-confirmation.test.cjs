'use strict';
// YES/NO "Is this a pre-inspection?" links: signed, test answers isolated from live decisions.
const assert = require('node:assert/strict');
const path = require('node:path');
const h = require('./harness.cjs');
const env = h.load(path.resolve(__dirname, '..', '..', 'apps-script'), {fixedNow: Date.parse('2026-11-12T15:00:00Z')});
const c = env.ctx, audits = [];
c.tmv3_audit_ = function() { audits.push(Array.from(arguments).slice(1, 5)); };
c.tmv3_piRecordConfirmationInReview_ = () => true;
env.props.setProperty('TMV3_PI_CONFIRM_WEBAPP_URL', 'https://script.google.com/a/macros/classicfireplace.ca/s/TESTID/exec');
const record = {vertical: 'PreInspection', eventId: '6574m7ljop14l7epuih3utso1v@google.com', title: '#59298-Nick Giardina:(416) 500-4002',
  start: '2026-11-18T20:00:00.000Z', end: '2026-11-18T21:00:00.000Z', location: '28 Wrenwood Ct', creator: 'edward@classicfireplace.ca', sourceCreators: []};
const params = href => Object.fromEntries(new URL(href.replace(/&amp;/g, '&')).searchParams);
const links = html => (html.match(/href="([^"]+)"/g) || []).map(x => x.slice(6, -1));

// TEST email: buttons carry t=1; answering does not change the live decision.
const base = c.tmv3_preInspectionMissingDetailEmailPreview_(Object.assign({}, record, {piOffSharedCandidate: true, notificationFollowupNumber: 0}));
const q = c.tmv3_piAddVisitQuestion_(base, record, true);
assert.equal(q.questionAdded, true);
assert.equal(q.recipient, 'edward@classicfireplace.ca');
assert.equal(q.subject, base.subject, 'approved subject unchanged');
assert.match(q.subject, /MISSING DETAILS/);
// Approved template is intact: removing the added section gives back the original email exactly.
const added = q.htmlBody.match(/<div style="margin-top:12px;padding:14px;[\s\S]*?ignore the rest of this email\.<\/div><\/div>/)[0];
assert.equal(q.htmlBody.replace(added, ''), base.htmlBody);
// Layout B: the question comes straight after "Hi <name>," and before the appointment details.
assert.ok(q.htmlBody.indexOf('Hi <strong>') < q.htmlBody.indexOf('First, is this appointment a pre-inspection?'));
assert.ok(q.htmlBody.indexOf('First, is this appointment a pre-inspection?') < q.htmlBody.indexOf('Appointment requiring attention'));
const [yesT, noT] = links(added);
assert.equal(params(yesT).a, 'YES'); assert.equal(params(yesT).t, '1'); assert.equal(params(noT).a, 'NO');
globalThis.__HARNESS_USER__ = 'pramodh@classicfireplace.ca';
let page = c.doGet({parameter: params(yesT)});
assert.match(page.html, /TEST — Thank you\. Recorded: YES/);
assert.equal(c.tmv3_piVisitConfirmation_(record.eventId), null, 'test answer must not set the live decision');
assert.deepEqual(audits.pop(), [record.eventId, '', 'PI_VISIT_CONFIRM_TEST', 'YES']);

// Tampered or swapped link is rejected and records nothing.
const bad = Object.assign({}, params(yesT), {a: 'NO'});
page = c.doGet({parameter: bad});
assert.match(page.html, /not valid/);
assert.equal(audits.length, 0);

// LIVE link records who answered and flips classification.
const live = c.tmv3_piAddVisitQuestion_(base, record, false);
const [yesL] = links(live.htmlBody.match(/First, is this appointment a pre-inspection\?[\s\S]*?ignore the rest/)[0]);
assert.equal(params(yesL).t, undefined);
globalThis.__HARNESS_USER__ = 'edward@classicfireplace.ca';
page = c.doGet({parameter: params(yesL)});
assert.match(page.html, /^<div[\s\S]*Thank you\. Recorded: YES, this is a pre-inspection\./);
const v = c.tmv3_piVisitConfirmation_(record.eventId);
assert.equal(v.answer, 'YES'); assert.equal(v.by, 'edward@classicfireplace.ca');
console.log('PASS: signed YES/NO links, test answers isolated, tamper rejected, live answer recorded with who/when');

// Dynamic rule: the YES/NO box appears only for a booking missing CF Preinspects (off-shared),
// only until it is answered, and never for events already on the shared calendar.
{
  const mails = [];
  c.MailApp = {getRemainingDailyQuota: () => 100, sendEmail: m => mails.push(m)};
  c.tmv3_preInspectionApprovalHold_ = () => null;
  c.tmv3_preInspectionNotificationDecision_ = () => ({send: true, followupNumber: 0, status: 'SEND'});
  const fresh = Object.assign({}, record, {eventId: 'z9offshared1@google.com'});
  const send = r => { mails.length = 0; Object.keys(env.props.store).filter(k => /NOTIFICATION/.test(k)).forEach(k => env.props.deleteProperty(k));
    const res = c.tmv3_piSendMissingDetailsAfterRelease_(r); return {res, html: (mails[0] || {}).htmlBody || ''}; };
  let o = send(Object.assign({}, fresh, {piOffSharedCandidate: true, piVisitQuestion: true}));
  assert.equal(o.res.status, 'SENT', JSON.stringify(o.res));
  assert.match(o.html, /First, is this appointment a pre-inspection\?/, 'off-shared + looks like PI: box shown');
  o = send(Object.assign({}, fresh, {piVisitQuestion: true}));
  assert.doesNotMatch(o.html, /First, is this appointment/, 'on the shared calendar: never a box');
  o = send(Object.assign({}, fresh, {piOffSharedCandidate: true, piVisitQuestion: false}));
  assert.doesNotMatch(o.html, /First, is this appointment/, 'questions switched off: no box');
  env.props.setProperty(c.tmv3_piConfirmKey_(fresh.eventId, false), '{"answer":"YES","by":"edward@classicfireplace.ca"}');
  o = send(Object.assign({}, fresh, {piOffSharedCandidate: true, piVisitQuestion: true}));
  assert.doesNotMatch(o.html, /First, is this appointment/, 'already answered: no box');
  console.log('PASS: YES/NO box only for off-shared look-alike pre-inspections, until answered');
}

// Steve and Colleen Trevor: an event created by either emails both.
{
  assert.deepEqual(Array.from(c.tmv3_piCreatorRecipients_('steve@classicfireplace.ca')), ['steve@classicfireplace.ca', 'colleen@classicfireplace.ca']);
  assert.deepEqual(Array.from(c.tmv3_piCreatorRecipients_('Colleen@classicfireplace.ca')), ['colleen@classicfireplace.ca', 'steve@classicfireplace.ca']);
  assert.deepEqual(Array.from(c.tmv3_piCreatorRecipients_('edward@classicfireplace.ca')), ['edward@classicfireplace.ca']);
  const mails = [];
  c.MailApp = {getRemainingDailyQuota: () => 100, sendEmail: m => mails.push(m)};
  Object.keys(env.props.store).filter(k => /NOTIFICATION/.test(k)).forEach(k => env.props.deleteProperty(k));
  const r = Object.assign({}, record, {eventId: 'z9steve1@google.com', creator: 'steve@classicfireplace.ca', piOffSharedCandidate: true});
  const res = c.tmv3_piSendMissingDetailsAfterRelease_(r);
  assert.equal(res.status, 'SENT', JSON.stringify(res));
  assert.equal(mails[0].to, 'steve@classicfireplace.ca,colleen@classicfireplace.ca');
  assert.doesNotMatch(mails[0].cc, /steve@|colleen@/);
  console.log('PASS: Steve and Colleen Trevor both receive emails for either one\'s bookings');
}
