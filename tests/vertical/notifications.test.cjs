'use strict';
// Notifications say which workflow they are from; the Install notice uses the standard layout
// and lists only what is wrong (SO missing / not Approved / not found, address missing).
const assert = require('node:assert/strict');
const path = require('node:path');
const h = require('../preinspection/harness.cjs');
const c = h.load(path.resolve(__dirname, '..', '..', 'apps-script'), {fixedNow: Date.parse('2026-11-13T15:00:00Z')}).ctx;

// Branding: subject prefix + header colour and label, idempotent, for all four workflows.
const std = {subject: 'MISSING DETAILS: X', htmlBody: '<td style="background:#202020;border-top:4px solid #b42b2f;padding:18px 24px;"><div>T</div></td>'};
const brands = {PreInspection: ['PRE-INSPECTION', '#b42b2f'], Install: ['INSTALL', '#1f5fa8'], Delivery: ['DELIVERY', '#2e7d32'], Service: ['SERVICE', '#c26a00']};
Object.keys(brands).forEach(v => {
  const [label, color] = brands[v];
  const e = c.tmv3_brandNotification_(v, std);
  assert.equal(e.subject, '[' + label + '] MISSING DETAILS: X');
  assert.ok(e.htmlBody.indexOf('border-top:4px solid ' + color) > 0 && e.htmlBody.indexOf('>' + label + '</div>') > 0, v);
  assert.equal(JSON.stringify(c.tmv3_brandNotification_(v, e)), JSON.stringify(e), 'idempotent ' + v);
});
assert.equal(c.tmv3_brandSubject_('Service', 'Service automation: action required'), '[SERVICE] Service automation: action required');

// Install notice: SO Quoted -> one item, standard layout, INSTALL branding, Steve + Colleen paired.
const record = {eventId: 'inst1@google.com', title: '63267 - Nickelo - 416-617-6614', orderNumber: '586903', customerNumber: '63267',
  location: '85 Northern Dancer Blvd', start: '2026-11-16T14:00:00.000Z', end: '2026-11-16T18:00:00.000Z',
  creator: 'steve@classicfireplace.ca', organizer: 'steve@classicfireplace.ca', sourceCreators: [], calendarId: 'inst-cal'};
let status = {Id: 19, Name: 'Quoted'};
c.tmv3_fetchJson_ = url => /\/search$/.test(url) ? {Data: [{Id: 900, Number: '586903'}], TotalCount: 1} : {Id: 900, Status: status};
let a = c.tmv3_installNoticeAssessment_(record);
assert.equal(JSON.stringify(a.missing.map(m => m.code)), JSON.stringify(['SO_NOT_APPROVED']));
const mail = c.tmv3_installNoticeEmail_(record, a, 0);
assert.match(mail.subject, /^\[INSTALL\] INSTALL DETAILS REQUIRED: .* \| SO not Approved$/);
assert.match(mail.htmlBody, /Install Update Required/);
assert.match(mail.htmlBody, /What it is now[\s\S]*SO #586903 — Quoted[\s\S]*SO #586903 — Approved/);
assert.match(mail.htmlBody, /Open Sales Order #586903/);
assert.ok(mail.htmlBody.indexOf('border-top:4px solid #1f5fa8') > 0);
assert.deepEqual(Array.from(c.tmv3_piCreatorRecipients_(mail.recipient)), ['steve@classicfireplace.ca', 'colleen@classicfireplace.ca']);
// Approved -> nothing to send; ID/name disagreement -> nothing; Pending Approval -> send.
status = {Id: 22, Name: 'Approved'}; assert.equal(c.tmv3_installNoticeAssessment_(record).shouldNotify, false);
status = {Id: 19, Name: 'Approved'}; assert.equal(c.tmv3_installNoticeAssessment_(record).shouldNotify, false);
status = {Id: 20, Name: 'Pending Approval'}; assert.match(c.tmv3_installNoticeAssessment_(record).missing[0].message, /Pending Approval, not Approved/);
// Dynamic: no SO and no address -> two items, no Striven read.
c.tmv3_fetchJson_ = () => { throw new Error('should not read Striven'); };
a = c.tmv3_installNoticeAssessment_(Object.assign({}, record, {orderNumber: '', location: ''}));
assert.equal(JSON.stringify(a.missing.map(m => m.code)), JSON.stringify(['SALES_ORDER', 'LOCATION']));
// Off by default.
assert.equal(c.tmv3_sendInstallNotices_('T').status, 'INSTALL_NOTICES_OFF');
console.log('PASS: workflow branding on all four; Install notice standard and dynamic (SO not Approved)');
