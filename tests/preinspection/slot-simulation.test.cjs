'use strict';
// piTest_SimulateSlot (95_PI_Operator_Tests.gs): runs the real daily slot with recorders,
// never writes Tasks/Calendar/email/audit, and restores PreInspection Script Properties.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const h = require('./harness.cjs');
const ROOT = path.resolve(__dirname, '..', '..');
const PRIMARY = 'c_3088a3989f3eb809957ed5c40137a7111a0ac97f68c29b40c144028cb14320dc@group.calendar.google.com';
const NOW = Date.parse('2026-11-12T15:00:00Z');

const env = h.load(path.join(ROOT, 'apps-script'), {fixedNow: NOW,
  transform: (f, t) => f === '30_PreInspection.gs' ? t.replace(/^var TMV3_PI_CLEANED_RELEASE_APPROVED = true;/m, 'var TMV3_PI_CLEANED_RELEASE_APPROVED = false;') : t});
vm.runInContext(fs.readFileSync(path.join(ROOT, 'apps-script', '95_PI_Operator_Tests.gs'), 'utf8'), env.ctx);
const c = env.ctx;
const day = n => new Date(NOW + n * 86400000).toISOString();
const rec = (i, extra) => Object.assign({vertical: 'PreInspection', eventId: 'e' + i + '@google.com', title: '1234' + i + ' - Cust ' + i + ' - (416) 555-010' + i,
  rawDescription: '', descriptionClean: '', calendarRole: 'PRIMARY_SHARED', sourceCalendarRoles: ['PRIMARY_SHARED'], calendarId: PRIMARY,
  customerNumber: '1234' + i, calendarCustomerName: 'Cust ' + i, phone: '(416) 555-010' + i, location: i === 2 ? '' : '1 Main St',
  orderNumber: '', status: 'confirmed', isAllDay: false, creator: 'adam@classicfireplace.ca', sourceCreators: [], guests: '',
  organizerCalendarId: 'stephen@classicfireplace.ca', organizer: 'adam@classicfireplace.ca', start: day(i + 1), end: day(i + 1),
  calendarCreatedAt: day(-3), existingTaskId: '', fingerprint: 'f' + i}, extra || {});
const records = [rec(0), rec(1), rec(2)];
const CONVERGED = {plan: 'NO_CHANGE', readStatus: 'FRESH_TASK_GET', taskId: '10', actions: [], customerCheck: 'MATCH', locationCheck: 'MATCH',
  requestedByCheck: 'MATCH', startCheck: 'MATCH', endCheck: 'MATCH', assignmentCheck: 'MATCH', calendarTitleCheck: 'MATCH', taskNameCheck: 'MATCH'};
const plans = {'e0@google.com': {plan: 'CREATE', actions: ['CREATE_TASK'], expectedCustomerId: '100', expectedRequestedByType: 'employee', expectedRequestedById: '15', eventId: 'e0@google.com'},
  'e1@google.com': CONVERGED, 'e2@google.com': CONVERGED};
c.tmv3_operationWritesEnabled_ = () => true;
c.tmv3_verticalCalendars_ = () => [{calendarId: PRIMARY, calendarName: 'CF Preinspects', role: 'PRIMARY_SHARED'}];
c.CalendarApp = {getCalendarById: id => id === PRIMARY ? {getEvents: () => records.map(r => ({__r: r, getId: () => r.eventId, isAllDayEvent: () => false}))} :
  {getEvents: () => []}};
c.tmv3_calendarEvent_ = (v, cfg, s, ev) => JSON.parse(JSON.stringify(ev.__r));
c.tmv3_findCalendarEventById_ = (cal, id) => { const r = records.find(x => x.eventId === id); return r ? {event: {__r: r}} : null; };
c.tmv3_findFreshEventRecord_ = (v, id) => JSON.parse(JSON.stringify(records.find(r => r.eventId === id)));
c.tmv3_strivenApiBudgetStatus_ = () => ({exhausted: false, remaining: 500, count: 1, date: '2026-11-12'});
c.tmv3_refreshOperationalSourcesIfNeeded_ = () => {};
c.tmv3_rows_ = () => [];
c.tmv3_step4IdentityIndex_ = () => ({customerById: {'100': {'Customer ID': '100', 'Customer Number': '12340', Name: 'Cust 0'}}});
c.tmv3_fetchJson_ = () => ({Data: [], TotalCount: 0});
c.tmv3_getTaskById_ = id => ({'Task ID': String(id), 'Task Type ID': 105, 'Customer ID': '100'});
c.tmv3_step7FreshPlansForEvent_ = (v, id) => [JSON.parse(JSON.stringify(plans[id]))];
let realExecutions = 0, realAudits = 0;
c.tmv3_executeFreshStep7Selection_ = () => { realExecutions++; return {status: 'VERIFIED'}; };
c.tmv3_audit_ = () => { realAudits++; };
env.props.setProperty('TMV3_PI_PENDING_EVENTS', '["old@google.com"]');
env.props.setProperty('TMV3_PI_SWEEP_STATE', '{"status":"IN_PROGRESS","lastKey":"x"}');
env.props.setProperty('OTHER_VERTICAL_KEY', 'keep');
const snap = () => JSON.stringify(Object.keys(env.props.store).sort().map(k => [k, env.props.store[k]]));
const before = snap();

const out = c.piTest_SimulateSlot();
assert.equal(out.status, 'SIMULATED_NO_WRITES', out.error);
assert.equal(out.wouldWriteTasks, 1, 'CREATE for e0 recorded, not executed');
assert.ok(out.events.some(e => e.eventId === 'e0@google.com' && e.outcome === 'WOULD_WRITE'));
assert.equal(out.events.length, 3, 'every event evaluated from the start');
assert.ok(out.emails.some(e => e.title.indexOf('Cust 2') >= 0 && e.missing.some(m => /Location/.test(m))), 'missing-location email previewed');
assert.equal(realExecutions, 0); assert.equal(realAudits, 0);
assert.equal(env.log.filter(e => e.kind === 'MAIL').length, 0);
assert.equal(snap(), before, 'Script Properties restored exactly');
assert.equal(c.TMV3_PI_CLEANED_RELEASE_APPROVED, false);
c.tmv3_executeFreshStep7Selection_({contract: {}});
assert.equal(realExecutions, 1, 'real executor restored');
// Test send: only to Pramodh, TEST prefix, no CC, no notification state written.
const stateBefore = snap();
const r2 = c.piTest_SendTestEmails();
const mails = env.log.filter(e => e.kind === 'MAIL');
assert.equal(r2.count, mails.length); assert.ok(mails.length >= 1);
mails.forEach(m => { assert.equal(m.to, 'pramodh@classicfireplace.ca'); assert.ok(!m.cc); assert.match(m.subject, /^TEST MISSING DETAILS: /); });
assert.equal(snap(), stateBefore, 'no notification state written by test send');
console.log('PASS: slot simulation evaluates all events, records writes/emails, writes nothing, restores state');
