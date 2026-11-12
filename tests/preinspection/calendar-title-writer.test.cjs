'use strict';
// Shared Calendar title/description writer for PreInspection: the organizer copy is
// written first, guest copies that received the propagated change are not re-written,
// and the previous title is recovered from any copy that still carries it.
const assert = require('node:assert/strict');
const path = require('node:path');
const h = require('./harness.cjs');
const PRIMARY = 'c_3088a3989f3eb809957ed5c40137a7111a0ac97f68c29b40c144028cb14320dc@group.calendar.google.com';
const STEPHEN = 'classicfireplace.ca_c20qcqfhvjbv784asn9pvuiaf4@group.calendar.google.com';

function scenario(guestTitle) {
  const env = h.load(path.resolve(__dirname, '..', '..', 'apps-script'));
  const c = env.ctx, writes = [];
  const shared = {title: '45176 - Doug Poad - 416-710-0088', description: 'SO586685'};
  const guestOverride = {title: guestTitle || null};
  const event = calendarId => ({
    getTitle: () => calendarId === PRIMARY && guestOverride.title ? guestOverride.title : shared.title,
    getDescription: () => shared.description,
    setDescription: v => { writes.push([calendarId, 'description']); if (calendarId === STEPHEN) shared.description = v; },
    setTitle: v => { writes.push([calendarId, 'title']); if (calendarId === STEPHEN) { shared.title = v; guestOverride.title = null; } else guestOverride.title = v; }
  });
  c.tmv3_assertTitleNormalizationWriteAuthorized_ = () => true;
  c.tmv3_titleNormalizationCalendarIds_ = () => [PRIMARY, STEPHEN];
  c.tmv3_findEventCopyRobust_ = id => ({event: event(id)});
  c.tmv3_titleCalendarProtectedSnapshot_ = () => ({});
  c.tmv3_assertTitleProtectedSnapshotUnchanged_ = () => true;
  c.tmv3_recordTitleMigrationState_ = () => true;
  c.tmv3_titleMigrationStateForEvent_ = () => ({preservedTitle: ''});
  c.tmv3_reconcilePreInspectionCalendarParity_ = () => ({status: 'PI_PARITY_READ_ONLY_MATCH'});
  c.tmv3_audit_ = () => {};
  const bundle = {eventRecord: {vertical: 'PreInspection', eventId: 'e@google.com', organizerCalendarId: STEPHEN}};
  const contract = {vertical: 'PreInspection', eventId: 'e@google.com', desiredCalendarTitle: '45176 - Doug Poad - (416) 710-0088'};
  let result, error = '';
  try { result = c.tmv3_operationNormalizeCalendarTitle_(bundle, contract, 'PREINSPECTIONAUTO'); } catch (e) { error = String(e.message); }
  return {result, error, writes, shared};
}

// Normal case: organizer written first; guest copy receives the change by propagation.
let r = scenario();
assert.equal(r.error, '');
assert.deepEqual(r.writes[0], [STEPHEN, 'description']);
assert.ok(!r.writes.some(w => w[0] === PRIMARY), 'guest copy must not be re-written after propagation');
assert.match(r.shared.description, /^45176 - Doug Poad - 416-710-0088/);
// Doug's state: guest copy already renamed, organizer still has the old title -> history recovered.
r = scenario('45176 - Doug Poad - (416) 710-0088');
assert.equal(r.error, '');
assert.match(r.shared.description, /^45176 - Doug Poad - 416-710-0088/, 'previous title recovered from organizer copy');
assert.equal(r.shared.title, '45176 - Doug Poad - (416) 710-0088');
console.log('PASS: organizer-first Calendar write, propagation skip, previous-title recovery');
