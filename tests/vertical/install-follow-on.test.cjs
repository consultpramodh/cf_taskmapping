'use strict';
// Install sweep: retries a busy lock, stops starting batches by 2.5 min, and hands the rest to a
// one-off follow-on run that itself continues until the sweep is done.
const assert = require('node:assert/strict');
const path = require('node:path');
const h = require('../preinspection/harness.cjs');
const env = h.load(path.resolve(__dirname, '..', '..', 'apps-script'), {fixedNow: Date.parse('2026-11-13T13:00:00Z')});
const c = env.ctx;
const made = [], audits = [];
let clock = Date.parse('2026-11-13T13:00:00Z');
c.Date.now = () => clock;
c.Utilities.sleep = ms => { clock += ms; };
c.ScriptApp.getProjectTriggers = () => [];
c.ScriptApp.newTrigger = handler => ({timeBased: () => ({after: ms => ({create: () => made.push({handler, ms})})})});
c.tmv3_audit_ = function() { audits.push(Array.from(arguments)); };
c.tmv3_installSchedulePublication_ = () => ({status: 'PUBLISHED'});
c.tmv3_installReminderCheck = () => ({status: 'OK'});

// Slot: lock busy twice, then two batches that leave work pending.
let calls = 0;
c.tmv3_verticalProductionCycle_ = () => {
  calls++;
  if (calls <= 2) return {status: 'LOCK_BUSY', writes: 0};
  clock += 70000;
  return {status: 'RUNTIME_CONTINUATION_REQUIRED', writes: 1, scanned: 3, pending: 20 - 3 * (calls - 2), results: [], sweep: {complete: false}};
};
const slot = c.tmv3_standardVerticalSlot_('Install', '13:00');
assert.equal(slot.lockRetries, 2);
assert.equal(slot.continuation, 'FOLLOW_ON_SCHEDULED_IN_1_MINUTE');
assert.deepEqual(made[0], {handler: 'tmv3_installSweepContinue', ms: 60000});
assert.ok(calls >= 3 && calls <= 5, 'no batch starts after 2.5 minutes: ' + calls);

// Follow-on: finishes the sweep, publishes the view, schedules nothing more.
calls = 0; made.length = 0;
c.tmv3_verticalProductionCycle_ = () => { calls++; clock += 30000;
  return calls === 1 ? {status: 'RUNTIME_CONTINUATION_REQUIRED', writes: 2, scanned: 5, pending: 2, results: [], sweep: {complete: false}}
    : {status: 'SWEEP_COMPLETE', writes: 1, scanned: 2, pending: 0, results: [], sweep: {complete: true}}; };
const follow = c.tmv3_installSweepContinue();
assert.equal(follow.status, 'SWEEP_COMPLETE'); assert.equal(follow.writes, 3);
assert.equal(made.length, 0, 'done: no further follow-on');
assert.equal(follow.sheetRefresh.status, 'PUBLISHED');
assert.ok(audits.some(a => a[3] === "STANDARD_DAILY_SLOT" && /"slot":"FOLLOW_ON"/.test(a[5])));

// Daily cap.
for (let i = 0; i < 12; i++) c.tmv3_verticalScheduleFollowOn_('Install');
assert.equal(c.tmv3_verticalScheduleFollowOn_('Install'), 'DEFERRED_TO_NEXT_STANDARD_SLOT_FOLLOW_ON_LIMIT');

// Service/Delivery unchanged: no lock retry, no follow-on.
calls = 0; made.length = 0;
c.tmv3_verticalProductionCycle_ = () => { calls++; return {status: 'LOCK_BUSY', writes: 0}; };
const svc = c.tmv3_runFullVerticalSweep_('Service');
assert.equal(calls, 1); assert.equal(svc.continuation, 'DEFERRED_TO_NEXT_STANDARD_SLOT'); assert.equal(made.length, 0);
// PreInspection never goes through the vertical sweep or its follow-ons.
assert.equal(c.TMV3_VERTICAL_FOLLOW_ON.PreInspection, undefined);
assert.equal(c.TMV3.OPERATIONS.verticalProductionVerticals.indexOf('PreInspection'), -1);
console.log('PASS: Install retries the lock and finishes its sweep in follow-on runs; other workflows unchanged');
