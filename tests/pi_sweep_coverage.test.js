// Standalone check of the PreInspection sweep coverage rule (operator 2026-11-23): a slot handles every
// current event, including events added or changed behind the sweep and today's events whose time
// has passed; today's events go first. Run: node tests/pi_sweep_coverage.test.js
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');

const props = {}, ephemeral = {};
const ctx = {
  console, JSON, Date, Math, Number, String, Object, Array, Error, isFinite,
  TMV3_TIMEZONE: 'America/Toronto',
  Utilities: {formatDate: (d, tz, f) => new Date(d.getTime() - 5 * 3600000).toISOString().slice(0, 10)},
  PropertiesService: {getScriptProperties: () => ({getProperty: k => (k in props ? props[k] : null),
    setProperty: (k, v) => { props[k] = String(v); }, deleteProperty: k => { delete props[k]; }})},
  CacheService: {getScriptCache: () => ({get: k => ephemeral[k] || null, put: (k, v) => { ephemeral[k] = v; }, remove: k => { delete ephemeral[k]; }})},
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'apps-script', '30_PreInspection.gs'), 'utf8'), ctx);
ctx.tmv3_hash_ = v => 'h' + String(v);

const NOW = Date.parse('2026-11-23T15:20:00Z'); // 10:20 Toronto
const ev = (id, start, fp) => ({eventId: id + '@google.com', start, fingerprint: (fp || id).padEnd(64, '0')});
function cycle(sweepContinue) {
  return {props: ctx.PropertiesService.getScriptProperties(), now: NOW, options: {sweepContinue: !!sweepContinue}, result: {processed: 0}};
}
function persist(c, selected, processed) {
  c.result = {processed, status: 'COMPLETE', results: []}; c.pending = [];
  c.result.notifications = 0;
  ctx.tmv3_piPersistCycle_(c, selected, 0);
  return c.result.sweep;
}

// Calendar at 08:00: two past-today events, one later today, one Nov 25.
const joy = ev('joy', '2026-11-23T15:00:00Z'), jackie = ev('jackie', '2026-11-23T16:00:00Z');
const junaid = ev('junaid', '2026-11-23T18:00:00Z'), matt = ev('matt', '2026-11-25T17:00:00Z');

// 1. Today's events come first; a batch of 3 covers them.
let c = cycle(false);
let sel = ctx.tmv3_piSelectSweepBatch_(c, [matt, junaid, jackie, joy]);
assert.deepStrictEqual(sel.map(r => r.eventId.split('@')[0]), ['joy', 'jackie', 'junaid', 'matt']);
assert.strictEqual(persist(c, sel, 3).remaining, 1);

// 2. Events added behind the sweep: a new 9:30 today (already passed) and a Nov 25 one. Both are picked
//    up, today's first, even though the sweep is past their start times.
const dave = ev('dave', '2026-11-25T19:00:00Z'), late = ev('late930', '2026-11-23T14:30:00Z');
c = cycle(true);
sel = ctx.tmv3_piSelectSweepBatch_(c, [joy, jackie, junaid, matt, dave, late]);
assert.deepStrictEqual(sel.map(r => r.eventId.split('@')[0]), ['late930', 'matt', 'dave']);
assert.strictEqual(persist(c, sel, 3).remaining, 0);

// 3. An event changed after it was handled (time moved) is handled again in the same sweep.
const jackieMoved = ev('jackie', '2026-11-23T16:30:00Z', 'jackie-moved');
props.TMV3_PI_SWEEP_STATE = JSON.stringify(Object.assign(JSON.parse(props.TMV3_PI_SWEEP_STATE), {status: 'IN_PROGRESS'}));
c = cycle(true);
sel = ctx.tmv3_piSelectSweepBatch_(c, [joy, jackieMoved, junaid, matt, dave, late]);
assert.deepStrictEqual(sel.map(r => r.eventId.split('@')[0]), ['jackie']);

// 4. The event's own verified write changed its fingerprint: not handled twice.
ephemeral['TMV3_PI_DONE_h' + junaid.eventId] = JSON.stringify({fingerprint: 'junaid-linked'.padEnd(64, '0'), at: Date.now()});
const junaidLinked = ev('junaid', '2026-11-23T18:00:00Z', 'junaid-linked');
// Ephemeral store may be Script Properties in this release: mirror the done key there too.
props['TMV3_PI_DONE_h' + junaid.eventId] = ephemeral['TMV3_PI_DONE_h' + junaid.eventId];
c = cycle(true);
sel = ctx.tmv3_piSelectSweepBatch_(c, [joy, jackie, junaidLinked, matt, dave, late]);
assert.ok(!sel.some(r => r.eventId.startsWith('junaid')), 'self-written change must not re-run the event');

// 5. A sweep saved by the old cursor-only release keeps its progress.
props.TMV3_PI_SWEEP_STATE = JSON.stringify({status: 'IN_PROGRESS', startedAt: new Date(NOW - 3600000).toISOString(),
  lastKey: ctx.tmv3_preInspectionSweepKey_(jackie), processed: 2});
c = cycle(true);
sel = ctx.tmv3_piSelectSweepBatch_(c, [joy, jackie, junaid, matt]);
assert.deepStrictEqual(sel.map(r => r.eventId.split('@')[0]), ['junaid', 'matt']);

console.log('pi_sweep_coverage: all 5 checks passed');
