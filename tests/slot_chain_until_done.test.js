// Standalone check of the slot chain rule (operator 2026-11-23): a slot keeps running a workflow,
// one pass per execution, until its sweep is done, and never defers events to the next slot.
// Run: node tests/slot_chain_until_done.test.js   (no dependencies; Apps Script services are mocked)
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');

function harness(plan) {
  const props = {}, triggers = [], audits = [], calls = [];
  const ctx = {
    console, JSON, Date, Math, Number, String, Object, Array, Error,
    TMV3_TIMEZONE: 'America/Toronto',
    PropertiesService: {getScriptProperties: () => ({
      getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); },
      deleteProperty: k => { delete props[k]; }})},
    ScriptApp: {
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger: t => { const i = triggers.indexOf(t); if (i >= 0) triggers.splice(i, 1); },
      newTrigger: h => ({timeBased: () => ({after: ms => ({create: () => {
        const t = {getHandlerFunction: () => h, after: ms}; triggers.push(t); return t; }})})})},
    Utilities: {formatDate: () => '2026-11-23'},
    LockService: {getScriptLock: () => ({tryLock: () => true, releaseLock: () => {}})},
    Logger: {log: () => {}},
    tmv3_audit_: function(v, e, t, action, result, detail) { audits.push({v, action, result, detail}); },
    tmv3_installReminderCheck: () => { calls.push('Install:reminders'); return {status:'OK'}; },
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'apps-script', '20_Master_Triggers.gs'), 'utf8'), ctx);
  // Each workflow pass returns the next scripted result for that workflow.
  ctx.tmv3_runPreInspectionDailySlot_ = function(slot, options) {
    calls.push('PreInspection:' + options.pass + (options.continuationPass ? ':cont' : ''));
    return plan.PreInspection.shift();
  };
  ctx.tmv3_runFullVerticalSweep_ = function(vertical, options) {
    calls.push(vertical + ':' + options.pass + (options.continuationPass ? ':cont' : ''));
    return plan[vertical].shift();
  };
  function drive(limit) {
    let out = ctx.tmv3_slotChainStart_('08:00'), n = 0;
    while (triggers.some(t => t.getHandlerFunction() === 'tmv3_slotChainNext') && n++ < limit) out = ctx.tmv3_slotChainNext();
    return out;
  }
  return {ctx, props, triggers, audits, calls, drive};
}

const pi = (processed, remaining, status) => ({status: status || (remaining ? 'IN_PROGRESS' : 'COMPLETE'), processed, sweep: {remaining}});
const v = (scanned, pending, status) => ({status: status || (pending ? 'RUNTIME_CONTINUATION_REQUIRED' : 'SWEEP_COMPLETE'), scanned, pending, results: []});

// 1. Today's 08:00 shape: PI leaves 7, Install 38, Service 10. Each workflow repeats until done, in order.
{
  const h = harness({
    PreInspection: [pi(3, 7), pi(3, 4), pi(3, 1), pi(1, 0)],
    Install: [v(1, 38), v(20, 18), v(18, 0)],
    Delivery: [v(0, 0)],
    Service: [v(3, 10), v(10, 0)],
  });
  const out = h.drive(50);
  assert.strictEqual(out.status, 'SLOT_CHAIN_COMPLETE');
  assert.deepStrictEqual(h.calls, [
    'PreInspection:1', 'PreInspection:2:cont', 'PreInspection:3:cont', 'PreInspection:4:cont',
    'Install:1', 'Install:reminders', 'Install:2:cont', 'Install:3:cont',
    'Delivery:1', 'Service:1', 'Service:2:cont']);
  const steps = JSON.parse(h.props.TMV3_SLOT_CHAIN).steps;
  assert.deepStrictEqual(steps.filter(s => s.status === 'DONE').map(s => s.vertical), ['PreInspection', 'Install', 'Delivery', 'Service']);
  assert.ok(!h.audits.some(a => a.action === 'SLOT_CHAIN' && /NO_PROGRESS|PASS_LIMIT/.test(a.result)));
}

// 2. API budget exhausted: stop that workflow (cannot progress), report it, move on.
{
  const h = harness({PreInspection: [pi(0, 5, 'API_BUDGET_WAIT')], Install: [v(0, 0)], Delivery: [v(0, 0)], Service: [v(0, 0)]});
  assert.strictEqual(h.drive(20).status, 'SLOT_CHAIN_COMPLETE');
  assert.ok(h.audits.some(a => a.action === 'SLOT_CHAIN' && a.result === 'API_BUDGET_WAIT'));
}

// 3. A sweep that makes no progress is reported, not looped forever.
{
  const h = harness({PreInspection: [pi(2, 3), pi(0, 3)], Install: [v(0, 0)], Delivery: [v(0, 0)], Service: [v(0, 0)]});
  assert.strictEqual(h.drive(20).status, 'SLOT_CHAIN_COMPLETE');
  assert.strictEqual(h.calls.filter(c => c.startsWith('PreInspection')).length, 2);
  assert.ok(h.audits.some(a => a.action === 'SLOT_CHAIN' && a.result === 'NO_PROGRESS'));
}

// 4. LOCK_BUSY retries the same workflow; the pass cap ends a sweep that never finishes.
{
  const busy = Array.from({length: 30}, () => ({status: 'LOCK_BUSY', processed: 0, sweep: {remaining: 4}}));
  const h = harness({PreInspection: busy, Install: [v(0, 0)], Delivery: [v(0, 0)], Service: [v(0, 0)]});
  assert.strictEqual(h.drive(60).status, 'SLOT_CHAIN_COMPLETE');
  assert.strictEqual(h.calls.filter(c => c.startsWith('PreInspection')).length, vm.runInContext('TMV3_SLOT_CHAIN_MAX_PASSES', h.ctx));
  assert.ok(h.audits.some(a => a.action === 'SLOT_CHAIN' && a.result === 'PASS_LIMIT'));
}

// 5. Install reminders run once per slot, never on continuation passes; audit rows carry the pass.
{
  const h = harness({PreInspection: [pi(0, 0)], Install: [v(5, 5), v(5, 0)], Delivery: [v(0, 0)], Service: [v(0, 0)]});
  h.drive(20);
  assert.strictEqual(h.calls.filter(c => c === 'Install:reminders').length, 1);
  const rows = h.audits.filter(a => a.v === 'Install' && a.action === 'STANDARD_DAILY_SLOT').map(a => JSON.parse(a.detail).pass);
  assert.deepStrictEqual(rows, [1, 2]);
}

console.log('slot_chain_until_done: all 5 checks passed');
