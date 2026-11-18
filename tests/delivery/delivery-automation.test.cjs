// Standalone Delivery automation test. Not part of the Apps Script project.
// Run: node tests/delivery/delivery-automation.test.cjs
// Loads apps-script/50_Delivery.gs plus the pure helpers and the follow-on map
// it needs from 10_Shared_Core.gs into an isolated context. Apps Script
// services are stubs: no Calendar, Striven, Sheet, Lock or Properties access.
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const dir = path.join(__dirname, '..', '..', 'apps-script');
const core = fs.readFileSync(path.join(dir, '10_Shared_Core.gs'), 'utf8');

function extract(name) {
  const start = core.indexOf('function ' + name + '(');
  if (start < 0) throw new Error('Missing ' + name + ' in 10_Shared_Core.gs');
  let depth = 0, i = core.indexOf('{', start);
  for (; i < core.length; i++) {
    if (core[i] === '{') depth++;
    else if (core[i] === '}' && --depth === 0) break;
  }
  return core.slice(start, i + 1);
}

const followOnLine = (core.match(/^var TMV3_VERTICAL_FOLLOW_ON = .*;$/m) || [])[0];
if (!followOnLine) throw new Error('Missing TMV3_VERTICAL_FOLLOW_ON in 10_Shared_Core.gs');

const calls = [];
const ctx = {
  // Stubs for the shared/runtime functions tmv3_deliverySweepContinue calls.
  tmv3_verticalClearFollowOn_: function(v) { calls.push(['clear', v]); },
  tmv3_runFullVerticalSweep_: function(v, o) { calls.push(['sweep', v, o]); return {status: 'SWEEP_COMPLETE', writes: 0}; },
  tmv3_audit_: function(v, e, t, action, status, detail) { calls.push(['audit', v, action, status, JSON.parse(detail).slot]); }
};
vm.createContext(ctx);
vm.runInContext([
  extract('tmv3_clean_'),
  extract('tmv3_stripPhonesForOrderParsing_'),
  followOnLine,
  fs.readFileSync(path.join(dir, '50_Delivery.gs'), 'utf8'),
  'this.extractOrder = tmv3_extractLegacyDeliveryOrderNumber_;',
  'this.followOn = TMV3_VERTICAL_FOLLOW_ON;',
  'this.continueSweep = tmv3_deliverySweepContinue;'
].join('\n'), ctx);

const cases = [];
function check(name, pass, detail) { cases.push({name: name, pass: pass === true, detail: detail}); }

[
  ['SO_HASH', 'SO#585275 - John Smith - (416) 555-1212', '585275'],
  ['SO_SPACE_HASH', 'Smith delivery SO #585289', '585289'],
  ['SO_DASH', 'Smith - SO-585123 - Jay', '585123'],
  ['SO_COLON', 'Smith SO: 585124', '585124'],
  ['SO_SLASH', 'Smith S/O 585125', '585125'],
  ['SALES_ORDER_NO', 'Smith Sales Order No. 585126', '585126'],
  ['SO_FOUR_DIGIT_MATCHES_NAME_STRIPPER', 'Smith SO 4567', '4567'],
  ['SO_WINS_OVER_CUSTOMER', 'Customer #63164 - Smith - SO#585127', '585127'],
  ['PHONE_NOT_ORDER', 'Smith 905-904-3605', ''],
  ['CUSTOMER_HASH_NOT_ORDER', 'Customer #631645 - Smith', ''],
  ['TASK_NOT_ORDER', 'Smith Task #123456', ''],
  ['WORK_ORDER_NOT_ORDER', 'Smith Work Order #585128', ''],
  ['PO_NOT_ORDER', 'Smith PO 585129', ''],
  ['URL_ID_NOT_ORDER', 'Smith https://example.striven.com/SalesOrders/Edit?id=123456', ''],
  ['ORDER_LABEL', 'Smith Order #585130', '585130'],
  ['SINGLE_HASH_FALLBACK', 'Smith #585131', '585131'],
  ['REPEATED_HASH_SAME_NUMBER', 'Smith #585132 see #585132', '585132'],
  ['AMBIGUOUS_HASH_REVIEW', 'Smith #585133 / Jones #585134', ''],
  ['SINGLE_BARE_SIX_FALLBACK', 'Smith 585135 Jay', '585135'],
  ['AMBIGUOUS_BARE_SIX_REVIEW', 'Smith 585136 585137', ''],
  ['BARE_FIVE_IGNORED', 'Smith 58513', ''],
  ['EMPTY', '', '']
].forEach(function(item) {
  const actual = ctx.extractOrder(item[1]);
  check('DELIVERY_ORDER_' + item[0], actual === item[2], {input: item[1], expected: item[2], actual: actual});
});

check('DELIVERY_FOLLOW_ON_ENABLED', ctx.followOn.Delivery === 'tmv3_deliverySweepContinue', ctx.followOn);
check('INSTALL_FOLLOW_ON_UNCHANGED', ctx.followOn.Install === 'tmv3_installSweepContinue', ctx.followOn);
check('SERVICE_FOLLOW_ON_UNCHANGED', !ctx.followOn.Service, ctx.followOn);

const out = ctx.continueSweep();
check('DELIVERY_CONTINUE_CLEARS_OWN_TRIGGER_FIRST', JSON.stringify(calls[0]) === JSON.stringify(['clear', 'Delivery']), calls);
check('DELIVERY_CONTINUE_RUNS_DELIVERY_FOLLOW_ON', calls[1] && calls[1][0] === 'sweep' && calls[1][1] === 'Delivery' && calls[1][2].followOn === true, calls);
check('DELIVERY_CONTINUE_AUDITS_FOLLOW_ON_SLOT', JSON.stringify(calls[2]) === JSON.stringify(['audit', 'Delivery', 'STANDARD_DAILY_SLOT', 'SWEEP_COMPLETE', 'FOLLOW_ON']), calls);
check('DELIVERY_CONTINUE_RETURNS_SWEEP_RESULT', out.status === 'SWEEP_COMPLETE', out);

const failures = cases.filter(function(c) { return !c.pass; });
failures.forEach(function(f) { console.error('FAIL ' + f.name + ' ' + JSON.stringify(f.detail)); });
console.log((failures.length ? 'FAIL' : 'PASS') + ' ' + (cases.length - failures.length) + '/' + cases.length);
process.exit(failures.length ? 1 : 0);
