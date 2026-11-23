// Proven pre-write rejections are review items, not uncertain writes (2026-11-23).
// Run: node tests/prewrite_rejection.test.js
'use strict';
const fs = require('fs'), path = require('path'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'apps-script', '10_Shared_Core.gs'), 'utf8');
const body = src.slice(src.indexOf('function tmv3_verticalPrewriteRejection_'));
const fn = new Function(body.slice(0, body.indexOf('\n}\n') + 3) + '; return tmv3_verticalPrewriteRejection_;')();
// Kerry Warren 2026-11-18: no Location create in the plan, source selection failed before any POST.
assert.strictEqual(fn('RECREATE requires exactly one completed source Task; found 2.', ''), true);
assert.strictEqual(fn('RECREATE requires exactly one completed source Task; found 2.', 'RECREATE_TASK__TYPE118'), true);
// A plan that created a Location first is not proven write-free.
assert.strictEqual(fn('RECREATE requires exactly one completed source Task; found 2.', 'CREATE_LOCATION_THEN_RECREATE_TASK'), false);
// Klement 2026-11-14: Location ownership failure may follow a Location create; stays held.
assert.strictEqual(fn('Fresh Location ownership verification expected one Location 58743; found 0.', ''), false);
assert.strictEqual(fn('VERTICAL_CONVERGENCE_NOT_VERIFIED:PARTIAL', ''), false);
console.log('prewrite_rejection: all 5 checks passed');
