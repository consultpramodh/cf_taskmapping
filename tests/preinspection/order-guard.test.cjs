'use strict';
// Pre-create Sales Order guard (tmv3_preInspectionOrderBeforeCreate_), live
// semantics plus the status ID/name pairing guard. Applies to every source
// calendar, including CF Preinspects (no shared-calendar early return).
const assert = require('node:assert/strict');
const path = require('node:path');
const h = require('./harness.cjs');

const env = h.load(path.resolve(__dirname, '..', '..', 'apps-script'));
let order = null;
env.ctx.tmv3_fetchJson_ = url => /\/search$/.test(url)
  ? {Data: [{Id: 123, Number: '456789'}], TotalCount: 1}
  : {Id: 123, OrderNumber: '456789', Customer: {Id: 10}, Status: order};
const guard = env.ctx.tmv3_preInspectionOrderBeforeCreate_;
const record = {vertical: 'PreInspection', calendarRole: 'PRIMARY_SHARED', calendarId: 'shared', orderNumber: '456789', eventId: 'e'};
const plan = {expectedCustomerId: 10};

order = {Id: 22, Name: 'Approved'};
assert.throws(() => guard(record, plan), /APPROVED_ORDER_INSTALLATION_REVIEW_NO_PI_CREATE/);
order = {Id: 25, Name: 'In Progress'};
assert.throws(() => guard(record, plan), /IN_PROGRESS_ORDER_VISIT_TYPE_REVIEW_NO_PI_CREATE/);
order = {Id: 19, Name: 'Quoted'};
assert.equal(guard(record, plan).status, 'SO_CORROBORATION_ONLY_NOT_ATTACHED');
order = {Id: 22, Name: 'Quoted'};
assert.throws(() => guard(record, plan), /PI_SO_STATUS_ID_NAME_MISMATCH_NO_CREATE/);
order = {Id: 30, Name: 'Closed'};
assert.throws(() => guard(record, plan), /PI_SO_STATUS_UNVERIFIED_NO_CREATE/);
assert.equal(guard(Object.assign({}, record, {orderNumber: ''}), plan).status, 'OPTIONAL_ORDER_MISSING_NOT_AN_ERROR');
assert.throws(() => guard(Object.assign({}, record, {orderNumber: '12'}), plan), /PI_SO_NUMBER_INVALID_REVIEW/);
console.log('PASS: seven Sales Order pre-create guard cases');
