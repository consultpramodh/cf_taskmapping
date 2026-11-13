const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const ctx={};vm.createContext(ctx);vm.runInContext(fs.readFileSync('apps-script/30_PreInspection.gs','utf8').split('function tmv3_piRebuildClassifyOrder_').pop().replace(/^/, 'function tmv3_piRebuildClassifyOrder_'),ctx);
const order=ctx.tmv3_piRebuildClassifyOrder_,loc=ctx.tmv3_piRebuildResolveLocation_;
assert.equal(order({hasOrderEvidence:false}).decision,'PREINSPECTION_ORDER_OPTIONAL');
assert.equal(order({}).decision,'REVIEW_ORDER_EVIDENCE_INCOMPLETE');
assert.equal(order({hasOrderEvidence:true,statusVerified:true,status:'APPROVED'}).decision,'ROUTE_INSTALLATION');
assert.equal(order({hasOrderEvidence:true,statusVerified:true,status:'PRE_APPROVAL'}).decision,'CONTINUE_PREINSPECTION');
assert.equal(order({hasOrderEvidence:true,status:'APPROVED'}).decision,'REVIEW_ORDER_STATUS_UNVERIFIED');
assert.equal(order({hasOrderEvidence:true,statusVerified:true,status:'IN_PROGRESS'}).decision,'REVIEW_ORDER_STATUS_UNSUPPORTED');
const a={calendarAddressPresent:true,customerId:'1',customerVerified:true,locationListComplete:true,locations:[{id:10,customerId:1,addressMatch:true},{id:11,customerId:1,addressMatch:true}]};
assert.equal(loc(a).decision,'REVIEW_DEFAULT_LOCATION_UNVERIFIED');
assert.equal(loc({...a,defaultLocationVerified:true,defaultLocationId:11}).locationId,'11');
assert.equal(loc({...a,defaultLocationVerified:true,defaultLocationId:12}).decision,'REVIEW_DEFAULT_NOT_MATCHING_ADDRESS');
assert.equal(loc({...a,locations:[a.locations[0]]}).basis,'UNIQUE_ADDRESS');
assert.equal(loc({...a,locationListComplete:false}).decision,'REVIEW_LOCATION_EVIDENCE_INCOMPLETE');
assert.equal(loc({...a,locations:[{id:10,customerId:2,addressMatch:true}]}).decision,'REVIEW_LOCATION_OWNERSHIP_CONFLICT');
assert.equal(loc({...a,locations:[{id:10,customerId:1,addressMatch:null}]}).decision,'REVIEW_ADDRESS_MATCH_UNVERIFIED');
assert.equal(loc({...a,locations:[a.locations[0],a.locations[0]]}).decision,'REVIEW_LOCATION_RECORD_INVALID');
assert.equal(loc({...a,locations:[]}).decision,'REVIEW_NO_MATCHING_LOCATION');
console.log('PASS: classification and Location decision assertions');
const task=ctx.tmv3_piRebuildSelectTask_,payload=ctx.tmv3_piRebuildValidatePayload_;
assert.equal(task({}).decision,'REVIEW_TASK_HISTORY_INCOMPLETE');
assert.equal(task({historyComplete:true,tasks:[],writeOutcomeUncertain:false}).decision,'CREATE_TASK');
assert.equal(task({historyComplete:true,tasks:[],writeOutcomeUncertain:true}).decision,'REVIEW_PRIOR_UNCERTAIN_CREATE');
assert.equal(task({historyComplete:true,tasks:[]}).decision,'REVIEW_PRIOR_UNCERTAIN_CREATE');
assert.equal(task({historyComplete:true,tasks:[{id:1,relationship:'SAME_VISIT',status:'OPEN'}]}).decision,'REUSE_TASK');
assert.equal(task({historyComplete:true,tasks:[{id:1,relationship:'SAME_VISIT',status:'COMPLETED'}]}).decision,'FULFILLED_NO_RECREATE');
assert.equal(task({historyComplete:true,tasks:[{id:1,relationship:'UNKNOWN',status:'OPEN'}]}).decision,'REVIEW_TASK_RELATIONSHIP');
assert.equal(task({historyComplete:true,tasks:[{id:1,relationship:'SAME_VISIT',status:'OPEN'},{id:2,relationship:'SAME_VISIT',status:'OPEN'}]}).decision,'REVIEW_MULTIPLE_SAME_VISIT_TASKS');
assert.equal(payload({Name:'Test'},'CREATE').decision,'ACCEPT');
assert.equal(payload({InfoCustomFields:[]},'CREATE').decision,'REJECT_PROTECTED_FIELDS');
assert.equal(payload({nested:{field854:'x'}},'PATCH').decision,'REJECT_PROTECTED_FIELDS');
assert.equal(payload({SalesOrder:{Id:1}},'CREATE').decision,'REJECT_ORDER_ATTACHMENT');
const cyclic={};cyclic.self=cyclic;assert.equal(payload(cyclic,'PATCH').decision,'REJECT_CYCLIC_PAYLOAD');
console.log('PASS: task lifecycle and payload assertions');

assert.equal(loc({...a,calendarAddressPresent:undefined}).decision,'REVIEW_CALENDAR_ADDRESS_EVIDENCE');
assert.equal(loc({...a,calendarAddressPresent:false}).decision,'QUEUE_MISSING_ADDRESS_NOTICE');
assert.equal(loc({...a,calendarAddressPresent:false,notificationStatus:'UNCERTAIN'}).decision,'REVIEW_NOTIFICATION_OUTCOME_NO_RETRY');
assert.equal(loc({...a,calendarAddressPresent:false,notificationStatus:'SENT',noticeMatchesEventAndAddress:true}).decision,'REVIEW_PRIMARY_LOCATION_UNVERIFIED');
assert.equal(loc({...a,calendarAddressPresent:false,notificationStatus:'SENT',noticeMatchesEventAndAddress:true,primaryLocationVerified:true,primaryLocationId:11}).basis,'PRIMARY_AFTER_NOTICE');
assert.equal(loc({...a,calendarAddressPresent:false,notificationStatus:'SENT',noticeMatchesEventAndAddress:true,primaryLocationVerified:true,primaryLocationId:12}).decision,'REVIEW_PRIMARY_LOCATION_NOT_OWNED');
console.log('PASS: missing-address notification and Primary Location sequence');

assert.equal(loc({...a,calendarAddressPresent:false,notificationStatus:'SENT'}).decision,'REVIEW_NOTIFICATION_EVIDENCE');
const flow=ctx.tmv3_piRebuildDryRun_;
const base={order:{hasOrderEvidence:false},location:{...a,locations:[a.locations[0]]},task:{historyComplete:true,tasks:[],writeOutcomeUncertain:false},payload:{Name:'Example'}};
assert.equal(flow(base).status,'DRY_RUN_READY_FOR_FURTHER_VERIFICATION');
assert.equal(flow({...base,order:{hasOrderEvidence:true,statusVerified:true,status:'APPROVED'}}).status,'ROUTE_INSTALLATION');
assert.equal(flow({...base,location:{...a,calendarAddressPresent:false}}).status,'QUEUE_MISSING_ADDRESS_NOTICE');
assert.equal(flow({...base,task:{historyComplete:true,tasks:[{id:1,relationship:'SAME_VISIT',status:'COMPLETED'}]}}).status,'FULFILLED_NO_RECREATE');
assert.equal(flow({...base,payload:{InfoCustomFields:[]}}).status,'REJECT_PROTECTED_FIELDS');
assert.equal(flow(base).sideEffects,0);
console.log('PASS: six connected-stage dry-run assertions');

assert.equal(payload({nested:[{fieldId:854,value:'private'}]},'PATCH').decision,'REJECT_PROTECTED_FIELDS');

// Grace period never runs past the appointment (Sara Elliott case: booked Thu 11:32, visit Mon 09:00).
{
  const path2 = require('node:path'), h2 = require('./harness.cjs');
  const e2 = h2.load(path2.resolve(__dirname, '..', '..', 'apps-script'), {fixedNow: Date.parse('2026-11-12T17:30:00Z')});
  const rec = {start: '2026-11-16T14:00:00.000Z', calendarCreatedAt: '2026-11-12T16:32:21.000Z', status: 'confirmed'};
  const d = e2.ctx.tmv3_preInspectionNotificationDecision_(rec, null);
  require('node:assert/strict').equal(d.send, true, JSON.stringify(d));
  require('node:assert/strict').equal(d.status, 'FIRST_DUE_GRACE_WOULD_MISS_APPOINTMENT');
  const later = e2.ctx.tmv3_preInspectionNotificationDecision_(Object.assign({}, rec, {start: '2026-11-26T14:00:00.000Z'}), null);
  require('node:assert/strict').equal(later.status, 'GRACE_PERIOD', 'normal grace kept when the visit is well after it');
  console.log('PASS: grace period never runs past the appointment');
}

// Requested By = the employee who created the event (Steve Trevor), not the calendar owner (Stephen).
{
  const p4 = require('node:path'), h4 = require('./harness.cjs'), a4 = require('node:assert/strict');
  const e4 = h4.load(p4.resolve(__dirname, '..', '..', 'apps-script'), {fixedNow: Date.parse('2026-11-13T14:30:00Z')});
  const c4 = e4.ctx;
  c4.tmv3_piVerifiedOwnerEmployee_ = () => ({status: 'MATCHED', employeeId: 15});
  c4.tmv3_piCreatorEmployeeId_ = () => 42;
  const rec = {organizer: 'steve@classicfireplace.ca', creator: 'steve@classicfireplace.ca'};
  a4.throws(() => c4.tmv3_piAssertRequestedByPlan_({expectedRequestedByType: 'employee', expectedRequestedById: '15'}, rec),
    /PI_REQUESTED_BY_MUST_BE_EVENT_CREATOR_NO_WRITE/, 'owner refused when the creator is an employee');
  a4.equal(c4.tmv3_piAssertRequestedByPlan_({expectedRequestedByType: 'employee', expectedRequestedById: '42'}, rec), true);
  c4.tmv3_piCreatorEmployeeId_ = () => null;
  a4.equal(c4.tmv3_piAssertRequestedByPlan_({expectedRequestedByType: 'employee', expectedRequestedById: '15'}, rec), true,
    'owner is the fallback when the creator is not an employee');
  console.log('PASS: Requested By is the event creator; owner only as fallback');
}
