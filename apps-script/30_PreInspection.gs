// Google Apps Script — CF Tasks V3 PreInspection vertical, rebuild 2026-11-11-r7.
//
// ONE vertical module of the V3 project. It depends on 00_Configuration and
// 10_Shared_Core and keeps every function name those files, 40_Installation,
// 90_Regression_Tests, installed triggers and the TMV3 menu call.
//
// Governing contract: docs/standards/preinspection.md (Standard 2026-11-10-r1).
// Behavioural baseline: live Apps Script capture of 2026-11-11 (staging r5).
//
// Pipeline (one appointment on CF Preinspects):
//   §2  SOURCE        authoritative intake, provenance, recovery, discovery
//   §3  CLASSIFY      Step 2 eligibility, explicit/likely PreInspection
//   §4  ANCHOR        Step 3 Customer / Sales Order evidence
//   §5  IDENTITY      Step 4 Customer-owned Location, optional order evidence
//   §6  TASK MATCH    Step 5 Task history, lifecycle and strong-match rules
//   §7  PRESENTATION  Calendar title/description/link block and copy parity
//   §8  WRITE GUARDS  create/patch payload rules, Requested By, order status
//   §9  NOTIFY        frozen missing-details email and failure alert
//   §10 CYCLE         batch runner: plan → guard → execute → read back → notify
//   §11 SCHEDULE      08:00 / 13:00 / 17:00 slots, recovery, pending review
//   §12 OPERATOR      read-only diagnostics, review sheets, dry run, self test
//   §13 STAGED        pure decision helpers awaiting verified adapters
//
// SAFETY: every Striven/Calendar/guest/email mutation path checks
// TMV3_PI_CLEANED_RELEASE_APPROVED first. It stays false until full V3
// integration, owner/Requested By rules and live read-back are verified.
var TMV3_PI_CLEANED_RELEASE_APPROVED = false;

const TMV3_PI_REVISION = '2026-11-11-r7';

// =====================================================================
// §1 RELEASE GATE, LOCKED STANDARD AND POLICY
// =====================================================================

/** Standard disabled result for operator/automation entry points. */
function tmv3_piStagingDisabled_() {
  if (TMV3_PI_CLEANED_RELEASE_APPROVED) return null;
  return {status:'PI_STAGING_OPERATOR_PATH_DISABLED',writes:0,emailsSent:0,guestInvites:0};
}

/** Positive-integer ID test used for Striven Customer/Task/Employee IDs. */
function tmv3_piPositiveId_(value) {
  return /^[1-9]\d*$/.test(String(value === null || value === undefined ? '' : value));
}

/** Configured hydration cap shared by task-history reads (1..100, default 60). */
function tmv3_piTaskHydrationCap_() {
  const configured = Number((TMV3.OPERATIONS || {}).preInspectionTaskHydrationMaxTasks || 60);
  return Number.isSafeInteger(configured) && configured > 0 ? Math.min(100, configured) : 60;
}

function tmv3_piIsAuthoritativeRecord_(record) {
  return !!record && record.calendarRole === 'PRIMARY_SHARED' &&
    String(record.calendarId) === String(TMV3.VERTICALS.PreInspection.primaryCalendarId);
}

function tmv3_piRendererFingerprint_() {
  const source = Function.prototype.toString.call(tmv3_preInspectionMissingDetailEmailPreview_);
  let hash = 2166136261;
  for (let i = 0; i < source.length; i++) {
    hash = Math.imul(hash ^ source.charCodeAt(i), 16777619) >>> 0;
  }
  return ('00000000' + hash.toString(16)).slice(-8);
}

function tmv3_piAssertLockedStandard_() {
  if (typeof TMV3_HARD_RULES === 'undefined' || !TMV3_HARD_RULES ||
      typeof TMV3 === 'undefined' || !TMV3.VERTICALS || !TMV3.OPERATIONS)
    throw new Error('PI_STANDARD_DEPENDENCY_MISSING_NO_WRITE');
  const root = TMV3_HARD_RULES, r = root.PreInspection || {};
  const live = TMV3.VERTICALS.PreInspection || {};
  if (root.version !== '2026-11-10-r1' || r.lockedStandardId !== '2026-11-10-r1' ||
      r.lockedStandardReference !== 'docs/workflows/preinspection/STANDARD_2026_11_10.md')
    throw new Error('PI_STANDARD_ID_OR_REFERENCE_DRIFT_NO_WRITE');
  if (Number(r.taskTypeId) !== 105 || Number(live.defaultPoolId) !== 8 ||
      r.orderRequired !== false || r.attachOrderToTask !== false ||
      r.field854Policy !== 'DO_NOT_MANAGE' || r.infoCustomFieldsAtCreate !== 'NONE' ||
      r.calendarNotesPolicy !== 'CALENDAR_ONLY' || live.descriptionMustBeBlankAtCreate !== true)
    throw new Error('PI_TASK_HARD_RULE_DRIFT_NO_WRITE');
  if (r.ownerGuestPolicy !== 'VERIFIED_EVENT_OWNER_ONLY')
    throw new Error('PI_GUEST_RULE_DRIFT_NO_WRITE');
  if (JSON.stringify((r.inspectorOwnedFields || []).map(Number).sort(function(a, b) { return a - b; })) !==
      JSON.stringify([852, 853, 854, 860]))
    throw new Error('PI_INSPECTOR_OWNED_FIELDS_DRIFT_NO_WRITE');
  if (JSON.stringify(TMV3.OPERATIONS.preInspectionAutomationScheduleHours || []) !== JSON.stringify([8, 13, 17]) ||
      TMV3.OPERATIONS.preInspectionAutomationScheduleTimezone !== 'America/Toronto')
    throw new Error('PI_DAILY_SLOTS_DRIFT_NO_WRITE');
  const presentation = r.presentation || {};
  if (presentation.linkHeading !== '-----Striven Links-----' ||
      (presentation.calendarTitleParts || []).join('|') !== 'customerNumber|customerName|phone' ||
      (presentation.taskTitleParts || []).join('|') !== 'customerName|address|phone')
    throw new Error('PI_PRESENTATION_DRIFT_NO_WRITE');
  if (r.notificationTemplatePolicy !== 'EXACT_EXISTING_RENDERER_NO_REDRAFT' ||
      r.notificationRendererFingerprint !== tmv3_piRendererFingerprint_())
    throw new Error('PI_NOTIFICATION_TEMPLATE_DRIFT_NO_SEND');
  return {status:'PASS',standard:root.version,reference:r.lockedStandardReference,
    requestedByTypeContract:tmv3_piRequestedByContract_()};
}

function tmv3_preInspectionLockedStandardStatus() {
  return tmv3_piAssertLockedStandard_();
}

// Original: 50_Verticals; Function: tmv3_preInspectionCreatePolicy
function tmv3_preInspectionCreatePolicy() {
  return {taskTypeId:105,requiredPoolId:8,attachSalesOrder:false,
    description:'',calendarNotes:'CALENDAR_ONLY',
    infoCustomFieldsAtCreate:'NONE',field854:'DO_NOT_MANAGE',
    inspectorOwnedFields:[852,853,854,860],
    previousCalendarTitlePolicy:'PRESERVE_ORIGINAL_UNMODIFIED',
    calendarLinksPreservationPolicy:'PRESERVE_AUTHORED_TEXT_AND_CLICKABLE_LINKS',
    requestedBy:'EMPLOYEE_VERIFIED_CALENDAR_OWNER_ONLY',
    assignedToFromOrganizer:'VERIFIED_OWNER_ONLY',
    guestChanges:'NO_UNATTENDED_INVITATIONS',
    calendarLinks:['CUSTOMER_SALES_ORDERS_PAGE','TASK']};
}

/** Operator intake holds from TMV3_HARD_RULES; a duplicate entry blocks writes. */
function tmv3_preInspectionApprovalHold_(eventId) {
  const normalized = String(eventId || '').trim().toLowerCase().replace(/@google[.]com$/, '');
  const entries = TMV3_HARD_RULES.PreInspection.operatorApprovedIntakeHolds || [];
  const matches = entries.filter(function(x) {
    return String(x.eventId || '').toLowerCase().replace(/@google[.]com$/, '') === normalized;
  });
  if (matches.length > 1) throw new Error('DUPLICATE_PI_OPERATOR_APPROVAL_NO_WRITE');
  return matches.length ? matches[0] : null;
}

function tmv3_piCreatorEmailExcluded_(creatorEmail) {
  const creator = String(creatorEmail || '').trim().toLowerCase();
  const blocked = TMV3_HARD_RULES.PreInspection.notificationExcludedCreatorEmails || [];
  return !!creator && blocked.indexOf(creator) >= 0;
}

/** Canonical rule confirmed by operator: RequestedBy.Type = employee.
 * Pure contract validation; never mutate the shared frozen policy in tests. */
function tmv3_piRequestedByContractFromRule_(rule) {
  const configured = String(rule && rule.requestedByType || '').trim().toLowerCase();
  if (configured && configured !== 'employee')
    throw new Error('PI_REQUESTED_BY_SHARED_STANDARD_CONFLICT_NO_WRITE');
  return 'employee';
}

function tmv3_piRequestedByContract_() {
  const rule = (typeof TMV3_HARD_RULES !== 'undefined' && TMV3_HARD_RULES &&
    TMV3_HARD_RULES.PreInspection) || {};
  return tmv3_piRequestedByContractFromRule_(rule);
}

/** Resolve the owning inspector strictly from a known Calendar identity.
 * Calendar list order never confers identity. Stephen (15) and Jay (1) come
 * from the approved employee table; Karen stays REVIEW until certified. A
 * secondary calendar maps to Stephen only when it is the single configured
 * secondary and the configuration names Stephen as its owner. */
function tmv3_piVerifiedOwnerEmployee_(record) {
  const cfg = (typeof TMV3 !== 'undefined' && TMV3.VERTICALS && TMV3.VERTICALS.PreInspection) || {};
  const owner = String(record && record.organizerCalendarId || '').trim().toLowerCase();
  const stephenEmail = 'stephen@classicfireplace.ca';
  const jayEmail = 'jay@classicfireplace.ca';
  const karenEmail = 'karen@classicfireplace.ca';
  const aliases = (cfg.secondaryCalendarIds || []).map(function(v) {
    return String(v || '').trim().toLowerCase();
  }).filter(Boolean);
  const primaryId = String(cfg.primaryCalendarId || '').trim().toLowerCase();
  const configuredStephen = String(cfg.secondaryOwnerEmail || '').trim().toLowerCase();
  if (!owner) return {status:'REVIEW',employeeId:null,blocker:'PI_OWNER_MISSING_NO_WRITE'};
  if (owner === stephenEmail)
    return {status:'MATCHED',employeeId:15,employeeName:'Stephen Foley',evidence:'EXACT_VERIFIED_OWNER_EMAIL'};
  if (owner === jayEmail)
    return {status:'MATCHED',employeeId:1,employeeName:'Jay Scott',evidence:'EXACT_VERIFIED_OWNER_EMAIL'};
  if (owner === karenEmail)
    return {status:'REVIEW',employeeId:null,blocker:'PI_KAREN_EMPLOYEE_ID_UNVERIFIED_NO_WRITE'};
  if (aliases.length === 1 && configuredStephen === stephenEmail &&
      aliases[0] === owner && owner !== primaryId &&
      owner !== jayEmail && owner !== karenEmail) {
    return {status:'MATCHED',employeeId:15,employeeName:'Stephen Foley',
      evidence:'SINGLE_SECONDARY_CALENDAR_EXPLICIT_STEPHEN_OWNER'};
  }
  return {status:'REVIEW',employeeId:null,blocker:'PI_OWNER_EMPLOYEE_MAPPING_UNVERIFIED_NO_WRITE'};
}

/** Assigned employee is the verified owner; a saved override may only reaffirm it. */
function tmv3_preInspectionAssignmentOverride_(record) {
  if (!record || !record.eventId) throw new Error('PI_EVENT_ID_REQUIRED_FOR_ASSIGNMENT');
  const owner = tmv3_piVerifiedOwnerEmployee_(record);
  if (owner.status !== 'MATCHED') throw new Error(owner.blocker);
  const key = 'TMV3_PI_ASSIGNMENT_OVERRIDE_' + tmv3_hash_(record.eventId);
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return [owner.employeeId];
  let saved;
  try { saved = JSON.parse(raw); } catch (error) { throw new Error('PI_ASSIGNMENT_OVERRIDE_INVALID'); }
  if (String(saved.eventId) !== String(record.eventId) ||
      String(saved.customerNumber) !== String(record.customerNumber) ||
      String(saved.organizerCalendarId) !== String(record.organizerCalendarId) ||
      Number(saved.employeeId) !== owner.employeeId)
    throw new Error('PI_OVERRIDE_CONFLICTS_WITH_VERIFIED_OWNER_NO_WRITE');
  return [owner.employeeId];
}

/** Public operator function; refuses any employee other than the verified owner. */
function tmv3_configurePreInspectionAssignmentOverride(eventId, customerNumber, organizerCalendarId, employeeId) {
  if (!/^[a-z0-9_]+@google\.com$/i.test(String(eventId || ''))) throw new Error('PI_EVENT_ID_INVALID');
  const record = tmv3_findFreshEventRecord_('PreInspection', eventId);
  if (!record || String(record.customerNumber) !== String(customerNumber) ||
      String(record.organizerCalendarId) !== String(organizerCalendarId))
    throw new Error('PI_ASSIGNMENT_SOURCE_EVIDENCE_CHANGED');
  const approved = tmv3_piVerifiedOwnerEmployee_(record);
  if (approved.status !== 'MATCHED' || Number(employeeId) !== approved.employeeId)
    throw new Error('PI_OPERATOR_OVERRIDE_MUST_MATCH_VERIFIED_OWNER');
  const value = {eventId:String(eventId),customerNumber:String(customerNumber),
    organizerCalendarId:String(organizerCalendarId),employeeId:approved.employeeId};
  const key = 'TMV3_PI_ASSIGNMENT_OVERRIDE_' + tmv3_hash_(eventId);
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(value));
  return {status:'VERIFIED_OWNER_OVERRIDE_SAVED',employeeId:approved.employeeId,businessWrites:0};
}

// =====================================================================
// §2 SOURCE — CF Preinspects is the only business source. Representative
// calendars (Stephen/Jay/Karen) are discovery-only and never auto-promoted.
// =====================================================================

/** Shared Step 1 hook: only the PRIMARY_SHARED calendar feeds intake. */
function tmv3_preInspectionStep1SourceAllowed_(cfg, calCfg, creator) {
  if (calCfg.role === 'PRIMARY_SHARED') return true;
  tmv3_audit_(
    'PreInspection','','','STEP1_SOURCE_SCOPE','SKIP',
    'PreInspection intake is authoritative only from CF Preinspects. Secondary Calendar copies are mirror-only.'
  );
  return false;
}

/** Shared Calendar hook: recorded origin creator for a promoted copy, if any. */
function tmv3_preInspectionProvenanceCreator_(vertical, cfg, calCfg, event, fallback) {
  if (vertical !== 'PreInspection' || calCfg.calendarId !== cfg.primaryCalendarId) return fallback;
  const raw = PropertiesService.getScriptProperties().getProperty('TMV3_PI_ORIGIN_' + tmv3_hash_(event.getId()));
  if (!raw) return fallback;
  let origin;
  try { origin = JSON.parse(raw); } catch (e) { throw new Error('PREINSPECTION_ORIGIN_METADATA_INVALID'); }
  if (origin.eventId !== event.getId() || (cfg.secondaryCalendarIds || []).indexOf(origin.calendarId) < 0)
    throw new Error('PREINSPECTION_ORIGIN_SCOPE_INVALID');
  const calendar = CalendarApp.getCalendarById(origin.calendarId);
  const source = calendar && calendar.getEventById(origin.eventId);
  const creators = source && source.getCreators ? source.getCreators().map(function(x) { return String(x).toLowerCase(); }) : [];
  if (!source || creators.length !== 1 || creators[0] !== origin.creator)
    throw new Error('PREINSPECTION_ORIGINAL_CREATOR_CHANGED_REVIEW');
  return origin.creator;
}

/** Today forward (lookaheadDays) from CF Preinspects only, sorted by start. */
function tmv3_preInspectionIntakeRecords_(options) {
  options = options || {};
  const cfg = TMV3.VERTICALS.PreInspection;
  const sources = tmv3_verticalCalendars_('PreInspection', cfg).filter(function(s) {
    return s.role === 'PRIMARY_SHARED' && String(s.calendarId) === String(cfg.primaryCalendarId);
  });
  if (sources.length !== 1) throw new Error('PI_AUTHORITATIVE_SOURCE_NOT_UNIQUE');
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const end = new Date(today); end.setDate(end.getDate() + Math.max(1, Number(cfg.lookaheadDays || 365)));
  const primary = CalendarApp.getCalendarById(cfg.primaryCalendarId);
  if (!primary) throw new Error('PI_PRIMARY_CALENDAR_INACCESSIBLE');
  const records = primary.getEvents(today, end).map(function(ev) {
    return tmv3_calendarEvent_('PreInspection', cfg, sources[0], ev, {applyEligibility:false});
  }).filter(function(r) { return r && !tmv3_step2PreInspectionNonCustomer_(r) && r.status !== 'cancelled'; });
  const seen = {};
  records.forEach(function(r) {
    if (!r.eventId || seen[r.eventId]) throw new Error('PI_PRIMARY_EVENT_ID_DUPLICATE_OR_BLANK');
    seen[r.eventId] = true;
  });
  // Cross-calendar discovery is published by tmv3_piCalendarReviewRefresh_, never imported here.
  return records.sort(function(a, b) { return new Date(a.start) - new Date(b.start); });
}

/** Re-read one event: primary first, then Stephen only if explicitly PreInspection. */
function tmv3_preInspectionRecoveryRecord_(eventId) {
  const cfg = TMV3.VERTICALS.PreInspection;
  const sources = tmv3_verticalCalendars_('PreInspection', cfg).concat(
    (cfg.secondaryCalendarIds || []).map(function(id) { return {calendarId:id,calendarName:'Stephen',role:'SECONDARY_STEPHEN'}; })
  );
  for (let i = 0; i < sources.length; i++) {
    const found = tmv3_findCalendarEventById_(sources[i].calendarId, eventId);
    if (!found || !found.event) continue;
    const record = tmv3_calendarEvent_('PreInspection', cfg, sources[i], found.event, {applyEligibility:false});
    if (record && !tmv3_step2PreInspectionNonCustomer_(record) &&
        (sources[i].role === 'PRIMARY_SHARED' || tmv3_preInspectionExplicitCandidate_(record))) return record;
  }
  throw new Error('RECOVERY_EVENT_MISSING_OR_NO_LONGER_PREINSPECTION');
}

/** Single-event scope for canary/targeted runs; out-of-window events are skipped unless pending. */
function tmv3_preInspectionTargetedIntake_(eventId, props) {
  const pending = JSON.parse(props.getProperty('TMV3_PI_PENDING_EVENTS') || '[]');
  const record = tmv3_preInspectionRecoveryRecord_(eventId);
  const start = new Date(record.start);
  if (!isFinite(start.getTime()) || record.status === 'cancelled') throw new Error('TARGET_EVENT_INVALID_OR_CANCELLED');
  const first = new Date(); first.setHours(0, 0, 0, 0);
  const last = new Date(first); last.setDate(last.getDate() + Number(TMV3.VERTICALS.PreInspection.lookaheadDays || 365));
  if (pending.indexOf(eventId) < 0 && (start < first || start >= last)) return [];
  return [record];
}

/** Only CF Preinspects events proceed; private-calendar candidates stay review-only. */
function tmv3_preInspectionPromoteCandidate_(record) {
  if (!record || record.vertical !== 'PreInspection') throw new Error('PI_EVENT_REQUIRED');
  if (tmv3_preInspectionApprovalHold_(record.eventId))
    return {status:'OPERATOR_INTAKE_HOLD',record:record,calendarWrites:0,guestInvites:0};
  if (tmv3_piIsAuthoritativeRecord_(record))
    return {status:'PRIMARY',record:record,calendarWrites:0,guestInvites:0};
  return {status:'SECONDARY_REVIEW_ONLY',record:record,calendarWrites:0,guestInvites:0};
}

function tmv3_preInspectionDiscoveryStatus() {
  const raw = PropertiesService.getScriptProperties().getProperty('TMV3_PI_DISCOVERY');
  if (!raw) return {status:'NOT_SCANNED',total:0,records:[]};
  try { return JSON.parse(raw); } catch (error) { return {status:'INVALID_DISCOVERY_STATE',total:0,records:[]}; }
}

/** Cross-calendar discovery is the read-only PI Calendar Review sheet. */
function tmv3_refreshPreInspectionDiscovery() {
  return tmv3_piCalendarReviewRefresh_();
}

// =====================================================================
// §3 CLASSIFY — Step 2 eligibility and PreInspection intent.
// =====================================================================

function tmv3_step2ClassifyPreInspection_(record) {
  if (tmv3_step2PreInspectionNonCustomer_(record)) {
    return tmv3_step2Decision_('SKIP', 'PREINSPECTION_NON_JOB',
      'Non-customer appointment excluded from PreInspection automation.', false);
  }
  const roles = record.sourceCalendarRoles || [record.calendarRole || ''];
  const warnings = tmv3_step2PreInspectionFormatWarnings_(record);
  if (roles.indexOf('PRIMARY_SHARED') !== -1) {
    return tmv3_step2Decision_('ELIGIBLE',
      warnings.length ? 'PREINSPECTION_SHARED_FORMAT_ATTENTION' : 'PREINSPECTION_SHARED_ELIGIBLE',
      'CF Preinspects is the authoritative PreInspection source.', false, warnings, 'PRIMARY_SOURCE');
  }
  if (roles.indexOf('SECONDARY_STEPHEN') !== -1) {
    return tmv3_step2Decision_('SKIP', 'PREINSPECTION_SECONDARY_MIRROR_ONLY',
      'Stephen Calendar is mirror-only. PreInspection intake must originate from CF Preinspects.',
      false, warnings, 'SECONDARY_NOT_AUTHORITATIVE');
  }
  return tmv3_step2Decision_('REVIEW', 'PREINSPECTION_UNKNOWN_SOURCE',
    'PreInspection event is not traceable to the authoritative CF Preinspects source.',
    false, warnings, 'UNKNOWN_SOURCE');
}

const TMV3_PI_NON_CUSTOMER_TITLE_PATTERNS = Object.freeze([
  /\bdo not book\b/i,
  /\bno more calls\b/i,
  /\bteam meeting\b/i,
  /\bnewmarket day\b/i,
  /\bnewmarket all day\b/i,
  /^newmarket$/i,
  /\ball day installing fireplaces\b/i,
  /\bon vacation\b/i,
  /\bday off\b/i,
  /\bcivic holiday\b/i,
  /^travel to\b/i,
  /\bout of office\b/i,
  /\binstall shut offs? and tags\b/i,
  /\binstall with scaffolding\b/i
]);

function tmv3_step2PreInspectionNonCustomer_(record) {
  if (record.isAllDay) return true;
  const title = tmv3_clean_(record.title);
  return TMV3_PI_NON_CUSTOMER_TITLE_PATTERNS.some(function(pattern) { return pattern.test(title); });
}

function tmv3_step2PreInspectionFormatWarnings_(record) {
  const warnings = [];
  if (!record.customerNumber) warnings.push('Title missing Customer #.');
  if (!tmv3_clean_(record.calendarCustomerName)) warnings.push('Title missing customer name.');
  if (!record.phone) warnings.push('Title missing phone.');
  if (!tmv3_preInspectionHasSalesOrderReference_(record)) {
    warnings.push('Sales Order Number missing (notification only; task creation allowed).');
  }
  if (!tmv3_clean_(record.location)) warnings.push('Calendar Location is blank.');
  return warnings;
}

/** Calendar-only intent: explicit instruction, or REVIEW. Never guesses a visit type. */
function tmv3_preInspectionDiscoveryDecision_(record, cfg) {
  if (!record || tmv3_step2PreInspectionNonCustomer_(record) || record.status === 'cancelled')
    return {status:'IGNORE',reason:'NON_CUSTOMER_OR_CANCELLED'};
  const creator = tmv3_preInspectionCreatorEmail_(record);
  const excluded = (cfg.secondaryIgnoredCreatorEmails || []).map(tmv3_normEmail_);
  if (creator && excluded.indexOf(creator) >= 0) return {status:'IGNORE',reason:'EXCLUDED_CREATOR'};
  const description = record.descriptionClean || record.rawDescription || '';
  const text = [record.title || '', description].join(' ');
  if (/\b(?:not|no)\s+(?:a\s+)?pre[\s-]*inspect(?:ion)?\b/i.test(text)) return {status:'IGNORE',reason:'EXPLICIT_NOT_PREINSPECTION'};
  if (/\bpre[\s-]*inspect(?:ion)?\s+(?:(?:is|was)\s+)?(?:not\b|required later|already completed)/i.test(text))
    return {status:'IGNORE',reason:'EXPLICIT_NOT_PREINSPECTION'};
  const instruction = /\bpre[\s-]*inspect(?:ion)?\b/i.test(record.title || '') ||
    /^\s*(?:(?:this is (?:a )?|appointment(?: type)?[: -]+))?pre[\s-]*inspect(?:ion)?\b/i.test(description);
  const hasPhone = !!tmv3_phone10_(record.phone), hasLocation = !!tmv3_clean_(record.location);
  if (!creator) return instruction || hasPhone || hasLocation ?
    {status:'REVIEW_VISIT_TYPE',reason:'CREATOR_UNVERIFIED'} : {status:'IGNORE',reason:'NO_CUSTOMER_VISIT_EVIDENCE'};
  if (instruction) return {status:'EXPLICIT_PREINSPECTION',reason:'EXPLICIT_SOURCE_INSTRUCTION'};
  if (!hasPhone || !hasLocation) return hasPhone || hasLocation ?
    {status:'REVIEW_VISIT_TYPE',reason:'INSUFFICIENT_CUSTOMER_VISIT_EVIDENCE'} :
    {status:'IGNORE',reason:'INSUFFICIENT_CUSTOMER_VISIT_EVIDENCE'};
  return {status:'REVIEW_VISIT_TYPE',reason:/\b(?:return|go back|measure|filler|trim)\b/i.test(text) ?
    'RETURN_OR_MEASURING_VISIT_UNCONFIRMED' : 'CUSTOMER_VISIT_TYPE_UNCONFIRMED'};
}

/** Reads a fresh customer order to resolve REVIEW_VISIT_TYPE where one quoted order at the job site proves intent. */
function tmv3_preInspectionClassifyLikely_(record, cfg, fresh) {
  const base = tmv3_preInspectionDiscoveryDecision_(record, cfg);
  if (base.status !== 'REVIEW_VISIT_TYPE' || base.reason === 'CREATOR_UNVERIFIED' ||
      base.reason === 'INSUFFICIENT_CUSTOMER_VISIT_EVIDENCE') return base;
  if (/\b(?:service|delivery|repair|maintenance|troubleshoot(?:ing)?)\b/i.test(record.title || '') ||
      /^\s*(?:service|delivery|repair|maintenance)\b/i.test(record.descriptionClean || record.rawDescription || ''))
    return {status:'REVIEW_VISIT_TYPE',reason:'OTHER_VERTICAL_INTENT_REQUIRES_REVIEW'};
  const key = record.eventId + '|' + record.fingerprint;
  if (!fresh && TMV3_PI_LIKELY_DECISIONS[key]) return TMV3_PI_LIKELY_DECISIONS[key];
  let result = base;
  try {
    const refs = TMV3_PI_LIKELY_REFS || (TMV3_PI_LIKELY_REFS = tmv3_step4IdentityIndex_());
    const identity = tmv3_step4ResolveCustomerFromCalendarEvidence_(record, refs);
    const customerId = Number(identity.customer && identity.customer['Customer ID']);
    if (!Number.isSafeInteger(customerId) || customerId <= 0)
      return {status:'REVIEW_VISIT_TYPE',reason:'CUSTOMER_IDENTITY_UNRESOLVED'};
    const explicit = String(record.orderNumber || '');
    // Never silently choose the first of multiple authored SO references.
    const text = [record.title, record.descriptionClean || record.rawDescription || ''].join(' ');
    const mentions = tmv3_unique_((text.match(/\b(?:SO|Sales\s+Order)\s*#?\s*\d{5,}/ig) || []).map(function(v) { return v.match(/\d+/)[0]; }));
    if (mentions.length > 1) return {status:'REVIEW_VISIT_TYPE',reason:'MULTIPLE_ORDER_REFERENCES'};
    const query = {PageIndex:0,PageSize:10,CustomerId:customerId};
    if (explicit) query.Number = explicit;
    const search = tmv3_fetchJson_(TMV3.API_BASE + '/v1/sales-orders/search',
      {method:'post',contentType:'application/json',payload:JSON.stringify(query)});
    const rows = tmv3_first_(search, ['Data']), total = Number(tmv3_first_(search, ['TotalCount']));
    if (!Array.isArray(rows) || !Number.isSafeInteger(total) || total !== rows.length || total > 10 ||
        (explicit && rows.length !== 1))
      return {status:'REVIEW_VISIT_TYPE',reason:'ORDER_SEARCH_INCOMPLETE_OR_AMBIGUOUS'};
    const locations = tmv3_preInspectionCustomerLocationsFresh_(customerId);
    const sites = locations.filter(function(l) {
      return Number(l['Customer ID']) === customerId &&
        tmv3_addressStrongMatch_(tmv3_addressParts_(record.location), tmv3_addressParts_(tmv3_locationFullAddress_(l)));
    });
    if (sites.length !== 1) return {status:'REVIEW_VISIT_TYPE',reason:'JOB_SITE_NOT_UNIQUE_FOR_CUSTOMER'};
    const relevant = [];
    rows.forEach(function(row) {
      const id = Number(tmv3_first_(row, ['Id'])), number = String(tmv3_first_(row, ['Number']));
      if (!Number.isSafeInteger(id) || id <= 0 || (explicit && number !== explicit)) throw Error('ORDER_SEARCH_IDENTITY_CONFLICT');
      const raw = tmv3_fetchJson_(TMV3.API_BASE + '/v1/sales-orders/' + id, {method:'get'});
      const customer = tmv3_first_(raw, ['Customer']), location = tmv3_first_(raw, ['ShipToLocation']), status = tmv3_first_(raw, ['Status']);
      if (Number(tmv3_first_(raw, ['Id'])) !== id || String(tmv3_first_(raw, ['OrderNumber', 'Number'])) !== number ||
          Number(tmv3_first_(customer, ['Id'])) !== customerId) throw Error('ORDER_READBACK_IDENTITY_CONFLICT');
      const statusId = String(tmv3_first_(status, ['Id'])), statusName = tmv3_norm_(tmv3_first_(status, ['Name']));
      const known = {'19':'quoted','22':'approved','25':'in progress'};
      if (!known[statusId] || known[statusId] !== statusName) return;
      if (Number(tmv3_first_(location, ['Id'])) !== Number(sites[0]['Location ID'])) return;
      relevant.push({orderId:id,orderNumber:number,customerId:customerId,
        locationId:Number(sites[0]['Location ID']),orderStatus:statusName});
    });
    const installations = relevant.filter(function(o) { return o.orderStatus === 'approved'; });
    const inProgress = relevant.filter(function(o) { return o.orderStatus === 'in progress'; });
    if (inProgress.length) result = {status:'REVIEW_VISIT_TYPE',reason:'IN_PROGRESS_ORDER_TYPE_UNVERIFIED',orders:inProgress};
    else if (installations.length) result = {status:'INSTALLATION',reason:'APPROVED_ORDER_AT_JOB_SITE',orders:installations};
    else if (relevant.length === 1 && relevant[0].orderStatus === 'quoted') {
      // Measuring return visits already booked for install are not initial inspections.
      if (/\binstall(?:ation)?\b/i.test(record.title || '') ||
          (/\b(?:return|go back)\b/i.test(text) && /\binstall(?:ation)?\b/i.test(text)))
        result = {status:'REVIEW_VISIT_TYPE',reason:'INSTALLATION_RETURN_VISIT_ORDER_NOT_APPROVED'};
      else result = {status:'EXPLICIT_PREINSPECTION',reason:'VERIFIED_QUOTED_ORDER_LIKELY_PREINSPECTION',orders:relevant};
    } else result = {status:'REVIEW_VISIT_TYPE',reason:'NO_UNIQUE_ACTIVE_ORDER_AT_JOB_SITE'};
  } catch (error) {
    result = {status:'REVIEW_VISIT_TYPE',reason:'LIVE_ORDER_CLASSIFICATION_FAILED',error:String(error).slice(0, 240)};
  }
  TMV3_PI_LIKELY_DECISIONS[key] = result;
  return result;
}

function tmv3_preInspectionExplicitCandidate_(record) {
  return !!record && tmv3_preInspectionClassifyLikely_(record, TMV3.VERTICALS.PreInspection).status === 'EXPLICIT_PREINSPECTION';
}

// =====================================================================
// §4 ANCHOR — Step 3: exact Customer #, deterministic SO corroboration,
// or a complete identity bundle (name + phone + address) for Step 4.
// =====================================================================

// Original: 30_Calendar; Function: tmv3_extractLegacyPreInspectionOrderNumber_
function tmv3_extractLegacyPreInspectionOrderNumber_(description) {
  const clean = tmv3_stripPhonesForOrderParsing_(description);
  if (!clean) return '';
  const match = clean.match(/\b(?:SO|S\/O|Sales\s*Order|Order)\s*(?:#|No\.?|Number)?\s*[:\-]?\s*(\d{5,8})\b/i);
  return match ? match[1] : '';
}

// Original: 30_Calendar; Function: tmv3_preInspectionCalendarCustomerName_
function tmv3_preInspectionCalendarCustomerName_(title, customerNumber) {
  let value = tmv3_clean_(title);
  if (!value) return '';
  value = value.replace(/(?:\+?1[\s.\-]?)?\(?[2-9]\d{2}\)?[\s.\-]?[2-9]\d{2}[\s.\-]?\d{4}/g, ' ');
  if (customerNumber) {
    const digits = String(customerNumber).replace(/[^0-9]/g, '');
    value = value
      .replace(new RegExp('(^|\\s)(?:C|Cust(?:omer)?)\\s*#?\\s*' + digits + '(?=\\s|[-–—:]|$)', 'ig'), ' ')
      .replace(new RegExp('(^|\\s)#?' + digits + '(?=\\s|[-–—:]|$)', 'g'), ' ');
  }
  return value
    .replace(/\bpre\s*-?inspect(?:ion)?\b/ig, ' ')
    .replace(/\s*[-–—:+&]+\s*/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function tmv3_step3AnchorFromPreInspection_(record, customer, order) {
  return {
    anchorType: customer ? 'CUSTOMER' : (order ? 'SALES_ORDER_EVIDENCE' : ''),
    customerId: tmv3_clean_(customer && customer['Customer ID']),
    customerNumber: tmv3_clean_(customer && customer['Customer Number']),
    customerName: tmv3_clean_(customer && customer['Name']),
    orderId: tmv3_clean_(order && order['Order ID']),
    orderNumber: tmv3_clean_(order && order['Order Number']),
    orderType: tmv3_clean_(order && order['Order Type']),
    orderStatus: tmv3_clean_(order && order['Status']),
    orderUrl: tmv3_clean_(order && order['URL']),
    orderLocationIdEvidence: tmv3_clean_(order && order['Location ID']),
    orderContactIdEvidence: tmv3_clean_(order && order['Contact ID']),
    orderEvidenceOnly: true
  };
}

function tmv3_step3ResolvePreInspectionAnchor_(record, refs) {
  const calendarCustomerNumber = tmv3_clean_(record.customerNumber);
  const evidence = [];
  const warnings = [];
  let customer = null;
  let unresolvedCalendarCustomerNumber = '';

  if (calendarCustomerNumber) {
    customer = refs.customerByNumber[calendarCustomerNumber] || null;
    if (customer) {
      evidence.push('CALENDAR_CUSTOMER_NUMBER_EXACT');
    } else {
      unresolvedCalendarCustomerNumber = calendarCustomerNumber;
      warnings.push('Calendar number ' + calendarCustomerNumber +
        ' did not match a Customer #; it will be tested only as deterministic SO evidence.');
    }
  }

  const orderResolution = tmv3_step3ResolvePreInspectionOrderEvidence_(record, refs, unresolvedCalendarCustomerNumber);
  if (orderResolution.ambiguous) {
    return tmv3_step3Decision_('REVIEW', 'PREINSPECTION_SO_AMBIGUOUS',
      'Multiple distinct Striven orders are supported by the Calendar numeric evidence.',
      customer ? tmv3_step3AnchorFromPreInspection_(record, customer, null) : null,
      evidence.concat(orderResolution.evidence), warnings.concat(orderResolution.warnings));
  }

  const order = orderResolution.order;
  Array.prototype.push.apply(evidence, orderResolution.evidence);
  Array.prototype.push.apply(warnings, orderResolution.warnings);

  if (customer && order) {
    const orderCustomerId = tmv3_clean_(order['Customer ID']);
    const customerId = tmv3_clean_(customer['Customer ID']);
    if (orderCustomerId && customerId && orderCustomerId !== customerId) {
      return tmv3_step3Decision_('REVIEW', 'PREINSPECTION_CUSTOMER_SO_CONFLICT',
        'Calendar Customer # and Calendar SO resolve to different Striven Customers.',
        tmv3_step3AnchorFromPreInspection_(record, customer, order), evidence, warnings);
    }
    if (orderCustomerId && customerId) evidence.push('SO_CORROBORATES_CUSTOMER');
  }

  if (!customer && order) {
    const orderCustomerId = tmv3_clean_(order['Customer ID']);
    customer = orderCustomerId ? (refs.customerById[orderCustomerId] || null) : null;
    if (customer) evidence.push('UNIQUE_SO_RESOLVES_CUSTOMER');
  }

  if (!customer) {
    const identityBundleReady =
      !unresolvedCalendarCustomerNumber &&
      !!tmv3_clean_(record.calendarCustomerName) &&
      !!tmv3_phone10_(record.phone) &&
      !!tmv3_clean_(record.location);
    if (identityBundleReady) {
      const anchor = tmv3_step3AnchorFromPreInspection_(record, null, order);
      anchor.anchorType = 'IDENTITY_EVIDENCE';
      anchor.customerName = tmv3_clean_(record.calendarCustomerName);
      return tmv3_step3Decision_('VERIFIED', 'PREINSPECTION_IDENTITY_EVIDENCE_READY',
        'No deterministic Customer/Order anchor is available, but Calendar name + phone + address are sufficient to proceed to Step 4 identity proof.',
        anchor, evidence.concat(['CALENDAR_IDENTITY_BUNDLE_READY']), warnings);
    }
    return tmv3_step3Decision_('REVIEW', 'PREINSPECTION_BUSINESS_ANCHOR_UNRESOLVED',
      'No exact Customer #, unique SO relationship, or complete identity bundle is available.',
      tmv3_step3AnchorFromPreInspection_(record, null, order), evidence, warnings);
  }

  // Exact Customer # is already a deterministic anchor; SO is corroboration only.
  return tmv3_step3Decision_('VERIFIED',
    order ? 'PREINSPECTION_CUSTOMER_SO_VERIFIED' : 'PREINSPECTION_CUSTOMER_VERIFIED',
    order ? 'PreInspection Customer resolved and the Calendar SO evidence points to the same Customer.'
          : 'PreInspection Customer resolved exactly; SO corroboration is unavailable or unresolved.',
    tmv3_step3AnchorFromPreInspection_(record, customer, order), evidence, warnings);
}

function tmv3_step3ResolvePreInspectionOrderEvidence_(record, refs, unresolvedCalendarCustomerNumber) {
  const candidateNumbers = [];
  const evidence = [];
  const warnings = [];
  const parsed = tmv3_clean_(record.orderNumber);
  if (parsed) candidateNumbers.push(parsed);
  // A six-digit leading number that failed Customer lookup is a common
  // legacy pattern where the title actually carried the Sales Order.
  if (unresolvedCalendarCustomerNumber && /^\d{6}$/.test(unresolvedCalendarCustomerNumber)) {
    candidateNumbers.push(unresolvedCalendarCustomerNumber);
  }
  if (!parsed) {
    tmv3_step3BareSixDigitOrderCandidates_(record).forEach(function(number) { candidateNumbers.push(number); });
  }
  const matchedOrders = [];
  tmv3_unique_(candidateNumbers.map(tmv3_clean_).filter(Boolean)).forEach(function(number) {
    const matches = tmv3_step3UniqueOrders_((refs.ordersByNumber[number] || []).filter(function(order) {
      return tmv3_step3OrderMatchesVertical_(order, 'PreInspection');
    }).slice());
    if (matches.length !== 1) return;
    matchedOrders.push(matches[0]);
    if (number === parsed) evidence.push('CALENDAR_SO_NUMBER_EXACT');
    else if (number === unresolvedCalendarCustomerNumber) evidence.push('LEADING_NUMBER_RECOVERED_AS_SO');
    else evidence.push('LEGACY_BARE_SIX_DIGIT_SO_EXACT');
  });
  const orders = tmv3_step3UniqueOrders_(matchedOrders);
  if (orders.length > 1) return {ambiguous:true,order:null,evidence:tmv3_unique_(evidence),warnings:warnings};
  if (orders.length === 1) return {ambiguous:false,order:orders[0],evidence:tmv3_unique_(evidence),warnings:warnings};
  warnings.push(parsed ? 'Calendar SO #' + parsed + ' was not found in the current Order source.'
                       : 'No deterministic SO corroboration was found.');
  return {ambiguous:false,order:null,evidence:tmv3_unique_(evidence),warnings:warnings};
}

// =====================================================================
// §5 IDENTITY — Step 4: verified Customer and Customer-owned Location.
// Customer Contact is identity evidence only, never Requested By.
// =====================================================================

function tmv3_step4ResolvePreInspectionIdentity_(record, refs, anchor) {
  const evidence = [];
  const warnings = [];
  let customer = null;
  const anchorCustomerId = tmv3_clean_(anchor.customerId);
  if (anchorCustomerId && refs.customerById[anchorCustomerId]) {
    customer = refs.customerById[anchorCustomerId];
    evidence.push('PREINSPECTION_CUSTOMER_FROM_STEP3');
  }
  if (!customer) {
    const recovery = tmv3_step4ResolveCustomerFromCalendarEvidence_(record, refs);
    if (recovery.status !== 'MATCHED') {
      return tmv3_step4Decision_('REVIEW', recovery.errorCode, recovery.reason,
        recovery.customer || null, recovery.location || null, null, 'NOT_RESOLVED',
        evidence.concat(recovery.evidence || []), warnings);
    }
    customer = recovery.customer;
    Array.prototype.push.apply(evidence, recovery.evidence || []);
  }

  const locationResult = tmv3_step4ResolveLocation_(record, refs, customer, anchor.orderLocationIdEvidence);
  if (locationResult.status !== 'MATCHED' && locationResult.status !== 'CREATE_REQUIRED') {
    return tmv3_step4Decision_('REVIEW', locationResult.errorCode, locationResult.reason,
      customer, null, null, 'NOT_RESOLVED', evidence.concat(locationResult.evidence || []), warnings);
  }
  if (locationResult.status === 'CREATE_REQUIRED') {
    warnings.push('Location will need to be created on the verified Customer before any Task write.');
  }

  const contactResult = tmv3_step4ResolveContact_(record, customer, '');
  if (contactResult.status === 'REVIEW') {
    warnings.push(contactResult.reason || 'Customer Contact requires review; not used as Requested By.');
  } else if (contactResult.status === 'NO_MATCH') {
    warnings.push('No unique Contact match. Requested By uses the verified Calendar owner employee, not Contact.');
  }

  return tmv3_step4Decision_('VERIFIED', 'PREINSPECTION_IDENTITY_VERIFIED',
    'PreInspection Customer and Customer-owned Location are verified.',
    customer, locationResult.location,
    contactResult.status === 'MATCHED' ? contactResult.contact : null,
    contactResult.status,
    evidence.concat(locationResult.evidence || [], contactResult.evidence || []),
    warnings);
}

/** Fresh GET of a Customer's Locations; incomplete or foreign-owned rows fail closed. */
function tmv3_preInspectionCustomerLocationsFresh_(customerId) {
  const id = Number(customerId);
  if (!Number.isInteger(id) || id <= 0) throw new Error('LOCATION_CUSTOMER_ID_INVALID');
  const response = tmv3_fetchJson_(TMV3.API_BASE + '/v1/customers/' + encodeURIComponent(id) + '/locations', {method:'get'});
  const rawData = response && (response.Data !== undefined ? response.Data : response.data);
  const count = response && (response.TotalCount !== undefined ? response.TotalCount : response.totalCount);
  // Live Striven represents an explicitly empty collection as count=0, data=null.
  const data = count === 0 && rawData === null ? [] : rawData;
  if (!response || !Array.isArray(data) || !Number.isInteger(count) || count !== data.length) {
    throw new Error('LOCATION_RESPONSE_INCOMPLETE_NO_CREATE');
  }
  const seen = {};
  return data.map(function(raw) {
    const customer = raw && (raw.Customer || raw.customer);
    const address = raw && (raw.Address || raw.address);
    const row = {Id:raw && (raw.Id !== undefined ? raw.Id : raw.id),
      Customer:customer && {Id:customer.Id !== undefined ? customer.Id : customer.id},
      Address:address && {Address1:address.Address1 || address.address1,FullAddress:address.FullAddress || address.fullAddress,
        Address2:address.Address2 || address.address2,City:address.City || address.city,
        State:address.State || address.state,PostalCode:address.PostalCode || address.postalCode}};
    if (!raw || !Number.isInteger(row.Id) || row.Id <= 0 || seen[row.Id] ||
        !row.Customer || Number(row.Customer.Id) !== id || !row.Address ||
        !tmv3_clean_(row.Address.Address1 || row.Address.FullAddress)) {
      throw new Error('LOCATION_RESPONSE_IDENTITY_INVALID_NO_CREATE');
    }
    seen[row.Id] = true;
    const a = row.Address;
    return tmv3_normalizedLocationObject_(tmv3_normalizeLocation_({
      Id:row.Id,CustomerId:id,Address1:a.Address1 || a.FullAddress,Address2:a.Address2 || '',
      City:a.City || '',State:a.State || '',PostalCode:a.PostalCode || ''
    }, {}));
  });
}

/** True only for the two known Striven 400 schema rejections of a never-created Location. */
function tmv3_preInspectionRejectedAddressGuard_(guard, preflight) {
  if (!guard || guard.state !== 'UNCERTAIN' || guard.locationId ||
      !preflight || !Array.isArray(preflight.owned) || !Array.isArray(preflight.matches) ||
      preflight.matches.length !== 0) return false;
  const error = String(guard.error || '');
  if (!/HTTP 400\b/.test(error)) return false;
  const oldAddress = !guard.addressSchema && /Address object is required/i.test(error);
  const oldCountry = guard.addressSchema === 'nested-address-v1' &&
    /country_Invalid/.test(error) && /country code Canada could not be read/i.test(error);
  return oldAddress || oldCountry;
}

/** When Step 4 finds ambiguous strong Locations, a unique explicit SO ship-to may repair the cache. */
function tmv3_preInspectionOptionalOrderLocationEvidence_(record) {
  const gate = tmv3_piStagingDisabled_(); if (gate) return gate;
  if (!record || record.vertical !== 'PreInspection' || !/^\d+$/.test(String(record.orderNumber || ''))) return {status:'NOT_REQUIRED'};
  const s2 = tmv3_step2CalendarRecords_([record]);
  const s3 = tmv3_step3BusinessAnchorRecords_(s2, tmv3_step3AnchorIndex_());
  const s4 = tmv3_step4IdentityRecords_(s3, tmv3_step4IdentityIndex_());
  if (s4.length !== 1 || !s4[0].step4 || s4[0].step4.code !== 'AMBIGUOUS_STRONG_LOCATION') return {status:'NOT_REQUIRED'};
  const customerId = Number(s4[0].step4.customer && s4[0].step4.customer['Customer ID']);
  if (!Number.isInteger(customerId) || customerId <= 0) return {status:'CUSTOMER_UNRESOLVED'};
  const number = String(record.orderNumber);
  const found = tmv3_fetchJson_(TMV3.API_BASE + '/v1/sales-orders/search', {method:'post',contentType:'application/json',
    payload:JSON.stringify({PageIndex:0,PageSize:2,Number:number,CustomerId:customerId})});
  const data = found && (found.Data || found.data);
  const count = found && (found.TotalCount !== undefined ? found.TotalCount : found.totalCount);
  if (!found || !Array.isArray(data) || count !== data.length || data.length !== 1) return {status:'ORDER_NOT_UNIQUE'};
  const rawCandidate = data[0];
  const candidate = {Id:tmv3_first_(rawCandidate, ['Id']),Number:tmv3_first_(rawCandidate, ['Number']),
    Customer:{Id:tmv3_first_(tmv3_first_(rawCandidate, ['Customer']), ['Id'])}};
  if (String(candidate.Number) !== number || !candidate.Customer || Number(candidate.Customer.Id) !== customerId ||
      !Number.isInteger(candidate.Id) || candidate.Id <= 0) return {status:'ORDER_IDENTITY_CONFLICT'};
  const rawOrder = tmv3_fetchJson_(TMV3.API_BASE + '/v1/sales-orders/' + candidate.Id, {method:'get'});
  const order = {Id:tmv3_first_(rawOrder, ['Id']),OrderNumber:tmv3_first_(rawOrder, ['OrderNumber']),OrderName:tmv3_first_(rawOrder, ['OrderName']),
    Customer:{Id:tmv3_first_(tmv3_first_(rawOrder, ['Customer']), ['Id'])},
    ShipToLocation:{Id:tmv3_first_(tmv3_first_(rawOrder, ['ShipToLocation']), ['Id'])},
    Contact:{Id:tmv3_first_(tmv3_first_(rawOrder, ['Contact']), ['Id'])},
    Status:{Name:tmv3_first_(tmv3_first_(rawOrder, ['Status']), ['Name'])}};
  if (!order || order.Id !== candidate.Id || String(order.OrderNumber) !== number || !order.Customer ||
      Number(order.Customer.Id) !== customerId) return {status:'ORDER_READBACK_CONFLICT'};
  const locationId = Number(order.ShipToLocation && order.ShipToLocation.Id);
  const owned = tmv3_preInspectionCustomerLocationsFresh_(customerId);
  const matches = owned.filter(function(location) {
    return Number(location['Location ID']) === locationId &&
      tmv3_addressStrongMatch_(tmv3_addressParts_(record.location), tmv3_addressParts_(tmv3_locationFullAddress_(location)));
  });
  if (matches.length !== 1) return {status:'ORDER_LOCATION_CONFLICT'};
  const row = tmv3_normalizeOrder_({Id:order.Id,OrderNumber:number,CustomerId:customerId,LocationId:locationId,
    ContactId:order.Contact && order.Contact.Id,Status:order.Status && order.Status.Name,Name:order.OrderName}, 'SALES_ORDER');
  const sheet = tmv3_sheet_(TMV3.SHEETS.ORDERS), values = sheet.getDataRange().getValues();
  if (!values.length || values[0][0] !== 'Order ID' || values[0][1] !== 'Order Number') throw new Error('ORDER_CACHE_SCHEMA_NO_WRITE');
  const existing = [];
  values.slice(1).forEach(function(value, index) {
    if (String(value[1]) === number && value[7] === 'SALES_ORDER') existing.push({value:value,row:index + 2});
  });
  if (existing.length > 1 || existing.some(function(item) {
    return Number(item.value[0]) !== order.Id || Number(item.value[2]) !== customerId;
  })) return {status:'ORDER_CACHE_CONFLICT'};
  const cachedLocations = tmv3_rows_(TMV3.SHEETS.LOCATIONS).filter(function(item) { return Number(item['Location ID']) === locationId; });
  if (cachedLocations.length > 1 || cachedLocations.some(function(item) { return Number(item['Customer ID']) !== customerId; }))
    return {status:'LOCATION_CACHE_CONFLICT'};
  tmv3_upsertLocationCacheObject_(matches[0]);
  sheet.getRange(existing.length ? existing[0].row : values.length + 1, 1, 1, row.length).setValues([row]);
  tmv3_audit_('PreInspection', record.eventId, '', 'OPTIONAL_ORDER_LOCATION_EVIDENCE', 'PASS',
    JSON.stringify({orderId:order.Id,customerId:customerId,locationId:locationId,orderAttachedToTask:false}));
  return {status:'CACHE_REPAIRED',orderId:order.Id,locationId:locationId};
}

// =====================================================================
// §6 TASK MATCH — Step 5. Customer ID alone never proves that an OPEN Type
// 105 Task belongs to this appointment; zero cached matches never proves
// that no Task exists. Pagination and ownership conflicts fail closed.
// =====================================================================

function tmv3_preInspectionTaskRowArray_(task) {
  return [
    task['Task ID'], task['Task Number'], task['Task Type ID'], task['Task Type'],
    task['Status'], task['Name'], task['Customer ID'], task['Location ID'],
    task['Contact ID'], task['Order ID'], task['Start'], task['Due'],
    task['Assignees'], task['Pools'], task['URL'], task.Fingerprint
  ];
}

/** Customers (and linked Task IDs) of Step-4-verified PreInspection events. */
function tmv3_preInspectionTaskRefreshTargets_() {
  const step2Records = tmv3_step2CalendarRecords_().filter(function(record) {
    return record.vertical === 'PreInspection';
  });
  const step3Records = tmv3_step3BusinessAnchorRecords_(step2Records, tmv3_step3AnchorIndex_());
  const step4Records = tmv3_step4IdentityRecords_(step3Records, tmv3_step4IdentityIndex_());
  return step4Records.filter(function(record) {
    return record.step4 && record.step4.disposition === 'VERIFIED' && record.step4.customer;
  }).map(function(record) {
    return {eventId:tmv3_clean_(record.eventId),existingTaskId:tmv3_clean_(record.existingTaskId),
      customer:record.step4.customer};
  });
}

/** All Type 105 Task IDs for a Customer; throws unless pagination provably ended. */
function tmv3_searchPreInspectionTaskIdsForCustomer_(customer) {
  const customerId = Number(tmv3_clean_(customer && customer['Customer ID']) || 0);
  if (!customerId) throw new Error('PreInspection task search requires Customer ID.');
  const customerPayload = {Id:customerId};
  const customerNumber = tmv3_clean_(customer['Customer Number']);
  const customerName = tmv3_clean_(customer['Name']);
  if (customerNumber) customerPayload.Number = customerNumber;
  if (customerName) customerPayload.Name = customerName;
  const maxPages = Math.max(1, Math.min(3, Number(TMV3.OPERATIONS && TMV3.OPERATIONS.preInspectionTaskSearchMaxPages || 3)));
  const rowId = function(row) { return tmv3_clean_(tmv3_first_(row, ['id','Id','taskId','TaskId'])); };
  const ids = [];
  let previousSignature = '';
  let terminatedNaturally = false;
  for (let pageIndex = 0; pageIndex < maxPages; pageIndex++) {
    const json = tmv3_fetchJson_(TMV3.API_BASE + '/v2/tasks/search', {
      method:'post',contentType:'application/json',
      payload:JSON.stringify({Customer:customerPayload,Type:[105],PageIndex:pageIndex,PageSize:100})
    }) || {};
    const rows = tmv3_extractTaskSearchRows_(json);
    const signature = rows.slice(0, 5).map(rowId).join('|');
    if (pageIndex > 0 && signature && signature === previousSignature) {
      throw new Error('TMV3_PREINSPECTION_CUSTOMER_TASK_PAGINATION_REPEATED: Customer ' +
        customerId + ' repeated page ' + pageIndex + '.');
    }
    previousSignature = signature;
    rows.forEach(function(row) { const id = rowId(row); if (id) ids.push(id); });
    if (!rows.length || rows.length < 100) { terminatedNaturally = true; break; }
  }
  if (!terminatedNaturally) {
    throw new Error('TMV3_PREINSPECTION_CUSTOMER_TASK_PAGINATION_TRUNCATED: Customer ' +
      customerId + ' reached the page cap without proving completion.');
  }
  return tmv3_unique_(ids);
}

/** Exhaustive, ownership-verified fresh Type 105 history for one Customer. */
function tmv3_searchPreInspectionTasks_(customer) {
  const customerId = String(tmv3_clean_(customer && customer['Customer ID']));
  if (!tmv3_piPositiveId_(customerId)) throw new Error('PI_CUSTOMER_ID_REQUIRED');
  const ids = tmv3_searchPreInspectionTaskIdsForCustomer_(customer);
  if (ids.length > tmv3_piTaskHydrationCap_()) throw new Error('PI_HISTORY_HYDRATION_CAP_REVIEW_NO_CREATE');
  const seen = {};
  return ids.map(function(id) {
    const key = String(id);
    if (seen[key]) throw new Error('PI_HISTORY_DUPLICATE_TASK_ID');
    seen[key] = true;
    const task = tmv3_getTaskById_(id);
    if (!task || String(task['Task ID']) !== key || Number(task['Task Type ID']) !== 105 ||
        String(task['Customer ID']) !== customerId) {
      throw new Error('PI_HISTORY_TASK_OWNERSHIP_CONFLICT_NO_CREATE:' + key);
    }
    return task;
  });
}

/** Shared Striven-cache hook: bounded targeted refresh that never erases unrelated rows. */
function tmv3_preInspectionTaskSourceRows_(options) {
  options = options || {};
  const prior = tmv3_rows_(TMV3.SHEETS.TASKS).filter(function(t) { return Number(t['Task Type ID']) === 105; });
  const props = PropertiesService.getScriptProperties();
  const key = 'TMV3_PREINSPECTION_TASK_CACHE_REFRESH_AT';
  const lastMs = Number(props.getProperty(key) || 0);
  const ttl = Math.max(60, Number((TMV3.OPERATIONS || {}).preInspectionTaskCacheMinMinutes || 360));
  const age = lastMs ? (Date.now() - lastMs) / 60000 : Infinity;
  const summary = {reused:true,targetCustomers:0,searchedCustomers:0,hydratedTasks:0,priorTasks:prior.length,
    ageMinutes:isFinite(age) ? Math.round(age) : null};
  if (options.force !== true && prior.length && age < ttl) {
    return Object.assign({rows:prior.map(tmv3_preInspectionTaskRowArray_)}, summary);
  }
  const targets = tmv3_preInspectionTaskRefreshTargets_();
  if (!targets.length) return Object.assign({rows:prior.map(tmv3_preInspectionTaskRowArray_)}, summary);
  const customers = {}, wanted = {};
  targets.forEach(function(target) {
    const id = String(tmv3_clean_(target.customer && target.customer['Customer ID']));
    if (id) customers[id] = target.customer;
    if (target.existingTaskId) wanted[String(target.existingTaskId)] = true;
  });
  const customerIds = Object.keys(customers);
  customerIds.forEach(function(customerId) {
    tmv3_searchPreInspectionTaskIdsForCustomer_(customers[customerId]).forEach(function(taskId) {
      wanted[String(taskId)] = true;
    });
  });
  const ids = Object.keys(wanted);
  if (ids.length > tmv3_piTaskHydrationCap_()) throw new Error('PI_CACHE_HYDRATION_CAP_NO_REPLACE');
  const byId = {};
  // This refresh is not an all-customer snapshot; keep prior tasks outside the targets.
  prior.forEach(function(task) {
    const id = String(tmv3_clean_(task['Task ID']));
    if (!id || byId[id]) throw new Error('PI_PRIOR_CACHE_DUPLICATE_OR_BLANK_TASK_ID');
    byId[id] = task;
  });
  ids.forEach(function(id) {
    const task = tmv3_getTaskById_(id);
    const owner = String(tmv3_clean_(task && task['Customer ID']));
    if (!task || String(tmv3_clean_(task['Task ID'])) !== id || Number(task['Task Type ID']) !== 105 || !customers[owner]) {
      throw new Error('PI_CACHE_TASK_IDENTITY_CONFLICT_NO_REPLACE:' + id);
    }
    byId[id] = task;
  });
  // Commit freshness only after the entire bounded retrieval succeeds.
  props.setProperty(key, String(Date.now()));
  return {rows:Object.keys(byId).sort(function(a, b) { return Number(a) - Number(b); })
      .map(function(id) { return tmv3_preInspectionTaskRowArray_(byId[id]); }),
    reused:false,targetCustomers:customerIds.length,searchedCustomers:customerIds.length,
    hydratedTasks:ids.length,priorTasks:prior.length,ageMinutes:isFinite(age) ? Math.round(age) : null};
}

/** Per-candidate match evidence for one OPEN/closed Type 105 Task. */
function tmv3_piEvaluateTaskCandidate_(task, eventRecord, customer, lid, linkedId) {
  const taskId = String(task['Task ID']);
  const sameStart = !!task.Start && tmv3_sameMinute_(eventRecord.start, task.Start);
  const sameDue = !!task.Due && tmv3_sameMinute_(eventRecord.end, task.Due);
  const sameDay = !!task.Start && tmv3_sameLocalDay_(eventRecord.start, task.Start);
  const sameLocation = !!lid && String(task['Location ID']) === lid;
  const taskLocation = String(tmv3_clean_(task['Location ID']));
  const phone = tmv3_phone10_(eventRecord.phone);
  const titlePhoneMatch = !!phone && tmv3_allPhones_(String(task.Name || '')).indexOf(phone) >= 0;
  const name = tmv3_norm_(customer.Name);
  const titleNameMatch = !!name && tmv3_norm_(task.Name).indexOf(name) >= 0;
  const explicit = !!linkedId && taskId === linkedId;
  const locationConflict = !!lid && !!taskLocation && taskLocation !== lid;
  const corroborated = titlePhoneMatch || titleNameMatch;
  const strong = (!locationConflict) && (
    (explicit && (sameDay || (sameLocation && corroborated))) ||
    (sameStart && sameLocation && corroborated) ||
    (sameDay && sameDue && sameLocation && corroborated)
  );
  return {task:task,open:tmv3_taskIsOpen_(task.Status),completed:tmv3_taskIsCompleted_(task.Status),
    sameStart:sameStart,sameDue:sameDue,sameDay:sameDay,sameLocation:sameLocation,
    titlePhoneMatch:titlePhoneMatch,titleNameMatch:titleNameMatch,identityCorroborated:corroborated,
    exactCalendarTaskLink:explicit,strongMatch:strong,locationConflict:locationConflict};
}

function tmv3_preInspectionTaskDecision_(eventRecord, customer, location, cachedCandidates) {
  const cid = String(tmv3_clean_(customer && customer['Customer ID']));
  const lid = String(tmv3_clean_(location && location['Location ID']));
  const linkedId = String(tmv3_clean_(eventRecord && eventRecord.existingTaskId));
  if (!tmv3_piPositiveId_(cid)) return {status:'REVIEW',errorCode:'PI_CUSTOMER_UNVERIFIED'};
  let candidates;
  if (Array.isArray(cachedCandidates)) candidates = cachedCandidates.slice();
  else if (linkedId) candidates = [tmv3_getTaskById_(linkedId)];
  else candidates = tmv3_searchPreInspectionTasks_(customer);
  if (linkedId && !candidates.some(function(t) { return String(t && t['Task ID']) === linkedId; })) {
    return {status:'REVIEW',errorCode:'PI_EXPLICIT_TASK_LINK_NOT_RESOLVED',
      reason:'Linked Task is absent; no fallback to another Task and no create.'};
  }
  const all = (candidates || []).filter(function(t) {
    return t && Number(t['Task Type ID']) === 105 && String(t['Customer ID']) === cid;
  });
  if (linkedId && !all.some(function(t) { return String(t['Task ID']) === linkedId; })) {
    return {status:'REVIEW',errorCode:'PI_LINKED_TASK_OWNERSHIP_OR_TYPE_CONFLICT'};
  }
  const evaluated = all.map(function(task) {
    return tmv3_piEvaluateTaskCandidate_(task, eventRecord, customer, lid, linkedId);
  });
  const history = evaluated.filter(function(x) { return x.completed; }).map(function(x) { return x.task; });
  const active = evaluated.filter(function(x) { return !x.open && !x.completed; });
  const open = evaluated.filter(function(x) { return x.open; });
  if (active.length) return {status:'REVIEW',historyTasks:history,
    errorCode:'PI_NONOPEN_ACTIVE_TASK_REQUIRES_REVIEW',
    reason:'A non-completed, non-open Task exists; lifecycle must be inspected.'};
  if (linkedId && !open.some(function(x) { return x.exactCalendarTaskLink; })) {
    return {status:'REVIEW',historyTasks:history,errorCode:'PI_EXPLICIT_TASK_LINK_NOT_OPEN_REVIEW',
      reason:'Calendar-linked Task is not open; verify fulfilled history before any create.'};
  }
  if (!open.length) return {status:'CLEAR',task:null,historyTasks:history,
    reason:'No open customer Task; completed Task history remains subject to fresh verification.'};
  if (open.length > 1) return {status:'REVIEW',historyTasks:history,
    errorCode:'PI_MULTIPLE_OPEN_TASKS_NO_AUTOMATIC_SELECTION',
    reason:'Multiple open Task candidates require event-specific reconciliation.'};
  if (!open[0].strongMatch) return {status:'REVIEW',historyTasks:history,
    errorCode:'PI_OPEN_TASK_APPOINTMENT_NOT_PROVEN',
    reason:'A single open Task for the customer is not enough to overwrite its relationship.'};
  return {status:'MATCHED',task:open[0].task,historyTasks:history,
    reason:'Exactly one customer-owned open Task is corroborated to this event.',
    evidence:tmv3_preInspectionCandidateEvidence_(open[0])};
}

function tmv3_preInspectionCandidateEvidence_(candidate) {
  const parts = [];
  if (candidate.exactCalendarTaskLink) parts.push('CALENDAR_TASK_LINK');
  if (candidate.sameStart) parts.push('START');
  if (candidate.sameDue) parts.push('DUE');
  if (candidate.sameDay) parts.push('DAY');
  if (candidate.sameLocation) parts.push('LOCATION');
  if (candidate.titlePhoneMatch) parts.push('PHONE');
  if (candidate.titleNameMatch) parts.push('NAME');
  return parts;
}

/** Shared Step 5 hook. Cache-only evidence; never certified as fresh history. */
function tmv3_step5ResolvePreInspection_(record, refs) {
  const identity = record.step4 || {};
  const cid = String(tmv3_clean_(identity.customer && identity.customer['Customer ID']));
  const cached = ((refs.tasksByCustomer || {})[cid] || []).filter(function(task) {
    return Number(task['Task Type ID']) === 105;
  });
  const linked = String(tmv3_clean_(record.existingTaskId));
  const evidence = ['SHARED_CACHE_ONLY_NOT_FULL_FRESH_HISTORY'];
  if (linked && !(refs.taskById || {})[linked]) {
    return tmv3_step5Decision_('REVIEW', 'PI_LINKED_TASK_NOT_CACHED_NO_FALLBACK',
      'Calendar carries a Task ID missing from the cache; hydrate exact ID before proceeding.',
      [], 0, evidence, ['Task ID not certified from fresh Striven.']);
  }
  try {
    const d = tmv3_preInspectionTaskDecision_(record, identity.customer, identity.location, cached);
    const history = d.historyTasks || [];
    if (d.status === 'MATCHED') return tmv3_step5Decision_('MATCHED', 'PREINSPECTION_TASK_MATCHED',
      d.reason, [d.task], history.length, evidence.concat(d.evidence || []), [], history);
    if (d.status === 'CLEAR') return tmv3_step5Decision_('NO_TASK', 'PREINSPECTION_NO_OPEN_TASK',
      d.reason, [], history.length, evidence, [], history);
    return tmv3_step5Decision_('REVIEW', d.errorCode || 'PI_TASK_MATCH_UNCERTAIN',
      d.reason || 'Task relationship requires review.', [], history.length, evidence, [], history);
  } catch (error) {
    return tmv3_step5Decision_('REVIEW', 'PI_TASK_RESOLUTION_FAILED',
      String(error && error.message || error), [], 0, evidence, []);
  }
}

/** Never invent a corrected Striven timestamp from a 12-hour discrepancy. */
function tmv3_step5NormalizePreInspectionTaskTimesAcrossRecords_(records) {
  return records;
}

/** Evidence-preserving read-only copy; no AM/PM adjustment without a verified contract. */
function tmv3_step5AlignPreInspectionTaskClock_(task, record) {
  return Object.assign({}, task || {});
}

/** Shared Step 7 hook: fresh-GET each Calendar-linked Task missing from the cache. */
function tmv3_step7HydrateLinkedPreInspectionTasks_(records, refs, runtime) {
  (records || []).forEach(function(record) {
    if (!record || record.vertical !== 'PreInspection') return;
    const step4 = record.step4 || {};
    const cid = String(tmv3_clean_(step4.customer && step4.customer['Customer ID']));
    const id = String(tmv3_clean_(record.existingTaskId));
    if (step4.disposition !== 'VERIFIED' || !cid || !id || (refs.taskById || {})[id]) return;
    if (!tmv3_piPositiveId_(id)) throw new Error('PI_LINKED_TASK_ID_INVALID_NO_CREATE');
    const fresh = tmv3_step7FreshTask_(id, runtime);
    if (!fresh.ok || !fresh.task || String(fresh.task['Task ID']) !== id ||
        Number(fresh.task['Task Type ID']) !== 105 || String(fresh.task['Customer ID']) !== cid) {
      throw new Error('PI_LINKED_TASK_FRESH_OWNERSHIP_UNVERIFIED_NO_CREATE:' + id);
    }
    const withoutId = function(list) { return list.filter(function(t) { return String(t['Task ID']) !== id; }); };
    refs.tasksByCustomer = refs.tasksByCustomer || {};
    refs.tasksByOrder = refs.tasksByOrder || {};
    refs.taskById[id] = fresh.task;
    refs.tasks = withoutId(refs.tasks || []);
    refs.tasks.push(fresh.task);
    Object.keys(refs.tasksByCustomer).forEach(function(k) { refs.tasksByCustomer[k] = withoutId(refs.tasksByCustomer[k]); });
    if (!refs.tasksByCustomer[cid]) refs.tasksByCustomer[cid] = [];
    refs.tasksByCustomer[cid].push(fresh.task);
    Object.keys(refs.tasksByOrder).forEach(function(k) { refs.tasksByOrder[k] = withoutId(refs.tasksByOrder[k]); });
    const orderId = String(tmv3_clean_(fresh.task['Order ID']));
    if (orderId) {
      if (!refs.tasksByOrder[orderId]) refs.tasksByOrder[orderId] = [];
      refs.tasksByOrder[orderId].push(fresh.task);
    }
    tmv3_upsertTaskCacheRow_(fresh.task);
  });
}

/** Cache row for a Type 105 Task carries the verified canonical schedule. */
function tmv3_preInspectionCanonicalCacheTask_(task, canonical) {
  if (Number(task && task['Task Type ID']) !== 105) return task;
  const startMs = canonical && new Date(canonical.startDateTime).getTime();
  const dueMs = canonical && new Date(canonical.dueDateTime).getTime();
  if (!canonical || Number(canonical.taskId) !== Number(task['Task ID']) ||
      !canonical.startDateTime || !canonical.dueDateTime ||
      !Number.isFinite(startMs) || !Number.isFinite(dueMs) || dueMs < startMs)
    throw new Error('CANONICAL_PI_CACHE_SCHEDULE_UNVERIFIED_NO_CACHE_WRITE');
  const verified = Object.assign({}, task, {'Start':canonical.startDateTime,'Due':canonical.dueDateTime});
  verified.Fingerprint = tmv3_hash_(['Task ID','Task Number','Task Type ID','Task Type','Status','Name',
    'Customer ID','Location ID','Contact ID','Order ID','Start','Due','Assignees','Pools']
    .map(function(key) { return verified[key] || ''; }).join('|'));
  return verified;
}

/** Complete, ownership-verified fresh Type 105 history for the plan's Customer (≤10 Tasks). */
function tmv3_piFreshCustomerHistory_(plan) {
  const customer = tmv3_step4IdentityIndex_().customerById[String(plan.expectedCustomerId)];
  if (!customer) throw new Error('CREATE_CUSTOMER_NOT_IN_VERIFIED_CACHE');
  const ids = tmv3_searchPreInspectionTaskIdsForCustomer_(customer);
  if (ids.length > 10) throw new Error('CREATE_HISTORY_API_LIMIT_REVIEW_NO_CREATE');
  return ids.map(function(id) {
    const task = tmv3_getTaskById_(Number(id));
    if (String(task['Task ID']) !== String(id) || String(task['Customer ID']) !== String(plan.expectedCustomerId) ||
        Number(task['Task Type ID']) !== 105) throw new Error('CREATE_HISTORY_OWNERSHIP_FAILED_NO_CREATE');
    tmv3_upsertTaskCacheRow_(task);
    return task;
  });
}

/** Before CREATE: re-read complete fresh Customer history and re-plan on it. */
function tmv3_preInspectionFreshHistoryBeforeCreate_(record, plan) {
  tmv3_piFreshCustomerHistory_(plan);
  return tmv3_step7FreshPlansForEvent_('PreInspection', record.eventId);
}

/** Operator rule 2026-11-11: a completed (or undated completed) Task on the appointment day blocks a new Task. */
function tmv3_piAssertNoSameDayCompletedTask_(record, history) {
  if (!Array.isArray(history)) throw new Error('PI_RECREATE_FRESH_HISTORY_MISSING_NO_CREATE');
  (history || []).forEach(function(task) {
    if (!tmv3_taskIsCompleted_(task.Status)) return;
    if (!task.Start) throw new Error('PI_COMPLETED_TASK_UNDATED_REVIEW_NO_RECREATE:' + task['Task ID']);
    if (tmv3_sameLocalDay_(record.start, task.Start)) throw new Error('PI_SAME_DAY_COMPLETED_TASK_NO_RECREATE:' + task['Task ID']);
  });
  return true;
}

/** A completed Task proves THIS visit only with same-day timing at the same site. */
function tmv3_preInspectionHistoricalCoverageVerified_(plan) {
  if (!plan || plan.vertical !== 'PreInspection' ||
      plan.disposition !== 'FULFILLED_NO_RECREATE' ||
      plan.plan !== 'VERIFY_ONLY_HISTORY_NO_WRITE' ||
      plan.readStatus !== 'FRESH_TASK_GET' ||
      !tmv3_piPositiveId_(plan.taskId || '') ||
      !Array.isArray(plan.actions) || plan.actions.length !== 0 ||
      !tmv3_step6IsFulfilledStatus_(plan.taskStatus) ||
      plan.customerCheck !== 'MATCH' ||
      String(plan.expectedCustomerId) !== String(plan.actualCustomerId)) return false;
  const exactLink = String(plan.existingTaskId || plan.calendarTaskId || '') === String(plan.taskId);
  const sameSite = plan.locationCheck === 'MATCH' && plan.calendarTitleCheck !== 'CONFLICT';
  const day = tmv3_step6TaskTouchesEventDay_({start:plan.calendarStart}, {Start:plan.taskStart,Due:plan.taskDue});
  return !!day && sameSite && (exactLink || plan.customerCheck === 'MATCH');
}

const TMV3_PI_COMPLETION_CHECKS = Object.freeze(['customerCheck','locationCheck','requestedByCheck',
  'startCheck','endCheck','assignmentCheck','calendarTitleCheck','taskNameCheck']);

/** Converged only on one fresh NO_CHANGE plan whose every managed check is MATCH. */
function tmv3_preInspectionCompletionVerified_(plans, execution) {
  if (!Array.isArray(plans) || plans.length !== 1) return false;
  const p = plans[0];
  if (tmv3_preInspectionHistoricalCoverageVerified_(p)) return true;
  if (!p || p.blocker || !Array.isArray(p.actions) || p.actions.length !== 0 ||
      p.plan !== 'NO_CHANGE' || p.readStatus !== 'FRESH_TASK_GET' ||
      !tmv3_piPositiveId_(p.taskId || '')) return false;
  return TMV3_PI_COMPLETION_CHECKS.every(function(key) { return p[key] === 'MATCH'; });
}

// =====================================================================
// §7 PRESENTATION — Calendar title, authored description, a single
// clickable "-----Striven Links-----" block, and read-only copy parity.
// Guest changes and title rewrites are never made by unattended automation.
// =====================================================================

// Original: 35_CalendarLinks; Function: tmv3_preInspectionClickableLinkPresent_
function tmv3_preInspectionClickableLinkPresent_(description, link) {
  const text = String(description || '');
  const url = tmv3_regexEscape_(String(link && link.url || ''));
  if (!url) return false;
  return new RegExp('<a\\b[^>]*href=["\\\']' + url + '["\\\'][^>]*>', 'i').test(text);
}

function tmv3_preInspectionHtmlLink_(link) {
  return '<a href="' + tmv3_calendarHtmlEscape_(String(link.url || '')) + '">' +
    tmv3_calendarHtmlEscape_(String(link.label || '')) + '</a>';
}

/** Previous title becomes the first description line when it is not already there. */
function tmv3_preInspectionEnsureLegacyTitle_(description, legacyTitle) {
  const authored = String(description || '');
  const title = String(legacyTitle || '').trim();
  if (!title) return authored;
  const readable = tmv3_calendarReadableDescription_(authored);
  if (String(readable || '').trim().indexOf(title) === 0) return authored;
  const html = /<[a-z][^>]*>/i.test(authored);
  return html ? tmv3_calendarHtmlEscape_(title) + '<br>*<br>' + authored :
    title + (authored ? '\n*\n' + authored : '');
}

/** Strip ONLY automation-managed blocks; keep headings, blank lines, notes and markup. */
function tmv3_preInspectionAuthoredMarkup_(description) {
  let raw = tmv3_stripV3ManagedBlock_(String(description || ''));
  raw = raw.replace(/<!--\s*PREINSPECT_STRIVEN_TASK_LINK_START\s*-->[\s\S]*?<!--\s*PREINSPECT_STRIVEN_TASK_LINK_END\s*-->/gi, '');
  const offset = tmv3_managedCalendarBlockStart_(raw);
  if (offset >= 0) {
    raw = raw.slice(0, offset);
    raw = raw.replace(/(?:<br\s*\/?>|\r?\n)\s*$/i, ''); // only the separator before our footer
  }
  if (/<[a-z][^>]*>/i.test(raw)) return raw.replace(/\s+$/, '');
  // Hyperlinks need markup: escape plain text and map each newline to one <br>.
  return tmv3_calendarHtmlEscape_(raw.replace(/\r\n?/g, '\n').replace(/\s+$/, '')).replace(/\n/g, '<br>');
}

function tmv3_preInspectionClickableManagedDescription_(existingDescription, plan) {
  const before = String(existingDescription || '');
  const links = (plan && plan.links) || [];
  const seen = {};
  links.forEach(function(link) {
    if (!link || !link.key || !/^https:\/\//i.test(String(link.url || '')) || seen[link.key])
      throw new Error('PI_MANAGED_LINK_INVALID_OR_DUPLICATE');
    seen[link.key] = true;
  });
  const authored = tmv3_preInspectionAuthoredMarkup_(before);
  const header = tmv3_calendarHtmlEscape_(TMV3_FINAL_LINK_HEADING);
  const footer = header + (links.length ? '<br>' + links.map(tmv3_preInspectionHtmlLink_).join('<br>') : '');
  const out = authored + (authored ? '<br>' : '') + footer;
  tmv3_assertPreInspectionLinkDescription_(out, plan);
  return {description:out,status:out === before ? 'LINKS_ALREADY_CANONICAL' : 'LINKS_REQUIRE_UPDATE',changed:out !== before};
}

function tmv3_preInspectionTitleVariants_(plan) {
  plan = plan || {};
  const number = tmv3_clean_(plan.customerNumber);
  const name = tmv3_clean_(plan.customerName);
  const phone = tmv3_clean_(plan.phone);
  const variants = [tmv3_clean_(plan.currentTitle)];
  if (number && name && phone) {
    variants.push(number + ' - ' + name + ' - ' + phone);
    variants.push('C#' + number + ' - ' + name + ' - ' + phone);
    variants.push('Cust#' + number + ' - ' + name + ' - ' + phone);
  }
  return tmv3_unique_(variants.map(tmv3_clean_).filter(Boolean));
}

function tmv3_stripPreInspectionPreservedTitle_(value, plan) {
  let text = tmv3_decodeCalendarHtmlEntities_(String(value || ''));
  tmv3_preInspectionTitleVariants_(plan).forEach(function(title) {
    text = text.replace(new RegExp('^\\s*(?:<p[^>]*>\\s*)?' + tmv3_regexEscape_(title) +
      '\\s*(?:<\\/p>)?\\s*(?:(?:<br\\s*\\/?>|\\r?\\n)\\s*){0,2}', 'i'), '');
  });
  return text.replace(/^\s+/, '');
}

/** Drop the presentation-only "Sales Order:" / "Notes:" labels at paragraph starts. */
function tmv3_piStripPresentationLabels_(text) {
  return text
    .replace(/(^|\n\n)\s*(?:\*\*)?Sales Order:(?:\*\*)?\s*/ig, function(match, prefix) { return prefix || ''; })
    .replace(/(^|\n\n)\s*(?:\*\*)?Notes:(?:\*\*)?\s*/ig, function(match, prefix) { return prefix || ''; });
}

function tmv3_normalizePreInspectionAuthoredWraps_(value) {
  let text = tmv3_decodeCalendarHtmlEntities_(String(value || ''))
    .replace(/<\s*br\s*\/?>/gi, '\n')
    .replace(/<\s*\/p\s*>/gi, '\n\n')
    .replace(/<\s*p[^>]*>/gi, '')
    .replace(/<\s*b\s*>/gi, '**')
    .replace(/<\s*\/b\s*>/gi, '**')
    .replace(/<[^>]+>/g, '')
    .replace(/\r\n?/g, '\n')
    .replace(/ {2,}\n/g, '\n')
    .trim();
  const marker = '__TMV3_PARAGRAPH_BREAK__';
  text = text
    .replace(/\n\s*\n+/g, marker)
    .replace(/[ \t]*\n[ \t]*/g, ' ')
    .replace(new RegExp(marker, 'g'), '\n\n');
  return tmv3_piStripPresentationLabels_(text).replace(/\n{3,}/g, '\n\n').trim();
}

/** Plain-text notes for display. The exact previous title stays as the first line. */
function tmv3_preInspectionCalendarNotesText_(description, plan) {
  const authored = tmv3_normalizePreInspectionAuthoredWraps_(tmv3_stripManagedLinkBlocks_(description || ''));
  if (!authored) return '';
  return tmv3_piStripPresentationLabels_(authored).replace(/\n{3,}/g, '\n\n').trim();
}

/** Exactly one link heading, one anchor per managed link with the exact label, no raw copies. */
function tmv3_assertPreInspectionLinkDescription_(description, plan) {
  const text = String(description || '');
  const headings = text.match(/Striven\s+Links/gi) || [];
  if (headings.length !== 1 || /Pre[- ]Inspection\s+Task\s+Link/i.test(text) ||
      /PREINSPECT_STRIVEN_TASK_LINK_(?:START|END)/i.test(text)) {
    throw new Error('PREINSPECTION_LINK_SECTION_NOT_UNIQUE');
  }
  (plan.links || []).forEach(function(link) {
    const url = tmv3_regexEscape_(tmv3_calendarHtmlEscape_(String(link.url || '')));
    if (!url) throw new Error('PREINSPECTION_LINK_URL_EMPTY');
    const anchors = text.match(new RegExp('<a\\b[^>]*href=["\\\']' + url + '["\\\'][^>]*>[\\s\\S]*?<\\/a>', 'gi')) || [];
    if (anchors.length !== 1) throw new Error('PREINSPECTION_LINK_NOT_UNIQUE: ' + link.key);
    const label = anchors[0].replace(/^[\s\S]*?>/, '').replace(/<\/a>$/i, '');
    if (label !== tmv3_calendarHtmlEscape_(String(link.label || ''))) {
      throw new Error('PREINSPECTION_LINK_LABEL_MISMATCH: ' + link.key);
    }
    if (text.replace(anchors[0], '').indexOf(String(link.url || '')) >= 0) {
      throw new Error('PREINSPECTION_DUPLICATE_RAW_LINK: ' + link.key);
    }
  });
  return true;
}

function tmv3_preInspectionCalendarBusinessSnapshot_(event) {
  if (!event) throw new Error('PreInspection Calendar snapshot requires an event.');
  const participants = {};
  try {
    (event.getCreators ? event.getCreators() : []).forEach(function(email) {
      const clean = tmv3_normEmail_(email);
      if (clean) participants[clean] = true;
    });
  } catch (ignored) {}
  try {
    (event.getGuestList ? event.getGuestList(true) : []).forEach(function(guest) {
      const clean = tmv3_normEmail_(guest && guest.getEmail ? guest.getEmail() : '');
      if (clean) participants[clean] = true;
    });
  } catch (ignored) {}
  const stephen = tmv3_normEmail_((TMV3.VERTICALS.PreInspection || {}).secondaryOwnerEmail || 'stephen@classicfireplace.ca');
  return {
    title:String(event.getTitle ? event.getTitle() || '' : ''),
    description:String(event.getDescription ? event.getDescription() || '' : ''),
    location:String(event.getLocation ? event.getLocation() || '' : ''),
    start:event.getStartTime ? event.getStartTime().getTime() : null,
    end:event.getEndTime ? event.getEndTime().getTime() : null,
    allDay:event.isAllDayEvent ? event.isAllDayEvent() === true : false,
    requiredStephenPresent:!!participants[stephen]
  };
}

function tmv3_preInspectionCalendarParityEqual_(expected, actual) {
  expected = expected || {};
  actual = actual || {};
  return ['title','description','location','start','end','allDay','requiredStephenPresent']
    .every(function(key) { return expected[key] === actual[key]; });
}

/** Audit required copies against CF Preinspects; never synchronizes or adds attendees. */
function tmv3_reconcilePreInspectionCalendarParity_(record, scope) {
  if (!record || record.vertical !== 'PreInspection') return {status:'NOT_NEEDED',copies:[],writes:0};
  const cfg = TMV3.VERTICALS.PreInspection || {};
  const primaryId = String(cfg.primaryCalendarId || '');
  if (!primaryId) throw new Error('PI_PRIMARY_CALENDAR_NOT_CONFIGURED');
  const master = tmv3_findEventCopyRobust_(primaryId, record.eventId);
  if (!master || !master.event) throw new Error('PI_AUTHORITY_EVENT_MISSING');
  const authority = tmv3_preInspectionCalendarBusinessSnapshot_(master.event);
  const ids = tmv3_unique_(tmv3_requiredCalendarCopiesForEvent_(record) || []);
  const differences = [];
  ids.forEach(function(calendarId) {
    if (String(calendarId) === primaryId) return;
    const found = tmv3_findEventCopyRobust_(calendarId, record.eventId);
    if (!found || !found.event) { differences.push({calendarId:calendarId,issue:'COPY_NOT_ACCESSIBLE_OR_MISSING'}); return; }
    const state = tmv3_preInspectionCalendarBusinessSnapshot_(found.event);
    const changed = ['title','description','location','start','end','allDay'].filter(function(key) {
      return authority[key] !== state[key];
    });
    if (changed.length) differences.push({calendarId:calendarId,issue:'PARITY_MISMATCH',fields:changed});
  });
  return {status:differences.length ? 'PI_PARITY_REVIEW_REQUIRED' : 'PI_PARITY_READ_ONLY_MATCH',
    authoritativeCalendarId:primaryId,eventId:record.eventId,copies:ids.length,
    differences:differences,writes:0,guestInvites:0};
}

/** Guest invitations notify people; unattended automation never changes guests. */
function tmv3_ensurePreInspectionMirror_(record, scope) {
  return {status:'PI_GUEST_CHANGE_OPERATOR_APPROVAL_REQUIRED',
    eventId:record && record.eventId || '',writePerformed:false,calendarWrites:0,guestInvites:0};
}

function tmv3_preInspectionGuestSyncDecision_(guestEmails, stephenEmail, ownerCalendarId) {
  const owner = String(ownerCalendarId || '').trim();
  if (!owner) return {status:'OWNER_NOT_VERIFIED',action:'NONE'};
  return {status:'REVIEW_ONLY_NO_UNATTENDED_GUEST_CHANGES',action:'NONE'};
}

function tmv3_reconcilePreInspectionSharedGuest_AUTO_() {
  return {status:'RETIRED_UNCONDITIONAL_GUEST_SYNC',writes:0,emailsSent:0,guestInvites:0};
}

/** Automatic title-history reconstruction wrote every copy without stable proof; read-only now. */
function tmv3_restorePreInspectionLegacyTitle_(record) {
  return {status:'PI_TITLE_HISTORY_READ_ONLY_REQUIRE_APPROVAL',eventId:record && record.eventId || '',writes:0};
}

function tmv3_piAssertCalendarTitle_(value) {
  const s = String(value || '').trim();
  if (!/^[0-9]+ - .+ - \([0-9]{3}\) [0-9]{3}-[0-9]{4}$/.test(s)) {
    throw new Error('PI_CALENDAR_TITLE_OUTSIDE_LOCKED_FORMAT_NO_WRITE');
  }
  return true;
}

function tmv3_piAssertTaskTitle_(value) {
  const s = String(value || '').trim();
  if (!/^.+ - .+, .+ - \([0-9]{3}\) [0-9]{3}-[0-9]{4}$/.test(s)) {
    throw new Error('PI_TASK_TITLE_OUTSIDE_LOCKED_FORMAT_NO_WRITE');
  }
  return true;
}

// =====================================================================
// §8 WRITE GUARDS — enforced at the real outgoing CREATE and PATCH
// boundaries in 10_Shared_Core. Inspector-owned fields 852/853/854/860 are
// never written; Sales Orders are never attached at CREATE; Description is
// blank at CREATE; Requested By is the verified owning-Calendar Employee.
// =====================================================================

const TMV3_PI_CREATE_FORBIDDEN_KEYS =
  /^(?:SalesOrder|SalesOrderId|SalesOrderID|SOId|OrderId|OrderID|InfoCustomFields|CustomFields|CustomFieldValues|TaskFields|InfoFields|Fields)$/i;

function tmv3_enforcePreInspectionCreatePayloadHardRules_(payload) {
  payload = payload || {};
  Object.keys(payload).forEach(function(k) {
    if (TMV3_PI_CREATE_FORBIDDEN_KEYS.test(k)) delete payload[k];
  });
  payload.Description = '';
  return payload;
}

function tmv3_assertPreInspectionCreatePayloadHardRules_(payload) {
  const p = payload || {}, type = p.Type || {};
  if (Number(type.Id || type.id) !== 105) throw new Error('PI_CREATE_TYPE_MUST_BE_105');
  if (Object.keys(p).some(function(k) { return TMV3_PI_CREATE_FORBIDDEN_KEYS.test(k); }))
    throw new Error('PI_CREATE_FORBIDDEN_ORDER_OR_CUSTOM_FIELDS');
  if (p.Description !== '') throw new Error('PI_CREATE_DESCRIPTION_NOT_BLANK');
  // Requested By must be a real Striven Employee, not a Customer Contact.
  tmv3_piAssertRequestedByEmployeeShape_(p.RequestedBy);
  // Branch guard: no protected field or order attachment at ANY nesting depth.
  tmv3_piAssertNoNestedProtectedFields_(p, 'CREATE');
  return true;
}

function tmv3_piAssertSafeTaskPatch_(payload) {
  tmv3_piAssertLockedStandard_();
  const p = payload || {};
  const protectedNames = /^(?:InfoCustomFields|CustomFields|Description|SalesOrder|SalesOrderID|SalesOrderId|SOId|Order|OrderId|Type)$/i;
  Object.keys(p).forEach(function(key) {
    if (protectedNames.test(key)) throw new Error('PI_PROTECTED_TASK_FIELD_NO_WRITE:' + key);
    if (key.toLowerCase() === 'requestedby' && key !== 'RequestedBy')
      throw new Error('PI_REQUESTED_BY_FIELD_CASE_INVALID_NO_WRITE');
  });
  if (p.RequestedBy !== undefined) tmv3_piAssertRequestedByEmployeeShape_(p.RequestedBy);
  if (p.Title !== undefined) tmv3_piAssertTaskTitle_(p.Title);
  tmv3_piAssertNoNestedProtectedFields_(p, 'PATCH');
  return true;
}

/** Throws when the deep payload scan (§13) rejects the outgoing body. */
function tmv3_piAssertNoNestedProtectedFields_(payload, mode) {
  const verdict = tmv3_piRebuildValidatePayload_(payload, mode);
  if (verdict.decision !== 'ACCEPT') throw new Error('PI_PAYLOAD_' + verdict.decision + '_NO_WRITE');
  return true;
}

/** Shape only: a positive integer Employee ID; provenance is checked at Step 7. */
function tmv3_piAssertRequestedByEmployeeShape_(requestedBy) {
  tmv3_piRequestedByContract_();
  const type = String(requestedBy && (requestedBy.Type || requestedBy.type) || '').trim().toLowerCase();
  const rawId = requestedBy && (requestedBy.Id !== undefined ? requestedBy.Id : requestedBy.id);
  const id = Number(rawId);
  if (type !== 'employee' || !tmv3_piPositiveId_(rawId || '') || !Number.isSafeInteger(id) || id <= 0)
    throw new Error('PI_REQUESTED_BY_VERIFIED_EMPLOYEE_REQUIRED_NO_WRITE');
  return true;
}

/** Plan requester must equal the employee mapped from this event's verified owning Calendar. */
function tmv3_piAssertRequestedByPlan_(plan, record) {
  tmv3_piRequestedByContract_();
  const owner = tmv3_piVerifiedOwnerEmployee_(record);
  if (owner.status !== 'MATCHED' || !Number.isSafeInteger(owner.employeeId))
    throw new Error(owner.blocker || 'PI_REQUESTED_BY_OWNER_UNVERIFIED_NO_WRITE');
  const type = String(plan && plan.expectedRequestedByType || '').trim().toLowerCase();
  const rawId = plan && plan.expectedRequestedById;
  const id = Number(rawId);
  if (type !== 'employee' || !tmv3_piPositiveId_(rawId || '') || !Number.isSafeInteger(id) || id <= 0)
    throw new Error('PI_EXPECTED_REQUESTED_BY_EMPLOYEE_NOT_PROVEN_NO_WRITE');
  if (id !== owner.employeeId) throw new Error('PI_REQUESTED_BY_MUST_MATCH_VERIFIED_EVENT_OWNER_NO_WRITE');
  return true;
}

/** Striven Sales Order status IDs paired with their names (branch guard: both must agree). */
const TMV3_PI_ORDER_STATUS_BY_ID = Object.freeze({
  '18':'incomplete','19':'quoted','20':'pending approval','22':'approved','25':'in progress'
});

/** Before CREATE: an explicit SO is optional, but an Approved/In Progress order is not a PreInspection. */
function tmv3_preInspectionOrderBeforeCreate_(record, plan) {
  if (!record || record.vertical !== 'PreInspection') throw new Error('PI_RECORD_REQUIRED');
  const number = String(tmv3_clean_(record.orderNumber));
  if (!number) return {status:'OPTIONAL_ORDER_MISSING_NOT_AN_ERROR'};
  if (!/^\d{4,8}$/.test(number)) throw new Error('PI_SO_NUMBER_INVALID_REVIEW');
  const cid = Number(plan.expectedCustomerId);
  if (!Number.isSafeInteger(cid) || cid <= 0) throw new Error('PI_CUSTOMER_NOT_VERIFIED_NO_CREATE');
  const response = tmv3_fetchJson_(TMV3.API_BASE + '/v1/sales-orders/search', {
    method:'post',contentType:'application/json',
    payload:JSON.stringify({PageIndex:0,PageSize:2,Number:number,CustomerId:cid})});
  const rows = response && (response.Data || response.data);
  const count = Number(response && (response.TotalCount !== undefined ? response.TotalCount : response.totalCount));
  if (!Array.isArray(rows) || count !== 1 || rows.length !== 1 ||
      String(tmv3_first_(rows[0], ['Number','OrderNumber'])) !== number)
    throw new Error('PI_SO_NOT_UNIQUE_OR_SOURCE_INCOMPLETE_NO_CREATE');
  const orderId = Number(tmv3_first_(rows[0], ['Id','id']));
  if (!Number.isSafeInteger(orderId) || orderId <= 0) throw new Error('PI_SO_ID_UNVERIFIED');
  const order = tmv3_fetchJson_(TMV3.API_BASE + '/v1/sales-orders/' + orderId, {method:'get'});
  const customer = tmv3_first_(order, ['Customer','customer']);
  const status = tmv3_first_(order, ['Status','status']);
  if (Number(tmv3_first_(order, ['Id','id'])) !== orderId ||
      String(tmv3_first_(order, ['OrderNumber','Number'])) !== number ||
      Number(tmv3_first_(customer, ['Id','id'])) !== cid)
    throw new Error('PI_SO_FRESH_IDENTITY_CONFLICT_NO_CREATE');
  const name = tmv3_norm_(tmv3_first_(status, ['Name','name']));
  if (name === 'approved') throw new Error('APPROVED_ORDER_INSTALLATION_REVIEW_NO_PI_CREATE');
  if (name === 'in progress') throw new Error('IN_PROGRESS_ORDER_VISIT_TYPE_REVIEW_NO_PI_CREATE');
  if (!['quoted','incomplete','pending approval'].includes(name)) throw new Error('PI_SO_STATUS_UNVERIFIED_NO_CREATE');
  // Branch guard: a renamed or re-numbered status must not pass on its name alone.
  const statusId = String(tmv3_first_(status, ['Id','id']));
  if (TMV3_PI_ORDER_STATUS_BY_ID[statusId] !== name) throw new Error('PI_SO_STATUS_ID_NAME_MISMATCH_NO_CREATE');
  return {status:'SO_CORROBORATION_ONLY_NOT_ATTACHED',orderId:orderId,orderNumber:number,orderStatus:name};
}

// =====================================================================
// §9 NOTIFY — the approved missing-details email. The renderer
// tmv3_preInspectionMissingDetailEmailPreview_ is FROZEN: its source text is
// fingerprinted against TMV3_HARD_RULES.PreInspection.notificationRendererFingerprint.
// Do not edit, reformat or re-indent it.
// =====================================================================

// Original: 58_Operations; Function: tmv3_preInspectionHasSalesOrderReference_
function tmv3_preInspectionHasSalesOrderReference_(eventRecord) {
  const record = eventRecord || {};
  if (tmv3_clean_(record.orderNumber)) return true;
  const haystack = [
    record.title || '',
    tmv3_calendarReadableDescription_(
      tmv3_identityEvidenceDescriptionRaw_(record.rawDescription || record.descriptionClean || '', '')
    )
  ].join(' ');
  // Any labelled SO / Sales Order / Order reference counts, including compact
  // slash notation such as SO#5853/67/80/85.
  const clean = tmv3_stripPhonesForOrderParsing_(haystack)
    .replace(/https?:\/\/\S+/gi, ' ')
    .replace(/(?:Customer|Cust|Task)\s*(?:#|Number|No\.?)?\s*[:\-]?\s*\d+/gi, ' ')
    .replace(/^\s*#?\d{4,8}\s*[-–—]/, ' ');
  return /\b\d{6}\b/.test(clean) ||
    /\b(?:SO|S\/O|Sales\s*Order|Order)\s*(?:#|No\.?|Number)?\s*[:\-]?\s*\d{4,8}(?:\s*\/\s*\d{2,8})*/i.test(String(haystack || ''));
}

/** The single company creator; missing or ambiguous creator means no send. */
function tmv3_preInspectionCreatorEmail_(eventRecord) {
  const sources = []
    .concat(eventRecord && eventRecord.sourceCreators || [])
    .concat(String(eventRecord && eventRecord.creator || '').split(','))
    .map(function(value) { return tmv3_clean_(value).toLowerCase(); })
    .filter(Boolean);
  const unique = tmv3_unique_(sources.filter(function(email) { return /@classicfireplace\.ca$/i.test(email); }));
  return unique.length === 1 ? unique[0] : '';
}

function tmv3_preInspectionCreatorFirstName_(email) {
  const local = tmv3_clean_(email).split('@')[0] || '';
  if (!local) return 'there';
  const token = local.split(/[._\-]+/).filter(Boolean)[0] || '';
  return token ? token.charAt(0).toUpperCase() + token.slice(1).toLowerCase() : 'there';
}

/** Customer # recognised in the title only when a unique profile corroborates it. */
function tmv3_preInspectionNotificationCustomerNumber_(record, customers) {
  const title = tmv3_clean_(record && record.title);
  const stripped = tmv3_stripPhonesForOrderParsing_(title)
    .replace(/https?:\/\/\S+/gi, ' ')
    .replace(/\b(?:SO|S\/O|Sales\s*Order|Order)\s*(?:#|No\.?|Number)?\s*[:\-]?\s*\d{4,8}(?:\s*\/\s*\d{2,8})*/ig, ' ');
  const numbers = stripped.match(/\b\d{4,8}\b/g) || [];
  const phone = tmv3_phone10_(record && record.phone || tmv3_extractPhone_(title));
  const words = title.toLowerCase().replace(/[^a-z0-9]+/g, ' ').split(' ').filter(Boolean);
  const matches = (customers || []).filter(function(customer) {
    const number = tmv3_clean_(customer['Customer Number']);
    if (numbers.indexOf(number) < 0) return false;
    const nameWords = tmv3_clean_(customer.Name).toLowerCase().replace(/[^a-z0-9]+/g, ' ').split(' ').filter(Boolean);
    const nameMatches = nameWords.length >= 2 && nameWords.every(function(word) { return words.indexOf(word) >= 0; });
    const profilePhone = tmv3_phone10_(customer['Primary Phone']);
    return (phone && profilePhone && phone === profilePhone) || nameMatches;
  });
  if (matches.length !== 1) return '';
  const number = tmv3_clean_(matches[0]['Customer Number']);
  if ((customers || []).filter(function(c) { return tmv3_clean_(c['Customer Number']) === number; }).length !== 1) return '';
  return number;
}

/** Human-supplied details only. The Sales Order is optional and never emailed about. */
function tmv3_preInspectionMissingDetailAssessment_(eventRecord) {
  const record = eventRecord || {};
  const text = [record.title || '', record.identityDescriptionClean ||
    record.descriptionClean || record.rawDescription || ''].join(' ');
  const customerNumber = tmv3_extractCustomerNumberForVertical_('PreInspection', text) ||
    tmv3_preInspectionNotificationCustomerNumber_(record, tmv3_rows_(TMV3.SHEETS.CUSTOMERS));
  const missing = [];
  if (!customerNumber) missing.push({code:'CUSTOMER_NUMBER',field:'Title',message:'Customer Number missing'});
  if (!tmv3_clean_(record.location)) missing.push({code:'LOCATION',field:'Location',
    message:'Location missing — please add the full job-site address'});
  return {eventId:record.eventId || '',title:tmv3_clean_(record.title),
    creatorEmail:tmv3_preInspectionCreatorEmail_(record),missing:missing,shouldNotify:missing.length > 0};
}

function tmv3_preInspectionHtmlEscape_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function tmv3_preInspectionGuestEmails_(eventRecord) {
  return tmv3_unique_(String(eventRecord && eventRecord.guests || '').split(',')
    .map(function(value) { return tmv3_norm_(value); }).filter(Boolean));
}

function tmv3_preInspectionDisplayGuests_(eventRecord) {
  const guests = tmv3_preInspectionGuestEmails_(eventRecord);
  if (!guests.length) return 'No guests added';
  return guests.map(function(email) {
    if (email === tmv3_norm_(TMV3.VERTICALS.PreInspection.secondaryOwnerEmail)) return 'Stephen Foley';
    if (email === tmv3_norm_(TMV3.VERTICALS.PreInspection.primaryCalendarId)) return 'CF Preinspects';
    return email;
  }).join('<br>');
}

function tmv3_preInspectionMissingFields_(assessment) {
  const fields = {};
  (assessment.missing || []).forEach(function(item) { fields[item.field] = true; });
  return fields;
}

// Template fragment used by the frozen renderer. Keep markup byte-identical.
function tmv3_preInspectionComparisonRowsHtml_(eventRecord, assessment) {
  const escape = tmv3_preInspectionHtmlEscape_;
  return (assessment.missing || []).map(function(item) {
    const current = item.code === 'CUSTOMER_NUMBER'
      ? escape(eventRecord.title || '') + '<br><strong style="color:#B42318;">Customer Number missing</strong>'
      : '<strong style="color:#B42318;">' + escape(item.message) + '</strong>';
    const standard = item.code === 'CUSTOMER_NUMBER'
      ? 'Customer Number - Customer Name - Phone Number'
      : item.code === 'SALES_ORDER' ? 'Sales Order #' : 'Full Job-Site Address';
    return '<tr><td valign="top" style="padding:12px;border:1px solid #dddddd;font-weight:700;">' +
      escape(item.field) + '</td><td valign="top" style="padding:12px;border:1px solid #dddddd;">' +
      current + '</td><td valign="top" style="padding:12px;border:1px solid #dddddd;color:#256b35;font-weight:700;">' +
      standard + '</td></tr>';
  }).join('');
}

function tmv3_preInspectionNotificationCc_() {
  return ['jay@classicfireplace.ca','thang@classicfireplace.ca'];
}

// Original: 58_Operations; Function: tmv3_preInspectionMissingDetailEmailPreview_
function tmv3_preInspectionMissingDetailEmailPreview_(eventRecord) {
  const assessment = tmv3_preInspectionMissingDetailAssessment_(eventRecord);
  if (!assessment.shouldNotify) return null;

  const creatorEmail = assessment.creatorEmail;
  const firstName = tmv3_preInspectionCreatorFirstName_(creatorEmail);
  const missingItemsHtml = assessment.missing.map(function(item) {
    return '<div style="color:#B42318;font-weight:700;">❌ ' + tmv3_preInspectionHtmlEscape_(item.message) + '</div>';
  }).join('');

  const when =
    Utilities.formatDate(eventRecord.start, TMV3_TIMEZONE, 'MMM d, yyyy') +
    ' · ' +
    Utilities.formatDate(eventRecord.start, TMV3_TIMEZONE, 'h:mm a') +
    ' – ' +
    Utilities.formatDate(eventRecord.end, TMV3_TIMEZONE, 'h:mm a');

  const followupNumber=Math.max(0,Number(eventRecord.notificationFollowupNumber || 0));
  const followupLabel=followupNumber ? 'FOLLOW-UP #'+followupNumber : '';
  const html =
    '<!doctype html><html lang="en"><body style="margin:0;padding:0;background:#f4f3f1;font-family:Arial,Helvetica,sans-serif;color:#222;">' +
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0"><tr><td align="center" style="padding:20px 12px;">' +
    '<table role="presentation" width="760" cellspacing="0" cellpadding="0" border="0" style="width:100%;max-width:760px;background:#ffffff;">' +
    '<tr><td style="background:#202020;border-top:4px solid #b42b2f;padding:18px 24px;">' +
      '<div style="font-size:18px;font-weight:700;color:#ffffff;">' + (followupLabel ? followupLabel + ' — ' : '') + 'Pre-Inspection Update Required</div>' +
      '<div style="font-size:12px;color:#aaaaaa;margin-top:3px;">Classic Fireplace &amp; BBQ Store</div>' +
    '</td></tr>' +
    '<tr><td style="padding:20px 24px 12px;">' +
      '<div style="font-size:12px;color:#666;font-style:italic;margin-bottom:16px;">This is an automated email notification from Classic Fireplace &amp; BBQ Store.</div>' +
      '<div style="font-size:14px;line-height:21px;">Hi <strong>' + tmv3_preInspectionHtmlEscape_(firstName) + '</strong>,</div>' +
      '<div style="margin-top:12px;padding:12px;background:#f6f5f3;border:1px solid #e2e0dc;"><strong>Appointment requiring attention</strong><br>' + tmv3_preInspectionHtmlEscape_(eventRecord.title || '(Untitled event)') + '<br><strong>' + tmv3_preInspectionHtmlEscape_(when) + '</strong><br>' + tmv3_preInspectionHtmlEscape_(eventRecord.location || 'Location missing') + '</div>' +
      '<div style="font-size:14px;line-height:21px;margin-top:8px;">This appointment is missing the following details. Please update the items highlighted below.</div>' +
      '<div style="margin-top:12px;background:#fff5f5;border-left:4px solid #b42b2f;padding:11px 13px;font-size:13px;line-height:22px;color:#8f2024;">' +
        missingItemsHtml +
      '</div>' +
    '</td></tr>' +
    '<tr><td style="padding:0 24px 14px;"><table width="100%" cellspacing="0" cellpadding="0" border="0" style="background:#f6f5f3;border:1px solid #e2e0dc;"><tr><td style="padding:12px 14px;">' +
      '<div style="font-size:15px;font-weight:700;">' + tmv3_preInspectionHtmlEscape_(eventRecord.title || '(Untitled event)') + '</div>' +
      '<div style="font-size:12px;color:#666;margin-top:4px;">' + tmv3_preInspectionHtmlEscape_(when) + '</div>' +
      (eventRecord.location ? '<div style="font-size:12px;color:#666;margin-top:4px;">' + tmv3_preInspectionHtmlEscape_(eventRecord.location) + '</div>' : '') +
    '</td></tr></table></td></tr>' +
    '<tr><td style="padding:0 24px 20px;"><table width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;border:1px solid #dddddd;font-size:13px;">' +
      '<tr>' +
        '<td width="18%" style="background:#f3f3f3;padding:10px 12px;border-right:1px solid #dddddd;border-bottom:1px solid #dddddd;"></td>' +
        '<td width="41%" style="background:#fff1f1;color:#9b1c1f;font-weight:700;padding:10px 12px;border-right:1px solid #dddddd;border-bottom:1px solid #dddddd;">What it is now</td>' +
        '<td width="41%" style="background:#f1f8f1;color:#28632f;font-weight:700;padding:10px 12px;border-bottom:1px solid #dddddd;">What it should be</td>' +
      '</tr>' +
      tmv3_preInspectionComparisonRowsHtml_(eventRecord, assessment) +
    '</table></td></tr>' +
    '<tr><td style="padding:0 24px 22px;">' +
      '<div style="font-size:14px;margin-bottom:14px;"><a href="' +
      tmv3_preInspectionHtmlEscape_(tmv3_calendarEventUrl_(eventRecord.eventId, eventRecord.calendarId || eventRecord.organizerCalendarId) || eventRecord.eventUrl) +
      '">Open calendar appointment</a></div>' +
      '<div style="font-size:14px;margin-bottom:14px;">Please update the missing details in the calendar event.</div>' +
      '<div style="font-size:14px;">Thank you.</div></td></tr>' +
    '<tr><td style="background:#202020;padding:12px 24px;text-align:center;"><div style="font-size:12px;font-weight:700;color:#ffffff;">Classic Fireplace &amp; BBQ Store</div></td></tr>' +
    '</table></td></tr></table></body></html>';

  return {
    eventId: eventRecord.eventId || '',
    recipient: creatorEmail,
    cc: tmv3_preInspectionNotificationCc_(),
    subject: (followupLabel ? followupLabel + ' | ' : '') + 'MISSING DETAILS: ' + (eventRecord.calendarCustomerName || eventRecord.title || 'Appointment') + ' | ' + when + ' | ' + assessment.missing.map(function(item){return item.code === 'CUSTOMER_NUMBER' ? 'Customer Number' : item.field;}).join(', '),
    missing: assessment.missing,
    htmlBody: html,
    emailDeliveryEnabled: TMV3.MODE === 'PRODUCTION_WRITE' &&
      TMV3.OPERATIONS.preInspectionMissingDetailEmailsEnabled === true,
    approvalRequiredBeforeSend: false
  };
}

function tmv3_preInspectionMissingDetailPreviewFromRecords_(records) {
  const preInspection = (records || []).filter(function(record) {
    return record && record.vertical === 'PreInspection';
  });
  const previews = preInspection.map(tmv3_preInspectionMissingDetailEmailPreview_).filter(Boolean);
  return {status:'PREVIEW_ONLY_NO_EMAILS',checked:preInspection.length,wouldNotify:previews.length,previews:previews};
}

/** Menu: Preview Pre-Inspection Missing Detail Emails. Sends nothing. */
function tmv3_previewPreInspectionMissingDetailEmails() {
  const result = tmv3_preInspectionMissingDetailPreviewFromRecords_(tmv3_step1CalendarRecords_());
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function tmv3_preInspectionNotificationDateParts_(value) {
  const date = new Date(value);
  if (!isFinite(date.getTime())) return null;
  const text = Utilities.formatDate(date, 'America/Toronto', 'yyyy-MM-dd HH:mm:ss');
  const day = text.slice(0, 10);
  const utc = new Date(day + 'T12:00:00Z');
  return {text:text,day:day,weekday:utc.getUTCDay(),utc:utc};
}

/** Monday (yyyy-MM-dd) of the Toronto week containing parts. */
function tmv3_preInspectionNotificationWeek_(parts) {
  const day = new Date(parts.utc.getTime());
  day.setUTCDate(day.getUTCDate() - ((day.getUTCDay() + 6) % 7));
  return day.toISOString().slice(0, 10);
}

/** Schedule rule: same-week appointment → next slot; otherwise two business days
 * after creation. Follow-ups only on a later business day, ≥24h apart, before start.
 * Any non-SENT prior state is uncertain and is never blindly retried. */
function tmv3_preInspectionNotificationDecision_(record, state, nowValue) {
  const now = new Date(nowValue || Date.now());
  const start = new Date(record.start);
  if (!isFinite(start.getTime())) return {status:'EVENT_START_UNKNOWN',send:false};
  if (record.status === 'cancelled' || now.getTime() >= start.getTime()) return {status:'APPOINTMENT_ENDED',send:false};
  if (state && state.status !== 'SENT') return {status:'UNCERTAIN_SEND_NO_RETRY',send:false};
  const today = tmv3_preInspectionNotificationDateParts_(now);
  if (today.weekday === 0 || today.weekday === 6) return {status:'WEEKEND_DEFERRED',send:false};
  if (state) {
    const last = new Date(state.lastSentAt);
    if (!isFinite(last.getTime())) return {status:'SEND_TIME_UNKNOWN_NO_RETRY',send:false};
    if (now.getTime() - last.getTime() < 86400000 ||
        today.day === tmv3_preInspectionNotificationDateParts_(last).day) return {status:'ALREADY_SENT',send:false};
    return {status:'FOLLOW_UP_DUE',send:true,followupNumber:Number(state.sentCount || 1)};
  }
  const created = tmv3_preInspectionNotificationDateParts_(record.calendarCreatedAt);
  if (!created) return {status:'EVENT_CREATION_TIME_UNKNOWN',send:false};
  const appointment = tmv3_preInspectionNotificationDateParts_(start);
  if (tmv3_preInspectionNotificationWeek_(created) === tmv3_preInspectionNotificationWeek_(appointment)) {
    return {status:'SAME_WEEK_DUE',send:true,followupNumber:0};
  }
  const due = new Date(created.utc.getTime());
  let remaining = 2;
  while (remaining) {
    due.setUTCDate(due.getUTCDate() + 1);
    if (due.getUTCDay() !== 0 && due.getUTCDay() !== 6) remaining--;
  }
  const eligibleAt = due.toISOString().slice(0, 10) + created.text.slice(10);
  return today.text >= eligibleAt ? {status:'FIRST_DUE',send:true,followupNumber:0} :
    {status:'GRACE_PERIOD',send:false,eligibleAt:eligibleAt,timezone:'America/Toronto'};
}

/** Send boundary. Authoritative source, no hold, gate open — then the delivery steps below. */
function tmv3_preInspectionSendMissingDetails_(record) {
  if (!TMV3_PI_CLEANED_RELEASE_APPROVED) return {status:'PI_EMAILS_DISABLED_PENDING_RELEASE',sent:0};
  tmv3_piAssertLockedStandard_();
  if (!record || record.vertical !== 'PreInspection' || !tmv3_piIsAuthoritativeRecord_(record))
    return {status:'NONAUTHORITATIVE_SOURCE_NO_SEND',sent:0};
  if (tmv3_preInspectionApprovalHold_(record.eventId)) return {status:'OPERATOR_APPROVAL_HOLD_NO_SEND',sent:0};
  return tmv3_piSendMissingDetailsAfterRelease_(record);
}

function tmv3_piSendMissingDetailsAfterRelease_(record) {
  tmv3_piAssertLockedStandard_();
  if (tmv3_preInspectionApprovalHold_(record && record.eventId)) return {status:'APPROVAL_NOTIFICATION_HOLD',sent:0};
  const assessment = tmv3_preInspectionMissingDetailAssessment_(record);
  if (!assessment.shouldNotify) return {status:'NO_MISSING_DETAILS',sent:0};
  if (!tmv3_normEmail_(assessment.creatorEmail)) return {status:'UNVERIFIED_CREATOR_NO_SEND',sent:0};
  if (tmv3_piCreatorEmailExcluded_(assessment.creatorEmail)) return {status:'EXCLUDED_CREATOR_NO_SEND',sent:0};
  if (TMV3.MODE !== 'PRODUCTION_WRITE' || TMV3.OPERATIONS.preInspectionMissingDetailEmailsEnabled !== true)
    return {status:'DELIVERY_GATED',sent:0};
  const props = PropertiesService.getScriptProperties();
  const key = 'TMV3_PI_NOTIFICATION_' + tmv3_hash_(record.eventId);
  const prior = props.getProperty(key);
  let state = null;
  // Old SENT markers lack timestamps: start a conservative 24-hour wait instead of resending.
  if (prior === 'SENT') {
    state = {status:'SENT',sentCount:1,lastSentAt:new Date().toISOString(),migrated:true};
    props.setProperty(key, JSON.stringify(state));
    return {status:'ALREADY_SENT',sent:0};
  }
  if (prior) {
    try { state = JSON.parse(prior); } catch (err) { return {status:'UNCERTAIN_SEND_NO_RETRY',sent:0}; }
  }
  const decision = tmv3_preInspectionNotificationDecision_(record, state);
  if (!decision.send) return Object.assign({sent:0}, decision);
  const email = tmv3_preInspectionMissingDetailEmailPreview_(Object.assign({}, record, {notificationFollowupNumber:decision.followupNumber}));
  if (!email.recipient) return {status:'CREATOR_MISSING_OR_AMBIGUOUS_NO_SEND',sent:0};
  if (MailApp.getRemainingDailyQuota() < 3) return {status:'EMAIL_QUOTA_DEFERRED',sent:0};
  // SENDING is written first: a crash mid-send leaves an uncertain state that is never retried.
  const next = {status:'SENDING',sentCount:Number(state && state.sentCount || 0) + 1,lastSentAt:new Date().toISOString()};
  props.setProperty(key, JSON.stringify(next));
  MailApp.sendEmail({to:email.recipient,cc:email.cc.join(','),subject:email.subject,
    body:(decision.followupNumber ? 'FOLLOW-UP #' + decision.followupNumber + ': ' : '') +
      'Please update the missing Pre-Inspection details: ' + assessment.missing.map(function(x) { return x.message; }).join('; '),
    htmlBody:email.htmlBody,name:'Classic Fireplace & BBQ Store'});
  next.status = 'SENT';
  props.setProperty(key, JSON.stringify(next));
  tmv3_audit_('PreInspection', record.eventId, '', 'MISSING_DETAILS_EMAIL', 'SENT',
    JSON.stringify({to:email.recipient,followupNumber:decision.followupNumber,missing:assessment.missing}));
  return {status:'SENT',sent:1,to:email.recipient,followupNumber:decision.followupNumber,missing:assessment.missing};
}

/** Operator failure alert (separate template; never the creator email). One per day per event+error. */
function tmv3_preInspectionFailureAlert_(eventId, error) {
  if (!TMV3_PI_CLEANED_RELEASE_APPROVED) return {status:'PI_FAILURE_EMAIL_DISABLED_PENDING_RELEASE',sent:0};
  if (TMV3.MODE !== 'PRODUCTION_WRITE' || TMV3.OPERATIONS.preInspectionFailureAlertsEnabled !== true) return {status:'ALERT_GATED'};
  const day = Utilities.formatDate(new Date(), TMV3_TIMEZONE, 'yyyy-MM-dd');
  const key = 'TMV3_PI_FAILURE_' + tmv3_hash_(day + '|' + eventId + '|' + error);
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(key)) return {status:'ALREADY_ALERTED'};
  try {
    if (MailApp.getRemainingDailyQuota() < 1) return {status:'QUOTA_DEFERRED'};
    props.setProperty(key, 'SENDING');
    const esc = tmv3_preInspectionHtmlEscape_;
    MailApp.sendEmail({to:TMV3.OPERATIONS.preInspectionFailureAlertEmail,
      subject:'Pre-Inspection Automation: action required',
      body:'Pre-Inspection automation needs review. Event: ' + (eventId || 'Source refresh') + '\n' + error,
      htmlBody:'<p><em>This is an automated email notification from Classic Fireplace &amp; BBQ Store.</em></p>' +
        '<p>Hi Pramodh,</p><p style="color:#B42318;font-weight:700;">Pre-Inspection automation requires review.</p>' +
        '<p>Event: ' + esc(eventId || 'Source refresh') + '</p><p>' + esc(error) + '</p>' +
        '<p>The affected action remains queued. Verify any uncertain write before retrying.</p>',
      name:'Classic Fireplace & BBQ Store'});
    props.setProperty(key, 'SENT');
    return {status:'SENT'};
  } catch (err) {
    tmv3_audit_('PreInspection', eventId, '', 'PREINSPECTION_FAILURE_ALERT', 'ATTENTION', String(err));
    return {status:'ALERT_UNCERTAIN_NO_RETRY'};
  }
}

// =====================================================================
// §10 CYCLE — one guarded batch. Stages:
//   A preconditions (gate, standard, ops gate, lock)
//   B intake (targeted or authoritative) and API budget
//   C throttle, source refresh, pending-event recovery
//   D batch selection (durable sweep position or rotating cursor)
//   E notification deadlines for every authoritative event
//   F per-event pipeline (≤3 events, ≤1 business write per batch)
//   G persist cursor/sweep/pending/last result
// =====================================================================

function tmv3_preInspectionSweepKey_(record) {
  const start = new Date(record.start);
  if (!isFinite(start.getTime())) throw new Error('PREINSPECTION_SWEEP_INVALID_START');
  return start.toISOString() + '|' + String(record.eventId || '');
}

function tmv3_piPlanHasAction_(plan, actions) {
  return Array.isArray(plan.actions) && plan.actions.some(function(a) { return actions.indexOf(a) >= 0; });
}

function tmv3_piPlanCreates_(plan) {
  return tmv3_piPlanHasAction_(plan, [TMV3_STEP7_ACTION.CREATE_TASK, TMV3_STEP7_ACTION.RECREATE_TASK]);
}

/** Any action beyond link verification patches or creates a Task. */
function tmv3_piPlanTouchesTask_(plan) {
  return Array.isArray(plan.actions) && plan.actions.some(function(a) { return a !== TMV3_STEP7_ACTION.VERIFY_CALENDAR_LINKS; });
}

function tmv3_piPendingAdd_(cycle, eventId) {
  if (cycle.pending.indexOf(eventId) < 0) cycle.pending.push(eventId);
}

function tmv3_piPendingRemove_(cycle, eventId) {
  cycle.pending = cycle.pending.filter(function(id) { return id !== eventId; });
}

/** Stage C: re-queue pending events still in scope; drop those no longer PreInspection. */
function tmv3_piRecoverPendingEvents_(cycle, records) {
  cycle.pending.slice().forEach(function(id) {
    if (records.some(function(r) { return r.eventId === id; })) return;
    try {
      const recovered = tmv3_preInspectionRecoveryRecord_(id);
      if (recovered.calendarRole === 'PRIMARY_SHARED' ||
          tmv3_preInspectionClassifyLikely_(recovered, TMV3.VERTICALS.PreInspection).status === 'EXPLICIT_PREINSPECTION') records.push(recovered);
      else tmv3_piPendingRemove_(cycle, id);
    } catch (err) {
      tmv3_audit_('PreInspection', id, '', 'PREINSPECTION_RECOVERY', 'ATTENTION', String(err));
    }
  });
}

/** Stage D (sweep): durable position by start|eventId survives batch boundaries. */
function tmv3_piSelectSweepBatch_(cycle, selected) {
  const raw = cycle.props.getProperty('TMV3_PI_SWEEP_STATE') || '';
  let prior = null;
  try { if (raw) prior = JSON.parse(raw); } catch (ignored) {}
  const valid = prior && prior.status === 'IN_PROGRESS' &&
    isFinite(new Date(prior.startedAt).getTime()) &&
    cycle.now - new Date(prior.startedAt).getTime() < 86400000;
  if (cycle.options.sweepContinue === true && !valid) return null;
  cycle.sweep = valid ? prior : {status:'IN_PROGRESS',startedAt:new Date(cycle.now).toISOString(),lastKey:'',processed:0};
  selected.sort(function(a, b) {
    const aa = tmv3_preInspectionSweepKey_(a), bb = tmv3_preInspectionSweepKey_(b);
    return aa < bb ? -1 : aa > bb ? 1 : 0;
  });
  return selected.filter(function(r) { return tmv3_preInspectionSweepKey_(r) > cycle.sweep.lastKey; });
}

/** Stage F for one authoritative event: plan → guards → execute → notify → verify. */
function tmv3_piReconcileAuthoritativeEvent_(cycle, record, one, doneKey) {
  const result = cycle.result;
  const preservation = tmv3_restorePreInspectionLegacyTitle_(record);
  if (preservation.writes) record = tmv3_preInspectionRecoveryRecord_(record.eventId);

  // F1 plan from fresh evidence; optional SO ship-to may repair an ambiguous Location.
  let plans = tmv3_step7FreshPlansForEvent_('PreInspection', record.eventId);
  if (plans.length === 1 && plans[0].blocker && !tmv3_preInspectionHistoricalCoverageVerified_(plans[0])) {
    one.optionalOrderEvidence = tmv3_preInspectionOptionalOrderLocationEvidence_(record);
    if (one.optionalOrderEvidence.status === 'CACHE_REPAIRED') plans = tmv3_step7FreshPlansForEvent_('PreInspection', record.eventId);
  }
  // F2 CREATE is re-planned on complete fresh Customer Task history.
  let history = null;
  if (plans.length === 1 && !plans[0].blocker && tmv3_piPlanCreates_(plans[0])) {
    history = tmv3_piFreshCustomerHistory_(plans[0]);
    plans = tmv3_step7FreshPlansForEvent_('PreInspection', record.eventId);
  }
  if (plans.length !== 1) throw new Error('MULTIPLE_CANONICAL_PLANS_NO_WRITE');
  const plan = plans[0];

  // F3 guards before any Task mutation.
  // Operator rule 2026-11-11: reuse the OPEN Task for this visit; if none is open, create a new
  // one — but never when a Task was already completed on the appointment day. RECREATE_TASK is
  // planned only after the fresh complete history above shows no open Task.
  if (!plan.blocker && Array.isArray(plan.actions) && plan.actions.indexOf(TMV3_STEP7_ACTION.RECREATE_TASK) >= 0) {
    if (plan.disposition === 'FULFILLED_NO_RECREATE') throw new Error('PI_SAME_VISIT_FULFILLED_NO_RECREATE');
    tmv3_piAssertNoSameDayCompletedTask_(record, history);
    one.recreateAuthorized = 'NO_OPEN_TASK_NEW_VISIT_OPERATOR_RULE_2026_11_11';
  }
  if (!plan.blocker && tmv3_piPlanTouchesTask_(plan)) tmv3_piAssertRequestedByPlan_(plan, record);
  if (!plan.blocker && tmv3_piPlanCreates_(plan)) one.orderStatusBeforeCreate = tmv3_preInspectionOrderBeforeCreate_(record, plan);
  one.plan = plan.plan; one.blocker = plan.blocker || ''; one.taskId = plan.taskId || '';

  // F4 one business transaction per batch; a deferred event keeps its sweep position.
  if (!plan.blocker && plan.actions.length && result.writes >= 1) return 'DEFER';
  if (!plan.blocker && plan.actions.length && result.writes < 1) {
    if (tmv3_piPlanTouchesTask_(plan)) result.writes++;
    one.execution = tmv3_executeFreshStep7Selection_({previousPlan:plan.plan,contract:plan}, 'PREINSPECTIONAUTO', 'ALL');
    if (!/VERIFIED|CONVERGED|NO_CHANGE/.test(String(one.execution.status || ''))) {
      throw new Error('PREINSPECTION_RECONCILIATION_NOT_VERIFIED: ' + one.execution.status);
    }
  }

  // F5 notification on the freshly re-read event.
  const fresh = tmv3_findFreshEventRecord_('PreInspection', record.eventId);
  one.notification = tmv3_preInspectionSendMissingDetails_(fresh);
  if (/CREATOR_MISSING|UNCERTAIN_SEND/.test(one.notification.status || '')) {
    result.status = 'ATTENTION';
    tmv3_preInspectionFailureAlert_(record.eventId, one.notification.status);
  }

  // F6 independent read-back; only a converged plan leaves the pending queue.
  const after = tmv3_step7FreshPlansForEvent_('PreInspection', record.eventId);
  one.afterPlans = after.map(function(p) { return {plan:p.plan,blocker:p.blocker,taskId:p.taskId,actions:p.actions}; });
  if (tmv3_preInspectionCompletionVerified_(after, one.execution)) {
    if (tmv3_preInspectionHistoricalCoverageVerified_(after[0])) {
      one.status = 'FULFILLED_NO_RECREATE'; one.historicalEvidence = one.blocker; one.blocker = '';
    }
    tmv3_piPendingRemove_(cycle, record.eventId);
    const inferredKey = 'TMV3_PI_INFERRED_' + tmv3_hash_(record.eventId);
    if (cycle.props.getProperty(inferredKey)) cycle.props.deleteProperty(inferredKey);
    cycle.props.setProperty(doneKey, JSON.stringify({fingerprint:fresh.fingerprint,at:Date.now()}));
  } else {
    one.status = 'ATTENTION';
    result.status = 'ATTENTION';
    tmv3_piPendingAdd_(cycle, record.eventId);
  }
  return 'DONE';
}

/** Stage F wrapper. Returns PUSH (record one), SKIP_BREAK (deferred, not counted) or PUSH_BREAK. */
function tmv3_piProcessCycleEvent_(cycle, record, one) {
  const result = cycle.result;
  try {
    if (tmv3_preInspectionApprovalHold_(record.eventId)) {
      one.status = 'APPROVED_INTAKE_TASK_AND_NOTIFICATION_HELD';
      one.notification = {status:'APPROVAL_NOTIFICATION_HOLD',sent:0};
      return 'PUSH';
    }
    const doneKey = 'TMV3_PI_DONE_' + tmv3_hash_(record.eventId);
    const done = JSON.parse(cycle.props.getProperty(doneKey) || 'null');
    if (!cycle.options.eventId && done && done.fingerprint === record.fingerprint && cycle.now - Number(done.at || 0) < 3600000) {
      one.status = 'VERIFIED_UNCHANGED_CACHE';
      return 'PUSH';
    }
    const intake = tmv3_preInspectionPromoteCandidate_(record);
    one.intake = intake.status;
    record = intake.record;
    if (intake.status !== 'PRIMARY') {
      tmv3_piPendingAdd_(cycle, record.eventId);
      one.notification = {status:'WAITING_FOR_INTAKE',sent:0}; // wait for authoritative intake and matching
    } else if (tmv3_piReconcileAuthoritativeEvent_(cycle, record, one, doneKey) === 'DEFER') {
      result.processed--;
      result.deferredWriteBudget = true;
      return 'SKIP_BREAK';
    }
    result.notifications += Number(one.notification && one.notification.sent || 0);
    return 'PUSH';
  } catch (err) {
    one.status = 'ATTENTION'; one.error = String(err && err.message || err); result.status = 'ATTENTION';
    tmv3_piPendingAdd_(cycle, record.eventId);
    tmv3_audit_('PreInspection', record.eventId, '', 'PREINSPECTIONAUTO', 'ATTENTION', one.error);
    tmv3_preInspectionFailureAlert_(record.eventId, one.error);
    if (/TMV3_STRIVEN_API_DAILY_GUARD|daily api usage limit/i.test(one.error)) {
      // Preserve this event's cursor position for a fresh reconciliation after reset.
      result.status = 'API_BUDGET_WAIT'; result.processed--;
      return 'PUSH_BREAK';
    }
    return 'PUSH';
  }
}

/** Stage G: durable sweep position or rotating cursor, pending queue, last result. */
function tmv3_piPersistCycle_(cycle, selected, cursor) {
  const result = cycle.result, props = cycle.props, sweep = cycle.sweep;
  if (sweep) {
    if (result.processed) sweep.lastKey = tmv3_preInspectionSweepKey_(selected[result.processed - 1]);
    sweep.processed = Number(sweep.processed || 0) + result.processed;
    const remaining = Math.max(0, selected.length - result.processed);
    sweep.status = remaining ? 'IN_PROGRESS' : (cycle.pending.length ? 'SCANNED_WITH_PENDING_REVIEW' : 'COMPLETE');
    props.setProperty('TMV3_PI_SWEEP_STATE', JSON.stringify(sweep));
    result.sweep = {status:sweep.status,eligible:result.eligible,attempted:sweep.processed,
      remaining:remaining,pendingReview:cycle.pending.length,lastKey:sweep.lastKey};
    // A successful batch is not a completed queue; keep honest backlog state.
    if (result.status === 'COMPLETE' && remaining > 0) result.status = 'IN_PROGRESS';
    else if (result.status === 'COMPLETE' && cycle.pending.length > 0) result.status = 'ATTENTION_PENDING_REVIEW';
  } else if (!cycle.options.eventId && selected.length) {
    props.setProperty('TMV3_PI_CURSOR', String((cursor + result.processed) % selected.length));
  }
  props.setProperty('TMV3_PI_PENDING_EVENTS', JSON.stringify(cycle.pending));
  props.setProperty('TMV3_PI_LAST_RESULT', JSON.stringify({status:result.status,
    processed:result.processed,writes:result.writes,notifications:result.notifications,
    events:result.results.map(function(x) {
      return {eventId:x.eventId,status:x.status || x.plan || x.intake,taskId:x.taskId,error:String(x.error || '').slice(0, 500)};
    })}));
}

function tmv3_piCycleBody_(cycle) {
  const options = cycle.options, props = cycle.props, result = cycle.result;
  // Stage B
  const scopeEventId = tmv3_clean_(TMV3.OPERATIONS.preInspectionAutomationCanaryEventId || options.eventId);
  const sweepRequested = options.sweepStart === true || options.sweepContinue === true;
  if (sweepRequested && scopeEventId) return {status:'SWEEP_BLOCKED_BY_CANARY_SCOPE',processed:0,writes:0};
  const records = scopeEventId ? tmv3_preInspectionTargetedIntake_(scopeEventId, props) :
    tmv3_preInspectionIntakeRecords_({recordDiscovery:true,resolveLikely:true});
  result.discovery = tmv3_preInspectionDiscoveryStatus();
  const budget = tmv3_strivenApiBudgetStatus_();
  if (budget.exhausted || budget.remaining < 30) {
    result.status = 'API_BUDGET_WAIT'; result.apiBudget = budget;
    result.resume = 'NEXT_SCHEDULED_SLOT_AFTER_DAILY_RESET';
    const pauseKey = 'TMV3_PI_BUDGET_PAUSE_DATE';
    if (props.getProperty(pauseKey) !== budget.date) {
      props.setProperty(pauseKey, budget.date);
      tmv3_audit_('PreInspection', '', '', 'PREINSPECTION_API_BUDGET_WAIT', 'WAIT', JSON.stringify({date:budget.date,
        count:budget.count,softLimit:budget.softLimit,resume:result.resume,triggers:tmv3_listTriggers()}));
    }
    props.setProperty('TMV3_PI_LAST_RESULT', JSON.stringify(result));
    return result;
  }

  // Stage C
  cycle.now = Date.now();
  const last = Number(props.getProperty('TMV3_PI_LAST_CYCLE') || 0);
  if (options.force !== true &&
      cycle.now - last < Math.max(1, Number(TMV3.OPERATIONS.preInspectionAutomationIntervalMinutes || 5)) * 60000) {
    return {status:'THROTTLED',writes:0};
  }
  props.setProperty('TMV3_PI_LAST_CYCLE', String(cycle.now));
  // Cache-age gates and last-good policy; a refresh failure stops before creates or emails.
  tmv3_refreshOperationalSourcesIfNeeded_(120, false);
  cycle.pending = JSON.parse(props.getProperty('TMV3_PI_PENDING_EVENTS') || '[]');
  if (!scopeEventId) tmv3_piRecoverPendingEvents_(cycle, records);

  // Stage D
  let selected = scopeEventId ? records.filter(function(r) { return r.eventId === scopeEventId; }) : records;
  if (sweepRequested) {
    selected = tmv3_piSelectSweepBatch_(cycle, selected);
    if (!selected) return {status:'NO_ACTIVE_PREINSPECTION_SWEEP',processed:0,writes:0};
    result.eligible = records.length;
  }

  // Stage E: slots evaluate deadlines for all authoritative events; continuations do not rescan mail.
  const notificationRecords = options.sweepContinue === true ? [] : (scopeEventId ? selected : records);
  notificationRecords.filter(function(r) {
    return r.calendarRole === 'PRIMARY_SHARED' && !tmv3_preInspectionApprovalHold_(r.eventId);
  }).forEach(function(r) {
    try { result.notifications += Number(tmv3_preInspectionSendMissingDetails_(r).sent || 0); }
    catch (err) { result.status = 'ATTENTION'; tmv3_preInspectionFailureAlert_(r.eventId, String(err)); }
  });

  // Stage F: tested safety budgets — at most three events and one business write.
  let cursor = cycle.sweep ? 0 : Number(props.getProperty('TMV3_PI_CURSOR') || 0);
  const cap = Math.max(1, Math.min(3, Number(TMV3.OPERATIONS.preInspectionAutomationMaxEventsPerRun || 3)));
  if (selected.length) cursor = cursor % selected.length;
  for (let i = 0; i < Math.min(cap, selected.length); i++) {
    const record = selected[(cursor + i) % selected.length];
    const one = {eventId:record.eventId};
    result.processed++;
    const outcome = tmv3_piProcessCycleEvent_(cycle, record, one);
    if (outcome === 'SKIP_BREAK') break;
    result.results.push(one);
    if (outcome === 'PUSH_BREAK') break;
  }

  // Stage G
  tmv3_piPersistCycle_(cycle, selected, cursor);
  return result;
}

function tmv3_preInspectionAutomationCycle_(options) {
  options = options || {};
  if (!TMV3_PI_CLEANED_RELEASE_APPROVED) {
    return {status:'PI_STAGING_RELEASE_GATE_NO_BUSINESS_WRITES',
      eventId:options.eventId || '',processed:0,writes:0,notifications:0,
      nextAction:'RUN_PI_CLEANED_VALIDATION_AND_VERIFY_FULL_V3_DEPENDENCIES'};
  }
  // Stage A
  tmv3_piAssertLockedStandard_();
  if (!tmv3_operationWritesEnabled_('PREINSPECTIONAUTO')) return {status:'PREINSPECTION_GATED',writes:0};
  if (TMV3.OPERATIONS.preInspectionAutomationScheduleHours && options.force !== true) return {status:'SCHEDULE_ONLY',writes:0};
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(options.force === true ? 30000 : 1000)) return {status:'LOCK_BUSY',writes:0};
  const cycle = {options:options,props:PropertiesService.getScriptProperties(),
    result:{status:'COMPLETE',processed:0,writes:0,notifications:0,results:[]},
    now:Date.now(),pending:[],sweep:null};
  const previousApiContext = TMV3_PREINSPECTION_API_CONTEXT_ACTIVE;
  TMV3_PREINSPECTION_API_CONTEXT_ACTIVE = true;
  try {
    return tmv3_piCycleBody_(cycle);
  } catch (err) {
    const result = cycle.result;
    result.status = 'ATTENTION'; result.error = String(err && err.message || err);
    tmv3_audit_('PreInspection', '', '', 'PREINSPECTION_CYCLE', 'ATTENTION', result.error);
    tmv3_preInspectionFailureAlert_('', result.error);
    return result;
  } finally {
    TMV3_PREINSPECTION_API_CONTEXT_ACTIVE = previousApiContext;
    lock.releaseLock();
  }
}

// =====================================================================
// §11 SCHEDULE — 08:00 / 13:00 / 17:00 America/Toronto daily slots.
// Retired handlers keep compatible no-op stubs for old installed triggers.
// =====================================================================

function tmv3_preInspectionCheck0800() { return tmv3_runPreInspectionDailySlot_('08:00'); }
function tmv3_preInspectionCheck1300() { return tmv3_runPreInspectionDailySlot_('13:00'); }
function tmv3_preInspectionCheck1700() { return tmv3_runPreInspectionDailySlot_('17:00'); }

function tmv3_preInspectionScheduledCycle() {
  return {status:'RETIRED_FIVE_MINUTE_SCHEDULE',writes:0};
}

function tmv3_preInspectionContinueEligibleSweep() {
  return {status:'RETIRED_OUTSIDE_STANDARD_SLOTS',writes:0,notifications:0};
}

function tmv3_preInspectionScheduleNextSweep_() {
  return 'DEFERRED_TO_NEXT_STANDARD_SLOT';
}

function tmv3_preInspectionClearContinuationTrigger_() {
  const gate = tmv3_piStagingDisabled_(); if (gate) return gate;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'tmv3_preInspectionContinueEligibleSweep') ScriptApp.deleteTrigger(t);
  });
}

function tmv3_installPreInspectionAutomationTrigger() {
  const gate = tmv3_piStagingDisabled_(); if (gate) return gate;
  return tmv3_installStandardDailyTriggers();
}

function tmv3_preInspectionLogSweep_(slot, result) {
  const summary = {
    slot:slot,status:result.status,processed:result.processed,
    writes:result.writes,notifications:result.notifications,
    sweep:result.sweep || null,continuation:result.continuation || null,
    error:result.error || '',
    calendarReview:result.calendarReview && result.calendarReview.status || 'NOT_RUN',
    operatorQueue:result.operatorQueue && result.operatorQueue.status || 'NOT_RUN',
    attentionEvents:(result.results || []).filter(function(r) {
      return r.status === 'ATTENTION' || !!r.blocker || !!r.error;
    }).map(function(r) { return {eventId:r.eventId,status:r.status || '',blocker:r.blocker || '',error:r.error || ''}; })
  };
  Logger.log(JSON.stringify(summary));
  tmv3_audit_('PreInspection', '', '', 'PREINSPECTION_SWEEP', result.status || 'UNKNOWN', JSON.stringify(summary));
}

function tmv3_preInspectionMaybeContinueSweep_(result) {
  const errors = [result.error || ''].concat((result.results || []).map(function(r) { return r.error || ''; })).join(' ');
  if (result.status === 'API_BUDGET_WAIT' || /TMV3_STRIVEN_API_DAILY_GUARD|daily api usage limit/i.test(errors)) {
    result.continuation = 'WAIT_FOR_DAILY_API_BUDGET';
    return result;
  }
  // A reviewed historical relationship is not an infrastructure failure.
  const allTechnicalFailures = (result.results || []).length > 0 &&
    result.results.every(function(r) { return !!r.error; });
  if (allTechnicalFailures) {
    result.continuation = 'PAUSED_TECHNICAL_FAILURE_AUTOMATIC_RECOVERY';
    return result;
  }
  if ((result.sweep && result.sweep.remaining > 0) || result.status === 'LOCK_BUSY') {
    try { result.continuation = tmv3_preInspectionScheduleNextSweep_(); }
    catch (err) {
      result.status = 'ATTENTION'; result.continuation = 'SCHEDULING_FAILED';
      result.error = String(err && err.message || err);
      tmv3_audit_('PreInspection', '', '', 'ELIGIBLE_SWEEP', 'ATTENTION', result.error);
      tmv3_preInspectionFailureAlert_('', result.error);
    }
  }
  return result;
}

/** One slot: read-only Calendar review, batches until done or ~4 minutes, then publish views. */
function tmv3_runPreInspectionDailySlot_(slot) {
  const gate = tmv3_piStagingDisabled_(); if (gate) return gate;
  tmv3_piAssertLockedStandard_();
  const sheetRefreshStarted = Date.now();
  // Report-only Calendar review; failure keeps the last-good view and never blocks Striven work.
  let calendarReview = null;
  try { calendarReview = tmv3_piCalendarReviewRefresh_(); }
  catch (err) {
    calendarReview = {status:'PI_REVIEW_ATTENTION',error:String(err && err.message || err).slice(0, 400)};
    try { tmv3_audit_('PreInspection', '', '', 'PI_CALENDAR_REVIEW', 'ATTENTION', calendarReview.error); } catch (ignored) {}
  }
  let result = tmv3_preInspectionAutomationCycle_({force:true,sweepStart:true,reason:'DAILY_SLOT_' + slot});
  const totals = {processed:Number(result.processed || 0),writes:Number(result.writes || 0),
    notifications:Number(result.notifications || 0),results:(result.results || []).slice()};
  let priorRemaining = result.sweep && result.sweep.remaining;
  const batchStatuses = [result.status];
  while (result.sweep && result.sweep.remaining > 0 && result.processed > 0 &&
      result.status !== 'API_BUDGET_WAIT' && result.status !== 'LOCK_BUSY' &&
      Date.now() - sheetRefreshStarted < 240000) {
    const next = tmv3_preInspectionAutomationCycle_({force:true,sweepContinue:true,reason:'DAILY_SLOT_' + slot});
    batchStatuses.push(next.status);
    totals.processed += Number(next.processed || 0); totals.writes += Number(next.writes || 0);
    totals.notifications += Number(next.notifications || 0); totals.results.push.apply(totals.results, next.results || []);
    if (!next.sweep || next.sweep.remaining >= priorRemaining || !next.processed) { result = next; break; }
    priorRemaining = next.sweep.remaining; result = next;
  }
  result.batchStatuses = batchStatuses;
  result.processed = totals.processed; result.writes = totals.writes;
  result.notifications = totals.notifications; result.results = totals.results;
  if (result.status === 'LOCK_BUSY') PropertiesService.getScriptProperties().setProperty('TMV3_PI_RETRY_SWEEP_START', '1');
  tmv3_preInspectionMaybeContinueSweep_(result);
  if (result.sweep && result.sweep.remaining > 0 && result.status === 'COMPLETE') result.status = 'IN_PROGRESS';
  if (result.sweep && result.sweep.remaining === 0 && result.sweep.pendingReview > 0 && result.status === 'COMPLETE')
    result.status = 'ATTENTION_PENDING_REVIEW';
  if (calendarReview && calendarReview.status === 'PI_REVIEW_PARTIAL_SOURCE_ACCESS' && result.status === 'COMPLETE')
    result.status = 'ATTENTION_CALENDAR_COVERAGE';
  result.sheetRefresh = tmv3_refreshAutomationSheetsAfter_(sheetRefreshStarted);
  try { result.operatorQueue = tmv3_preInspectionRefreshOperatorView_(); }
  catch (err) {
    result.operatorQueue = {status:'PI_OPERATOR_QUEUE_ATTENTION',error:String(err && err.message || err).slice(0, 350)};
    try { tmv3_audit_('PreInspection', '', '', 'PI_OPERATOR_QUEUE', 'ATTENTION', result.operatorQueue.error); } catch (ignore) {}
  }
  result.calendarReview = calendarReview;
  if (result.status === 'COMPLETE' && (!result.operatorQueue || result.operatorQueue.status !== 'PI_OPERATOR_VIEW_REFRESHED'))
    result.status = 'ATTENTION_OPERATOR_QUEUE';
  if (result.status === 'COMPLETE' && (!calendarReview ||
      !/^(?:PI_REVIEW_REFRESHED|PI_REVIEW_UNCHANGED)$/.test(String(calendarReview.status || ''))))
    result.status = 'ATTENTION_CALENDAR_COVERAGE';
  tmv3_preInspectionLogSweep_(slot, result); // audit only after both publication outcomes are known
  return result;
}

/** Stalled sweep recovery inside business hours; reuses the scheduled machinery. */
function tmv3_preInspectionAutomaticRecovery_(reason) {
  const gate = tmv3_piStagingDisabled_(); if (gate) return gate;
  const hour = Number(Utilities.formatDate(new Date(), TMV3_TIMEZONE, 'H'));
  if (hour < 8 || hour > 18) return {status:'OUTSIDE_PREINSPECTION_RECOVERY_HOURS',writes:0};
  const props = PropertiesService.getScriptProperties();
  let state = null;
  try { state = JSON.parse(props.getProperty('TMV3_PI_SWEEP_STATE') || 'null'); } catch (ignored) {}
  if (!state || state.status !== 'IN_PROGRESS') return {status:'NO_STALLED_PREINSPECTION_SWEEP',writes:0};
  const last = Number(props.getProperty('TMV3_PI_LAST_CYCLE') || 0);
  if (Date.now() - last < 20 * 60000) return {status:'RECENT_PREINSPECTION_BATCH',writes:0};
  tmv3_preInspectionClearContinuationTrigger_();
  const result = tmv3_preInspectionAutomationCycle_({force:true,sweepContinue:true,reason:'AUTOMATIC_RECOVERY_' + reason});
  tmv3_preInspectionMaybeContinueSweep_(result);
  tmv3_preInspectionLogSweep_('RECOVERY', result);
  return result;
}

/** Menu: Run Pre-Inspection Automation Now. */
function tmv3_runPreInspectionAutomation() {
  const gate = tmv3_piStagingDisabled_(); if (gate) return gate;
  const sheetRefreshStarted = Date.now();
  const result = tmv3_preInspectionAutomationCycle_({force:true,reason:'MANUAL_KICK'});
  result.sheetRefresh = tmv3_refreshAutomationSheetsAfter_(sheetRefreshStarted);
  Logger.log(JSON.stringify(result));
  return result;
}

function tmv3_reconcilePreInspectionEvent(eventId) {
  const gate = tmv3_piStagingDisabled_(); if (gate) return gate;
  if (!/^[a-z0-9_]+@google\.com$/.test(String(eventId || ''))) throw new Error('INVALID_PREINSPECTION_EVENT_ID');
  const started = Date.now();
  const result = tmv3_preInspectionAutomationCycle_({force:true,eventId:eventId,reason:'TARGETED_RECONCILIATION'});
  result.sheetRefresh = tmv3_refreshAutomationSheetsAfter_(started);
  return result;
}

/** Clears pending events whose completed Task is proven by fresh history. No business writes. */
function tmv3_reconcilePreInspectionPendingReview() {
  const gate = tmv3_piStagingDisabled_(); if (gate) return gate;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return {status:'LOCK_BUSY',businessWrites:0};
  try {
    const props = PropertiesService.getScriptProperties();
    const stored = JSON.parse(props.getProperty('TMV3_PI_PENDING_EVENTS') || '[]');
    if (!Array.isArray(stored) || stored.some(function(id) { return typeof id !== 'string' || !id; })) {
      throw new Error('INVALID_PENDING_QUEUE_NO_CHANGE');
    }
    let pending = stored.slice();
    const results = [];
    stored.filter(function(id, index) { return stored.indexOf(id) === index; }).slice(0, 5).forEach(function(id) {
      try {
        const plans = tmv3_step7FreshPlansForEvent_('PreInspection', id);
        if (plans.length === 1 && tmv3_preInspectionHistoricalCoverageVerified_(plans[0])) {
          pending = pending.filter(function(candidate) { return candidate !== id; });
          results.push({eventId:id,taskId:plans[0].taskId,status:'FULFILLED_NO_RECREATE',readbackVerified:true,businessWrites:0});
        } else {
          results.push({eventId:id,status:'REVIEW_RETAINED',plans:plans.map(function(p) {
            return {plan:p.plan,blocker:p.blocker,taskId:p.taskId};
          }),businessWrites:0});
        }
      } catch (err) {
        results.push({eventId:id,status:'READ_FAILED_REVIEW_RETAINED',error:String(err),businessWrites:0});
      }
    });
    props.setProperty('TMV3_PI_PENDING_EVENTS', JSON.stringify(pending));
    const readback = JSON.parse(props.getProperty('TMV3_PI_PENDING_EVENTS') || '[]');
    if (JSON.stringify(readback) !== JSON.stringify(pending)) throw new Error('PENDING_QUEUE_READBACK_FAILED');
    const result = {status:pending.length ? 'REVIEW_REMAINS' : 'COMPLETE',
      before:stored.length,remaining:pending.length,cleared:stored.length - pending.length,
      results:results,businessWrites:0};
    props.setProperty('TMV3_PI_LAST_PENDING_RECONCILIATION', JSON.stringify(result));
    tmv3_audit_('PreInspection', '', '', 'PENDING_HISTORY_RECONCILE', result.status, JSON.stringify(result));
    Logger.log(JSON.stringify(result));
    return result;
  } finally { lock.releaseLock(); }
}

// =====================================================================
// §12 OPERATOR — read-only diagnostics and review sheets. None of these
// create/patch Tasks, change Calendar events or guests, or send email.
// =====================================================================

/** Fast cache-only diagnostic of the PreInspection operator sheet, three events per call. */
function tmv3_verifyPreInspectionAutomation(eventId) {
  const startedMs = Date.now();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1500)) {
    const busy = {status:'PI_VERIFY_LOCK_BUSY',businessWrites:0};
    Logger.log('PI_VERIFICATION_BATCH=' + JSON.stringify(busy));
    return busy;
  }
  try {
    const today = Utilities.formatDate(new Date(), TMV3_TIMEZONE, 'yyyy-MM-dd');
    const rows = tmv3_rows_(TMV3.VERTICALS.PreInspection.sheet);
    const startText = function(row) {
      const value = row['Calendar Start'];
      return value instanceof Date ? Utilities.formatDate(value, TMV3_TIMEZONE, 'yyyy-MM-dd HH:mm:ss') : String(value || '');
    };
    const queue = rows.filter(function(row) {
      if (String(row['Vertical'] || '') !== 'PreInspection' || String(row['Calendar Role'] || '') !== 'PRIMARY_SHARED') return false;
      const start = startText(row);
      return !/^\d{4}-\d{2}-\d{2}/.test(start) || start.slice(0, 10) >= today;
    }).sort(function(a, b) {
      return startText(a).localeCompare(startText(b)) || String(a['Event ID'] || '').localeCompare(String(b['Event ID'] || ''));
    });
    const seen = {};
    queue.forEach(function(row) {
      const id = String(row['Event ID'] || '');
      if (!id || seen[id]) throw new Error('PI_DIAGNOSTIC_EVENT_ID_EMPTY_OR_DUPLICATE');
      seen[id] = true;
    });
    const target = String(eventId || '').trim();
    const selectedQueue = target ? queue.filter(function(row) { return String(row['Event ID']) === target; }) : queue;
    const fingerprint = tmv3_hash_(today + '|' + queue.map(function(row) {
      return String(row['Event ID']) + '|' + startText(row) + '|' + String(row['Fingerprint'] || '');
    }).join('|'));
    const props = PropertiesService.getScriptProperties();
    const cursorKey = 'TMV3_PI_VERIFY_CURSOR';
    let offset = 0;
    if (!target) {
      let state = null;
      try { state = JSON.parse(props.getProperty(cursorKey) || 'null'); } catch (ignored) {}
      if (state && state.fingerprint === fingerprint && Number.isInteger(state.offset) &&
          state.offset >= 0 && state.offset < queue.length) offset = state.offset;
    }
    const batch = target ? selectedQueue.slice(0, 1) : selectedQueue.slice(offset, offset + 3);
    const records = batch.map(tmv3_piVerifyDiagnosticRow_);
    const hasMore = !target && offset + batch.length < queue.length;
    const result = {
      status:target && !batch.length ? 'TARGET_NOT_IN_OPERATOR_VIEW' :
        !queue.length ? 'NO_CURRENT_OPERATOR_EVENTS_REVIEW_REFRESH' : 'CACHE_ONLY_BATCH',
      version:TMV3.VERSION,
      source:'PREINSPECTION_OPERATOR_SNAPSHOT',
      evidence:'CACHE_ONLY_NOT_LIVE_STRIVEN_OR_CALENDAR_VERIFICATION',
      automationEnabled:TMV3.OPERATIONS.preInspectionAutomationEnabled,
      emailEnabled:TMV3.OPERATIONS.preInspectionMissingDetailEmailsEnabled,
      total:queue.length,
      processed:batch.length,
      startIndex:target ? null : offset,
      nextIndex:hasMore ? offset + batch.length : null,
      remaining:target ? null : Math.max(0, queue.length - offset - batch.length),
      hasMore:hasMore,
      records:records,
      elapsedMs:Date.now() - startedMs,
      strivenCalls:0,calendarReads:0,calendarWrites:0,taskWrites:0,emailsSent:0,
      diagnosticCursorPropertyOnly:!target
    };
    if (!target) {
      if (hasMore) props.setProperty(cursorKey, JSON.stringify({fingerprint:fingerprint,offset:offset + batch.length}));
      else props.deleteProperty(cursorKey);
    }
    Logger.log('PI_VERIFICATION_BATCH=' + JSON.stringify(result));
    return result;
  } catch (error) {
    Logger.log('PI_VERIFICATION_ERROR=' + String(error && error.message || error).slice(0, 300));
    throw error;
  } finally { lock.releaseLock(); }
}

function tmv3_piVerifyDiagnosticRow_(row) {
  const startValue = row['Calendar Start'];
  const start = startValue instanceof Date ? Utilities.formatDate(startValue, TMV3_TIMEZONE, 'yyyy-MM-dd HH:mm:ss') : String(startValue || '');
  const linked = String(row['Existing Task Link'] || '');
  const detail = String(row['Task'] || '');
  const linkedId = (/[?&]TaskID=(\d+)/i.exec(linked) || [])[1] || '';
  const detailId = (/\bTask\s*#\s*(\d+)/i.exec(detail) || [])[1] || '';
  const taskId = linkedId || detailId;
  const mapping = String(row['Mapping Status'] || '');
  const checklist = String(row['Data Checklist'] || '').split(/\r?\n/);
  const warnings = checklist.filter(function(line) { return line.indexOf('⚠') >= 0; }).slice(0, 5);
  const pending = checklist.filter(function(line) { return line.indexOf('⏳') >= 0; }).slice(0, 5);
  let disposition;
  if (linkedId && detailId && linkedId !== detailId) disposition = 'TASK_ID_CONFLICT_REVIEW_NO_WRITE';
  else if (/\b(?:NOT_RUN|REVIEW|BLOCKED)\b/i.test(mapping)) disposition = 'OPERATOR_REVIEW_REQUIRED';
  else if (/\b(?:CREATE_TASK|RECREATE_TASK)\b/i.test(mapping)) disposition = 'FRESH_TASK_HISTORY_REQUIRED_NO_CREATE';
  else if (!linkedId && taskId) disposition = 'TASK_LINK_MISSING_LIVE_CHECK_REQUIRED';
  else if (!taskId) disposition = 'TASK_ID_NOT_IN_OPERATOR_SNAPSHOT';
  else if (warnings.length) disposition = 'OPERATOR_WARNING_LIVE_CHECK_REQUIRED';
  else disposition = 'CACHED_MATCH_LIVE_NOT_VERIFIED';
  return {eventId:String(row['Event ID'] || ''),customerNumber:String(row['Customer # (Calendar)'] || ''),
    start:start,taskId:taskId,taskLinkPresent:!!linkedId,mapping:mapping,disposition:disposition,
    warnings:warnings,pending:pending,nextAction:String(row['Issue / Next Action'] || '').slice(0, 180)};
}

/** Representative calendars scanned (read-only) for the PI Calendar Review sheet. */
function tmv3_piReviewCalendarSources_() {
  return [
    {name:'CF Preinspects',id:TMV3.VERTICALS.PreInspection.primaryCalendarId},
    {name:'Stephen',id:'classicfireplace.ca_c20qcqfhvjbv784asn9pvuiaf4@group.calendar.google.com'},
    {name:'Jay',id:'jay@classicfireplace.ca'},
    {name:'Karen',id:'karen@classicfireplace.ca'}
  ];
}

/** Classify one logical event (all its calendar copies) for the review sheet. */
function tmv3_piReviewClassify_(copies, coverageErrors) {
  const shared = copies['CF Preinspects'];
  const ev = shared || copies.Stephen || copies.Karen || copies.Jay;
  const text = ev.title + ' ' + ev.description;
  const knownOther = (ev.creators || []).some(function(email) {
    return /(?:vpg05s56lp9f7kaegsiqask34o|2pbtbkjl1g2sos1fspgfq029h0|2h90kflud3re121hga9bvsi33o|service@classicfireplace\.ca|steve\.jensen83@gmail\.com)/i.test(email);
  });
  const internal = /do not book|newmarket|team meeting|flash|rylander|layout meeting|follow up appointment|meet with steve|vacation|personal|holiday|golf club/i.test(ev.title);
  const serviceOnly = /\bservice\b/i.test(ev.title) && !/\bpre[\s-]?inspect/i.test(text);
  const ownerVisit = Boolean(copies.Stephen || copies.Jay || copies.Karen);
  const visitClue = /\bpre[\s-]?inspect|\bon[\s-]?site\b|\bsite visit\b|\bmeasure\b|\binspect\b/i.test(text);
  const productClue = Boolean(ev.location) && /\b(fireplace|insert|gas unit|stove|mantle|chimney|remove|replace)\b/i.test(text);
  const customerClue = /(?:\b\d{5}\b|(?:\d{3}[-. ]?\d{3}[-. ]?\d{4}))/i.test(text);
  const linked = /Tasks\/TaskInfo\.aspx\?TaskID|Task\s*#\s*\d+/i.test(ev.description);
  let status = '', reason = '', mirror = 'NOT_EXPECTED';
  if (shared) {
    const st = copies.Stephen;
    const exact = st && ['title','location','description','start','end'].every(function(field) {
      return String(shared[field] || '') === String(st[field] || '');
    });
    mirror = st ? (exact ? 'EXACT' : 'MISMATCH') :
      coverageErrors.some(function(x) { return x.source === 'Stephen'; }) ? 'ACCESS_UNVERIFIED' : 'NOT_PRESENT';
    status = mirror === 'MISMATCH' ? 'REVIEW_MIRROR_MISMATCH' :
      mirror === 'ACCESS_UNVERIFIED' ? 'REVIEW_MIRROR_ACCESS' :
      linked ? 'SHARED_LINKED' : 'SHARED_UNLINKED';
    reason = linked ? 'Task link is present; verify Striven before marking fully complete' :
      'Task link absent; search existing Striven Task before any create';
  } else if (knownOther) {
    status = 'EXCLUDED_OTHER_WORKFLOW';
    reason = 'Creator identifies Installation/Delivery/Service or contractor workflow; no automatic promotion';
  } else if (internal) { status = 'EXCLUDED_NON_CUSTOMER'; reason = 'Internal block/meeting/non-customer'; }
  else if (serviceOnly) { status = 'EXCLUDED_SERVICE'; reason = 'Service-specific appointment'; }
  else if (ownerVisit && (visitClue || productClue)) {
    status = 'REVIEW_POSSIBLE_PI';
    reason = 'Possible on-site preinspection omitted from shared Calendar; confirm visit type and ownership';
  } else if (ownerVisit && customerClue) {
    status = 'REVIEW_UNCLEAR_VISIT';
    reason = 'Customer/sales visit with no confirmed preinspection intention or adequate location';
  } else if (ownerVisit && (ev.creators || []).some(function(email) {
    const id = String(email || '').toLowerCase();
    return (copies.Stephen && (id === 'classicfireplace.ca_c20qcqfhvjbv784asn9pvuiaf4@group.calendar.google.com' || id === 'stephen@classicfireplace.ca')) ||
      (copies.Jay && id === 'jay@classicfireplace.ca') ||
      (copies.Karen && id === 'karen@classicfireplace.ca');
  })) {
    status = 'REVIEW_IDENTITY_MISSING';
    reason = 'Owner-created appointment missing evidence to classify; verify identity, location and visit type, do not silently exclude';
  } else { status = 'EXCLUDED_UNVERIFIED'; reason = 'Insufficient evidence to treat as Preinspection'; }
  return {ev:ev,shared:shared,status:status,reason:reason,mirror:mirror,linked:linked};
}

/**
 * PI Calendar Review — today forward 31 days across CF Preinspects, Stephen,
 * Jay and Karen. Writes ONLY the 'PI Calendar Review' sheet; operator decisions
 * persist by event+start. No Striven calls, Calendar mutations or email.
 */
function tmv3_piCalendarReviewRefresh_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1200)) return {status:'PI_REVIEW_LOCK_BUSY',calendarWrites:0,strivenCalls:0};
  try {
    const spreadsheet = SpreadsheetApp.openById(TMV3_SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('PI Calendar Review');
    if (!sheet) throw new Error('PI_REVIEW_SHEET_NOT_FOUND_NO_WRITE');
    const timezone = 'America/Toronto';
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const until = new Date(today.getTime()); until.setDate(until.getDate() + 31);
    const checkedDate = Utilities.formatDate(today, timezone, 'yyyy-MM-dd');
    const copiesByKey = {};
    const sourceCounts = {};
    const coverageErrors = [];
    tmv3_piReviewCalendarSources_().forEach(function(source) {
      let calendar = null, events = null;
      try {
        calendar = CalendarApp.getCalendarById(source.id);
        if (calendar) events = calendar.getEvents(today, until);
      } catch (error) {
        coverageErrors.push({source:source.name,reason:'CALENDAR_ACCESS_DENIED'});
      }
      if (!calendar || !events) {
        if (!coverageErrors.some(function(x) { return x.source === source.name; }))
          coverageErrors.push({source:source.name,reason:'CALENDAR_UNAVAILABLE'});
        if (source.name === 'CF Preinspects') throw new Error('PI_PRIMARY_CALENDAR_REQUIRED_NO_WRITE');
        sourceCounts[source.name] = null;
        return;
      }
      if (events.length > 750) throw new Error('PI_CALENDAR_EVENT_LIMIT_REVIEW_NO_WRITE:' + source.name);
      sourceCounts[source.name] = events.length;
      events.forEach(function(event) {
        if (event.isAllDayEvent() && event.getStartTime() < today) return;
        const rawId = String(event.getId() || '');
        const id = rawId.replace(/@google\.com$/i, '');
        const start = event.getStartTime();
        const key = id + '|' + start.getTime();
        const record = {id:id,rawId:rawId,start:start.toISOString(),end:event.getEndTime().toISOString(),
          title:String(event.getTitle() || ''),location:String(event.getLocation() || ''),
          description:String(event.getDescription() || ''),creators:event.getCreators().map(String),
          calendarId:source.id,calendarName:source.name};
        if (!copiesByKey[key]) copiesByKey[key] = {};
        if (copiesByKey[key][source.name]) throw new Error('PI_DUPLICATE_SOURCE_EVENT_NO_WRITE:' + key);
        copiesByKey[key][source.name] = record;
      });
    });
    const old = sheet.getLastRow() > 3 ? sheet.getRange(4, 1, Math.min(sheet.getLastRow() - 3, 1500), 20).getValues() : [];
    const oldByKey = {};
    old.forEach(function(row) {
      const id = String(row[8] || '');
      const start = String(row[1] || '');
      if (!id || !start) return;
      const key = id + '|' + start;
      if (!oldByKey[key]) oldByKey[key] = {decision:String(row[18] || ''),notes:String(row[19] || ''),row:row};
    });
    const headers = [
      'Status','Start (Toronto)','End (Toronto)','Event / Customer','Calendar Copies',
      'Organizer / Origin','Location','Customer #','Google Event ID','Shared CF?',
      'Stephen Copy','Jay Copy','Karen Copy','Task Link?','Review Reason',
      'Original Notes Preview','Calendar URL','Checked Date','Operator Decision','Operator Notes'
    ];
    const rows = [];
    const norm = function(s) { return String(s || '').replace(/\s+/g, ' ').trim().toLowerCase(); };
    const stamp = function(iso) { return Utilities.formatDate(new Date(iso), timezone, "yyyy-MM-dd'T'HH:mm:ssXXX"); };
    const statusOrder = {REVIEW_MIRROR_MISMATCH:0,SHARED_UNLINKED:1,SHARED_LINKED:2,
      REVIEW_POSSIBLE_PI:3,REVIEW_UNCLEAR_VISIT:4,REVIEW_IDENTITY_MISSING:5,EXCLUDED_SERVICE:6,
      EXCLUDED_OTHER_WORKFLOW:7,EXCLUDED_NON_CUSTOMER:8,EXCLUDED_UNVERIFIED:9};
    Object.keys(copiesByKey).forEach(function(key) {
      const copies = copiesByKey[key];
      const c = tmv3_piReviewClassify_(copies, coverageErrors);
      const ev = c.ev;
      const number = ev.title.match(/(?:^|#)\s*(\d{5})\b/);
      const startText = stamp(ev.start), endText = stamp(ev.end);
      const orig = oldByKey[ev.id + '|' + startText];
      const link = 'https://www.google.com/calendar/event?eid=' + Utilities.base64EncodeWebSafe(ev.id + ' ' + ev.calendarId).replace(/=+$/, '');
      rows.push([c.status,startText,endText,ev.title,Object.keys(copies).join(', '),(ev.creators || []).join(', '),
        ev.location,number ? number[1] : '',ev.id,c.shared ? 'YES' : 'NO',c.shared ? c.mirror : copies.Stephen ? 'PRIVATE' : '—',
        copies.Jay ? 'PRESENT' : '—',copies.Karen ? 'PRESENT' : '—',c.linked ? 'YES' : 'NO',c.reason,
        norm(ev.description).slice(0, 330),link,checkedDate,orig ? orig.decision : '',orig ? orig.notes : '']);
    });
    // Inaccessible source: keep last-known rows; never report absence as fact.
    Object.keys(oldByKey).forEach(function(key) {
      const prev = oldByKey[key];
      const sourceNames = String(prev.row[4] || '');
      if (!coverageErrors.some(function(x) { return sourceNames.split(', ').indexOf(x.source) >= 0; })) return;
      if (rows.some(function(r) { return r[8] + '|' + r[1] === key; })) return;
      const preserved = prev.row.slice(0, 20);
      preserved[0] = 'SOURCE_ACCESS_UNVERIFIED';
      preserved[14] = 'Calendar source access missing; last-known event retained, no auto-promotion or Task creation';
      rows.push(preserved);
    });
    // Never erase human decisions for events that moved beyond the review horizon.
    Object.keys(oldByKey).forEach(function(key) {
      const prev = oldByKey[key];
      if (!prev.decision && !prev.notes) return;
      if (rows.some(function(row) { return row[8] + '|' + row[1] === key; })) return;
      const preserved = prev.row.slice(0, 20);
      preserved[0] = 'OUT_OF_WINDOW_REVIEW';
      preserved[14] = 'Retained due to operator decision; event not in current 31-day Calendar window';
      rows.push(preserved);
    });
    rows.sort(function(a, b) {
      const aa = typeof statusOrder[a[0]] === 'number' ? statusOrder[a[0]] : 10;
      const bb = typeof statusOrder[b[0]] === 'number' ? statusOrder[b[0]] : 10;
      return aa - bb || String(a[1]).localeCompare(String(b[1]));
    });
    const statuses = {};
    rows.forEach(function(r) { statuses[r[0]] = (statuses[r[0]] || 0) + 1; });
    const summary = 'Calendar-only review ' + checkedDate + ' | ' + sourceCounts['CF Preinspects'] + ' shared | ' +
      (statuses.REVIEW_POSSIBLE_PI || 0) + ' possible | ' + (statuses.REVIEW_UNCLEAR_VISIT || 0) + ' unclear | ' +
      (coverageErrors.length ? 'PARTIAL COVERAGE: ' + coverageErrors.map(function(x) { return x.source; }).join(', ') : 'all sources accessible') +
      ' | zero Striven writes';
    const report = [
      ['PI CALENDAR INTAKE REVIEW — READ ONLY'].concat(Array(19).fill('')),
      [summary].concat(Array(19).fill('')),
      headers
    ].concat(rows);
    const lastRow = sheet.getLastRow();
    if (report.length > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), report.length - sheet.getMaxRows());
    let changed = false;
    // A1:T1 and A2:T2 are merged. Set only the anchor cell to protect the merge.
    if (sheet.getRange(1, 1).getDisplayValue() !== report[0][0]) { sheet.getRange(1, 1).setValue(report[0][0]); changed = true; }
    if (sheet.getRange(2, 1).getDisplayValue() !== report[1][0]) { sheet.getRange(2, 1).setValue(report[1][0]); changed = true; }
    const headerNow = sheet.getRange(3, 1, 1, 20).getDisplayValues()[0];
    if (!headerNow.every(function(v, i) { return String(v || '') === String(headers[i] || ''); })) {
      sheet.getRange(3, 1, 1, 20).setValues([headers]); changed = true;
    }
    if (rows.length) {
      const bodyNow = sheet.getRange(4, 1, rows.length, 20).getDisplayValues();
      const bodyMatches = bodyNow.every(function(row, i) {
        return row.every(function(v, j) { return String(v || '') === String(rows[i][j] || ''); });
      });
      if (!bodyMatches) { sheet.getRange(4, 1, rows.length, 20).setValues(rows); changed = true; }
    }
    if (lastRow > report.length) { sheet.getRange(report.length + 1, 1, lastRow - report.length, 20).clearContent(); changed = true; }
    if (changed) SpreadsheetApp.flush();
    if (sheet.getRange(3, 1).getDisplayValue() !== 'Status' ||
        sheet.getRange(1, 1).getDisplayValue() !== 'PI CALENDAR INTAKE REVIEW — READ ONLY')
      throw new Error('PI_REVIEW_SHEET_READBACK_FAILED');
    return {status:coverageErrors.length ? 'PI_REVIEW_PARTIAL_SOURCE_ACCESS' :
        (changed ? 'PI_REVIEW_REFRESHED' : 'PI_REVIEW_UNCHANGED'),
      calendarCounts:sourceCounts,coverageErrors:coverageErrors,
      logicalEvents:rows.length,statuses:statuses,
      scope:'TODAY_FORWARD_31_DAYS',calendarWrites:0,strivenCalls:0,emailsSent:0};
  } finally { lock.releaseLock(); }
}

/** Manual entry point; writes only the review sheet. */
function tmv3_preInspectionCalendarReviewRefresh() {
  return tmv3_piCalendarReviewRefresh_();
}

/** Rebuild the PreInspection operator sheet from CF Preinspects + cached Striven sources. */
function tmv3_preInspectionRefreshOperatorView_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1500)) return {status:'PI_OPERATOR_LOCK_BUSY',strivenCalls:0};
  try {
    const cfg = TMV3.VERTICALS.PreInspection;
    const source = tmv3_verticalCalendars_('PreInspection', cfg)[0];
    const calendar = CalendarApp.getCalendarById(cfg.primaryCalendarId);
    if (!calendar) throw new Error('PI_OPERATOR_PRIMARY_CALENDAR_NOT_ACCESSIBLE');
    const start = new Date(); start.setHours(0, 0, 0, 0);
    const end = new Date(start); end.setDate(end.getDate() + Math.max(1, Number(cfg.lookaheadDays || 365)));
    const step1 = calendar.getEvents(start, end).map(function(event) {
      return tmv3_calendarEvent_('PreInspection', cfg, source, event, {applyEligibility:false});
    }).filter(function(x) { return x && !tmv3_step2PreInspectionNonCustomer_(x); });
    const step2 = tmv3_step2CalendarRecords_(step1);
    const step3 = tmv3_step3BusinessAnchorRecords_(step2, tmv3_step3AnchorIndex_());
    const step4 = tmv3_step4IdentityRecords_(step3, tmv3_step4IdentityIndex_());
    const step5 = tmv3_step5TaskRecords_(step4, tmv3_step5TaskIndex_());
    const step6 = tmv3_step6DecisionRecords_(step5).filter(function(r) { return r.vertical === 'PreInspection'; });
    step6.sort(function(a, b) { return new Date(a.start) - new Date(b.start); });
    const sheet = tmv3_sheet_(cfg.sheet), oldLast = sheet.getLastRow();
    tmv3_replaceRows_(cfg.sheet, TMV3_STEP1_HEADERS.slice(), step6.map(tmv3_step6OperatorRow_));
    tmv3_step1FormatOperatorView_(cfg.sheet, 'PreInspection', step6, oldLast);
    const actual = tmv3_rows_(cfg.sheet), seen = {};
    actual.forEach(function(row) {
      const id = String(row['Event ID'] || '');
      if (!id || seen[id]) throw new Error('PI_OPERATOR_EVENT_ID_EMPTY_OR_DUPLICATE');
      seen[id] = true;
    });
    const expected = step6.map(function(r) { return String(r.eventId || ''); });
    if (actual.length !== expected.length || expected.some(function(id) { return !seen[id]; }))
      throw new Error('PI_OPERATOR_ROW_READBACK_MISMATCH');
    const result = {status:'PI_OPERATOR_VIEW_REFRESHED',at:Utilities.formatDate(new Date(), TMV3_TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
      events:actual.length,mode:'CURRENT_SHARED_CALENDAR_PLUS_CACHED_STRIVEN_SOURCES',
      strivenCalls:0,calendarWrites:0,taskWrites:0};
    tmv3_audit_('PreInspection', '', '', 'PI_OPERATOR_QUEUE', 'PASS', JSON.stringify(result));
    return result;
  } finally { lock.releaseLock(); }
}

/** Manual entry point for the operator sheet. */
function tmv3_preInspectionRefreshOperatorView() {
  return tmv3_preInspectionRefreshOperatorView_();
}

function tmv3_preInspectionTodayRecords_(targetDate) {
  const today = targetDate || Utilities.formatDate(new Date(), TMV3_TIMEZONE, 'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(today)) throw new Error('INVALID_PI_TARGET_DATE');
  return tmv3_preInspectionIntakeRecords_().filter(function(record) {
    return record.calendarRole === 'PRIMARY_SHARED' &&
      Utilities.formatDate(new Date(record.start), TMV3_TIMEZONE, 'yyyy-MM-dd') === today;
  });
}

function tmv3_preInspectionCompactPlan_(plan) {
  return {eventId:plan.eventId,taskId:plan.taskId,taskStatus:plan.taskStatus,plan:plan.plan,blocker:plan.blocker,
    customer:plan.customerCheck,location:plan.locationCheck,requestedBy:plan.requestedByCheck,
    start:plan.startCheck,end:plan.endCheck,assignment:plan.assignmentCheck,
    calendarTitle:plan.calendarTitleCheck,taskName:plan.taskNameCheck};
}

function tmv3_piRequireReadbackHeadroom_() {
  const before = tmv3_strivenApiBudgetStatus_();
  if (before.exhausted || !Number.isFinite(before.remaining) || before.remaining < 100) throw new Error('INSUFFICIENT_READBACK_HEADROOM');
  return before;
}

/** Logs fresh plans for one day's CF Preinspects events (default today). */
function tmv3_preInspectionTodayPreview(targetDate) {
  const before = tmv3_piRequireReadbackHeadroom_();
  const records = tmv3_preInspectionTodayRecords_(targetDate);
  records.forEach(function(record) {
    const plans = tmv3_step7FreshPlansForEvent_('PreInspection', record.eventId);
    Logger.log(JSON.stringify({title:record.title,plans:plans.map(tmv3_preInspectionCompactPlan_)}));
  });
  Logger.log(JSON.stringify({status:'PREVIEW_ONLY',events:records.length,before:before.count,
    after:tmv3_strivenApiBudgetStatus_().count}));
}

function tmv3_preInspectionTodayLocationEvidence() {
  const before = tmv3_piRequireReadbackHeadroom_();
  const results = tmv3_preInspectionTodayRecords_().map(function(record) {
    return {eventId:record.eventId,title:record.title,evidence:tmv3_preInspectionOptionalOrderLocationEvidence_(record)};
  });
  Logger.log(JSON.stringify({results:results,before:before.count,afterCalls:tmv3_strivenApiBudgetStatus_().count}));
}

/** READ ONLY authorization diagnostic: Calendar, mail quota and installed slot handlers. */
function tmv3_preInspectionAuthorizationCheck() {
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/script.send_mail',
    'https://www.googleapis.com/auth/script.scriptapp'
  ]);
  const cfg = TMV3.VERTICALS.PreInspection;
  const calendarId = String(cfg && cfg.primaryCalendarId || '');
  if (!calendarId) throw new Error('PREINSPECTION_CALENDAR_ID_MISSING');
  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) throw new Error('PREINSPECTION_CALENDAR_INACCESSIBLE');
  const mailQuota = MailApp.getRemainingDailyQuota();
  const handlers = ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); })
    .filter(function(name) { return /^tmv3_preInspectionCheck(0800|1300|1700)$/.test(name); });
  const result = {
    status:'CALENDAR_MAIL_AUTHORIZATION_CHECK_PASSED',
    calendarName:calendar.getName(),
    calendarReadable:true,
    mailQuotaReadable:Number.isInteger(mailQuota),
    mailQuotaRecipientsRemaining:mailQuota,
    currentUserPreinspectionTriggerHandlers:handlers.sort(),
    calendarWrites:0,strivenCalls:0,emailsSent:0,triggerChanges:0
  };
  if (!result.mailQuotaReadable) throw new Error('MAIL_QUOTA_NOT_INTEGER');
  Logger.log(JSON.stringify(result));
  return result;
}

function tmv3_preInspectionCleanedReleaseStatus() {
  return {revision:TMV3_PI_REVISION,releaseApproved:TMV3_PI_CLEANED_RELEASE_APPROVED,
    productionWritesEnabledByFile:TMV3_PI_CLEANED_RELEASE_APPROVED,
    outstanding:['VERIFY_SHARED_STEP7_EMITS_OWNER_EMPLOYEE_REQUESTER','VERIFY_OWNER_TO_EMPLOYEE_MAPPING_KAREN',
      'RUN_FULL_V3_SHARED_DEPENDENCIES_AND_LIVE_READBACK','VERIFY_SO_STATUS_ID_TABLE_18_20_AGAINST_STRIVEN']};
}

/**
 * Read-only dry run for ONE CF Preinspects event: source, classification,
 * discovery intent, hold, owner/Requested By, notification eligibility and
 * the frozen email preview metadata. Reads the authoritative Calendar event and
 * cached sheets only — no Striven calls, no Calendar/Task/guest writes, no email,
 * no Script Property writes.
 */
function tmv3_preInspectionDryRun(eventId) {
  const id = String(eventId || '').trim();
  if (!/^[a-z0-9_]+@google\.com$/i.test(id)) throw new Error('INVALID_PREINSPECTION_EVENT_ID');
  const cfg = TMV3.VERTICALS.PreInspection;
  const source = tmv3_verticalCalendars_('PreInspection', cfg).filter(function(s) {
    return s.role === 'PRIMARY_SHARED' && String(s.calendarId) === String(cfg.primaryCalendarId);
  })[0];
  if (!source) throw new Error('PI_AUTHORITATIVE_SOURCE_NOT_UNIQUE');
  const found = tmv3_findCalendarEventById_(cfg.primaryCalendarId, id);
  const out = (!found || !found.event) ? {status:'EVENT_NOT_ON_CF_PREINSPECTS',eventId:id,sideEffects:0} :
    tmv3_piDryRunRecord_(tmv3_calendarEvent_('PreInspection', cfg, source, found.event, {applyEligibility:false}), Date.now());
  Logger.log('PI_DRY_RUN=' + JSON.stringify(out));
  return out;
}

/** Pure staged evaluation of a normalized record (see tmv3_preInspectionDryRun). */
function tmv3_piDryRunRecord_(record, nowValue) {
  const stages = [];
  function stage(name, fn) {
    let out;
    try { out = fn(); } catch (error) { out = {status:'ERROR',error:String(error && error.message || error)}; }
    stages.push({stage:name,result:out});
    return out;
  }
  stage('RELEASE_GATE', function() {
    return {status:TMV3_PI_CLEANED_RELEASE_APPROVED ? 'OPEN' : 'CLOSED_NO_BUSINESS_WRITES'};
  });
  stage('LOCKED_STANDARD', function() { return tmv3_piAssertLockedStandard_(); });
  const authoritative = stage('SOURCE', function() {
    return {status:tmv3_piIsAuthoritativeRecord_(record) ? 'AUTHORITATIVE' : 'NOT_AUTHORITATIVE',
      calendarRole:record && record.calendarRole || ''};
  });
  stage('NON_CUSTOMER', function() { return {nonCustomer:tmv3_step2PreInspectionNonCustomer_(record)}; });
  stage('STEP2_CLASSIFICATION', function() { return tmv3_step2ClassifyPreInspection_(record); });
  stage('DISCOVERY_INTENT', function() { return tmv3_preInspectionDiscoveryDecision_(record, TMV3.VERTICALS.PreInspection); });
  const hold = stage('OPERATOR_HOLD', function() {
    const h = tmv3_preInspectionApprovalHold_(record.eventId);
    return h ? {status:'HELD',mode:h.mode} : {status:'NONE'};
  });
  stage('OWNER_REQUESTED_BY', function() { return tmv3_piVerifiedOwnerEmployee_(record); });
  const assessment = stage('MISSING_DETAILS', function() { return tmv3_preInspectionMissingDetailAssessment_(record); });
  stage('NOTIFICATION_DECISION', function() {
    if (authoritative.status !== 'AUTHORITATIVE') return {status:'NONAUTHORITATIVE_SOURCE_NO_SEND',send:false};
    if (hold.status === 'ERROR') return {status:'HOLD_CHECK_FAILED_NO_SEND',send:false};
    if (assessment.status === 'ERROR') return {status:'ASSESSMENT_FAILED_NO_SEND',send:false};
    if (hold.status === 'HELD') return {status:'OPERATOR_APPROVAL_HOLD_NO_SEND',send:false};
    if (!assessment.shouldNotify) return {status:'NO_MISSING_DETAILS',send:false};
    if (!tmv3_normEmail_(assessment.creatorEmail)) return {status:'UNVERIFIED_CREATOR_NO_SEND',send:false};
    if (tmv3_piCreatorEmailExcluded_(assessment.creatorEmail)) return {status:'EXCLUDED_CREATOR_NO_SEND',send:false};
    const raw = PropertiesService.getScriptProperties().getProperty('TMV3_PI_NOTIFICATION_' + tmv3_hash_(record.eventId));
    let state = null;
    if (raw === 'SENT') return {status:'ALREADY_SENT',send:false};
    if (raw) { try { state = JSON.parse(raw); } catch (e) { return {status:'UNCERTAIN_SEND_NO_RETRY',send:false}; } }
    const decision = tmv3_preInspectionNotificationDecision_(record, state, nowValue);
    if (decision.send) {
      const email = tmv3_preInspectionMissingDetailEmailPreview_(Object.assign({}, record, {notificationFollowupNumber:decision.followupNumber}));
      decision.preview = email && {recipient:email.recipient,cc:email.cc,subject:email.subject};
    }
    decision.wouldSendIfReleased = !!decision.send;
    decision.send = false;
    return decision;
  });
  const errors = stages.filter(function(s) { return s.result && s.result.status === 'ERROR'; });
  return {status:errors.length ? 'DRY_RUN_ATTENTION' : 'DRY_RUN_COMPLETE',eventId:record && record.eventId || '',
    stages:stages,sideEffects:0,strivenCalls:0,calendarWrites:0,taskWrites:0,emailsSent:0,
    evidence:'CALENDAR_AND_CACHE_ONLY_TASK_PLAN_NOT_EVALUATED'};
}

// =====================================================================
// §13 STAGED — pure, fail-closed decision helpers from the preinspection
// branch. tmv3_piRebuildValidatePayload_ is now wired into the CREATE and
// PATCH guards (§8). The order, Location and Task helpers still need verified
// Striven adapters (default-Location evidence, complete history) before they
// may drive writes; tmv3_piRebuildDryRun_ composes them without side effects.
// =====================================================================

function tmv3_piRebuildClassifyOrder_(input) {
  input = input || {};
  if (input.hasOrderEvidence === false) return {decision:'PREINSPECTION_ORDER_OPTIONAL'};
  if (input.hasOrderEvidence !== true) return {decision:'REVIEW_ORDER_EVIDENCE_INCOMPLETE'};
  if (input.statusVerified !== true) return {decision:'REVIEW_ORDER_STATUS_UNVERIFIED'};
  if (input.status === 'APPROVED') return {decision:'ROUTE_INSTALLATION'};
  if (input.status === 'PRE_APPROVAL') return {decision:'CONTINUE_PREINSPECTION'};
  return {decision:'REVIEW_ORDER_STATUS_UNSUPPORTED'};
}

function tmv3_piRebuildResolveLocation_(input) {
  input = input || {};
  if (input.calendarAddressPresent !== true && input.calendarAddressPresent !== false)
    return {decision:'REVIEW_CALENDAR_ADDRESS_EVIDENCE'};
  if (input.calendarAddressPresent === false) {
    if (input.notificationStatus === 'UNCERTAIN') return {decision:'REVIEW_NOTIFICATION_OUTCOME_NO_RETRY'};
    if (input.notificationStatus !== 'SENT') return {decision:'QUEUE_MISSING_ADDRESS_NOTICE',fallback:'WAIT_FOR_NOTIFICATION'};
    if (input.noticeMatchesEventAndAddress !== true) return {decision:'REVIEW_NOTIFICATION_EVIDENCE'};
  }
  const id = String(input.customerId || '').trim();
  if (!/^[1-9]\d*$/.test(id) || input.customerVerified !== true || input.locationListComplete !== true || !Array.isArray(input.locations))
    return {decision:'REVIEW_LOCATION_EVIDENCE_INCOMPLETE'};
  const seen = {};
  const matches = [];
  for (const candidate of input.locations) {
    if (!candidate || !/^[1-9]\d*$/.test(String(candidate.id || '')) || seen[String(candidate.id)])
      return {decision:'REVIEW_LOCATION_RECORD_INVALID'};
    seen[String(candidate.id)] = true;
    if (String(candidate.customerId || '') !== id) return {decision:'REVIEW_LOCATION_OWNERSHIP_CONFLICT'};
    if (input.calendarAddressPresent === true) {
      if (candidate.addressMatch === true) matches.push(candidate);
      else if (candidate.addressMatch !== false) return {decision:'REVIEW_ADDRESS_MATCH_UNVERIFIED'};
    }
  }
  if (input.calendarAddressPresent === false) {
    if (input.primaryLocationVerified !== true || !/^[1-9]\d*$/.test(String(input.primaryLocationId || '')))
      return {decision:'REVIEW_PRIMARY_LOCATION_UNVERIFIED'};
    const primary = input.locations.filter(function(x){return String(x.id)===String(input.primaryLocationId);});
    if (primary.length !== 1) return {decision:'REVIEW_PRIMARY_LOCATION_NOT_OWNED'};
    return {decision:'MATCH_LOCATION',locationId:String(primary[0].id),basis:'PRIMARY_AFTER_NOTICE'};
  }
  if (matches.length === 0) return {decision:'REVIEW_NO_MATCHING_LOCATION'};
  if (matches.length === 1) return {decision:'MATCH_LOCATION',locationId:String(matches[0].id),basis:'UNIQUE_ADDRESS'};
  if (input.defaultLocationVerified !== true || !/^[1-9]\d*$/.test(String(input.defaultLocationId || '')))
    return {decision:'REVIEW_DEFAULT_LOCATION_UNVERIFIED'};
  const chosen = matches.filter(function(x){ return String(x.id) === String(input.defaultLocationId); });
  if (chosen.length !== 1) return {decision:'REVIEW_DEFAULT_NOT_MATCHING_ADDRESS'};
  return {decision:'MATCH_LOCATION',locationId:String(chosen[0].id),basis:'VERIFIED_DEFAULT_AMONG_DUPLICATES'};
}

function tmv3_piRebuildSelectTask_(input) {
  input=input||{};
  if (input.historyComplete !== true || !Array.isArray(input.tasks)) return {decision:'REVIEW_TASK_HISTORY_INCOMPLETE'};
  const ids={}, linked=[];
  for (const task of input.tasks) {
    if (!task || !/^[1-9]\d*$/.test(String(task.id||'')) || ids[String(task.id)]) return {decision:'REVIEW_TASK_HISTORY_INVALID'};
    ids[String(task.id)]=true;
    if (task.relationship === 'CONFLICT' || task.relationship === 'UNKNOWN') return {decision:'REVIEW_TASK_RELATIONSHIP'};
    if (task.relationship === 'SAME_VISIT') linked.push(task);
    else if (task.relationship !== 'OTHER_VISIT') return {decision:'REVIEW_TASK_RELATIONSHIP'};
  }
  if (linked.length>1) return {decision:'REVIEW_MULTIPLE_SAME_VISIT_TASKS'};
  if (linked.length===1) {
    const task=linked[0];
    if (task.status==='OPEN') return {decision:'REUSE_TASK',taskId:String(task.id)};
    if (task.status==='COMPLETED') return {decision:'FULFILLED_NO_RECREATE',taskId:String(task.id)};
    return {decision:'REVIEW_TASK_STATUS'};
  }
  if (input.writeOutcomeUncertain!==false) return {decision:'REVIEW_PRIOR_UNCERTAIN_CREATE'};
  return {decision:'CREATE_TASK'};
}

function tmv3_piRebuildValidatePayload_(payload,mode) {
  if (!payload || typeof payload!=='object' || Array.isArray(payload)) return {decision:'REJECT_INVALID_PAYLOAD'};
  if (mode!=='CREATE' && mode!=='PATCH') return {decision:'REJECT_INVALID_MODE'};
  const seen=[];
  function visit(value) {
    if (!value || typeof value!=='object') return '';
    if (seen.indexOf(value)!==-1) return 'REJECT_CYCLIC_PAYLOAD';
    seen.push(value);
    for (const key of Object.keys(value)) {
      const normalized=key.replace(/[^a-z0-9]/gi,'').toLowerCase();
      if (/^(infocustomfields|customfields|installnotes|field854|field852|field853|field860)$/.test(normalized)) return 'REJECT_PROTECTED_FIELDS';
      if (/^(customfieldid|fieldid)$/.test(normalized) && /^(852|853|854|860)$/.test(String(value[key]))) return 'REJECT_PROTECTED_FIELDS';
      if (mode==='CREATE' && /^(salesorder|salesorderid|soid|orderid)$/.test(normalized)) return 'REJECT_ORDER_ATTACHMENT';
      const nested=visit(value[key]);if(nested)return nested;
    }
    seen.pop();return '';
  }
  const violation=visit(payload);return {decision:violation||'ACCEPT'};
}

/** Pure connected-stage preview. No Calendar, Striven, Sheet, trigger, or email actions. */
function tmv3_piRebuildDryRun_(input) {
  input=input||{};
  const stages=[];
  function step(name,result) {stages.push({stage:name,result:result});return result;}
  const order=step('ORDER_CLASSIFICATION',tmv3_piRebuildClassifyOrder_(input.order));
  if (['PREINSPECTION_ORDER_OPTIONAL','CONTINUE_PREINSPECTION'].indexOf(order.decision)<0)
    return {status:order.decision,stages:stages,sideEffects:0};
  const location=step('LOCATION',tmv3_piRebuildResolveLocation_(input.location));
  if (location.decision!=='MATCH_LOCATION') return {status:location.decision,stages:stages,sideEffects:0};
  const task=step('TASK_LIFECYCLE',tmv3_piRebuildSelectTask_(input.task));
  if (task.decision==='FULFILLED_NO_RECREATE') return {status:task.decision,stages:stages,sideEffects:0};
  if (['CREATE_TASK','REUSE_TASK'].indexOf(task.decision)<0) return {status:task.decision,stages:stages,sideEffects:0};
  const mode=task.decision==='CREATE_TASK'?'CREATE':'PATCH';
  const valid=step('PAYLOAD_CHECK',tmv3_piRebuildValidatePayload_(input.payload,mode));
  if (valid.decision!=='ACCEPT') return {status:valid.decision,stages:stages,sideEffects:0};
  return {status:'DRY_RUN_READY_FOR_FURTHER_VERIFICATION',stages:stages,sideEffects:0};
}


// =====================================================================
// SELF TEST — synthetic, pure functions and configuration only.
// =====================================================================

/** Google Apps Script: one-click, read-only PreInspection staging validation.
 * No Calendar writes, Striven requests, notifications or trigger changes. */
function tmv3_runPreInspectionCleanedSelfTest() {
  const results=[];
  function run(name,fn){
    try {const actual=fn();results.push({check:name,status:'PASS',result:actual||''});}
    catch(e){results.push({check:name,status:'FAIL',error:String(e&&e.message||e)});}
  }
  const start='2026-11-12T13:00:00-05:00',end='2026-11-12T14:00:00-05:00';
  const customer={'Customer ID':'100',Name:'A. Customer'};
  const location={'Location ID':'200'};
  const event={vertical:'PreInspection',eventId:'pi-test@google.com',start:start,end:end,phone:'(416) 555-0101'};
  const one={'Task ID':'10','Task Type ID':105,'Customer ID':'100','Location ID':'200',
    Start:start,Due:end,Status:'Open',Name:'A. Customer - 123 King St, Toronto - (416) 555-0101'};
  run('01_STAGING_RELEASE_GATE',function(){if(TMV3_PI_CLEANED_RELEASE_APPROVED)throw new Error('STAGING_GATE_ON');return 'OFF';});
  run('02_LOCKED_STANDARD_CHECK',function(){return tmv3_piAssertLockedStandard_().status;});
  run('03_AUTHORITATIVE_SOURCE',function(){return TMV3.VERTICALS.PreInspection.primaryCalendarId?'CONFIGURED':'MISSING';});
  run('04_TASK_TYPE_AND_POOL',function(){const p=tmv3_preInspectionCreatePolicy();if(p.taskTypeId!==105||p.requiredPoolId!==8||p.attachSalesOrder!==false)throw Error('POLICY_DRIFT');return 'PASS';});
  run('05_INSPECTOR_FIELDS_UNMANAGED',function(){const p=tmv3_preInspectionCreatePolicy();if(p.field854!=='DO_NOT_MANAGE'||p.infoCustomFieldsAtCreate!=='NONE')throw Error('FIELD_DRIFT');return 'PASS';});
  run('06_NO_UNATTENDED_GUEST_SYNC',function(){if(tmv3_preInspectionGuestSyncDecision_([], 'stephen@classicfireplace.ca','stephen@classicfireplace.ca').action!=='NONE')throw Error('UNSAFE_GUEST_ACTION');return 'NO_ACTION';});
  run('07_OWNER_UNKNOWN_BLOCKED',function(){if(tmv3_piVerifiedOwnerEmployee_({organizerCalendarId:'unknown'}).status!=='REVIEW')throw Error('OWNER_NOT_BLOCKED');return 'REVIEW';});
  run('08_STRONG_TASK_MATCH',function(){const x=tmv3_preInspectionTaskDecision_(event,customer,location,[one]);if(x.status!=='MATCHED')throw Error('STRONG_MATCH_FAILED:'+x.status);return x.status;});
  run('09_UNRELATED_SINGLE_OPEN_REVIEW',function(){const t=Object.assign({},one,{'Location ID':'999',Start:'2026-11-20T13:00:00-05:00'});const x=tmv3_preInspectionTaskDecision_(event,customer,location,[t]);if(x.status!=='REVIEW')throw Error('BAD_REUSE');return x.errorCode;});
  run('10_MULTIPLE_OPEN_REVIEW',function(){const x=tmv3_preInspectionTaskDecision_(event,customer,location,[one,Object.assign({},one,{'Task ID':'11'})]);if(x.status!=='REVIEW')throw Error('DUPLICATE_RISK');return x.errorCode;});
  run('11_NO_TWELVE_HOUR_REWRITE',function(){const t=Object.assign({},one,{Start:'2026-11-12T01:00:00-05:00'});if(tmv3_step5AlignPreInspectionTaskClock_(t,event).Start!==t.Start)throw Error('CLOCK_MUTATION');return 'UNCHANGED';});
  run('12_REQUESTED_BY_CONTRACT',function(){const type=tmv3_piRequestedByContract_();if(type!=='employee')throw Error('PI_REQUESTED_BY_CONTRACT_NOT_EMPLOYEE');return type;});
  run('13_DISABLED_NOTIFICATION',function(){if(tmv3_preInspectionSendMissingDetails_(event).sent!==0)throw Error('EMAIL_SENT');return 'NO_SEND';});
  run('14_DISABLED_BUSINESS_CYCLE',function(){const x=tmv3_preInspectionAutomationCycle_({eventId:'pi-test@google.com'});if(x.writes!==0)throw Error('UNEXPECTED_WRITE');return x.status;});
  run('15_CONTACT_REQUESTED_BY_REJECTED',function(){try{tmv3_piAssertRequestedByEmployeeShape_({Type:'contact',Id:123});}catch(e){if(/EMPLOYEE_REQUIRED/.test(String(e)))return 'REJECTED';throw e;}throw Error('CONTACT_REQUESTER_ACCEPTED');});
  run('16_EMPLOYEE_REQUESTED_BY_VALID_SHAPE',function(){tmv3_piAssertRequestedByEmployeeShape_({Type:'employee',Id:15});return 'EMPLOYEE_SHAPE_VALID';});
  run('17_REQUESTED_BY_PLAN_MISSING_REJECTED',function(){try{tmv3_piAssertRequestedByPlan_({expectedRequestedByType:'employee'},{organizerCalendarId:'stephen@classicfireplace.ca'});}catch(e){if(/NOT_PROVEN/.test(String(e)))return 'REJECTED';throw e;}throw Error('MISSING_REQUESTER_ACCEPTED');});
  run('18_REQUESTED_BY_PLAN_OWNER_EMPLOYEE_VALID',function(){tmv3_piAssertRequestedByPlan_({expectedRequestedByType:'employee',expectedRequestedById:'15'},{organizerCalendarId:'stephen@classicfireplace.ca'});return 'OWNER_EMPLOYEE_VALID';});
  run('19_SHARED_CONTACT_CONTRACT_REJECTED',function(){try{tmv3_piRequestedByContractFromRule_(Object.freeze({requestedByType:'contact'}));}catch(e){if(/SHARED_STANDARD_CONFLICT/.test(String(e)))return 'REJECTED';throw e;}throw Error('CONTACT_CONTRACT_ACCEPTED');});
  run('20_CREATE_REQUIRES_EMPLOYEE',function(){const p={Type:{Id:105},Description:'',RequestedBy:{Type:'employee',Id:15}};tmv3_assertPreInspectionCreatePayloadHardRules_(p);try{tmv3_assertPreInspectionCreatePayloadHardRules_(Object.assign({},p,{RequestedBy:{Type:'contact',Id:123}}));}catch(e){if(/EMPLOYEE_REQUIRED/.test(String(e)))return 'CONTACT_REJECTED';throw e;}throw Error('CREATE_ACCEPTED_CONTACT');});
  run('21_PLAN_REJECTS_OTHER_EMPLOYEE',function(){try{tmv3_piAssertRequestedByPlan_({expectedRequestedByType:'employee',expectedRequestedById:'1'},{organizerCalendarId:'stephen@classicfireplace.ca'});}catch(e){if(/MUST_MATCH_VERIFIED_EVENT_OWNER/.test(String(e)))return 'REJECTED';throw e;}throw Error('OTHER_EMPLOYEE_ACCEPTED');});
  run('22_PLAN_REJECTS_UNKNOWN_OWNER',function(){try{tmv3_piAssertRequestedByPlan_({expectedRequestedByType:'employee',expectedRequestedById:'15'},{organizerCalendarId:'unknown@example.com'});}catch(e){if(/MAPPING_UNVERIFIED/.test(String(e)))return 'REJECTED';throw e;}throw Error('UNVERIFIED_OWNER_ACCEPTED');});
  run('23_PLAN_REJECTS_CONTACT',function(){try{tmv3_piAssertRequestedByPlan_({expectedRequestedByType:'contact',expectedRequestedById:'123'},{organizerCalendarId:'stephen@classicfireplace.ca'});}catch(e){if(/EMPLOYEE_NOT_PROVEN/.test(String(e)))return 'REJECTED';throw e;}throw Error('CONTACT_PLAN_ACCEPTED');});
  run('24_PATCH_REJECTS_CONTACT',function(){try{tmv3_piAssertSafeTaskPatch_({RequestedBy:{Type:'contact',Id:123}});}catch(e){if(/EMPLOYEE_REQUIRED/.test(String(e)))return 'REJECTED';throw e;}throw Error('CONTACT_PATCH_ACCEPTED');});

  run('25_OWNER_EMAIL_NOT_ARRAY_ORDER',function(){const x=tmv3_piVerifiedOwnerEmployee_({organizerCalendarId:'stephen@classicfireplace.ca'});if(x.status!=='MATCHED'||x.employeeId!==15)throw Error('STEPHEN_OWNER_NOT_DETERMINISTIC');return 'STEPHEN_15';});
  run('26_JAY_EMAIL_OWNED_ASSIGNMENT',function(){const x=tmv3_piVerifiedOwnerEmployee_({organizerCalendarId:'jay@classicfireplace.ca'});if(x.status!=='MATCHED'||x.employeeId!==1)throw Error('JAY_OWNER_INVALID');return 'JAY_1';});
  run('27_KAREN_NOT_GUESSED',function(){const x=tmv3_piVerifiedOwnerEmployee_({organizerCalendarId:'karen@classicfireplace.ca'});if(x.status!=='REVIEW'||x.employeeId!==null)throw Error('KAREN_WRONGLY_ASSIGNED');return x.blocker;});
  run('28_SHARED_RULE_IMMUTABLE',function(){const x=tmv3_piRequestedByContractFromRule_(Object.freeze({requestedByType:'employee'}));if(x!=='employee')throw Error('FROZEN_RULE_INVALID');return x;});
  run('29_UNKNOWN_CALENDAR_NOT_POSITIONALLY_TRUSTED',function(){const x=tmv3_piVerifiedOwnerEmployee_({organizerCalendarId:'pi-qa-never-mapped@invalid.example'});if(x.status!=='REVIEW')throw Error('UNKNOWN_CALENDAR_TRUSTED');return x.blocker;});
  run('30_EMPLOYEE_PLAN_JAY',function(){tmv3_piAssertRequestedByPlan_({expectedRequestedByType:'employee',expectedRequestedById:'1'},{organizerCalendarId:'jay@classicfireplace.ca'});return 'JAY_MATCH';});
  run('31_MULTI_CALENDAR_ORDER_NOT_TRUSTED',function(){const cfg=TMV3.VERTICALS.PreInspection||{};const arr=cfg.secondaryCalendarIds||[];if(arr.length<2)return 'NOT_APPLICABLE_SINGLE_CALENDAR';const unknown=arr.map(function(s){return String(s||'').toLowerCase();}).filter(function(s){return s&&s!=='stephen@classicfireplace.ca'&&s!=='jay@classicfireplace.ca'&&s!=='karen@classicfireplace.ca';})[0];if(!unknown)return 'NO_UNKNOWN_SECONDARY_ID';const x=tmv3_piVerifiedOwnerEmployee_({organizerCalendarId:unknown});if(x.status!=='REVIEW')throw Error('CALENDAR_ARRAY_POSITION_WAS_TRUSTED');return 'POSITION_NOT_TRUSTED';});
  // r7 additions: wired branch guards and the read-only dry run.
  const canonicalCreate=function(){return {Type:{Id:105,Name:'Pre Inspection'},Title:'A. Customer - 123 King St, Toronto - (416) 555-0101',Description:'',
    Customer:{Id:100,Number:'12345',Name:'A. Customer'},Location:{Id:200,Name:'123 King St'},RequestedBy:{Id:15,Name:'Stephen Foley',Type:'employee'}};};
  run('32_CREATE_ACCEPTS_CANONICAL_PAYLOAD',function(){tmv3_assertPreInspectionCreatePayloadHardRules_(canonicalCreate());return 'ACCEPTED';});
  run('33_CREATE_REJECTS_NESTED_CUSTOM_FIELDS',function(){const p=canonicalCreate();p.Location.CustomFields=[{Id:854,Value:'x'}];try{tmv3_assertPreInspectionCreatePayloadHardRules_(p);}catch(e){if(/PROTECTED_FIELDS/.test(String(e)))return 'REJECTED';throw e;}throw Error('NESTED_CUSTOM_FIELDS_ACCEPTED');});
  run('34_PATCH_REJECTS_NESTED_FIELD_854',function(){try{tmv3_piAssertSafeTaskPatch_({Id:10,Location:{Id:200,Extra:[{fieldId:854,value:'x'}]}});}catch(e){if(/PROTECTED_FIELDS/.test(String(e)))return 'REJECTED';throw e;}throw Error('NESTED_854_PATCH_ACCEPTED');});
  run('35_ORDER_STATUS_ID_TABLE',function(){const t=TMV3_PI_ORDER_STATUS_BY_ID;if(t['19']!=='quoted'||t['22']!=='approved'||t['25']!=='in progress')throw Error('STATUS_TABLE_DRIFT');return 'PINNED';});
  run('36_DRY_RUN_HAS_NO_SIDE_EFFECTS',function(){const x=tmv3_piDryRunRecord_(Object.assign({},event,{calendarRole:'PRIMARY_SHARED',calendarId:TMV3.VERTICALS.PreInspection.primaryCalendarId,title:'12345 - A. Customer - (416) 555-0101',location:'123 King St, Toronto',creator:'adam@classicfireplace.ca',organizerCalendarId:'stephen@classicfireplace.ca'}),Date.parse('2026-11-10T14:00:00Z'));const n=x.stages.filter(function(s){return s.stage==='NOTIFICATION_DECISION';})[0];if(x.sideEffects!==0||!n||n.result.send!==false)throw Error('DRY_RUN_UNSAFE');return x.status;});
  const failed=results.filter(function(x){return x.status==='FAIL';});
  const out={status:failed.length?'ATTENTION_UNRESOLVED_RULES':'SYNTHETIC_SELF_TEST_PASS',
    revision:TMV3_PI_REVISION,checks:results.length,passed:results.length-failed.length,
    failures:failed,results:results,calendarWrites:0,strivenCalls:0,emailsSent:0,
    evidence:'PURE_FUNCTIONS_AND_CONFIG_ONLY_NOT_LIVE_READBACK'};
  Logger.log('PI_CLEANED_SELF_TEST='+JSON.stringify(out));
  return out;
}
