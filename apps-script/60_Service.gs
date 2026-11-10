// MIGRATION STAGING ONLY. Source: reconstructed V3 3.12.21. NOT APPROVED FOR DEPLOYMENT.

// Original: 30_Calendar; Function: tmv3_serviceStep1LegacyAllowed_
function tmv3_serviceStep1LegacyAllowed_(title, isAllDay, creator) {
  const cleanTitle = tmv3_clean_(title);
  const normalizedTitle = tmv3_norm_(cleanTitle);
  const creatorText = String(creator || '').toLowerCase();

  // Original Service calendar intake rule:
  // only events actually created by @classicfireplace.ca.
  if (creatorText.indexOf('@classicfireplace.ca') === -1) return false;

  if (!cleanTitle) return false;
  if (isAllDay) return false;

  if (
    normalizedTitle === 'chris' ||
    normalizedTitle === 'travis' ||
    normalizedTitle === 'matt' ||
    normalizedTitle === 'matthew' ||
    normalizedTitle === 'matthew thompson'
  ) {
    return false;
  }

  if (
    normalizedTitle === 'classic fireplace' ||
    normalizedTitle === 'classic fireplace bbq' ||
    normalizedTitle === 'classic fireplace and bbq'
  ) {
    return false;
  }

  if (
    normalizedTitle.indexOf('take van home') !== -1 ||
    normalizedTitle.indexOf('van home') !== -1 ||
    normalizedTitle.indexOf('day off') !== -1 ||
    normalizedTitle.indexOf('vacation') !== -1 ||
    normalizedTitle.indexOf('sick') !== -1 ||
    normalizedTitle.indexOf('off') === normalizedTitle.length - 3
  ) {
    return false;
  }

  return true;
}

// Original: 30_Calendar; Function: tmv3_serviceExternalCreatorRecoveryAllowed_
function tmv3_serviceExternalCreatorRecoveryAllowed_(evidence) {
  evidence = evidence || {};

  if (evidence.isAllDay) return false;

  const title = tmv3_norm_(evidence.title || '');

  if (
    !title ||
    title === 'chris' ||
    title === 'travis' ||
    title === 'matt' ||
    title === 'matthew' ||
    title === 'matthew thompson' ||
    title === 'classic fireplace' ||
    title === 'classic fireplace bbq' ||
    title === 'classic fireplace and bbq' ||
    /\b(day off|vacation|sick|take van home|van home|do not book|meeting|lunch|training|warehouse|pickup|pick up)\b/i.test(title)
  ) {
    return false;
  }

  // A managed Striven link or labelled Task number is strong Calendar-carried
  // evidence. Otherwise require SO/WO + phone + physical location together.
  if (evidence.existingTaskId || evidence.existingOrderId || evidence.taskNumber) {
    return true;
  }

  return !!(
    evidence.orderNumber &&
    evidence.phone &&
    tmv3_clean_(evidence.location)
  );
}

// Original: 30_Calendar; Function: tmv3_extractLegacyServiceOrderNumber_
function tmv3_extractLegacyServiceOrderNumber_(text) {
  const clean = tmv3_clean_(text);
  if (!clean) return '';

  const patterns = [
    /\bSO\s*#?\s*(\d{5,8})\b/i,
    /\bS\/O\s*#?\s*(\d{5,8})\b/i,
    /\bSales\s*Order\s*#?\s*(\d{5,8})\b/i,
    /\bOrder\s*#?\s*(\d{5,8})\b/i,
    /\b(\d{6})\b/
  ];

  for (let i = 0; i < patterns.length; i++) {
    const match = clean.match(patterns[i]);
    if (match && match[1]) return match[1];
  }

  return '';
}

// Original: 30_Calendar; Function: tmv3_serviceLegacyCustomerName_
function tmv3_serviceLegacyCustomerName_(title) {
  const clean = tmv3_clean_(title);
  if (!clean) return '';

  const value = clean
    .replace(/\bSO\s*#?\s*\d{5,8}\b/gi, '')
    .replace(/\bSales\s*Order\s*#?\s*\d{5,8}\b/gi, '')
    .replace(/\bOrder\s*#?\s*\d{5,8}\b/gi, '')
    .replace(/\bTask\s*#?\s*\d{4,8}\b/gi, '')
    .replace(/\bService\b/gi, '')
    .replace(/\bChris\b/gi, '')
    .replace(/\bTravis\b/gi, '')
    .replace(/\bMatt\b/gi, '')
    .replace(/\bMatthew\b/gi, '')
    .replace(/[|:;,#\-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  return value.length >= 3 ? value : '';
}

// Original: 34_Step2_Eligibility; Function: tmv3_step2ClassifyService_
function tmv3_step2ClassifyService_(record) {
  if (record.serviceLegacyAllowed === true) {
    return tmv3_step2Decision_(
      'ELIGIBLE',
      'SERVICE_LEGACY_ELIGIBLE',
      'Service event passed the original creator and blocker rules.',
      false
    );
  }

  if (record.serviceExternalRecovery === true) {
    return tmv3_step2Decision_(
      'ELIGIBLE',
      'SERVICE_EXTERNAL_CREATOR_RECOVERY',
      'External creator recovered because Calendar already carries strong Striven task/order evidence.',
      false,
      ['External creator — verify against Striven in the business-anchor stage.']
    );
  }

  return tmv3_step2Decision_(
    'SKIP',
    'SERVICE_CREATOR_OR_BLOCKER_REJECTED',
    'Service event did not pass the legacy creator/blocker rules and had no strong recovery evidence.',
    false
  );
}

// Original: 38_Step4_Identity; Function: tmv3_step4ResolveServiceIdentity_
function tmv3_step4ResolveServiceIdentity_(record, refs, anchor) {
  const evidence = [];
  const warnings = [];

  let customer = null;
  let location = null;

  const anchorCustomerId = tmv3_clean_(anchor.customerId);

  if (anchorCustomerId && refs.customerById[anchorCustomerId]) {
    customer = refs.customerById[anchorCustomerId];
    evidence.push('SERVICE_CUSTOMER_FROM_WORK_ORDER');
  }

  const orderLocationId = tmv3_clean_(anchor.orderLocationIdEvidence);

  if (!customer && orderLocationId && refs.locationById[orderLocationId]) {
    const candidateLocation = refs.locationById[orderLocationId];
    const ownerId = tmv3_clean_(candidateLocation['Customer ID']);

    if (ownerId && refs.customerById[ownerId]) {
      customer = refs.customerById[ownerId];
      location = candidateLocation;
      evidence.push('SERVICE_CUSTOMER_FROM_WORK_ORDER_LOCATION');
      evidence.push('CUSTOMER_LOCATION_FROM_WORK_ORDER');
    }
  }

  if (!customer) {
    const recovery = tmv3_step4ResolveCustomerFromCalendarEvidence_(
      record,
      refs
    );

    if (recovery.status !== 'MATCHED') {
      return tmv3_step4Decision_(
        recovery.status === 'BLOCKED' ? 'BLOCKED' : 'REVIEW',
        recovery.errorCode,
        recovery.reason,
        recovery.customer || null,
        recovery.location || null,
        null,
        'NOT_RESOLVED',
        evidence.concat(recovery.evidence || []),
        warnings
      );
    }

    customer = recovery.customer;
    location = recovery.location || null;
    Array.prototype.push.apply(evidence, recovery.evidence || []);
  }

  if (!location) {
    const locationResult = tmv3_step4ResolveLocation_(
      record,
      refs,
      customer,
      orderLocationId
    );

    if (
      locationResult.status !== 'MATCHED' &&
      locationResult.status !== 'CREATE_REQUIRED'
    ) {
      return tmv3_step4Decision_(
        locationResult.status === 'BLOCKED' ? 'BLOCKED' : 'REVIEW',
        locationResult.errorCode,
        locationResult.reason,
        customer,
        null,
        null,
        'NOT_RESOLVED',
        evidence.concat(locationResult.evidence || []),
        warnings
      );
    }

    if (locationResult.status === 'CREATE_REQUIRED') {
      warnings.push('Location will need to be created on the verified Customer before any Task write.');
    }

    location = locationResult.location || null;
    Array.prototype.push.apply(evidence, locationResult.evidence || []);
  }

  const contactResult = tmv3_step4ResolveContact_(record, customer, '');

  if (contactResult.status === 'REVIEW') {
    return tmv3_step4Decision_(
      'REVIEW',
      contactResult.errorCode,
      contactResult.reason,
      customer,
      location,
      null,
      'REVIEW',
      evidence.concat(contactResult.evidence || []),
      warnings
    );
  }

  if (contactResult.status === 'NO_MATCH') {
    warnings.push('No deterministic Contact match; Customer and Location remain verified.');
  }

  return tmv3_step4Decision_(
    'VERIFIED',
    'SERVICE_IDENTITY_VERIFIED',
    'Service Customer and Customer-owned Location are verified from Work Order + Calendar evidence.',
    customer,
    location,
    contactResult.contact || null,
    contactResult.status,
    evidence.concat(contactResult.evidence || []),
    warnings
  );
}

// Original: 40_Engine; Function: tmv3_extractServiceFireplaceNumber_
function tmv3_extractServiceFireplaceNumber_(taskName) {
  const clean = tmv3_clean_(taskName);
  if (!clean) return 0;

  const patterns = [
    /\bFP\s*#?\s*(\d+)\b/i,
    /\bF\.?P\.?\s*#?\s*(\d+)\b/i,
    /\bFireplace\s*#?\s*(\d+)\b/i
  ];

  for (let i = 0; i < patterns.length; i++) {
    const match = clean.match(patterns[i]);
    if (match && match[1]) {
      const value = Number(match[1]);
      if (!isNaN(value) && value > 0) return value;
    }
  }

  return 0;
}

// Original: 50_Verticals; Function: tmv3_servicePolicy
function tmv3_servicePolicy() {
  return {
    orderRequired: true,
    operationalName: 'Work Order',
    sameUnderlyingRecordAsSalesOrder: true,
    calendarLinks: [
      'WORK_ORDER',
      'TASK'
    ]
  };
}

// Original: 58_Operations; Function: tmv3_serviceTaskTypeForOrder_
function tmv3_serviceTaskTypeForOrder_(orderId) {
  const rows = tmv3_reportRows_(TMV3.PROPERTIES.SERVICE_WORK_ORDERS);
  const target = String(orderId || '');

  for (let i = 0; i < rows.length; i++) {
    const id = tmv3_clean_(tmv3_first_(rows[i], [
      'WorkOrderId','Work Order ID','WorkOrder ID','SalesOrderId','OrderId','Id'
    ]));
    if (id !== target) continue;
    return tmv3_clean_(tmv3_first_(rows[i], [
      'CategoryName','Category Name','Category'
    ]));
  }

  return '';
}

// Original: 60_Automation; Function: tmv3_runServiceAutomation
function tmv3_runServiceAutomation() {const r=tmv3_runFullVerticalSweep_('Service');Logger.log(JSON.stringify(r));return r;}

// Original: 60_Automation; Function: tmv3_serviceCheck0800
function tmv3_serviceCheck0800() { return tmv3_standardVerticalSlot_('Service','08:00'); }

// Original: 60_Automation; Function: tmv3_serviceCheck1300
function tmv3_serviceCheck1300() { return tmv3_standardVerticalSlot_('Service','13:00'); }

// Original: 60_Automation; Function: tmv3_serviceCheck1700
function tmv3_serviceCheck1700() { return tmv3_standardVerticalSlot_('Service','17:00'); }

// Original: 60_Automation; Function: tmv3_serviceSweepContinue
function tmv3_serviceSweepContinue(){return {status:'RETIRED_OUTSIDE_STANDARD_SLOTS',writes:0};}