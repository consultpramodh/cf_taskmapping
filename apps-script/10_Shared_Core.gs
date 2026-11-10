// MIGRATION STAGING ONLY. Source: reconstructed V3 3.12.21. NOT APPROVED FOR DEPLOYMENT.

// Original: 10_Core
function tmv3_ss_() {
  return SpreadsheetApp.openById(TMV3_SPREADSHEET_ID);
}

function tmv3_sheet_(name) {
  const sh = tmv3_ss_().getSheetByName(name);
  if (!sh) throw new Error('Missing required V3 sheet: ' + name);
  return sh;
}

function tmv3_clean_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}

function tmv3_norm_(value) {
  return tmv3_clean_(value).toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function tmv3_digits_(value) {
  return tmv3_clean_(value).replace(/\D+/g, '');
}

function tmv3_phone10_(value) {
  const d = tmv3_digits_(value);
  return d.length >= 10 ? d.slice(-10) : '';
}

function tmv3_now_() {
  return Utilities.formatDate(new Date(), 'America/Toronto', "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function tmv3_date_(date) {
  return Utilities.formatDate(date, 'America/Toronto', 'yyyy-MM-dd');
}

function tmv3_time_(date) {
  return Utilities.formatDate(date, 'America/Toronto', 'HH:mm');
}

function tmv3_iso_(date) {
  return Utilities.formatDate(date, 'America/Toronto', "yyyy-MM-dd'T'HH:mm:ssXXX");
}

// External Striven Task date write contract.
// Striven's Task PATCH/POST parser requires an explicit 12-hour clock marker;
// V3 keeps ISO internally and converts only at the API boundary.
function tmv3_strivenTaskDateTime_(value) {
  const date = tmv3_parseDateTime_(value);
  if (!date) {
    throw new Error('Invalid Striven Task date/time value: ' + tmv3_clean_(value));
  }
  return Utilities.formatDate(date, TMV3_TIMEZONE, 'MM/dd/yyyy hh:mm:ss a');
}

function tmv3_parseDateTime_(value) {
  if (!value) return null;
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }

  const text = tmv3_clean_(value);
  if (!text) return null;

  // Explicit zone/offset: preserve the actual instant.
  if (/Z$/i.test(text) || /[+-]\d{2}:?\d{2}$/.test(text)) {
    const zoned = new Date(text);
    return isNaN(zoned.getTime()) ? null : zoned;
  }

  const patterns = [
    "yyyy-MM-dd'T'HH:mm:ss",
    "yyyy-MM-dd'T'HH:mm",
    'MM/dd/yyyy hh:mm a',
    'M/d/yyyy h:mm a',
    'MM/dd/yyyy HH:mm',
    'M/d/yyyy H:mm'
  ];

  for (let i = 0; i < patterns.length; i++) {
    try {
      const parsed = Utilities.parseDate(text, TMV3_TIMEZONE, patterns[i]);
      if (parsed && !isNaN(parsed.getTime())) return parsed;
    } catch (ignored) {}
  }

  const fallback = new Date(text);
  return isNaN(fallback.getTime()) ? null : fallback;
}


function tmv3_sameCalendarDate_(a, b) {
  const left = tmv3_parseDateTime_(a);
  const right = tmv3_parseDateTime_(b);
  if (!left || !right) return false;

  return (
    Utilities.formatDate(left, TMV3_TIMEZONE, 'yyyy-MM-dd') ===
    Utilities.formatDate(right, TMV3_TIMEZONE, 'yyyy-MM-dd')
  );
}

function tmv3_taskDueDateOnly_(vertical, task) {
  return (
    tmv3_clean_(vertical) === 'Install' &&
    Number(task && task['Task Type ID'] || 0) === 92
  );
}

function tmv3_hash_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    tmv3_clean_(value),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) {
    return ('0' + ((b + 256) % 256).toString(16)).slice(-2);
  }).join('');
}

function tmv3_first_(obj, aliases) {
  if (!obj) return '';
  const keyMap = {};
  Object.keys(obj).forEach(function(k) { keyMap[tmv3_norm_(k)] = k; });

  for (let i = 0; i < aliases.length; i++) {
    const actual = keyMap[tmv3_norm_(aliases[i])];
    if (actual !== undefined) {
      const value = obj[actual];
      if (value !== null && value !== undefined && tmv3_clean_(value) !== '') return value;
    }
  }
  return '';
}

function tmv3_unique_(values) {
  const seen = {};
  const out = [];
  (values || []).forEach(function(v) {
    const s = tmv3_clean_(v);
    if (s && !seen[s]) {
      seen[s] = true;
      out.push(s);
    }
  });
  return out;
}

function tmv3_property_(aliases, required) {
  const props = PropertiesService.getScriptProperties();
  for (let i = 0; i < aliases.length; i++) {
    const value = tmv3_clean_(props.getProperty(aliases[i]));
    if (value) return value;
  }
  if (required) {
    throw new Error('Missing required Script Property capability: ' + aliases.join(' | '));
  }
  return '';
}

function tmv3_configCapabilities_() {
  const out = {};
  Object.keys(TMV3.PROPERTIES).forEach(function(k) {
    const aliases = TMV3.PROPERTIES[k];
    out[k] = { present: !!tmv3_property_(aliases, false), aliases: aliases.slice() };
  });
  return out;
}

function tmv3_rows_(sheetName) {
  const sh = tmv3_sheet_(sheetName);
  if (sh.getLastRow() < 2) return [];

  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const headers = values[0].map(tmv3_clean_);

  return values.slice(1)
    .filter(function(row) {
      return row.some(function(v) { return tmv3_clean_(v) !== ''; });
    })
    .map(function(row) {
      const obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    });
}

function tmv3_replaceRows_(sheetName, headers, rows) {
  const sh = tmv3_sheet_(sheetName);
  const width = headers.length;
  const needed = Math.max(2, rows.length + 1);

  if (sh.getMaxRows() < needed) {
    sh.insertRowsAfter(sh.getMaxRows(), needed - sh.getMaxRows());
  }

  if (sh.getMaxColumns() < width) {
    sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  }

  sh.clearContents();
  sh.getRange(1, 1, 1, width).setValues([headers]);

  if (rows.length) {
    sh.getRange(2, 1, rows.length, width).setValues(rows);
  }

  sh.setFrozenRows(1);
  return rows.length;
}

function tmv3_audit_(vertical, eventId, taskId, action, result, detail) {
  const sh = tmv3_sheet_(TMV3.SHEETS.AUDIT);
  sh.appendRow([
    tmv3_now_(),
    vertical || '',
    eventId || '',
    taskId || '',
    action || '',
    result || '',
    detail || '',
    TMV3.VERSION,
    TMV3.MODE
  ]);
}

function tmv3_stateKey_(vertical, eventId, taskId) {
  return [
    tmv3_clean_(vertical),
    tmv3_clean_(eventId),
    tmv3_clean_(taskId)
  ].join('|');
}

function tmv3_eventStateIndex_() {
  const rows = tmv3_rows_(TMV3.SHEETS.STATE);
  const byKey = {};

  rows.forEach(function(r) {
    const vertical = tmv3_clean_(r['Vertical']);
    const eventId = tmv3_clean_(r['Event ID']);
    const taskId = tmv3_clean_(r['Task ID']);
    const exactKey = tmv3_stateKey_(vertical, eventId, taskId);
    const eventKey = vertical + '|' + eventId;

    byKey[exactKey] = r;

    // Event-level alias is retained for classification overrides and
    // acknowledgement even when Service expands to multiple fireplace Tasks.
    if (!byKey[eventKey]) byKey[eventKey] = r;
  });

  return byKey;
}


function tmv3_titleMigrationStateIndex_() {
  if (tmv3_titleMigrationStateIndex_._cache) {
    return tmv3_titleMigrationStateIndex_._cache;
  }

  const out = {};
  tmv3_rows_(TMV3.SHEETS.STATE).forEach(function(row) {
    const vertical = tmv3_clean_(row['Vertical']);
    const eventId = tmv3_clean_(row['Event ID']);
    if (!vertical || !eventId) return;

    const key = vertical + '|' + eventId;
    if (!out[key]) out[key] = row;
  });

  tmv3_titleMigrationStateIndex_._cache = out;
  return out;
}

function tmv3_titleMigrationStateForEvent_(vertical, eventId, calendarId) {
  const row =
    tmv3_titleMigrationStateIndex_()[
      tmv3_clean_(vertical) + '|' + tmv3_clean_(eventId)
    ] || {};

  let map = {};
  const rawMap = tmv3_clean_(row['Title Migration Map']);

  if (rawMap) {
    try {
      const parsed = JSON.parse(rawMap);
      if (parsed && typeof parsed === 'object') map = parsed;
    } catch (ignored) {}
  }

  const id = tmv3_clean_(calendarId);

  return {
    preservedTitle:
      id && map[id]
        ? tmv3_clean_(map[id])
        : '',
    canonicalTitle:tmv3_clean_(row['Canonical Calendar Title']),
    normalizedAt:tmv3_clean_(row['Title Normalized At']),
    migrationMap:map
  };
}

function tmv3_recordTitleMigrationState_(
  vertical,
  eventId,
  calendarId,
  preservedTitle,
  canonicalTitle
) {
  const cleanVertical = tmv3_clean_(vertical);
  const cleanEventId = tmv3_clean_(eventId);
  const cleanCalendarId = tmv3_clean_(calendarId);
  const oldTitle = tmv3_clean_(preservedTitle);
  const newTitle = tmv3_clean_(canonicalTitle);

  if (
    !cleanVertical ||
    !cleanEventId ||
    !cleanCalendarId ||
    !oldTitle ||
    !newTitle
  ) {
    throw new Error(
      'Title migration state requires vertical, Event ID, Calendar ID, old title and canonical title.'
    );
  }

  const sh = tmv3_sheet_(TMV3.SHEETS.STATE);

  if (sh.getLastRow() < 1) {
    sh.getRange(1,1,1,6).setValues([[
      'Vertical',
      'Event ID',
      'Title Migration Map',
      'Canonical Calendar Title',
      'Title Normalized At',
      'Engine Version'
    ]]);
  }

  let headers = sh
    .getRange(1,1,1,Math.max(1,sh.getLastColumn()))
    .getValues()[0]
    .map(tmv3_clean_);

  [
    'Title Migration Map',
    'Canonical Calendar Title',
    'Title Normalized At'
  ].forEach(function(header) {
    if (headers.indexOf(header) !== -1) return;
    const col = headers.length + 1;
    if (sh.getMaxColumns() < col) {
      sh.insertColumnsAfter(sh.getMaxColumns(), col - sh.getMaxColumns());
    }
    sh.getRange(1,col).setValue(header);
    headers.push(header);
  });

  const verticalIx = headers.indexOf('Vertical');
  const eventIx = headers.indexOf('Event ID');
  const mapIx = headers.indexOf('Title Migration Map');
  const canonicalIx = headers.indexOf('Canonical Calendar Title');
  const normalizedAtIx = headers.indexOf('Title Normalized At');

  if (verticalIx < 0 || eventIx < 0) {
    throw new Error(
      'TM State is missing Vertical/Event ID columns required for title migration state.'
    );
  }

  const lastRow = sh.getLastRow();
  const values =
    lastRow >= 2
      ? sh.getRange(2,1,lastRow-1,headers.length).getValues()
      : [];

  const targets = [];

  values.forEach(function(row, index) {
    if (
      tmv3_clean_(row[verticalIx]) === cleanVertical &&
      tmv3_clean_(row[eventIx]) === cleanEventId
    ) {
      targets.push(index + 2);
    }
  });

  if (!targets.length) {
    const row = headers.map(function() { return ''; });
    row[verticalIx] = cleanVertical;
    row[eventIx] = cleanEventId;
    targets.push(sh.getLastRow() + 1);
    sh.getRange(targets[0],1,1,headers.length).setValues([row]);
  }

  targets.forEach(function(rowNumber) {
    const currentRaw = tmv3_clean_(
      sh.getRange(rowNumber,mapIx+1).getValue()
    );

    let map = {};
    if (currentRaw) {
      try {
        const parsed = JSON.parse(currentRaw);
        if (parsed && typeof parsed === 'object') map = parsed;
      } catch (ignored) {}
    }

    // The first sales-rep title is permanent history. Later normalization or
    // description-only parity repairs must not replace it with the new title.
    if (!tmv3_clean_(map[cleanCalendarId])) map[cleanCalendarId] = oldTitle;

    sh.getRange(rowNumber,mapIx+1).setValue(JSON.stringify(map));
    sh.getRange(rowNumber,canonicalIx+1).setValue(newTitle);
    sh.getRange(rowNumber,normalizedAtIx+1).setValue(tmv3_now_());
  });

  tmv3_titleMigrationStateIndex_._cache = null;

  return {
    status:'TITLE_MIGRATION_STATE_RECORDED',
    vertical:cleanVertical,
    eventId:cleanEventId,
    calendarId:cleanCalendarId,
    preservedTitle:oldTitle,
    canonicalTitle:newTitle,
    rowsUpdated:targets.length
  };
}

function tmv3_upsertState_(records) {
  const headers = [
    'Vertical','Event ID','Calendar ID','Customer ID','Location ID','Contact ID',
    'Order ID','Task ID','Source Fingerprint','Calendar Updated At','Last Verified At',
    'Engine Version','Classification Override','State','Next Action','Error Code',
    'First Seen At','Last Seen At','Last State Change At','Resolved At',
    'First Detected At','Last Attempt At','Attempt Count',
    'Acknowledged By','Acknowledged At','Acknowledgement Note','Last Error Class',
    'Title Migration Map','Canonical Calendar Title','Title Normalized At'
  ];

  const existingIndex = tmv3_eventStateIndex_();
  const rowsByExactKey = {};

  // Preserve every existing physical row, including multiple Tasks for one Event.
  tmv3_rows_(TMV3.SHEETS.STATE).forEach(function(r) {
    rowsByExactKey[
      tmv3_stateKey_(r['Vertical'], r['Event ID'], r['Task ID'])
    ] = r;
  });

  (records || []).forEach(function(r) {
    const exactKey = tmv3_stateKey_(r.vertical, r.eventId, r.taskId);
    const eventKey = r.vertical + '|' + r.eventId;
    const prior =
      rowsByExactKey[exactKey] ||
      existingIndex[exactKey] ||
      existingIndex[eventKey] ||
      {};

    const status = r.status || prior['State'] || '';
    const unresolved =
      ['MATCHED','IGNORED'].indexOf(tmv3_clean_(status).toUpperCase()) === -1;

    const now = tmv3_now_();
    const firstDetected =
      unresolved
        ? (prior['First Detected At'] || now)
        : '';

    const attemptCount =
      unresolved
        ? Number(prior['Attempt Count'] || 0) + 1
        : 0;

    const priorState = tmv3_clean_(prior['State']);
    const stateChanged = priorState !== tmv3_clean_(status);
    const resolvedAt =
      !unresolved && priorState && ['MATCHED','IGNORED'].indexOf(priorState.toUpperCase()) === -1
        ? now
        : (!unresolved ? (prior['Resolved At'] || '') : '');

    rowsByExactKey[exactKey] = {
      'Vertical': r.vertical,
      'Event ID': r.eventId,
      'Calendar ID': r.calendarId || prior['Calendar ID'] || '',
      'Customer ID': r.customerId || prior['Customer ID'] || '',
      'Location ID': r.locationId || prior['Location ID'] || '',
      'Contact ID': r.contactId || prior['Contact ID'] || '',
      'Order ID': r.orderId || prior['Order ID'] || '',
      'Task ID': r.taskId || prior['Task ID'] || '',
      'Source Fingerprint': r.fingerprint || prior['Source Fingerprint'] || '',
      'Calendar Updated At': r.calendarUpdatedAt || prior['Calendar Updated At'] || '',
      'Last Verified At': r.lastVerified || prior['Last Verified At'] || '',
      'Engine Version': TMV3.VERSION,
      'Classification Override': prior['Classification Override'] || '',
      'State': status,
      'Next Action': r.nextAction || prior['Next Action'] || '',
      'Error Code': r.errorCode || '',
      'First Seen At': prior['First Seen At'] || now,
      'Last Seen At': now,
      'Last State Change At': stateChanged ? now : (prior['Last State Change At'] || now),
      'Resolved At': resolvedAt,
      'First Detected At': firstDetected,
      'Last Attempt At': now,
      'Attempt Count': attemptCount,
      'Acknowledged By': unresolved ? (prior['Acknowledged By'] || '') : '',
      'Acknowledged At': unresolved ? (prior['Acknowledged At'] || '') : '',
      'Acknowledgement Note': unresolved ? (prior['Acknowledgement Note'] || '') : '',
      'Last Error Class': unresolved ? (r.errorCode || prior['Last Error Class'] || '') : '',
      'Title Migration Map': r.titleMigrationMap || prior['Title Migration Map'] || '',
      'Canonical Calendar Title': r.canonicalCalendarTitle || prior['Canonical Calendar Title'] || '',
      'Title Normalized At': r.titleNormalizedAt || prior['Title Normalized At'] || ''
    };
  });

  const rows = Object.keys(rowsByExactKey).sort().map(function(k) {
    const r = rowsByExactKey[k];
    return headers.map(function(h) {
      return r[h] === undefined ? '' : r[h];
    });
  });

  tmv3_replaceRows_(TMV3.SHEETS.STATE, headers, rows);
}

function tmv3_assertHardRules_() {
  const errors = [];
  const pre = TMV3.VERTICALS.PreInspection;
  const rules = TMV3_HARD_RULES.PreInspection;
  const globalRules = TMV3_HARD_RULES.GLOBAL;

  function requireRule(ok, message) {
    if (!ok) errors.push(message);
  }

  requireRule(globalRules.titleDelimiter===' - ', 'Invalid title separator standard.');
  requireRule(globalRules.executionGate==='CANONICAL_STAGE7_FRESH_EVIDENCE_READBACK', 'Stage 7 contract drift.');
  requireRule(TMV3_HARD_RULES.Install.presentation.calendarTitleParts.join('|')==='orderNumber|customerName|phone', 'Install title parts drift.');
  requireRule(TMV3_HARD_RULES.PreInspection.presentation.calendarTitleParts.join('|')==='customerNumber|customerName|phone', 'PreInspection Calendar title parts drift.');
  requireRule(TMV3_HARD_RULES.PreInspection.presentation.taskTitleParts.join('|')==='customerName|address|phone', 'PreInspection Task title parts drift.');
  requireRule(TMV3_HARD_RULES.PreInspection.presentation.linkHeading==='-----Striven Links-----', 'PreInspection link format drift.');
  requireRule(TMV3_HARD_RULES.Install.presentation.linkHeading==='------- Install Striven Task Link -------', 'Install link format drift.');
  requireRule(TMV3_HARD_RULES.Delivery.presentation.linkHeading==='------- Delivery Striven Links -------', 'Delivery link format drift.');
  requireRule(TMV3_HARD_RULES.Service.presentation.linkHeading==='------- Striven Links -------', 'Service link format drift.');
  requireRule(globalRules.taskLifecyclePolicy === 'SHARED_STAGE6_HISTORY_GATE', 'All verticals must use the shared Stage 6 Task history gate.');
  requireRule(globalRules.singleOpenDecision === 'MATCH_EXISTING', 'Exactly one valid OPEN Task must resolve to MATCH_EXISTING.');
  requireRule(globalRules.multipleOpenDecision === 'REVIEW', 'Multiple OPEN Tasks must require REVIEW unless the explicit Service multi-fireplace exception applies.');
  requireRule(globalRules.serviceMultiOpenException === 'UNIQUE_FIREPLACE_MARKERS_ONLY', 'Service multi-task exception must require unique fireplace markers.');
  requireRule(globalRules.noHistoryDecision === 'CREATE_TASK', 'No Task/history must resolve to CREATE_TASK after prerequisites are verified.');
  requireRule(globalRules.sameDayCompletedDecision === 'FULFILLED_NO_RECREATE', 'Same-day fulfilled history must suppress recreate.');
  requireRule(globalRules.anyCancelledHistoryDecision === 'REVIEW', 'Cancelled Task history must never authorize automatic RECREATE.');
  requireRule(globalRules.unprovenHistoryDateDecision === 'REVIEW', 'Unproven Task history dates must require REVIEW.');
  requireRule(globalRules.futureHistoryDecision === 'REVIEW', 'Future Task history must require REVIEW.');
  requireRule(globalRules.olderFulfilledHistoryDecision === 'RECREATE_TASK', 'Only older fulfilled history may resolve to RECREATE_TASK.');
  requireRule(globalRules.wrongVerticalTaskDecision === 'EVIDENCE_ONLY_NO_EXECUTABLE_MATCH', 'Wrong-vertical Tasks must never become executable matches.');
  requireRule(globalRules.missingPrerequisiteDecision === 'BLOCK_OR_REVIEW_BEFORE_TASK_WRITE', 'Missing required Customer/Location/Order prerequisites must block before Task writes.');
  requireRule(globalRules.missingLinkedTaskCacheDecision === 'REVIEW_FAIL_CLOSED', 'Unverified linked Tasks must fail closed to REVIEW.');

  requireRule(pre.orderRequired === false, 'PreInspection must not require a Sales Order.');
  requireRule(pre.attachOrderToTask === false, 'PreInspection must not attach a Sales Order to the Task.');
  requireRule(Number(pre.taskTypeId || 0) === 105, 'PreInspection Task Type must remain 105.');
  requireRule(pre.descriptionMustBeBlankAtCreate === true, 'PreInspection Task Description must be blank at CREATE.');
  requireRule(pre.calendarNotesPolicy === 'CALENDAR_ONLY', 'PreInspection Calendar notes must remain Calendar-only.');
  requireRule(pre.previousCalendarTitlePolicy === 'PRESERVE_AS_FIRST_DESCRIPTION_LINE_NO_HEADER', 'PreInspection previous Calendar title must be preserved as the first Description line with no added heading.');
  requireRule(pre.descriptionHeaderPolicy === 'REMOVE_NOTES_AND_SALES_ORDER_LABELS_PRESERVE_CONTENT', 'PreInspection Notes:/Sales Order: presentation headers must be removed while preserving content.');
  requireRule(pre.calendarCopyParityPolicy === 'FULL_BUSINESS_VISIBLE_PARITY_REQUIRED', 'PreInspection required Calendar copies must converge to full business-visible parity.');
  requireRule(pre.calendarCopyParityFields === 'TITLE_DESCRIPTION_LOCATION_START_END_REQUIRED_PARTICIPATION', 'PreInspection parity fields must include title, description, location, start/end and required participation.');
  requireRule(pre.calendarLinksPreservationPolicy === 'PRESERVE_EXISTING_CLICKABLE_BLOCK_REPAIR_NONCLICKABLE_OR_MISSING_ONLY', 'PreInspection Striven Links block must preserve existing clickable hyperlinks and repair only missing or non-clickable links.');
  requireRule(pre.calendarTaskLinkLabelPolicy === 'TASK_ID_PLUS_ACTUAL_STRIVEN_TASK_NAME', 'PreInspection Calendar Task link label must use Task ID plus the actual Striven Task Name.');
  requireRule(pre.field854Policy === 'DO_NOT_MANAGE', 'PreInspection Field 854 must not be managed by V3.');
  requireRule(pre.infoCustomFieldsAtCreate === 'NONE', 'PreInspection CREATE must not prefill InfoCustomFields.');
  requireRule(pre.technicianFieldsPrefill === false, 'PreInspection technician-completed fields must not be prefilled.');
  requireRule(Number(pre.defaultPoolId || 0) === 8, 'PreInspection must use Pool 8.');
  requireRule(pre.requestedByFromOrganizer === true, 'PreInspection Requested By must resolve from Calendar organizer/creator.');

  requireRule(rules.orderRequired === false, 'Hard-rule mirror says PreInspection orderRequired must be false.');
  requireRule(rules.attachOrderToTask === false, 'Hard-rule mirror says PreInspection SO attachment is forbidden.');
  requireRule(Number(rules.taskTypeId || 0) === 105, 'Hard-rule mirror says PreInspection Task Type must be 105.');
  requireRule(rules.descriptionPolicy === 'BLANK_AT_CREATE', 'Hard-rule mirror says Description must be blank.');
  requireRule(rules.calendarNotesPolicy === 'CALENDAR_ONLY', 'Hard-rule mirror says Calendar notes stay on Calendar.');
  requireRule(rules.previousCalendarTitlePolicy === 'PRESERVE_AS_FIRST_DESCRIPTION_LINE_NO_HEADER', 'Hard-rule mirror says prior Calendar title is preserved as first Description line without a heading.');
  requireRule(rules.descriptionHeaderPolicy === 'REMOVE_NOTES_AND_SALES_ORDER_LABELS_PRESERVE_CONTENT', 'Hard-rule mirror says Notes:/Sales Order: labels are removed while content is preserved.');
  requireRule(rules.calendarCopyParityPolicy === 'FULL_BUSINESS_VISIBLE_PARITY_REQUIRED', 'Hard-rule mirror says required PreInspection Calendar copies must reach full business-visible parity.');
  requireRule(rules.calendarCopyParityFields === 'TITLE_DESCRIPTION_LOCATION_START_END_REQUIRED_PARTICIPATION', 'Hard-rule mirror parity field set is incomplete.');
  requireRule(rules.calendarLinksPreservationPolicy === 'PRESERVE_EXISTING_CLICKABLE_BLOCK_REPAIR_NONCLICKABLE_OR_MISSING_ONLY', 'Hard-rule mirror says existing clickable PreInspection Striven Links formatting must be preserved; missing or non-clickable links may be repaired.');
  requireRule(rules.calendarTaskLinkLabelPolicy === 'TASK_ID_PLUS_ACTUAL_STRIVEN_TASK_NAME', 'Hard-rule mirror says PreInspection Calendar Task links use actual Striven Task Name.');
  requireRule(rules.field854Policy === 'DO_NOT_MANAGE', 'Hard-rule mirror says Field 854 is removed from automation.');
  requireRule(rules.infoCustomFieldsAtCreate === 'NONE', 'Hard-rule mirror says no PreInspection custom-field prefill.');
  requireRule(rules.technicianFieldsPolicy === 'DO_NOT_PREFILL', 'Hard-rule mirror says technician fields are not prefilled.');
  requireRule(Number(rules.defaultPoolId || 0) === 8, 'Hard-rule mirror says Pool 8 is required.');
  requireRule(rules.organizerIsAssignee === false, 'PreInspection organizer must not become Assigned To.');
  requireRule(rules.calendarTitlePrefix === '', 'PreInspection canonical Calendar title must use the bare Customer Number with no prefix.');
  requireRule(rules.noOpenTaskPolicy === 'HISTORY_GATE_BEFORE_CREATE', 'PreInspection NO OPEN TASK must pass the history gate before CREATE.');
  requireRule(rules.noHistoryDecision === 'CREATE_TASK', 'PreInspection no-history outcome must be CREATE_TASK.');
  requireRule(rules.sameDayCompletedDecision === 'FULFILLED_NO_RECREATE', 'PreInspection same-day completed history must suppress recreate.');
  requireRule(rules.sameDayCancelledDecision === 'REVIEW', 'PreInspection same-day cancelled history must require REVIEW.');
  requireRule(rules.unprovenHistoryDateDecision === 'REVIEW', 'PreInspection unproven history date must require REVIEW.');
  requireRule(rules.futureHistoryDecision === 'REVIEW', 'PreInspection future history must require REVIEW.');
  requireRule(rules.olderHistoryDecision === 'RECREATE_TASK', 'PreInspection older history must resolve to RECREATE_TASK.');
  requireRule(rules.missingLinkedTaskCacheDecision === 'REVIEW_FAIL_CLOSED', 'PreInspection missing linked Task cache evidence must fail closed to REVIEW.');

  if (
    typeof TMV3_STEP7_ACTION !== 'undefined' &&
    TMV3_STEP7_ACTION.PATCH_FIELD854
  ) {
    errors.push('PATCH_FIELD854 must not exist in the Step 7 action vocabulary.');
  }

  if (errors.length) {
    throw new Error(
      'TMV3 HARD RULE VIOLATION [' +
      TMV3_HARD_RULES.version +
      ']: ' +
      errors.join(' | ')
    );
  }

  return {
    status:'PASS',
    rulesVersion:TMV3_HARD_RULES.version,
    engineVersion:TMV3.VERSION
  };
}

function tmv3_assertShadow_() {
  tmv3_assertHardRules_();
  const allowedReadPipelineModes = [
    'SHADOW_READ_ONLY',
    'CANARY_WRITE',
    'PRODUCTION_WRITE'
  ];

  if (allowedReadPipelineModes.indexOf(TMV3.MODE) === -1) {
    throw new Error(
      'V3 read/planning pipeline blocked: unsupported mode "' +
      String(TMV3.MODE || '') +
      '".'
    );
  }

  return true;
}



// Original: 15_Identity
/************************************************************
 * TM V3 — SHARED IDENTITY RESOLUTION
 *
 * Goals:
 * - One evidence ladder across all four verticals.
 * - Order-required verticals derive Customer from the verified
 *   Sales Order / Work Order rather than fuzzy Calendar text.
 * - PreInspection can resolve without an Order, but explicit
 *   Customer Number evidence is authoritative: if it is wrong,
 *   V3 REVIEWs rather than silently overriding it.
 * - Same surname alone is never enough.
 * - Customer-owned Location / Contact relationships are mandatory.
 ************************************************************/


function tmv3_identityDescriptionForRecord_(record) {
  record = record || {};

  if (record.identityDescriptionClean !== undefined) {
    return tmv3_clean_(record.identityDescriptionClean);
  }

  if (record.descriptionClean !== undefined) {
    return tmv3_clean_(record.descriptionClean);
  }

  return tmv3_calendarReadableDescription_(
    record.description || ''
  );
}

function tmv3_identityIndex_(refs) {
  const out = {
    contactsByCustomer: {},
    contactsByPhone: {},
    contactsByEmail: {},
    locationsByCustomer: refs.locationsByCustomer || {},
    locationOwnersByAddress: {},
    customerByNumber: refs.customerByNumber || {},
    customerByPhone: refs.customerByPhone || {},
    customerById: refs.customerById || {}
  };

  (refs.contacts || []).forEach(function(r) {
    const customerId = tmv3_clean_(r['Customer ID']);
    const phone = tmv3_phone10_(r['Phone']);
    const email = tmv3_normEmail_(r['Email']);

    if (customerId) {
      if (!out.contactsByCustomer[customerId]) out.contactsByCustomer[customerId] = [];
      out.contactsByCustomer[customerId].push(r);
    }

    if (phone) {
      if (!out.contactsByPhone[phone]) out.contactsByPhone[phone] = [];
      out.contactsByPhone[phone].push(r);
    }

    if (email) {
      if (!out.contactsByEmail[email]) out.contactsByEmail[email] = [];
      out.contactsByEmail[email].push(r);
    }
  });

  (refs.locations || []).forEach(function(r) {
    const key = tmv3_normalizeAddress_(tmv3_locationFullAddress_(r));
    if (!key) return;
    if (!out.locationOwnersByAddress[key]) out.locationOwnersByAddress[key] = [];
    out.locationOwnersByAddress[key].push(r);
  });

  return out;
}

function tmv3_resolveIdentity_(eventRecord, cfg, refs, order) {
  const ix = tmv3_identityIndex_(refs);
  let evidence = [];

  let customer = null;
  let location = null;
  let contact = null;

  if (cfg.orderRequired) {
    if (!order) {
      return tmv3_identityFail_(
        'BLOCKED',
        'ORDER_REQUIRED',
        cfg.orderLabel + ' is required before identity can be resolved.'
      );
    }

    const customerId = tmv3_clean_(order['Customer ID']);
    if (!customerId || !refs.customerById[customerId]) {
      return tmv3_identityFail_(
        'BLOCKED',
        'ORDER_CUSTOMER_MISSING',
        'Verified ' + cfg.orderLabel + ' does not resolve to a valid Customer.'
      );
    }

    customer = refs.customerById[customerId];
    evidence.push('CUSTOMER_FROM_' + tmv3_normCode_(cfg.orderLabel));

    const orderLocationId = tmv3_clean_(order['Location ID']);
    if (orderLocationId) {
      const orderLocation =
        refs.locationById && refs.locationById[orderLocationId]
          ? refs.locationById[orderLocationId]
          : null;

      if (orderLocation) {
        if (tmv3_clean_(orderLocation['Customer ID']) !== customerId) {
          return tmv3_identityFail_(
            'BLOCKED',
            'ORDER_LOCATION_OWNERSHIP_CONFLICT',
            cfg.orderLabel + ' Location resolves to a Location owned by a different Customer.',
            customer,
            null,
            null,
            evidence
          );
        }

        location = orderLocation;
        evidence.push('LOCATION_FROM_' + tmv3_normCode_(cfg.orderLabel));
      } else {
        // Missing source evidence is not the same as a proven ownership
        // conflict. Fall back only to the normal customer-owned Location
        // resolver, which still requires a strong address match.
        evidence.push('ORDER_LOCATION_ID_NOT_IN_LOCATION_SOURCE');
      }
    }

    if (!location) {
      const locResult = tmv3_resolveOwnedLocation_(
        customer,
        eventRecord.location,
        refs.locationsByCustomer[customerId] || []
      );

      if (locResult.status !== 'MATCHED') {
        return tmv3_identityFail_(
          locResult.status,
          locResult.errorCode,
          locResult.reason,
          customer,
          null,
          null,
          evidence.concat(locResult.evidence || [])
        );
      }

      location = locResult.location;
      evidence = evidence.concat(locResult.evidence || []);
    }

    const orderContactId = tmv3_clean_(order['Contact ID']);
    const contactResult = tmv3_resolveOwnedContact_(
      customer,
      orderContactId,
      eventRecord,
      ix
    );

    if (contactResult.status === 'REVIEW') {
      return tmv3_identityFail_(
        'REVIEW',
        contactResult.errorCode,
        contactResult.reason,
        customer,
        location,
        null,
        evidence.concat(contactResult.evidence || [])
      );
    }

    contact = contactResult.contact || null;
    evidence = evidence.concat(contactResult.evidence || []);

    return {
      status: 'MATCHED',
      customer: customer,
      location: location,
      contact: contact,
      evidence: evidence
    };
  }

  // PreInspection / no-order-required path.
  const explicitCustomerNumber = tmv3_clean_(eventRecord.customerNumber);

  if (explicitCustomerNumber) {
    customer = refs.customerByNumber[explicitCustomerNumber] || null;

    if (!customer) {
      return tmv3_identityFail_(
        'REVIEW',
        'EXPLICIT_CUSTOMER_NUMBER_NOT_FOUND',
        'Explicit Customer Number ' + explicitCustomerNumber +
          ' was not found. Other evidence is retained but cannot silently override it.'
      );
    }

    evidence.push('CUSTOMER_NUMBER_EXACT');
  }

  if (!customer && eventRecord.phone) {
    const phoneMatches = refs.customerByPhone[tmv3_phone10_(eventRecord.phone)] || [];

    if (phoneMatches.length === 1) {
      customer = phoneMatches[0];
      evidence.push('CUSTOMER_PHONE_EXACT');
    } else if (phoneMatches.length > 1) {
      return tmv3_identityFail_(
        'REVIEW',
        'AMBIGUOUS_CUSTOMER_PHONE',
        'Calendar phone matches multiple Customers.'
      );
    }
  }

  if (!customer && eventRecord.location) {
    const addressKey = tmv3_normalizeAddress_(eventRecord.location);
    const ownedLocations = ix.locationOwnersByAddress[addressKey] || [];
    const ownerIds = tmv3_unique_(ownedLocations.map(function(r) {
      return tmv3_clean_(r['Customer ID']);
    }));

    if (ownerIds.length === 1 && refs.customerById[ownerIds[0]]) {
      const candidateCustomer = refs.customerById[ownerIds[0]];
      const nameCheck = tmv3_nameCorroboratesCustomer_(
        eventRecord.title + ' ' + tmv3_identityDescriptionForRecord_(eventRecord),
        candidateCustomer['Name']
      );

      // Exact address can identify a household, but same surname alone is not
      // treated as proof. Require at least useful name or phone corroboration.
      if (nameCheck || tmv3_phoneCorroboratesCustomer_(eventRecord, candidateCustomer, ix)) {
        customer = candidateCustomer;
        evidence.push('CUSTOMER_FROM_EXACT_ADDRESS_PLUS_IDENTITY');
      }
    }
  }

  if (!customer) {
    return tmv3_identityFail_(
      'REVIEW',
      'CUSTOMER_UNRESOLVED',
      'No deterministic Customer match was proven.'
    );
  }

  const customerId = tmv3_clean_(customer['Customer ID']);
  const locResult = tmv3_resolveOwnedLocation_(
    customer,
    eventRecord.location,
    refs.locationsByCustomer[customerId] || []
  );

  if (locResult.status !== 'MATCHED') {
    return tmv3_identityFail_(
      locResult.status,
      locResult.errorCode,
      locResult.reason,
      customer,
      null,
      null,
      evidence.concat(locResult.evidence || [])
    );
  }

  location = locResult.location;
  evidence = evidence.concat(locResult.evidence || []);

  const contactResult = tmv3_resolveOwnedContact_(
    customer,
    '',
    eventRecord,
    ix
  );

  // PreInspection contact ambiguity does not invent a Contact. It can continue
  // without one unless another rule explicitly requires it.
  if (contactResult.status === 'MATCHED') {
    contact = contactResult.contact || null;
    evidence = evidence.concat(contactResult.evidence || []);
  } else if (contactResult.status === 'REVIEW') {
    evidence.push('CONTACT_UNRESOLVED_NO_WRITE');
  }

  return {
    status: 'MATCHED',
    customer: customer,
    location: location,
    contact: contact,
    evidence: evidence
  };
}

function tmv3_resolveOwnedLocation_(customer, calendarAddress, ownedLocations) {
  const customerId = tmv3_clean_(customer && customer['Customer ID']);
  const target = tmv3_normalizeAddress_(calendarAddress);

  if (!target) {
    return {
      status: 'REVIEW',
      errorCode: 'CALENDAR_LOCATION_MISSING',
      reason: 'Calendar appointment has no usable Location.',
      evidence: []
    };
  }

  const exact = (ownedLocations || []).filter(function(r) {
    return tmv3_normalizeAddress_(tmv3_locationFullAddress_(r)) === target;
  });

  if (exact.length === 1) {
    return {
      status: 'MATCHED',
      location: exact[0],
      evidence: ['CUSTOMER_LOCATION_EXACT']
    };
  }

  if (exact.length > 1) {
    return {
      status: 'REVIEW',
      errorCode: 'AMBIGUOUS_EXACT_LOCATION',
      reason: 'Multiple Customer-owned Locations normalize to the Calendar address.',
      evidence: []
    };
  }

  const targetParts = tmv3_addressParts_(calendarAddress);
  const strong = (ownedLocations || []).filter(function(r) {
    const parts = tmv3_addressParts_(tmv3_locationFullAddress_(r));
    return tmv3_addressStrongMatch_(targetParts, parts);
  });

  if (strong.length === 1) {
    return {
      status: 'MATCHED',
      location: strong[0],
      evidence: ['CUSTOMER_LOCATION_STRONG']
    };
  }

  if (strong.length > 1) {
    return {
      status: 'REVIEW',
      errorCode: 'AMBIGUOUS_STRONG_LOCATION',
      reason: 'Multiple Customer-owned Locations strongly match the Calendar address.',
      evidence: []
    };
  }

  return {
    status: 'REVIEW',
    errorCode: 'LOCATION_UNRESOLVED',
    reason: customerId
      ? 'Customer is proven, but the Calendar address does not match a known Customer-owned Location.'
      : 'Location cannot be resolved because Customer ownership is unknown.',
    evidence: []
  };
}

function tmv3_resolveOwnedContact_(customer, preferredContactId, eventRecord, ix) {
  const customerId = tmv3_clean_(customer && customer['Customer ID']);
  let owned = (ix.contactsByCustomer[customerId] || []).slice();

  if (preferredContactId) {
    const preferred = owned.filter(function(r) {
      return tmv3_clean_(r['Contact ID']) === tmv3_clean_(preferredContactId);
    });

    if (preferred.length === 1) {
      return {
        status: 'MATCHED',
        contact: preferred[0],
        evidence: ['CONTACT_FROM_CACHE']
      };
    }

    try {
      const direct = tmv3_getContactById_(preferredContactId, customerId);

      if (direct && direct.__ownershipVerified) {
        return {
          status: 'MATCHED',
          contact: direct,
          evidence: ['CONTACT_FROM_ORDER_VERIFIED_BY_API']
        };
      }
    } catch (err) {
      return {
        status: 'REVIEW',
        errorCode: 'CONTACT_READ_FAILED',
        reason:
          'Order Contact ' +
          preferredContactId +
          ' could not be freshly verified: ' +
          String(err && err.message || err),
        evidence: []
      };
    }

    return {
      status: 'REVIEW',
      errorCode: 'ORDER_CONTACT_OWNERSHIP_CONFLICT',
      reason: 'Order Contact is not proven to belong to the resolved Customer.',
      evidence: []
    };
  }

  try {
    owned = tmv3_getCustomerContacts_(customerId);
  } catch (err) {
    return {
      status: 'REVIEW',
      errorCode: 'CUSTOMER_CONTACTS_READ_FAILED',
      reason:
        'Customer Contacts could not be read: ' +
        String(err && err.message || err),
      evidence: []
    };
  }

  const emails = tmv3_extractEmails_(
    eventRecord.title + ' ' + tmv3_identityDescriptionForRecord_(eventRecord)
  );
  const phones = tmv3_allPhones_(
    eventRecord.title + ' ' + tmv3_identityDescriptionForRecord_(eventRecord) + ' ' + eventRecord.location
  );

  let candidates = owned.filter(function(r) {
    const emailList = r.Emails || (r['Email'] ? [tmv3_normEmail_(r['Email'])] : []);
    const phoneList = r.Phones || (r['Phone'] ? [tmv3_phone10_(r['Phone'])] : []);

    const emailHit = emailList.some(function(email) {
      return email && emails.indexOf(tmv3_normEmail_(email)) !== -1;
    });

    const phoneHit = phoneList.some(function(phone) {
      const cleanPhone = tmv3_phone10_(phone);
      return cleanPhone && phones.indexOf(cleanPhone) !== -1;
    });

    return emailHit || phoneHit;
  });

  const seen = {};
  candidates = candidates.filter(function(r) {
    const id = tmv3_clean_(r['Contact ID']);
    if (!id || seen[id]) return false;
    seen[id] = true;
    return true;
  });

  if (candidates.length === 1) {
    return {
      status: 'MATCHED',
      contact: candidates[0],
      evidence: ['CUSTOMER_CONTACT_PHONE_EMAIL_EXACT']
    };
  }

  if (candidates.length > 1) {
    return {
      status: 'REVIEW',
      errorCode: 'AMBIGUOUS_CONTACT',
      reason: 'Calendar phone/email matches multiple Contacts owned by the Customer.',
      evidence: []
    };
  }

  if (owned.length === 1) {
    return {
      status: 'MATCHED',
      contact: owned[0],
      evidence: ['ONLY_CUSTOMER_CONTACT']
    };
  }

  return {
    status: 'NO_MATCH',
    contact: null,
    evidence: []
  };
}

function tmv3_identityFail_(status, errorCode, reason, customer, location, contact, evidence) {
  return {
    status: status || 'REVIEW',
    errorCode: errorCode || 'IDENTITY_REVIEW',
    reason: reason || 'Identity requires review.',
    customer: customer || null,
    location: location || null,
    contact: contact || null,
    evidence: evidence || []
  };
}

function tmv3_normalizeAddress_(value) {
  return tmv3_clean_(value)
    .toLowerCase()
    .replace(/&amp;/g, ' and ')
    .replace(/\bcanada\b/g, ' ')
    .replace(/\bontario\b/g, ' on ')
    .replace(/\broad\b/g, ' rd ')
    .replace(/\bstreet\b/g, ' st ')
    .replace(/\bavenue\b/g, ' ave ')
    .replace(/\bdrive\b/g, ' dr ')
    .replace(/\bcrescent\b/g, ' cres ')
    .replace(/\bcourt\b/g, ' ct ')
    .replace(/\bboulevard\b/g, ' blvd ')
    .replace(/\bplace\b/g, ' pl ')
    .replace(/\blane\b/g, ' ln ')
    .replace(/[.,#()]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function tmv3_addressParts_(value) {
  const normalized = tmv3_normalizeAddress_(value);
  const postal = (tmv3_clean_(value).toUpperCase().match(/\b[A-Z]\d[A-Z]\s?\d[A-Z]\d\b/) || [])[0] || '';
  const streetNo = (normalized.match(/^\s*(\d+[a-z]?)/i) || [])[1] || '';
  const tokens = normalized.split(' ').filter(Boolean);

  return {
    normalized: normalized,
    postal: postal.replace(/\s+/g, ''),
    streetNo: streetNo,
    tokens: tokens
  };
}

function tmv3_addressStrongMatch_(a, b) {
  if (!a || !b || !a.normalized || !b.normalized) return false;
  if (a.streetNo && b.streetNo && a.streetNo !== b.streetNo) return false;

  if (a.postal && b.postal && a.postal !== b.postal) return false;

  const aSet = {};
  a.tokens.forEach(function(t) { if (t.length > 2) aSet[t] = true; });
  const overlap = b.tokens.filter(function(t) {
    return t.length > 2 && aSet[t];
  }).length;

  return overlap >= 3 || (
    !!a.streetNo &&
    a.streetNo === b.streetNo &&
    overlap >= 2
  );
}

function tmv3_locationFullAddress_(r) {
  return [
    r['Address 1'],
    r['Address 2'],
    r['City'],
    r['Province'],
    r['Postal Code']
  ].map(tmv3_clean_).filter(Boolean).join(' ');
}

function tmv3_extractEmails_(text) {
  const matches = tmv3_clean_(text).match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || [];
  return tmv3_unique_(matches.map(tmv3_normEmail_));
}

function tmv3_normEmail_(value) {
  return tmv3_clean_(value).toLowerCase();
}

function tmv3_nameCorroboratesCustomer_(haystack, customerName) {
  const hay = tmv3_norm_(haystack);
  const name = tmv3_norm_(customerName);
  if (!hay || !name) return false;

  if (hay.indexOf(name) !== -1) return true;

  const parts = name.split(' ').filter(function(x) { return x.length > 1; });
  if (parts.length < 2) return false;

  const first = parts[0];
  const last = parts[parts.length - 1];

  // Minor first-name tolerance is allowed only with the same surname and
  // other evidence. Surname alone is never enough.
  return hay.indexOf(last) !== -1 && hay.indexOf(first) !== -1;
}

function tmv3_phoneCorroboratesCustomer_(eventRecord, customer, ix) {
  const phones = tmv3_allPhones_(
    eventRecord.title + ' ' + tmv3_identityDescriptionForRecord_(eventRecord) + ' ' + eventRecord.location
  );

  if (!phones.length) return false;

  const customerId = tmv3_clean_(customer && customer['Customer ID']);
  const primary = tmv3_phone10_(customer && customer['Primary Phone']);
  if (primary && phones.indexOf(primary) !== -1) return true;

  return (ix.contactsByCustomer[customerId] || []).some(function(r) {
    const p = tmv3_phone10_(r['Phone']);
    return p && phones.indexOf(p) !== -1;
  });
}

function tmv3_normCode_(value) {
  return tmv3_clean_(value).toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}


// Original: 20_Striven
var TMV3_RUNTIME_METRICS = {
  strivenApiCalls: 0,
  strivenApiFailures: 0,
  strivenApiBudgetBlocks: 0
};

var TMV3_CANARY_API_ALLOWANCE_ACTIVE = false;
var TMV3_PREINSPECTION_API_CONTEXT_ACTIVE = false;

function tmv3_resetRuntimeMetrics_() {
  TMV3_RUNTIME_METRICS = {
    strivenApiCalls: 0,
    strivenApiFailures: 0,
    strivenApiBudgetBlocks: 0
  };
}

function tmv3_runtimeMetrics_() {
  return {
    strivenApiCalls: Number(TMV3_RUNTIME_METRICS.strivenApiCalls || 0),
    strivenApiFailures: Number(TMV3_RUNTIME_METRICS.strivenApiFailures || 0),
    strivenApiBudgetBlocks: Number(TMV3_RUNTIME_METRICS.strivenApiBudgetBlocks || 0),
    dailyBudget: tmv3_strivenApiBudgetStatus_()
  };
}

function tmv3_strivenApiBudgetConfig_() {
  return TMV3.API_BUDGET || {
    planDailyLimit: 5000,
    v3DailySoftLimit: 1200,
    reserveForOtherWorkflows: 3800,
    propertyPrefix: 'TMV3_STRIVEN_API_BUDGET'
  };
}

function tmv3_strivenApiBudgetDate_() {
  return Utilities.formatDate(
    new Date(),
    TMV3_TIMEZONE || Session.getScriptTimeZone() || 'America/Toronto',
    'yyyy-MM-dd'
  );
}

function tmv3_strivenApiBudgetKeys_() {
  const prefix = String(
    tmv3_strivenApiBudgetConfig_().propertyPrefix ||
    'TMV3_STRIVEN_API_BUDGET'
  );

  return {
    date: prefix + '_DATE',
    count: prefix + '_COUNT',
    exhaustedDate: prefix + '_EXHAUSTED_DATE',
    exhaustedReason: prefix + '_EXHAUSTED_REASON'
  };
}

function tmv3_strivenApiBudgetStatus_() {
  const cfg = tmv3_strivenApiBudgetConfig_();
  const keys = tmv3_strivenApiBudgetKeys_();
  const props = PropertiesService.getScriptProperties();
  const today = tmv3_strivenApiBudgetDate_();
  const storedDate = String(props.getProperty(keys.date) || '');
  const count = storedDate === today
    ? Number(props.getProperty(keys.count) || 0)
    : 0;
  const exhausted =
    String(props.getProperty(keys.exhaustedDate) || '') === today;
  const internalLimitEnabled = cfg.enabled !== false;
  const softLimit = internalLimitEnabled ? Number(cfg.v3DailySoftLimit || 1200) : null;

  return {
    date: today,
    count: count,
    enabled: internalLimitEnabled,
    limitMode: internalLimitEnabled ? 'ENFORCED' : 'TRACKING_ONLY',
    softLimit: softLimit,
    remaining: exhausted ? 0 : Math.max(0, tmv3_strivenApiEffectiveSoftLimit_({softLimit:softLimit},cfg) - count),
    preInspectionReservedCalls: internalLimitEnabled ? Number(cfg.preInspectionReservedCalls || 0) : 0,
    planDailyLimit: Number(cfg.planDailyLimit || 5000),
    reserveForOtherWorkflows: internalLimitEnabled ? Number(cfg.reserveForOtherWorkflows || 0) : 0,
    exhausted: exhausted,
    exhaustedReason: exhausted
      ? String(props.getProperty(keys.exhaustedReason) || '')
      : ''
  };
}

function tmv3_strivenApiEffectiveSoftLimit_(status, cfg, context) {
  status = status || {};
  cfg = cfg || tmv3_strivenApiBudgetConfig_();
  // Finite sentinel preserves existing numeric queue checks and JSON serialization.
  if (cfg.enabled === false) return Number.MAX_SAFE_INTEGER;
  context = context || {};

  const baseSoftLimit = Number(
    status.softLimit || cfg.v3DailySoftLimit || 1200
  );
  const planDailyLimit = Number(cfg.planDailyLimit || 5000);
  const active =
    Object.prototype.hasOwnProperty.call(context, 'canaryActive')
      ? context.canaryActive === true
      : TMV3_CANARY_API_ALLOWANCE_ACTIVE === true;
  const mode =
    Object.prototype.hasOwnProperty.call(context, 'mode')
      ? String(context.mode || '')
      : String(TMV3.MODE || '');
  const configuredExtra =
    Object.prototype.hasOwnProperty.call(context, 'extraCalls')
      ? Number(context.extraCalls || 0)
      : Number(
          TMV3.OPERATIONS &&
          TMV3.OPERATIONS.canaryApiExtraCalls ||
          0
        );
  const extra =
    active && mode === 'CANARY_WRITE'
      ? Math.max(0, configuredExtra)
      : 0;

  const ceiling=Math.min(planDailyLimit,baseSoftLimit+extra);
  // Ring-fence within the existing 1,200-call V3 allowance, never above it.
  const reserve=mode==='PRODUCTION_WRITE' && TMV3_PREINSPECTION_API_CONTEXT_ACTIVE!==true
    ? Math.min(ceiling,Math.max(0,Number(cfg.preInspectionReservedCalls || 0))) : 0;
  return ceiling-reserve;
}

function tmv3_reserveStrivenApiCall_(label) {
  const cfg = tmv3_strivenApiBudgetConfig_();
  const keys = tmv3_strivenApiBudgetKeys_();
  const props = PropertiesService.getScriptProperties();
  const status = tmv3_strivenApiBudgetStatus_();
  const effectiveSoftLimit =
    tmv3_strivenApiEffectiveSoftLimit_(status, cfg);

  if (status.exhausted || (cfg.enabled !== false && status.count >= effectiveSoftLimit)) {
    TMV3_RUNTIME_METRICS.strivenApiBudgetBlocks =
      Number(TMV3_RUNTIME_METRICS.strivenApiBudgetBlocks || 0) + 1;

    throw new Error(
      'TMV3_STRIVEN_API_DAILY_GUARD: blocked ' +
      String(label || 'Striven API call') +
      '. V3 daily usage=' + status.count +
      '/' + effectiveSoftLimit +
      (effectiveSoftLimit !== status.softLimit
        ? ' (base soft limit ' + status.softLimit + ')'
        : '') +
      '; reserved for other workflows=' +
      Number(cfg.reserveForOtherWorkflows || 0) +
      (status.exhaustedReason
        ? '; Striven reported: ' + status.exhaustedReason
        : '')
    );
  }

  const nextCount = status.count + 1;
  const values = {};
  values[keys.date] = status.date;
  values[keys.count] = String(nextCount);
  props.setProperties(values, false);

  return {
    date: status.date,
    count: nextCount,
    enabled: cfg.enabled !== false,
    softLimit: cfg.enabled === false ? null : effectiveSoftLimit,
    baseSoftLimit: status.softLimit
  };
}

function tmv3_markStrivenDailyLimitExhausted_(reason) {
  const keys = tmv3_strivenApiBudgetKeys_();
  const props = PropertiesService.getScriptProperties();
  const values = {};
  values[keys.exhaustedDate] = tmv3_strivenApiBudgetDate_();
  values[keys.exhaustedReason] = String(reason || '').slice(0, 500);
  props.setProperties(values, false);
}

function tmv3_isStrivenDailyLimitResponse_(code, text) {
  const body = String(text || '').toLowerCase();
  return (
    Number(code) === 429 &&
    (
      body.indexOf('daily api usage limit') !== -1 ||
      body.indexOf('5000 api calls per day') !== -1
    )
  );
}

// Public diagnostic. Does not call Striven.
function tmv3_strivenApiBudgetStatus() {
  const status = tmv3_strivenApiBudgetStatus_();
  Logger.log(JSON.stringify(status, null, 2));
  return status;
}

function tmv3_token_() {
  const cache = CacheService.getScriptCache();
  const cached = tmv3_clean_(cache.get('TMV3_STRIVEN_TOKEN'));
  if (cached) return cached;

  const clientId = tmv3_property_(TMV3.PROPERTIES.CLIENT_ID, true);
  const clientSecret = tmv3_property_(TMV3.PROPERTIES.CLIENT_SECRET, true);

  tmv3_reserveStrivenApiCall_('TOKEN_REFRESH');

  const response = UrlFetchApp.fetch(TMV3.API_BASE + '/accesstoken', {
    method: 'post',
    payload: {
      grant_type: 'client_credentials',
      client_id: clientId,
      client_secret: clientSecret,
      scope: 'api'
    },
    contentType: 'application/x-www-form-urlencoded',
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    TMV3_RUNTIME_METRICS.strivenApiFailures =
      Number(TMV3_RUNTIME_METRICS.strivenApiFailures || 0) + 1;

    const tokenErrorText = response.getContentText() || '';
    if (tmv3_isStrivenDailyLimitResponse_(code, tokenErrorText)) {
      tmv3_markStrivenDailyLimitExhausted_(tokenErrorText);
    }

    throw new Error(
      'Striven authentication failed HTTP ' +
      code +
      ': ' +
      tokenErrorText.slice(0, 600)
    );
  }

  const body = JSON.parse(response.getContentText() || '{}');
  const token = tmv3_clean_(body.access_token);

  if (!token) {
    throw new Error('Striven authentication returned no access_token.');
  }

  const ttl = Math.max(
    60,
    Math.min(Number(body.expires_in || 3600) - 300, 3300)
  );

  cache.put('TMV3_STRIVEN_TOKEN', token, ttl);
  return token;
}

function tmv3_fetchJson_(url, options) {
  const opts = options || {};
  opts.muteHttpExceptions = true;
  opts.headers = Object.assign({
    Authorization: 'Bearer ' + tmv3_token_(),
    Accept: 'application/json'
  }, opts.headers || {});

  tmv3_reserveStrivenApiCall_(
    String(opts.method || 'get').toUpperCase() + ' ' + String(url || '')
  );

  TMV3_RUNTIME_METRICS.strivenApiCalls =
    Number(TMV3_RUNTIME_METRICS.strivenApiCalls || 0) + 1;

  const response = UrlFetchApp.fetch(url, opts);
  const code = response.getResponseCode();
  const text = response.getContentText();

  let json = null;
  try {
    json = text ? JSON.parse(text) : null;
  } catch (err) {}

  if (code < 200 || code >= 300) {
    TMV3_RUNTIME_METRICS.strivenApiFailures =
      Number(TMV3_RUNTIME_METRICS.strivenApiFailures || 0) + 1;

    if (tmv3_isStrivenDailyLimitResponse_(code, text)) {
      tmv3_markStrivenDailyLimitExhausted_(text);
    }

    throw new Error(
      'Striven request failed HTTP ' +
      code +
      ': ' +
      text.slice(0, 600)
    );
  }

  return json;
}

function tmv3_reportRows_(propertyAliases) {
  const baseUrl = tmv3_property_(propertyAliases, true);
  const all = [];
  let naturalTermination = false;

  for (let pageIndex = 0; pageIndex < TMV3.MAX_PAGES; pageIndex++) {
    const sep = baseUrl.indexOf('?') === -1 ? '?' : '&';
    const url =
      baseUrl +
      sep +
      'PageIndex=' +
      pageIndex +
      '&PageSize=' +
      TMV3.PAGE_SIZE;

    const json = tmv3_fetchJson_(url, { method: 'get' });
    const rows = tmv3_extractReportRows_(json);

    if (!rows.length) {
      naturalTermination = true;
      break;
    }

    Array.prototype.push.apply(all, rows);

    if (rows.length < TMV3.PAGE_SIZE) {
      naturalTermination = true;
      break;
    }
  }

  if (!naturalTermination) {
    throw new Error(
      'TMV3_REPORT_PAGINATION_TRUNCATED: report remained full through ' +
      TMV3.MAX_PAGES +
      ' pages at PageSize ' +
      TMV3.PAGE_SIZE +
      '. Existing cache was preserved.'
    );
  }

  return all;
}

function tmv3_extractReportRows_(json) {
  if (Array.isArray(json)) return json;

  if (!json || typeof json !== 'object') {
    throw new Error(
      'TMV3_REPORT_SCHEMA_INVALID: expected an array or report object.'
    );
  }

  const keys = ['data','Data','rows','Rows','Items','items'];

  for (let i = 0; i < keys.length; i++) {
    if (Object.prototype.hasOwnProperty.call(json, keys[i])) {
      if (!Array.isArray(json[keys[i]])) {
        throw new Error(
          'TMV3_REPORT_SCHEMA_INVALID: field ' +
          keys[i] +
          ' is present but is not an array.'
        );
      }
      return json[keys[i]];
    }
  }

  throw new Error(
    'TMV3_REPORT_SCHEMA_INVALID: no recognized report-row array was returned.'
  );
}

function tmv3_assertSourceReplacementSafe_(label, rows, sheetName) {
  if (!Array.isArray(rows)) {
    throw new Error(
      'TMV3_SOURCE_REPLACEMENT_BLOCKED: ' +
      String(label || 'source') +
      ' did not produce an array.'
    );
  }

  const priorRows = sheetName
    ? tmv3_rows_(sheetName).length
    : 0;

  if (!rows.length && priorRows > 0) {
    throw new Error(
      'TMV3_SOURCE_EMPTY_KEEP_LAST_GOOD: ' +
      String(label || 'source') +
      ' returned zero rows while the existing verified cache contains ' +
      priorRows +
      ' row(s). Existing cache was preserved.'
    );
  }

  return {
    label: String(label || ''),
    incomingRows: rows.length,
    priorRows: priorRows
  };
}

function tmv3_refreshMasterSources_() {
  tmv3_assertShadow_();

  const customers =
    tmv3_reportRows_(TMV3.PROPERTIES.CUSTOMERS)
      .map(tmv3_normalizeCustomer_);

  const customerNumberToId = {};
  customers.forEach(function(r) {
    if (tmv3_clean_(r[1]) && tmv3_clean_(r[0])) {
      customerNumberToId[tmv3_clean_(r[1])] = tmv3_clean_(r[0]);
    }
  });

  const locations =
    tmv3_reportRows_(TMV3.PROPERTIES.LOCATIONS)
      .map(function(r) {
        return tmv3_normalizeLocation_(r, customerNumberToId);
      });

  tmv3_assertSourceReplacementSafe_(
    'Customers',
    customers,
    TMV3.SHEETS.CUSTOMERS
  );
  tmv3_assertSourceReplacementSafe_(
    'Locations',
    locations,
    TMV3.SHEETS.LOCATIONS
  );

  tmv3_replaceRows_(
    TMV3.SHEETS.CUSTOMERS,
    [
      'Customer ID',
      'Customer Number',
      'Name',
      'Primary Phone',
      'Primary Email',
      'Fingerprint'
    ],
    customers
  );

  tmv3_replaceRows_(
    TMV3.SHEETS.LOCATIONS,
    [
      'Location ID',
      'Customer ID',
      'Address 1',
      'Address 2',
      'City',
      'Province',
      'Postal Code',
      'Phone',
      'Fingerprint'
    ],
    locations
  );

  // Stage 4 no longer performs routine Contact API reads. Keep the optional
  // Contact cache surface explicitly empty so stale Contact rows cannot become
  // planning evidence.
  tmv3_replaceRows_(
    TMV3.SHEETS.CONTACTS,
    [
      'Contact ID',
      'Customer ID',
      'Name',
      'Phone',
      'Email',
      'Fingerprint'
    ],
    []
  );

  return {
    customers: customers.length,
    locations: locations.length,
    contacts: 0
  };
}









function tmv3_refreshTransactionSources_(options) {
  options = options || {};
  tmv3_assertShadow_();

  const approved =
    tmv3_reportRows_(TMV3.PROPERTIES.APPROVED_ORDERS)
      .map(function(r) {
        return tmv3_normalizeOrder_(r, 'SALES_ORDER');
      });

  const deliveryApproved =
    tmv3_reportRows_(TMV3.PROPERTIES.DELIVERY_APPROVED_ORDERS)
      .map(function(r) {
        return tmv3_normalizeOrder_(r, 'DELIVERY_APPROVED');
      });

  const serviceOrders =
    tmv3_reportRows_(TMV3.PROPERTIES.SERVICE_WORK_ORDERS)
      .map(function(r) {
        return tmv3_normalizeOrder_(r, 'WORK_ORDER');
      });

  const orders = tmv3_mergeOrderRows_(
    approved.concat(deliveryApproved, serviceOrders)
  );

  const orderIdByNumber = {};
  orders.forEach(function(r) {
    if (tmv3_clean_(r[1]) && tmv3_clean_(r[0])) {
      orderIdByNumber[tmv3_clean_(r[1])] = tmv3_clean_(r[0]);
    }
  });

  const installTasks =
    tmv3_reportRows_(TMV3.PROPERTIES.INSTALL_TASKS)
      .map(function(r) {
        return tmv3_normalizeTask_(r, 'Install', orderIdByNumber);
      });

  const deliveryTasks =
    tmv3_reportRows_(TMV3.PROPERTIES.DELIVERY_TASKS)
      .map(function(r) {
        return tmv3_normalizeTask_(r, 'Delivery', orderIdByNumber);
      });

  const serviceTasks =
    tmv3_reportRows_(TMV3.PROPERTIES.SERVICE_TASKS)
      .map(function(r) {
        return tmv3_normalizeTask_(r, 'Service', orderIdByNumber);
      });

  const preInspectionSearch =
    tmv3_preInspectionTaskSourceRows_({
      force:options.forcePreInspectionTasks === true
    });
  const preInspectionTasks =
    preInspectionSearch.rows;

  const tasks = tmv3_dedupeObjects_(
    installTasks.concat(
      deliveryTasks,
      serviceTasks,
      preInspectionTasks
    ),
    function(r) {
      return r[0];
    }
  );

  tmv3_assertSourceReplacementSafe_(
    'Orders',
    orders,
    TMV3.SHEETS.ORDERS
  );
  tmv3_assertSourceReplacementSafe_(
    'Tasks',
    tasks,
    TMV3.SHEETS.TASKS
  );

  tmv3_replaceRows_(
    TMV3.SHEETS.ORDERS,
    [
      'Order ID',
      'Order Number',
      'Customer ID',
      'Location ID',
      'Contact ID',
      'Status',
      'Name',
      'Order Type',
      'URL',
      'Fingerprint'
    ],
    orders
  );

  tmv3_replaceRows_(
    TMV3.SHEETS.TASKS,
    [
      'Task ID',
      'Task Number',
      'Task Type ID',
      'Task Type',
      'Status',
      'Name',
      'Customer ID',
      'Location ID',
      'Contact ID',
      'Order ID',
      'Start',
      'Due',
      'Assignees',
      'Pools',
      'URL',
      'Fingerprint'
    ],
    tasks
  );

  return {
    orders: orders.length,
    tasks: tasks.length,
    installTasks: installTasks.length,
    deliveryTasks: deliveryTasks.length,
    serviceTasks: serviceTasks.length,
    preInspectionTasks: preInspectionTasks.length,
    preInspectionCacheReused: preInspectionSearch.reused === true,
    preInspectionTargetCustomers: Number(
      preInspectionSearch.targetCustomers || 0
    ),
    preInspectionSearchedCustomers: Number(
      preInspectionSearch.searchedCustomers || 0
    ),
    preInspectionHydratedTasks: Number(
      preInspectionSearch.hydratedTasks || 0
    ),
    preInspectionPriorTasks: Number(
      preInspectionSearch.priorTasks || 0
    ),
    preInspectionCacheAgeMinutes:
      preInspectionSearch.ageMinutes
  };
}

function tmv3_refreshSources() {
  tmv3_assertShadow_();
  tmv3_resetRuntimeMetrics_();

  const master = tmv3_refreshMasterSources_();
  const transactions = tmv3_refreshTransactionSources_();
  const refreshedAt = new Date().toISOString();

  PropertiesService.getScriptProperties().setProperties({
    TMV3_LAST_MASTER_SOURCE_REFRESH: refreshedAt,
    TMV3_LAST_OPERATIONAL_SOURCE_REFRESH: refreshedAt
  }, false);

  tmv3_audit_(
    'SYSTEM',
    '',
    '',
    'REFRESH_SOURCES',
    'PASS',
    'Customers ' +
      master.customers +
      '; Locations ' +
      master.locations +
      '; Contacts 0; Orders ' +
      transactions.orders +
      '; Tasks ' +
      transactions.tasks +
      '; API ' +
      JSON.stringify(tmv3_runtimeMetrics_())
  );

  return {
    customers: master.customers,
    locations: master.locations,
    contacts: 0,
    orders: transactions.orders,
    tasks: transactions.tasks,
    refreshedAt: refreshedAt,
    api: tmv3_runtimeMetrics_()
  };
}

function tmv3_normalizeCustomer_(r) {
  const id =
    tmv3_first_(r, [
      'CustomerId',
      'Customer ID',
      'CustomerCustomerId',
      'Id',
      'AccountId',
      'Account ID'
    ]);

  const no =
    tmv3_first_(r, [
      'CustomerNumber',
      'Customer Number',
      'AccountNumber',
      'Account Number',
      'Customer #'
    ]);

  const name =
    tmv3_first_(r, [
      'CustomerName',
      'Customer Name',
      'FullName',
      'Name',
      'AccountName',
      'Account Name'
    ]);

  const phone =
    tmv3_first_(r, [
      'Phone',
      'Phone Number',
      'PrimaryPhone',
      'Primary Phone',
      'Mobile'
    ]);

  const email =
    tmv3_first_(r, [
      'Email',
      'Email Address',
      'PrimaryEmail',
      'Primary Email'
    ]);

  return [
    id,
    no,
    name,
    phone,
    email,
    tmv3_hash_([id,no,name,phone,email].join('|'))
  ];
}

function tmv3_normalizeLocation_(r, customerNumberToId) {
  const id =
    tmv3_first_(r, [
      'LocationId',
      'Location ID',
      'Id'
    ]);

  let customerId =
    tmv3_first_(r, [
      'CustomerId',
      'Customer ID',
      'CustomerCustomerId',
      'AccountId',
      'Account ID'
    ]);

  if (!customerId) {
    const customerNumber = tmv3_first_(r, ['CustomerNumber','Customer Number','Customer #']);
    customerId = customerNumberToId && customerNumber
      ? (customerNumberToId[tmv3_clean_(customerNumber)] || '')
      : '';
  }

  const a1 =
    tmv3_first_(r, [
      'Address1',
      'Address 1',
      'AddressFullAddress',
      'FullAddress',
      'Street',
      'Street1',
      'Street 1'
    ]);

  const a2 =
    tmv3_first_(r, [
      'Address2',
      'Address 2',
      'Street2',
      'Street 2'
    ]);

  const city = tmv3_first_(r, ['City']);
  const province = tmv3_first_(r, ['State','Province','State/Province']);
  const postal = tmv3_first_(r, ['PostalCode','Postal Code','AddressZip','Zip','ZIP']);
  const phone = tmv3_first_(r, ['Phone','Phone Number']);

  return [
    id,
    customerId,
    a1,
    a2,
    city,
    province,
    postal,
    phone,
    tmv3_hash_(
      [
        id,
        customerId,
        a1,
        a2,
        city,
        province,
        postal,
        phone
      ].join('|')
    )
  ];
}

function tmv3_normalizeContact_(r) {
  const id =
    tmv3_first_(r, [
      'ContactId',
      'Contact ID',
      'Id'
    ]);

  const customerId =
    tmv3_first_(r, [
      'CustomerId',
      'Customer ID',
      'CustomerCustomerId',
      'AccountId',
      'Account ID'
    ]);

  const name =
    tmv3_first_(r, [
      'ContactName',
      'Contact Name',
      'Name',
      'FullName',
      'Full Name'
    ]);

  const phone =
    tmv3_first_(r, [
      'Phone',
      'Phone Number',
      'Mobile'
    ]);

  const email =
    tmv3_first_(r, [
      'Email',
      'Email Address'
    ]);

  return [
    id,
    customerId,
    name,
    phone,
    email,
    tmv3_hash_(
      [
        id,
        customerId,
        name,
        phone,
        email
      ].join('|')
    )
  ];
}

function tmv3_normalizeOrder_(r, type) {
  const id =
    tmv3_first_(r, [
      'SalesOrderId',
      'Sales Order ID',
      'SOId',
      'SO ID',
      'OrderId',
      'Order ID',
      'Id'
    ]);

  const no =
    tmv3_first_(r, [
      'SONumber',
      'SO Number',
      'SalesOrderNumber',
      'Sales Order Number',
      'OrderNumber',
      'Order Number',
      'SO #',
      'Service Order Number'
    ]);

  const customerId =
    tmv3_first_(r, [
      'CustomerId',
      'Customer ID',
      'CustomerCustomerId',
      'AccountId',
      'Account ID'
    ]);

  const locationId =
    tmv3_first_(r, [
      'LocationId',
      'Location ID',
      'CustomerAddressAddressId',
      'ShipToAddressAddressId',
      'CustomerAddressId'
    ]);

  const contactId =
    tmv3_first_(r, [
      'ContactId',
      'Contact ID',
      'ContactContactId'
    ]);

  const status =
    tmv3_first_(r, [
      'Status',
      'SOStatus',
      'SalesOrderStatus',
      'Order Status'
    ]);

  const name =
    tmv3_first_(r, [
      'SalesOrderName',
      'Sales Order Name',
      'SOName',
      'OrderName',
      'Order Name',
      'Name',
      'Description'
    ]);

  const url =
    id
      ? TMV3.ORDER_URL_BASE + encodeURIComponent(id)
      : '';

  return [
    id,
    no,
    customerId,
    locationId,
    contactId,
    status,
    name,
    type,
    url,
    tmv3_hash_(
      [
        id,
        no,
        customerId,
        locationId,
        contactId,
        status,
        name,
        type
      ].join('|')
    )
  ];
}

function tmv3_normalizeTask_(r, vertical, orderIdByNumber) {
  const id =
    tmv3_first_(r, [
      'TaskId',
      'Task ID',
      'Id',
      'ServiceTaskId',
      'DeliveryTaskId'
    ]);

  const no =
    tmv3_first_(r, [
      'TaskNumber',
      'Task Number',
      'Task #'
    ]) || id;

  const typeId =
    tmv3_first_(r, [
      'TaskTypeId',
      'Task Type ID',
      'TypeId',
      'Type ID'
    ]);

  const typeName =
    tmv3_first_(r, [
      'TaskType',
      'Task Type',
      'TaskTypeName',
      'Task Type Name',
      'Type'
    ]);

  const status =
    tmv3_first_(r, [
      'TaskStatus',
      'Task Status',
      'Status'
    ]);

  const name =
    tmv3_first_(r, [
      'TaskName',
      'Task Name',
      'Name',
      'Title',
      'Subject'
    ]);

  const customerId =
    tmv3_first_(r, [
      'CustomerId',
      'Customer ID',
      'AccountId',
      'Account ID'
    ]);

  const locationId =
    tmv3_first_(r, [
      'LocationId',
      'Location ID',
      'LocationLocationId',
      'CustomerLocationId'
    ]);

  const contactId =
    tmv3_first_(r, [
      'ContactId',
      'Contact ID'
    ]);

  let orderId =
    tmv3_first_(r, [
      'SalesOrderId',
      'Sales Order ID',
      'SOId',
      'SO ID',
      'OrderId',
      'Order ID'
    ]);

  if (!orderId) {
    const orderNumber = tmv3_first_(r, [
      'SalesOrderNumber',
      'Sales Order Number',
      'SONumber',
      'SO Number',
      'SO #',
      'OrderNumber',
      'Order Number'
    ]);
    if (orderNumber && orderIdByNumber) {
      orderId = orderIdByNumber[tmv3_clean_(orderNumber)] || '';
    }
  }

  const start =
    tmv3_first_(r, [
      'StartDateTime',
      'Start Date Time',
      'StartDate',
      'Start Date',
      'TaskStartDate'
    ]);

  const due =
    tmv3_first_(r, [
      'DueDateTime',
      'Due Date Time',
      'DueDate',
      'Due Date',
      'TaskDueDate'
    ]);

  const assignees =
    tmv3_first_(r, [
      'AssignedTo',
      'Assigned To',
      'Assignees',
      'Employees'
    ]);

  const pools =
    tmv3_first_(r, [
      'Pools',
      'AssignedPools',
      'Assigned Pools',
      'Pool'
    ]);

  const url =
    id
      ? TMV3.TASK_URL_BASE + encodeURIComponent(id)
      : '';

  return [
    id,
    no,
    typeId,
    typeName || vertical,
    status,
    name,
    customerId,
    locationId,
    contactId,
    orderId,
    start,
    due,
    assignees,
    pools,
    url,
    tmv3_hash_(
      [
        id,
        no,
        typeId,
        typeName,
        status,
        name,
        customerId,
        locationId,
        contactId,
        orderId,
        start,
        due,
        assignees,
        pools,
        vertical
      ].join('|')
    )
  ];
}

function tmv3_dedupeObjects_(rows, keyFn) {
  const map = {};

  (rows || []).forEach(function(r) {
    const key = tmv3_clean_(keyFn(r));
    if (key) map[key] = r;
  });

  return Object.keys(map).map(function(k) {
    return map[k];
  });
}


function tmv3_mergeOrderRows_(rows) {
  const map = {};

  (rows || []).forEach(function(r) {
    const key = tmv3_clean_(r[0] || r[1]);
    if (!key) return;

    if (!map[key]) {
      map[key] = r.slice();
      return;
    }

    const existing = map[key];

    // Preserve the richest nonblank relationship data from any approved-order feed.
    [0,1,2,3,4,5,6,8].forEach(function(i) {
      if (!tmv3_clean_(existing[i]) && tmv3_clean_(r[i])) existing[i] = r[i];
    });

    const types = tmv3_unique_(
      tmv3_clean_(existing[7]).split('|').concat(tmv3_clean_(r[7]).split('|'))
    );
    existing[7] = types.join('|');

    existing[9] = tmv3_hash_(existing.slice(0, 9).join('|'));
  });

  return Object.keys(map).map(function(k) { return map[k]; });
}


function tmv3_getTaskById_(taskId) {
  const id = Number(taskId || 0);
  if (!id) throw new Error('A positive Task ID is required.');

  const raw = tmv3_fetchJson_(
    TMV3.API_BASE + '/v2/tasks/' + encodeURIComponent(id),
    { method: 'get' }
  );

  return tmv3_normalizeV2TaskModel_(raw || {});
}


/**
 * Canonical Task schedule fallback.
 *
 * Striven v2 Task start/due values have a proven PM representation defect.
 * The legacy R4.6 workflow verified /v1/tasks/{id} DesiredStartDate /
 * DesiredEndDate as the canonical schedule model. V3 calls this only when
 * the v2 schedule disagrees with Calendar.
 */
function tmv3_getCanonicalV1TaskSchedule_(taskId) {
  const id = Number(taskId || 0);
  if (!id) throw new Error('A positive Task ID is required for canonical v1 schedule read.');

  const raw = tmv3_fetchJson_(
    TMV3.API_BASE + '/v1/tasks/' + encodeURIComponent(id),
    { method: 'get' }
  ) || {};

  const start =
    raw.desiredStartDate !== undefined ? raw.desiredStartDate :
    (raw.DesiredStartDate !== undefined ? raw.DesiredStartDate :
    (raw.startDate !== undefined ? raw.startDate : raw.StartDate));

  const due =
    raw.desiredEndDate !== undefined ? raw.desiredEndDate :
    (raw.DesiredEndDate !== undefined ? raw.DesiredEndDate :
    (raw.dueDate !== undefined ? raw.dueDate : raw.DueDate));

  return {
    taskId: id,
    startDateTime: start || null,
    dueDateTime: due || null,
    source: 'V1_DESIRED_START_END'
  };
}

function tmv3_taskScheduleCheck_(vertical, calendarStart, calendarEnd, taskId, v2Task) {
  const task = v2Task || {};
  const dueDateOnly = tmv3_taskDueDateOnly_(vertical, task);
  const v2Start = task['Start'];
  const v2Due = task['Due'];

  const v2StartMatch = tmv3_sameMinute_(calendarStart, v2Start);
  const v2DueMatch = dueDateOnly
    ? tmv3_sameCalendarDate_(calendarEnd, v2Due)
    : tmv3_sameMinute_(calendarEnd, v2Due);

  if (v2StartMatch && v2DueMatch) {
    return {
      startCheck: 'MATCH',
      endCheck: 'MATCH',
      start: v2Start,
      due: v2Due,
      source: 'V2_TASK_MODEL',
      canonicalRead: 'NOT_NEEDED'
    };
  }

  let canonical;
  try {
    canonical = tmv3_getCanonicalV1TaskSchedule_(taskId);
  } catch (err) {
    return {
      startCheck: v2StartMatch ? 'MATCH' : 'CANONICAL_READ_FAILED',
      endCheck: v2DueMatch ? 'MATCH' : 'CANONICAL_READ_FAILED',
      start: v2Start,
      due: v2Due,
      source: 'V2_MISMATCH_V1_READ_FAILED',
      canonicalRead: 'FAILED',
      error: String(err && err.message || err)
    };
  }

  if (!canonical.startDateTime || !canonical.dueDateTime) {
    return {
      startCheck: v2StartMatch ? 'MATCH' : 'CANONICAL_READ_FAILED',
      endCheck: v2DueMatch ? 'MATCH' : 'CANONICAL_READ_FAILED',
      start: v2Start,
      due: v2Due,
      source: 'V2_MISMATCH_V1_FIELDS_MISSING',
      canonicalRead: 'FIELDS_MISSING',
      error: 'Canonical v1 Task response did not expose DesiredStartDate and DesiredEndDate.'
    };
  }

  const canonicalStartMatch =
    tmv3_sameMinute_(calendarStart, canonical.startDateTime);
  const canonicalDueMatch = dueDateOnly
    ? tmv3_sameCalendarDate_(calendarEnd, canonical.dueDateTime)
    : tmv3_sameMinute_(calendarEnd, canonical.dueDateTime);

  return {
    startCheck: canonicalStartMatch ? 'MATCH' : 'MISMATCH',
    endCheck: canonicalDueMatch ? 'MATCH' : 'MISMATCH',
    start: canonical.startDateTime,
    due: canonical.dueDateTime,
    source: canonical.source || 'V1_DESIRED_START_END',
    canonicalRead: 'VERIFIED',
    v2: {
      start: v2Start || '',
      due: v2Due || ''
    }
  };
}



function tmv3_extractTaskSearchRows_(json) {
  if (!json) return [];
  if (Array.isArray(json)) return json;

  const direct = [
    json.data, json.Data, json.items, json.Items,
    json.results, json.Results, json.tasks, json.Tasks
  ];

  for (let i = 0; i < direct.length; i++) {
    if (Array.isArray(direct[i])) return direct[i];
  }

  const nested = [json.data, json.Data, json.result, json.Result];

  for (let j = 0; j < nested.length; j++) {
    const obj = nested[j];
    if (!obj || typeof obj !== 'object') continue;

    const arrays = [
      obj.items, obj.Items, obj.results, obj.Results,
      obj.tasks, obj.Tasks, obj.rows, obj.Rows, obj.data, obj.Data
    ];

    for (let k = 0; k < arrays.length; k++) {
      if (Array.isArray(arrays[k])) return arrays[k];
    }
  }

  return [];
}

function tmv3_normalizeV2TaskModel_(raw) {
  raw = raw || {};
  const type = raw.type || raw.Type || {};
  const status = raw.status || raw.Status || {};
  const customer = raw.customer || raw.Customer || {};
  const location = raw.location || raw.Location || {};
  const salesOrder = raw.salesOrder || raw.SalesOrder || {};
  const requestedBy = raw.requestedBy || raw.RequestedBy || {};
  const requestedByContact = raw.requestedByContact || raw.RequestedByContact || {};
  const assignments = raw.assignments || raw.Assignments || [];

  const taskId = tmv3_first_(raw, ['id','Id','taskId','TaskId']);
  const taskNumber = tmv3_first_(raw, ['number','Number','taskNumber','TaskNumber']) || taskId;
  const typeId = tmv3_first_(type, ['id','Id']);
  const typeName = tmv3_first_(type, ['name','Name']);
  const statusName = tmv3_first_(status, ['name','Name']);
  const customerId = tmv3_first_(customer, ['id','Id']);
  const locationId = tmv3_first_(location, ['id','Id']);
  const orderId = tmv3_first_(salesOrder, ['id','Id']);
  const requestedById =
    tmv3_first_(requestedBy, ['id','Id']) ||
    tmv3_first_(requestedByContact, ['contactId','ContactId','id','Id']) ||
    tmv3_first_(raw, ['contactId','ContactId','requestedByContactId','RequestedByContactId']);

  const employeeNames = [];
  const poolNames = [];

  (assignments || []).forEach(function(a) {
    const kind = tmv3_norm_(tmv3_first_(a, ['type','Type']));
    const name = tmv3_clean_(tmv3_first_(a, ['name','Name']));
    const id = tmv3_clean_(tmv3_first_(a, ['id','Id']));
    if (kind === 'pool') poolNames.push(name || id);
    else if (name || id) employeeNames.push(name || id);
  });

  const start = tmv3_first_(raw, ['startDateTime','StartDateTime','startDate','StartDate']);
  const due = tmv3_first_(raw, ['dueDateTime','DueDateTime','dueDate','DueDate']);
  // Striven returns HTML entities in task names (for example &amp;).
  // Decode once at the read boundary so verified display text compares with
  // the submitted canonical name without repeated PATCH attempts.
  const name = tmv3_decodeCalendarHtmlEntities_(
    tmv3_first_(raw, ['title','Title','name','Name']) || ''
  );

  const row = {
    'Task ID': taskId,
    'Task Number': taskNumber,
    'Task Type ID': typeId,
    'Task Type': typeName,
    'Status': statusName,
    'Name': name,
    'Customer ID': customerId,
    'Location ID': locationId,
    'Contact ID': requestedById,
    'Order ID': orderId,
    'Start': start,
    'Due': due,
    'Assignees': employeeNames.join(', '),
    'Pools': poolNames.join(', '),
    'URL': taskId ? TMV3.TASK_URL_BASE + encodeURIComponent(taskId) : ''
  };

  row.Fingerprint = tmv3_hash_([
    taskId, taskNumber, typeId, typeName, statusName, name, customerId,
    locationId, requestedById, orderId, start, due,
    row.Assignees, row.Pools
  ].join('|'));

  return row;
}


function tmv3_employeeDirectory_() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'TMV3_EMPLOYEE_DIRECTORY';
  const cached = cache.get(cacheKey);

  if (cached) {
    try { return JSON.parse(cached); } catch (ignored) {}
  }

  const json = tmv3_fetchJson_(
    TMV3.API_BASE + '/v1/employees',
    { method: 'get' }
  );

  const rows = Array.isArray(json)
    ? json
    : (
        (json && (json.Data || json.data || json.Items || json.items)) ||
        []
      );

  const employees = rows.map(function(row) {
    return {
      id: Number(
        row.Id ||
        row.id ||
        row.EmployeeId ||
        row.employeeId ||
        0
      ),
      name: tmv3_clean_(
        row.Name ||
        row.name ||
        row.EmployeeName ||
        row.employeeName
      ),
      email: tmv3_normEmail_(
        row.Email ||
        row.email ||
        row.EmailAddress ||
        row.emailAddress
      )
    };
  }).filter(function(employee) {
    return employee.id > 0 && employee.email;
  });

  try {
    cache.put(cacheKey, JSON.stringify(employees), 21600);
  } catch (ignored) {}

  return employees;
}

function tmv3_resolveOrganizerEmployee_(eventRecord) {
  const emails = tmv3_extractEmails_(eventRecord && eventRecord.organizer);

  if (emails.length !== 1) {
    return {
      status: 'REVIEW',
      errorCode: 'ORGANIZER_EMAIL_NOT_UNIQUE',
      reason:
        'Expected exactly one Calendar organizer email; found ' +
        emails.length +
        '.'
    };
  }

  const organizerEmail = emails[0];
  const matches = tmv3_employeeDirectory_().filter(function(employee) {
    return employee.email === organizerEmail;
  });

  if (matches.length !== 1) {
    return {
      status: 'REVIEW',
      errorCode: 'ORGANIZER_EMPLOYEE_NOT_UNIQUE',
      reason:
        'Calendar organizer ' +
        organizerEmail +
        ' resolved to ' +
        matches.length +
        ' Striven Employees.'
    };
  }

  return {
    status: 'MATCHED',
    employee: matches[0],
    evidence: 'CALENDAR_ORGANIZER_EMAIL_EXACT_EMPLOYEE'
  };
}


function tmv3_getCustomerContacts_(customerId, options) {
  const id = Number(customerId || 0);
  if (!id) throw new Error('Customer ID is required for contact lookup.');

  options = options || {};
  const forceFresh = options.forceFresh === true;
  const cache = CacheService.getScriptCache();
  const cacheKey = 'TMV3_CUSTOMER_CONTACTS_' + id;
  const cached = forceFresh ? null : cache.get(cacheKey);

  if (cached) {
    try { return JSON.parse(cached); } catch (ignored) {}
  }

  const raw = tmv3_fetchJson_(
    TMV3.API_BASE + '/v1/customers/' + encodeURIComponent(id) + '/contacts',
    { method: 'get' }
  );

  const rows = tmv3_arrayFromApiResponse_(raw);
  if(options.requireComplete===true && !Array.isArray(raw)) {
    const total=raw && (raw.TotalCount!==undefined?raw.TotalCount:raw.totalCount);
    if(!Number.isInteger(total)||total!==rows.length)throw Error('CONTACT_COLLECTION_INCOMPLETE');
  }
  const contacts = rows
    .map(function(item) {
      return tmv3_normalizeApiContact_(item, id);
    })
    .filter(function(item) {
      return !!item['Contact ID'];
    });

  if (!forceFresh) {
    try {
      cache.put(
        cacheKey,
        JSON.stringify(contacts),
        Math.max(
          1800,
          Number(
            TMV3.OPERATIONS &&
            TMV3.OPERATIONS.customerContactsCacheSeconds ||
            21600
          )
        )
      );
    } catch (ignored) {}
  }

  return contacts;
}

function tmv3_getContactById_(contactId, expectedCustomerId) {
  const id = Number(contactId || 0);
  if (!id) throw new Error('Contact ID is required.');

  const raw = tmv3_fetchJson_(
    TMV3.API_BASE + '/v1/contacts/' + encodeURIComponent(id),
    { method:'get' }
  ) || {};

  // Fresh ownership must come from fields actually returned by Striven.
  // Never backfill the expected Customer ID into the model before proving it.
  const contact = tmv3_normalizeApiContact_(raw, '');
  const expected = String(
    Number(expectedCustomerId || 0) || ''
  );

  const directCustomerId = tmv3_clean_(
    contact['Customer ID']
  );

  const ownerObject =
    raw.customer ||
    raw.Customer ||
    raw.account ||
    raw.Account ||
    {};

  const nestedOwnerId = tmv3_clean_(
    tmv3_first_(ownerObject, [
      'id','Id','customerId','CustomerId',
      'accountId','AccountId'
    ])
  );

  let scopedCustomerEvidence = false;
  let scopedLookupError = '';

  if (
    expected &&
    directCustomerId !== expected &&
    nestedOwnerId !== expected
  ) {
    try {
      scopedCustomerEvidence = tmv3_getCustomerContacts_(
        Number(expected),
        { forceFresh:true }
      ).some(function(item) {
        return String(item['Contact ID'] || '') === String(id);
      });
    } catch (err) {
      scopedLookupError = String(err && err.message || err);
    }
  }

  contact.__ownershipVerified =
    !!expected &&
    (
      directCustomerId === expected ||
      nestedOwnerId === expected ||
      scopedCustomerEvidence
    );

  contact.__ownershipEvidence =
    contact.__ownershipVerified
      ? (
          directCustomerId === expected
            ? 'CONTACT_DIRECT_CUSTOMER_ID'
            : (
                nestedOwnerId === expected
                  ? 'CONTACT_NESTED_OWNER_ID'
                  : 'CONTACT_CUSTOMER_SCOPED_LIST'
              )
        )
      : 'CONTACT_OWNER_NOT_PROVEN';

  if (scopedLookupError) {
    contact.__ownershipLookupError = scopedLookupError;
  }

  return contact;
}

function tmv3_normalizeApiContact_(raw, fallbackCustomerId) {
  raw = raw || {};

  const id = tmv3_first_(raw, ['id','Id','contactId','ContactId']);
  const first = tmv3_clean_(tmv3_first_(raw, ['firstName','FirstName']));
  const last = tmv3_clean_(tmv3_first_(raw, ['lastName','LastName']));
  const name =
    tmv3_clean_(tmv3_first_(raw, ['name','Name','fullName','FullName'])) ||
    [first,last].filter(Boolean).join(' ');

  const customerId =
    tmv3_first_(raw, ['customerId','CustomerId','accountId','AccountId']) ||
    tmv3_first_(raw.customer || raw.Customer || {}, ['id','Id']) ||
    fallbackCustomerId ||
    '';

  const text = JSON.stringify(raw);
  const phones = tmv3_allPhones_(text);
  const emails = tmv3_extractEmails_(text);

  return {
    'Contact ID': tmv3_clean_(id),
    'Customer ID': tmv3_clean_(customerId),
    'Name': name,
    'Phone': phones[0] || '',
    'Email': emails[0] || '',
    Phones: phones,
    Emails: emails
  };
}

function tmv3_arrayFromApiResponse_(json) {
  if (!json) return [];
  if (Array.isArray(json)) return json;

  const direct = [
    json.data,
    json.Data,
    json.items,
    json.Items,
    json.results,
    json.Results,
    json.contacts,
    json.Contacts
  ];

  for (let i = 0; i < direct.length; i++) {
    if (Array.isArray(direct[i])) return direct[i];
  }

  return [];
}


// Original: 30_Calendar
function tmv3_calendarRecords_() {
  return tmv3_calendarRecordsWithOptions_({
    stage: 'RUNTIME',
    applyEligibility: true,
    mergeLogical: false
  });
}

function tmv3_step1CalendarRecords_() {
  return tmv3_calendarRecordsWithOptions_({
    stage: 'STEP1',
    applyEligibility: false,
    mergeLogical: true
  });
}

function tmv3_calendarRecordsWithOptions_(options) {
  tmv3_assertShadow_();
  options = options || {};
  const out = [];

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    const cfg = TMV3.VERTICALS[vertical];
    const calendars = tmv3_verticalCalendars_(vertical, cfg);

    const start = new Date();
    start.setDate(start.getDate() - Number(cfg.lookbackDays || 0));
    start.setHours(0,0,0,0);

    const end = new Date();
    end.setDate(end.getDate() + Number(cfg.lookaheadDays || 365));
    end.setHours(23,59,59,999);

    calendars.forEach(function(calCfg) {
      const cal = CalendarApp.getCalendarById(calCfg.calendarId);

      if (!cal) {
        tmv3_audit_(
          vertical,'','','CALENDAR_ACCESS','BLOCKED',
          'Calendar unavailable: ' + calCfg.calendarId
        );
        return;
      }

      const resolvedCalCfg = Object.assign({}, calCfg, {
        calendarName: tmv3_safeCalendar_(function() {
          return cal.getName ? cal.getName() : '';
        }, '')
      });

      cal.getEvents(start, end).forEach(function(event) {
        const record = tmv3_calendarEvent_(vertical, cfg, resolvedCalCfg, event, options);
        if (record) out.push(record);
      });
    });
  });

  return options.mergeLogical === true
    ? tmv3_mergeLogicalCalendarRecords_(out)
    : out;
}

function tmv3_verticalCalendars_(vertical, cfg) {
  if (vertical === 'Service') {
    return (cfg.calendars || []).map(function(item) {
      return Object.assign({}, item, { role: 'TECHNICIAN' });
    });
  }

  if (vertical === 'PreInspection') {
    // CF Preinspects is the single authoritative appointment source.
    // Stephen receives the same event as a guest; his Calendar is mirror-only
    // and must never introduce an independent PreInspection record.
    return [
      {
        calendarId: cfg.primaryCalendarId,
        role: 'PRIMARY_SHARED',
        calendarName: 'CF Preinspects'
      }
    ];
  }

  return (cfg.calendarIds || []).map(function(id) {
    return {
      calendarId: id,
      role: vertical.toUpperCase()
    };
  });
}

function tmv3_calendarEvent_(vertical, cfg, calCfg, event, options) {
  options = options || {};

  const title = tmv3_clean_(event.getTitle());
  const eventId = event.getId();
  const rawDescription = String(event.getDescription() || '');
  const descriptionClean = tmv3_calendarReadableDescription_(rawDescription);
  const titleMigration = tmv3_titleMigrationStateForEvent_(
    vertical,
    eventId,
    calCfg.calendarId
  );
  const identityRawDescription =
    tmv3_identityEvidenceDescriptionRaw_(
      rawDescription,
      titleMigration.preservedTitle
    );
  const identityDescriptionClean =
    tmv3_calendarReadableDescription_(identityRawDescription);
  const assignmentSourceTitle =
    tmv3_clean_(titleMigration.preservedTitle) || title;
  const location = tmv3_clean_(event.getLocation());

  let creator = tmv3_safeCalendar_(
    function() {
      return event.getCreators ? (event.getCreators() || []).join(',') : '';
    },
    ''
  );

  creator=tmv3_preInspectionProvenanceCreator_(vertical,cfg,calCfg,event,creator);

  const guests = tmv3_safeCalendar_(
    function() {
      return event.getGuestList().map(function(g) {
        return g.getEmail();
      }).join(',');
    },
    ''
  );

  const organizerCalendarId = tmv3_safeCalendar_(
    function() {
      return event.getOriginalCalendarId ? event.getOriginalCalendarId() : '';
    },
    ''
  );

  const isAllDay = event.isAllDayEvent();

  if (
    options.stage === 'STEP1' &&
    vertical === 'PreInspection' &&
    !tmv3_preInspectionStep1SourceAllowed_(cfg, calCfg, creator)
  ) {
    return null;
  }

  if (
    options.applyEligibility !== false &&
    tmv3_shouldIgnoreEvent_(vertical, title, rawDescription, isAllDay, creator)
  ) {
    return null;
  }

  const start = event.getStartTime();
  const end = event.getEndTime();
  const links = tmv3_extractLinks_(rawDescription);

  const orderNumber = tmv3_extractOrderNumberForVertical_(
    vertical,
    title,
    identityDescriptionClean,
    location
  );

  const customerNumber = tmv3_extractCustomerNumberForVertical_(
    vertical,
    title + ' ' + identityDescriptionClean
  );

  const phone = tmv3_extractPhone_(
    title + ' ' + identityDescriptionClean + ' ' + location
  );

  const taskNumber = tmv3_extractCalendarTaskNumber_(
    title + ' ' + identityDescriptionClean
  );

  const serviceLegacyAllowed =
    vertical === 'Service'
      ? tmv3_serviceStep1LegacyAllowed_(title, isAllDay, creator)
      : true;

  const serviceExternalRecovery =
    vertical === 'Service' &&
    !serviceLegacyAllowed &&
    tmv3_serviceExternalCreatorRecoveryAllowed_({
      title: title,
      description: identityDescriptionClean,
      location: location,
      isAllDay: isAllDay,
      creator: creator,
      orderNumber: orderNumber,
      phone: phone,
      taskNumber: taskNumber,
      existingTaskId: links.taskId,
      existingOrderId: links.orderId
    });

  if (
    options.stage === 'STEP1' &&
    vertical === 'Service' &&
    !serviceLegacyAllowed &&
    !serviceExternalRecovery
  ) {
    return null;
  }

  const calendarCustomerName = tmv3_extractCalendarCustomerName_(
    vertical, title, customerNumber, phone, orderNumber, calCfg
  );

  const fingerprint = tmv3_hash_(
    [
      vertical, calCfg.calendarId, eventId, title, location,
      tmv3_iso_(start), tmv3_iso_(end), rawDescription, creator, guests
    ].join('|')
  );

  return {
    vertical: vertical,
    calendarId: calCfg.calendarId,
    calendarName: calCfg.calendarName || '',
    calendarRole: calCfg.role || '',
    technician: calCfg.technician || '',
    technicianEmployeeId: calCfg.employeeId || '',
    eventId: eventId,
    title: title,
    description: rawDescription,
    descriptionClean: descriptionClean,
    identityDescriptionClean: identityDescriptionClean,
    rawDescription: rawDescription,
    preservedCalendarTitle: tmv3_clean_(titleMigration.preservedTitle),
    assignmentSourceTitle: assignmentSourceTitle,
    location: location,
    start: start,
    end: end,
    isAllDay: isAllDay,
    organizer: creator,
    creator: creator,
    organizerCalendarId: organizerCalendarId,
    guests: guests,
    orderNumber: orderNumber,
    customerNumber: customerNumber,
    calendarCustomerName: calendarCustomerName,
    phone: phone,
    taskNumber: taskNumber,
    serviceLegacyAllowed: serviceLegacyAllowed,
    serviceExternalRecovery: serviceExternalRecovery,
    existingTaskId: links.taskId,
    existingTaskUrl: links.taskUrl,
    existingOrderId: links.orderId,
    existingOrderUrl: links.orderUrl,
    eventUrl: tmv3_calendarEventUrl_(eventId, calCfg.calendarId),
    fingerprint: fingerprint,
    logicalKey: tmv3_calendarLogicalKey_(eventId, start, title, location),
    sourceCalendarIds: [calCfg.calendarId],
    sourceCalendarNames: [calCfg.calendarName || calCfg.calendarId],
    sourceCalendarRoles: [calCfg.role || ''],
    sourceCreators: creator ? [creator] : [],
    sourceGuests: guests ? [guests] : [],
    duplicateSourceCount: 1,
    calendarCreatedAt: tmv3_safeCalendar_(
      function() { return tmv3_iso_(event.getDateCreated()); }, ''),
    calendarUpdatedAt: tmv3_safeCalendar_(
      function() { return tmv3_iso_(event.getLastUpdated()); },
      ''
    )
  };
}



function tmv3_mergeLogicalCalendarRecords_(records) {
  const byKey = {};
  const order = [];

  (records || []).forEach(function(record) {
    const key = record.logicalKey || tmv3_calendarLogicalKey_(
      record.eventId, record.start, record.title, record.location
    );

    if (!byKey[key]) {
      byKey[key] = Object.assign({}, record);
      order.push(key);
      return;
    }

    const base = byKey[key];

    base.sourceCalendarIds = tmv3_unique_(
      (base.sourceCalendarIds || []).concat(record.sourceCalendarIds || [record.calendarId])
    );
    base.sourceCalendarNames = tmv3_unique_(
      (base.sourceCalendarNames || []).concat(record.sourceCalendarNames || [record.calendarName || record.calendarId])
    );
    base.sourceCalendarRoles = tmv3_unique_(
      (base.sourceCalendarRoles || []).concat(record.sourceCalendarRoles || [record.calendarRole || ''])
    );
    base.sourceCreators = tmv3_unique_(
      (base.sourceCreators || []).concat(record.sourceCreators || (record.creator ? [record.creator] : []))
    );
    base.sourceGuests = tmv3_unique_(
      (base.sourceGuests || []).concat(record.sourceGuests || (record.guests ? [record.guests] : []))
    );

    base.duplicateSourceCount = Number(base.duplicateSourceCount || 1) + 1;

    if (!base.existingTaskId && record.existingTaskId) {
      base.existingTaskId = record.existingTaskId;
      base.existingTaskUrl = record.existingTaskUrl;
    }
    if (!base.existingOrderId && record.existingOrderId) {
      base.existingOrderId = record.existingOrderId;
      base.existingOrderUrl = record.existingOrderUrl;
    }
    if (!base.orderNumber && record.orderNumber) base.orderNumber = record.orderNumber;
    if (!base.customerNumber && record.customerNumber) base.customerNumber = record.customerNumber;
    if (!base.phone && record.phone) base.phone = record.phone;
    if (!base.calendarCustomerName && record.calendarCustomerName) {
      base.calendarCustomerName = record.calendarCustomerName;
    }

    base.guests = tmv3_unique_([base.guests, record.guests].filter(Boolean)).join(',');
    base.creator = tmv3_unique_([base.creator, record.creator].filter(Boolean)).join(',');
    base.organizer = base.creator;
    base.fingerprint = tmv3_hash_(
      [
        base.fingerprint, record.fingerprint,
        base.sourceCalendarIds.join(','), base.sourceCalendarRoles.join(',')
      ].join('|')
    );
  });

  return order.map(function(key) { return byKey[key]; });
}

function tmv3_calendarLogicalKey_(eventId, start, title, location) {
  const id = tmv3_norm_(eventId).replace(/@google\.com$/i, '');
  const startIso = start ? tmv3_iso_(start) : '';

  if (id) return id + '|' + startIso;

  return tmv3_hash_([
    startIso, tmv3_norm_(title), tmv3_norm_(location)
  ].join('|'));
}

function tmv3_calendarEventUrl_(eventId, calendarId) {
  const id = tmv3_clean_(eventId).replace(/@google\.com$/i, '');
  const calendar = tmv3_clean_(calendarId).replace(/@group\.calendar\.google\.com$/i, '@g');
  if (!id || !calendar) return '';
  const eid = Utilities.base64EncodeWebSafe(id + ' ' + calendar).replace(/=+$/, '');
  return 'https://www.google.com/calendar/event?eid=' + eid;
}


function tmv3_identityEvidenceDescriptionRaw_(rawDescription, preservedTitle) {
  let text = String(
    rawDescription === null || rawDescription === undefined
      ? ''
      : rawDescription
  );

  const oldTitle = tmv3_clean_(preservedTitle);

  if (oldTitle) {
    const lfPrefix = oldTitle + '\n\n';
    const crlfPrefix = oldTitle + '\r\n\r\n';

    if (text.indexOf(lfPrefix) === 0) {
      text = text.slice(lfPrefix.length);
    } else if (text.indexOf(crlfPrefix) === 0) {
      text = text.slice(crlfPrefix.length);
    } else if (text === oldTitle) {
      text = '';
    }
  }

  // Older PreInspection standardization used a deterministic managed context
  // block. Ignore that block for identity evidence only. Never mutate it here.
  text = text.replace(
    /<!--\s*PREINSPECT_TITLE_CONTEXT_START\s*-->[\s\S]*?<!--\s*PREINSPECT_TITLE_CONTEXT_END\s*-->/gi,
    ' '
  );

  return text;
}

function tmv3_decodeCalendarHtmlEntities_(value) {
  let text = String(value === null || value === undefined ? '' : value);

  // Google Calendar can preserve already-HTML-formatted text as escaped HTML.
  // Decode before stripping tags so escaped markup becomes markup first.
  // Two bounded passes cover double-escaped entities such as &amp;#45;.
  for (let pass = 0; pass < 2; pass++) {
    const decoded = text
      .replace(/&amp;/gi, '&')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/&quot;/gi, '"')
      .replace(/&apos;/gi, "'")
      .replace(/&#39;/gi, "'")
      .replace(/&#x([0-9a-f]+);/gi, function(match, hex) {
        const codePoint = parseInt(hex, 16);
        return isNaN(codePoint) ? match : String.fromCodePoint(codePoint);
      })
      .replace(/&#(\d+);/g, function(match, decimal) {
        const codePoint = parseInt(decimal, 10);
        return isNaN(codePoint) ? match : String.fromCodePoint(codePoint);
      });

    if (decoded === text) break;
    text = decoded;
  }

  return text;
}

function tmv3_calendarReadableDescription_(value) {
  const decoded = tmv3_decodeCalendarHtmlEntities_(value);

  return tmv3_clean_(
    decoded
      .replace(/<\s*br\s*\/?>/gi, '\n')
      .replace(/<\s*\/p\s*>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/[ \t]+\n/g, '\n')
      .replace(/\n[ \t]+/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
  );
}





function tmv3_extractCalendarTaskNumber_(text) {
  const value = tmv3_clean_(text);
  if (!value) return '';

  const match = value.match(
    /\bTask\s*(?:#|No\.?|Number)?\s*[:\-]?\s*(\d{4,8})\b/i
  );

  return match ? match[1] : '';
}

function tmv3_extractOrderNumberForVertical_(vertical, title, description, location) {
  const titleText = tmv3_clean_(title);
  const descriptionText = tmv3_clean_(description);
  const locationText = tmv3_clean_(location);

  if (vertical === 'Install') {
    return tmv3_extractLegacyInstallOrderNumber_(titleText + ' ' + descriptionText);
  }

  if (vertical === 'Delivery') {
    return tmv3_extractLegacyDeliveryOrderNumber_(titleText + ' ' + descriptionText);
  }

  if (vertical === 'Service') {
    return tmv3_extractLegacyServiceOrderNumber_(
      [titleText, descriptionText, locationText].join(' ')
    );
  }

  if (vertical === 'PreInspection') {
    // Only labelled SO patterns are accepted here, so a leading Customer #
    // can never be mistaken for the Sales Order. Step 2 still separately
    // verifies that the SO is present in the description per the new format.
    return tmv3_extractLegacyPreInspectionOrderNumber_(
      titleText + ' ' + descriptionText
    );
  }

  return tmv3_extractOrderNumber_(
    [titleText, descriptionText, locationText].join(' ')
  );
}

function tmv3_stripPhonesForOrderParsing_(text) {
  return String(text || '')
    .replace(
      /(?:\+?1[\s\-.]?)?(?:\(?\d{3}\)?[\s\-.]?\d{3}[\s\-.]?\d{4})(?:\s*(?:x|ext\.?)\s*\d+)?/gi,
      ' '
    )
    .replace(/\s+/g, ' ')
    .trim();
}









function tmv3_extractCustomerNumberForVertical_(vertical, text) {
  const s = tmv3_clean_(text);

  const labelled = s.match(
    /(?:Customer|Cust(?:omer)?\s*#?|C\s*#)\s*(?:#|No\.?|Number)?\s*[:\-]?\s*(\d{3,8})/i
  );

  if (labelled) return labelled[1];

  if (vertical === 'PreInspection') {
    const leading = s.match(/^\s*#?(\d{4,8})\s*[-–—]/);
    return leading ? leading[1] : '';
  }

  return '';
}

function tmv3_extractCalendarCustomerName_(vertical, title, customerNumber, phone, orderNumber, calCfg) {
  if (vertical === 'Install') {
    return tmv3_installLegacyCustomerName_(title);
  }

  if (vertical === 'Delivery') {
    return tmv3_deliveryLegacyCustomerName_(title);
  }

  if (vertical === 'Service') {
    return tmv3_serviceLegacyCustomerName_(title);
  }

  return tmv3_preInspectionCalendarCustomerName_(
    title,
    customerNumber
  );
}









function tmv3_shouldIgnoreEvent_(vertical, title, description, isAllDay, organizer) {
  const t = tmv3_norm_(title);

  if (isAllDay) return true;

  if (
    vertical !== 'PreInspection' &&
    organizer &&
    organizer.indexOf(TMV3.DOMAIN) === -1
  ) return true;

  if (/\b(off|vacation|pto|sick|holiday|unavailable|blocked|personal)\b/i.test(t)) {
    return true;
  }

  if (vertical === 'Install' && /\bcrane job\b/i.test(t)) return true;

  if (
    vertical === 'Service' &&
    /\b(admin|meeting|lunch|training|warehouse|pickup|pick up)\b/i.test(t)
  ) return true;

  return false;
}

function tmv3_extractLinks_(html) {
  const text = String(html || '');
  const hrefs = [];
  const rx = /href=["']([^"']+)["']/gi;
  let m;

  while ((m = rx.exec(text)) !== null) hrefs.push(m[1]);

  const rawUrls = text.match(/https?:\/\/[^\s<>"']+/gi) || [];
  const urls = tmv3_unique_(hrefs.concat(rawUrls));

  let taskUrl = '';
  let taskId = '';
  let orderUrl = '';
  let orderId = '';

  urls.forEach(function(url) {
    let mm = String(url).match(/TaskInfo\.aspx\?[^"']*TaskID=(\d+)/i);
    if (mm && !taskUrl) {
      taskUrl = url;
      taskId = mm[1];
    }

    mm = String(url).match(/\/sales-orders\/(\d+)/i);
    if (mm && !orderUrl) {
      orderUrl = url;
      orderId = mm[1];
    }
  });

  return {
    taskUrl: taskUrl,
    taskId: taskId,
    orderUrl: orderUrl,
    orderId: orderId
  };
}

function tmv3_extractOrderNumber_(text) {
  const s = tmv3_clean_(text);
  const labelled = s.match(
    /(?:SO|Sales\s*Order|Work\s*Order|Service\s*Order)\s*(?:#|No\.?|Number)?\s*[:\-]?\s*(\d{4,8})/i
  );
  return labelled ? labelled[1] : '';
}

function tmv3_extractCustomerNumber_(text) {
  const s = tmv3_clean_(text);
  const m = s.match(
    /(?:Customer|Cust(?:omer)?\s*#?|C\s*#)\s*(?:#|No\.?|Number)?\s*[:\-]?\s*(\d{3,8})/i
  );

  if (m) return m[1];

  const leading = s.match(/^\s*#?(\d{4,8})\s*[-–—]/);
  return leading ? leading[1] : '';
}

function tmv3_extractPhone_(text) {
  const m = tmv3_clean_(text).match(
    /(?:\+?1[\s.\-]?)?\(?([2-9]\d{2})\)?[\s.\-]?([2-9]\d{2})[\s.\-]?(\d{4})/
  );

  return m ? (m[1] + m[2] + m[3]) : '';
}

function tmv3_safeCalendar_(fn, fallback) {
  try {
    return fn();
  } catch (err) {
    return fallback;
  }
}


// Preserve the verified original creator of an owner-confirmed imported copy.



// Original: 32_Step1_Calendar_View
const TMV3_STEP1_VISIBLE_COLUMN_COUNT = 11;

const TMV3_STEP1_HEADERS = Object.freeze([
  'Calendar Event',
  'Calendar Customer',
  'Description',
  'Created / Guests',
  'Data Checklist',
  'Task Customer / IDs (Evidence)',
  'Task',
  'Task Schedule',
  'Task Status',
  'Mapping Status',
  'Issue / Next Action',
  'Vertical',
  'Source Calendar',
  'Calendar Role',
  'Event ID',
  'Calendar Start',
  'Calendar End',
  'Calendar Location',
  'Creator',
  'Guests',
  'Customer # (Calendar)',
  'Customer Name (Calendar)',
  'Phone (Calendar)',
  'Order / Work Order (Calendar)',
  'Existing Task Link',
  'Existing Order Link',
  'Calendar Event Link',
  'Raw Description',
  'Calendar Updated',
  'Fingerprint',
  'Source Count'
]);

function tmv3_executionStage_() {
  return Number(
    typeof TMV3_EXECUTION_STAGE !== 'undefined'
      ? TMV3_EXECUTION_STAGE
      : 1
  );
}

function tmv3_step1CalendarRun(reason) {
  tmv3_assertShadow_();

  const access = tmv3_step1CalendarDiagnostics_();
  const records = tmv3_step1CalendarRecords_();
  const write = tmv3_step1WriteOperatorViews_(records);
  const verify = tmv3_step1VerifyOperatorViews_(records);

  const pass = access.blocked === 0 && verify.pass === true;

  const result = {
    version: TMV3.VERSION,
    executionStage: tmv3_executionStage_(),
    stage: 1,
    status: pass ? 'PASS' : 'REVIEW',
    mode: 'ONE_WAY_GOOGLE_CALENDAR_TO_V3_SHEETS',
    reason: tmv3_clean_(reason || 'MANUAL'),
    calendarWritesPerformed: false,
    strivenReadsPerformed: false,
    strivenWritesPerformed: false,
    calendars: access,
    logicalEvents: records.length,
    counts: tmv3_step1Counts_(records),
    write: write,
    readBack: verify,
    preInspectionMissingDetails:
      typeof tmv3_preInspectionMissingDetailPreviewFromRecords_ === 'function'
        ? tmv3_preInspectionMissingDetailPreviewFromRecords_(records)
        : { status:'UNAVAILABLE', checked:0, wouldNotify:0, previews:[] }
  };

  tmv3_audit_(
    'SYSTEM','','','STEP1_CALENDAR_INTAKE',
    pass ? 'PASS' : 'REVIEW',
    JSON.stringify(result)
  );

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function tmv3_step1CalendarDiagnostics_() {
  const rows = [];
  let blocked = 0;

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    const cfg = TMV3.VERTICALS[vertical];

    tmv3_verticalCalendars_(vertical, cfg).forEach(function(calCfg) {
      const cal = CalendarApp.getCalendarById(calCfg.calendarId);
      const accessible = !!cal;
      if (!accessible) blocked++;

      rows.push({
        vertical: vertical,
        calendarId: calCfg.calendarId,
        role: calCfg.role || '',
        technician: calCfg.technician || '',
        accessible: accessible,
        name: accessible
          ? tmv3_safeCalendar_(function() { return cal.getName(); }, '')
          : ''
      });
    });
  });

  return {
    expected: rows.length,
    accessible: rows.length - blocked,
    blocked: blocked,
    calendars: rows
  };
}

function tmv3_step1WriteOperatorViews_(records) {
  const written = {};

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    const verticalRecords = (records || [])
      .filter(function(record) { return record.vertical === vertical; })
      .sort(function(a, b) {
        const aTime = a.start instanceof Date ? a.start.getTime() : 0;
        const bTime = b.start instanceof Date ? b.start.getTime() : 0;
        if (aTime !== bTime) return aTime - bTime;

        if (vertical === 'Service') {
          const techCompare = tmv3_clean_(a.technician)
            .localeCompare(tmv3_clean_(b.technician));
          if (techCompare !== 0) return techCompare;
        }

        return tmv3_clean_(a.title).localeCompare(tmv3_clean_(b.title));
      });

    const rows = verticalRecords.map(tmv3_step1OperatorRow_);
    const sheetName = TMV3.VERTICALS[vertical].sheet;
    const priorLastRow = tmv3_sheet_(sheetName).getLastRow();

    tmv3_replaceRows_(sheetName, TMV3_STEP1_HEADERS.slice(), rows);
    tmv3_step1FormatOperatorView_(
      sheetName,
      vertical,
      verticalRecords,
      priorLastRow
    );
    written[vertical] = rows.length;
  });

  return written;
}

function tmv3_step1OperatorRow_(record) {
  return [
    tmv3_step1EventDetails_(record),
    tmv3_step1CalendarCustomerDetails_(record),
    record.descriptionClean || '',
    tmv3_step1People_(record),
    tmv3_step1CalendarChecklist_(record),
    'NOT RUN — STEP 1',
    'NOT RUN — STEP 1',
    'NOT RUN — STEP 1',
    'NOT RUN — STEP 1',
    'STEP 1 — CALENDAR INTAKE',
    'Calendar captured. Customer / task / Striven resolution intentionally not run.',
    record.vertical || '',
    tmv3_step1SourceLabel_(record),
    (record.sourceCalendarRoles || [record.calendarRole]).filter(Boolean).join(' + '),
    record.eventId || '',
    record.start ? tmv3_iso_(record.start) : '',
    record.end ? tmv3_iso_(record.end) : '',
    record.location || '',
    record.creator || '',
    record.guests || '',
    record.customerNumber || '',
    record.calendarCustomerName || '',
    record.phone || '',
    record.orderNumber || '',
    record.existingTaskUrl || '',
    record.existingOrderUrl || '',
    record.eventUrl || '',
    record.rawDescription || '',
    record.calendarUpdatedAt || '',
    record.fingerprint || '',
    Number(record.duplicateSourceCount || 1)
  ];
}

function tmv3_step1EventDetails_(record) {
  const lines = [];
  lines.push(record.title || '(Untitled Calendar event)');

  if (record.start) {
    lines.push(
      tmv3_step1CalendarScheduleDisplay_(
        record.start,
        record.end,
        record.isAllDay
      )
    );
  }

  if (record.location) lines.push('Location: ' + record.location);
  lines.push('Source: ' + tmv3_step1SourceLabel_(record));
  lines.push('Event ID: ' + (record.eventId || '—'));

  return lines.join('\n');
}

function tmv3_step1CalendarCustomerDetails_(record) {
  return [
    'Customer #: ' + (record.customerNumber || '—'),
    'Name: ' + (record.calendarCustomerName || '—'),
    'Phone: ' + tmv3_step1FormatPhone_(record.phone),
    (record.vertical === 'Service' ? 'Work Order / SO #: ' : 'SO #: ') +
      (record.orderNumber || '—')
  ].join('\n');
}

function tmv3_step1People_(record) {
  const creators = (record.sourceCreators || []).filter(Boolean).join(', ') ||
    record.creator || '—';

  const sourceIds = record.sourceCalendarIds || [];
  const sourceNames = record.sourceCalendarNames || [];
  const calendarNamesById = {};

  sourceIds.forEach(function(id, index) {
    const key = tmv3_norm_(id);
    const name = tmv3_clean_(sourceNames[index]);
    if (key && name) calendarNamesById[key] = name;
  });

  const rawGuests = []
    .concat(record.sourceGuests || [])
    .concat(record.guests || '')
    .join(',')
    .split(',')
    .map(function(value) { return tmv3_clean_(value); })
    .filter(Boolean);

  const guests = tmv3_unique_(
    rawGuests.map(function(value) {
      return calendarNamesById[tmv3_norm_(value)] || value;
    })
  ).join(', ') || '—';

  return [
    'Created by: ' + creators,
    'Guests: ' + guests
  ].join('\n');
}

function tmv3_step1CalendarScheduleDisplay_(startValue, endValue, isAllDay) {
  const start = tmv3_parseDateTime_(startValue);
  const end = tmv3_parseDateTime_(endValue);

  if (!start) return '';

  const dateLabel = Utilities.formatDate(
    start,
    TMV3_TIMEZONE,
    'MMM d, yyyy'
  );

  if (isAllDay) return dateLabel + ' (All day)';

  if (!end) {
    return dateLabel + ' (' +
      Utilities.formatDate(start, TMV3_TIMEZONE, 'h:mm a') +
      ')';
  }

  const startDay = Utilities.formatDate(start, TMV3_TIMEZONE, 'yyyy-MM-dd');
  const endDay = Utilities.formatDate(end, TMV3_TIMEZONE, 'yyyy-MM-dd');

  if (startDay === endDay) {
    return (
      dateLabel +
      ' (' +
      Utilities.formatDate(start, TMV3_TIMEZONE, 'h:mm a') +
      ' - ' +
      Utilities.formatDate(end, TMV3_TIMEZONE, 'h:mm a') +
      ')'
    );
  }

  return (
    Utilities.formatDate(start, TMV3_TIMEZONE, 'MMM d, yyyy h:mm a') +
    ' - ' +
    Utilities.formatDate(end, TMV3_TIMEZONE, 'MMM d, yyyy h:mm a')
  );
}

function tmv3_step1CalendarChecklist_(record) {
  const lines = [];

  lines.push(record.eventId ? '✅ Event ID' : '❌ Event ID');
  lines.push(record.title ? '✅ Title' : '❌ Title');
  lines.push(record.start ? '✅ Date / Time' : '❌ Date / Time');
  lines.push(record.location ? '✅ Location' : '⚪ Location');
  lines.push(record.descriptionClean ? '✅ Description' : '⚪ Description');
  lines.push(record.creator ? '✅ Creator' : '⚪ Creator');
  lines.push(record.guests ? '✅ Guests' : '⚪ Guests');

  if (record.vertical === 'PreInspection') {
    lines.push(record.customerNumber ? '✅ Customer #' : '⚪ Customer #');
    lines.push(record.phone ? '✅ Phone' : '⚪ Phone');
    lines.push(record.orderNumber ? '✅ SO #' : '⚪ SO #');
  } else if (record.vertical === 'Service') {
    lines.push(record.orderNumber ? '✅ Work Order / SO #' : '⚪ Work Order / SO #');
  } else {
    lines.push(record.orderNumber ? '✅ SO #' : '⚪ SO #');
  }

  if (record.existingTaskUrl) lines.push('✅ Existing Task Link');
  if (record.existingOrderUrl) lines.push('✅ Existing Order Link');

  return lines.join('\n');
}

function tmv3_step1SourceLabel_(record) {
  const names = tmv3_unique_((record.sourceCalendarNames || []).filter(Boolean));

  if (record.vertical === 'Service' && record.technician) {
    return record.technician + (names.length ? ' · ' + names.join(' + ') : '');
  }

  return names.length
    ? names.join(' + ')
    : (record.calendarName || record.calendarId || '');
}

function tmv3_step1FormatPhone_(value) {
  const digits = String(value || '').replace(/\D/g, '');
  const phone = digits.length === 11 && digits.charAt(0) === '1'
    ? digits.slice(1)
    : digits;

  return phone.length === 10
    ? '(' + phone.slice(0,3) + ') ' + phone.slice(3,6) + '-' + phone.slice(6)
    : (tmv3_clean_(value) || '—');
}

function tmv3_step1FormatOperatorView_(sheetName, vertical, records, priorLastRow) {
  const sh = tmv3_sheet_(sheetName);
  const visible = TMV3_STEP1_VISIBLE_COLUMN_COUNT;
  const maxColumns = sh.getMaxColumns();
  const lastRow = sh.getLastRow();
  const clearThroughRow = Math.max(
    Number(lastRow || 0),
    Number(priorLastRow || 0)
  );

  sh.showColumns(1, Math.min(visible, maxColumns));

  // Customer/IDs, Schedule, and Status remain as hidden evidence columns.
  // Their operator-facing presentation is combined into the single Task column.
  if (maxColumns >= 9) {
    sh.hideColumns(6, 1);
    sh.hideColumns(8, 2);
  }

  if (maxColumns > visible) {
    sh.hideColumns(visible + 1, maxColumns - visible);
  }

  const widths = [330,220,330,260,190,190,430,180,120,170,320];

  widths.forEach(function(width, index) {
    if (index + 1 <= maxColumns) sh.setColumnWidth(index + 1, width);
  });

  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);

  // Visual grouping kept intentionally light:
  // Calendar A:E, Task G (F/H/I hidden evidence), decision J:K.
  if (maxColumns >= 5) {
    sh.getRange(1, 1, 1, 5)
      .setBackground('#EAF3FF')
      .setFontColor('#1F2937')
      .setFontWeight('bold');
  }

  if (maxColumns >= 9) {
    sh.getRange(1, 6, 1, 4)
      .setBackground('#EDF7EE')
      .setFontColor('#1F2937')
      .setFontWeight('bold');
  }

  if (maxColumns >= 11) {
    sh.getRange(1, 10, 1, 2)
      .setBackground('#F3F4F6')
      .setFontColor('#1F2937')
      .setFontWeight('bold');
  }

  if (clearThroughRow >= 2) {
    sh.getRange(2, 1, clearThroughRow - 1, visible)
      .setBackground('#FFFFFF');
  }

  if (lastRow >= 2) {
    if (vertical === 'Service' && records && records.length) {
      const techColors = {
        chris: '#EEF7FF',
        travis: '#F0FAF3',
        matt: '#FFF7ED',
        matthew: '#FFF7ED'
      };

      const backgrounds = records.map(function(record) {
        const tech = String(record.technician || '').trim().toLowerCase();
        const color = techColors[tech] || '#FFFFFF';
        return new Array(visible).fill(color);
      });

      sh.getRange(2, 1, backgrounds.length, visible)
        .setBackgrounds(backgrounds);
    }
  }

  if (lastRow >= 1) {
    sh.getRange(1, 1, lastRow, Math.min(visible, maxColumns))
      .setWrap(true)
      .setVerticalAlignment('top');
  }

  if (lastRow >= 2 && maxColumns >= 9) {
    tmv3_step1FormatTaskCells_(sh, records || []);
  }
}

function tmv3_step1FormatTaskCells_(sh, records) {
  const rowCount = Math.max(0, sh.getLastRow() - 1);
  if (!rowCount) return;

  const range = sh.getRange(2, 7, rowCount, 1);
  const values = range.getDisplayValues();

  const statusStyles = {
    'open': SpreadsheetApp.newTextStyle()
      .setForegroundColor('#137333')
      .setBold(true)
      .build(),
    'done': SpreadsheetApp.newTextStyle()
      .setForegroundColor('#5F6368')
      .setBold(true)
      .build(),
    'on hold': SpreadsheetApp.newTextStyle()
      .setForegroundColor('#B06000')
      .setBold(true)
      .build(),
    'cancelled': SpreadsheetApp.newTextStyle()
      .setForegroundColor('#B3261E')
      .setBold(true)
      .build()
  };

  const rich = values.map(function(row, index) {
    const text = String(row[0] || '');
    const record = records[index] || {};
    const builder = SpreadsheetApp.newRichTextValue().setText(text);

    // Customer line: open the Striven Customer Dashboard.
    const customer =
      record.step4 && record.step4.customer
        ? record.step4.customer
        : null;
    const customerId = tmv3_clean_(customer && customer['Customer ID']);
    if (customerId && TMV3.CUSTOMER_DASHBOARD_BASE) {
      const customerLine = text.split('\n').filter(function(line) {
        return /^Customer #/i.test(String(line || '').trim());
      })[0] || '';

      if (customerLine) {
        const start = text.indexOf(customerLine);
        if (start >= 0) {
          builder.setLinkUrl(
            start,
            start + customerLine.length,
            TMV3.CUSTOMER_DASHBOARD_BASE + encodeURIComponent(customerId)
          );
        }
      }
    }

    // Each exact Task line opens the exact Striven Task.
    const tasks =
      record.step5 && Array.isArray(record.step5.tasks)
        ? record.step5.tasks
        : [];

    tasks.forEach(function(task) {
      const taskId = tmv3_clean_(task && task['Task ID']);
      const taskUrl =
        tmv3_clean_(task && task['URL']) ||
        (taskId ? TMV3.TASK_URL_BASE + encodeURIComponent(taskId) : '');

      if (!taskId || !taskUrl) return;

      const prefix = 'Task #' + taskId;
      const lines = text.split('\n');
      const taskLine = lines.filter(function(line) {
        return String(line || '').indexOf(prefix) === 0;
      })[0] || '';

      if (taskLine) {
        const start = text.indexOf(taskLine);
        if (start >= 0) {
          builder.setLinkUrl(start, start + taskLine.length, taskUrl);
        }
      }
    });

    // Status stays in the same Task cell and is visually coded.
    text.split('\n').forEach(function(line) {
      const clean = String(line || '').trim().toLowerCase();
      const style = statusStyles[clean];
      if (!style) return;

      const start = text.indexOf(line);
      if (start >= 0) {
        builder.setTextStyle(start, start + line.length, style);
      }
    });

    return [builder.build()];
  });

  range.setRichTextValues(rich);
}

function tmv3_step1VerifyOperatorViews_(records) {
  const expected = tmv3_step1Counts_(records);
  const actual = {};
  let pass = true;

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    const sh = tmv3_sheet_(TMV3.VERTICALS[vertical].sheet);
    const rows = Math.max(0, sh.getLastRow() - 1);
    actual[vertical] = rows;
    if (rows !== Number(expected[vertical] || 0)) pass = false;
  });

  return {
    pass: pass,
    expectedRows: expected,
    actualRows: actual
  };
}

function tmv3_step1Counts_(records) {
  const counts = { Install:0, Delivery:0, Service:0, PreInspection:0 };

  (records || []).forEach(function(record) {
    if (counts[record.vertical] !== undefined) counts[record.vertical]++;
  });

  return counts;
}

function tmv3_step1CalendarIds_() {
  const ids = [];

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    tmv3_verticalCalendars_(vertical, TMV3.VERTICALS[vertical])
      .forEach(function(calCfg) {
        if (calCfg.calendarId) ids.push(calCfg.calendarId);
      });
  });

  return tmv3_unique_(ids);
}



function tmv3_step1CalendarLiveSyncStatus() {
  const handler = 'tmv3_calendarEventUpdated';
  const expected = tmv3_step1CalendarIds_();

  const active = ScriptApp.getProjectTriggers()
    .filter(function(trigger) {
      return trigger.getHandlerFunction() === handler &&
        String(trigger.getTriggerSource()) === String(ScriptApp.TriggerSource.CALENDAR);
    })
    .map(function(trigger) {
      return tmv3_safeCalendar_(function() {
        return trigger.getTriggerSourceId();
      }, '');
    })
    .filter(Boolean);

  const missing = expected.filter(function(id) {
    return active.indexOf(id) === -1;
  });

  const fallbackCount = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getHandlerFunction() === 'tmv3_calendarReconciliationFallback';
  }).length;

  return {
    handler: handler,
    expectedCalendarIds: expected,
    activeCalendarIds: tmv3_unique_(active),
    missingCalendarIds: missing,
    reconciliationFallbackCount: fallbackCount,
    complete: missing.length === 0 && fallbackCount === 1
  };
}


// Original: 34_Step2_Eligibility
/************************************************************
 * TM V3 — STAGE 2 CALENDAR NORMALIZATION / ELIGIBILITY
 *
 * Stage contract:
 * - Consumes only tmv3_step1CalendarRecords_().
 * - Performs no Striven reads/writes.
 * - Performs no Google Calendar writes.
 * - Classifies every Step-1 record as ELIGIBLE / REVIEW / SKIP.
 * - Later stages may consume only tmv3_step2EligibleCalendarRecords_().
 ************************************************************/

function tmv3_step2CalendarRecords_(step1Records) {
  const sourceRecords = step1Records || tmv3_step1CalendarRecords_();

  return sourceRecords.map(function(record) {
    const classified = tmv3_step2ClassifyCalendarRecord_(record);
    return Object.assign({}, record, { step2: classified });
  });
}

function tmv3_step2EligibleCalendarRecords_() {
  return tmv3_step2CalendarRecords_().filter(function(record) {
    return record.step2 && record.step2.disposition === 'ELIGIBLE';
  });
}

function tmv3_step2ClassifyCalendarRecord_(record) {
  if (!record || !record.vertical) {
    return tmv3_step2Decision_(
      'REVIEW',
      'INVALID_RECORD',
      'Calendar record is missing its vertical.',
      false
    );
  }

  if (record.vertical === 'PreInspection') {
    return tmv3_step2ClassifyPreInspection_(record);
  }

  if (record.vertical === 'Service') {
    return tmv3_step2ClassifyService_(record);
  }

  if (
    tmv3_shouldIgnoreEvent_(
      record.vertical,
      record.title,
      record.rawDescription,
      record.isAllDay,
      record.creator
    )
  ) {
    return tmv3_step2Decision_(
      'SKIP',
      'LEGACY_NON_JOB_FILTER',
      'Legacy Calendar non-job / creator filter matched.',
      false
    );
  }

  return tmv3_step2Decision_(
    'ELIGIBLE',
    'LEGACY_ELIGIBLE',
    'Calendar event passed the legacy ' + record.vertical + ' eligibility rules.',
    false
  );
}









function tmv3_step2Decision_(disposition, code, reason, mirrorRequired, warnings, mirrorStatus) {
  return {
    disposition: disposition,
    code: code,
    reason: reason,
    mirrorRequired: mirrorRequired === true,
    mirrorStatus: mirrorStatus || '',
    warnings: warnings || []
  };
}

function tmv3_step2CalendarRun(reason) {
  tmv3_assertShadow_();

  const step1Snapshot = tmv3_step1CalendarRecords_();
  const records = tmv3_step2CalendarRecords_(step1Snapshot);
  const counts = tmv3_step2Counts_(records);
  const write = tmv3_step2WriteOperatorViews_(records);
  const verify = tmv3_step2Verify_(records, step1Snapshot);

  const result = {
    version: TMV3.VERSION,
    executionStage: tmv3_executionStage_(),
    stage: 2,
    status: verify.pass ? 'PASS' : 'REVIEW',
    mode: 'CALENDAR_NORMALIZATION_AND_ELIGIBILITY_ONLY',
    reason: tmv3_clean_(reason || 'MANUAL'),
    calendarWritesPerformed: false,
    strivenReadsPerformed: false,
    strivenWritesPerformed: false,
    counts: counts,
    write: write,
    verification: verify
  };

  tmv3_audit_(
    'SYSTEM','','','STEP2_CALENDAR_ELIGIBILITY',
    result.status,
    JSON.stringify(result)
  );

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function tmv3_step2Counts_(records) {
  const counts = {
    total: 0,
    eligible: 0,
    review: 0,
    skip: 0,
    mirrorRequired: 0,
    serviceExternalRecovery: 0,
    byVertical: {}
  };

  (records || []).forEach(function(record) {
    counts.total++;
    if (!counts.byVertical[record.vertical]) {
      counts.byVertical[record.vertical] = {
        total:0, eligible:0, review:0, skip:0
      };
    }

    const bucket = counts.byVertical[record.vertical];
    bucket.total++;

    const disposition = record.step2 ? record.step2.disposition : 'REVIEW';
    if (disposition === 'ELIGIBLE') {
      counts.eligible++;
      bucket.eligible++;
    } else if (disposition === 'SKIP') {
      counts.skip++;
      bucket.skip++;
    } else {
      counts.review++;
      bucket.review++;
    }

    if (record.step2 && record.step2.mirrorRequired) counts.mirrorRequired++;
    if (record.step2 && record.step2.code === 'SERVICE_EXTERNAL_CREATOR_RECOVERY') {
      counts.serviceExternalRecovery++;
    }
  });

  return counts;
}

function tmv3_step2WriteOperatorViews_(records) {
  const written = {};

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    const verticalRecords = (records || [])
      .filter(function(record) { return record.vertical === vertical; })
      .sort(function(a, b) {
        const aTime = a.start instanceof Date ? a.start.getTime() : 0;
        const bTime = b.start instanceof Date ? b.start.getTime() : 0;
        if (aTime !== bTime) return aTime - bTime;

        if (vertical === 'Service') {
          const techCompare = tmv3_clean_(a.technician)
            .localeCompare(tmv3_clean_(b.technician));
          if (techCompare !== 0) return techCompare;
        }

        return tmv3_clean_(a.title).localeCompare(tmv3_clean_(b.title));
      });

    const rows = verticalRecords.map(tmv3_step2OperatorRow_);
    const sheetName = TMV3.VERTICALS[vertical].sheet;
    const priorLastRow = tmv3_sheet_(sheetName).getLastRow();

    tmv3_replaceRows_(sheetName, TMV3_STEP1_HEADERS.slice(), rows);
    tmv3_step1FormatOperatorView_(sheetName, vertical, verticalRecords, priorLastRow);
    written[vertical] = rows.length;
  });

  return written;
}

function tmv3_step2OperatorRow_(record) {
  const row = tmv3_step1OperatorRow_(record);
  const step2 = record.step2 || {};
  const disposition = step2.disposition || 'REVIEW';

  row[4] = tmv3_step1CalendarChecklist_(record) +
    '\n' +
    (disposition === 'ELIGIBLE' ? '✅ Step 2 Eligible' :
      disposition === 'SKIP' ? '⏭ Step 2 Skip' : '⚠ Step 2 Review');

  row[5] = 'NOT RUN — STEP 2';
  row[6] = 'NOT RUN — STEP 2';
  row[7] = 'NOT RUN — STEP 2';
  row[8] = 'NOT RUN — STEP 2';
  row[9] = 'STEP 2 — ' + disposition;

  const details = [step2.reason || ''];
  if (step2.mirrorStatus) details.push('Mirror: ' + step2.mirrorStatus);
  (step2.warnings || []).forEach(function(warning) {
    details.push('Attention: ' + warning);
  });
  row[10] = details.filter(Boolean).join('\n');

  return row;
}

function tmv3_step2Verify_(records, step1Snapshot) {
  const step1 = step1Snapshot || tmv3_step1CalendarRecords_();
  const step1Keys = {};
  const step2Keys = {};
  let missingDecision = 0;
  let badIdentity = 0;

  step1.forEach(function(record) {
    step1Keys[record.logicalKey] = true;
  });

  records.forEach(function(record) {
    step2Keys[record.logicalKey] = true;

    if (
      !record.step2 ||
      ['ELIGIBLE','REVIEW','SKIP'].indexOf(record.step2.disposition) === -1
    ) {
      missingDecision++;
    }

    if (!record.eventId || !record.start) badIdentity++;
  });

  const missingFromStep2 = Object.keys(step1Keys).filter(function(key) {
    return !step2Keys[key];
  });

  const introducedByStep2 = Object.keys(step2Keys).filter(function(key) {
    return !step1Keys[key];
  });

  const expectedRows = {};
  const actualRows = {};

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    expectedRows[vertical] = records.filter(function(record) {
      return record.vertical === vertical;
    }).length;

    actualRows[vertical] = Math.max(
      0,
      tmv3_sheet_(TMV3.VERTICALS[vertical].sheet).getLastRow() - 1
    );
  });

  const rowCountsMatch = Object.keys(expectedRows).every(function(vertical) {
    return expectedRows[vertical] === actualRows[vertical];
  });

  return {
    pass:
      missingDecision === 0 &&
      badIdentity === 0 &&
      missingFromStep2.length === 0 &&
      introducedByStep2.length === 0 &&
      rowCountsMatch,
    step1Records: step1.length,
    step2Records: records.length,
    missingDecision: missingDecision,
    badIdentity: badIdentity,
    missingFromStep2: missingFromStep2.length,
    introducedByStep2: introducedByStep2.length,
    expectedRows: expectedRows,
    actualRows: actualRows
  };
}



// Original: 35_CalendarLinks
/************************************************************
 * TM V3 — SINGLE MANAGED CALENDAR LINK ENGINE
 ************************************************************/

const TMV3_LINK_BLOCK_START = '<!-- TMV3_STRIVEN_LINKS_START -->';
const TMV3_LINK_BLOCK_END = '<!-- TMV3_STRIVEN_LINKS_END -->';
const TMV3_FINAL_LINK_HEADING = TMV3_HARD_RULES.PreInspection.presentation.linkHeading;

function tmv3_buildCalendarLinkPlan_(eventRecord, resolved) {
  const cfg = TMV3.VERTICALS[eventRecord.vertical];
  const records = Array.isArray(resolved) ? resolved : [resolved];
  const links = [];

  const first = records[0] || {};

  if (eventRecord.vertical === 'PreInspection') {
    if (!first.customerId) {
      throw new Error('PreInspection Calendar links require Customer ID.');
    }

    const customer = eventRecord.step4 && eventRecord.step4.customer || null;
    const customerName = tmv3_clean_(
      first.customer ||
      (customer && customer['Name']) ||
      eventRecord.calendarCustomerName ||
      tmv3_preInspectionCalendarCustomerName_(
        eventRecord.title || '',
        eventRecord.customerNumber || ''
      )
    );
    const customerNumber = tmv3_clean_(
      (customer && customer['Customer Number']) ||
      eventRecord.customerNumber ||
      tmv3_extractCustomerNumberForVertical_(
        'PreInspection',
        eventRecord.title || ''
      )
    );
    const phone = tmv3_titlePhoneDisplay_(
      eventRecord.phone ||
      (customer && customer['Primary Phone']) ||
      ''
    );

    if (!customerName || !customerNumber || !phone) {
      throw new Error(
        'PreInspection Calendar links require verified Customer Name, Customer Number and Phone.'
      );
    }

    links.push({
      key:'CUSTOMER_SALES_ORDERS_PAGE',
      label:
        TMV3_HARD_RULES.PreInspection.presentation.customerLinkPrefix +
        customerName +
        ' (#' + customerNumber + ')',
      url:
        TMV3.CUSTOMER_ORDERS_PAGE_BASE +
        encodeURIComponent(first.customerId)
    });

    eventRecord.__tmv3PreInspectionLinkIdentity = {
      customerName:customerName,
      customerNumber:customerNumber,
      phone:phone
    };
  } else if (cfg.calendarOrderLinkRequired) {
    if (!first.orderId) {
      throw new Error(cfg.orderLabel + ' Calendar link requires internal Order ID.');
    }

    links.push({
      key: 'ORDER',
      label:
        cfg.orderLabel +
        (first.order ? ' – ' + first.order : ''),
      url:
        TMV3.ORDER_URL_BASE +
        encodeURIComponent(first.orderId)
    });
  }

  const taskSeen = {};

  if (cfg.calendarTaskLinkRequired) {
    records.forEach(function(record) {
      if (!record || !record.taskId) {
        throw new Error('Calendar Task link requires verified Task ID for every mapped task.');
      }

      const taskId = String(record.taskId);
      if (taskSeen[taskId]) return;
      taskSeen[taskId] = true;

      let label = '';

      if (eventRecord.vertical === 'PreInspection') {
        const taskName = tmv3_clean_(record.task);

        if (!taskName || /^(?:task\s*#?\s*)?\d+$/i.test(taskName)) {
          throw new Error(
            'PreInspection Calendar Task link requires the actual verified Striven Task Name.'
          );
        }

        label = TMV3_HARD_RULES.PreInspection.presentation.taskLinkPrefix + taskId + TMV3_HARD_RULES.PreInspection.presentation.taskLinkSeparator + taskName;
      } else {
        label =
          (
            record.serviceFireplaceNumber
              ? 'FP#' + record.serviceFireplaceNumber + ' – '
              : ''
          ) +
          'Task #' +
          taskId +
          (record.task ? ' – ' + record.task : '');
      }

      links.push({
        key:
          record.serviceFireplaceNumber
            ? 'TASK_FP_' + record.serviceFireplaceNumber
            : 'TASK_' + taskId,
        label:label,
        url:
          TMV3.TASK_URL_BASE +
          encodeURIComponent(taskId)
      });
    });
  }

  const preIdentity =
    eventRecord.__tmv3PreInspectionLinkIdentity || null;
  try { delete eventRecord.__tmv3PreInspectionLinkIdentity; } catch (ignored) {}

  return {
    vertical: eventRecord.vertical,
    eventId: eventRecord.eventId,
    occurrenceStart: eventRecord.start ? tmv3_iso_(eventRecord.start) : '',
    calendarIds: tmv3_requiredCalendarCopiesForEvent_(eventRecord),
    currentTitle:tmv3_clean_(eventRecord.title),
    customerName:preIdentity ? preIdentity.customerName : '',
    customerNumber:preIdentity ? preIdentity.customerNumber : '',
    phone:preIdentity ? preIdentity.phone : '',
    links: links
  };
}

function tmv3_requiredCalendarCopiesForEvent_(eventRecord) {
  const record = eventRecord || {};
  const vertical = tmv3_clean_(record.vertical);
  const sourceIds = tmv3_unique_(
    []
      .concat(record.sourceCalendarIds || [])
      .concat(record.calendarId || [])
      .map(tmv3_clean_)
      .filter(Boolean)
  );

  if (vertical === 'PreInspection') {
    const cfg = TMV3.VERTICALS.PreInspection || {};
    const primaryId = tmv3_clean_(cfg.primaryCalendarId);
    const secondaryIds = (cfg.secondaryCalendarIds || [])
      .map(tmv3_clean_)
      .filter(Boolean);

    // CF Preinspects remains the sole intake authority, but presentation
    // state must converge on every configured PreInspection Calendar copy.
    const organizer = tmv3_clean_(record.organizerCalendarId);
    // Stephen's copy is required only for events organized on Stephen's
    // configured calendar, or for legacy records with unknown organizer.
    const requiredSecondary = !organizer || secondaryIds.indexOf(organizer)>=0 ? secondaryIds : [];
    return tmv3_unique_([primaryId].concat(requiredSecondary).filter(Boolean));
  }

  if (sourceIds.length) return sourceIds;
  return tmv3_requiredCalendarCopies_(vertical);
}

function tmv3_requiredCalendarCopies_(vertical) {
  const cfg = TMV3.VERTICALS[vertical];

  if (vertical === 'Service') {
    return [];
  }

  if (vertical === 'PreInspection') {
    return tmv3_unique_(
      [cfg.primaryCalendarId]
        .concat(cfg.secondaryCalendarIds || [])
        .map(tmv3_clean_)
        .filter(Boolean)
    );
  }

  return (cfg.calendarIds || []).slice();
}

function tmv3_calendarHtmlEscape_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function tmv3_calendarHtmlText_(value) {
  return tmv3_calendarHtmlEscape_(value)
    .replace(/\r\n?/g, '\n')
    .replace(/\n/g, '<br>');
}

function tmv3_managedCalendarDescription_(existingDescription, plan) {
  if(plan && ['Install','Delivery','Service'].indexOf(plan.vertical)>=0)return tmv3_verticalClickableDescription_(existingDescription,plan);
  const authored = plan && plan.vertical === 'PreInspection'
    ? tmv3_preInspectionCalendarNotesText_(
        existingDescription,
        plan
      )
    : tmv3_stripManagedLinkBlocks_(existingDescription);

  const linkLines = (plan.links || []).map(function(link) {
    return '[' +
      tmv3_calendarMarkdownLabel_(link.label) +
      '](' +
      String(link.url || '') +
      ')';
  });

  const managed =
    TMV3_FINAL_LINK_HEADING +
    (linkLines.length ? '  \n' + linkLines.join('  \n') : '');

  return (
    authored
      ? authored.replace(/\s+$/, '') + '\n\n'
      : ''
  ) + managed;
}





















function tmv3_calendarMarkdownLabel_(value) {
  return String(value == null ? '' : value)
    .replace(/\\/g, '\\\\')
    .replace(/\[/g, '\\[')
    .replace(/\]/g, '\\]');
}


function tmv3_managedCalendarBlockStart_(value) {
  const text = String(value || '');
  const patterns = [
    /(?:<b>\s*)?(?:\*\*\s*)?\\?-{5}\s*Striven\s+Links\s*-{5}(?:\s*\*\*)?(?:\s*<\/b>)?/i,
    /(?:<b>\s*)?(?:\*\*\s*)?Striven\s+Links(?:\s*\*\*)?(?:\s*<\/b>)?/i,
    /-{5,}\s*Pre-Inspection Task Link\s*-{5,}/i,
    /-{5,}\s*Pre Inspection Task Link\s*-{5,}/i,
    /-{5,}\s*Install Striven Task Link\s*-{5,}/i,
    /-{5,}\s*Delivery Striven\s+Links\s*-{5,}/i,
    /-{5,}\s*Service Striven\s+Links\s*-{5,}/i,
    /-{5,}\s*Delivery Task Link\s*-{5,}/i
  ];

  let best = -1;

  patterns.forEach(function(pattern) {
    const match = pattern.exec(text);
    if (!match) return;
    if (best === -1 || match.index < best) best = match.index;
  });

  return best;
}

function tmv3_stripManagedLinkBlocks_(description) {
  let text = tmv3_stripV3ManagedBlock_(description);

  text = text.replace(
    /<!--\s*PREINSPECT_STRIVEN_TASK_LINK_START\s*-->[\s\S]*?<!--\s*PREINSPECT_STRIVEN_TASK_LINK_END\s*-->/gi,
    ''
  );

  const managedStart = tmv3_managedCalendarBlockStart_(text);
  if (managedStart >= 0) {
    text = text.slice(0, managedStart);
  }

  return text
    .replace(/(?:<br\s*\/?>\s*){3,}/gi, '<br><br>')
    .replace(/(?:\r?\n\s*){3,}/g, '\n\n')
    .replace(/^\s+|\s+$/g, '');
}

function tmv3_stripV3ManagedBlock_(description) {
  const raw = String(description || '');
  const escapedStart = tmv3_regexEscape_(TMV3_LINK_BLOCK_START);
  const escapedEnd = tmv3_regexEscape_(TMV3_LINK_BLOCK_END);
  const rx = new RegExp(escapedStart + '[\\s\\S]*?' + escapedEnd, 'g');
  return raw.replace(rx, '').replace(/\s+$/, '');
}

function tmv3_regexEscape_(value) {
  return String(value || '').replace(/[.*+?^\$\{\}()|[\]\\]/g, '\\$&');
}

function tmv3_findEventCopies_(eventId, calendarIds, occurrenceStart) {
  const copies = [];
  const target = occurrenceStart ? new Date(occurrenceStart) : null;

  (calendarIds || []).forEach(function(calendarId) {
    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) return;

    const event = tmv3_findCalendarOccurrence_(
      calendar,
      eventId,
      target
    );

    if (!event) return;

    copies.push({
      calendarId: calendarId,
      calendarName: calendar.getName(),
      event: event
    });
  });

  return copies;
}

function tmv3_findCalendarOccurrence_(calendar, eventId, targetStart) {
  const direct = calendar.getEventById(eventId);

  if (!targetStart) return direct || null;

  if (
    direct &&
    Math.abs(direct.getStartTime().getTime() - targetStart.getTime()) < 1000
  ) {
    return direct;
  }

  const from = new Date(targetStart.getTime() - 60000);
  const to = new Date(targetStart.getTime() + 60000);

  const matches = calendar.getEvents(from, to).filter(function(event) {
    return (
      String(event.getId()) === String(eventId) &&
      Math.abs(event.getStartTime().getTime() - targetStart.getTime()) < 1000
    );
  });

  if (matches.length > 1) {
    throw new Error(
      'Calendar occurrence lookup is ambiguous for Event ID ' +
      eventId +
      ' at ' +
      targetStart.toISOString() +
      '.'
    );
  }

  return matches.length === 1 ? matches[0] : null;
}

function tmv3_previewCalendarLinks_(eventRecord, resolved) {
  return {
    mode: TMV3.MODE,
    writeAllowed: tmv3_writesEnabled_(),
    plan: tmv3_buildCalendarLinkPlan_(eventRecord, resolved)
  };
}

function tmv3_assertLegacyCalendarMutationDisabled_(helperName) {
  throw new Error(
    'STEP7_AUTHORITY_REQUIRED: legacy direct Calendar mutation helper ' +
    String(helperName || 'UNKNOWN') +
    ' is disabled. Use the Step 7 executor (for example tmv3_fixSelectedCalendarLinks).'
  );
}

function tmv3_writeCalendarLinks_(eventRecord, resolved) {
  tmv3_assertLegacyCalendarMutationDisabled_('tmv3_writeCalendarLinks_');
  return tmv3_writeCalendarLinksForEvent_(eventRecord, [resolved]);
}


function tmv3_writeCalendarLinksForEvent_(eventRecord, resolvedRecords) {
  tmv3_assertLegacyCalendarMutationDisabled_('tmv3_writeCalendarLinksForEvent_');
  tmv3_assertBusinessWritesEnabled_();

  const records = Array.isArray(resolvedRecords)
    ? resolvedRecords
    : [resolvedRecords];

  const plan = tmv3_buildCalendarLinkPlan_(eventRecord, records);
  const copies = tmv3_findEventCopies_(
    plan.eventId,
    plan.calendarIds,
    plan.occurrenceStart
  );

  if (!copies.length) {
    throw new Error('No Calendar copy found for Event ID ' + plan.eventId + '.');
  }

  const results = [];

  copies.forEach(function(copy) {
    const before = String(copy.event.getDescription() || '');
    const desired = tmv3_managedCalendarDescription_(before, plan);

    if (before !== desired) copy.event.setDescription(desired);

    const after = String(copy.event.getDescription() || '');
    const missing = (plan.links || []).filter(function(link) {
      return after.indexOf(link.url) === -1;
    });

    if (missing.length) {
      throw new Error(
        'Calendar read-back failed for ' +
        copy.calendarName +
        ': missing ' +
        missing.map(function(x) { return x.key; }).join(', ')
      );
    }

    results.push({
      calendarId: copy.calendarId,
      calendarName: copy.calendarName,
      status: before === desired ? 'ALREADY_CORRECT' : 'WRITTEN_AND_VERIFIED'
    });
  });

  return {
    status: 'CALENDAR_LINKS_VERIFIED',
    eventId: plan.eventId,
    copies: results,
    links: plan.links
  };
}



// Google Apps Script: retain valid per-vertical blocks; repair only managed links.
function tmv3_verticalClickableDescription_(description,plan) {
  const raw=String(description||''),links=plan.links||[];
  const rules=TMV3_HARD_RULES[plan.vertical];
  if(!rules||!rules.presentation)throw new Error('UNDEFINED_VERTICAL_LINK_STANDARD_NO_WRITE');
  const heading=rules.presentation.linkHeading;
  const correct=raw.indexOf(heading)>=0 && links.length>0 && links.every(function(link){
    const esc=tmv3_regexEscape_(String(link.url));
    const rx=new RegExp('<a\\b[^>]*\\bhref=["\']'+esc+'["\'][^>]*>([\\s\\S]*?)<\\/a>','i');
    const m=rx.exec(raw);
    if(!m)return false;
    const label=tmv3_decodeCalendarHtmlEntities_(m[1].replace(/<[^>]*>/g,''));
    if(/^\s*(?:Task\s*#?\s*)?\d+\s*(?:[-–]\s*\d+)?\s*$/i.test(label))return false;
    return !/^TASK/.test(link.key)||label.indexOf(link.label.replace(/^.*?Task #\d+\s*[–-]\s*/,''))>=0;
  });
  if(correct)return raw;
  const authored=tmv3_stripManagedLinkBlocks_(raw),esc=tmv3_calendarHtmlEscape_;
  const anchor=function(link,label){return '<a href="'+esc(link.url)+'">'+esc(label||link.label)+'</a>';};
  const tasks=links.filter(function(l){return /^TASK/.test(l.key);});
  const order=links.filter(function(l){return l.key==='ORDER';})[0];
  const taskLines=tasks.map(function(l){return anchor(l,l.label.replace(/^(.*?Task #\d+)\s*–\s*/,'$1 - '));});
  let block;
  if(plan.vertical==='Install')block=heading+'<br><br>'+taskLines.join('<br>')+'<br><br>------------------------------------';
  else if(plan.vertical==='Delivery')block=heading+'<br><br>'+rules.presentation.orderSection+'<br>'+(order?anchor(order,order.label.replace(/^Sales Order\s*[–-]\s*/,'SO #')):'')+'<br><br>'+rules.presentation.taskSection+'<br>'+taskLines.join('<br>')+'<br><br>------------------------------------';
  else block=heading+'<br>'+(order?anchor(order,order.label.replace(/\s*[–-]\s*/,' #'))+'<br>':'')+taskLines.join('<br>')+'<br>------------------------------------';
  return (authored?authored+'<br><br>':'')+block;
}


// Original: 36_Step3_Business_Anchor
/************************************************************
 * TM V3 — STAGE 3 BUSINESS-ANCHOR RESOLUTION
 *
 * Contract:
 * - Consumes the exact Step-2 event set.
 * - Resolves only business anchors.
 * - Does NOT resolve Location / Contact / Task.
 * - Does NOT create or update Striven.
 * - Does NOT write Google Calendar.
 ************************************************************/

function tmv3_step3RefreshAnchorSources_() {
  tmv3_assertShadow_();
  tmv3_resetRuntimeMetrics_();

  const customers =
    tmv3_reportRows_(TMV3.PROPERTIES.CUSTOMERS)
      .map(tmv3_normalizeCustomer_);

  const approved =
    tmv3_reportRows_(TMV3.PROPERTIES.APPROVED_ORDERS)
      .map(function(r) {
        return tmv3_normalizeOrder_(r, 'SALES_ORDER');
      });

  const deliveryApproved =
    tmv3_reportRows_(TMV3.PROPERTIES.DELIVERY_APPROVED_ORDERS)
      .map(function(r) {
        return tmv3_normalizeOrder_(r, 'DELIVERY_APPROVED');
      });

  const serviceOrders =
    tmv3_reportRows_(TMV3.PROPERTIES.SERVICE_WORK_ORDERS)
      .map(function(r) {
        return tmv3_normalizeOrder_(r, 'WORK_ORDER');
      });

  const orders = tmv3_mergeOrderRows_(
    approved.concat(deliveryApproved, serviceOrders)
  );

  tmv3_assertSourceReplacementSafe_(
    'Stage 3 Customers',
    customers,
    TMV3.SHEETS.CUSTOMERS
  );
  tmv3_assertSourceReplacementSafe_(
    'Stage 3 Orders',
    orders,
    TMV3.SHEETS.ORDERS
  );

  tmv3_replaceRows_(
    TMV3.SHEETS.CUSTOMERS,
    [
      'Customer ID',
      'Customer Number',
      'Name',
      'Primary Phone',
      'Primary Email',
      'Fingerprint'
    ],
    customers
  );

  tmv3_replaceRows_(
    TMV3.SHEETS.ORDERS,
    [
      'Order ID',
      'Order Number',
      'Customer ID',
      'Location ID',
      'Contact ID',
      'Status',
      'Name',
      'Order Type',
      'URL',
      'Fingerprint'
    ],
    orders
  );

  const result = {
    customers: customers.length,
    orders: orders.length,
    api: tmv3_runtimeMetrics_()
  };

  tmv3_audit_(
    'SYSTEM','','','STEP3_REFRESH_ANCHOR_SOURCES','PASS',
    JSON.stringify(result)
  );

  return result;
}

function tmv3_step3AnchorSourceSummary_() {
  return {
    customers: tmv3_rows_(TMV3.SHEETS.CUSTOMERS).length,
    orders: tmv3_rows_(TMV3.SHEETS.ORDERS).length,
    source: 'CACHED_ANCHOR_SOURCES'
  };
}

function tmv3_step3AnchorIndex_() {
  const customers = tmv3_rows_(TMV3.SHEETS.CUSTOMERS);
  const orders = tmv3_rows_(TMV3.SHEETS.ORDERS);

  const refs = {
    customers: customers,
    orders: orders,
    customerById: {},
    customerByNumber: {},
    orderById: {},
    ordersByNumber: {}
  };

  customers.forEach(function(customer) {
    const id = tmv3_clean_(customer['Customer ID']);
    const number = tmv3_clean_(customer['Customer Number']);

    if (id) refs.customerById[id] = customer;
    if (number) refs.customerByNumber[number] = customer;
  });

  orders.forEach(function(order) {
    const id = tmv3_clean_(order['Order ID']);
    const number = tmv3_clean_(order['Order Number']);

    if (id) refs.orderById[id] = order;

    if (number) {
      if (!refs.ordersByNumber[number]) refs.ordersByNumber[number] = [];
      refs.ordersByNumber[number].push(order);
    }
  });

  return refs;
}

function tmv3_step3AllowedOrderTypes_(vertical) {
  const v = tmv3_clean_(vertical);

  if (v === 'Install') return ['SALES_ORDER'];
  if (v === 'Delivery') return ['SALES_ORDER', 'DELIVERY_APPROVED'];
  if (v === 'Service') return ['WORK_ORDER'];
  if (v === 'PreInspection') {
    return ['SALES_ORDER', 'DELIVERY_APPROVED'];
  }

  return [];
}

function tmv3_step3OrderTypes_(order) {
  return tmv3_unique_(
    tmv3_clean_(order && order['Order Type'])
      .split('|')
      .map(function(type) {
        return tmv3_clean_(type).toUpperCase();
      })
      .filter(Boolean)
  );
}

function tmv3_step3OrderMatchesVertical_(order, vertical) {
  const allowed = tmv3_step3AllowedOrderTypes_(vertical);
  const actual = tmv3_step3OrderTypes_(order);

  if (!allowed.length || !actual.length) return false;

  return actual.some(function(type) {
    return allowed.indexOf(type) !== -1;
  });
}

function tmv3_step3BusinessAnchorRecords_(step2Records, refs) {
  return (step2Records || []).map(function(record) {
    const step2 = record.step2 || {};

    if (step2.disposition !== 'ELIGIBLE') {
      return Object.assign({}, record, {
        step3: {
          disposition: 'NOT_RUN',
          code: 'STEP2_' + (step2.disposition || 'UNKNOWN'),
          reason: 'Business-anchor resolution did not run because Step 2 is not ELIGIBLE.',
          evidence: [],
          warnings: []
        }
      });
    }

    const decision =
      record.vertical === 'PreInspection'
        ? tmv3_step3ResolvePreInspectionAnchor_(record, refs)
        : tmv3_step3ResolveOrderAnchor_(record, refs);

    return Object.assign({}, record, { step3: decision });
  });
}

function tmv3_step3ResolveOrderAnchor_(record, refs) {
  const cfg = TMV3.VERTICALS[record.vertical];
  const label = cfg.orderLabel || 'Order';
  const orderNumber = tmv3_clean_(record.orderNumber);
  const existingOrderId = tmv3_clean_(record.existingOrderId);
  let candidates = [];
  const evidence = [];

  if (existingOrderId && refs.orderById[existingOrderId]) {
    candidates = [refs.orderById[existingOrderId]];
    evidence.push('CALENDAR_ORDER_LINK_ID_EXACT');
  } else if (orderNumber) {
    candidates = (refs.ordersByNumber[orderNumber] || []).slice();
    evidence.push('CALENDAR_ORDER_NUMBER_EXACT');
  }

  candidates = tmv3_step3UniqueOrders_(candidates);
  const unfilteredCandidates = candidates.slice();

  candidates = candidates.filter(function(order) {
    return tmv3_step3OrderMatchesVertical_(order, record.vertical);
  });

  if (!existingOrderId && !orderNumber) {
    return tmv3_step3Decision_(
      'BLOCKED',
      'ANCHOR_ORDER_NUMBER_MISSING',
      label + ' number/link is required before business-anchor resolution.',
      null,
      evidence
    );
  }

  if (!candidates.length) {
    if (unfilteredCandidates.length) {
      return tmv3_step3Decision_(
        'BLOCKED',
        'ANCHOR_TRANSACTION_TYPE_MISMATCH',
        label +
          ' number/link exists in the shared Order cache, but not as the transaction type allowed for ' +
          record.vertical +
          '. Expected: ' +
          tmv3_step3AllowedOrderTypes_(record.vertical).join(' or ') +
          '.',
        null,
        evidence.concat([
          'ORDER_NUMBER_PRESENT_WRONG_TRANSACTION_TYPE'
        ])
      );
    }

    return tmv3_step3Decision_(
      'BLOCKED',
      'ANCHOR_ORDER_NOT_FOUND',
      label + ' was not found in the current Striven anchor source.',
      null,
      evidence
    );
  }

  if (candidates.length > 1) {
    return tmv3_step3Decision_(
      'REVIEW',
      'ANCHOR_ORDER_AMBIGUOUS',
      'Multiple distinct Striven records match ' + label + ' ' + (orderNumber || existingOrderId) + '.',
      null,
      evidence
    );
  }

  const order = candidates[0];
  const customerId = tmv3_clean_(order['Customer ID']);
  const customer = customerId ? (refs.customerById[customerId] || null) : null;
  const anchor = tmv3_step3AnchorFromOrder_(record, order, customer);

  if (!customerId && record.vertical === 'Service') {
    return tmv3_step3Decision_(
      'VERIFIED',
      'SERVICE_WORK_ORDER_ANCHOR_VERIFIED',
      'Work Order resolved uniquely. Customer identity is intentionally deferred to Step 4.',
      anchor,
      evidence.concat(['WORK_ORDER_UNIQUE']),
      ['Work Order source does not expose Customer ID; Step 4 will resolve Customer from independent evidence.']
    );
  }

  if (!customerId) {
    return tmv3_step3Decision_(
      'REVIEW',
      'ANCHOR_ORDER_CUSTOMER_MISSING',
      label + ' resolved, but it does not expose a Customer ID.',
      anchor,
      evidence
    );
  }

  if (!customer) {
    return tmv3_step3Decision_(
      'REVIEW',
      'ANCHOR_CUSTOMER_CACHE_MISSING',
      label + ' resolved to Customer ID ' + customerId + ', but that Customer is not present in the current Customer source.',
      anchor,
      evidence
    );
  }

  return tmv3_step3Decision_(
    'VERIFIED',
    'ANCHOR_ORDER_VERIFIED',
    label + ' resolved uniquely and exposes a verified Customer relationship.',
    anchor,
    evidence.concat(['ORDER_CUSTOMER_RELATIONSHIP_PRESENT'])
  );
}





function tmv3_step3BareSixDigitOrderCandidates_(record) {
  const text = tmv3_stripPhonesForOrderParsing_(
    [
      tmv3_clean_(record.title),
      tmv3_clean_(record.descriptionClean)
    ].join(' ')
  );

  const matches = text.match(/\b\d{6}\b/g) || [];
  const customerNumber = tmv3_clean_(record.customerNumber);

  return tmv3_unique_(
    matches.filter(function(number) {
      return number !== customerNumber;
    })
  );
}

function tmv3_step3UniqueOrders_(orders) {
  const byKey = {};

  (orders || []).forEach(function(order) {
    const key =
      tmv3_clean_(order['Order ID']) ||
      ('NUMBER:' + tmv3_clean_(order['Order Number']));

    if (key) byKey[key] = order;
  });

  return Object.keys(byKey).map(function(key) {
    return byKey[key];
  });
}

function tmv3_step3AnchorFromOrder_(record, order, customer) {
  return {
    anchorType: record.vertical === 'Service' ? 'WORK_ORDER' : 'SALES_ORDER',
    orderId: tmv3_clean_(order && order['Order ID']),
    orderNumber: tmv3_clean_(order && order['Order Number']),
    orderType: tmv3_clean_(order && order['Order Type']),
    orderStatus: tmv3_clean_(order && order['Status']),
    orderUrl: tmv3_clean_(order && order['URL']),
    customerId: tmv3_clean_(order && order['Customer ID']),
    customerNumber: tmv3_clean_(customer && customer['Customer Number']),
    customerName: tmv3_clean_(customer && customer['Name']),
    orderLocationIdEvidence: tmv3_clean_(order && order['Location ID']),
    orderContactIdEvidence: tmv3_clean_(order && order['Contact ID']),
    orderEvidenceOnly: false
  };
}



function tmv3_step3Decision_(disposition, code, reason, anchor, evidence, warnings) {
  return {
    disposition: disposition,
    code: code,
    reason: reason,
    anchor: anchor || {},
    evidence: evidence || [],
    warnings: warnings || []
  };
}

function tmv3_step3BusinessAnchorRun(reason, refreshSources) {
  tmv3_assertShadow_();

  const sourceSummary =
    refreshSources === true
      ? tmv3_step3RefreshAnchorSources_()
      : tmv3_step3AnchorSourceSummary_();

  const step2Snapshot = tmv3_step2CalendarRecords_();
  const refs = tmv3_step3AnchorIndex_();
  const records = tmv3_step3BusinessAnchorRecords_(step2Snapshot, refs);
  const counts = tmv3_step3Counts_(records);
  const write = tmv3_step3WriteOperatorViews_(records);
  const verify = tmv3_step3Verify_(records, step2Snapshot);

  const result = {
    version: TMV3.VERSION,
    executionStage: tmv3_executionStage_(),
    stage: 3,
    status: verify.pass ? 'PASS' : 'REVIEW',
    mode: 'BUSINESS_ANCHOR_RESOLUTION_ONLY',
    reason: tmv3_clean_(reason || 'MANUAL'),
    sourceSummary: sourceSummary,
    counts: counts,
    write: write,
    verification: verify,
    calendarWritesPerformed: false,
    strivenWritesPerformed: false,
    taskResolutionPerformed: false,
    locationResolutionPerformed: false,
    contactResolutionPerformed: false
  };

  tmv3_audit_(
    'SYSTEM','','','STEP3_BUSINESS_ANCHOR',
    result.status,
    JSON.stringify(result)
  );

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function tmv3_step3BusinessAnchorRunCached(reason) {
  return tmv3_step3BusinessAnchorRun(
    reason || 'CACHED_ANCHOR_SOURCES',
    false
  );
}

function tmv3_step3Counts_(records) {
  const counts = {
    total: 0,
    eligible: 0,
    verified: 0,
    review: 0,
    blocked: 0,
    notRun: 0,
    byVertical: {}
  };

  (records || []).forEach(function(record) {
    counts.total++;

    if (!counts.byVertical[record.vertical]) {
      counts.byVertical[record.vertical] = {
        total:0, eligible:0, verified:0, review:0, blocked:0, notRun:0
      };
    }

    const bucket = counts.byVertical[record.vertical];
    bucket.total++;

    if (record.step2 && record.step2.disposition === 'ELIGIBLE') {
      counts.eligible++;
      bucket.eligible++;
    }

    const disposition = record.step3 ? record.step3.disposition : 'NOT_RUN';

    if (disposition === 'VERIFIED') {
      counts.verified++;
      bucket.verified++;
    } else if (disposition === 'REVIEW') {
      counts.review++;
      bucket.review++;
    } else if (disposition === 'BLOCKED') {
      counts.blocked++;
      bucket.blocked++;
    } else {
      counts.notRun++;
      bucket.notRun++;
    }
  });

  return counts;
}

function tmv3_step3WriteOperatorViews_(records) {
  const written = {};

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    const verticalRecords = (records || [])
      .filter(function(record) { return record.vertical === vertical; })
      .sort(function(a, b) {
        const aTime = a.start instanceof Date ? a.start.getTime() : 0;
        const bTime = b.start instanceof Date ? b.start.getTime() : 0;
        if (aTime !== bTime) return aTime - bTime;

        if (vertical === 'Service') {
          const techCompare = tmv3_clean_(a.technician)
            .localeCompare(tmv3_clean_(b.technician));
          if (techCompare !== 0) return techCompare;
        }

        return tmv3_clean_(a.title).localeCompare(tmv3_clean_(b.title));
      });

    const rows = verticalRecords.map(tmv3_step3OperatorRow_);
    const sheetName = TMV3.VERTICALS[vertical].sheet;
    const priorLastRow = tmv3_sheet_(sheetName).getLastRow();

    tmv3_replaceRows_(sheetName, TMV3_STEP1_HEADERS.slice(), rows);
    tmv3_step1FormatOperatorView_(
      sheetName,
      vertical,
      verticalRecords,
      priorLastRow
    );

    written[vertical] = rows.length;
  });

  return written;
}

function tmv3_step3OperatorRow_(record) {
  const row = tmv3_step2OperatorRow_(record);
  const step2 = record.step2 || {};
  const step3 = record.step3 || {};

  if (step2.disposition !== 'ELIGIBLE') {
    return row;
  }

  const anchor = step3.anchor || {};
  const disposition = step3.disposition || 'REVIEW';

  row[4] = tmv3_step1CalendarChecklist_(record) +
    '\n✅ Step 2 Eligible' +
    '\n' +
    (disposition === 'VERIFIED'
      ? '✅ Business Anchor'
      : disposition === 'BLOCKED'
        ? '❌ Business Anchor'
        : '⚠ Business Anchor');

  row[5] = tmv3_step3AnchorSummary_(record, anchor);
  row[6] = 'NOT RUN — STEP 3';
  row[7] = 'NOT RUN — STEP 3';
  row[8] = 'NOT RUN — STEP 3';
  row[9] = 'STEP 3 — ' + disposition;

  const detail = [step3.reason || ''];

  if (step3.evidence && step3.evidence.length) {
    detail.push('Evidence: ' + step3.evidence.join(', '));
  }

  (step3.warnings || []).forEach(function(warning) {
    detail.push('Attention: ' + warning);
  });

  if (record.vertical === 'PreInspection' && anchor.orderId) {
    detail.push('SO is corroborating evidence only; it is not attached to the PreInspection Task.');
  }

  row[10] = detail.filter(Boolean).join('\n');

  return row;
}

function tmv3_step3AnchorSummary_(record, anchor) {
  if (!anchor || !Object.keys(anchor).length) return 'Anchor: —';

  const lines = [];

  if (anchor.customerNumber) {
    lines.push(
      'Customer #: ' +
      anchor.customerNumber +
      (anchor.customerName ? ' · ' + anchor.customerName : '')
    );
  }

  if (anchor.customerId) lines.push('Customer ID: ' + anchor.customerId);

  if (anchor.orderNumber) {
    lines.push(
      (record.vertical === 'Service' ? 'Work Order / SO #: ' : 'SO #: ') +
      anchor.orderNumber
    );
  }

  if (anchor.orderId) lines.push('Order ID: ' + anchor.orderId);
  if (anchor.orderType) lines.push('Order Type: ' + anchor.orderType);

  if (anchor.orderLocationIdEvidence) {
    lines.push('Order Location ID (evidence): ' + anchor.orderLocationIdEvidence);
  }

  return lines.length ? lines.join('\n') : 'Anchor: —';
}

function tmv3_step3Verify_(records, step2Snapshot) {
  const step2Keys = {};
  const step3Keys = {};
  let missingDecision = 0;
  let eligibleWithoutDecision = 0;
  let nonEligibleResolved = 0;

  (step2Snapshot || []).forEach(function(record) {
    step2Keys[record.logicalKey] = true;
  });

  (records || []).forEach(function(record) {
    step3Keys[record.logicalKey] = true;

    const s2 = record.step2 ? record.step2.disposition : '';
    const s3 = record.step3 ? record.step3.disposition : '';

    if (!s3) missingDecision++;

    if (
      s2 === 'ELIGIBLE' &&
      ['VERIFIED','REVIEW','BLOCKED'].indexOf(s3) === -1
    ) {
      eligibleWithoutDecision++;
    }

    if (
      s2 !== 'ELIGIBLE' &&
      s3 !== 'NOT_RUN'
    ) {
      nonEligibleResolved++;
    }
  });

  const missingFromStep3 = Object.keys(step2Keys).filter(function(key) {
    return !step3Keys[key];
  });

  const introducedByStep3 = Object.keys(step3Keys).filter(function(key) {
    return !step2Keys[key];
  });

  const expectedRows = {};
  const actualRows = {};

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    expectedRows[vertical] = records.filter(function(record) {
      return record.vertical === vertical;
    }).length;

    actualRows[vertical] = Math.max(
      0,
      tmv3_sheet_(TMV3.VERTICALS[vertical].sheet).getLastRow() - 1
    );
  });

  const rowCountsMatch = Object.keys(expectedRows).every(function(vertical) {
    return expectedRows[vertical] === actualRows[vertical];
  });

  return {
    pass:
      missingDecision === 0 &&
      eligibleWithoutDecision === 0 &&
      nonEligibleResolved === 0 &&
      missingFromStep3.length === 0 &&
      introducedByStep3.length === 0 &&
      rowCountsMatch,
    step2Records: (step2Snapshot || []).length,
    step3Records: (records || []).length,
    missingDecision: missingDecision,
    eligibleWithoutDecision: eligibleWithoutDecision,
    nonEligibleResolved: nonEligibleResolved,
    missingFromStep3: missingFromStep3.length,
    introducedByStep3: introducedByStep3.length,
    expectedRows: expectedRows,
    actualRows: actualRows
  };
}


// Original: 38_Step4_Identity
/************************************************************
 * TM V3 — STAGE 4 CUSTOMER / CONTACT / LOCATION
 *
 * Contract:
 * - Consumes the exact Step-3 record set.
 * - Resolves Customer + Customer-owned Location.
 * - Resolves Contact when deterministic; Contact no-match is non-blocking,
 *   ambiguity or read failure is REVIEW.
 * - Does NOT resolve or mutate Tasks.
 * - Does NOT write Google Calendar.
 * - Does NOT write Striven.
 ************************************************************/

function tmv3_step4RefreshIdentitySources_() {
  tmv3_assertShadow_();
  tmv3_resetRuntimeMetrics_();

  const customers =
    tmv3_reportRows_(TMV3.PROPERTIES.CUSTOMERS)
      .map(tmv3_normalizeCustomer_);

  const customerNumberToId = {};
  customers.forEach(function(row) {
    const id = tmv3_clean_(row[0]);
    const number = tmv3_clean_(row[1]);
    if (id && number) customerNumberToId[number] = id;
  });

  const locations =
    tmv3_reportRows_(TMV3.PROPERTIES.LOCATIONS)
      .map(function(row) {
        return tmv3_normalizeLocation_(row, customerNumberToId);
      });

  tmv3_assertSourceReplacementSafe_(
    'Stage 4 Customers',
    customers,
    TMV3.SHEETS.CUSTOMERS
  );
  tmv3_assertSourceReplacementSafe_(
    'Stage 4 Locations',
    locations,
    TMV3.SHEETS.LOCATIONS
  );

  tmv3_replaceRows_(
    TMV3.SHEETS.CUSTOMERS,
    [
      'Customer ID',
      'Customer Number',
      'Name',
      'Primary Phone',
      'Primary Email',
      'Fingerprint'
    ],
    customers
  );

  tmv3_replaceRows_(
    TMV3.SHEETS.LOCATIONS,
    [
      'Location ID',
      'Customer ID',
      'Address 1',
      'Address 2',
      'City',
      'Province',
      'Postal Code',
      'Primary Phone',
      'Fingerprint'
    ],
    locations
  );

  const result = {
    customers: customers.length,
    locations: locations.length,
    contacts: 'DEFERRED_TO_WRITE_GATE_NO_MAPPING_API_READS',
    api: tmv3_runtimeMetrics_()
  };

  tmv3_audit_(
    'SYSTEM','','','STEP4_REFRESH_IDENTITY_SOURCES','PASS',
    JSON.stringify(result)
  );

  return result;
}

function tmv3_step4IdentityIndex_() {
  const customers = tmv3_rows_(TMV3.SHEETS.CUSTOMERS);
  const locations = tmv3_rows_(TMV3.SHEETS.LOCATIONS);
  const orders = tmv3_rows_(TMV3.SHEETS.ORDERS);
  const tasks = tmv3_rows_(TMV3.SHEETS.TASKS);

  const refs = {
    customers: customers,
    locations: locations,
    contacts: [],
    orders: orders,
    tasks: tasks,
    customerById: {},
    customerByNumber: {},
    customerByPhone: {},
    locationsByCustomer: {},
    locationById: {},
    locationsByNormalizedAddress: {},
    locationsByPostal: {},
    locationsByStreetNo: {},
    orderById: {},
    taskById: {}
  };

  customers.forEach(function(row) {
    const id = tmv3_clean_(row['Customer ID']);
    const number = tmv3_clean_(row['Customer Number']);
    const phone = tmv3_phone10_(row['Primary Phone']);

    if (id) refs.customerById[id] = row;
    if (number) refs.customerByNumber[number] = row;

    if (phone) {
      if (!refs.customerByPhone[phone]) refs.customerByPhone[phone] = [];
      refs.customerByPhone[phone].push(row);
    }
  });

  locations.forEach(function(row) {
    const id = tmv3_clean_(row['Location ID']);
    const customerId = tmv3_clean_(row['Customer ID']);

    if (id) refs.locationById[id] = row;

    if (customerId) {
      if (!refs.locationsByCustomer[customerId]) {
        refs.locationsByCustomer[customerId] = [];
      }
      refs.locationsByCustomer[customerId].push(row);
    }

    const fullAddress = tmv3_locationFullAddress_(row);
    const normalizedAddress = tmv3_normalizeAddress_(fullAddress);
    const addressParts = tmv3_addressParts_(fullAddress);

    if (normalizedAddress) {
      if (!refs.locationsByNormalizedAddress[normalizedAddress]) {
        refs.locationsByNormalizedAddress[normalizedAddress] = [];
      }
      refs.locationsByNormalizedAddress[normalizedAddress].push(row);
    }

    if (addressParts.postal) {
      if (!refs.locationsByPostal[addressParts.postal]) {
        refs.locationsByPostal[addressParts.postal] = [];
      }
      refs.locationsByPostal[addressParts.postal].push(row);
    }

    if (addressParts.streetNo) {
      if (!refs.locationsByStreetNo[addressParts.streetNo]) {
        refs.locationsByStreetNo[addressParts.streetNo] = [];
      }
      refs.locationsByStreetNo[addressParts.streetNo].push(row);
    }
  });

  orders.forEach(function(row) {
    const id = tmv3_clean_(row['Order ID']);
    if (id) refs.orderById[id] = row;
  });

  tasks.forEach(function(row) {
    const id = tmv3_clean_(row['Task ID']);
    if (id) refs.taskById[id] = row;
  });

  return refs;
}

function tmv3_step4AnchorPendingIdentityAllowed_(step3) {
  const code = tmv3_clean_(step3 && step3.code);

  const allowed = {
    ANCHOR_ORDER_NUMBER_MISSING:true,
    ANCHOR_ORDER_NOT_FOUND:true,
    ANCHOR_TRANSACTION_TYPE_MISMATCH:true,
    ANCHOR_ORDER_AMBIGUOUS:true,
    PREINSPECTION_BUSINESS_ANCHOR_UNRESOLVED:true,
    PREINSPECTION_SO_AMBIGUOUS:true
  };

  return (
    step3 &&
    step3.disposition !== 'VERIFIED' &&
    allowed[code] === true
  );
}

function tmv3_step4IdentityRecords_(step3Records, refs) {
  return (step3Records || []).map(function(record) {
    const step3 = record.step3 || {};

    if (step3.disposition === 'VERIFIED') {
      return Object.assign({}, record, {
        step4: tmv3_step4ResolveIdentity_(record, refs)
      });
    }

    if (tmv3_step4AnchorPendingIdentityAllowed_(step3)) {
      return Object.assign({}, record, {
        step4: tmv3_step4ResolveIdentityWithoutAnchor_(
          record,
          refs,
          step3
        )
      });
    }

    return Object.assign({}, record, {
      step4: {
        disposition: 'NOT_RUN',
        code: 'STEP3_' + (step3.disposition || 'UNKNOWN'),
        reason: 'Identity resolution did not run because Step 3 is not VERIFIED.',
        customer: null,
        location: null,
        contact: null,
        contactStatus: 'NOT_RUN',
        evidence: [],
        warnings: []
      }
    });
  });
}

function tmv3_step4ResolveIdentityWithoutAnchor_(record, refs, step3) {
  let resolved = null;

  if (record.vertical === 'Service') {
    resolved = tmv3_step4ResolveServiceIdentity_(
      record,
      refs,
      {}
    );
  } else if (record.vertical === 'PreInspection') {
    resolved = tmv3_step4ResolvePreInspectionIdentity_(
      record,
      refs,
      {}
    );
  } else {
    resolved = tmv3_step4ResolveCalendarIdentityOnly_(
      record,
      refs
    );
  }

  if (!resolved || resolved.disposition !== 'VERIFIED') {
    return resolved || tmv3_step4Decision_(
      'REVIEW',
      'IDENTITY_ONLY_RESOLUTION_FAILED',
      'Business anchor is pending and Calendar identity could not be resolved deterministically.',
      null,
      null,
      null,
      'NOT_RESOLVED',
      [],
      [tmv3_clean_(step3 && step3.reason)]
    );
  }

  const cfg = TMV3.VERTICALS[record.vertical] || {};
  const label = cfg.orderLabel ||
    (record.vertical === 'Service' ? 'Work Order' : 'Sales Order');

  if (
    record.vertical === 'PreInspection' &&
    cfg.orderRequired !== true
  ) {
    return tmv3_step4Decision_(
      'VERIFIED',
      'PREINSPECTION_IDENTITY_VERIFIED_ORDER_OPTIONAL',
      'PreInspection Customer and Customer-owned Location are verified. ' +
        label +
        ' corroboration is unresolved or unavailable, but it is not required before Task resolution.',
      resolved.customer,
      resolved.location,
      resolved.contact,
      resolved.contactStatus,
      (resolved.evidence || []).concat([
        'IDENTITY_VERIFIED_WITHOUT_BUSINESS_ANCHOR',
        'PREINSPECTION_ORDER_OPTIONAL_STAGE5_ALLOWED'
      ]),
      (resolved.warnings || []).concat([
        tmv3_clean_(step3 && step3.reason),
        label + ' is corroborating evidence only for PreInspection; Stage 5 may continue.'
      ]).filter(Boolean)
    );
  }

  return tmv3_step4Decision_(
    'IDENTITY_ONLY',
    'IDENTITY_VERIFIED_ANCHOR_PENDING',
    'Customer and Customer-owned Location are verified. ' +
      label +
      ' remains unresolved and is still required before Task resolution.',
    resolved.customer,
    resolved.location,
    resolved.contact,
    resolved.contactStatus,
    (resolved.evidence || []).concat([
      'IDENTITY_VERIFIED_WITHOUT_BUSINESS_ANCHOR'
    ]),
    (resolved.warnings || []).concat([
      tmv3_clean_(step3 && step3.reason),
      label + ' must be resolved before Stage 5.'
    ]).filter(Boolean)
  );
}

function tmv3_step4ResolveCalendarIdentityOnly_(record, refs) {
  const recovery = tmv3_step4ResolveCustomerFromCalendarEvidence_(
    record,
    refs
  );

  if (recovery.status !== 'MATCHED') {
    return tmv3_step4Decision_(
      recovery.status === 'BLOCKED' ? 'BLOCKED' : 'REVIEW',
      recovery.errorCode || 'IDENTITY_CUSTOMER_UNRESOLVED',
      recovery.reason || 'Customer could not be resolved deterministically from Calendar evidence.',
      recovery.customer || null,
      recovery.location || null,
      null,
      'NOT_RESOLVED',
      recovery.evidence || [],
      []
    );
  }

  const customer = recovery.customer;
  let location = recovery.location || null;
  let locationEvidence = recovery.evidence || [];
  let locationStatus = location ? 'MATCHED' : '';

  if (!location) {
    const locationResult = tmv3_step4ResolveLocation_(
      record,
      refs,
      customer,
      ''
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
        (recovery.evidence || []).concat(
          locationResult.evidence || []
        ),
        []
      );
    }

    location = locationResult.location || null;
    locationStatus = locationResult.status;
    locationEvidence = (recovery.evidence || []).concat(
      locationResult.evidence || []
    );
  }

  const contactResult = tmv3_step4ResolveContact_(
    record,
    customer,
    ''
  );

  const warnings = [];
  if (locationStatus === 'CREATE_REQUIRED') {
    warnings.push(
      'Location will need to be created on the verified Customer before any Task write.'
    );
  }
  if (contactResult.status === 'REVIEW') {
    warnings.push(
      contactResult.reason || 'Contact requires review.'
    );
  } else if (contactResult.status === 'NO_MATCH') {
    warnings.push(
      'No deterministic Contact match; Customer and Location identity remain usable.'
    );
  }

  return tmv3_step4Decision_(
    'VERIFIED',
    'CALENDAR_IDENTITY_VERIFIED',
    'Customer and Customer-owned Location are verified from deterministic Calendar evidence.',
    customer,
    location,
    contactResult.status === 'MATCHED'
      ? contactResult.contact
      : null,
    contactResult.status,
    locationEvidence.concat(contactResult.evidence || []),
    warnings
  );
}

function tmv3_step4ResolveIdentity_(record, refs) {
  const step3 = record.step3 || {};
  const anchor = step3.anchor || {};

  if (record.vertical === 'Service') {
    return tmv3_step4ResolveServiceIdentity_(record, refs, anchor);
  }

  if (record.vertical === 'PreInspection') {
    return tmv3_step4ResolvePreInspectionIdentity_(record, refs, anchor);
  }

  return tmv3_step4ResolveOrderIdentity_(record, refs, anchor);
}

function tmv3_step4ResolveOrderIdentity_(record, refs, anchor) {
  const evidence = [];
  const warnings = [];

  const customerId = tmv3_clean_(anchor.customerId);
  const customer = customerId ? refs.customerById[customerId] : null;

  if (!customer) {
    return tmv3_step4Decision_(
      'BLOCKED',
      'CUSTOMER_FROM_ORDER_MISSING',
      'Step-3 Order anchor does not resolve to a current Customer.',
      null,
      null,
      null,
      'NOT_RESOLVED',
      evidence,
      warnings
    );
  }

  evidence.push('CUSTOMER_FROM_VERIFIED_ORDER');

  const locationResult = tmv3_step4ResolveLocation_(
    record,
    refs,
    customer,
    anchor.orderLocationIdEvidence
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

  const contactResult = tmv3_step4ResolveContact_(
    record,
    customer,
    anchor.orderContactIdEvidence
  );

  if (contactResult.status === 'REVIEW') {
    return tmv3_step4Decision_(
      'REVIEW',
      contactResult.errorCode,
      contactResult.reason,
      customer,
      locationResult.location,
      null,
      'REVIEW',
      evidence.concat(locationResult.evidence || [], contactResult.evidence || []),
      warnings
    );
  }

  if (contactResult.status === 'NO_MATCH') {
    warnings.push('No deterministic Contact match; Customer and Location are still verified.');
  }

  return tmv3_step4Decision_(
    'VERIFIED',
    'ORDER_IDENTITY_VERIFIED',
    'Customer and Customer-owned Location are verified from the Step-3 business anchor.',
    customer,
    locationResult.location || null,
    contactResult.contact || null,
    contactResult.status,
    evidence.concat(locationResult.evidence || [], contactResult.evidence || []),
    warnings
  );
}





function tmv3_step4ResolveCustomerFromCalendarEvidence_(record, refs) {
  const explicit = tmv3_clean_(record.customerNumber);
  const exactCustomer = explicit
    ? (refs.customerByNumber[explicit] || null)
    : null;

  if (exactCustomer) {
    return {
      status: 'MATCHED',
      customer: exactCustomer,
      location: null,
      evidence: ['CUSTOMER_NUMBER_EXACT']
    };
  }

  // A Calendar Task relationship is stronger than a loose name/address
  // search, but only when the cached Task belongs to the intended vertical
  // and its Customer-owned Location matches the Calendar job site.
  const taskIdentity = tmv3_step4TaskIdentityCandidate_(
    record,
    refs
  );

  if (taskIdentity.status === 'MATCHED') {
    if (explicit) {
      taskIdentity.evidence = [
        'EXPLICIT_CUSTOMER_NUMBER_STALE_RECOVERED_BY_TASK'
      ].concat(taskIdentity.evidence || []);
    }
    return taskIdentity;
  }

  const phone = tmv3_phone10_(record.phone);
  const phoneMatches = phone
    ? (refs.customerByPhone[phone] || []).slice()
    : [];

  const addressCandidates = tmv3_step4AddressCandidates_(
    record.location,
    refs
  );

  const ownerIds = tmv3_unique_(
    addressCandidates.locations
      .map(function(location) {
        return tmv3_clean_(location['Customer ID']);
      })
      .filter(Boolean)
  );

  // An explicit Calendar Customer # that is not present in the current cache
  // may be stale/incorrect. Do not silently ignore it; recover only when
  // independent evidence proves exactly one different Customer.
  if (explicit) {
    if (phoneMatches.length === 1) {
      const phoneCustomer = phoneMatches[0];
      const phoneCustomerId = tmv3_clean_(
        phoneCustomer['Customer ID']
      );
      const ownedByPhoneCustomer =
        addressCandidates.locations.filter(function(location) {
          return (
            tmv3_clean_(location['Customer ID']) ===
            phoneCustomerId
          );
        });

      if (
        ownedByPhoneCustomer.length &&
        (
          addressCandidates.matchType === 'EXACT' ||
          addressCandidates.matchType === 'STRONG'
        )
      ) {
        return {
          status: 'MATCHED',
          customer: phoneCustomer,
          location:
            ownedByPhoneCustomer.length === 1
              ? ownedByPhoneCustomer[0]
              : null,
          evidence: [
            'EXPLICIT_CUSTOMER_NUMBER_STALE',
            'CUSTOMER_PRIMARY_PHONE_EXACT',
            'CUSTOMER_' +
              addressCandidates.matchType +
              '_ADDRESS_CORROBORATES_PHONE'
          ]
        };
      }
    }

    if (ownerIds.length === 1 && refs.customerById[ownerIds[0]]) {
      const addressCustomer = refs.customerById[ownerIds[0]];
      const nameMatch = tmv3_nameCorroboratesCustomer_(
        record.title + ' ' +
          tmv3_identityDescriptionForRecord_(record),
        addressCustomer['Name']
      );
      const surnameMatch =
        addressCandidates.matchType === 'EXACT' &&
        tmv3_step4SurnameCorroboratesCustomer_(
          record,
          addressCustomer
        );

      if (nameMatch || surnameMatch) {
        const owned =
          addressCandidates.locations.filter(function(location) {
            return (
              tmv3_clean_(location['Customer ID']) ===
              ownerIds[0]
            );
          });

        return {
          status: 'MATCHED',
          customer: addressCustomer,
          location: owned.length === 1 ? owned[0] : null,
          evidence: [
            'EXPLICIT_CUSTOMER_NUMBER_STALE',
            surnameMatch
              ? 'EXACT_ADDRESS_SAME_SURNAME_RECOVERY'
              : 'CUSTOMER_NAME_AND_ADDRESS_RECOVERY'
          ]
        };
      }
    }

    return {
      status: 'REVIEW',
      errorCode: 'EXPLICIT_CUSTOMER_NUMBER_NOT_FOUND',
      reason:
        'Explicit Calendar Customer # ' +
        explicit +
        ' does not resolve to a Customer and no independent phone/address proof identifies one replacement Customer.',
      evidence: []
    };
  }

  if (phoneMatches.length === 1) {
    return {
      status: 'MATCHED',
      customer: phoneMatches[0],
      location: null,
      evidence: ['CUSTOMER_PRIMARY_PHONE_EXACT']
    };
  }

  if (ownerIds.length) {
    // Shared phone numbers are not an immediate failure. If the Calendar
    // address narrows the same phone to one Customer, that is deterministic.
    if (phoneMatches.length > 1) {
      const phoneIds = phoneMatches.map(function(customer) {
        return tmv3_clean_(customer['Customer ID']);
      });

      const intersect = ownerIds.filter(function(ownerId) {
        return phoneIds.indexOf(ownerId) !== -1;
      });

      if (intersect.length === 1 && refs.customerById[intersect[0]]) {
        const customer = refs.customerById[intersect[0]];
        const owned = addressCandidates.locations.filter(function(location) {
          return tmv3_clean_(location['Customer ID']) === intersect[0];
        });

        return {
          status: 'MATCHED',
          customer: customer,
          location: owned.length === 1 ? owned[0] : null,
          evidence: [
            'SHARED_PHONE_RESOLVED_BY_' +
              addressCandidates.matchType +
              '_ADDRESS'
          ]
        };
      }
    }

    if (ownerIds.length === 1 && refs.customerById[ownerIds[0]]) {
      const customer = refs.customerById[ownerIds[0]];
      const corroboration = tmv3_step4CorroborateCustomer_(
        record,
        customer
      );
      const surnameMatch =
        addressCandidates.matchType === 'EXACT' &&
        tmv3_step4SurnameCorroboratesCustomer_(
          record,
          customer
        );

      if (corroboration.matched || surnameMatch) {
        const owned = addressCandidates.locations.filter(function(location) {
          return tmv3_clean_(location['Customer ID']) === ownerIds[0];
        });

        return {
          status: 'MATCHED',
          customer: customer,
          location: owned.length === 1 ? owned[0] : null,
          evidence: [
            'CUSTOMER_FROM_' +
              addressCandidates.matchType +
              '_ADDRESS'
          ].concat(
            corroboration.evidence || [],
            surnameMatch
              ? ['EXACT_ADDRESS_SAME_SURNAME_HOUSEHOLD']
              : []
          )
        };
      }

      return {
        status: 'REVIEW',
        errorCode: 'ADDRESS_OWNER_NOT_CORROBORATED',
        reason:
          'Calendar address points to one Customer, but name/phone/surname evidence does not corroborate that household.',
        customer: customer,
        evidence: [
          addressCandidates.matchType + '_ADDRESS_OWNER_ONLY'
        ]
      };
    }

    if (ownerIds.length > 1) {
      const nameHits = ownerIds.filter(function(ownerId) {
        const customer = refs.customerById[ownerId];
        return customer && tmv3_nameCorroboratesCustomer_(
          record.title + ' ' +
            tmv3_identityDescriptionForRecord_(record),
          customer['Name']
        );
      });

      if (nameHits.length === 1 && refs.customerById[nameHits[0]]) {
        const customer = refs.customerById[nameHits[0]];
        const owned = addressCandidates.locations.filter(function(location) {
          return tmv3_clean_(location['Customer ID']) === nameHits[0];
        });

        return {
          status: 'MATCHED',
          customer: customer,
          location: owned.length === 1 ? owned[0] : null,
          evidence: [
            'AMBIGUOUS_ADDRESS_RESOLVED_BY_CUSTOMER_NAME'
          ]
        };
      }

      if (addressCandidates.matchType === 'EXACT') {
        const surnameHits = ownerIds.filter(function(ownerId) {
          const customer = refs.customerById[ownerId];
          return customer &&
            tmv3_step4SurnameCorroboratesCustomer_(
              record,
              customer
            );
        });

        if (
          surnameHits.length === 1 &&
          refs.customerById[surnameHits[0]]
        ) {
          const customer = refs.customerById[surnameHits[0]];
          const owned =
            addressCandidates.locations.filter(function(location) {
              return (
                tmv3_clean_(location['Customer ID']) ===
                surnameHits[0]
              );
            });

          return {
            status: 'MATCHED',
            customer: customer,
            location: owned.length === 1 ? owned[0] : null,
            evidence: [
              'AMBIGUOUS_EXACT_ADDRESS_RESOLVED_BY_SURNAME'
            ]
          };
        }
      }

      return {
        status: 'REVIEW',
        errorCode: 'AMBIGUOUS_ADDRESS_OWNER',
        reason:
          'Calendar address is associated with multiple Customers and no unique phone/name/surname corroboration wins.',
        evidence: []
      };
    }
  }

  if (phoneMatches.length > 1) {
    const nameHits = phoneMatches.filter(function(customer) {
      return tmv3_nameCorroboratesCustomer_(
        record.title + ' ' +
          tmv3_identityDescriptionForRecord_(record),
        customer['Name']
      );
    });

    if (nameHits.length === 1) {
      return {
        status: 'MATCHED',
        customer: nameHits[0],
        location: null,
        evidence: ['SHARED_PHONE_RESOLVED_BY_CUSTOMER_NAME']
      };
    }

    return {
      status: 'REVIEW',
      errorCode: 'AMBIGUOUS_CUSTOMER_PHONE',
      reason:
        'Calendar phone matches multiple Customers and address/name evidence does not identify one.',
      evidence: []
    };
  }

  return {
    status: 'REVIEW',
    errorCode: 'CUSTOMER_UNRESOLVED',
    reason:
      'No deterministic Customer match was proven from Customer #, cached Task, phone, or Calendar address.',
    evidence: []
  };
}

function tmv3_step4TaskMatchesVertical_(task, vertical) {
  const cfg = TMV3.VERTICALS[vertical] || {};
  const taskTypeId = Number(task && task['Task Type ID'] || 0);

  if (
    Number(cfg.taskTypeId || 0) > 0 &&
    taskTypeId === Number(cfg.taskTypeId)
  ) {
    return true;
  }

  const hay = tmv3_norm_(
    tmv3_clean_(task && task['Task Type']) +
    ' ' +
    tmv3_clean_(task && task['Name'])
  );

  return (cfg.taskTypeNames || []).some(function(name) {
    const needle = tmv3_norm_(name);
    return needle && hay.indexOf(needle) !== -1;
  });
}

function tmv3_step4TaskIdentityCandidate_(record, refs) {
  const taskId = tmv3_clean_(
    record &&
    (
      record.existingTaskId ||
      record.taskNumber
    )
  );

  if (!taskId) return { status:'NO_MATCH', evidence:[] };

  const task = refs.taskById && refs.taskById[taskId];
  if (!task) {
    return {
      status:'NO_MATCH',
      evidence:['CALENDAR_TASK_ID_NOT_IN_TASK_CACHE']
    };
  }

  if (!tmv3_step4TaskMatchesVertical_(task, record.vertical)) {
    return {
      status:'NO_MATCH',
      evidence:['CALENDAR_TASK_WRONG_VERTICAL']
    };
  }

  const customerId = tmv3_clean_(task['Customer ID']);
  const locationId = tmv3_clean_(task['Location ID']);
  const customer = customerId
    ? refs.customerById[customerId]
    : null;
  const location = locationId
    ? refs.locationById[locationId]
    : null;

  if (!customer || !location) {
    return {
      status:'NO_MATCH',
      evidence:['CALENDAR_TASK_IDENTITY_INCOMPLETE']
    };
  }

  if (tmv3_clean_(location['Customer ID']) !== customerId) {
    return {
      status:'NO_MATCH',
      evidence:['CALENDAR_TASK_LOCATION_CUSTOMER_CONFLICT']
    };
  }

  const addressMatch = tmv3_resolveOwnedLocation_(
    customer,
    record.location,
    [location]
  );

  if (addressMatch.status !== 'MATCHED') {
    return {
      status:'NO_MATCH',
      evidence:['CALENDAR_TASK_LOCATION_ADDRESS_MISMATCH']
    };
  }

  return {
    status:'MATCHED',
    customer:customer,
    location:location,
    evidence:[
      'CUSTOMER_FROM_CACHED_CALENDAR_TASK',
      'LOCATION_FROM_CACHED_CALENDAR_TASK'
    ]
  };
}

function tmv3_step4SurnameCorroboratesCustomer_(record, customer) {
  const customerName = tmv3_norm_(
    customer && customer['Name']
  );

  if (!customerName) return false;

  // Do not turn business/common suffixes into pseudo-surnames.
  if (
    /\b(ltd|inc|corp|corporation|company|design|build|homes|store)\b/.test(
      customerName
    )
  ) {
    return false;
  }

  const parts = customerName
    .split(' ')
    .filter(function(part) {
      return part && part.length >= 3 && part !== 'and';
    });

  if (!parts.length) return false;

  const surname = parts[parts.length - 1];
  const hay = tmv3_norm_(
    tmv3_clean_(record && record.title) +
    ' ' +
    tmv3_identityDescriptionForRecord_(record || {})
  );
  const tokens = hay.split(' ').filter(Boolean);

  return tokens.indexOf(surname) !== -1;
}

function tmv3_step4AddressCandidates_(calendarAddress, refs) {
  const target = tmv3_normalizeAddress_(calendarAddress);

  if (!target) {
    return {
      matchType: 'NO',
      locations: []
    };
  }

  const exactIndex =
    refs &&
    refs.locationsByNormalizedAddress &&
    refs.locationsByNormalizedAddress[target];

  const exact = exactIndex
    ? exactIndex.slice()
    : (refs.locations || []).filter(function(location) {
        return tmv3_normalizeAddress_(
          tmv3_locationFullAddress_(location)
        ) === target;
      });

  if (exact.length) {
    return {
      matchType: 'EXACT',
      locations: exact
    };
  }

  const targetParts = tmv3_addressParts_(calendarAddress);
  let candidates = refs.locations || [];

  if (
    targetParts.postal &&
    refs.locationsByPostal &&
    refs.locationsByPostal[targetParts.postal]
  ) {
    candidates =
      refs.locationsByPostal[targetParts.postal].slice();
  } else if (
    targetParts.streetNo &&
    refs.locationsByStreetNo &&
    refs.locationsByStreetNo[targetParts.streetNo]
  ) {
    candidates =
      refs.locationsByStreetNo[targetParts.streetNo].slice();
  }

  const strong = candidates.filter(function(location) {
    return tmv3_addressStrongMatch_(
      targetParts,
      tmv3_addressParts_(
        tmv3_locationFullAddress_(location)
      )
    );
  });

  return {
    matchType: strong.length ? 'STRONG' : 'NO',
    locations: strong
  };
}

function tmv3_step4CorroborateCustomer_(record, customer) {
  const evidence = [];
  // A named person after c/o is explicit account evidence, not a fuzzy company-name match.
  const careOf=String(customer['Name']||'').match(/\bc\s*\/\s*o\s+([A-Za-z][A-Za-z .'-]+)$/i);
  if(record.vertical==='PreInspection'&&careOf&&careOf[1].trim().split(/\s+/).length>=2&&
      tmv3_nameCorroboratesCustomer_(record.title+' '+tmv3_identityDescriptionForRecord_(record),careOf[1])) {
    return {matched:true,evidence:['CUSTOMER_CARE_OF_FULL_NAME_CORROBORATES_ADDRESS']};
  }


  if (
    tmv3_nameCorroboratesCustomer_(
      record.title + ' ' + tmv3_identityDescriptionForRecord_(record),
      customer['Name']
    )
  ) {
    evidence.push('CUSTOMER_NAME_CORROBORATES_ADDRESS');
    return { matched: true, evidence: evidence };
  }

  const eventPhones = tmv3_allPhones_(
    record.title + ' ' + tmv3_identityDescriptionForRecord_(record) + ' ' + record.location
  );

  const primary = tmv3_phone10_(customer['Primary Phone']);

  if (primary && eventPhones.indexOf(primary) !== -1) {
    evidence.push('CUSTOMER_PRIMARY_PHONE_CORROBORATES_ADDRESS');
    return { matched: true, evidence: evidence };
  }

  // Stage 4 is a planning/read-model stage. Do not spend one Striven API call
  // per Customer just to search Contacts. If Customer name/primary phone does
  // not corroborate the address from cached sources, leave the identity in
  // REVIEW. A later write-capable path may fresh-verify Contact ownership only
  // when a Contact is actually required for that mutation.
  evidence.push('CONTACT_CORROBORATION_DEFERRED_TO_WRITE_GATE');

  return { matched: false, evidence: evidence };
}

function tmv3_step4ResolveLocation_(record, refs, customer, preferredLocationId) {
  const customerId = tmv3_clean_(customer && customer['Customer ID']);
  const fallbackEvidence = [];

  if (!customerId) {
    return {
      status: 'BLOCKED',
      errorCode: 'LOCATION_CUSTOMER_REQUIRED',
      reason: 'Location cannot be resolved before Customer.',
      evidence: []
    };
  }

  if (preferredLocationId) {
    const preferred = refs.locationById[tmv3_clean_(preferredLocationId)] || null;

    if (preferred) {
      if (tmv3_clean_(preferred['Customer ID']) !== customerId) {
        return {
          status: 'BLOCKED',
          errorCode: 'ORDER_LOCATION_OWNERSHIP_CONFLICT',
          reason: 'Order Location resolves to a Location owned by a different Customer.',
          evidence: []
        };
      }

      return {
        status: 'MATCHED',
        location: preferred,
        evidence: ['LOCATION_FROM_BUSINESS_ANCHOR']
      };
    }

    fallbackEvidence.push('ORDER_LOCATION_ID_NOT_IN_LOCATION_SOURCE');
  }

  const owned = tmv3_resolveOwnedLocation_(
    customer,
    record.location,
    refs.locationsByCustomer[customerId] || []
  );

  if (owned.status === 'MATCHED') {
    return {
      status: 'MATCHED',
      location: owned.location,
      evidence: fallbackEvidence.concat(owned.evidence || [])
    };
  }

  // Prefer the unique address as written on the Calendar before resolving a
  // broad street-abbreviation tie. Never choose by label, ID or primary flag.
  if (record.vertical === 'PreInspection' &&
      /^AMBIGUOUS_(EXACT|STRONG)_LOCATION$/.test(owned.errorCode || '')) {
    const writtenKey = tmv3_step4WrittenAddressKey_(record.location);
    const written = (refs.locationsByCustomer[customerId] || []).filter(function(location) {
      return tmv3_clean_(location['Customer ID']) === customerId &&
        writtenKey &&
        tmv3_step4WrittenAddressKey_(tmv3_locationFullAddress_(location)) === writtenKey &&
        tmv3_addressStrongMatch_(tmv3_addressParts_(record.location),
          tmv3_addressParts_(tmv3_locationFullAddress_(location)));
    });
    if (written.length === 1) {
      return {status: 'MATCHED', location: written[0],
        evidence: fallbackEvidence.concat(['CUSTOMER_LOCATION_CALENDAR_WRITTEN_ADDRESS'])};
    }
  }

  // An explicit Calendar Task link is strong relationship evidence, but it is
  // only allowed to break an address tie when the linked Task is consistent
  // with the verified Order, the Task Location belongs to the same Customer,
  // and that one Location strongly matches the Calendar job-site address.
  const linkedTaskLocation = tmv3_step4LinkedTaskLocation_(
    record,
    refs,
    customer
  );

  if (linkedTaskLocation.status === 'MATCHED') {
    return {
      status: 'MATCHED',
      location: linkedTaskLocation.location,
      evidence: fallbackEvidence.concat(linkedTaskLocation.evidence || [])
    };
  }

  // Historical rule: once Customer identity is proven, a genuinely different
  // Calendar job-site address may be added to that Customer later.
  if (
    owned.errorCode === 'LOCATION_UNRESOLVED' &&
    tmv3_clean_(record.location)
  ) {
    return {
      status: 'CREATE_REQUIRED',
      errorCode: 'LOCATION_CREATE_REQUIRED',
      reason: 'Customer is verified and Calendar job-site address is not yet a Customer-owned Location.',
      proposedAddress: tmv3_clean_(record.location),
      evidence: fallbackEvidence.concat(['NEW_CUSTOMER_LOCATION_REQUIRED'])
    };
  }

  return {
    status: owned.status,
    errorCode: owned.errorCode,
    reason: preferredLocationId
      ? owned.reason + ' Order Location ID ' + preferredLocationId +
        ' was not directly present in the Customer Location source.'
      : owned.reason,
    evidence: fallbackEvidence.concat(owned.evidence || [])
  };
}

function tmv3_step4LinkedTaskLocation_(record, refs, customer) {
  const taskId = tmv3_clean_(record && record.existingTaskId);
  const customerId = tmv3_clean_(customer && customer['Customer ID']);

  if (!taskId || !customerId) {
    return { status: 'NO_MATCH', evidence: [] };
  }

  const task = refs.taskById && refs.taskById[taskId];
  if (!task) {
    return {
      status: 'NO_MATCH',
      evidence: ['EXPLICIT_TASK_ID_NOT_IN_TASK_SOURCE']
    };
  }

  const expectedOrderId = tmv3_clean_(
    record &&
    record.step3 &&
    record.step3.anchor &&
    record.step3.anchor.orderId
  );
  const taskOrderId = tmv3_clean_(task['Order ID']);

  if (
    expectedOrderId &&
    taskOrderId &&
    expectedOrderId !== taskOrderId
  ) {
    return {
      status: 'NO_MATCH',
      evidence: ['EXPLICIT_TASK_ORDER_CONFLICT']
    };
  }

  const locationId = tmv3_clean_(task['Location ID']);
  const location = locationId && refs.locationById
    ? refs.locationById[locationId]
    : null;

  if (!location) {
    return {
      status: 'NO_MATCH',
      evidence: ['EXPLICIT_TASK_LOCATION_NOT_IN_LOCATION_SOURCE']
    };
  }

  if (tmv3_clean_(location['Customer ID']) !== customerId) {
    return {
      status: 'NO_MATCH',
      evidence: ['EXPLICIT_TASK_LOCATION_CUSTOMER_CONFLICT']
    };
  }

  const corroborated = tmv3_resolveOwnedLocation_(
    customer,
    record.location,
    [location]
  );

  if (corroborated.status !== 'MATCHED') {
    return {
      status: 'NO_MATCH',
      evidence: ['EXPLICIT_TASK_LOCATION_DOES_NOT_MATCH_CALENDAR_ADDRESS']
    };
  }

  return {
    status: 'MATCHED',
    location: location,
    evidence: [
      'LOCATION_FROM_EXPLICIT_LINKED_TASK',
      'EXPLICIT_TASK_LOCATION_CUSTOMER_OWNED',
      'EXPLICIT_TASK_LOCATION_CALENDAR_ADDRESS_MATCH'
    ]
  };
}

function tmv3_step4ResolveContact_(record, customer, preferredContactId) {
  const customerId = tmv3_clean_(customer && customer['Customer ID']);

  if (!customerId) {
    return {
      status: 'NO_MATCH',
      contact: null,
      evidence: []
    };
  }

  // Order Contact is deterministic relationship evidence. Fresh ownership is
  // still re-verified immediately before any later write.
  if (preferredContactId) {
    return {
      status: 'MATCHED',
      contact: {
        'Contact ID': tmv3_clean_(preferredContactId),
        'Customer ID': customerId,
        'Name': '',
        'Phone': '',
        'Email': ''
      },
      evidence: ['CONTACT_ID_FROM_VERIFIED_ORDER']
    };
  }

  // Routine Stage-4 mapping must be cache/report driven. Source Contacts is not
  // a maintained master cache, so do not call Striven Contacts customer by
  // customer here. Contact is optional for identity planning unless the
  // verified Order already supplied a Contact ID. Any mutation that needs a
  // Contact must fresh-verify ownership in Step 7 immediately before writing.
  return {
    status: 'NO_MATCH',
    contact: null,
    evidence: ['CONTACT_LOOKUP_DEFERRED_TO_WRITE_GATE']
  };
}

function tmv3_step4Decision_(
  disposition,
  code,
  reason,
  customer,
  location,
  contact,
  contactStatus,
  evidence,
  warnings
) {
  return {
    disposition: disposition,
    code: code,
    reason: reason,
    customer: customer || null,
    location: location || null,
    contact: contact || null,
    contactStatus: contactStatus || 'NO_MATCH',
    locationStatus:
      location && tmv3_clean_(location['Location ID'])
        ? 'MATCHED'
        : ((evidence || []).indexOf('NEW_CUSTOMER_LOCATION_REQUIRED') !== -1
          ? 'CREATE_REQUIRED'
          : 'NOT_RESOLVED'),
    evidence: evidence || [],
    warnings: warnings || []
  };
}

function tmv3_step4IdentityRun(reason, refreshSources) {
  tmv3_assertShadow_();

  const sourceSummary =
    refreshSources === true
      ? tmv3_step4RefreshIdentitySources_()
      : {
          customers: tmv3_rows_(TMV3.SHEETS.CUSTOMERS).length,
          locations: tmv3_rows_(TMV3.SHEETS.LOCATIONS).length,
          contacts: 'DEFERRED_TO_WRITE_GATE_NO_MAPPING_API_READS',
          source: 'CACHED_IDENTITY_SOURCES'
        };

  const step2Snapshot = tmv3_step2CalendarRecords_();
  const step3Refs = tmv3_step3AnchorIndex_();
  const step3Records = tmv3_step3BusinessAnchorRecords_(
    step2Snapshot,
    step3Refs
  );

  const refs = tmv3_step4IdentityIndex_();
  const records = tmv3_step4IdentityRecords_(step3Records, refs);
  const counts = tmv3_step4Counts_(records);
  const write = tmv3_step4WriteOperatorViews_(records);
  const verify = tmv3_step4Verify_(records, step3Records);

  const result = {
    version: TMV3.VERSION,
    executionStage: tmv3_executionStage_(),
    stage: 4,
    status: verify.pass ? 'PASS' : 'REVIEW',
    mode: 'CUSTOMER_CONTACT_LOCATION_RESOLUTION_ONLY',
    reason: tmv3_clean_(reason || 'MANUAL'),
    sourceSummary: sourceSummary,
    counts: counts,
    write: write,
    verification: verify,
    taskResolutionPerformed: false,
    calendarWritesPerformed: false,
    strivenWritesPerformed: false
  };

  tmv3_audit_(
    'SYSTEM','','','STEP4_IDENTITY',
    result.status,
    JSON.stringify(result)
  );

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function tmv3_step4IdentityRunCached(reason) {
  return tmv3_step4IdentityRun(
    reason || 'CACHED_IDENTITY_SOURCES',
    false
  );
}

function tmv3_step4Counts_(records) {
  const counts = {
    total:0,
    step3Verified:0,
    verified:0,
    identityOnly:0,
    review:0,
    blocked:0,
    notRun:0,
    contactMatched:0,
    contactNoMatch:0,
    byVertical:{}
  };

  (records || []).forEach(function(record) {
    counts.total++;

    if (!counts.byVertical[record.vertical]) {
      counts.byVertical[record.vertical] = {
        total:0,
        step3Verified:0,
        verified:0,
        identityOnly:0,
        review:0,
        blocked:0,
        notRun:0
      };
    }

    const bucket = counts.byVertical[record.vertical];
    bucket.total++;

    if (record.step3 && record.step3.disposition === 'VERIFIED') {
      counts.step3Verified++;
      bucket.step3Verified++;
    }

    const disposition = record.step4 ? record.step4.disposition : 'NOT_RUN';

    if (disposition === 'VERIFIED') {
      counts.verified++;
      bucket.verified++;
    } else if (disposition === 'IDENTITY_ONLY') {
      counts.identityOnly++;
      bucket.identityOnly++;
    } else if (disposition === 'REVIEW') {
      counts.review++;
      bucket.review++;
    } else if (disposition === 'BLOCKED') {
      counts.blocked++;
      bucket.blocked++;
    } else {
      counts.notRun++;
      bucket.notRun++;
    }

    if (record.step4 && record.step4.contactStatus === 'MATCHED') {
      counts.contactMatched++;
    }

    if (record.step4 && record.step4.contactStatus === 'NO_MATCH') {
      counts.contactNoMatch++;
    }
  });

  return counts;
}

function tmv3_step4WriteOperatorViews_(records) {
  const written = {};

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    const verticalRecords = (records || [])
      .filter(function(record) { return record.vertical === vertical; })
      .sort(function(a, b) {
        const aTime = a.start instanceof Date ? a.start.getTime() : 0;
        const bTime = b.start instanceof Date ? b.start.getTime() : 0;
        if (aTime !== bTime) return aTime - bTime;

        if (vertical === 'Service') {
          const techCompare = tmv3_clean_(a.technician)
            .localeCompare(tmv3_clean_(b.technician));
          if (techCompare !== 0) return techCompare;
        }

        return tmv3_clean_(a.title).localeCompare(tmv3_clean_(b.title));
      });

    const rows = verticalRecords.map(tmv3_step4OperatorRow_);
    const sheetName = TMV3.VERTICALS[vertical].sheet;
    const priorLastRow = tmv3_sheet_(sheetName).getLastRow();

    tmv3_replaceRows_(sheetName, TMV3_STEP1_HEADERS.slice(), rows);
    tmv3_step1FormatOperatorView_(
      sheetName,
      vertical,
      verticalRecords,
      priorLastRow
    );

    written[vertical] = rows.length;
  });

  return written;
}

function tmv3_step4OperatorRow_(record) {
  const row = tmv3_step3OperatorRow_(record);
  const step3 = record.step3 || {};
  const step4 = record.step4 || {};

  if (!step4.disposition || step4.disposition === 'NOT_RUN') {
    return row;
  }

  const customer = step4.customer || null;
  const location = step4.location || null;
  const contact = step4.contact || null;
  const disposition = step4.disposition || 'REVIEW';

  row[4] = tmv3_step4SyncChecklist_(record, step4);
  row[5] = tmv3_step4IdentitySummary_(
    customer,
    location,
    contact,
    step4.contactStatus,
    step4.locationStatus
  );
  row[6] =
    tmv3_step4TaskColumnIdentity_(record, step4) +
    '\n\nTask not resolved yet';
  row[7] = '';
  row[8] = '';
  row[9] = 'STEP 4 — ' + disposition;

  const details = [step4.reason || ''];

  (step4.warnings || []).forEach(function(warning) {
    details.push('Attention: ' + warning);
  });

  row[10] = details.filter(Boolean).join('\n');

  return row;
}

function tmv3_step4SyncChecklist_(record, step4) {
  const customerOk = !!(
    step4.customer &&
    tmv3_clean_(step4.customer['Customer ID'])
  );

  const locationOk = !!(
    step4.location &&
    tmv3_clean_(step4.location['Location ID'])
  );
  const locationCreate = step4.locationStatus === 'CREATE_REQUIRED';

  return [
    'Customer ' + (customerOk ? '✅' : '⚠'),
    'Location ' + (locationOk ? '✅' : (locationCreate ? '🟡' : '⚠')),
    'Date / Time ⏳',
    'Assignee ⏳',
    'Notes ⏳',
    'Task Link ' + (record.existingTaskUrl ? '✅' : '⏳')
  ].join('\n');
}

function tmv3_step4TaskColumnIdentity_(record, step4) {
  step4 = step4 || {};

  const customer = step4.customer || null;
  const contact = step4.contact || null;
  const location = step4.location || null;
  const lines = [];

  if (customer) {
    lines.push(
      'Customer #' +
      (tmv3_clean_(customer['Customer Number']) || '—') +
      ' - ' +
      (tmv3_clean_(customer['Name']) || '—')
    );
  } else {
    lines.push('Customer —');
  }

  if (contact && tmv3_clean_(contact['Contact ID'])) {
    const contactPhone = tmv3_step1FormatPhone_(
      tmv3_clean_(contact['Phone']) ||
      tmv3_clean_(record && record.phone)
    );

    lines.push(
      'Contact ID ' +
      tmv3_clean_(contact['Contact ID']) +
      ' - ' +
      contactPhone
    );
  } else {
    lines.push(
      'Contact: ' +
      (step4.contactStatus === 'REVIEW'
        ? 'Review'
        : 'Not required / not resolved')
    );
  }

  if (location && tmv3_clean_(location['Location ID'])) {
    const address = tmv3_step4LocationDisplay_(location);
    lines.push(
      'Location: ' +
      tmv3_clean_(location['Location ID']) +
      (address ? ' ' + address : '')
    );
  } else if (step4.locationStatus === 'CREATE_REQUIRED') {
    lines.push('Location: Create required');
  } else {
    lines.push('Location: —');
  }

  return lines.join('\n\n');
}

function tmv3_step4LocationDisplay_(location) {
  const a1 = tmv3_clean_(location && location['Address 1']);
  const a2 = tmv3_clean_(location && location['Address 2']);
  const city = tmv3_clean_(location && location['City']);
  const province = tmv3_clean_(location && location['Province']);
  const postal = tmv3_clean_(location && location['Postal Code']);

  const parts = [];
  if (a1) parts.push(a1);

  const currentText = function() {
    return parts.join(' ').toLowerCase().replace(/[^a-z0-9]/g, '');
  };

  [a2, city, province, postal].forEach(function(part) {
    if (!part) return;
    const key = part.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (!key) return;
    if (currentText().indexOf(key) !== -1) return;
    parts.push(part);
  });

  return parts.join(', ');
}

function tmv3_step4IdentitySummary_(customer, location, contact, contactStatus, locationStatus) {
  const lines = [];

  if (customer) {
    lines.push(
      'Customer: #' +
      (tmv3_clean_(customer['Customer Number']) || '—') +
      ' · ' +
      (tmv3_clean_(customer['Name']) || '—')
    );
  } else {
    lines.push('Customer: —');
  }

  if (location) {
    lines.push(
      'Location ID: ' +
      (tmv3_clean_(location['Location ID']) || '—')
    );
  } else if (locationStatus === 'CREATE_REQUIRED') {
    lines.push('Location: Create required');
  } else {
    lines.push('Location ID: —');
  }

  if (contact && tmv3_clean_(contact['Contact ID'])) {
    lines.push('Contact ID: ' + tmv3_clean_(contact['Contact ID']));
  } else {
    lines.push(
      'Contact: ' +
      (contactStatus === 'REVIEW' ? 'Review' : 'Not required / not resolved')
    );
  }

  return lines.join('\n');
}

function tmv3_step4Verify_(records, step3Records) {
  const priorKeys = {};
  const currentKeys = {};
  let missingDecision = 0;
  let verifiedWithoutIdentity = 0;
  let nonVerifiedAdvanced = 0;

  (step3Records || []).forEach(function(record) {
    priorKeys[record.logicalKey] = true;
  });

  (records || []).forEach(function(record) {
    currentKeys[record.logicalKey] = true;

    const s3 = record.step3 ? record.step3.disposition : '';
    const s4 = record.step4 ? record.step4.disposition : '';

    if (!s4) missingDecision++;

    if (
      (s4 === 'VERIFIED' || s4 === 'IDENTITY_ONLY') &&
      (
        !record.step4.customer ||
        !tmv3_clean_(record.step4.customer['Customer ID']) ||
        (
          !record.step4.location ||
          !tmv3_clean_(record.step4.location['Location ID'])
        ) &&
        record.step4.locationStatus !== 'CREATE_REQUIRED'
      )
    ) {
      verifiedWithoutIdentity++;
    }

    if (s3 !== 'VERIFIED' && s4 !== 'NOT_RUN') {
      const pendingAnchorAllowed =
        tmv3_step4AnchorPendingIdentityAllowed_(record.step3) &&
        ['IDENTITY_ONLY','REVIEW','BLOCKED'].indexOf(s4) !== -1;

      if (!pendingAnchorAllowed) {
        nonVerifiedAdvanced++;
      }
    }

    if (
      s4 === 'IDENTITY_ONLY' &&
      !tmv3_step4AnchorPendingIdentityAllowed_(record.step3)
    ) {
      nonVerifiedAdvanced++;
    }
  });

  const missing = Object.keys(priorKeys).filter(function(key) {
    return !currentKeys[key];
  });

  const introduced = Object.keys(currentKeys).filter(function(key) {
    return !priorKeys[key];
  });

  const expectedRows = {};
  const actualRows = {};

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    expectedRows[vertical] = records.filter(function(record) {
      return record.vertical === vertical;
    }).length;

    actualRows[vertical] = Math.max(
      0,
      tmv3_sheet_(TMV3.VERTICALS[vertical].sheet).getLastRow() - 1
    );
  });

  const rowCountsMatch = Object.keys(expectedRows).every(function(vertical) {
    return expectedRows[vertical] === actualRows[vertical];
  });

  return {
    pass:
      missingDecision === 0 &&
      verifiedWithoutIdentity === 0 &&
      nonVerifiedAdvanced === 0 &&
      missing.length === 0 &&
      introduced.length === 0 &&
      rowCountsMatch,
    step3Records: (step3Records || []).length,
    step4Records: (records || []).length,
    missingDecision: missingDecision,
    verifiedWithoutIdentity: verifiedWithoutIdentity,
    nonVerifiedAdvanced: nonVerifiedAdvanced,
    missingFromStep4: missing.length,
    introducedByStep4: introduced.length,
    expectedRows: expectedRows,
    actualRows: actualRows
  };
}


// Preserve street words and unit information; ignore only formatting and a
// trailing country spelling absent from the location-cache schema.
function tmv3_step4WrittenAddressKey_(address) {
  return String(address || '').toLowerCase().trim()
    .replace(/(?:[,\s]+)(?:canada|ca)\.?\s*$/, '')
    .replace(/[.,]/g, '').replace(/\b([a-z]\d[a-z])\s+(\d[a-z]\d)\b/g, '$1$2')
    .replace(/\s+/g, ' ').trim();
}


// Original: 40_Engine
function tmv3_shadowRun() {
  tmv3_assertShadow_();

  const sourceSummary =
    tmv3_refreshSources();

  return tmv3_shadowMapFromCache_(
    sourceSummary
  );
}

function tmv3_shadowMapFromCache() {
  tmv3_assertShadow_();

  return tmv3_shadowMapFromCache_(
    tmv3_sourceCacheSummary_()
  );
}

function tmv3_sourceCacheSummary_() {
  return {
    customers: tmv3_rows_(TMV3.SHEETS.CUSTOMERS).length,
    locations: tmv3_rows_(TMV3.SHEETS.LOCATIONS).length,
    contacts: tmv3_rows_(TMV3.SHEETS.CONTACTS).length,
    orders: tmv3_rows_(TMV3.SHEETS.ORDERS).length,
    tasks: tmv3_rows_(TMV3.SHEETS.TASKS).length,
    source: 'CACHED'
  };
}

function tmv3_shadowMapFromCache_(sourceSummary) {
  tmv3_resetRuntimeMetrics_();

  if (
    typeof tmv3_executionStage_ === 'function' &&
    tmv3_executionStage_() < 6
  ) {
    throw new Error(
      'Execution stage ' + tmv3_executionStage_() +
      ' blocks Striven/customer/task mapping. Run the current Calendar stage instead.'
    );
  }

  const events =
    typeof tmv3_step2EligibleCalendarRecords_ === 'function'
      ? tmv3_step2EligibleCalendarRecords_()
      : tmv3_calendarRecords_();

  const refs =
    tmv3_referenceIndex_();

  const state =
    tmv3_eventStateIndex_();

  const resolved = [];

  events.forEach(function(e) {
    const records = tmv3_resolveEventRecords_(e, refs, state);
    Array.prototype.push.apply(resolved, records);
  });

  tmv3_writeOperatorViews_(
    resolved
  );

  tmv3_upsertState_(
    resolved
  );

  tmv3_updateOverview_(
    resolved
  );

  const regression =
    tmv3_updateRegression_(
      resolved
    );

  tmv3_audit_(
    'SYSTEM',
    '',
    '',
    'SHADOW_RUN',
    'PASS',
    'Events ' +
      resolved.length +
      '; Sources ' +
      JSON.stringify(sourceSummary || {}) +
      '; API ' +
      JSON.stringify(tmv3_runtimeMetrics_())
  );

  const morningOps =
    tmv3_refreshMorningOps();

  return {
    version:
      TMV3.VERSION,
    mode:
      TMV3.MODE,
    sources:
      sourceSummary,
    events:
      resolved.length,
    counts:
      tmv3_statusCounts_(
        resolved
      ),
    regression:
      regression,
    morningOps:
      morningOps
  };
}

function tmv3_referenceIndex_() {
  const customers =
    tmv3_rows_(
      TMV3.SHEETS.CUSTOMERS
    );

  const locations =
    tmv3_rows_(
      TMV3.SHEETS.LOCATIONS
    );

  const contacts =
    tmv3_rows_(
      TMV3.SHEETS.CONTACTS
    );

  const orders =
    tmv3_rows_(
      TMV3.SHEETS.ORDERS
    );

  const tasks =
    tmv3_rows_(
      TMV3.SHEETS.TASKS
    );

  const index = {
    customers: customers,
    locations: locations,
    contacts: contacts,
    orders: orders,
    tasks: tasks,
    customerById: {},
    customerByNumber: {},
    customerByPhone: {},
    locationsByCustomer: {},
    orderById: {},
    ordersByNumber: {},
    taskById: {},
    tasksByOrder: {},
    tasksByCustomer: {}
  };

  customers.forEach(
    function(r) {
      const id =
        tmv3_clean_(
          r['Customer ID']
        );

      const no =
        tmv3_clean_(
          r['Customer Number']
        );

      const phone =
        tmv3_phone10_(
          r['Primary Phone']
        );

      if (id) {
        index.customerById[id] =
          r;
      }

      if (no) {
        index.customerByNumber[no] =
          r;
      }

      if (phone) {
        if (
          !index
            .customerByPhone[
              phone
            ]
        ) {
          index
            .customerByPhone[
              phone
            ] = [];
        }

        index
          .customerByPhone[
            phone
          ]
          .push(r);
      }
    }
  );

  locations.forEach(
    function(r) {
      const id =
        tmv3_clean_(
          r['Customer ID']
        );

      if (id) {
        if (
          !index
            .locationsByCustomer[
              id
            ]
        ) {
          index
            .locationsByCustomer[
              id
            ] = [];
        }

        index
          .locationsByCustomer[
            id
          ]
          .push(r);
      }
    }
  );

  orders.forEach(
    function(r) {
      const id =
        tmv3_clean_(
          r['Order ID']
        );

      const no =
        tmv3_clean_(
          r['Order Number']
        );

      if (id) {
        index.orderById[id] =
          r;
      }

      if (no) {
        if (
          !index
            .ordersByNumber[
              no
            ]
        ) {
          index
            .ordersByNumber[
              no
            ] = [];
        }

        index
          .ordersByNumber[
            no
          ]
          .push(r);
      }
    }
  );

  tasks.forEach(
    function(r) {
      const id =
        tmv3_clean_(
          r['Task ID']
        );

      const orderId =
        tmv3_clean_(
          r['Order ID']
        );

      const customerId =
        tmv3_clean_(
          r['Customer ID']
        );

      if (id) {
        index.taskById[id] =
          r;
      }

      if (orderId) {
        if (
          !index
            .tasksByOrder[
              orderId
            ]
        ) {
          index
            .tasksByOrder[
              orderId
            ] = [];
        }

        index
          .tasksByOrder[
            orderId
          ]
          .push(r);
      }

      if (customerId) {
        if (
          !index
            .tasksByCustomer[
              customerId
            ]
        ) {
          index
            .tasksByCustomer[
              customerId
            ] = [];
        }

        index
          .tasksByCustomer[
            customerId
          ]
          .push(r);
      }
    }
  );

  return index;
}

function tmv3_resolveEvent_(
  e,
  refs,
  stateIndex
) {
  const cfg =
    TMV3.VERTICALS[
      e.vertical
    ];

  const key =
    e.vertical +
    '|' +
    e.eventId;

  const prior =
    stateIndex[key] ||
    {};

  const override =
    tmv3_clean_(
      prior[
        'Classification Override'
      ]
    );

  if (
    e.vertical ===
      'PreInspection' &&
    tmv3_norm_(override) ===
      'not preinspect'
  ) {
    return tmv3_result_(
      e,
      {
        status:
          'IGNORED',
        nextAction:
          'NONE',
        issue:
          'Manual override: NOT Preinspect'
      }
    );
  }

  let order = null;
  let evidence = [];

  if (cfg.orderRequired) {
    const orderCandidates =
      e.existingOrderId && refs.orderById[e.existingOrderId]
        ? [refs.orderById[e.existingOrderId]]
        : (
            e.orderNumber
              ? (refs.ordersByNumber[e.orderNumber] || [])
              : []
          );

    if (orderCandidates.length > 1) {
      return tmv3_result_(
        e,
        {
          status: 'REVIEW',
          nextAction: 'RESOLVE ' + cfg.orderLabel.toUpperCase(),
          issue: 'Multiple matching ' + cfg.orderLabel + ' records.',
          errorCode: 'AMBIGUOUS_ORDER'
        }
      );
    }

    if (orderCandidates.length === 0) {
      return tmv3_result_(
        e,
        {
          status: 'BLOCKED',
          nextAction: 'ADD / VERIFY ' + cfg.orderLabel.toUpperCase(),
          issue: cfg.orderLabel + ' is required and was not resolved.',
          errorCode: 'ORDER_REQUIRED'
        }
      );
    }

    order = orderCandidates[0];
    evidence.push(cfg.orderLabel + ' exact');
  }

  const identity = tmv3_resolveIdentity_(
    e,
    cfg,
    refs,
    order
  );

  if (identity.status !== 'MATCHED') {
    return tmv3_result_(
      e,
      {
        status: identity.status || 'REVIEW',
        nextAction:
          identity.errorCode === 'LOCATION_UNRESOLVED'
            ? (
                e.vertical === 'PreInspection'
                  ? 'VERIFY / CREATE LOCATION'
                  : 'RESOLVE LOCATION'
              )
            : 'RESOLVE IDENTITY',
        issue: identity.reason || 'Identity requires review.',
        customer: identity.customer || null,
        location: identity.location || null,
        order: order,
        evidence: evidence.concat(identity.evidence || []),
        errorCode: identity.errorCode || 'IDENTITY_REVIEW'
      }
    );
  }

  const customer = identity.customer;
  const location = identity.location;
  const contact = identity.contact || null;
  const customerId = tmv3_clean_(customer['Customer ID']);
  const contactId = contact
    ? tmv3_clean_(contact['Contact ID'])
    : (
        order
          ? tmv3_clean_(order['Contact ID'])
          : ''
      );

  Array.prototype.push.apply(evidence, identity.evidence || []);

  let taskCandidates =
    [];

  if (
    e.forcedTaskId &&
    refs.taskById[String(e.forcedTaskId)]
  ) {
    taskCandidates = [refs.taskById[String(e.forcedTaskId)]];
    if (e.forcedMatchEvidence) evidence.push(e.forcedMatchEvidence);

  } else if (e.vertical === 'PreInspection') {
    let piDecision;

    try {
      piDecision = tmv3_preInspectionTaskDecision_(
        e,
        customer,
        location
      );
    } catch (err) {
      return tmv3_result_(
        e,
        {
          status: 'REVIEW',
          nextAction: 'REVIEW TASK LOOKUP',
          issue: String(err && err.message || err),
          customer: customer,
          location: location,
          evidence: evidence,
          errorCode: 'PREINSPECTION_TASK_LOOKUP_ERROR'
        }
      );
    }

    if (piDecision.status === 'REVIEW') {
      return tmv3_result_(
        e,
        {
          status: 'REVIEW',
          nextAction: 'REVIEW TASK CANDIDATES',
          issue: piDecision.reason,
          customer: customer,
          location: location,
          evidence: evidence.concat(piDecision.evidence || []),
          errorCode: piDecision.errorCode || 'PREINSPECTION_TASK_REVIEW'
        }
      );
    }

    taskCandidates = piDecision.task ? [piDecision.task] : [];

    if (piDecision.task) {
      evidence.push(piDecision.reason || 'PreInspection task verified');
    }

  } else if (
    e.existingTaskId &&
    refs
      .taskById[
        e.existingTaskId
      ]
  ) {
    taskCandidates =
      [
        refs
          .taskById[
            e.existingTaskId
          ]
      ];

  } else if (
    order &&
    tmv3_clean_(
      order[
        'Order ID'
      ]
    )
  ) {
    taskCandidates =
      (
        refs
          .tasksByOrder[
            tmv3_clean_(
              order[
                'Order ID'
              ]
            )
          ] ||
        []
      )
        .filter(
          function(t) {
            return tmv3_taskFitsVertical_(
              t,
              e.vertical,
              cfg
            );
          }
        );

  } else {
    taskCandidates =
      (
        refs
          .tasksByCustomer[
            customerId
          ] ||
        []
      )
        .filter(
          function(t) {
            return (
              tmv3_taskFitsVertical_(
                t,
                e.vertical,
                cfg
              ) &&
              tmv3_clean_(
                t[
                  'Location ID'
                ]
              ) ===
              tmv3_clean_(
                location[
                  'Location ID'
                ]
              )
            );
          }
        );
  }

  const open =
    taskCandidates
      .filter(
        function(t) {
          return tmv3_taskIsOpen_(
            t[
              'Status'
            ]
          );
        }
      );

  const completed =
    taskCandidates
      .filter(
        function(t) {
          return tmv3_taskIsCompleted_(
            t[
              'Status'
            ]
          );
        }
      );

  if (
    open.length >
    1
  ) {
    return tmv3_result_(
      e,
      {
        status:
          'REVIEW',
        nextAction:
          'SELECT TASK',
        issue:
          'Multiple applicable OPEN tasks found.',
        customer:
          customer,
        location:
          location,
        order:
          order,
        contact:
          contact,
        evidence:
          evidence,
        errorCode:
          'MULTIPLE_OPEN_TASKS'
      }
    );
  }

  if (
    open.length ===
    0
  ) {
    const safeCreate =
      !!customerId &&
      !!tmv3_clean_(
        location[
          'Location ID'
        ]
      ) &&
      (
        !cfg.orderRequired ||
        !!order
      );

    if (!safeCreate) {
      return tmv3_result_(
        e,
        {
          status:'BLOCKED',
          nextAction:'RESOLVE REQUIRED DATA',
          issue:'Required Customer / Location / Order prerequisites are not fully verified.',
          customer:customer,
          location:location,
          order:order,
          contact:contact,
          evidence:evidence,
          errorCode:'TASK_PREREQUISITE_UNRESOLVED'
        }
      );
    }

    const lifecycle = tmv3_step6NoOpenTaskDecision_({
      vertical:e.vertical,
      start:e.start,
      step5:{
        disposition:'NO_TASK',
        historyTasks:completed
      }
    });

    if (lifecycle.disposition === 'REVIEW') {
      return tmv3_result_(
        e,
        {
          status:'REVIEW',
          nextAction:'REVIEW TASK HISTORY',
          issue:lifecycle.reason,
          customer:customer,
          location:location,
          order:order,
          contact:contact,
          evidence:evidence.concat(['SHARED_STAGE6_HISTORY_GATE']),
          errorCode:lifecycle.code
        }
      );
    }

    if (lifecycle.disposition === 'FULFILLED_NO_RECREATE') {
      const fulfilledTask =
        (lifecycle.tasks || [])[0] ||
        completed[0] ||
        null;

      return tmv3_result_(
        e,
        {
          status:'MATCHED',
          nextAction:'NONE',
          issue:lifecycle.reason,
          customer:customer,
          location:location,
          order:order,
          contact:contact,
          task:fulfilledTask,
          evidence:evidence.concat(['SHARED_STAGE6_HISTORY_GATE']),
          verification:'LIFECYCLE FULFILLED · NO RECREATE',
          errorCode:lifecycle.code,
          lastVerified:tmv3_now_()
        }
      );
    }

    const recreate =
      lifecycle.disposition === 'RECREATE_TASK';

    return tmv3_result_(
      e,
      {
        status:recreate ? 'READY RECREATE' : 'READY CREATE',
        nextAction:recreate ? 'RECREATE TASK' : 'CREATE TASK',
        issue:lifecycle.reason,
        customer:customer,
        location:location,
        order:order,
        contact:contact,
        evidence:evidence.concat(['SHARED_STAGE6_HISTORY_GATE']),
        errorCode:lifecycle.code
      }
    );
  }

  const task =
    open[0];

  evidence.push(
    'Single OPEN task'
  );

  const differences =
    tmv3_taskDifferences_(
      e,
      cfg,
      task,
      order,
      customer,
      location
    );

  const links =
    tmv3_calendarLinkAcceptance_(
      e,
      cfg,
      task,
      order,
      customer
    );

  const allGood =
    differences.length ===
      0 &&
    links.ok;

  const verification =
    tmv3_verificationSummary_(
      e,
      cfg,
      task,
      order,
      customer,
      location,
      differences,
      links
    );

  return tmv3_result_(
    e,
    {
      status:
        allGood
          ? 'MATCHED'
          : 'READY',
      nextAction:
        allGood
          ? 'NONE'
          : 'SYNC + VERIFY',
      issue:
        allGood
          ? ''
          : differences
              .concat(
                links.issues
              )
              .join('; '),
      customer:
        customer,
      location:
        location,
      order:
        order,
      contact:
        contact,
      task:
        task,
      evidence:
        evidence,
      verification:
        verification,
      lastVerified:
        allGood
          ? tmv3_now_()
          : ''
    }
  );
}

function tmv3_taskFitsVertical_(
  task,
  vertical,
  cfg
) {
  const typeId =
    Number(
      tmv3_clean_(
        task[
          'Task Type ID'
        ]
      ) ||
      0
    );

  if (
    cfg.taskTypeId &&
    typeId ===
      Number(
        cfg.taskTypeId
      )
  ) {
    return true;
  }

  const hay =
    tmv3_norm_(
      task[
        'Task Type'
      ] +
      ' ' +
      task[
        'Name'
      ]
    );

  return (
    cfg.taskTypeNames ||
    []
  )
    .some(
      function(name) {
        return (
          hay.indexOf(
            tmv3_norm_(
              name
            )
          ) !== -1
        );
      }
    );
}

function tmv3_taskIsOpen_(
  status
) {
  return [
    'open',
    'in progress',
    'inprogress'
  ].indexOf(
    tmv3_norm_(
      status
    )
  ) !== -1;
}

function tmv3_taskIsCompleted_(
  status
) {
  return [
    'done',
    'complete',
    'completed',
    'closed',
    'cancelled',
    'canceled'
  ].indexOf(
    tmv3_norm_(
      status
    )
  ) !== -1;
}

/**
 * Google Apps Script — replace existing tmv3_taskDifferences_ in 10_Shared_Core.gs.
 * PreInspection 2026-11-10-r1 targeted source fix.
 * Preserves all other customer, location, SO and technician checks.
 * Does not call Striven/Calendar, write to Sheets or mutate any source record.
 */
function tmv3_taskDifferences_(e, cfg, task, order, customer, location) {
  const issues = [];
  if (tmv3_clean_(task['Customer ID']) &&
      tmv3_clean_(task['Customer ID']) !== tmv3_clean_(customer['Customer ID'])) {
    issues.push('Task customer differs');
  }
  if (tmv3_clean_(task['Location ID']) &&
      tmv3_clean_(task['Location ID']) !== tmv3_clean_(location['Location ID'])) {
    issues.push('Task location differs');
  }
  if (cfg.orderRequired && tmv3_clean_(task['Order ID']) &&
      tmv3_clean_(task['Order ID']) !== tmv3_clean_(order['Order ID'])) {
    issues.push('Task ' + cfg.orderLabel + ' differs');
  }
  if (e.vertical === 'PreInspection' && tmv3_clean_(task['Order ID'])) {
    issues.push('PreInspection task must not be attached to Sales Order');
  }

  const taskStart = tmv3_clean_(task['Start']);
  const taskDue = tmv3_clean_(task['Due']);
  // Compare actual instants to the MINUTE, rather than searching an ISO date
  // string inside normalized text. The former code created false discrepancies.
  if (taskStart && !tmv3_sameMinute_(e.start, taskStart)) {
    issues.push('Start date/time differs');
  }
  if (taskDue && !tmv3_sameMinute_(e.end, taskDue)) {
    issues.push('Due date/time differs');
  }

  const assigned = tmv3_norm_(task['Assignees']);
  if (e.vertical === 'Service' && e.technician &&
      assigned.indexOf(tmv3_norm_(e.technician)) === -1) {
    issues.push('Technician assignment differs');
  }
  if (e.vertical === 'PreInspection') {
    // Canonicalize only the known 'Preinspection' / 'Pre-Inspection'
    // spelling difference in the read-only text cache. Task write authorization
    // must continue to verify the actual Striven pool ID 8.
    function canonicalPool(value) {
      return tmv3_norm_(value).replace(/\bpre\s+inspection\b/g, 'preinspection');
    }
    const pools=canonicalPool(task['Pools']);
    const expected=canonicalPool(cfg.defaultPoolName);
    if (!pools || (expected && pools.indexOf(expected) === -1 &&
        !pools.split(' ').includes(String(cfg.defaultPoolId)))) {
      issues.push('Pre-Inspection Pool 8 missing');
    }
  }
  return issues;
}

function tmv3_calendarLinkAcceptance_(
  e,
  cfg,
  task,
  order,
  customer
) {
  const issues =
    [];

  const taskId =
    tmv3_clean_(
      task[
        'Task ID'
      ]
    );

  if (
    cfg.calendarTaskLinkRequired &&
    tmv3_clean_(
      e.existingTaskId
    ) !==
      taskId
  ) {
    issues.push(
      'Calendar Task link missing/wrong'
    );
  }

  if (
    cfg.calendarOrderLinkRequired &&
    order &&
    tmv3_clean_(
      e.existingOrderId
    ) !==
      tmv3_clean_(
        order[
          'Order ID'
        ]
      )
  ) {
    issues.push(
      'Calendar ' +
      cfg.orderLabel +
      ' link missing/wrong'
    );
  }

  if (
    cfg.calendarOrdersPageLinkRequired
  ) {
    const expected =
      TMV3
        .CUSTOMER_ORDERS_PAGE_BASE +
      encodeURIComponent(
        tmv3_clean_(
          customer[
            'Customer ID'
          ]
        )
      );

    if (
      e.description.indexOf(
        expected
      ) === -1
    ) {
      issues.push(
        'Calendar Sales Orders page link missing'
      );
    }
  }

  return {
    ok:
      issues.length ===
      0,
    issues:
      issues
  };
}

function tmv3_result_(
  e,
  x
) {
  x = x || {};

  const customer = x.customer || null;
  const location = x.location || null;
  const order = x.order || null;
  const task = x.task || null;
  const contact = x.contact || null;
  const cfg = TMV3.VERTICALS[e.vertical];

  const result = {
    vertical: e.vertical,
    calendarId: e.calendarId,
    eventId: e.eventId,
    fingerprint: e.fingerprint,
    calendarUpdatedAt: e.calendarUpdatedAt,
    status: x.status || 'REVIEW',
    nextAction: x.nextAction || 'REVIEW',
    date: tmv3_date_(e.start),
    time: tmv3_time_(e.start),
    title: e.title,
    customer: customer ? tmv3_clean_(customer['Name']) : '',
    customerId: customer ? tmv3_clean_(customer['Customer ID']) : '',
    location: location
      ? [location['Address 1'], location['City']].map(tmv3_clean_).filter(Boolean).join(', ')
      : e.location,
    locationId: location ? tmv3_clean_(location['Location ID']) : '',
    contactId: contact
      ? tmv3_clean_(contact['Contact ID'])
      : (order ? tmv3_clean_(order['Contact ID']) : ''),
    order: order
      ? (
          tmv3_clean_(order['Order Number']) +
          (tmv3_clean_(order['Name']) ? ' - ' + tmv3_clean_(order['Name']) : '')
        )
      : '',
    orderId: order ? tmv3_clean_(order['Order ID']) : '',
    task: task
      ? (
          tmv3_clean_(task['Task Number']) +
          (tmv3_clean_(task['Name']) ? ' - ' + tmv3_clean_(task['Name']) : '')
        )
      : '',
    taskId: task ? tmv3_clean_(task['Task ID']) : '',
    taskStatus: task ? tmv3_clean_(task['Status']) : '',
    assignedTo: task ? tmv3_clean_(task['Assignees']) : (e.technician || ''),
    matchEvidence: (x.evidence || []).join(' + '),
    verification: x.verification || (
      x.status === 'MATCHED'
        ? 'CALENDAR ↔ STRIVEN PASS'
        : (
            x.status === 'READY CREATE' || x.status === 'READY RECREATE'
              ? 'IDENTITY VERIFIED · TASK ACTION PENDING'
              : 'NOT V3 VERIFIED'
          )
    ),
    issue: x.issue || '',
    calendarLinks: tmv3_calendarLinkSummary_(e),
    lastVerified: x.lastVerified || '',
    errorCode: x.errorCode || ''
  };

  result.dataChecklist = tmv3_buildDataChecklist_(
    e,
    cfg,
    customer,
    location,
    order,
    contact,
    task,
    result
  );

  return result;
}

function tmv3_buildDataChecklist_(eventRecord, cfg, customer, location, order, contact, task, result) {
  const lines = [];
  const issue = tmv3_clean_(result && result.issue);
  const verification = tmv3_clean_(result && result.verification);
  const hasTask = !!(task && tmv3_clean_(task['Task ID']));
  const hasContact = !!(
    (contact && tmv3_clean_(contact['Contact ID'])) ||
    (order && tmv3_clean_(order['Contact ID']))
  );

  lines.push(eventRecord && eventRecord.eventId ? '✅ Calendar' : '❌ Calendar');

  if (!hasTask) {
    lines.push('⚪ Time');
  } else if (/start date\/time differs|due date\/time differs/i.test(issue)) {
    lines.push('❌ Time');
  } else {
    lines.push('✅ Time');
  }

  lines.push(customer && tmv3_clean_(customer['Customer ID']) ? '✅ Customer' : '❌ Customer');

  if (cfg && cfg.orderRequired) {
    if (eventRecord.vertical === 'Service') {
      lines.push(eventRecord.orderNumber ? '✅ SO' : '❌ SO');
      lines.push(order && tmv3_clean_(order['Order ID']) ? '✅ WO' : '⚪ WO');
    } else {
      lines.push(order && tmv3_clean_(order['Order ID']) ? '✅ SO' : '❌ SO');
    }
  }

  lines.push(hasContact ? '✅ Contact' : '⚪ Contact');
  lines.push(location && tmv3_clean_(location['Location ID']) ? '✅ Location' : '❌ Location');

  if (hasTask) {
    lines.push('✅ Task');
    lines.push(tmv3_taskIsOpen_(task['Status']) ? '✅ Open Task' : '❌ Open Task');
  } else {
    const plannedCreate = /READY CREATE/i.test(tmv3_clean_(result && result.status));
    lines.push(plannedCreate ? '⚪ Task' : '❌ Task');
    lines.push('⚪ Open Task');
  }

  lines.push(tmv3_assignmentChecklistLine_(eventRecord, task, verification, issue));

  if (!hasTask) {
    lines.push('⚪ Calendar Link');
  } else if (/LINKS\s*✓/i.test(verification)) {
    lines.push('✅ Calendar Link');
  } else if (/LINKS\s*[△✕]/i.test(verification)) {
    lines.push('❌ Calendar Link');
  } else {
    lines.push(eventRecord && eventRecord.existingTaskUrl ? '✅ Calendar Link' : '❌ Calendar Link');
  }

  return lines.join('\n');
}

function tmv3_assignmentChecklistLine_(eventRecord, task, verification, issue) {
  if (!task) return '⚪ Assignee';

  if (/technician assignment differs|pool 8 missing|assignment differs/i.test(issue)) {
    return '❌ Assignee';
  }

  if (/ASSIGN\s*✓/i.test(verification)) {
    return '✅ Assignee';
  }

  if (/ASSIGN\s*[△✕]/i.test(verification)) {
    return '❌ Assignee';
  }

  if (
    eventRecord &&
    eventRecord.vertical === 'PreInspection' &&
    /POOL 8\s*✓/i.test(verification) &&
    !/INSPECTOR CHECK PENDING/i.test(verification)
  ) {
    return '✅ Assignee';
  }

  return '⚪ Assignee';
}

function tmv3_calendarLinkSummary_(
  e
) {
  const a =
    [];

  if (
    e.existingOrderUrl
  ) {
    a.push(
      'Order'
    );
  }

  if (
    e.existingTaskUrl
  ) {
    a.push(
      'Task'
    );
  }

  if (
    e.vertical ===
      'PreInspection' &&
    e.description.indexOf(
      TMV3
        .CUSTOMER_ORDERS_PAGE_BASE
    ) !== -1
  ) {
    a.push(
      'Sales Orders Page'
    );
  }

  return a.length
    ? a.join(' + ')
    : 'Missing';
}

function tmv3_writeOperatorViews_(
  resolved
) {
  Object.keys(
    TMV3.VERTICALS
  ).forEach(
    function(vertical) {
      const rows =
        resolved
          .filter(
            function(r) {
              return (
                r.vertical ===
                vertical
              );
            }
          )
          .sort(
            function(a,b) {
              return (
                a.date +
                a.time
              )
                .localeCompare(
                  b.date +
                  b.time
                );
            }
          )
          .map(
            function(r) {
              return [
                r.status,
                r.dataChecklist,
                r.date,
                r.time,
                r.customer,
                r.task,
                r.issue,
                r.nextAction,
                r.title,
                r.order,
                r.location,
                r.taskStatus,
                r.assignedTo,
                r.verification,
                r.calendarLinks,
                r.lastVerified,
                r.eventId
              ];
            }
          );

      const sheetName =
        TMV3
          .VERTICALS[
            vertical
          ]
          .sheet;

      tmv3_replaceRows_(
        sheetName,
        TMV3_OPERATOR_HEADERS
          .slice(),
        rows
      );

      tmv3_formatOperatorView_(
        sheetName
      );
    }
  );
}

function tmv3_formatOperatorView_(sheetName) {
  const sh = tmv3_sheet_(sheetName);
  const visible = Number(TMV3_OPERATOR_VISIBLE_COLUMN_COUNT || 8);
  const lastRow = sh.getLastRow();
  const maxColumns = sh.getMaxColumns();

  sh.showColumns(1, Math.min(visible, maxColumns));

  if (maxColumns > visible) {
    sh.hideColumns(visible + 1, maxColumns - visible);
  }

  const widths = [110, 185, 95, 85, 180, 220, 320, 145];
  widths.forEach(function(width, index) {
    if (index + 1 <= maxColumns) {
      sh.setColumnWidth(index + 1, width);
    }
  });

  if (lastRow >= 1) {
    sh.getRange(1, 1, lastRow, Math.min(visible, maxColumns))
      .setWrap(true)
      .setVerticalAlignment('middle');
  }

  sh.getRange(1, 1, 1, Math.min(visible, maxColumns))
    .setFontWeight('bold');

  if (lastRow > 1) {
    sh.autoResizeRows(2, lastRow - 1);
  }
}

function tmv3_statusCounts_(
  records
) {
  const out =
    {};

  (records || [])
    .forEach(
      function(r) {
        out[
          r.status
        ] =
          (
            out[
              r.status
            ] ||
            0
          ) +
          1;
      }
    );

  return out;
}

function tmv3_updateOverview_(
  records
) {
  const sh =
    tmv3_sheet_(
      TMV3
        .SHEETS
        .OVERVIEW
    );

  const verticals =
    [
      'Install',
      'Delivery',
      'Service',
      'PreInspection'
    ];

  verticals.forEach(
    function(v,i) {
      const rows =
        records.filter(
          function(r) {
            return (
              r.vertical ===
              v
            );
          }
        );

      const review =
        rows.filter(
          function(r) {
            return (
              r.status ===
              'REVIEW'
            );
          }
        ).length;

      const blocked =
        rows.filter(
          function(r) {
            return (
              r.status ===
              'BLOCKED'
            );
          }
        ).length;

      const matched =
        rows.filter(
          function(r) {
            return (
              r.status ===
              'MATCHED'
            );
          }
        ).length;

      sh
        .getRange(
          7 + i,
          1,
          1,
          6
        )
        .setValues(
          [[
            v,
            'SHADOW',
            rows.length,
            review,
            blocked,
            matched
              ? tmv3_now_()
              : ''
          ]]
        );
    }
  );
}






function tmv3_sameMinute_(a, b) {
  const da = tmv3_parseDateTime_(a);
  const db = tmv3_parseDateTime_(b);
  if (!da || !db) return false;
  return Math.floor(da.getTime() / 60000) === Math.floor(db.getTime() / 60000);
}

function tmv3_sameLocalDay_(a, b) {
  const da = tmv3_parseDateTime_(a);
  const db = tmv3_parseDateTime_(b);
  if (!da || !db) return false;
  return tmv3_date_(da) === tmv3_date_(db);
}

function tmv3_allPhones_(text) {
  const raw = tmv3_clean_(text);
  const out = [];
  const rx = /(?:\+?1[\s.\-]?)?\(?([2-9]\d{2})\)?[\s.\-]?([2-9]\d{2})[\s.\-]?(\d{4})/g;
  let m;
  while ((m = rx.exec(raw)) !== null) {
    out.push(m[1] + m[2] + m[3]);
  }
  return tmv3_unique_(out);
}


function tmv3_verificationSummary_(
  eventRecord,
  cfg,
  task,
  order,
  customer,
  location,
  differences,
  links
) {
  const diff = (differences || []).join(' | ');
  const parts = [];

  parts.push('CUSTOMER ✓');
  parts.push('LOCATION ✓');

  if (cfg.orderRequired) {
    parts.push(order ? (cfg.orderLabel.toUpperCase() + ' ✓') : (cfg.orderLabel.toUpperCase() + ' ✕'));
  }

  parts.push(tmv3_taskIsOpen_(task && task['Status']) ? 'TASK OPEN ✓' : 'TASK OPEN ✕');

  const scheduleOk =
    !/start date\/time differs|due date\/time differs/i.test(diff);
  parts.push(scheduleOk ? 'TIME ✓' : 'TIME △');

  if (eventRecord.vertical === 'Service') {
    parts.push(/technician assignment differs/i.test(diff) ? 'ASSIGN △' : 'ASSIGN ✓');
  } else if (eventRecord.vertical === 'PreInspection') {
    parts.push(/pool 8 missing/i.test(diff) ? 'POOL 8 △' : 'POOL 8 ✓');
    parts.push('INSPECTOR CHECK PENDING');
  } else {
    parts.push('ASSIGN CHECK PENDING');
  }

  parts.push(links && links.ok ? 'LINKS ✓' : 'LINKS △');

  return (
    ((differences || []).length === 0 && links && links.ok)
      ? 'CALENDAR ↔ STRIVEN PASS · ' + parts.join(' · ')
      : 'CALENDAR ↔ STRIVEN CHECK · ' + parts.join(' · ')
  );
}


/************************************************************
 * SERVICE MULTI-FIREPLACE EXPANSION
 *
 * Old-system rule preserved:
 * If the same Work Order has multiple OPEN Service Tasks with
 * explicit FP#n / Fireplace #n markers, one Calendar event
 * expands to one V3 operational record per fireplace task.
 *
 * Arbitrary multiple tasks without those markers still REVIEW.
 ************************************************************/
function tmv3_resolveEventRecords_(eventRecord, refs, stateIndex) {
  if (eventRecord.vertical !== 'Service') {
    return [tmv3_resolveEvent_(eventRecord, refs, stateIndex)];
  }

  const orderCandidates =
    eventRecord.existingOrderId && refs.orderById[eventRecord.existingOrderId]
      ? [refs.orderById[eventRecord.existingOrderId]]
      : (
          eventRecord.orderNumber
            ? (refs.ordersByNumber[eventRecord.orderNumber] || [])
            : []
        );

  if (orderCandidates.length !== 1) {
    return [tmv3_resolveEvent_(eventRecord, refs, stateIndex)];
  }

  const order = orderCandidates[0];
  const orderId = tmv3_clean_(order['Order ID']);
  const cfg = TMV3.VERTICALS.Service;

  const candidates = (refs.tasksByOrder[orderId] || [])
    .filter(function(task) {
      return (
        tmv3_taskFitsVertical_(task, 'Service', cfg) &&
        tmv3_taskIsOpen_(task['Status'])
      );
    })
    .map(function(task) {
      return {
        task: task,
        fireplaceNumber: tmv3_extractServiceFireplaceNumber_(task['Name'])
      };
    })
    .filter(function(item) {
      return item.fireplaceNumber > 0;
    })
    .sort(function(a, b) {
      return a.fireplaceNumber - b.fireplaceNumber;
    });

  if (candidates.length <= 1) {
    return [tmv3_resolveEvent_(eventRecord, refs, stateIndex)];
  }

  const numberCounts = {};
  candidates.forEach(function(item) {
    numberCounts[item.fireplaceNumber] =
      (numberCounts[item.fireplaceNumber] || 0) + 1;
  });

  const duplicates = Object.keys(numberCounts).filter(function(n) {
    return numberCounts[n] > 1;
  });

  if (duplicates.length) {
    return [tmv3_result_(eventRecord, {
      status: 'REVIEW',
      nextAction: 'REVIEW MULTI-FIREPLACE TASKS',
      issue:
        'Multiple OPEN Service Tasks use the same fireplace marker: FP#' +
        duplicates.join(', FP#') +
        '.',
      errorCode: 'DUPLICATE_SERVICE_FIREPLACE_NUMBER'
    })];
  }

  return candidates.map(function(item) {
    const clone = Object.assign({}, eventRecord, {
      forcedTaskId: tmv3_clean_(item.task['Task ID']),
      forcedMatchEvidence: 'WORK ORDER + FP#' + item.fireplaceNumber,
      serviceFireplaceNumber: item.fireplaceNumber
    });

    const result = tmv3_resolveEvent_(clone, refs, stateIndex);
    result.serviceFireplaceNumber = item.fireplaceNumber;
    return result;
  });
}




// Original: 42_Step5_Task_Resolution
/************************************************************
 * TM V3 — STAGE 5 TASK RESOLUTION
 *
 * Narrow contract:
 * - Consumes exact Step-4 records.
 * - Finds the current applicable Striven Task(s).
 * - No Task create / patch / assignment writes.
 * - No Calendar writes.
 * - No reconciliation plan yet.
 ************************************************************/

function tmv3_step5CachedSourceSummary_() {
  const tasks = tmv3_rows_(TMV3.SHEETS.TASKS);
  const preInspectionTasks = tasks.filter(function(task) {
    return Number(task['Task Type ID'] || 0) === 105;
  }).length;

  return {
    tasks: tasks.length,
    preInspectionTasks: preInspectionTasks,
    source: 'CACHED_TASK_SOURCES'
  };
}

function tmv3_step5RefreshTaskSources_(options) {
  options = options || {};
  tmv3_assertShadow_();
  tmv3_resetRuntimeMetrics_();

  // One canonical transaction refresh owns all Task cache population.
  // PreInspection Task hydration is centralized here and reused by routine
  // Stage-5 mapping for the configured TTL.
  const refreshed = tmv3_refreshTransactionSources_({
    forcePreInspectionTasks:
      options.forcePreInspectionTasks === true
  });
  const result = Object.assign({}, refreshed, {
    api: tmv3_runtimeMetrics_()
  });

  tmv3_audit_(
    'SYSTEM','','','STEP5_REFRESH_TASK_SOURCES','PASS',
    JSON.stringify(result)
  );

  return result;
}

function tmv3_step5TaskIndex_() {
  const tasks = tmv3_rows_(TMV3.SHEETS.TASKS);

  const refs = {
    tasks: tasks,
    taskById: {},
    tasksByOrder: {},
    tasksByCustomer: {}
  };

  tasks.forEach(function(task) {
    const taskId = tmv3_clean_(task['Task ID']);
    const orderId = tmv3_clean_(task['Order ID']);
    const customerId = tmv3_clean_(task['Customer ID']);

    if (taskId) refs.taskById[taskId] = task;

    if (orderId) {
      if (!refs.tasksByOrder[orderId]) refs.tasksByOrder[orderId] = [];
      refs.tasksByOrder[orderId].push(task);
    }

    if (customerId) {
      if (!refs.tasksByCustomer[customerId]) refs.tasksByCustomer[customerId] = [];
      refs.tasksByCustomer[customerId].push(task);
    }
  });

  return refs;
}

function tmv3_step5TaskRecords_(step4Records, refs) {
  const records = (step4Records || []).map(function(record) {
    const step4 = record.step4 || {};

    if (step4.disposition !== 'VERIFIED') {
      return Object.assign({}, record, {
        step5: {
          disposition: 'NOT_RUN',
          code: 'STEP4_' + (step4.disposition || 'UNKNOWN'),
          reason: 'Task resolution did not run because Step 4 is not VERIFIED.',
          tasks: [],
          completedHistory: 0,
          historyTasks: [],
          warnings: []
        }
      });
    }

    const decision =
      record.vertical === 'PreInspection'
        ? tmv3_step5ResolvePreInspection_(record, refs)
        : tmv3_step5ResolveStandardTask_(record, refs);

    return Object.assign({}, record, { step5: decision });
  });

  tmv3_step5NormalizePreInspectionTaskTimesAcrossRecords_(records);
  tmv3_step5GuardTaskLinkedAcrossEventDates_(records);

  return records;
}



function tmv3_step5ResolveStandardTask_(record, refs) {
  const cfg = TMV3.VERTICALS[record.vertical];
  const identity = record.step4 || {};
  const anchor = record.step3 && record.step3.anchor
    ? record.step3.anchor
    : {};

  const customer = identity.customer || null;
  const location = identity.location || null;

  const customerId = tmv3_clean_(customer && customer['Customer ID']);
  const locationId = tmv3_clean_(location && location['Location ID']);
  const orderId = tmv3_clean_(anchor.orderId);

  let candidates = [];
  const evidence = [];
  const warnings = [];

  let wrongVerticalLinkedTask = null;
  let linkedTaskMissingFromCache = false;

  if (record.existingTaskId) {
    const linkedId = tmv3_clean_(record.existingTaskId);
    let linked = refs.taskById[linkedId] || null;

    if (!linked) {
      linkedTaskMissingFromCache = true;
      warnings.push(
        'Calendar carries Task ' + linkedId +
        ', but it is not present in the current Task cache. Stage 5 will not spend a live Task GET; cached Order or Customer + Location evidence will be used instead.'
      );
      evidence.push(
        'CALENDAR_TASK_LINK_NOT_IN_CACHE_DEFER_LIVE_READ'
      );
    } else {
      evidence.push('CALENDAR_TASK_LINK_CACHE_EXACT');
    }

    if (
      linked &&
      !tmv3_taskFitsVertical_(linked, record.vertical, cfg)
    ) {
      wrongVerticalLinkedTask = linked;
      warnings.push(
        'Calendar Task link points to Task ' + linkedId +
        ', which does not fit the ' + record.vertical +
        ' vertical. The stale link was ignored for resolution.'
      );
      evidence.push('CALENDAR_TASK_LINK_WRONG_VERTICAL_IGNORED');
      linked = null;
    }

    if (linked) candidates = [linked];
  }

  if (!candidates.length && orderId) {
    candidates = (refs.tasksByOrder[orderId] || [])
      .filter(function(task) {
        return tmv3_taskFitsVertical_(task, record.vertical, cfg);
      });

    evidence.push('TASKS_FROM_VERIFIED_BUSINESS_ANCHOR');

  } else if (
    !candidates.length &&
    customerId &&
    locationId
  ) {
    candidates = (refs.tasksByCustomer[customerId] || [])
      .filter(function(task) {
        return (
          tmv3_taskFitsVertical_(task, record.vertical, cfg) &&
          tmv3_clean_(task['Location ID']) === locationId
        );
      });

    evidence.push('TASKS_FROM_VERIFIED_CUSTOMER_LOCATION');
  }

  if (
    linkedTaskMissingFromCache &&
    !candidates.length
  ) {
    return tmv3_step5Decision_(
      'REVIEW',
      'CALENDAR_LINKED_TASK_NOT_IN_CACHE_NO_SAFE_FALLBACK',
      'Calendar carries a Task link that is absent from the current Task cache, and no verified cached Task fallback was found. Do not create a replacement Task until the linked Task is freshly verified at the write gate.',
      [],
      0,
      evidence,
      warnings
    );
  }

  if (
    wrongVerticalLinkedTask &&
    !candidates.length &&
    !orderId &&
    !(customerId && locationId)
  ) {
    return tmv3_step5Decision_(
      'REVIEW',
      'CALENDAR_LINKED_TASK_WRONG_TYPE_NO_SAFE_FALLBACK',
      'Calendar Task link is for another vertical and no verified Order or Customer + Location fallback is available.',
      [wrongVerticalLinkedTask],
      0,
      evidence,
      warnings
    );
  }

  const open = candidates.filter(function(task) {
    return tmv3_taskIsOpen_(task['Status']);
  });

  const completed = candidates.filter(function(task) {
    return tmv3_taskIsCompleted_(task['Status']);
  });

  const nonOpenActive = candidates.filter(function(task) {
    return (
      !tmv3_taskIsOpen_(task['Status']) &&
      !tmv3_taskIsCompleted_(task['Status'])
    );
  });

  if (open.length === 0 && nonOpenActive.length) {
    return tmv3_step5Decision_(
      'REVIEW',
      'NON_OPEN_ACTIVE_TASK_EXISTS',
      'Applicable Task exists in a non-open active status and must not be duplicated.',
      nonOpenActive,
      completed.length,
      evidence,
      [],
      completed
    );
  }

  if (open.length === 0) {
    return tmv3_step5Decision_(
      'NO_TASK',
      completed.length
        ? 'NO_OPEN_TASK_COMPLETED_HISTORY'
        : 'NO_TASK_FOUND',
      completed.length
        ? 'No applicable OPEN Task remains; completed history exists.'
        : 'No applicable Task was found.',
      [],
      completed.length,
      evidence,
      [],
      completed
    );
  }

  if (open.length === 1) {
    return tmv3_step5Decision_(
      'MATCHED',
      'SINGLE_OPEN_TASK_MATCHED',
      'Exactly one applicable OPEN Task was found.',
      [open[0]],
      completed.length,
      evidence,
      [],
      completed
    );
  }

  if (record.vertical === 'Service') {
    const expanded = open.map(function(task) {
      return {
        task: task,
        fireplaceNumber: tmv3_extractServiceFireplaceNumber_(task['Name'])
      };
    });

    const everyNumbered = expanded.every(function(item) {
      return item.fireplaceNumber > 0;
    });

    const numbers = expanded.map(function(item) {
      return item.fireplaceNumber;
    });

    const uniqueNumbers = tmv3_unique_(numbers.map(String));

    if (
      everyNumbered &&
      uniqueNumbers.length === expanded.length
    ) {
      expanded.sort(function(a, b) {
        return a.fireplaceNumber - b.fireplaceNumber;
      });

      return tmv3_step5Decision_(
        'MATCHED_MULTI',
        'SERVICE_MULTI_FIREPLACE_TASKS_MATCHED',
        'Multiple OPEN Service Tasks are valid because each has a unique fireplace marker.',
        expanded.map(function(item) { return item.task; }),
        completed.length,
        evidence.concat(
          expanded.map(function(item) {
            return 'FP#' + item.fireplaceNumber;
          })
        ),
        [],
        completed
      );
    }
  }

  return tmv3_step5Decision_(
    'REVIEW',
    'MULTIPLE_OPEN_TASKS',
    'Multiple applicable OPEN Tasks were found.',
    open,
    completed.length,
    evidence,
    warnings,
    completed
  );
}

function tmv3_step5Decision_(
  disposition,
  code,
  reason,
  tasks,
  completedHistory,
  evidence,
  warnings,
  historyTasks
) {
  return {
    disposition: disposition,
    code: code,
    reason: reason,
    tasks: tasks || [],
    completedHistory: Number(completedHistory || 0),
    historyTasks: historyTasks || [],
    evidence: evidence || [],
    warnings: warnings || []
  };
}

function tmv3_step5TaskResolutionRun(reason, refreshSources) {
  tmv3_assertShadow_();

  const sourceSummary =
    refreshSources === true
      ? tmv3_step5RefreshTaskSources_({
          forcePreInspectionTasks:true
        })
      : tmv3_step5CachedSourceSummary_();

  const step2Snapshot = tmv3_step2CalendarRecords_();
  const step3Refs = tmv3_step3AnchorIndex_();
  const step3Records = tmv3_step3BusinessAnchorRecords_(
    step2Snapshot,
    step3Refs
  );

  const step4Refs = tmv3_step4IdentityIndex_();
  const step4Records = tmv3_step4IdentityRecords_(
    step3Records,
    step4Refs
  );

  const refs = tmv3_step5TaskIndex_();
  const records = tmv3_step5TaskRecords_(step4Records, refs);

  const counts = tmv3_step5Counts_(records);
  const write = tmv3_step5WriteOperatorViews_(records);
  const verify = tmv3_step5Verify_(records, step4Records);

  const result = {
    version: TMV3.VERSION,
    executionStage: tmv3_executionStage_(),
    stage: 5,
    status: verify.pass ? 'PASS' : 'REVIEW',
    mode: 'TASK_RESOLUTION_ONLY',
    reason: tmv3_clean_(reason || 'MANUAL'),
    sourceSummary: sourceSummary,
    counts: counts,
    write: write,
    verification: verify,
    taskWritesPerformed: false,
    calendarWritesPerformed: false,
    strivenWritesPerformed: false,
    reconciliationPerformed: false
  };

  tmv3_audit_(
    'SYSTEM','','','STEP5_TASK_RESOLUTION',
    result.status,
    JSON.stringify(result)
  );

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function tmv3_step5TaskResolutionRunCached(reason) {
  return tmv3_step5TaskResolutionRun(
    reason || 'CACHED_TASK_SOURCES',
    false
  );
}

function tmv3_step5Counts_(records) {
  const counts = {
    total:0,
    step4Verified:0,
    matched:0,
    matchedMulti:0,
    noTask:0,
    review:0,
    notRun:0,
    byVertical:{}
  };

  (records || []).forEach(function(record) {
    counts.total++;

    if (!counts.byVertical[record.vertical]) {
      counts.byVertical[record.vertical] = {
        total:0,
        step4Verified:0,
        matched:0,
        matchedMulti:0,
        noTask:0,
        review:0,
        notRun:0
      };
    }

    const bucket = counts.byVertical[record.vertical];
    bucket.total++;

    if (record.step4 && record.step4.disposition === 'VERIFIED') {
      counts.step4Verified++;
      bucket.step4Verified++;
    }

    const disposition = record.step5 ? record.step5.disposition : 'NOT_RUN';

    if (disposition === 'MATCHED') {
      counts.matched++;
      bucket.matched++;
    } else if (disposition === 'MATCHED_MULTI') {
      counts.matchedMulti++;
      bucket.matchedMulti++;
    } else if (disposition === 'NO_TASK') {
      counts.noTask++;
      bucket.noTask++;
    } else if (disposition === 'REVIEW') {
      counts.review++;
      bucket.review++;
    } else {
      counts.notRun++;
      bucket.notRun++;
    }
  });

  return counts;
}

function tmv3_step5WriteOperatorViews_(records) {
  const written = {};

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    const verticalRecords = (records || [])
      .filter(function(record) { return record.vertical === vertical; })
      .sort(function(a, b) {
        const aTime = a.start instanceof Date ? a.start.getTime() : 0;
        const bTime = b.start instanceof Date ? b.start.getTime() : 0;
        if (aTime !== bTime) return aTime - bTime;

        if (vertical === 'Service') {
          const techCompare = tmv3_clean_(a.technician)
            .localeCompare(tmv3_clean_(b.technician));
          if (techCompare !== 0) return techCompare;
        }

        return tmv3_clean_(a.title).localeCompare(tmv3_clean_(b.title));
      });

    const rows = verticalRecords.map(tmv3_step5OperatorRow_);
    const sheetName = TMV3.VERTICALS[vertical].sheet;
    const priorLastRow = tmv3_sheet_(sheetName).getLastRow();

    tmv3_replaceRows_(sheetName, TMV3_STEP1_HEADERS.slice(), rows);
    tmv3_step1FormatOperatorView_(
      sheetName,
      vertical,
      verticalRecords,
      priorLastRow
    );

    written[vertical] = rows.length;
  });

  return written;
}

function tmv3_step5OperatorRow_(record) {
  const row = tmv3_step4OperatorRow_(record);
  const step4 = record.step4 || {};
  const step5 = record.step5 || {};

  if (step4.disposition !== 'VERIFIED') {
    return row;
  }

  const disposition = step5.disposition || 'REVIEW';
  const tasks = step5.tasks || [];

  row[4] = tmv3_step5SyncChecklist_(record, tasks);
  row[6] = tmv3_step5TaskColumnSummary_(record, tasks, disposition);
  row[7] = tmv3_step5TaskScheduleSummary_(tasks);
  row[8] = tmv3_step5TaskStatusSummary_(tasks, disposition);
  row[9] = 'STEP 5 — ' + disposition;

  const details = [];

  if (disposition === 'REVIEW') {
    details.push(step5.reason || 'Task candidates require review.');
  } else if (disposition === 'NO_TASK' && step5.completedHistory) {
    details.push('No OPEN Task; completed history exists.');
  } else if (disposition === 'NO_TASK') {
    details.push('No applicable Task found.');
  }

  (step5.warnings || []).forEach(function(warning) {
    details.push(warning);
  });

  row[10] = details.join('\n');

  return row;
}

function tmv3_step5SyncChecklist_(record, tasks) {
  const step4 = record.step4 || {};
  const customer = step4.customer || null;
  const location = step4.location || null;
  const matched = (tasks || []).length > 0;

  let customerSync = '⏳';
  let locationSync = '⏳';
  let timeSync = '⏳';
  let assigneeSync = '⏳';
  let taskLinkSync = '⏳';

  if (matched) {
    customerSync = tasks.every(function(task) {
      const taskCustomer = tmv3_clean_(task['Customer ID']);
      return (
        !taskCustomer ||
        taskCustomer === tmv3_clean_(customer && customer['Customer ID'])
      );
    }) ? '✅' : '⚠';

    locationSync = tasks.every(function(task) {
      const taskLocation = tmv3_clean_(task['Location ID']);
      return (
        !taskLocation ||
        taskLocation === tmv3_clean_(location && location['Location ID'])
      );
    }) ? '✅' : '⚠';

    timeSync = tasks.every(function(task) {
      return (
        tmv3_sameMinute_(record.start, task['Start']) &&
        tmv3_sameMinute_(record.end, task['Due'])
      );
    }) ? '✅' : '⚠';

    if (record.vertical === 'Service' && record.technician) {
      assigneeSync = tasks.every(function(task) {
        return tmv3_norm_(task['Assignees']).indexOf(
          tmv3_norm_(record.technician)
        ) !== -1;
      }) ? '✅' : '⚠';
    }

    const linkedId = tmv3_clean_(record.existingTaskId);
    if (tasks.length === 1) {
      taskLinkSync =
        linkedId &&
        linkedId === tmv3_clean_(tasks[0]['Task ID'])
          ? '✅'
          : '⚠';
    }
  }

  return [
    'Customer ' + customerSync,
    'Location ' + locationSync,
    'Date / Time ' + timeSync,
    'Assignee ' + assigneeSync,
    'Task Link ' + taskLinkSync
  ].join('\n');
}

function tmv3_step5TaskSummary_(tasks, disposition) {
  if (!tasks || !tasks.length) {
    return disposition === 'NO_TASK'
      ? 'No OPEN Task'
      : 'Task not resolved';
  }

  return tasks.map(function(task) {
    const id = tmv3_clean_(task['Task ID']);
    const name = tmv3_clean_(task['Name']);
    return 'Task #' + id + (name ? ' · ' + name : '');
  }).join('\n');
}



function tmv3_step5GuardTaskLinkedAcrossEventDates_(records) {
  const usage = {};

  (records || []).forEach(function(record) {
    if (
      record.vertical !== 'PreInspection' ||
      !record.step5 ||
      record.step5.disposition !== 'MATCHED'
    ) {
      return;
    }

    (record.step5.tasks || []).forEach(function(task) {
      const taskId = tmv3_clean_(task && task['Task ID']);
      if (!taskId) return;

      if (!usage[taskId]) usage[taskId] = [];
      usage[taskId].push({
        record:record,
        eventDay:tmv3_step6LocalDay_(record.start)
      });
    });
  });

  Object.keys(usage).forEach(function(taskId) {
    const entries = usage[taskId];
    const days = tmv3_unique_(
      entries.map(function(entry) {
        return entry.eventDay;
      }).filter(Boolean)
    );

    if (days.length <= 1) return;

    entries.forEach(function(entry) {
      const prior = entry.record.step5 || {};
      entry.record.step5 = Object.assign({}, prior, {
        disposition:'REVIEW',
        code:'TASK_LINKED_TO_MULTIPLE_EVENT_DATES',
        reason:
          'Task #' + taskId +
          ' is linked to multiple PreInspection event dates (' +
          days.join(', ') +
          '). Do not reconcile either appointment automatically.'
      });
    });
  });

  return records;
}



function tmv3_step5TaskColumnSummary_(record, tasks, disposition) {
  const identity = tmv3_step4TaskColumnIdentity_(
    record,
    record.step4 || {}
  );

  const task = tmv3_step5TaskCombinedSummary_(tasks, disposition);

  return [identity, task]
    .filter(Boolean)
    .join('\n\n');
}

function tmv3_step5TaskCombinedSummary_(tasks, disposition) {
  if (!tasks || !tasks.length) {
    return disposition === 'NO_TASK'
      ? 'No OPEN Task'
      : 'Task not resolved';
  }

  return tasks.map(function(task) {
    const id = tmv3_clean_(task['Task ID']);
    const name = tmv3_step5DisplayTaskName_(task['Name']);
    const detail = 'Task #' + id + (name ? ' · ' + name : '');
    const schedule = tmv3_step5TaskScheduleOne_(task);
    const status = tmv3_step5DisplayStatus_(task['Status']);

    return [detail, schedule, status]
      .filter(Boolean)
      .join('\n\n');
  }).join('\n\n');
}

function tmv3_step5DisplayTaskName_(value) {
  return tmv3_clean_(value)
    .replace(/&amp;/gi, '&')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&nbsp;/gi, ' ');
}

function tmv3_step5TaskScheduleSummary_(tasks) {
  if (!tasks || !tasks.length) return '';

  return tasks.map(function(task) {
    return tmv3_step5TaskScheduleOne_(task);
  }).filter(Boolean).join('\n');
}

function tmv3_step5TaskScheduleOne_(task) {
  const startRaw = tmv3_clean_(task && task['Start']);
  const dueRaw = tmv3_clean_(task && task['Due']);
  if (!startRaw && !dueRaw) return '';

  const start = startRaw ? tmv3_parseDateTime_(startRaw) : null;
  const due = dueRaw ? tmv3_parseDateTime_(dueRaw) : null;

  const startOk = !!start;
  const dueOk = !!due;

  if (!startOk && !dueOk) return [startRaw, dueRaw].filter(Boolean).join(' - ');

  if (startOk && dueOk) {
    const startDay = Utilities.formatDate(start, TMV3_TIMEZONE, 'yyyy-MM-dd');
    const dueDay = Utilities.formatDate(due, TMV3_TIMEZONE, 'yyyy-MM-dd');

    if (startDay === dueDay) {
      return (
        Utilities.formatDate(start, TMV3_TIMEZONE, 'MMM d, yyyy') +
        ' (' +
        Utilities.formatDate(start, TMV3_TIMEZONE, 'h:mm a') +
        ' - ' +
        Utilities.formatDate(due, TMV3_TIMEZONE, 'h:mm a') +
        ')'
      );
    }

    return (
      Utilities.formatDate(start, TMV3_TIMEZONE, 'MMM d, yyyy h:mm a') +
      ' - ' +
      Utilities.formatDate(due, TMV3_TIMEZONE, 'MMM d, yyyy h:mm a')
    );
  }

  const only = startOk ? start : due;
  return Utilities.formatDate(only, TMV3_TIMEZONE, 'MMM d, yyyy h:mm a');
}

function tmv3_step5DisplayStatus_(status) {
  const clean = tmv3_norm_(status);

  if (['open','in progress','inprogress'].indexOf(clean) !== -1) {
    return 'Open';
  }

  if (['done','complete','completed','closed'].indexOf(clean) !== -1) {
    return 'Done';
  }

  if (['on hold','onhold'].indexOf(clean) !== -1) {
    return 'On Hold';
  }

  if (['cancelled','canceled'].indexOf(clean) !== -1) {
    return 'Cancelled';
  }

  return tmv3_clean_(status);
}

function tmv3_step5TaskStatusSummary_(tasks, disposition) {
  if (!tasks || !tasks.length) {
    return disposition === 'NO_TASK' ? 'NO OPEN TASK' : '';
  }

  return tmv3_unique_(
    tasks.map(function(task) {
      return tmv3_step5DisplayStatus_(task['Status']);
    }).filter(Boolean)
  ).join(', ');
}

function tmv3_step5Verify_(records, step4Records) {
  const priorKeys = {};
  const currentKeys = {};
  let missingDecision = 0;
  let verifiedWithoutDecision = 0;
  let nonVerifiedAdvanced = 0;

  (step4Records || []).forEach(function(record) {
    priorKeys[record.logicalKey] = true;
  });

  (records || []).forEach(function(record) {
    currentKeys[record.logicalKey] = true;

    const s4 = record.step4 ? record.step4.disposition : '';
    const s5 = record.step5 ? record.step5.disposition : '';

    if (!s5) missingDecision++;

    if (
      s4 === 'VERIFIED' &&
      ['MATCHED','MATCHED_MULTI','NO_TASK','REVIEW'].indexOf(s5) === -1
    ) {
      verifiedWithoutDecision++;
    }

    if (s4 !== 'VERIFIED' && s5 !== 'NOT_RUN') {
      nonVerifiedAdvanced++;
    }
  });

  const missing = Object.keys(priorKeys).filter(function(key) {
    return !currentKeys[key];
  });

  const introduced = Object.keys(currentKeys).filter(function(key) {
    return !priorKeys[key];
  });

  const expectedRows = {};
  const actualRows = {};

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    expectedRows[vertical] = records.filter(function(record) {
      return record.vertical === vertical;
    }).length;

    actualRows[vertical] = Math.max(
      0,
      tmv3_sheet_(TMV3.VERTICALS[vertical].sheet).getLastRow() - 1
    );
  });

  const rowCountsMatch = Object.keys(expectedRows).every(function(vertical) {
    return expectedRows[vertical] === actualRows[vertical];
  });

  return {
    pass:
      missingDecision === 0 &&
      verifiedWithoutDecision === 0 &&
      nonVerifiedAdvanced === 0 &&
      missing.length === 0 &&
      introduced.length === 0 &&
      rowCountsMatch,
    step4Records: (step4Records || []).length,
    step5Records: (records || []).length,
    missingDecision: missingDecision,
    verifiedWithoutDecision: verifiedWithoutDecision,
    nonVerifiedAdvanced: nonVerifiedAdvanced,
    missingFromStep5: missing.length,
    introducedByStep5: introduced.length,
    expectedRows: expectedRows,
    actualRows: actualRows
  };
}


// Original: 44_Step6_Task_Decision
/************************************************************
 * TM V3 — STAGE 6 TASK DECISION
 *
 * Narrow contract:
 * - Consumes exact Step-5 records.
 * - Decides MATCH / FULFILLED / CREATE / RECREATE / REVIEW.
 * - Does not create, patch, assign, or link anything.
 * - Same-day completed-task protection is shared across all verticals.
 ************************************************************/

function tmv3_step6TaskDecisionRun(reason, refreshSources) {
  tmv3_assertShadow_();

  const sourceSummary =
    refreshSources === true
      ? tmv3_step5RefreshTaskSources_()
      : tmv3_step5CachedSourceSummary_();

  const step2Snapshot = tmv3_step2CalendarRecords_();
  const step3Refs = tmv3_step3AnchorIndex_();
  const step3Records = tmv3_step3BusinessAnchorRecords_(
    step2Snapshot,
    step3Refs
  );

  const step4Refs = tmv3_step4IdentityIndex_();
  const step4Records = tmv3_step4IdentityRecords_(
    step3Records,
    step4Refs
  );

  const step5Refs = tmv3_step5TaskIndex_();
  const step5Records = tmv3_step5TaskRecords_(
    step4Records,
    step5Refs
  );

  const records = tmv3_step6DecisionRecords_(step5Records);
  const counts = tmv3_step6Counts_(records);
  const write = tmv3_step6WriteOperatorViews_(records);
  const verification = tmv3_step6Verify_(records, step5Records);

  const result = {
    version: TMV3.VERSION,
    executionStage: tmv3_executionStage_(),
    stage: 6,
    status: verification.pass ? 'PASS' : 'REVIEW',
    mode: 'TASK_DECISION_ONLY',
    reason: tmv3_clean_(reason || 'MANUAL'),
    sourceSummary: sourceSummary,
    counts: counts,
    write: write,
    verification: verification,
    taskWritesPerformed: false,
    calendarWritesPerformed: false,
    strivenWritesPerformed: false,
    reconciliationPerformed: false
  };

  tmv3_audit_(
    'SYSTEM','','','STEP6_TASK_DECISION',
    result.status,
    JSON.stringify(result)
  );

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function tmv3_step6TaskDecisionRunCached(reason) {
  return tmv3_step6TaskDecisionRun(
    reason || 'CACHED_TASK_SOURCES',
    false
  );
}

function tmv3_step6DecisionRecords_(step5Records) {
  return (step5Records || []).map(function(record) {
    const step5 = record.step5 || {};
    let decision;

    if (step5.disposition === 'NOT_RUN') {
      decision = tmv3_step6Decision_(
        'NOT_RUN',
        'STEP5_NOT_RUN',
        'Task decision did not run because Step 5 did not run.',
        []
      );

    } else if (step5.disposition === 'REVIEW') {
      decision = tmv3_step6Decision_(
        'REVIEW',
        'STEP5_REVIEW',
        step5.reason || 'Step 5 requires review.',
        step5.tasks || []
      );

    } else if (step5.disposition === 'MATCHED') {
      decision = tmv3_step6Decision_(
        'MATCH_EXISTING',
        'OPEN_TASK_MATCHED',
        'Use the verified OPEN Task.',
        step5.tasks || []
      );

    } else if (step5.disposition === 'MATCHED_MULTI') {
      decision = tmv3_step6Decision_(
        'MATCH_EXISTING_MULTI',
        'OPEN_MULTI_TASKS_MATCHED',
        'Use the verified OPEN Service fireplace Tasks.',
        step5.tasks || []
      );

    } else if (step5.disposition === 'NO_TASK') {
      decision = tmv3_step6NoOpenTaskDecision_(record);

    } else {
      decision = tmv3_step6Decision_(
        'REVIEW',
        'UNKNOWN_STEP5_DISPOSITION',
        'Unknown Step 5 disposition: ' + tmv3_clean_(step5.disposition),
        []
      );
    }

    return Object.assign({}, record, { step6: decision });
  });
}

function tmv3_step6NoOpenTaskDecision_(record) {
  const history = (record.step5 && record.step5.historyTasks) || [];
  const eventDay = tmv3_step6LocalDay_(record.start);

  if (record.vertical === 'Service') {
    return tmv3_step6Decision_(
      'BLOCKED',
      'SERVICE_TASK_CREATION_OWNED_BY_STRIVEN',
      'Service Task creation/recreation is owned by the Striven Work Order workflow. Task Mapping waits for or reconciles that Service Task instead of creating one.',
      [],
      history
    );
  }

  if (!history.length) {
    return tmv3_step6Decision_(
      'CREATE_TASK',
      'NO_TASK_OR_HISTORY',
      'No applicable OPEN Task or historical Task was found.',
      []
    );
  }

  const fulfilled = history.filter(function(task) {
    return tmv3_step6IsFulfilledStatus_(task['Status']);
  });

  const cancelled = history.filter(function(task) {
    return tmv3_step6IsCancelledStatus_(task['Status']);
  });

  const sameDayFulfilled = fulfilled.filter(function(task) {
    return tmv3_step6TaskTouchesEventDay_(record, task);
  });

  if (sameDayFulfilled.length) {
    return tmv3_step6Decision_(
      'FULFILLED_NO_RECREATE',
      'COMPLETED_TASK_ON_EVENT_DAY',
      'Applicable completed Task is scheduled on the Calendar event date. Do not recreate.',
      sameDayFulfilled,
      history
    );
  }

  const sameDayCancelled = cancelled.filter(function(task) {
    return tmv3_step6TaskTouchesEventDay_(record, task);
  });

  if (sameDayCancelled.length) {
    return tmv3_step6Decision_(
      'REVIEW',
      'CANCELLED_TASK_ON_EVENT_DAY',
      'A cancelled Task is scheduled on the Calendar event date. Do not auto-create or recreate.',
      sameDayCancelled,
      history
    );
  }

  if (cancelled.length) {
    return tmv3_step6Decision_(
      'REVIEW',
      'CANCELLED_TASK_HISTORY',
      'Cancelled Task history exists for this appointment context. Cancelled history never authorizes automatic CREATE or RECREATE.',
      cancelled,
      history
    );
  }

  const historyDays = history.map(function(task) {
    return tmv3_step6TaskLocalDay_(task);
  });

  if (historyDays.some(function(day) { return !day; })) {
    return tmv3_step6Decision_(
      'REVIEW',
      'HISTORY_DATE_UNPROVEN',
      'Historical Task exists but its appointment date cannot be proven.',
      history,
      history
    );
  }

  if (
    eventDay &&
    historyDays.some(function(day) { return day > eventDay; })
  ) {
    return tmv3_step6Decision_(
      'REVIEW',
      'HISTORY_AFTER_EVENT_DATE',
      'Historical Task evidence falls after the Calendar event date.',
      history,
      history
    );
  }

  return tmv3_step6Decision_(
    'RECREATE_TASK',
    'ONLY_OLDER_HISTORY_REMAINS',
    'No OPEN Task remains and only older fulfilled/completed Task history exists.',
    history,
    history
  );
}

function tmv3_step6Decision_(
  disposition,
  code,
  reason,
  tasks,
  historyTasks
) {
  return {
    disposition: disposition,
    code: code,
    reason: reason,
    tasks: tasks || [],
    historyTasks: historyTasks || [],
    writesAuthorized: false
  };
}

function tmv3_step6IsFulfilledStatus_(status) {
  const clean = tmv3_norm_(status);
  return ['done','complete','completed','closed'].indexOf(clean) !== -1;
}

function tmv3_step6IsCancelledStatus_(status) {
  const clean = tmv3_norm_(status);
  return ['canceled','cancelled'].indexOf(clean) !== -1;
}

function tmv3_step6TaskTouchesEventDay_(record, task) {
  return (
    tmv3_sameLocalDay_(record.start, task && task['Start']) ||
    tmv3_sameLocalDay_(record.start, task && task['Due'])
  );
}

function tmv3_step6TaskLocalDay_(task) {
  const start = tmv3_clean_(task && task['Start']);
  const due = tmv3_clean_(task && task['Due']);
  return tmv3_step6LocalDay_(start || due);
}

function tmv3_step6LocalDay_(value) {
  if (!value) return '';
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, TMV3_TIMEZONE, 'yyyy-MM-dd');
}

function tmv3_step6Counts_(records) {
  const counts = {
    total:0,
    matchExisting:0,
    matchExistingMulti:0,
    fulfilledNoRecreate:0,
    createTask:0,
    recreateTask:0,
    review:0,
    notRun:0,
    byVertical:{}
  };

  (records || []).forEach(function(record) {
    counts.total++;

    if (!counts.byVertical[record.vertical]) {
      counts.byVertical[record.vertical] = {
        total:0,
        matchExisting:0,
        matchExistingMulti:0,
        fulfilledNoRecreate:0,
        createTask:0,
        recreateTask:0,
        review:0,
        notRun:0
      };
    }

    const bucket = counts.byVertical[record.vertical];
    bucket.total++;

    const d = record.step6 ? record.step6.disposition : 'NOT_RUN';
    const map = {
      MATCH_EXISTING:'matchExisting',
      MATCH_EXISTING_MULTI:'matchExistingMulti',
      FULFILLED_NO_RECREATE:'fulfilledNoRecreate',
      CREATE_TASK:'createTask',
      RECREATE_TASK:'recreateTask',
      REVIEW:'review',
      NOT_RUN:'notRun'
    };

    const key = map[d] || 'review';
    counts[key]++;
    bucket[key]++;
  });

  return counts;
}

function tmv3_step6WriteOperatorViews_(records) {
  const written = {};

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    const verticalRecords = (records || [])
      .filter(function(record) { return record.vertical === vertical; })
      .sort(function(a, b) {
        const aTime = a.start instanceof Date ? a.start.getTime() : 0;
        const bTime = b.start instanceof Date ? b.start.getTime() : 0;
        if (aTime !== bTime) return aTime - bTime;
        return tmv3_clean_(a.title).localeCompare(tmv3_clean_(b.title));
      });

    const rows = verticalRecords.map(tmv3_step6OperatorRow_);
    const sheetName = TMV3.VERTICALS[vertical].sheet;
    const priorLastRow = tmv3_sheet_(sheetName).getLastRow();

    tmv3_replaceRows_(sheetName, TMV3_STEP1_HEADERS.slice(), rows);
    tmv3_step1FormatOperatorView_(
      sheetName,
      vertical,
      verticalRecords,
      priorLastRow
    );

    written[vertical] = rows.length;
  });

  return written;
}

function tmv3_step6OperatorRow_(record) {
  const row = tmv3_step5OperatorRow_(record);
  const decision = record.step6 || {};

  row[9] = 'STEP 6 — ' + (decision.disposition || 'NOT_RUN');

  if (
    decision.disposition === 'FULFILLED_NO_RECREATE' ||
    decision.disposition === 'RECREATE_TASK'
  ) {
    const historical = decision.tasks || decision.historyTasks || [];
    if (historical.length) {
      row[6] = tmv3_step5TaskColumnSummary_(
        record,
        historical,
        'HISTORY'
      );
      row[7] = tmv3_step5TaskScheduleSummary_(historical);
      row[8] = tmv3_step5TaskStatusSummary_(historical, 'HISTORY');
    }
  }

  if (decision.disposition === 'MATCH_EXISTING') {
    row[10] = 'Use existing OPEN Task.';
  } else if (decision.disposition === 'MATCH_EXISTING_MULTI') {
    row[10] = 'Use existing OPEN fireplace Tasks.';
  } else if (decision.disposition === 'FULFILLED_NO_RECREATE') {
    row[10] = 'Completed Task covers this event date — do not recreate.';
  } else if (decision.disposition === 'CREATE_TASK') {
    row[10] = 'No Task/history found — Task creation candidate.';
  } else if (decision.disposition === 'RECREATE_TASK') {
    row[10] = 'Only older Task history remains — Task recreation candidate.';
  } else if (decision.disposition === 'REVIEW') {
    row[10] = decision.reason || 'Task decision requires review.';
  }

  return row;
}

function tmv3_step6Verify_(records, step5Records) {
  const prior = {};
  const current = {};
  let missingDecision = 0;
  let invalidTransition = 0;
  let sameDayCompletedRecreateError = 0;

  (step5Records || []).forEach(function(record) {
    prior[record.logicalKey] = true;
  });

  (records || []).forEach(function(record) {
    current[record.logicalKey] = true;

    const s5 = record.step5 ? record.step5.disposition : '';
    const s6 = record.step6 ? record.step6.disposition : '';

    if (!s6) missingDecision++;

    const allowed = {
      NOT_RUN:['NOT_RUN'],
      REVIEW:['REVIEW'],
      MATCHED:['MATCH_EXISTING'],
      MATCHED_MULTI:['MATCH_EXISTING_MULTI'],
      NO_TASK:[
        'FULFILLED_NO_RECREATE',
        'CREATE_TASK',
        'RECREATE_TASK',
        'REVIEW',
        'BLOCKED'
      ]
    };

    if (!allowed[s5] || allowed[s5].indexOf(s6) === -1) {
      invalidTransition++;
    }

    if (
      s5 === 'NO_TASK' &&
      ((record.step5 && record.step5.historyTasks) || [])
        .some(function(task) {
          return (
            tmv3_step6IsFulfilledStatus_(task['Status']) &&
            tmv3_step6TaskTouchesEventDay_(record, task)
          );
        }) &&
      // REVIEW/BLOCKED are safe holds, including Striven-owned Service
      // creation and conflicting history. Only a create decision violates
      // same-day completion protection; transition checks reject other drift.
      (s6 === 'CREATE_TASK' || s6 === 'RECREATE_TASK')
    ) {
      sameDayCompletedRecreateError++;
    }
  });

  const missing = Object.keys(prior).filter(function(key) {
    return !current[key];
  });

  const introduced = Object.keys(current).filter(function(key) {
    return !prior[key];
  });

  const expectedRows = {};
  const actualRows = {};

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    expectedRows[vertical] = records.filter(function(record) {
      return record.vertical === vertical;
    }).length;

    actualRows[vertical] = Math.max(
      0,
      tmv3_sheet_(TMV3.VERTICALS[vertical].sheet).getLastRow() - 1
    );
  });

  const rowCountsMatch = Object.keys(expectedRows).every(function(vertical) {
    return expectedRows[vertical] === actualRows[vertical];
  });

  return {
    pass:
      missingDecision === 0 &&
      invalidTransition === 0 &&
      sameDayCompletedRecreateError === 0 &&
      missing.length === 0 &&
      introduced.length === 0 &&
      rowCountsMatch,
    step5Records:(step5Records || []).length,
    step6Records:(records || []).length,
    missingDecision:missingDecision,
    invalidTransition:invalidTransition,
    sameDayCompletedRecreateError:sameDayCompletedRecreateError,
    missingFromStep6:missing.length,
    introducedByStep6:introduced.length,
    expectedRows:expectedRows,
    actualRows:actualRows
  };
}


// Original: 46_Step7_Reconciliation
/************************************************************
 * TM V3 — STEP 7 RECONCILIATION PLAN
 *
 * Purpose:
 * - consume the verified Step 6 decision; never re-decide identity/task;
 * - fresh-read canonical Striven Tasks on demand;
 * - compare expected vs actual relationships, dates and assignments;
 * - fresh-verify expected Contact ownership before any future mutation;
 * - write a hidden, inspectable reconciliation plan;
 * - perform ZERO Striven/Calendar mutations.
 ************************************************************/

const TMV3_STEP7_HEADERS = Object.freeze([
  'Vertical',
  'Event ID',
  'Logical Key',
  'Event',
  'Step 6 Decision',
  'Task ID',
  'Task Status',
  'Expected Customer ID',
  'Actual Customer ID',
  'Customer Check',
  'Expected Order ID',
  'Actual Order ID',
  'Order Check',
  'Expected Location ID',
  'Actual Location ID',
  'Location Check',
  'Expected Requested By ID',
  'Expected Requested By Type',
  'Actual Requested By ID',
  'Actual Requested By Type',
  'Requested By Check',
  'Contact Ownership',
  'Calendar Start',
  'Task Start',
  'Start Check',
  'Calendar End',
  'Task Due',
  'End Check',
  'Desired Assignment',
  'Actual Assignment',
  'Assignment Check',
  'Relationship Plan',
  'Write Gate',
  'Blocker',
  'Evidence',
  'Source Task IDs',
  'Engine Version',
  'Planned At',
  'Read Status'
]);

const TMV3_STEP7_ACTION = Object.freeze({
  PATCH_DATES:'PATCH_DATES',
  PATCH_LOCATION:'PATCH_LOCATION',
  PATCH_ORDER:'PATCH_ORDER',
  PATCH_REQUESTED_BY:'PATCH_REQUESTED_BY',
  PATCH_ASSIGNMENTS:'PATCH_ASSIGNMENTS',
  PATCH_CALENDAR_TITLE:'PATCH_CALENDAR_TITLE',
  PATCH_TASK_NAME:'PATCH_TASK_NAME',
  VERIFY_CALENDAR_LINKS:'VERIFY_CALENDAR_LINKS',
  CREATE_LOCATION:'CREATE_LOCATION',
  CREATE_TASK:'CREATE_TASK',
  RECREATE_TASK:'RECREATE_TASK'
});

const TMV3_STEP7_ACTION_VALUES = Object.freeze(
  Object.keys(TMV3_STEP7_ACTION).map(function(key) {
    return TMV3_STEP7_ACTION[key];
  })
);

function tmv3_step7ValidateExecutionContract_(contract) {
  if (!contract || typeof contract !== 'object') {
    throw new Error('Step 7 execution contract is required.');
  }

  const required = [
    'vertical','eventId','logicalKey','disposition','plan',
    'actions','engineVersion','inputFingerprint'
  ];

  required.forEach(function(key) {
    if (
      contract[key] === undefined ||
      contract[key] === null ||
      (key !== 'actions' && tmv3_clean_(contract[key]) === '')
    ) {
      throw new Error('Step 7 execution contract is missing ' + key + '.');
    }
  });

  if (!TMV3.VERTICALS[contract.vertical]) {
    throw new Error('Step 7 execution contract has an unknown vertical.');
  }

  if (!Array.isArray(contract.actions)) {
    throw new Error('Step 7 execution contract actions must be an array.');
  }

  const seen = {};
  contract.actions.forEach(function(action) {
    if (TMV3_STEP7_ACTION_VALUES.indexOf(action) === -1) {
      throw new Error('Step 7 execution contract has unsupported action ' + action + '.');
    }
    if (seen[action]) {
      throw new Error('Step 7 execution contract contains duplicate action ' + action + '.');
    }
    seen[action] = true;
  });

  if (tmv3_clean_(contract.blocker) && contract.actions.length) {
    throw new Error('Blocked Step 7 execution contract must not expose executable actions.');
  }

  if (
    contract.actions.indexOf(TMV3_STEP7_ACTION.CREATE_LOCATION) !== -1 &&
    !tmv3_clean_(contract.expectedLocationAddress)
  ) {
    throw new Error(
      'Step 7 CREATE_LOCATION contract is missing the verified Calendar job-site address.'
    );
  }

  return contract;
}




function tmv3_step7FreshPlansForEvent_(vertical, eventId) {
  const wantedVertical = tmv3_clean_(vertical);
  const wantedEventId = tmv3_clean_(eventId);

  if (!TMV3.VERTICALS[wantedVertical] || !wantedEventId) {
    throw new Error('Fresh Step 7 event planning requires vertical and Event ID.');
  }

  const freshEvent = tmv3_findFreshEventRecord_(wantedVertical, wantedEventId);
  const step2 = tmv3_step2CalendarRecords_([freshEvent]);
  const step3 = tmv3_step3BusinessAnchorRecords_(step2, tmv3_step3AnchorIndex_());
  const step4 = tmv3_step4IdentityRecords_(step3, tmv3_step4IdentityIndex_());
  const runtime = { taskById:{}, contactByKey:{}, organizerByEmail:{} };
  const taskRefs = tmv3_step5TaskIndex_();
  tmv3_step7HydrateLinkedPreInspectionTasks_(step4, taskRefs, runtime);
  const step5 = tmv3_step5TaskRecords_(step4, taskRefs);
  const step6 = tmv3_step6DecisionRecords_(step5);

  return tmv3_step7Plans_(step6, runtime);
}

function tmv3_step7FreshExecutionContract_(vertical, eventId, taskId) {
  const wantedTaskId = Number(taskId || 0);
  const plans = tmv3_step7FreshPlansForEvent_(vertical, eventId);
  const matches = plans.filter(function(plan) {
    return wantedTaskId
      ? Number(plan.taskId || 0) === wantedTaskId
      : !Number(plan.taskId || 0);
  });

  if (matches.length !== 1) {
    throw new Error(
      'Expected one fresh canonical Step 7 execution contract for ' +
      tmv3_clean_(vertical) + ' Event ' + tmv3_clean_(eventId) +
      (wantedTaskId ? ' Task ' + wantedTaskId : '') +
      '; found ' + matches.length + '.'
    );
  }

  return tmv3_step7ValidateExecutionContract_(matches[0]);
}

function tmv3_step7ReconciliationRun(
  reason,
  refreshSources,
  verticalFilter,
  batchOptions
) {
  tmv3_assertShadow_();

  if (tmv3_executionStage_() < 7) {
    throw new Error(
      'Execution stage ' + tmv3_executionStage_() +
      ' blocks Step 7 reconciliation.'
    );
  }

  const filter = tmv3_clean_(verticalFilter);
  if (filter && !TMV3.VERTICALS[filter]) {
    throw new Error('Unknown Step 7 vertical filter: ' + filter);
  }

  tmv3_resetRuntimeMetrics_();

  const sourceSummary =
    refreshSources === true
      ? tmv3_step5RefreshTaskSources_()
      : tmv3_step5CachedSourceSummary_();

  batchOptions = batchOptions || {};
  const batchOffset = Math.max(0, Number(batchOptions.offset || 0));
  const requestedLimit = Math.max(0, Number(batchOptions.limit || 0));
  const batchEnabled = requestedLimit > 0;

  const allStep2Snapshot = tmv3_step2CalendarRecords_();
  const filteredStep2Snapshot = filter
    ? allStep2Snapshot.filter(function(record) {
        return record.vertical === filter;
      })
    : allStep2Snapshot;

  // API-budget rule: for batched Step 7 runs, slice the Calendar input
  // BEFORE identity/task resolution. Previously batching happened after
  // Step 4/5, so a 35-row Service batch still resolved all Service events
  // and could exceed Striven's 100-calls/minute ceiling.
  const step2Snapshot = batchEnabled
    ? filteredStep2Snapshot.slice(
        batchOffset,
        batchOffset + requestedLimit
      )
    : filteredStep2Snapshot;

  const step3Refs = tmv3_step3AnchorIndex_();
  const step3Records = tmv3_step3BusinessAnchorRecords_(
    step2Snapshot,
    step3Refs
  );

  const step4Refs = tmv3_step4IdentityIndex_();
  const step4Records = tmv3_step4IdentityRecords_(
    step3Records,
    step4Refs
  );

  const step5Refs = tmv3_step5TaskIndex_();
  const step5Records = tmv3_step5TaskRecords_(
    step4Records,
    step5Refs
  );

  const step6Records = tmv3_step6DecisionRecords_(step5Records);
  const selected = step6Records;
  const batch = {
    enabled: batchEnabled,
    offset: batchOffset,
    limit: batchEnabled ? requestedLimit : filteredStep2Snapshot.length,
    selected: step2Snapshot.length,
    step7Rows: selected.length,
    totalAvailable: filteredStep2Snapshot.length,
    nextOffset:
      batchEnabled &&
      batchOffset + step2Snapshot.length < filteredStep2Snapshot.length
        ? batchOffset + step2Snapshot.length
        : null,
    complete:
      !batchEnabled ||
      batchOffset + step2Snapshot.length >= filteredStep2Snapshot.length
  };

  const runtime = {
    taskById: {},
    contactByKey: {},
    organizerByEmail: {}
  };

  const plans = tmv3_step7Plans_(selected, runtime);
  const write = tmv3_step7WritePlans_(plans, filter, batch);
  const visibleWrite = tmv3_publishLatestStep7ToVisibleSheets();
  const verification = tmv3_step7Verify_(selected, plans);
  const counts = tmv3_step7Counts_(plans);

  const result = {
    version: TMV3.VERSION,
    executionStage: tmv3_executionStage_(),
    stage: 7,
    status: verification.pass ? 'PASS' : 'REVIEW',
    mode: 'RECONCILIATION_PLAN_ONLY',
    reason: tmv3_clean_(reason || 'MANUAL'),
    verticalFilter: filter || 'ALL',
    batch: batch,
    sourceSummary: sourceSummary,
    counts: counts,
    write: write,
    visibleWrite: visibleWrite,
    verification: verification,
    runtime: tmv3_runtimeMetrics_(),
    reconciliationPlanned: true,
    taskWritesPerformed: false,
    calendarWritesPerformed: false,
    strivenWritesPerformed: false
  };

  tmv3_audit_(
    'SYSTEM',
    '',
    '',
    'STEP7_RECONCILIATION_PLAN',
    result.status,
    JSON.stringify({
      verticalFilter: result.verticalFilter,
      counts: counts,
      verification: verification,
      runtime: result.runtime,
      taskWritesPerformed: false,
      calendarWritesPerformed: false,
      strivenWritesPerformed: false
    })
  );

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function tmv3_step7CreateCandidateScan(reason, refreshSources) {
  tmv3_assertShadow_();

  if (tmv3_executionStage_() < 7) {
    throw new Error(
      'Execution stage ' + tmv3_executionStage_() +
      ' blocks Step 7 CREATE candidate scan.'
    );
  }

  tmv3_resetRuntimeMetrics_();

  const sourceSummary =
    refreshSources === true
      ? tmv3_step5RefreshTaskSources_()
      : tmv3_step5CachedSourceSummary_();

  const step2Snapshot = tmv3_step2CalendarRecords_();
  const step3Records = tmv3_step3BusinessAnchorRecords_(
    step2Snapshot,
    tmv3_step3AnchorIndex_()
  );
  const step4Records = tmv3_step4IdentityRecords_(
    step3Records,
    tmv3_step4IdentityIndex_()
  );
  const step5Records = tmv3_step5TaskRecords_(
    step4Records,
    tmv3_step5TaskIndex_()
  );
  const step6Records = tmv3_step6DecisionRecords_(step5Records);

  // Critical API-budget rule:
  // only CREATE / RECREATE dispositions advance into Step 7 planning here.
  // Existing-task rows require fresh Task GETs and can exhaust Striven's
  // per-minute budget before CREATE ownership checks are reached.
  const candidates = step6Records.filter(function(record) {
    const disposition = tmv3_clean_(
      record && record.step6 && record.step6.disposition
    );
    return disposition === 'CREATE_TASK' || disposition === 'RECREATE_TASK';
  });

  const runtime = {
    taskById: {},
    contactByKey: {},
    organizerByEmail: {}
  };

  const plans = tmv3_step7Plans_(candidates, runtime);
  const eligible = plans.filter(function(plan) {
    return (
      (
        String(plan.plan || '').indexOf('CREATE_TASK') !== -1 ||
        String(plan.plan || '').indexOf('RECREATE_TASK') !== -1
      ) &&
      !tmv3_clean_(plan.blocker)
    );
  });

  const result = {
    version: TMV3.VERSION,
    executionStage: tmv3_executionStage_(),
    stage: 7,
    status: 'PASS',
    mode: 'CREATE_CANDIDATE_SCAN_ONLY',
    reason: tmv3_clean_(reason || 'CREATE_CANDIDATE_SCAN'),
    sourceSummary: sourceSummary,
    candidateRecords: candidates.length,
    plans: plans,
    eligible: eligible,
    runtime: tmv3_runtimeMetrics_(),
    taskWritesPerformed: false,
    calendarWritesPerformed: false,
    strivenWritesPerformed: false
  };

  tmv3_audit_(
    'SYSTEM',
    '',
    '',
    'STEP7_CREATE_CANDIDATE_SCAN',
    'PASS',
    JSON.stringify({
      candidateRecords: candidates.length,
      eligible: eligible.length,
      runtime: result.runtime,
      taskWritesPerformed: false,
      calendarWritesPerformed: false,
      strivenWritesPerformed: false
    })
  );

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function tmv3_step7ReconciliationRunCached(reason, verticalFilter) {
  return tmv3_step7ReconciliationRun(
    reason || 'CACHED_TASK_SOURCES',
    false,
    verticalFilter || ''
  );
}



function tmv3_step7ReconciliationBatchRun(
  reason,
  verticalFilter,
  offset,
  limit,
  refreshSources
) {
  const filter = tmv3_clean_(verticalFilter);
  if (!filter) {
    throw new Error('Step 7 batch run requires one vertical.');
  }

  return tmv3_step7ReconciliationRun(
    reason || 'BATCHED_STEP7_VERIFY',
    refreshSources === true,
    filter,
    {
      offset: Number(offset || 0),
      limit: Number(limit || 20)
    }
  );
}

function tmv3_step7FreshPlanForTask_(vertical, eventId, taskId) {
  const wantedTaskId = Number(taskId || 0);
  if (!wantedTaskId) {
    throw new Error('Fresh Step 7 plan requires a positive Task ID.');
  }
  return tmv3_step7FreshExecutionContract_(vertical, eventId, wantedTaskId);
}

function tmv3_step7Plans_(records, runtime) {
  const plans = [];

  (records || []).forEach(function(record) {
    const decision = record.step6 || {};
    const disposition = tmv3_clean_(decision.disposition) || 'NOT_RUN';
    const sourceTasks = decision.tasks || [];
    const sourceTaskIds = tmv3_unique_(
      sourceTasks.map(function(task) {
        return tmv3_clean_(task && task['Task ID']);
      })
    );

    if (disposition === 'NOT_RUN') {
      plans.push(
        tmv3_step7NoTaskPlan_(
          record,
          disposition,
          'NO_RECONCILIATION_STEP6_NOT_RUN',
          decision.reason || 'Step 6 did not run.',
          sourceTaskIds
        )
      );
      return;
    }

    if (disposition === 'CREATE_TASK' || disposition === 'RECREATE_TASK') {
      plans.push(
        tmv3_step7CreatePlan_(
          record,
          disposition,
          sourceTaskIds,
          runtime
        )
      );
      return;
    }

    if (disposition === 'REVIEW') {
      plans.push(
        tmv3_step7NoTaskPlan_(
          record,
          disposition,
          'REVIEW_NO_AUTOMATIC_MUTATION',
          decision.reason || 'Step 6 requires review.',
          sourceTaskIds
        )
      );
      return;
    }

    if (disposition === 'BLOCKED') {
      plans.push(
        tmv3_step7NoTaskPlan_(
          record,
          disposition,
          'BLOCKED_EXTERNAL_NO_AUTOMATIC_MUTATION',
          decision.reason || 'Step 6 is blocked by an external prerequisite.',
          sourceTaskIds
        )
      );
      return;
    }

    if (!sourceTasks.length) {
      plans.push(
        tmv3_step7NoTaskPlan_(
          record,
          disposition,
          'REVIEW_STEP6_TASK_ID_MISSING',
          'Step 6 requires an existing Task but exposed no Task ID.',
          sourceTaskIds
        )
      );
      return;
    }

    sourceTasks.forEach(function(task) {
      plans.push(
        tmv3_step7ExistingTaskPlan_(
          record,
          disposition,
          task,
          sourceTaskIds,
          runtime
        )
      );
    });
  });

  return plans;
}

function tmv3_step7Expected_(record, runtime, verifyContactOwnership) {
  const step3 = record.step3 || {};
  const step4 = record.step4 || {};
  const anchor = step3.anchor || {};
  const cfg = TMV3.VERTICALS[record.vertical];

  const customerId = tmv3_clean_(
    step4.customer && step4.customer['Customer ID']
  );
  const locationId = tmv3_clean_(
    step4.location && step4.location['Location ID']
  );
  const identityContactId = tmv3_clean_(
    (step4.contact && step4.contact['Contact ID']) ||
    (record.vertical !== 'PreInspection' && anchor.orderContactIdEvidence) || ''
  );
  const orderId = cfg.orderRequired
    ? tmv3_clean_(anchor.orderId)
    : '';

  const expected = {
    customerId: customerId,
    locationId: locationId,
    locationStatus: tmv3_clean_(step4.locationStatus),
    locationAddress:
      tmv3_clean_(step4.locationStatus) === 'CREATE_REQUIRED'
        ? tmv3_clean_(record.location)
        : '',
    identityContactId: identityContactId,
    orderId: orderId,
    requestedById: '',
    requestedByType: '',
    requestedByError: '',
    contactOwnership: identityContactId ? 'NOT_CHECKED' : 'N/A',
    evidence: (step4.evidence || []).slice()
  };

  if (record.vertical === 'PreInspection') {
    const organizerKey = tmv3_norm_(record.organizer || '');

    if (!runtime.organizerByEmail[organizerKey]) {
      try {
        runtime.organizerByEmail[organizerKey] =
          tmv3_resolveOrganizerEmployee_(record);
      } catch (err) {
        runtime.organizerByEmail[organizerKey] = {
          status: 'REVIEW',
          reason: String(err && err.message || err)
        };
      }
    }

    const requested = runtime.organizerByEmail[organizerKey];
    if (requested && requested.status === 'MATCHED') {
      expected.requestedById = String(requested.employee.id || '');
      expected.requestedByType = 'employee';
      expected.evidence.push(
        requested.evidence || 'CALENDAR_ORGANIZER_EMAIL_EXACT_EMPLOYEE'
      );
    } else {
      expected.requestedByError =
        (requested && requested.reason) ||
        'PreInspection Requested By employee is unresolved.';
    }

    // Customer Contact remains identity evidence only for PreInspection.
    expected.contactOwnership = 'NOT_APPLICABLE_PREINSPECTION';
    return expected;
  }

  if (identityContactId) {
    expected.requestedById = identityContactId;
    expected.requestedByType = 'contact';

    if (verifyContactOwnership) {
      const ownership = tmv3_step7ContactOwnership_(
        identityContactId,
        customerId,
        runtime
      );
      expected.contactOwnership = ownership.status;
      if (ownership.error) {
        expected.requestedByError = ownership.error;
      }
    }
  }

  return expected;
}

function tmv3_step7ContactOwnership_(contactId, customerId, runtime) {
  const key = tmv3_clean_(contactId) + '|' + tmv3_clean_(customerId);

  if (runtime.contactByKey[key]) {
    return runtime.contactByKey[key];
  }

  let result;

  try {
    const contact = tmv3_getContactById_(contactId, customerId);
    result = {
      status: contact && contact.__ownershipVerified
        ? 'VERIFIED'
        : 'UNVERIFIED',
      error:
        contact && contact.__ownershipVerified
          ? ''
          : (
              'Contact ' + contactId +
              ' ownership is not proven for Customer ' + customerId + '.'
            )
    };
  } catch (err) {
    result = {
      status: 'READ_FAILED',
      error:
        'Contact ' + contactId + ' ownership read failed: ' +
        String(err && err.message || err)
    };
  }

  runtime.contactByKey[key] = result;
  return result;
}

function tmv3_step7FreshTask_(taskId, runtime) {
  const id = tmv3_clean_(taskId);

  if (runtime.taskById[id]) {
    return runtime.taskById[id];
  }

  let result;

  try {
    const raw = tmv3_rawTaskById_(id);
    const normalized = tmv3_normalizeV2TaskModel_(raw);
    result = {
      ok: true,
      raw: raw,
      task: normalized,
      requestedBy: tmv3_step7RequestedByFromRaw_(raw),
      assignments: tmv3_taskAssignmentsFromRaw_(raw)
    };
  } catch (err) {
    result = {
      ok: false,
      error: String(err && err.message || err),
      raw: null,
      task: null,
      requestedBy: { id:'', type:'' },
      assignments: []
    };
  }

  runtime.taskById[id] = result;
  return result;
}

function tmv3_step7RequestedByFromRaw_(raw) {
  raw = raw || {};

  const requested = raw.requestedBy || raw.RequestedBy || {};
  const requestedContact =
    raw.requestedByContact || raw.RequestedByContact || {};

  let id = tmv3_clean_(
    tmv3_first_(requested, ['id','Id'])
  );
  let type = tmv3_norm_(
    tmv3_first_(requested, ['type','Type'])
  );

  if (!id) {
    id = tmv3_clean_(
      tmv3_first_(
        requestedContact,
        ['contactId','ContactId','id','Id']
      )
    );

    if (id && !type) type = 'contact';
  }

  if (!id) {
    id = tmv3_clean_(
      tmv3_first_(
        raw,
        [
          'contactId',
          'ContactId',
          'requestedByContactId',
          'RequestedByContactId'
        ]
      )
    );

    if (id && !type) type = 'contact';
  }

  if (
    type &&
    type !== 'contact' &&
    type !== 'employee'
  ) {
    type = tmv3_clean_(
      tmv3_first_(requested, ['type','Type'])
    ).toLowerCase();
  }

  return {
    id: id,
    type: type
  };
}


function tmv3_step7TitleState_(record, actualTaskName, taskStatus, historical) {
  const desiredCalendar = tmv3_desiredCalendarTitle_(record);
  const desiredTask = tmv3_desiredTaskName_(
    record,
    { sourceTaskTitle:actualTaskName || '' }
  );

  const actualCalendarTitle = tmv3_clean_(record && record.title);
  const actualName = tmv3_clean_(actualTaskName);
  const openTask =
    !historical &&
    tmv3_taskNameNormalizationStatusAllowed_(taskStatus);

  const calendarCheck =
    desiredCalendar.status === 'READY'
      ? (
          actualCalendarTitle === tmv3_clean_(desiredCalendar.value)
            ? 'MATCH'
            : 'MISMATCH'
        )
      : desiredCalendar.status;

  const taskCheck =
    desiredTask.status === 'READY'
      ? (
          actualName === tmv3_clean_(desiredTask.value)
            ? 'MATCH'
            : 'MISMATCH'
        )
      : desiredTask.status;

  return {
    openTask:openTask,
    jev:'JEV NOT USED — DETERMINISTIC',
    desiredCalendarTitle:desiredCalendar.value || '',
    actualCalendarTitle:actualCalendarTitle,
    calendarTitleStatus:desiredCalendar.status || '',
    calendarTitleCheck:calendarCheck,
    calendarTitleBlocker:
      desiredCalendar.status === 'BLOCKED'
        ? (desiredCalendar.reason || '')
        : '',
    desiredTaskName:desiredTask.value || '',
    actualTaskName:actualName,
    taskNameStatus:desiredTask.status || '',
    taskNameCheck:taskCheck,
    taskNameBlocker:
      desiredTask.status === 'BLOCKED'
        ? (desiredTask.reason || '')
        : '',
    calendarActionRequired:
      openTask &&
      desiredCalendar.status === 'READY' &&
      calendarCheck === 'MISMATCH',
    taskActionRequired:
      openTask &&
      desiredTask.status === 'READY' &&
      taskCheck === 'MISMATCH'
  };
}

function tmv3_step7ExistingTaskPlan_(
  record,
  disposition,
  sourceTask,
  sourceTaskIds,
  runtime
) {
  const taskId = tmv3_clean_(sourceTask && sourceTask['Task ID']);
  const fresh = tmv3_step7FreshTask_(taskId, runtime);
  const historical = disposition === 'FULFILLED_NO_RECREATE';

  const expected = tmv3_step7Expected_(
    record,
    runtime,
    false
  );

  if (!fresh.ok) {
    return tmv3_step7PlanRow_({
      record: record,
      disposition: disposition,
      taskId: taskId,
      taskStatus: tmv3_clean_(sourceTask && sourceTask['Status']),
      expected: expected,
      checks: tmv3_step7BlankChecks_('READ_FAILED'),
      plan: historical
        ? 'VERIFY_HISTORY_READ_FAILED'
        : 'DIRECT_TASK_READ_FAILED_NO_MUTATION',
      blocker: fresh.error,
      evidence: tmv3_step7Evidence_(record, [
        'FRESH_TASK_GET_FAILED'
      ]),
      sourceTaskIds: sourceTaskIds,
      readStatus: 'READ_FAILED'
    });
  }

  const actual = fresh.task || {};
  const requestedBy = fresh.requestedBy || { id:'', type:'' };
  const desiredAssignment = tmv3_desiredAssignment_(record);
  const actualAssignments = fresh.assignments || [];

  const customerCheck = tmv3_step7IdCheck_(
    expected.customerId,
    actual['Customer ID']
  );

  let orderCheck;
  if (record.vertical === 'PreInspection') {
    orderCheck = tmv3_clean_(actual['Order ID'])
      ? 'FORBIDDEN_PRESENT'
      : 'MATCH_NOT_ATTACHED';
  } else {
    orderCheck = tmv3_step7IdCheck_(
      expected.orderId,
      actual['Order ID']
    );
  }

  const locationCheck =
    expected.locationStatus === 'CREATE_REQUIRED' &&
    !expected.locationId
      ? (
          tmv3_clean_(actual['Location ID'])
            ? 'EXPECTED_LOCATION_UNRESOLVED'
            : 'CREATE_REQUIRED'
        )
      : tmv3_step7IdCheck_(
          expected.locationId,
          actual['Location ID']
        );

  const requestedCheck = tmv3_step7RequestedByCheck_(
    expected,
    requestedBy
  );

  // Fresh Contact ownership is only needed when Step 7 would propose a
  // Requested By mutation. Matching Requested By relationships do not consume
  // an extra Contact API call; Step 8 must re-verify ownership immediately
  // before any future write.
  if (
    expected.requestedByType === 'contact' &&
    expected.requestedById
  ) {
    if (requestedCheck === 'MATCH') {
      expected.contactOwnership = 'NOT_NEEDED_REQUESTED_BY_MATCH';
    } else if (!historical) {
      const ownership = tmv3_step7ContactOwnership_(
        expected.requestedById,
        expected.customerId,
        runtime
      );
      expected.contactOwnership = ownership.status;
      if (ownership.error) {
        expected.requestedByError = ownership.error;
      }
    }
  }

  const scheduleCheck = tmv3_taskScheduleCheck_(
    record.vertical,
    record.start,
    record.end,
    taskId,
    actual
  );
  const startCheck = scheduleCheck.startCheck;
  const endCheck = scheduleCheck.endCheck;

  const assignmentCheck = tmv3_step7AssignmentCheck_(
    record,
    desiredAssignment,
    actualAssignments
  );

  const titleState = tmv3_step7TitleState_(
    record,
    tmv3_clean_(actual['Name']),
    tmv3_clean_(actual['Status']),
    historical
  );

  const checks = {
    customer: customerCheck,
    order: orderCheck,
    location: locationCheck,
    requestedBy: requestedCheck,
    start: startCheck,
    end: endCheck,
    assignment: assignmentCheck.status
  };

  let action;
  if (
    checks.start === 'CANONICAL_READ_FAILED' ||
    checks.end === 'CANONICAL_READ_FAILED'
  ) {
    action = {
      plan: 'REVIEW_DATE_CANONICAL_READ_FAILED',
      blocker:
        'v2 Task schedule disagrees with Calendar and canonical v1 DesiredStartDate/DesiredEndDate could not be verified. Date mutation is blocked.'
    };
  } else {
    action = tmv3_step7ExistingAction_(
      record,
      disposition,
      expected,
      actual,
      checks,
      assignmentCheck,
      titleState,
      historical
    );
  }

  return tmv3_step7PlanRow_({
    record: record,
    disposition: disposition,
    taskId: taskId,
    taskStatus: tmv3_clean_(actual['Status']),
    expected: expected,
    actual: {
      customerId: tmv3_clean_(actual['Customer ID']),
      orderId: tmv3_clean_(actual['Order ID']),
      locationId: tmv3_clean_(actual['Location ID']),
      requestedById: tmv3_clean_(requestedBy.id),
      requestedByType: tmv3_clean_(requestedBy.type),
      start: tmv3_clean_(scheduleCheck.start),
      due: tmv3_clean_(scheduleCheck.due),
      assignment: tmv3_step7AssignmentText_(actualAssignments),
      taskName:tmv3_clean_(actual['Name']),
      scheduleSource: scheduleCheck.source || 'V2_TASK_MODEL'
    },
    titleState:titleState,
    desiredAssignment: tmv3_step7DesiredAssignmentText_(
      desiredAssignment
    ),
    desiredAssignmentObject: desiredAssignment,
    titleState:titleState,
    checks: checks,
    actions: action.actions || [],
    plan: action.plan,
    blocker: action.blocker,
    evidence: tmv3_step7Evidence_(record, [
      'FRESH_TASK_GET',
      'TASK_ID_' + taskId,
      'TASK_SCHEDULE_' + (scheduleCheck.source || 'V2_TASK_MODEL')
    ]),
    sourceTaskIds: sourceTaskIds,
    readStatus: 'FRESH_TASK_GET'
  });
}

function tmv3_step7ExistingAction_(
  record,
  disposition,
  expected,
  actual,
  checks,
  assignmentCheck,
  titleState,
  historical
) {
  function blocked(plan, blocker) {
    return { plan:plan, blocker:blocker, actions:[] };
  }

  if (historical) {
    return blocked(
      'VERIFY_ONLY_HISTORY_NO_WRITE',
      'Completed Task covers this appointment. Historical relationship drift is evidence only.'
    );
  }

  if (checks.customer !== 'MATCH') {
    return blocked(
      'REVIEW_CUSTOMER_RELATIONSHIP_CONFLICT',
      'Fresh Task Customer must exactly match the Step 4 verified Customer before any mutation.'
    );
  }

  if (
    record.vertical === 'PreInspection' &&
    checks.order === 'FORBIDDEN_PRESENT'
  ) {
    return blocked(
      'REVIEW_PREINSPECTION_ORDER_ATTACHED',
      'PreInspection Task is attached to a Sales Order; automatic detachment is not authorized.'
    );
  }

  if (
    checks.order === 'MISMATCH' ||
    checks.location === 'MISMATCH' ||
    checks.location === 'EXPECTED_LOCATION_UNRESOLVED'
  ) {
    return blocked(
      'REVIEW_RELATIONSHIP_CONFLICT',
      'Fresh Task relationship conflicts with the verified appointment relationship. Do not overwrite automatically.'
    );
  }

  if (expected.requestedByError) {
    return blocked('REVIEW_REQUESTED_BY_EVIDENCE', expected.requestedByError);
  }

  if (
    expected.requestedByType === 'contact' &&
    checks.requestedBy !== 'MATCH' &&
    expected.contactOwnership !== 'VERIFIED'
  ) {
    return blocked(
      'REVIEW_CONTACT_OWNERSHIP',
      'Expected Contact ownership is not freshly verified for the resolved Customer.'
    );
  }

  if (assignmentCheck && assignmentCheck.blocker) {
    return blocked('REVIEW_ASSIGNMENT_UNRESOLVED', assignmentCheck.blocker);
  }

  if (
    titleState &&
    titleState.taskNameStatus === 'BLOCKED'
  ) {
    return blocked(
      'REVIEW_TASK_NAME_EVIDENCE',
      titleState.taskNameBlocker || 'Canonical Task Name evidence is incomplete.'
    );
  }

  if (
    titleState &&
    titleState.calendarTitleStatus === 'BLOCKED'
  ) {
    return blocked(
      'REVIEW_CALENDAR_TITLE_EVIDENCE',
      titleState.calendarTitleBlocker || 'Canonical Calendar title evidence is incomplete.'
    );
  }

  const changes = [];
  const actions = [];

  if (checks.location === 'MISSING') {
    changes.push('LOCATION');
    actions.push(TMV3_STEP7_ACTION.PATCH_LOCATION);
  }

  if (checks.location === 'CREATE_REQUIRED') {
    changes.push('CREATE_LOCATION_THEN_LOCATION');
    actions.push(TMV3_STEP7_ACTION.CREATE_LOCATION);
    actions.push(TMV3_STEP7_ACTION.PATCH_LOCATION);
  }

  if (checks.order === 'MISSING') {
    changes.push('ORDER');
    actions.push(TMV3_STEP7_ACTION.PATCH_ORDER);
  }

  if (expected.requestedById && checks.requestedBy !== 'MATCH') {
    changes.push('REQUESTED_BY');
    actions.push(TMV3_STEP7_ACTION.PATCH_REQUESTED_BY);
  }

  if (checks.start !== 'MATCH' || checks.end !== 'MATCH') {
    changes.push('DATES');
    actions.push(TMV3_STEP7_ACTION.PATCH_DATES);
  }

  if (
    assignmentCheck.status !== 'MATCH' &&
    assignmentCheck.status !== 'N/A'
  ) {
    changes.push('ASSIGNMENTS');
    actions.push(TMV3_STEP7_ACTION.PATCH_ASSIGNMENTS);
  }

  if (titleState && titleState.calendarActionRequired) {
    changes.push('CALENDAR_TITLE');
    actions.push(TMV3_STEP7_ACTION.PATCH_CALENDAR_TITLE);
  }

  if (titleState && titleState.taskActionRequired) {
    changes.push('TASK_NAME');
    actions.push(TMV3_STEP7_ACTION.PATCH_TASK_NAME);
  }

  actions.push(TMV3_STEP7_ACTION.VERIFY_CALENDAR_LINKS);

  return {
    plan:changes.length ? 'PATCH_' + changes.join('_AND_') : 'NO_CHANGE',
    blocker:'',
    actions:tmv3_unique_(actions)
  };
}

function tmv3_step7CreatePlan_(
  record,
  disposition,
  sourceTaskIds,
  runtime
) {
  const expected = tmv3_step7Expected_(record, runtime, true);
  const desiredAssignment = tmv3_desiredAssignment_(record);

  const desiredTask = tmv3_desiredTaskName_(
    record,
    { sourceTaskTitle:'' }
  );
  const desiredCalendar = tmv3_desiredCalendarTitle_(record);
  const calendarTitleRequired =
    record.vertical === 'Install' ||
    record.vertical === 'PreInspection';

  const titleState = {
    openTask:false,
    jev:'JEV NOT USED — DETERMINISTIC',
    desiredCalendarTitle:desiredCalendar.value || '',
    actualCalendarTitle:tmv3_clean_(record && record.title),
    calendarTitleStatus:desiredCalendar.status || '',
    calendarTitleCheck:
      desiredCalendar.status === 'READY'
        ? (
            tmv3_clean_(record && record.title) ===
            tmv3_clean_(desiredCalendar.value)
              ? 'MATCH'
              : 'MISMATCH'
          )
        : desiredCalendar.status,
    calendarTitleBlocker:
      desiredCalendar.status === 'BLOCKED'
        ? (desiredCalendar.reason || '')
        : '',
    desiredTaskName:desiredTask.value || '',
    actualTaskName:'',
    taskNameStatus:desiredTask.status || '',
    taskNameCheck:'CREATE_CANONICAL',
    taskNameBlocker:
      desiredTask.status === 'BLOCKED'
        ? (desiredTask.reason || '')
        : '',
    calendarActionRequired:
      desiredCalendar.status === 'READY' &&
      tmv3_clean_(record && record.title) !==
        tmv3_clean_(desiredCalendar.value),
    taskActionRequired:false
  };

  let plan;
  let blocker = '';
  let actions = [];

  if (!expected.customerId) {
    plan = 'REVIEW_CREATE_CUSTOMER_UNRESOLVED';
    blocker = 'Task CREATE requires a verified Customer.';
  } else if (
    TMV3.VERTICALS[record.vertical].orderRequired &&
    !expected.orderId
  ) {
    plan = 'REVIEW_CREATE_ORDER_UNRESOLVED';
    blocker = 'Task CREATE requires the verified Order / Work Order relationship.';
  } else if (expected.requestedByError) {
    plan = 'REVIEW_CREATE_REQUESTED_BY_UNRESOLVED';
    blocker = expected.requestedByError;
  } else if (
    expected.requestedByType === 'contact' &&
    expected.contactOwnership !== 'VERIFIED'
  ) {
    plan = 'REVIEW_CREATE_CONTACT_OWNERSHIP';
    blocker = 'Expected Contact ownership is not freshly verified for the resolved Customer.';
  } else if (desiredTask.status !== 'READY') {
    plan = 'REVIEW_CREATE_TASK_NAME_UNRESOLVED';
    blocker =
      desiredTask.reason ||
      'Canonical Task Name cannot be built from verified evidence.';
  } else if (
    calendarTitleRequired &&
    desiredCalendar.status !== 'READY'
  ) {
    plan = 'REVIEW_CREATE_CALENDAR_TITLE_UNRESOLVED';
    blocker =
      desiredCalendar.reason ||
      'Canonical Calendar title cannot be built from verified evidence.';
  } else {
    const createVerb =
      disposition === 'RECREATE_TASK' ? 'RECREATE_TASK' : 'CREATE_TASK';
    const createAction =
      disposition === 'RECREATE_TASK'
        ? TMV3_STEP7_ACTION.RECREATE_TASK
        : TMV3_STEP7_ACTION.CREATE_TASK;

    if (
      expected.locationStatus === 'CREATE_REQUIRED' &&
      !expected.locationId
    ) {
      actions.push(TMV3_STEP7_ACTION.CREATE_LOCATION);
      plan =
        record.vertical === 'PreInspection'
          ? (
              'CREATE_LOCATION_THEN_' +
              createVerb +
              '__TYPE105__BLANK_DESCRIPTION__NO_SO__POOL8'
            )
          : 'CREATE_LOCATION_THEN_' + createVerb;
    } else if (record.vertical === 'PreInspection') {
      plan =
        createVerb +
        '__TYPE105__BLANK_DESCRIPTION__NO_SO__POOL8';
    } else {
      plan = createVerb;
    }

    actions.push(createAction);

    if (
      (desiredAssignment.employeeIds || []).length ||
      (desiredAssignment.poolIds || []).length
    ) {
      actions.push(TMV3_STEP7_ACTION.PATCH_ASSIGNMENTS);
    }

    actions.push(TMV3_STEP7_ACTION.VERIFY_CALENDAR_LINKS);

    if (titleState.calendarActionRequired) {
      actions.push(TMV3_STEP7_ACTION.PATCH_CALENDAR_TITLE);
      plan += '__CALENDAR_TITLE';
    }
  }

  return tmv3_step7PlanRow_({
    record:record,
    disposition:disposition,
    taskId:'',
    taskStatus:'NOT_CREATED',
    expected:expected,
    checks:{
      customer:expected.customerId ? 'VERIFIED_EXPECTED' : 'MISSING',
      order:
        record.vertical === 'PreInspection'
          ? 'NOT_APPLICABLE'
          : (expected.orderId ? 'VERIFIED_EXPECTED' : 'MISSING'),
      location:
        expected.locationId
          ? 'VERIFIED_EXPECTED'
          : (
              expected.locationStatus === 'CREATE_REQUIRED'
                ? 'CREATE_REQUIRED'
                : 'MISSING'
            ),
      requestedBy:
        expected.requestedByError
          ? 'UNRESOLVED'
          : (expected.requestedById ? 'VERIFIED_EXPECTED' : 'N/A'),
      start:'VERIFIED_EXPECTED',
      end:'VERIFIED_EXPECTED',
      assignment:'VERIFIED_EXPECTED'
    },
    desiredAssignment:tmv3_step7DesiredAssignmentText_(desiredAssignment),
    desiredAssignmentObject:desiredAssignment,
    titleState:titleState,
    actions:actions,
    plan:plan,
    blocker:blocker,
    evidence:tmv3_step7Evidence_(record, ['CREATE_PLAN_ONLY','NO_TASK_WRITE']),
    sourceTaskIds:sourceTaskIds,
    readStatus:'NO_TASK_READ_REQUIRED'
  });
}

function tmv3_step7NoTaskPlan_(
  record,
  disposition,
  plan,
  blocker,
  sourceTaskIds
) {
  const expected = {
    customerId: tmv3_clean_(
      record.step4 &&
      record.step4.customer &&
      record.step4.customer['Customer ID']
    ),
    orderId: tmv3_clean_(
      record.step3 &&
      record.step3.anchor &&
      record.step3.anchor.orderId
    ),
    locationId: tmv3_clean_(
      record.step4 &&
      record.step4.location &&
      record.step4.location['Location ID']
    ),
    requestedById: '',
    requestedByType: '',
    contactOwnership: 'NOT_CHECKED',
    evidence: []
  };

  return tmv3_step7PlanRow_({
    record: record,
    disposition: disposition,
    taskId: '',
    taskStatus: '',
    expected: expected,
    checks: tmv3_step7BlankChecks_('NOT_RUN'),
    plan: plan,
    blocker: blocker,
    evidence: tmv3_step7Evidence_(record, [
      'NO_MUTATION_AUTHORIZED'
    ]),
    sourceTaskIds: sourceTaskIds,
    readStatus: 'NOT_RUN'
  });
}

function tmv3_step7PlanRow_(input) {
  input = input || {};
  const record = input.record || {};
  const expected = input.expected || {};
  const actual = input.actual || {};
  const checks = input.checks || tmv3_step7BlankChecks_('');
  const now = tmv3_now_();
  const sourceTaskIdList = (input.sourceTaskIds || [])
    .map(tmv3_clean_)
    .filter(Boolean);
  const desired =
    input.desiredAssignmentObject || { employeeIds:[], poolIds:[] };
  const actions = tmv3_unique_(
    (input.actions || []).map(tmv3_clean_).filter(Boolean)
  );
  const inputFingerprint =
    tmv3_clean_(record.fingerprint) ||
    tmv3_hash_([
      record.vertical || '',
      record.eventId || '',
      record.calendarId || '',
      record.start ? tmv3_iso_(record.start) : '',
      record.end ? tmv3_iso_(record.end) : '',
      record.title || '',
      record.location || '',
      record.description || '',
      record.organizer || ''
    ].join('|'));

  return {
    vertical:record.vertical || '',
    eventId:record.eventId || '',
    calendarId:record.calendarId || '',
    logicalKey:record.logicalKey || (
      tmv3_clean_(record.vertical) + '|' + tmv3_clean_(record.eventId)
    ),
    event:record.title || '',
    disposition:input.disposition || '',
    taskId:input.taskId || '',
    taskStatus:input.taskStatus || '',
    expectedCustomerId:expected.customerId || '',
    actualCustomerId:actual.customerId || '',
    customerCheck:checks.customer || '',
    expectedOrderId:expected.orderId || '',
    actualOrderId:actual.orderId || '',
    orderCheck:checks.order || '',
    expectedLocationId:expected.locationId || '',
    locationStatus:expected.locationStatus || '',
    expectedLocationAddress:expected.locationAddress || '',
    actualLocationId:actual.locationId || '',
    locationCheck:checks.location || '',
    expectedRequestedById:expected.requestedById || '',
    expectedRequestedByType:expected.requestedByType || '',
    actualRequestedById:actual.requestedById || '',
    actualRequestedByType:actual.requestedByType || '',
    requestedByCheck:checks.requestedBy || '',
    contactOwnership:expected.contactOwnership || 'N/A',
    calendarStart:record.start ? tmv3_iso_(record.start) : '',
    expectedStart:record.start ? tmv3_iso_(record.start) : '',
    taskStart:actual.start || '',
    startCheck:checks.start || '',
    calendarEnd:record.end ? tmv3_iso_(record.end) : '',
    expectedDue:record.end ? tmv3_iso_(record.end) : '',
    taskDue:actual.due || '',
    endCheck:checks.end || '',
    scheduleSource:actual.scheduleSource || 'V2_TASK_MODEL',
    desiredAssignment:input.desiredAssignment || '',
    desiredAssignmentEmployeeIds:(desired.employeeIds || []).map(Number),
    desiredAssignmentPoolIds:(desired.poolIds || []).map(Number),
    actualAssignment:actual.assignment || '',
    assignmentCheck:checks.assignment || '',
    desiredCalendarTitle:
      (input.titleState && input.titleState.desiredCalendarTitle) || '',
    actualCalendarTitle:
      (input.titleState && input.titleState.actualCalendarTitle) ||
      tmv3_clean_(record.title),
    calendarTitleStatus:
      (input.titleState && input.titleState.calendarTitleStatus) || '',
    calendarTitleCheck:
      (input.titleState && input.titleState.calendarTitleCheck) || '',
    calendarTitleBlocker:
      (input.titleState && input.titleState.calendarTitleBlocker) || '',
    desiredTaskName:
      (input.titleState && input.titleState.desiredTaskName) || '',
    actualTaskName:
      (input.titleState && input.titleState.actualTaskName) ||
      (actual.taskName || ''),
    taskNameStatus:
      (input.titleState && input.titleState.taskNameStatus) || '',
    taskNameCheck:
      (input.titleState && input.titleState.taskNameCheck) || '',
    taskNameBlocker:
      (input.titleState && input.titleState.taskNameBlocker) || '',
    titleJev:
      (input.titleState && input.titleState.jev) ||
      'JEV NOT USED — DETERMINISTIC',
    plan:input.plan || '',
    actions:actions,
    writeGate:
      TMV3.MODE === 'SHADOW_READ_ONLY'
        ? 'SHADOW_ONLY__NO_WRITES'
        : record.vertical === 'PreInspection' && TMV3.MODE === 'PRODUCTION_WRITE' &&
          TMV3.OPERATIONS.preInspectionAutomationEnabled === true
          ? (input.blocker ? 'PREINSPECTION_AUTO_HELD' : 'PREINSPECTION_AUTO_ENABLED')
          : 'EXECUTION_CONFIG_GATED',
    blocker:input.blocker || '',
    evidence:input.evidence || '',
    sourceTaskIds:sourceTaskIdList.join(','),
    sourceTaskIdList:sourceTaskIdList,
    engineVersion:TMV3.VERSION,
    plannedAt:now,
    readStatus:input.readStatus || '',
    freshTaskReadStatus:input.readStatus || '',
    inputFingerprint:inputFingerprint,
    calendarUpdatedAt:tmv3_clean_(record.calendarUpdatedAt)
  };
}

function tmv3_step7BlankChecks_(value) {
  return {
    customer: value || '',
    order: value || '',
    location: value || '',
    requestedBy: value || '',
    start: value || '',
    end: value || '',
    assignment: value || ''
  };
}

function tmv3_step7IdCheck_(expected, actual) {
  const e = tmv3_clean_(expected);
  const a = tmv3_clean_(actual);

  if (!e && !a) return 'N/A';
  if (!e && a) return 'SOURCE_ONLY';
  if (e && !a) return 'MISSING';
  return e === a ? 'MATCH' : 'MISMATCH';
}

function tmv3_step7RequestedByCheck_(expected, actual) {
  const expectedId = tmv3_clean_(expected.requestedById);
  const expectedType = tmv3_norm_(expected.requestedByType);
  const actualId = tmv3_clean_(actual && actual.id);
  const actualType = tmv3_norm_(actual && actual.type);

  if (!expectedId && !actualId) return 'N/A';
  if (!expectedId && actualId) return 'SOURCE_ONLY';
  if (expectedId && !actualId) return 'MISSING';
  if (expectedId !== actualId) return 'MISMATCH';
  if (expectedType && !actualType) return 'TYPE_UNKNOWN';
  if (expectedType && expectedType !== actualType) return 'TYPE_MISMATCH';
  return 'MATCH';
}

function tmv3_step7DateCheck_(expectedDate, actualValue) {
  if (!expectedDate && !actualValue) return 'N/A';
  if (expectedDate && !actualValue) return 'MISSING';
  if (!expectedDate && actualValue) return 'SOURCE_ONLY';

  return tmv3_sameMinute_(expectedDate, actualValue)
    ? 'MATCH'
    : 'MISMATCH';
}

function tmv3_step7DesiredAssignmentText_(desired) {
  desired = desired || {};

  const parts = [];

  (desired.employeeIds || []).forEach(function(id) {
    parts.push('employee:' + Number(id));
  });

  (desired.poolIds || []).forEach(function(id) {
    parts.push('pool:' + Number(id));
  });

  return parts.sort().join(',');
}

function tmv3_step7AssignmentText_(assignments) {
  return (assignments || [])
    .map(function(item) {
      return (
        tmv3_clean_(item.type) +
        ':' +
        Number(item.id || 0)
      );
    })
    .filter(Boolean)
    .sort()
    .join(',');
}

function tmv3_step7AssignmentCheck_(
  record,
  desired,
  actualAssignments
) {
  desired = desired || {
    employeeIds: [],
    poolIds: []
  };

  const actual = (actualAssignments || []).map(function(item) {
    return tmv3_clean_(item.type) + '|' + Number(item.id || 0);
  });

  const desiredEmployeeIds = (desired.employeeIds || []).map(Number);
  const desiredPoolIds = (desired.poolIds || []).map(Number);

  // Legacy Delivery / Service parity:
  // existing employee assignments are preserved. The only managed
  // conflicting assignment is Pool 4 ("To Be Assigned").
  if (
    (record.vertical === 'Delivery' || record.vertical === 'Service') &&
    actual.indexOf('pool|4') !== -1 &&
    !desiredEmployeeIds.length
  ) {
    return {
      status: 'BLOCKED',
      missing: [],
      conflicts: ['pool|4'],
      blocker:
        'To Be Assigned pool exists but no valid intended employee was resolved from the legacy assignment source.'
    };
  }

  const required = [];

  desiredEmployeeIds.forEach(function(id) {
    required.push('employee|' + id);
  });

  desiredPoolIds.forEach(function(id) {
    required.push('pool|' + id);
  });

  if (!required.length) {
    return {
      status: 'N/A',
      missing: [],
      conflicts: [],
      blocker: ''
    };
  }

  const missing = required.filter(function(key) {
    return actual.indexOf(key) === -1;
  });

  const conflicts = [];

  if (
    (record.vertical === 'Delivery' || record.vertical === 'Service') &&
    desiredEmployeeIds.length &&
    actual.indexOf('pool|4') !== -1
  ) {
    conflicts.push('pool|4');
  }

  if (
    record.vertical === 'PreInspection' &&
    actual.indexOf('pool|4') !== -1
  ) {
    conflicts.push('pool|4');
  }

  return {
    status:
      missing.length || conflicts.length
        ? 'MISMATCH'
        : 'MATCH',
    missing: missing,
    conflicts: conflicts,
    blocker: ''
  };
}

function tmv3_step7Evidence_(record, extra) {
  const evidence = [];

  if (record.step6) {
    if (record.step6.code) {
      evidence.push('STEP6_' + record.step6.code);
    }
    if (record.step6.reason) {
      evidence.push(tmv3_clean_(record.step6.reason));
    }
  }

  ((record.step4 && record.step4.evidence) || [])
    .forEach(function(item) {
      evidence.push(tmv3_clean_(item));
    });

  (extra || []).forEach(function(item) {
    evidence.push(tmv3_clean_(item));
  });

  return tmv3_unique_(
    evidence.filter(Boolean)
  ).join(' | ');
}

function tmv3_step7EnsureSheet_() {
  const ss = tmv3_ss_();
  let sh = ss.getSheetByName(TMV3.SHEETS.RECONCILE);

  if (!sh) {
    sh = ss.insertSheet(TMV3.SHEETS.RECONCILE);
  }

  const width = TMV3_STEP7_HEADERS.length;

  if (sh.getMaxColumns() < width) {
    sh.insertColumnsAfter(
      sh.getMaxColumns(),
      width - sh.getMaxColumns()
    );
  }

  if (sh.getMaxRows() < 2) {
    sh.insertRowsAfter(sh.getMaxRows(), 2 - sh.getMaxRows());
  }

  return sh;
}

function tmv3_step7WritePlans_(plans, verticalFilter, batch) {
  const sh = tmv3_step7EnsureSheet_();
  const filter = tmv3_clean_(verticalFilter);

  const newRows = (plans || []).map(function(plan) {
    return [
      plan.vertical,
      plan.eventId,
      plan.logicalKey,
      plan.event,
      plan.disposition,
      plan.taskId,
      plan.taskStatus,
      plan.expectedCustomerId,
      plan.actualCustomerId,
      plan.customerCheck,
      plan.expectedOrderId,
      plan.actualOrderId,
      plan.orderCheck,
      plan.expectedLocationId,
      plan.actualLocationId,
      plan.locationCheck,
      plan.expectedRequestedById,
      plan.expectedRequestedByType,
      plan.actualRequestedById,
      plan.actualRequestedByType,
      plan.requestedByCheck,
      plan.contactOwnership,
      plan.calendarStart,
      plan.taskStart,
      plan.startCheck,
      plan.calendarEnd,
      plan.taskDue,
      plan.endCheck,
      plan.desiredAssignment,
      plan.actualAssignment,
      plan.assignmentCheck,
      plan.plan,
      plan.writeGate,
      plan.blocker,
      plan.evidence,
      plan.sourceTaskIds,
      plan.engineVersion,
      plan.plannedAt,
      plan.readStatus
    ];
  });

  let rows = newRows.slice();

  if (filter && sh.getLastRow() > 1) {
    const existing = sh
      .getRange(1,1,sh.getLastRow(),TMV3_STEP7_HEADERS.length)
      .getValues();

    const headerMatches =
      existing.length &&
      String(existing[0][0] || '') === TMV3_STEP7_HEADERS[0] &&
      existing[0].length === TMV3_STEP7_HEADERS.length;

    if (headerMatches) {
      const replacementKeys = {};
      newRows.forEach(function(row) {
        replacementKeys[tmv3_clean_(row[2])] = true;
      });

      const preserved = existing.slice(1).filter(function(row) {
        const rowVertical = tmv3_clean_(row[0]);
        if (!rowVertical) return false;
        if (rowVertical !== filter) return true;
        if (!(batch && batch.enabled)) return false;
        return !replacementKeys[tmv3_clean_(row[2])];
      });

      rows = preserved.concat(newRows);
    }
  }

  rows.sort(function(a, b) {
    const verticalCompare =
      tmv3_clean_(a[0]).localeCompare(tmv3_clean_(b[0]));
    if (verticalCompare !== 0) return verticalCompare;

    const startCompare =
      tmv3_clean_(a[22]).localeCompare(tmv3_clean_(b[22]));
    if (startCompare !== 0) return startCompare;

    return tmv3_clean_(a[3]).localeCompare(tmv3_clean_(b[3]));
  });

  tmv3_replaceRows_(
    TMV3.SHEETS.RECONCILE,
    TMV3_STEP7_HEADERS.slice(),
    rows
  );

  try {
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,TMV3_STEP7_HEADERS.length)
      .setFontWeight('bold')
      .setWrap(true);

    sh.getRange(1,1,Math.max(1,rows.length + 1),TMV3_STEP7_HEADERS.length)
      .setWrap(true);

    sh.setColumnWidth(4, 320);
    sh.setColumnWidth(32, 280);
    sh.setColumnWidth(34, 320);
    sh.setColumnWidth(35, 360);

    if (!sh.isSheetHidden()) {
      sh.hideSheet();
    }
  } catch (ignored) {}

  return {
    sheet: TMV3.SHEETS.RECONCILE,
    verticalFilter: filter || 'ALL',
    batch: batch || { enabled:false },
    rowsWritten: newRows.length,
    totalRows: rows.length,
    preservedOtherVerticalRows: rows.length - newRows.length,
    hidden: true
  };
}

function tmv3_step7OperatorNextAction_(plans) {
  const list = (plans || []).map(function(plan) {
    return tmv3_clean_(plan);
  }).filter(Boolean);

  if (!list.length) return 'No Step 7 reconciliation result published yet.';
  if (list.every(function(plan) { return plan === 'NO_CHANGE'; })) {
    return 'Verified by Step 7 — no change required.';
  }
  if (list.some(function(plan) { return /^REVIEW_|READ_FAILED|DIRECT_TASK_READ_FAILED/.test(plan); })) {
    return 'Step 7 requires review before any write.';
  }
  if (list.some(function(plan) { return plan.indexOf('CREATE_LOCATION_THEN_') === 0; })) {
    return 'Step 7 requires customer-owned Location creation before Task creation.';
  }
  if (list.some(function(plan) { return plan.indexOf('RECREATE_TASK') !== -1; })) {
    return 'Step 7 verified a guarded Task recreation candidate.';
  }
  if (list.some(function(plan) { return plan.indexOf('CREATE_TASK') !== -1; })) {
    return 'Step 7 verified a guarded Task creation candidate.';
  }
  if (list.some(function(plan) { return plan.indexOf('PATCH_') === 0; })) {
    return 'Step 7 identified a guarded existing-Task correction.';
  }
  if (list.some(function(plan) { return plan.indexOf('VERIFY_ONLY') === 0; })) {
    return 'Step 7 verified historical coverage — no automatic write.';
  }
  return 'Step 7 reconciliation published. Review the plan shown here.';
}

function tmv3_publishLatestStep7ToVisibleSheets() {
  const reconcileRows = tmv3_rows_(TMV3.SHEETS.RECONCILE);
  const byEvent = {};

  reconcileRows.forEach(function(row) {
    const vertical = tmv3_clean_(row['Vertical']);
    const eventId = tmv3_clean_(row['Event ID']);
    if (!vertical || !eventId) return;

    const key = vertical + '|' + eventId;
    if (!byEvent[key]) byEvent[key] = [];
    byEvent[key].push(row);
  });

  const written = {};
  let publishedRows = 0;

  Object.keys(TMV3.VERTICALS).forEach(function(vertical) {
    const sheetName = TMV3.VERTICALS[vertical].sheet;
    const sh = tmv3_sheet_(sheetName);
    const lastRow = sh.getLastRow();
    const lastColumn = sh.getLastColumn();

    if (lastRow < 2 || lastColumn < 1) {
      written[vertical] = 0;
      return;
    }

    const values = sh.getRange(1,1,lastRow,lastColumn).getValues();
    const headers = values[0].map(tmv3_clean_);
    const eventIx = headers.indexOf('Event ID');
    const mappingIx = headers.indexOf('Mapping Status');
    const issueIx = headers.indexOf('Issue / Next Action');

    if (eventIx < 0 || mappingIx < 0 || issueIx < 0) {
      throw new Error(
        'Operator sheet ' + sheetName +
        ' is missing Event ID / Mapping Status / Issue columns.'
      );
    }

    let changed = 0;

    for (let i = 1; i < values.length; i++) {
      const eventId = tmv3_clean_(values[i][eventIx]);
      if (!eventId) continue;

      const plans = byEvent[vertical + '|' + eventId] || [];
      if (!plans.length) continue;

      const planNames = tmv3_unique_(
        plans.map(function(plan) {
          return tmv3_clean_(plan['Relationship Plan']);
        }).filter(Boolean)
      );

      const blockers = tmv3_unique_(
        plans.map(function(plan) {
          return tmv3_clean_(plan['Blocker']);
        }).filter(Boolean)
      );

      const engineVersions = tmv3_unique_(
        plans.map(function(plan) {
          return tmv3_clean_(plan['Engine Version']);
        }).filter(Boolean)
      );

      values[i][mappingIx] =
        'STEP 7 — ' +
        (planNames.length ? planNames.join(' + ') : 'NO PLAN');

      values[i][issueIx] =
        blockers.length
          ? blockers.join(' | ')
          : tmv3_step7OperatorNextAction_(planNames);

      // Keep a visible, non-invasive proof of which engine produced this
      // projection without changing the operator column structure.
      const checklistIx = headers.indexOf('Data Checklist');
      if (checklistIx >= 0 && engineVersions.length) {
        const existing = tmv3_clean_(values[i][checklistIx])
          .replace(/\n?V3 Engine:\s*[^\n]+/gi, '')
          .trim();
        values[i][checklistIx] =
          (existing ? existing + '\n' : '') +
          'V3 Engine: ' +
          engineVersions.join(', ');
      }

      changed++;
      publishedRows++;
    }

    if (changed) {
      sh.getRange(1,1,values.length,values[0].length).setValues(values);
    }

    written[vertical] = changed;
  });

  const overview = tmv3_sheet_(TMV3.SHEETS.OVERVIEW);
  overview.getRange('B2').setValue(
    'V3 ' + TMV3.VERSION + ' — STAGE ' + tmv3_executionStage_() +
    ' — VISIBLE SHEET SYNCED'
  );
  overview.getRange('B4').setValue(
    TMV3.MODE + ' — WORKBOOK UPDATES ENABLED; BUSINESS WRITES GATED'
  );
  overview.getRange('B5').setValue(
    'Visible operator tabs are projected from the latest Step 7 reconciliation.'
  );
  overview.getRange('B6').setValue(
    'Mapping Status and Issue / Next Action now show the latest published Step 7 plan.'
  );
  overview.getRange('B14').setValue(
    'Step 7 visible-sheet publication active; guarded CREATE/RECREATE verification remains next.'
  );
  overview.getRange('B16').setValue(
    'Operator tabs reflect the latest TM Reconcile publication; TM Reconcile remains the detailed evidence source.'
  );

  const morning = tmv3_refreshMorningOps();

  tmv3_audit_(
    'SYSTEM','','','STEP7_VISIBLE_SHEET_PUBLISH','PASS',
    JSON.stringify({
      version: TMV3.VERSION,
      reconcileRows: reconcileRows.length,
      publishedRows: publishedRows,
      written: written
    })
  );

  return {
    status:'STEP7_VISIBLE_SHEETS_SYNCED',
    version:TMV3.VERSION,
    reconcileRows:reconcileRows.length,
    publishedRows:publishedRows,
    written:written,
    morningOps:morning
  };
}

function tmv3_step7Counts_(plans) {
  const counts = {
    total: 0,
    byVertical: {},
    byPlan: {},
    taskReads: 0,
    taskReadFailures: 0,
    safeMutationCandidates: 0,
    reviewPlans: 0,
    noChange: 0
  };

  (plans || []).forEach(function(plan) {
    counts.total++;

    if (!counts.byVertical[plan.vertical]) {
      counts.byVertical[plan.vertical] = 0;
    }
    counts.byVertical[plan.vertical]++;

    if (!counts.byPlan[plan.plan]) {
      counts.byPlan[plan.plan] = 0;
    }
    counts.byPlan[plan.plan]++;

    if (plan.readStatus === 'FRESH_TASK_GET') {
      counts.taskReads++;
    }
    if (plan.readStatus === 'READ_FAILED') {
      counts.taskReadFailures++;
    }

    if (/^PATCH_|^CREATE_|^RECREATE_/.test(plan.plan)) {
      counts.safeMutationCandidates++;
    }
    if (/^REVIEW_|READ_FAILED/.test(plan.plan)) {
      counts.reviewPlans++;
    }
    if (plan.plan === 'NO_CHANGE') {
      counts.noChange++;
    }
  });

  return counts;
}

function tmv3_step7Verify_(step6Records, plans) {
  const covered = {};
  const expected = {};
  let invalidWriteGate = 0;
  let missingPlan = 0;
  let readFailures = 0;
  let multiSiblingMismatch = 0;

  (plans || []).forEach(function(plan) {
    covered[plan.logicalKey] =
      (covered[plan.logicalKey] || 0) + 1;

    const expectedGate = TMV3.MODE === 'SHADOW_READ_ONLY'
      ? 'SHADOW_ONLY__NO_WRITES'
      : plan.vertical === 'PreInspection' && TMV3.MODE === 'PRODUCTION_WRITE' &&
        TMV3.OPERATIONS.preInspectionAutomationEnabled === true
        ? (plan.blocker ? 'PREINSPECTION_AUTO_HELD' : 'PREINSPECTION_AUTO_ENABLED')
        : 'EXECUTION_CONFIG_GATED';
    if (plan.writeGate !== expectedGate) {
      invalidWriteGate++;
    }

    if (!plan.plan) {
      missingPlan++;
    }

    if (plan.readStatus === 'READ_FAILED') {
      readFailures++;
    }
  });

  (step6Records || []).forEach(function(record) {
    const key = record.logicalKey || (
      tmv3_clean_(record.vertical) + '|' +
      tmv3_clean_(record.eventId)
    );
    expected[key] = true;

    const disposition = tmv3_clean_(
      record.step6 && record.step6.disposition
    );

    if (
      disposition === 'MATCH_EXISTING_MULTI'
    ) {
      const expectedSiblingRows =
        ((record.step6 && record.step6.tasks) || []).length;
      const actualSiblingRows = Number(covered[key] || 0);

      if (expectedSiblingRows !== actualSiblingRows) {
        multiSiblingMismatch++;
      }
    }
  });

  const missingKeys = Object.keys(expected).filter(function(key) {
    return !covered[key];
  });

  const introducedKeys = Object.keys(covered).filter(function(key) {
    return !expected[key];
  });

  return {
    pass:
      invalidWriteGate === 0 &&
      missingPlan === 0 &&
      readFailures === 0 &&
      multiSiblingMismatch === 0 &&
      missingKeys.length === 0 &&
      introducedKeys.length === 0,
    step6Records: (step6Records || []).length,
    step7Rows: (plans || []).length,
    missingPlan: missingPlan,
    invalidWriteGate: invalidWriteGate,
    taskReadFailures: readFailures,
    multiSiblingMismatch: multiSiblingMismatch,
    missingFromStep7: missingKeys.length,
    introducedByStep7: introducedKeys.length,
    missingKeys: missingKeys.slice(0,20),
    introducedKeys: introducedKeys.slice(0,20)
  };
}


/**
 * V3-001: Calendar-independent, read-only relationship audit.
 *
 * A Striven Service/Delivery Task can be created before scheduling. Do not
 * infer its Customer, Location or Requested By from a task title, and do not
 * generate a date/assignment plan without Calendar evidence.
 *
 * This audit makes at most one fresh Task API request; related SO identity is
 * only a candidate unless independently verified. It never writes to Striven.
 */
function tmv3_taskFirstRelationshipAudit(taskId, vertical) {
  const id = Number(taskId);
  const kind = tmv3_clean_(vertical);
  if (!Number.isSafeInteger(id) || id <= 0 ||
      (kind !== 'Service' && kind !== 'Delivery')) {
    throw new Error('V3 task-first audit requires a positive Task ID and Service or Delivery.');
  }

  const sourceRows = tmv3_rows_(TMV3.SHEETS.TASKS);
  const candidate = sourceRows.filter(function(row) {
    return Number(row['Task ID']) === id;
  });
  if (candidate.length !== 1) {
    return {
      status:'REVIEW_TASK_SOURCE_MISSING_OR_DUPLICATED', taskId:id,
      vertical:kind, candidateCount:candidate.length, writes:0
    };
  }

  const source = candidate[0];
  const sourceType = tmv3_norm_(source['Task Type']);
  if ((kind === 'Service' && sourceType.indexOf('service') < 0) ||
      (kind === 'Delivery' && sourceType.indexOf('deliver') < 0)) {
    return {status:'REVIEW_WRONG_TASK_VERTICAL',taskId:id,vertical:kind,writes:0};
  }

  const runtime = {taskById:{},contactByKey:{},organizerByEmail:{}};
  const fresh = tmv3_step7FreshTask_(String(id), runtime);
  if (!fresh.ok) {
    return {
      status:'READ_FAILED', taskId:id, vertical:kind,
      reason:fresh.error, writes:0
    };
  }

  const actual = fresh.task || {};
  const requested = fresh.requestedBy || {};
  const rawName = tmv3_clean_(actual['Name'] || source['Name']);
  const orderNumberMatches = rawName.match(/\bSO\s*#\s*(\d+)\b/i);
  const orderNumber = orderNumberMatches ? orderNumberMatches[1] : '';
  const orders = tmv3_rows_(TMV3.SHEETS.ORDERS);
  const orderCandidates = orderNumber ? orders.filter(function(order) {
    return tmv3_clean_(order['Order Number']) === orderNumber;
  }) : [];
  const orderCandidate = orderCandidates.length === 1 ? orderCandidates[0] : null;
  const currentOrderId = tmv3_clean_(actual['Order ID']);
  const proposedOrderId = orderCandidate
    ? tmv3_clean_(orderCandidate['Order ID']) : '';
  // Cached Work Order reports may omit Customer ID. A report omission is
  // unknown evidence, not proof that a successfully repaired link is wrong.
  // One targeted authoritative GET proves ownership without a full refresh.
  let safeOrderCandidate = false;
  let ownershipError = '';
  if (proposedOrderId && (!currentOrderId || currentOrderId === proposedOrderId) &&
      /^[1-9]\d*$/.test(tmv3_clean_(actual['Customer ID']))) {
    try {
      const proof = tmv3_taskFirstFreshOrderOwnershipById_(
        proposedOrderId,tmv3_clean_(actual['Customer ID']),rawName
      );
      safeOrderCandidate = proof && proof.verified === true;
    } catch(err) {
      ownershipError = String(err && err.message || err);
    }
  }

  return {
    status:'AUDIT_ONLY_NO_MUTATION',
    vertical:kind,taskId:id,taskStatus:tmv3_clean_(actual['Status']),
    actual:{
      customerId:tmv3_clean_(actual['Customer ID']),
      orderId:currentOrderId,
      locationId:tmv3_clean_(actual['Location ID']),
      requestedById:tmv3_clean_(requested.id),
      requestedByType:tmv3_clean_(requested.type)
    },
    orderCandidate:{
      number:orderNumber,
      id:proposedOrderId,
      count:orderCandidates.length,
      customerOwnershipVerified:safeOrderCandidate,
      ownershipEvidence:safeOrderCandidate ? 'FRESH_STRIVEN_ORDER_GET' : 'UNPROVEN',
      ownershipError:ownershipError
    },
    requiredNextAction:safeOrderCandidate
      ? 'FRESH_VERIFY_ORDER_LOCATION_AND_CONTACT_BEFORE_GUARDED_PATCH'
      : 'REVIEW_ORDER_CUSTOMER_OWNERSHIP_NO_AUTOMATIC_PATCH',
    scheduling:'AWAIT_CALENDAR_EVENT',
    writes:0
  };
}




// Original: 50_Verticals
function tmv3_desiredAssignment_(eventRecord) {
  const cfg = TMV3.VERTICALS[eventRecord.vertical];

  if (eventRecord.vertical === 'Service') {
    return {
      employeeIds:
        eventRecord.technicianEmployeeId
          ? [Number(eventRecord.technicianEmployeeId)]
          : [],
      poolIds: []
    };
  }

  if (eventRecord.vertical === 'PreInspection') {
    // A targeted operator instruction may add an employee for one verified
    // event. It does not turn every organizer into an assignee.
    const override = tmv3_preInspectionAssignmentOverride_(eventRecord);
    const owner = tmv3_piVerifiedOwnerEmployee_(eventRecord);
    return {
      employeeIds: override.length ? override : (owner.status==='MATCHED' ? [owner.employeeId] : []),
      poolIds: [Number(cfg.defaultPoolId)]
    };
  }

  // Install / Delivery parity rule:
  // assignment names must be explicit in the Calendar event TITLE.
  // Description/notes are narrative context and must never create an
  // assignment instruction (for example, "Let SF know the results").
  const title = tmv3_norm_(eventRecord.assignmentSourceTitle || eventRecord.title);

  const ids = [];

  (cfg.assignees || []).forEach(function(p) {
    if (p.ignoreAssignment) return;

    if (
      (p.patterns || []).some(function(q) {
        const pattern = tmv3_norm_(q);
        return pattern &&
          (' ' + title + ' ').indexOf(' ' + pattern + ' ') !== -1;
      })
    ) {
      ids.push(Number(p.employeeId));
    }
  });

  return {
    employeeIds: tmv3_unique_(ids),
    poolIds: []
  };
}

// Overrides are private Script Properties, keyed to the immutable event ID.
// Mismatched or incomplete evidence blocks the plan rather than assigning a
// different appointment. No property means the normal Pool 8 rule applies.


// Explicitly configure one appointment after rereading its current Calendar
// identity. No task or Calendar event is changed by this operation.



/************************************************************
 * TM V3 — CANONICAL TITLE / TASK NAME POLICY
 *
 * Deterministic only. JEV is intentionally not used here.
 * Identity, ownership and canonical Task selection must already
 * be verified by the existing Step 3-7 pipeline.
 ************************************************************/

function tmv3_titlePhoneDisplay_(value) {
  const phone = tmv3_phone10_(value);
  if (!phone || phone.length !== 10) return '';
  return '(' + phone.slice(0,3) + ') ' + phone.slice(3,6) + '-' + phone.slice(6);
}


function tmv3_taskNameNormalizationStatusAllowed_(status) {
  return tmv3_clean_(status).toUpperCase() === 'OPEN';
}

function tmv3_titleCityFromAddress_(address) {
  const parts = tmv3_clean_(address)
    .split(',')
    .map(tmv3_clean_)
    .filter(Boolean);

  if (parts.length < 2) return '';

  let provinceIndex = -1;

  for (let i = 1; i < parts.length; i++) {
    if (
      /^(?:ON|QC|BC|AB|MB|SK|NS|NB|NL|PE|NT|NU|YT)\b/i.test(parts[i]) ||
      /\b[A-Z]\d[A-Z]\s?\d[A-Z]\d\b/i.test(parts[i])
    ) {
      provinceIndex = i;
      break;
    }
  }

  if (provinceIndex > 1) return parts[provinceIndex - 1];
  if (parts.length === 2) return parts[1];
  return '';
}

function tmv3_titleLocationDisplay_(location, calendarLocation) {
  if (!location) return '';

  const rawAddress = tmv3_clean_(location['Address 1']);
  const structuredCity = tmv3_clean_(location['City']);

  if (!rawAddress) {
    return (
      structuredCity ||
      tmv3_titleCityFromAddress_(calendarLocation)
    );
  }

  const parts = rawAddress
    .split(',')
    .map(tmv3_clean_)
    .filter(Boolean);

  const street = parts.length ? parts[0] : rawAddress;

  // Prefer the verified Calendar locality when Striven stores a complete
  // formatted address in Address 1 and leaves City blank. This keeps the
  // operational city that staff actually booked (e.g. North York) instead
  // of a broader hierarchy segment (e.g. Toronto).
  const city =
    structuredCity ||
    tmv3_titleCityFromAddress_(calendarLocation) ||
    tmv3_titleCityFromAddress_(rawAddress);

  return [street, city]
    .map(tmv3_clean_)
    .filter(Boolean)
    .join(', ');
}

function tmv3_verifiedTitleContext_(record) {
  record = record || {};

  const step3 = record.step3 || {};
  const step4 = record.step4 || {};
  const anchor = step3.anchor || {};
  const customer = step4.customer || null;
  const location = step4.location || null;

  const verified =
    step4.disposition === 'VERIFIED' &&
    (
      step3.disposition === 'VERIFIED' ||
      (
        record.vertical === 'PreInspection' &&
        (TMV3.VERTICALS.PreInspection || {}).orderRequired !== true
      )
    );

  const customerName = tmv3_clean_(
    customer && customer['Name']
  );

  const customerNumber = tmv3_clean_(
    (customer && customer['Customer Number']) ||
    anchor.customerNumber
  );

  const orderNumber = tmv3_clean_(
    anchor.orderNumber || record.orderNumber
  );

  const phone = tmv3_titlePhoneDisplay_(
    record.phone ||
    (customer && customer['Primary Phone']) ||
    (location && location['Primary Phone'])
  );

  const address =
    tmv3_titleLocationDisplay_(location, record.location) ||
    (
      step4.locationStatus === 'CREATE_REQUIRED' &&
      tmv3_clean_(record.location)
        ? tmv3_titleLocationDisplay_(
            {
              'Address 1':tmv3_clean_(record.location),
              'City':tmv3_titleCityFromAddress_(record.location)
            },
            record.location
          )
        : ''
    );

  return {
    verified: verified,
    customerName: customerName,
    customerNumber: customerNumber,
    orderNumber: orderNumber,
    phone: phone,
    address: address
  };
}

function tmv3_desiredTaskName_(record, options) {
  options = options || {};
  const context = tmv3_verifiedTitleContext_(record);

  if (!context.verified) {
    return {
      status:'BLOCKED',
      value:'',
      reason:'Customer/Location identity is not fully verified.'
    };
  }

  // Locked standard: require verified name, street/city and phone.
  const missing = [];
  if (!context.customerName) missing.push('Customer Name');
  if (!context.address) missing.push('Address');
  if (!context.phone) missing.push('Phone');

  if (missing.length) {
    return {
      status:'BLOCKED',
      value:'',
      reason:'Canonical Task Name is missing verified ' + missing.join(', ') + '.'
    };
  }

  if(record.vertical!=='PreInspection') {
    if(!/^\d+$/.test(String(context.orderNumber||'')))return {status:'BLOCKED',value:'',reason:'Verified Sales Order number required for standard Task naming.'};
    const source=tmv3_clean_(options.sourceTaskTitle);
    const numbers=[];let match;const pattern=/\bSO\s*#?\s*(\d{4,9})\b/gi;
    while((match=pattern.exec(source))!==null)if(numbers.indexOf(match[1])<0)numbers.push(match[1]);
    if(numbers.length>1)return {status:'BLOCKED',value:'',reason:'Existing Task title carries multiple SO numbers; review before renaming.'};
    let standard;
    if(source && numbers.length===1)standard=source.replace(/\bSO\s*#?\s*\d{4,9}\b/gi,'SO#'+context.orderNumber);
    else if(source){
      const level=source.match(/^(Level\s+\d+)\s*-\s*(.*)$/i);
      standard=level ? level[1]+' - SO#'+context.orderNumber+' - '+level[2] : 'SO#'+context.orderNumber+' - '+source;
    } else standard='SO#'+context.orderNumber+' - '+context.customerName+' - '+context.phone;
    const fireplaceNumber=Number(record.serviceFireplaceNumber||0);
    if(record.vertical==='Service' && fireplaceNumber>0 && !new RegExp('\\bFP#'+fireplaceNumber+'\\b','i').test(standard))standard+=' - FP#'+fireplaceNumber;
    return {status:'READY',value:standard,reason:'',source:'VERIFIED_ORDER_PRESERVE_EXISTING_STANDARD_TASK_NAME'};
  }

  let value = TMV3_HARD_RULES.PreInspection.presentation.taskTitleParts.map(function(key) {
    return context[key];
  }).join(TMV3_HARD_RULES.GLOBAL.titleDelimiter);

  if (record.vertical === 'Service') {
    const sourceTitle = tmv3_clean_(options.sourceTaskTitle);
    const fireplaceNumber = Number(
      record.serviceFireplaceNumber ||
      (
        sourceTitle &&
        typeof tmv3_extractServiceFireplaceNumber_ === 'function'
          ? tmv3_extractServiceFireplaceNumber_(sourceTitle)
          : 0
      ) ||
      0
    );

    if (fireplaceNumber > 0) {
      value += ' - FP#' + fireplaceNumber;
    }
  }

  return {
    status:'READY',
    value:value,
    reason:'',
    source:'VERIFIED_CUSTOMER_LOCATION_PHONE'
  };
}

function tmv3_desiredCalendarTitle_(record) {
  record = record || {};
  const context = tmv3_verifiedTitleContext_(record);

  if (!context.verified) {
    return {
      status:'BLOCKED',
      value:'',
      reason:'Calendar title normalization requires verified Customer/Location identity.'
    };
  }

  if (record.vertical === 'Install') {
    if (!context.orderNumber || !context.customerName || !context.phone) {
      return {
        status:'BLOCKED',
        value:'',
        reason:'Install Calendar title requires verified Sales Order Number, Customer Name and Phone.'
      };
    }

    return {
      status:'READY',
      value:TMV3_HARD_RULES.Install.presentation.calendarTitleParts.map(function(key) {
        return key==='orderNumber' ? TMV3_HARD_RULES.Install.presentation.orderNumberPrefix + context.orderNumber : context[key];
      }).join(TMV3_HARD_RULES.GLOBAL.titleDelimiter),
      reason:'',
      source:'VERIFIED_SO_CUSTOMER_PHONE'
    };
  }

  if (record.vertical === 'PreInspection') {
    // Locked standard: incomplete identity blocks non-standard title writes.
    if (!context.customerNumber || !context.customerName || !context.phone) {
      return {
        status:'BLOCKED',
        value:'',
        reason:'PreInspection Calendar title requires verified Customer Number, Customer Name and Phone.'
      };
    }

    return {
      status:'READY',
      value:TMV3_HARD_RULES.PreInspection.presentation.calendarTitleParts.map(function(key) {
        return context[key];
      }).join(TMV3_HARD_RULES.GLOBAL.titleDelimiter),
      reason:'',
      source:'VERIFIED_CUSTOMER_NUMBER_NAME_PHONE'
    };
  }

  return {
    status:'NOT_AUTHORIZED',
    value:'',
    reason:
      record.vertical +
      ' Calendar title normalization has no explicitly approved canonical format. Existing Calendar title must remain unchanged.'
  };
}






// Original: 55_Writes
/************************************************************
 * TM V3 — GUARDED WRITE ENGINE
 ************************************************************/

function tmv3_writesEnabled_() {
  return (TMV3.MODE === 'CANARY_WRITE' || TMV3.MODE === 'PRODUCTION_WRITE') &&
    (TMV3.OPERATIONS.manualWritesEnabled === true ||
      TMV3.OPERATIONS.automationWritesEnabled === true);
}

function tmv3_assertBusinessWritesEnabled_() {
  if (!tmv3_writesEnabled_()) {
    throw new Error('TM V3 external write blocked. Current mode: ' + TMV3.MODE);
  }
}

function tmv3_assertLegacyMutationHelperDisabled_(helperName) {
  throw new Error(
    'STEP7_AUTHORITY_REQUIRED: legacy direct Task mutation helper ' +
    String(helperName || 'UNKNOWN') +
    ' is disabled. Use the canonical Step 7 executor in 58_Operations.js.'
  );
}

function tmv3_buildTaskPatchPayload_(eventRecord, resolved, currentTask) {
  if (!eventRecord || !resolved || !currentTask) {
    throw new Error('Event, resolved identity and current Task are required.');
  }

  const cfg = TMV3.VERTICALS[eventRecord.vertical];
  const payload = {
    Id: Number(currentTask['Task ID']),
    StartDateTime: tmv3_strivenTaskDateTime_(eventRecord.start),
    DueDateTime: tmv3_strivenTaskDateTime_(eventRecord.end)
  };

  const locationId = Number(resolved.locationId || 0);
  const orderId = Number(resolved.orderId || 0);
  const contactId = Number(resolved.contactId || 0);

  if (locationId) payload.Location = { Id: locationId };

  if (cfg.orderRequired && orderId) {
    payload.SalesOrder = { Id: orderId };
  }

  if (eventRecord.vertical !== 'PreInspection' && contactId) {
    payload.RequestedBy = { Id: contactId, Type: 'contact' };
  }

  if (eventRecord.vertical === 'PreInspection') {
    delete payload.SalesOrder;

    const requestedBy = tmv3_resolveOrganizerEmployee_(eventRecord);
    if (requestedBy.status !== 'MATCHED') {
      throw new Error(requestedBy.reason || 'PreInspection Requested By unresolved.');
    }

    payload.RequestedBy = {
      Id: Number(requestedBy.employee.id),
      Name: requestedBy.employee.name || '',
      Type: 'employee'
    };
  }

  return payload;
}

function tmv3_buildAssignmentPlan_(eventRecord, currentTask) {
  const desired = tmv3_desiredAssignment_(eventRecord);
  const currentAssigneeText = tmv3_norm_(currentTask && currentTask['Assignees']);
  const currentPoolText = tmv3_norm_(currentTask && currentTask['Pools']);

  const addEmployeeIds = (desired.employeeIds || []).filter(function(id) {
    const name = tmv3_employeeNameById_(id);
    return (
      currentAssigneeText.indexOf(String(id)) === -1 &&
      (!name || currentAssigneeText.indexOf(tmv3_norm_(name)) === -1)
    );
  });

  const addPoolIds = (desired.poolIds || []).filter(function(id) {
    const cfg = TMV3.VERTICALS[eventRecord.vertical];
    const poolName = cfg.defaultPoolName || '';
    return (
      currentPoolText.indexOf(String(id)) === -1 &&
      (!poolName || currentPoolText.indexOf(tmv3_norm_(poolName)) === -1)
    );
  });

  return {
    addEmployeeIds: addEmployeeIds,
    addPoolIds: addPoolIds,
    desiredEmployeeIds: desired.employeeIds || [],
    desiredPoolIds: desired.poolIds || []
  };
}

function tmv3_employeeNameById_(employeeId) {
  const id = Number(employeeId || 0);
  const known = {
    1: 'Jay Scott',
    6: 'Pramodh',
    9: 'Thang',
    15: 'Stephen Foley',
    18: 'John Hoang',
    20: 'Spencer',
    26: 'Matthew Thompson',
    38: 'Travis Mason',
    39: 'Chris Olomojobi'
  };
  return known[id] || '';
}

function tmv3_safeTaskPatchPayload_(payload, taskId) {
  const allowed = {
    Id: true,
    StartDateTime: true,
    DueDateTime: true,
    SalesOrder: true,
    Location: true,
    RequestedBy: true,
    Title: true,
    Description: true,
    InfoCustomFields: true
  };

  // Striven's Task PATCH contract uses PascalCase date fields.
  // Normalize any internal lower-camel aliases before the external write.
  const normalized = Object.assign({}, payload || {});
  if (
    normalized.startDateTime !== undefined &&
    normalized.StartDateTime === undefined
  ) {
    normalized.StartDateTime = normalized.startDateTime;
  }
  if (
    normalized.dueDateTime !== undefined &&
    normalized.DueDateTime === undefined
  ) {
    normalized.DueDateTime = normalized.dueDateTime;
  }
  delete normalized.startDateTime;
  delete normalized.dueDateTime;

  const safe = { Id: Number(taskId) };

  Object.keys(normalized).forEach(function(key) {
    if (!allowed[key]) {
      throw new Error('Blocked unknown Task PATCH field: ' + key);
    }
    if (key === 'Id') return;
    const value = normalized[key];
    if (value === undefined || value === null || value === '') return;
    safe[key] = value;
  });

  return safe;
}

function tmv3_patchTaskById_(taskId, payload) {
  tmv3_assertLegacyMutationHelperDisabled_('tmv3_patchTaskById_');
  tmv3_assertBusinessWritesEnabled_();

  const id = Number(taskId || 0);
  if (!id) throw new Error('Task ID is required.');

  const safe = tmv3_safeTaskPatchPayload_(payload, id);
  if (Object.keys(safe).length <= 1) {
    return { status: 'NOT_NEEDED', taskId: id, payload: safe };
  }

  tmv3_fetchJson_(
    TMV3.API_BASE + '/v2/tasks/' + encodeURIComponent(id),
    {
      method: 'patch',
      contentType: 'application/json',
      payload: JSON.stringify(safe)
    }
  );

  return {
    status: 'PATCHED_READBACK_REQUIRED',
    taskId: id,
    payload: safe,
    readback: tmv3_getTaskById_(id)
  };
}

function tmv3_addAssignment_(taskId, id, type, name) {
  tmv3_assertLegacyMutationHelperDisabled_('tmv3_addAssignment_');
  tmv3_assertBusinessWritesEnabled_();

  return tmv3_fetchJson_(
    TMV3.API_BASE + '/v2/tasks/' + encodeURIComponent(Number(taskId)) + '/assignments',
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        Id: Number(id),
        Name: name || '',
        Type: type
      })
    }
  );
}

function tmv3_applyAssignmentPlan_(taskId, plan) {
  tmv3_assertLegacyMutationHelperDisabled_('tmv3_applyAssignmentPlan_');
  tmv3_assertBusinessWritesEnabled_();

  const results = [];

  (plan.addEmployeeIds || []).forEach(function(id) {
    results.push({
      type: 'employee',
      id: Number(id),
      response: tmv3_addAssignment_(taskId, id, 'employee', tmv3_employeeNameById_(id))
    });
  });

  (plan.addPoolIds || []).forEach(function(id) {
    const poolName = Number(id) === 8
      ? 'Pre-Inspection Pool'
      : (Number(id) === 4 ? 'To Be Assigned' : '');

    results.push({
      type: 'pool',
      id: Number(id),
      response: tmv3_addAssignment_(taskId, id, 'pool', poolName)
    });
  });

  return results;
}

function tmv3_previewTaskMutation_(eventRecord, resolved, refs) {
  if (!resolved || !resolved.status) {
    throw new Error('Resolved record is required.');
  }

  if (['REVIEW','BLOCKED','IGNORED'].indexOf(resolved.status) !== -1) {
    return {
      allowed: false,
      action: 'NONE',
      reason: resolved.issue || resolved.status
    };
  }

  if (resolved.status === 'MATCHED') {
    return {
      allowed: true,
      action: 'NONE',
      reason: 'Already verified.'
    };
  }

  if (resolved.taskId) {
    const currentTask = refs.taskById[String(resolved.taskId)] ||
      tmv3_getTaskById_(resolved.taskId);

    return {
      allowed: true,
      action: 'PATCH_TASK',
      taskId: Number(resolved.taskId),
      payload: tmv3_buildTaskPatchPayload_(eventRecord, resolved, currentTask),
      assignments: tmv3_buildAssignmentPlan_(eventRecord, currentTask),
      calendarLinksAfterTaskVerification: true
    };
  }

  if (eventRecord.vertical === 'PreInspection' && resolved.status === 'READY CREATE') {
    const requestedBy = tmv3_resolveOrganizerEmployee_(eventRecord);

    if (requestedBy.status !== 'MATCHED') {
      return {
        allowed: false,
        action: 'NONE',
        reason: requestedBy.reason
      };
    }

    return {
      allowed: true,
      action: 'CREATE_PREINSPECTION_TASK',
      sourceTemplateTaskId: 11138,
      invariant: {
        taskTypeId: 105,
        salesOrderAttached: false,
        description: '',
        field854: 'DO_NOT_MANAGE',
        technicianFields: 'DO_NOT_PREFILL'
      },
      event: {
        startDateTime: tmv3_strivenTaskDateTime_(eventRecord.start),
        dueDateTime: tmv3_strivenTaskDateTime_(eventRecord.end)
      },
      customerId: Number(resolved.customerId),
      locationId: Number(resolved.locationId),
      requestedBy: {
        Id: Number(requestedBy.employee.id),
        Name: requestedBy.employee.name || '',
        Type: 'employee'
      },
      assignments: tmv3_desiredAssignment_(eventRecord),
      calendarLinksAfterTaskVerification: true
    };
  }

  return {
    allowed: false,
    action: resolved.status === 'READY RECREATE'
      ? 'RECREATE_PENDING_SOURCE_TEMPLATE'
      : 'CREATE_PENDING_VERTICAL_TEMPLATE',
    reason: 'CREATE source template is not yet proven for this vertical. Write remains blocked.'
  };
}


// Original: 58_Operations
/************************************************************
 * TM V3 — FEATURE PARITY OPERATIONS
 *
 * Purpose:
 * - Restore the proven operational "hands" from the original
 *   Task Mapping system behind the cleaner V3 resolver/UI.
 * - Manual actions always fresh-read Calendar + Striven first.
 * - Scheduled external writes remain independently gated.
 * - CREATE outcomes are never blindly retried.
 ************************************************************/

function tmv3_operationPolicy_() {
  return TMV3.OPERATIONS || {};
}

function tmv3_operationWritesEnabled_(scope) {
  const policy = tmv3_operationPolicy_();
  const mode = tmv3_norm_(scope || '');
  const externalWriteMode =
    TMV3.MODE === 'CANARY_WRITE' ||
    TMV3.MODE === 'PRODUCTION_WRITE';

  // A policy flag may narrow an authorized write mode, but it must never
  // elevate SHADOW_READ_ONLY into a write-capable mode.
  if (!externalWriteMode) return false;

  if (mode === 'verticalauto') {
    const context = typeof tmv3_verticalProductionCycle_ === 'function' && tmv3_verticalProductionCycle_.context;
    return TMV3.MODE === 'PRODUCTION_WRITE' && policy.verticalProductionEnabled === true &&
      policy.manualWritesEnabled === false && policy.automationWritesEnabled === false &&
      !!context && ['Install','Delivery','Service'].indexOf(context.vertical)>=0 &&
      (policy.verticalProductionVerticals || []).indexOf(context.vertical)>=0;
  }

  if (mode === 'preinspectionauto') {
    return TMV3.MODE === 'PRODUCTION_WRITE' &&
      policy.preInspectionAutomationEnabled === true &&
      policy.automationWritesEnabled === false &&
      policy.manualWritesEnabled === false;
  }

  if (mode === 'servicedeliverydetails') {
    const context = typeof tmv3_taskFirstProductionCycle_ === 'function' &&
      tmv3_taskFirstProductionCycle_.detailContext;
    return TMV3.MODE === 'PRODUCTION_WRITE' &&
      policy.taskFirstProductionEnabled === true &&
      policy.manualWritesEnabled === false && policy.automationWritesEnabled === false &&
      Number(policy.taskFirstProductionMaxWritesPerRun) === 1 && !!context &&
      (context.vertical === 'Service' || context.vertical === 'Delivery');
  }

  // The dedicated production permission cannot unlock legacy/manual/calendar writers.
  if (mode === 'taskfirstproduction') {
    return TMV3.MODE === 'PRODUCTION_WRITE' &&
      policy.taskFirstProductionEnabled === true &&
      policy.manualWritesEnabled === false &&
      policy.automationWritesEnabled === false &&
      Number(policy.taskFirstProductionMaxWritesPerRun) === 1;
  }

  if (mode === 'taskfirstcanary19083') {
    return TMV3.MODE === 'CANARY_WRITE' &&
      policy.taskFirstCanaryEnabled === true &&
      Number(policy.taskFirstCanaryTaskId) === 19083 &&
      policy.manualWritesEnabled === false &&
      policy.automationWritesEnabled === false;
  }

  if (mode === 'manual') {
    return policy.manualWritesEnabled === true;
  }

  if (mode === 'auto') {
    return policy.automationWritesEnabled === true;
  }

  return false;
}

function tmv3_assertOperationWrite_(scope, action) {
  tmv3_assertHardRules_();

  if (!tmv3_operationWritesEnabled_(scope)) {
    throw new Error(
      'TM V3 ' + String(scope || 'UNKNOWN') +
      ' external write is gated by configuration.'
    );
  }
  const normalizedScope = tmv3_norm_(scope);
  if (normalizedScope === 'servicedeliverydetails' &&
      ['TASK_RELATIONSHIP_PATCH','TASK_DETAIL_ASSIGNMENT'].indexOf(action) < 0) {
    throw new Error('TASK_DETAILS_ONLY_SCOPE: no Calendar, creation, notes or title writes.');
  }
  if ((normalizedScope === 'taskfirstproduction' || normalizedScope === 'taskfirstcanary19083') &&
      action !== 'TASK_RELATIONSHIP_REPAIR' && action !== 'TASK_RELATIONSHIP_PATCH') {
    throw new Error('RELATIONSHIP_ONLY_SCOPE: Task-first permission cannot authorize this action.');
  }
}

function tmv3_selectedContext_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const range = sh && sh.getActiveRange();

  if (!sh || !range) {
    throw new Error('Select an appointment row first.');
  }

  const sheetName = sh.getName();
  let headerRow = 1;
  let vertical = '';

  if (sheetName === TMV3.SHEETS.MORNING) {
    headerRow = 16;
  } else {
    Object.keys(TMV3.VERTICALS).some(function(key) {
      if (TMV3.VERTICALS[key].sheet === sheetName) {
        vertical = key;
        return true;
      }
      return false;
    });
  }

  const rowNumber = range.getRow();
  if (rowNumber <= headerRow) {
    throw new Error('Select a data row, not a header.');
  }

  const lastColumn = sh.getLastColumn();
  const headers = sh.getRange(headerRow, 1, 1, lastColumn).getValues()[0]
    .map(tmv3_clean_);
  const values = sh.getRange(rowNumber, 1, 1, lastColumn).getValues()[0];
  const row = {};

  headers.forEach(function(header, i) {
    if (header) row[header] = values[i];
  });

  if (!vertical) {
    vertical = tmv3_clean_(row['Vertical']);
  }

  if (!TMV3.VERTICALS[vertical]) {
    throw new Error(
      'Selected row does not identify Install, Delivery, Service, or PreInspection.'
    );
  }

  const eventId = tmv3_clean_(row['Event ID']);
  if (!eventId) {
    throw new Error('Selected row has no Event ID. Refresh V3 mapping first.');
  }

  return {
    spreadsheet: ss,
    sheet: sh,
    sheetName: sheetName,
    headerRow: headerRow,
    rowNumber: rowNumber,
    headers: headers,
    row: row,
    vertical: vertical,
    eventId: eventId,
    taskHint: tmv3_clean_(row['Task']),
    status: tmv3_clean_(row['Status']).toUpperCase()
  };
}

function tmv3_taskIdHintFromText_(text) {
  const raw = tmv3_clean_(text);
  if (!raw) return '';

  let m = raw.match(/\bTask\s*#?\s*(\d+)\b/i);
  if (m) return m[1];

  m = raw.match(/^\s*(\d{3,})\b/);
  return m ? m[1] : '';
}

function tmv3_findCalendarEventById_(calendarId, eventId) {
  const calendar = CalendarApp.getCalendarById(tmv3_clean_(calendarId));
  if (!calendar) return null;

  const raw = tmv3_clean_(eventId);
  const normalized = raw.replace(/@google\.com$/i, '');
  const candidates = tmv3_unique_([
    raw,
    normalized,
    normalized ? normalized + '@google.com' : ''
  ]);

  for (let i = 0; i < candidates.length; i++) {
    try {
      const event = calendar.getEventById(candidates[i]);
      if (event) {
        return { calendar: calendar, event: event, match: 'DIRECT' };
      }
    } catch (ignored) {}
  }

  return null;
}

function tmv3_findFreshEventRecord_(vertical, eventId) {
  const cfg = TMV3.VERTICALS[vertical];
  const calendars = tmv3_verticalCalendars_(vertical, cfg);

  for (let i = 0; i < calendars.length; i++) {
    const calCfg = calendars[i];
    const found = tmv3_findCalendarEventById_(calCfg.calendarId, eventId);
    if (!found || !found.event) continue;

    const record = tmv3_calendarEvent_(vertical, cfg, calCfg, found.event);
    if (record) return record;
  }

  throw new Error(
    vertical + ' Calendar event not found for Event ID ' + eventId + '.'
  );
}

function tmv3_freshSelectedResolution_() {
  const context = tmv3_selectedContext_();
  const eventRecord = tmv3_findFreshEventRecord_(context.vertical, context.eventId);
  const refs = tmv3_referenceIndex_();
  const state = tmv3_eventStateIndex_();
  const records = tmv3_resolveEventRecords_(eventRecord, refs, state);

  if (!records.length) {
    throw new Error('Fresh V3 resolution returned no appointment record.');
  }

  let selected = null;
  const taskHintId = tmv3_taskIdHintFromText_(context.taskHint);

  if (taskHintId) {
    const matches = records.filter(function(record) {
      return String(record.taskId || '') === String(taskHintId);
    });
    if (matches.length === 1) selected = matches[0];
  }

  if (!selected && context.taskHint) {
    const normalizedHint = tmv3_norm_(context.taskHint);
    const matches = records.filter(function(record) {
      return normalizedHint && tmv3_norm_(record.task).indexOf(normalizedHint) !== -1;
    });
    if (matches.length === 1) selected = matches[0];
  }

  if (!selected && records.length === 1) selected = records[0];

  if (!selected) {
    throw new Error(
      'Selected Calendar event currently maps to ' + records.length +
      ' V3 records. Select the exact task row and refresh if needed.'
    );
  }

  return {
    context: context,
    eventRecord: eventRecord,
    refs: refs,
    state: state,
    records: records,
    resolved: selected
  };
}

function tmv3_step7PublishedPlan_(vertical, eventId, taskId) {
  const rows = tmv3_rows_(TMV3.SHEETS.RECONCILE);
  const wantedTaskId = Number(taskId || 0);
  const matches = rows.filter(function(row) {
    if (
      tmv3_clean_(row['Vertical']) !== tmv3_clean_(vertical) ||
      tmv3_clean_(row['Event ID']) !== tmv3_clean_(eventId)
    ) return false;

    const rowTaskId = Number(row['Task ID'] || 0);
    return wantedTaskId ? rowTaskId === wantedTaskId : rowTaskId === 0;
  });

  return matches.length === 1
    ? tmv3_clean_(matches[0]['Relationship Plan'])
    : '';
}

function tmv3_step7FreshSelectedContract_() {
  const context = tmv3_selectedContext_();
  const taskId = Number(tmv3_taskIdHintFromText_(context.taskHint) || 0);
  const previousPlan = tmv3_step7PublishedPlan_(
    context.vertical,
    context.eventId,
    taskId
  );
  const contract = tmv3_step7FreshExecutionContract_(
    context.vertical,
    context.eventId,
    taskId
  );

  return {
    context:context,
    previousPlan:previousPlan,
    contract:contract
  };
}

function tmv3_step7PlanChangedResult_(selection) {
  const previous = tmv3_clean_(
    selection && selection.previousPlan
  );
  const fresh = tmv3_clean_(
    selection &&
    selection.contract &&
    selection.contract.plan
  );

  if (!previous) {
    return {
      status:'PLAN_CHANGED_NO_WRITE',
      previousPlan:'',
      freshPlan:fresh,
      reason:
        'No previously published Step 7 plan exists for this exact Event/Task row. ' +
        'Refresh V3 before execution.',
      vertical:selection.contract.vertical,
      eventId:selection.contract.eventId,
      taskId:selection.contract.taskId || ''
    };
  }

  if (previous === fresh) return null;

  return {
    status:'PLAN_CHANGED_NO_WRITE',
    previousPlan:previous,
    freshPlan:fresh,
    reason:
      'The canonical Step 7 plan changed after the operator view was published. ' +
      'Refresh and review the new plan before execution.',
    vertical:selection.contract.vertical,
    eventId:selection.contract.eventId,
    taskId:selection.contract.taskId || ''
  };
}

function tmv3_step7BundleFromContract_(contract) {
  contract = tmv3_step7ValidateExecutionContract_(contract);
  const eventRecord = tmv3_findFreshEventRecord_(
    contract.vertical,
    contract.eventId
  );

  if (
    contract.calendarId &&
    tmv3_clean_(eventRecord.calendarId) !==
      tmv3_clean_(contract.calendarId)
  ) {
    throw new Error(
      'PLAN_CHANGED_NO_WRITE: Calendar ID changed after Step 7 planning.'
    );
  }

  if (
    contract.inputFingerprint &&
    tmv3_clean_(eventRecord.fingerprint) !==
      tmv3_clean_(contract.inputFingerprint)
  ) {
    throw new Error(
      'PLAN_CHANGED_NO_WRITE: Calendar event changed after Step 7 planning.'
    );
  }

  const resolved = {
    taskId:Number(contract.taskId || 0) || '',
    task:contract.taskId ? (tmv3_clean_(contract.desiredTaskName || contract.actualTaskName) || String(contract.taskId)) : '',
    order:tmv3_clean_(eventRecord.orderNumber),
    customerId:tmv3_clean_(contract.expectedCustomerId),
    locationId:tmv3_clean_(contract.expectedLocationId),
    contactId:
      tmv3_norm_(contract.expectedRequestedByType) === 'contact'
        ? tmv3_clean_(contract.expectedRequestedById)
        : '',
    orderId:tmv3_clean_(contract.expectedOrderId),
    status:contract.taskId
      ? 'READY'
      : (
          contract.actions.indexOf(TMV3_STEP7_ACTION.RECREATE_TASK) !== -1
            ? 'READY RECREATE'
            : 'READY CREATE'
        )
  };

  return {
    context:null,
    eventRecord:eventRecord,
    refs:tmv3_referenceIndex_(),
    state:tmv3_eventStateIndex_(),
    records:[resolved],
    resolved:resolved,
    contract:contract
  };
}

function tmv3_step7FreshLocationOwnership_(contract) {
  const locationId = tmv3_clean_(
    contract.expectedLocationId
  );
  const customerId = tmv3_clean_(
    contract.expectedCustomerId
  );

  if (!locationId) return { status:'N/A' };

  if (contract.vertical === 'PreInspection') {
    const owned=tmv3_preInspectionCustomerLocationsFresh_(customerId);
    const matches=owned.filter(function(row){return tmv3_clean_(row['Location ID'])===locationId;});
    if (matches.length!==1) throw new Error('Fresh PreInspection Location ownership could not be verified.');
    return {status:'VERIFIED',locationId:locationId,customerId:customerId};
  }

  // The Customer-number map is also freshly read from Striven so a stale
  // Source Customers sheet cannot manufacture Location ownership.
  const customerNumberToId = {};

  tmv3_reportRows_(TMV3.PROPERTIES.CUSTOMERS)
    .map(tmv3_normalizeCustomer_)
    .forEach(function(row) {
      const number = tmv3_clean_(row[1]);
      const id = tmv3_clean_(row[0]);
      if (number && id) customerNumberToId[number] = id;
    });

  const matches = tmv3_reportRows_(
    TMV3.PROPERTIES.LOCATIONS
  )
    .map(function(row) {
      return tmv3_normalizeLocation_(
        row,
        customerNumberToId
      );
    })
    .filter(function(row) {
      return tmv3_clean_(row[0]) === locationId;
    });

  if (matches.length !== 1) {
    throw new Error(
      'Fresh Location ownership verification expected one Location ' +
      locationId + '; found ' + matches.length + '.'
    );
  }

  if (tmv3_clean_(matches[0][1]) !== customerId) {
    throw new Error(
      'Fresh Location ownership mismatch: Location ' +
      locationId + ' does not belong to Customer ' +
      customerId + '.'
    );
  }

  return {
    status:'VERIFIED',
    locationId:locationId,
    customerId:customerId
  };
}

function tmv3_step7FreshOrderOwnership_(contract, eventRecord) {
  if (contract.vertical === 'PreInspection') return {status:'N/A'};
  const number=tmv3_clean_(eventRecord && eventRecord.orderNumber);
  if (!/^\d+$/.test(number)) throw new Error('FRESH_CALENDAR_ORDER_NUMBER_REQUIRED');
  const proof=tmv3_taskFirstFreshOrderOwnershipById_(contract.expectedOrderId,
    contract.expectedCustomerId,'SO#'+number,contract.vertical);
  return {status:'VERIFIED',orderId:proof.orderId,customerId:proof.customerId,orderStatus:proof.status};
}


function tmv3_step7FreshCriticalOwnership_(contract, eventRecord) {
  contract = tmv3_step7ValidateExecutionContract_(contract);

  if (!tmv3_clean_(contract.expectedCustomerId)) {
    throw new Error(
      'Step 7 contract is missing the resolved Customer ID.'
    );
  }

  let task = null;
  if (contract.taskId) {
    const raw = tmv3_rawTaskById_(contract.taskId);
    task = tmv3_normalizeV2TaskModel_(raw || {});

    if (!tmv3_taskIsOpen_(task['Status'])) {
      throw new Error('Fresh Task state is no longer OPEN.');
    }

    if (
      tmv3_clean_(task['Customer ID']) !==
      tmv3_clean_(contract.expectedCustomerId)
    ) {
      throw new Error(
        'Fresh Task Customer no longer matches the canonical Step 7 Customer.'
      );
    }
  } else {
    const customers = tmv3_reportRows_(TMV3.PROPERTIES.CUSTOMERS)
      .map(tmv3_normalizeCustomer_)
      .filter(function(row) {
        return tmv3_clean_(row[0]) ===
          tmv3_clean_(contract.expectedCustomerId);
      });

    // A joined report may repeat the same primary Customer ID. This check
    // establishes existence only; fresh Order/Location ownership still gates writes.
    if (customers.length === 0) {
      throw new Error(
        'Fresh Customer existence verification expected one Customer ' +
        contract.expectedCustomerId + '; found ' +
        customers.length + '.'
      );
    }
  }

  const location = tmv3_step7FreshLocationOwnership_(contract);
  const order = tmv3_step7FreshOrderOwnership_(contract, eventRecord);

  let requestedBy = { status:'N/A' };

  if (tmv3_norm_(contract.expectedRequestedByType) === 'contact') {
    const contact = tmv3_getContactById_(
      contract.expectedRequestedById,
      contract.expectedCustomerId
    );

    if (!contact.__ownershipVerified) {
      throw new Error(
        'Fresh Requested By Contact ownership could not be proven for Customer ' +
        contract.expectedCustomerId + '.'
      );
    }

    requestedBy = {
      status:'VERIFIED',
      type:'contact',
      id:contract.expectedRequestedById
    };
  } else if (
    contract.vertical === 'PreInspection' &&
    tmv3_norm_(contract.expectedRequestedByType) === 'employee'
  ) {
    const organizer = tmv3_resolveOrganizerEmployee_(eventRecord);

    if (
      !organizer ||
      organizer.status !== 'MATCHED' ||
      Number(organizer.employee && organizer.employee.id || 0) !==
        Number(contract.expectedRequestedById || 0)
    ) {
      throw new Error(
        'Fresh PreInspection organizer Employee no longer matches ' +
        'the canonical Requested By.'
      );
    }

    requestedBy = {
      status:'VERIFIED',
      type:'employee',
      id:contract.expectedRequestedById
    };
  }

  return {
    task:task,
    location:location,
    order:order,
    requestedBy:requestedBy
  };
}

function tmv3_step7DatePayload_(contract) {
  return {
    Id:Number(contract.taskId),
    StartDateTime:tmv3_strivenTaskDateTime_(
      new Date(contract.expectedStart)
    ),
    DueDateTime:tmv3_strivenTaskDateTime_(
      new Date(contract.expectedDue)
    )
  };
}

function tmv3_step7RelationshipPayload_(contract) {
  const payload = { Id:Number(contract.taskId) };
  const actions = contract.actions || [];

  if (actions.indexOf(TMV3_STEP7_ACTION.PATCH_LOCATION) !== -1) {
    payload.Location = { Id:Number(contract.expectedLocationId) };
  }

  if (actions.indexOf(TMV3_STEP7_ACTION.PATCH_ORDER) !== -1) {
    payload.SalesOrder = { Id:Number(contract.expectedOrderId) };
  }

  if (
    actions.indexOf(TMV3_STEP7_ACTION.PATCH_REQUESTED_BY) !== -1
  ) {
    payload.RequestedBy = {
      Id:Number(contract.expectedRequestedById),
      Type:tmv3_clean_(contract.expectedRequestedByType)
    };
  }

  return payload;
}

function tmv3_step7DesiredAssignmentFromContract_(contract) {
  return {
    employeeIds:
      (contract.desiredAssignmentEmployeeIds || []).map(Number),
    poolIds:
      (contract.desiredAssignmentPoolIds || []).map(Number)
  };
}

function tmv3_step7ActionsForMode_(contract, mode) {
  const actions = (contract.actions || []).slice();
  const map = {
    DATES:[TMV3_STEP7_ACTION.PATCH_DATES],
    RELATIONSHIPS:[
      TMV3_STEP7_ACTION.CREATE_LOCATION,
      TMV3_STEP7_ACTION.PATCH_LOCATION,
      TMV3_STEP7_ACTION.PATCH_ORDER,
      TMV3_STEP7_ACTION.PATCH_REQUESTED_BY
    ],
    ASSIGNEE:[TMV3_STEP7_ACTION.PATCH_ASSIGNMENTS],
    LINKS:[TMV3_STEP7_ACTION.VERIFY_CALENDAR_LINKS],
    TITLE:[
      TMV3_STEP7_ACTION.PATCH_CALENDAR_TITLE,
      TMV3_STEP7_ACTION.PATCH_TASK_NAME
    ]
  };

  if (mode === 'ALL') return actions;

  const allowed = map[mode];
  if (!allowed) {
    throw new Error(
      'Unsupported Step 7 execution mode: ' + mode + '.'
    );
  }

  return actions.filter(function(action) {
    return allowed.indexOf(action) !== -1;
  });
}

function tmv3_executeFreshStep7Selection_(selection, scope, mode) {
  const target=selection&&selection.contract;
  if (target && target.vertical === 'PreInspection') tmv3_piAssertLockedStandard_();
  if(target&&target.vertical==='PreInspection'&&tmv3_preInspectionApprovalHold_(target.eventId))
    return {status:'BLOCKED_NO_WRITE',eventId:target.eventId,vertical:'PreInspection',
      blocker:'OPERATOR_APPROVED_INTAKE_AWAITING_FRESH_VERIFICATION',
      taskWrites:0,calendarWrites:0,emailsSent:0};
  if (tmv3_norm_(scope) === 'preinspectionauto' &&
      (!selection || !selection.contract || selection.contract.vertical !== 'PreInspection')) {
    throw new Error('PREINSPECTION_ONLY_SCOPE: other verticals cannot execute.');
  }
  if (tmv3_norm_(scope)==='verticalauto') {
    const c=tmv3_verticalProductionCycle_.context,p=selection && selection.contract;
    if (!c || !p || p.vertical!==c.vertical || p.eventId!==c.eventId ||
        JSON.stringify(p)!==c.contractJson ||
        (p.vertical==='Service' && (p.actions||[]).some(function(a){return /CREATE|RECREATE/.test(a);})) ||
        p.vertical==='PreInspection') throw new Error('VERTICAL_AUTO_OUTSIDE_FRESH_CONTRACT');
  }
  const changed = tmv3_step7PlanChangedResult_(selection);
  if (changed) {
    tmv3_audit_(
      changed.vertical || 'SYSTEM',
      changed.eventId || '',
      changed.taskId || '',
      'STEP7_PLAN_GUARD',
      'NO_WRITE',
      JSON.stringify(changed)
    );
    return changed;
  }

  let contract = tmv3_step7ValidateExecutionContract_(
    selection.contract
  );

  if (tmv3_clean_(contract.blocker)) {
    return {
      status:'BLOCKED_NO_WRITE',
      blocker:contract.blocker,
      vertical:contract.vertical,
      eventId:contract.eventId,
      taskId:contract.taskId || '',
      plan:contract.plan
    };
  }

  let bundle = tmv3_step7BundleFromContract_(contract);
  if (contract.vertical === 'PreInspection') {
    const owner = tmv3_piVerifiedOwnerEmployee_(bundle.eventRecord);
    if (owner.status !== 'MATCHED') return {
      status:'BLOCKED_NO_WRITE',vertical:'PreInspection',eventId:contract.eventId,
      blocker:owner.blocker || 'PI_OWNER_UNVERIFIED_NO_WRITE',
      taskWrites:0,calendarWrites:0,emailsSent:0
    };
  }

  if (
    TMV3.MODE === 'CANARY_WRITE' &&
    !tmv3_canaryWriteEventAllowed_(bundle.eventRecord)
  ) {
    throw new Error(
      'CANARY_WRITE blocked: this Step 7 contract is not the authorized canary event.'
    );
  }

  tmv3_step7FreshCriticalOwnership_(
    contract,
    bundle.eventRecord
  );

  const requestedActions = tmv3_step7ActionsForMode_(
    contract,
    mode || 'ALL'
  );

  if (
    requestedActions.indexOf(TMV3_STEP7_ACTION.CREATE_LOCATION) !== -1
  ) {
    const locationResult = tmv3_ensureStep7Location_(
      bundle,
      scope,
      contract
    );

    bundle = locationResult.bundle;
    contract = locationResult.contract;

    tmv3_step7FreshCriticalOwnership_(
      contract,
      bundle.eventRecord
    );
  }

  const createAction =
    contract.actions.indexOf(TMV3_STEP7_ACTION.CREATE_TASK) !== -1 ||
    contract.actions.indexOf(TMV3_STEP7_ACTION.RECREATE_TASK) !== -1;

  return createAction
    ? tmv3_createOrRecreateFromBundle_(
        bundle,
        scope,
        contract
      )
    : tmv3_executeExistingTaskSync_(
        bundle,
        scope,
        mode || 'ALL',
        contract
      );
}

function tmv3_assertResolvedRecordWritable_(resolved, options) {
  options = options || {};
  const status = tmv3_clean_(resolved && resolved.status).toUpperCase();

  if (!resolved) throw new Error('Resolved appointment is required.');

  if (['REVIEW','BLOCKED','IGNORED','NOT MATCHED'].indexOf(status) !== -1) {
    throw new Error(
      'Write blocked because fresh V3 status is ' + status + '. ' +
      tmv3_clean_(resolved.issue)
    );
  }

  if (options.requireTask && !resolved.taskId) {
    throw new Error('Write requires an existing verified Task ID.');
  }

  return true;
}

function tmv3_rawTaskById_(taskId) {
  const id = Number(taskId || 0);
  if (!id) throw new Error('Positive Task ID is required.');

  return tmv3_fetchJson_(
    TMV3.API_BASE + '/v2/tasks/' + encodeURIComponent(id),
    { method: 'get' }
  ) || {};
}

function tmv3_taskAssignmentsFromRaw_(raw) {
  raw = raw || {};
  const arrays = [
    raw.assignments, raw.Assignments, raw.assignedTo, raw.AssignedTo,
    raw.assignedUsers, raw.AssignedUsers, raw.assignees, raw.Assignees
  ];
  const output = [];

  arrays.forEach(function(source) {
    if (!source) return;
    const items = Array.isArray(source) ? source : [source];

    items.forEach(function(item) {
      if (!item || typeof item !== 'object') return;
      const id = Number(
        item.Id || item.id ||
        item.EmployeeId || item.employeeId || item.EmployeeID || item.employeeID ||
        item.PoolId || item.poolId || item.PoolID || item.poolID || 0
      );
      if (!id) return;

      const rawType = tmv3_norm_(
        item.Type || item.type || item.AssignmentType || item.assignmentType ||
        item.EntityType || item.entityType || ''
      );
      const name = tmv3_clean_(
        item.Name || item.name || item.DisplayName || item.displayName || ''
      );
      const looksLikePool =
        rawType === 'pool' || rawType === 'queue' ||
        tmv3_norm_(name) === 'to be assigned' ||
        tmv3_norm_(name) === 'pre-inspection pool' ||
        Number(item.PoolId || item.poolId || 0) > 0 ||
        id === 4 || id === 8;

      const key = (looksLikePool ? 'pool' : 'employee') + '|' + id;
      if (output.some(function(existing) { return existing.__key === key; })) return;
      output.push({
        id: id,
        type: looksLikePool ? 'pool' : 'employee',
        name: name,
        __key: key
      });
    });
  });

  return output.map(function(item) {
    return { id: item.id, type: item.type, name: item.name };
  });
}

function tmv3_upsertTaskCacheRow_(task) {
  if (!task || !task['Task ID']) return;
  if(Number(task['Task Type ID'])===105) task=tmv3_preInspectionCanonicalCacheTask_(task,tmv3_getCanonicalV1TaskSchedule_(task['Task ID']));

  const sh = tmv3_sheet_(TMV3.SHEETS.TASKS);
  const headers = [
    'Task ID','Task Number','Task Type ID','Task Type','Status','Name',
    'Customer ID','Location ID','Contact ID','Order ID','Start','Due',
    'Assignees','Pools','URL','Fingerprint'
  ];

  const row = headers.map(function(header) {
    return task[header] === undefined ? '' : task[header];
  });

  if (sh.getLastRow() < 1 || sh.getLastColumn() < headers.length) {
    tmv3_replaceRows_(TMV3.SHEETS.TASKS, headers, [row]);
    return;
  }

  const currentHeaders = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(tmv3_clean_);
  const idCol = currentHeaders.indexOf('Task ID') + 1;
  if (!idCol) {
    tmv3_replaceRows_(TMV3.SHEETS.TASKS, headers, [row]);
    return;
  }

  let targetRow = 0;
  if (sh.getLastRow() >= 2) {
    const ids = sh.getRange(2,idCol,sh.getLastRow()-1,1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '') === String(task['Task ID'])) {
        targetRow = i + 2;
        break;
      }
    }
  }

  if (!targetRow) targetRow = sh.getLastRow() + 1;
  sh.getRange(targetRow,1,1,headers.length).setValues([row]);
}

function tmv3_operationPatchTask_(taskId, payload, scope) {
  tmv3_assertOperationWrite_(scope, 'TASK_RELATIONSHIP_PATCH');
  const id = Number(taskId || 0);
  if (!id) throw new Error('Task ID is required.');
  const normalizedScope = tmv3_norm_(scope);
  if (normalizedScope.indexOf('preinspection') >= 0 ||
      (typeof TMV3_PREINSPECTION_API_CONTEXT_ACTIVE !== 'undefined' && TMV3_PREINSPECTION_API_CONTEXT_ACTIVE === true)) {
    tmv3_piAssertSafeTaskPatch_(payload);
  }
  if (normalizedScope === 'taskfirstproduction' || normalizedScope === 'taskfirstcanary19083') {
    const allowed = ['Id','SalesOrder','RequestedBy','Location'];
    if (!Number.isSafeInteger(id) || id <= 0 || Number((payload || {}).Id) !== id ||
        Object.keys(payload || {}).some(function(key) { return allowed.indexOf(key) < 0; }) ||
        (normalizedScope === 'taskfirstcanary19083' && id !== Number(TMV3.OPERATIONS.taskFirstCanaryTaskId))) {
      throw new Error('RELATIONSHIP_ONLY_SCOPE: Exact Task ID and relationship-only payload required.');
    }
    if (normalizedScope === 'taskfirstcanary19083' && (
        Boolean(payload.Location) ||
        (payload.SalesOrder && Number(payload.SalesOrder.Id) !== Number(TMV3.OPERATIONS.taskFirstCanaryExpectedOrderId)) ||
        (payload.RequestedBy && (Number(payload.RequestedBy.Id) !== Number(TMV3.OPERATIONS.taskFirstCanaryExpectedContactId) ||
          payload.RequestedBy.Type !== 'contact')))) {
      throw new Error('RELATIONSHIP_ONLY_SCOPE: Canary payload differs from the exact approved tuple.');
    }
  }

  if (normalizedScope === 'servicedeliverydetails') {
    const context = tmv3_taskFirstProductionCycle_.detailContext;
    const allowed = ['Id','SalesOrder','RequestedBy','Location','StartDateTime','DueDateTime'];
    if (!context || Number(context.taskId) !== id ||
        Number((payload || {}).Id) !== id ||
        Object.keys(payload || {}).some(function(k){return allowed.indexOf(k)<0;}) ||
        JSON.stringify(payload) !== JSON.stringify(context.payload)) {
      throw new Error('TASK_DETAILS_PAYLOAD_OUTSIDE_FRESH_CONTRACT');
    }
  }

  const safe = tmv3_safeTaskPatchPayload_(payload || {}, id);
  if (Object.keys(safe).length <= 1) {
    return {
      status: 'NOT_NEEDED',
      taskId: id,
      payload: safe,
      readback: tmv3_getTaskById_(id)
    };
  }

  tmv3_fetchJson_(
    TMV3.API_BASE + '/v2/tasks/' + encodeURIComponent(id),
    {
      method: 'patch',
      contentType: 'application/json',
      payload: JSON.stringify(safe)
    }
  );

  const readback = tmv3_getTaskById_(id);
  tmv3_upsertTaskCacheRow_(readback);
  return {
    status: 'PATCHED_AND_READ_BACK',
    taskId: id,
    payload: safe,
    readback: readback
  };
}

function tmv3_operationAssignmentRequest_(taskId, assignment, method, scope) {
  tmv3_assertOperationWrite_(scope, 'TASK_DETAIL_ASSIGNMENT');

  if (tmv3_norm_(scope) === 'servicedeliverydetails') {
    const context = tmv3_taskFirstProductionCycle_.detailContext;
    const type = tmv3_norm_(assignment.type), id = Number(assignment.id);
    const desired = context && context.desiredAssignment || {};
    const allowedAdd = method === 'post' && (
      (type === 'employee' && (desired.employeeIds || []).indexOf(id)>=0) ||
      (type === 'pool' && (desired.poolIds || []).indexOf(id)>=0));
    const allowedRemove = method === 'delete' && type === 'pool' && id === 4 &&
      (desired.employeeIds || []).length > 0;
    if (!context || Number(context.taskId)!==Number(taskId) || (!allowedAdd && !allowedRemove)) {
      throw new Error('TASK_DETAILS_ASSIGNMENT_OUTSIDE_FRESH_CONTRACT');
    }
    if (allowedRemove) {
      const actual = tmv3_taskAssignmentsFromRaw_(tmv3_rawTaskById_(taskId));
      if (!(desired.employeeIds || []).every(function(employeeId){
        return actual.some(function(a){return a.type==='employee' && Number(a.id)===employeeId;});
      })) throw new Error('TASK_DETAILS_EMPLOYEE_NOT_READBACK_VERIFIED_NO_POOL_REMOVAL');
    }
  }

  return tmv3_fetchJson_(
    TMV3.API_BASE + '/v2/tasks/' + encodeURIComponent(Number(taskId)) + '/assignments',
    {
      method: method,
      contentType: 'application/json',
      payload: JSON.stringify({
        Id: Number(assignment.id),
        Name: assignment.name || '',
        Type: assignment.type || 'employee'
      })
    }
  );
}

function tmv3_reconcileTaskAssignments_(eventRecord, taskId, scope, options) {
  options = options || {};
  tmv3_assertOperationWrite_(scope, 'TASK_DETAIL_ASSIGNMENT');

  const desired = options.desiredAssignment || tmv3_desiredAssignment_(eventRecord);
  const raw = tmv3_rawTaskById_(taskId);
  let current = tmv3_taskAssignmentsFromRaw_(raw);
  const added = [];
  const removed = [];

  function has(type, id) {
    return current.some(function(item) {
      return item.type === type && Number(item.id) === Number(id);
    });
  }

  (desired.employeeIds || []).forEach(function(id) {
    if (has('employee', id)) return;
    const assignment = {
      id: Number(id),
      type: 'employee',
      name: tmv3_employeeNameById_(id)
    };
    tmv3_operationAssignmentRequest_(taskId, assignment, 'post', scope);
    added.push(assignment);
    current.push(assignment);
  });

  (desired.poolIds || []).forEach(function(id) {
    if (has('pool', id)) return;
    const cfg = TMV3.VERTICALS[eventRecord.vertical];
    const assignment = {
      id: Number(id),
      type: 'pool',
      name: cfg.defaultPoolName || ''
    };
    tmv3_operationAssignmentRequest_(taskId, assignment, 'post', scope);
    added.push(assignment);
    current.push(assignment);
  });

  // Original Delivery/Service rule: only remove To Be Assigned after the
  // intended employee has been successfully established.
  if (
    (eventRecord.vertical === 'Delivery' || eventRecord.vertical === 'Service') &&
    (desired.employeeIds || []).length &&
    has('pool', 4)
  ) {
    const assignment = { id: 4, type: 'pool', name: 'To Be Assigned' };
    tmv3_operationAssignmentRequest_(taskId, assignment, 'delete', scope);
    removed.push(assignment);
    current = current.filter(function(item) {
      return !(item.type === 'pool' && Number(item.id) === 4);
    });
  }

  // Legacy parity: Install / Delivery / Service employee assignment
  // writes are additive. Preserve existing employee assignments.
  // Delivery / Service remove only Pool 4 (To Be Assigned), and only
  // after the intended employee has been established above.

  if (eventRecord.vertical === 'PreInspection') {
    // Pool 8 is mandatory. New tasks must not inherit the historical legacy
    // employees 6/20 from old templates.
    if (options.newTask === true) {
      [6, 20].forEach(function(id) {
        const found = current.filter(function(item) {
          return item.type === 'employee' && Number(item.id) === Number(id);
        });
        found.forEach(function(item) {
          tmv3_operationAssignmentRequest_(taskId, item, 'delete', scope);
          removed.push(item);
        });
      });
    }

    if (has('pool', 4)) {
      const wrongPool = { id: 4, type: 'pool', name: 'To Be Assigned' };
      tmv3_operationAssignmentRequest_(taskId, wrongPool, 'delete', scope);
      removed.push(wrongPool);
    }
  }

  const afterRaw = tmv3_rawTaskById_(taskId);
  const after = tmv3_taskAssignmentsFromRaw_(afterRaw);
  tmv3_upsertTaskCacheRow_(tmv3_normalizeV2TaskModel_(afterRaw));

  (desired.employeeIds || []).forEach(function(id) {
    if (!after.some(function(item) {
      return item.type === 'employee' && Number(item.id) === Number(id);
    })) {
      throw new Error('Assignment read-back failed for employee ' + id + '.');
    }
  });

  (desired.poolIds || []).forEach(function(id) {
    if (!after.some(function(item) {
      return item.type === 'pool' && Number(item.id) === Number(id);
    })) {
      throw new Error('Assignment read-back failed for pool ' + id + '.');
    }
  });

  return {
    status: added.length || removed.length ? 'RECONCILED' : 'NOT_NEEDED',
    taskId: Number(taskId),
    desired: desired,
    added: added,
    removed: removed,
    after: after
  };
}

function tmv3_assertResolvedOwnership_(bundle) {
  const resolved = bundle.resolved;
  const refs = bundle.refs;
  const customerId = tmv3_clean_(resolved.customerId);
  const locationId = tmv3_clean_(resolved.locationId);
  const orderId = tmv3_clean_(resolved.orderId);
  const contactId = tmv3_clean_(resolved.contactId);

  if (!customerId) throw new Error('Resolved Customer ID is missing.');

  if (locationId) {
    const owned = (refs.locationsByCustomer[customerId] || []).some(function(row) {
      return tmv3_clean_(row['Location ID']) === locationId;
    });
    if (!owned) {
      throw new Error(
        'Location ' + locationId + ' is not proven to belong to Customer ' + customerId + '.'
      );
    }
  }

  if (orderId) {
    const order = refs.orderById[orderId];
    if (!order || tmv3_clean_(order['Customer ID']) !== customerId) {
      throw new Error(
        'Order / Work Order ' + orderId + ' is not proven to belong to Customer ' + customerId + '.'
      );
    }
  }

  if (contactId && bundle.eventRecord.vertical !== 'PreInspection') {
    const contact = tmv3_getContactById_(contactId, customerId);
    if (!contact.__ownershipVerified) {
      throw new Error(
        'Contact ' + contactId + ' ownership could not be proven for Customer ' + customerId + '.'
      );
    }
  }

  return true;
}

function tmv3_selectedDatePayload_(bundle) {
  return {
    Id: Number(bundle.resolved.taskId),
    StartDateTime: tmv3_strivenTaskDateTime_(bundle.eventRecord.start),
    DueDateTime: tmv3_strivenTaskDateTime_(bundle.eventRecord.end)
  };
}

function tmv3_selectedRelationshipPayload_(bundle) {
  tmv3_assertResolvedOwnership_(bundle);

  const resolved = bundle.resolved;
  const vertical = bundle.eventRecord.vertical;
  const cfg = TMV3.VERTICALS[vertical];
  const payload = { Id: Number(resolved.taskId) };

  if (resolved.locationId) {
    payload.Location = { Id: Number(resolved.locationId) };
  }

  if (cfg.orderRequired && resolved.orderId) {
    payload.SalesOrder = { Id: Number(resolved.orderId) };
  }

  if (vertical === 'PreInspection') {
    const requestedBy = tmv3_resolveOrganizerEmployee_(bundle.eventRecord);
    if (requestedBy.status !== 'MATCHED') {
      throw new Error(requestedBy.reason || 'Requested By employee is unresolved.');
    }
    payload.RequestedBy = {
      Id: Number(requestedBy.employee.id),
      Name: requestedBy.employee.name || '',
      Type: 'employee'
    };
  } else if (resolved.contactId) {
    payload.RequestedBy = {
      Id: Number(resolved.contactId),
      Type: 'contact'
    };
  }

  return payload;
}

function tmv3_verifyTaskPatch_(bundle, readback, mode) {
  const issues = [];
  const resolved = bundle.resolved;
  const eventRecord = bundle.eventRecord;
  const cfg = TMV3.VERTICALS[eventRecord.vertical];

  let scheduleVerification = null;
  if (mode === 'DATES' || mode === 'ALL') {
    scheduleVerification = tmv3_taskScheduleCheck_(
      eventRecord.vertical,
      eventRecord.start,
      eventRecord.end,
      resolved.taskId,
      readback
    );

    if (scheduleVerification.startCheck !== 'MATCH') {
      issues.push(
        scheduleVerification.startCheck === 'CANONICAL_READ_FAILED'
          ? 'Start date/time canonical v1 read-back unavailable'
          : 'Start date/time read-back mismatch'
      );
    }

    if (scheduleVerification.endCheck !== 'MATCH') {
      issues.push(
        scheduleVerification.endCheck === 'CANONICAL_READ_FAILED'
          ? 'Due date/time canonical v1 read-back unavailable'
          : (
              tmv3_taskDueDateOnly_(eventRecord.vertical, readback)
                ? 'Due date read-back mismatch'
                : 'Due date/time read-back mismatch'
            )
      );
    }
  }

  if (mode === 'RELATIONSHIPS' || mode === 'ALL') {
    if (
      resolved.locationId &&
      String(readback['Location ID'] || '') !== String(resolved.locationId)
    ) {
      issues.push('Location read-back mismatch');
    }

    if (
      cfg.orderRequired && resolved.orderId &&
      String(readback['Order ID'] || '') !== String(resolved.orderId)
    ) {
      issues.push(cfg.orderLabel + ' read-back mismatch');
    }

    if (
      eventRecord.vertical !== 'PreInspection' && resolved.contactId &&
      String(readback['Contact ID'] || '') !== String(resolved.contactId)
    ) {
      issues.push('Requested By / Contact read-back mismatch');
    }

    if (eventRecord.vertical === 'PreInspection' && readback['Order ID']) {
      issues.push('PreInspection Task unexpectedly has a Sales Order');
    }
  }

  if (issues.length) {
    throw new Error('Task PATCH verification failed: ' + issues.join('; '));
  }

  return {
    status: 'VERIFIED',
    taskId: Number(resolved.taskId),
    mode: mode,
    readback: readback,
    scheduleVerification: scheduleVerification
  };
}

function tmv3_stripAuthoredCalendarNotes_(description) {
  return tmv3_stripManagedLinkBlocks_(description || '')
    .replace(/<!--\s*PREINSPECT_STRIVEN_TASK_LINK_START\s*-->[\s\S]*?<!--\s*PREINSPECT_STRIVEN_TASK_LINK_END\s*-->/gi, '')
    .replace(/(?:\r?\n\s*){3,}/g, '\n\n')
    .trim();
}


/************************************************************
 * TM V3 — TITLE NORMALIZATION WRITE PRIMITIVES
 *
 * These helpers are deliberately narrow:
 * - Calendar mutation: title plus exact legacy-title prepend only.
 * - Task mutation: Title field only.
 * - No generic description cleanup is called.
 * - No caller is authorized unless the canonical Step 7 contract
 *   explicitly exposes the corresponding title action.
 ************************************************************/

function tmv3_titleDescriptionPlan_(currentTitle, desiredTitle, rawDescription, options) {
  options = options || {};
  const oldTitle = tmv3_clean_(currentTitle);
  const newTitle = tmv3_clean_(desiredTitle);
  const before = String(rawDescription === undefined || rawDescription === null ? '' : rawDescription);

  if (!oldTitle) {
    return {
      status:'BLOCKED',
      reason:'Current Calendar title is blank; legacy title cannot be preserved.',
      currentTitle:oldTitle,
      desiredTitle:newTitle,
      before:before,
      after:before,
      titleChange:false,
      descriptionChange:false,
      alreadyPreserved:false
    };
  }

  if (!newTitle) {
    return {
      status:'BLOCKED',
      reason:'Desired Calendar title is blank.',
      currentTitle:oldTitle,
      desiredTitle:newTitle,
      before:before,
      after:before,
      titleChange:false,
      descriptionChange:false,
      alreadyPreserved:false
    };
  }

  if (oldTitle === newTitle) {
    return {
      status:'NO_CHANGE',
      reason:'',
      currentTitle:oldTitle,
      desiredTitle:newTitle,
      before:before,
      after:before,
      titleChange:false,
      descriptionChange:false,
      alreadyPreserved:false
    };
  }

  const preserveOldTitle = options.preserveOldTitle !== false;
  const prefix = oldTitle + '\n\n';
  const alreadyPreserved =
    preserveOldTitle &&
    (
      before === oldTitle ||
      before.indexOf(prefix) === 0
    );
  const after =
    !preserveOldTitle
      ? before
      : (
          alreadyPreserved
            ? before
            : (
                before
                  ? prefix + before
                  : oldTitle
              )
        );

  return {
    status:'READY',
    reason:'',
    currentTitle:oldTitle,
    desiredTitle:newTitle,
    before:before,
    after:after,
    titleChange:true,
    descriptionChange:after !== before,
    alreadyPreserved:alreadyPreserved
  };
}

function tmv3_assertTitleNormalizationWriteAuthorized_(contract, action, scope) {
  tmv3_assertOperationWrite_(scope);
  tmv3_step7ValidateExecutionContract_(contract);

  if (tmv3_clean_(contract.blocker)) {
    throw new Error(
      'Title normalization blocked by Step 7: ' +
      contract.blocker
    );
  }

  if ((contract.actions || []).indexOf(action) === -1) {
    throw new Error(
      'Step 7 did not authorize title action ' + action + '.'
    );
  }

  if (
    action === 'PATCH_TASK_NAME' &&
    !tmv3_taskNameNormalizationStatusAllowed_(contract.taskStatus)
  ) {
    throw new Error(
      'Task Name normalization is authorized only for the canonical OPEN Task.'
    );
  }
}

function tmv3_titleNormalizationCalendarIds_(eventRecord) {
  return tmv3_requiredCalendarCopiesForEvent_(
    eventRecord || {}
  );
}

function tmv3_titleMissingRequiredCalendarIds_(requiredIds, copies) {
  const found = {};
  (copies || []).forEach(function(copy) {
    const id = tmv3_clean_(copy && copy.calendarId);
    if (id) found[id] = true;
  });

  return (requiredIds || []).map(tmv3_clean_).filter(function(id) {
    return id && !found[id];
  });
}

function tmv3_titleCalendarProtectedSnapshot_(event) {
  const guestEmails = [];
  const creatorEmails = [];

  try {
    (event.getGuestList ? event.getGuestList(true) : []).forEach(function(guest) {
      const email = tmv3_normEmail_(guest && guest.getEmail ? guest.getEmail() : '');
      if (email) guestEmails.push(email);
    });
  } catch (ignored) {}

  try {
    (event.getCreators ? event.getCreators() : []).forEach(function(email) {
      const clean = tmv3_normEmail_(email);
      if (clean) creatorEmails.push(clean);
    });
  } catch (ignored) {}

  return {
    start:event && event.getStartTime ? event.getStartTime().getTime() : null,
    end:event && event.getEndTime ? event.getEndTime().getTime() : null,
    allDay:event && event.isAllDayEvent ? event.isAllDayEvent() === true : null,
    location:String(event && event.getLocation ? event.getLocation() || '' : ''),
    guests:guestEmails.sort(),
    creators:creatorEmails.sort()
  };
}

function tmv3_titleTaskProtectedSnapshot_(task) {
  task = task || {};

  function sortedCsv(value) {
    return tmv3_clean_(value)
      .split(',')
      .map(tmv3_clean_)
      .filter(Boolean)
      .sort();
  }

  return {
    taskId:tmv3_clean_(task['Task ID']),
    taskNumber:tmv3_clean_(task['Task Number']),
    taskTypeId:tmv3_clean_(task['Task Type ID']),
    taskType:tmv3_clean_(task['Task Type']),
    status:tmv3_clean_(task['Status']),
    customerId:tmv3_clean_(task['Customer ID']),
    locationId:tmv3_clean_(task['Location ID']),
    contactId:tmv3_clean_(task['Contact ID']),
    orderId:tmv3_clean_(task['Order ID']),
    start:tmv3_clean_(task['Start']),
    due:tmv3_clean_(task['Due']),
    assignees:sortedCsv(task['Assignees']),
    pools:sortedCsv(task['Pools'])
  };
}

function tmv3_assertTitleProtectedSnapshotUnchanged_(before, after, label) {
  const beforeJson = JSON.stringify(before || {});
  const afterJson = JSON.stringify(after || {});

  if (beforeJson !== afterJson) {
    throw new Error(
      label + ' unrelated-field read-back mismatch. Before=' +
      beforeJson + ' After=' + afterJson
    );
  }

  return true;
}

function tmv3_operationNormalizeCalendarTitle_(bundle, contract, scope) {
  tmv3_assertTitleNormalizationWriteAuthorized_(
    contract,
    'PATCH_CALENDAR_TITLE',
    scope
  );

  const eventRecord = bundle && bundle.eventRecord;
  if (!eventRecord) throw new Error('Calendar title normalization requires an Event record.');

  const desired = tmv3_clean_(contract.desiredCalendarTitle);
  if (contract.vertical === 'PreInspection') tmv3_piAssertCalendarTitle_(desired);
  if (!desired) throw new Error('Step 7 contract has no desired Calendar title.');

  const ids = tmv3_titleNormalizationCalendarIds_(eventRecord);
  if (!ids.length) throw new Error('No Calendar ID is available for title normalization.');

  const copies = [];
  ids.forEach(function(calendarId) {
    const found = tmv3_findEventCopyRobust_(calendarId, contract.eventId);
    if (found && found.event) {
      copies.push({
        calendarId:calendarId,
        event:found.event,
        protectedBefore:tmv3_titleCalendarProtectedSnapshot_(found.event),
        beforeTitle:String(found.event.getTitle() || ''),
        beforeDescription:String(found.event.getDescription() || '')
      });
    }
  });

  const missing = tmv3_titleMissingRequiredCalendarIds_(ids, copies);
  if (missing.length) {
    throw new Error(
      'Required Calendar event copy/copies were not found: ' +
      missing.join(', ') +
      '. No Calendar title write was performed.'
    );
  }

  let authoritativePreInspectionPlan = null;
  if (eventRecord.vertical === 'PreInspection') {
    const cfg = TMV3.VERTICALS.PreInspection || {};
    const primaryId = tmv3_clean_(cfg.primaryCalendarId);
    const authoritativeCopy = copies.filter(function(copy) {
      return tmv3_clean_(copy.calendarId) === primaryId;
    })[0];

    if (!authoritativeCopy) {
      throw new Error(
        'Authoritative CF Preinspects copy is required before title normalization.'
      );
    }

    const recordedLegacy = tmv3_titleMigrationStateForEvent_(
      'PreInspection', contract.eventId, primaryId
    ).preservedTitle;
    authoritativePreInspectionPlan = tmv3_titleDescriptionPlan_(
      authoritativeCopy.beforeTitle,
      desired,
      authoritativeCopy.beforeDescription,
      { preserveOldTitle:!recordedLegacy }
    );
    const savedLegacy = recordedLegacy || (authoritativeCopy.beforeTitle !== desired ? authoritativeCopy.beforeTitle : '');
    authoritativePreInspectionPlan.after = tmv3_preInspectionEnsureLegacyTitle_(
      authoritativePreInspectionPlan.after, savedLegacy
    );

    if (authoritativePreInspectionPlan.status === 'BLOCKED') {
      throw new Error(
        'Authoritative Calendar title preflight blocked: ' +
        authoritativePreInspectionPlan.reason
      );
    }
  }

  copies.forEach(function(copy) {
    if (eventRecord.vertical === 'PreInspection') {
      copy.plan = {
        status:'READY',
        reason:'',
        currentTitle:copy.beforeTitle,
        desiredTitle:desired,
        before:copy.beforeDescription,
        after:authoritativePreInspectionPlan.after,
        titleChange:tmv3_clean_(copy.beforeTitle) !== desired,
        descriptionChange:
          copy.beforeDescription !== authoritativePreInspectionPlan.after,
        alreadyPreserved:
          copy.beforeDescription === authoritativePreInspectionPlan.after
      };
    } else {
      copy.plan = tmv3_titleDescriptionPlan_(
        copy.beforeTitle,
        desired,
        copy.beforeDescription,
        { preserveOldTitle:true }
      );
    }

    if (copy.plan.status === 'BLOCKED') {
      throw new Error(
        'Calendar title preflight blocked on ' +
        copy.calendarId + ': ' +
        copy.plan.reason
      );
    }
  });

  const results = [];

  copies.forEach(function(copy) {
    const plan = copy.plan;

    if (plan.descriptionChange) {
      copy.event.setDescription(plan.after);
    }

    if (plan.titleChange) {
      copy.event.setTitle(desired);
    }

    const readbackFound = tmv3_findEventCopyRobust_(
      copy.calendarId,
      contract.eventId
    );
    if (!readbackFound || !readbackFound.event) {
      throw new Error('Calendar event disappeared during title read-back.');
    }

    const afterTitle = String(readbackFound.event.getTitle() || '');
    const afterDescription = String(readbackFound.event.getDescription() || '');
    const protectedAfter =
      tmv3_titleCalendarProtectedSnapshot_(readbackFound.event);

    if (tmv3_clean_(afterTitle) !== desired) {
      throw new Error(
        'Calendar title read-back mismatch on ' + copy.calendarId + '.'
      );
    }

    if (afterDescription !== plan.after) {
      throw new Error(
        'Calendar description preservation read-back mismatch on ' +
        copy.calendarId + '.'
      );
    }

    tmv3_assertTitleProtectedSnapshotUnchanged_(
      copy.protectedBefore,
      protectedAfter,
      'Calendar ' + copy.calendarId
    );

    if (plan.titleChange || plan.descriptionChange) {
      tmv3_recordTitleMigrationState_(
        eventRecord.vertical,
        contract.eventId,
        copy.calendarId,
        plan.currentTitle,
        desired
      );
    }

    results.push({
      calendarId:copy.calendarId,
      status:
        !plan.titleChange && !plan.descriptionChange
          ? 'ALREADY_CORRECT'
          : 'WRITTEN_AND_VERIFIED',
      writePerformed:plan.titleChange || plan.descriptionChange,
      oldTitle:plan.currentTitle,
      newTitle:desired,
      descriptionChanged:plan.descriptionChange,
      oldDescriptionLength:copy.beforeDescription.length,
      newDescriptionLength:afterDescription.length,
      oldDescriptionHash:tmv3_hash_(copy.beforeDescription),
      newDescriptionHash:tmv3_hash_(afterDescription),
      unrelatedFieldsVerified:true
    });
  });

  const parity =
    eventRecord.vertical === 'PreInspection'
      ? tmv3_reconcilePreInspectionCalendarParity_(
          eventRecord,
          scope
        )
      : null;

  return {
    status:'CALENDAR_TITLE_VERIFIED',
    vertical:eventRecord.vertical,
    eventId:contract.eventId,
    desiredCalendarTitle:desired,
    requiredCalendarIds:ids,
    copies:results,
    parity:parity
  };
}

function tmv3_operationNormalizeTaskName_(taskId, desiredTaskName, contract, scope) {
  tmv3_assertTitleNormalizationWriteAuthorized_(
    contract,
    'PATCH_TASK_NAME',
    scope
  );

  const id = Number(taskId || 0);
  const desired = tmv3_clean_(desiredTaskName);
  if (contract && contract.vertical === 'PreInspection') tmv3_piAssertTaskTitle_(desired);

  if (!id || !desired) {
    throw new Error('Task Name normalization requires Task ID and desired Task Name.');
  }

  const before = tmv3_getTaskById_(id);
  const beforeName = tmv3_clean_(before && before['Name']);
  const protectedBefore = tmv3_titleTaskProtectedSnapshot_(before);

  if (beforeName === desired) {
    return {
      status:'ALREADY_CORRECT',
      taskId:id,
      beforeName:beforeName,
      afterName:beforeName,
      unrelatedFieldsVerified:true
    };
  }

  const patch = tmv3_operationPatchTask_(
    id,
    { Title:desired },
    scope
  );

  const after = tmv3_getTaskById_(id);
  const afterName = tmv3_clean_(after && after['Name']);
  const protectedAfter = tmv3_titleTaskProtectedSnapshot_(after);

  if (afterName !== desired) {
    throw new Error('Task Name read-back mismatch for Task ' + id + '.');
  }

  tmv3_assertTitleProtectedSnapshotUnchanged_(
    protectedBefore,
    protectedAfter,
    'Task ' + id
  );

  return {
    status:'TASK_NAME_WRITTEN_AND_VERIFIED',
    taskId:id,
    beforeName:beforeName,
    afterName:afterName,
    patchStatus:patch && patch.status,
    unrelatedFieldsVerified:true
  };
}

function tmv3_infoCustomFieldsRaw_(raw) {
  const direct = raw && (raw.infoCustomFields || raw.InfoCustomFields);
  return Array.isArray(direct) ? direct : [];
}

function tmv3_fieldId_(field) {
  return Number(tmv3_first_(field || {}, ['id','Id','customFieldId','CustomFieldId']) || 0);
}

function tmv3_fieldValue_(field) {
  if (!field) return null;
  if (field.value !== undefined) return field.value;
  if (field.Value !== undefined) return field.Value;
  if (field.valueText !== undefined) return field.valueText;
  if (field.ValueText !== undefined) return field.ValueText;
  return null;
}

function tmv3_setExistingFieldValue_(field, value) {
  if (Object.prototype.hasOwnProperty.call(field, 'value')) field.value = value;
  else if (Object.prototype.hasOwnProperty.call(field, 'Value')) field.Value = value;
  else field.Value = value;

  if (Object.prototype.hasOwnProperty.call(field, 'valueText')) field.valueText = null;
  if (Object.prototype.hasOwnProperty.call(field, 'ValueText')) field.ValueText = null;
}

function tmv3_findEventCopyRobust_(calendarId, eventId) {
  const direct = tmv3_findCalendarEventById_(calendarId, eventId);
  if (direct) return direct;

  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) return null;

  const raw = tmv3_clean_(eventId);
  const normalized = raw.replace(/@google\.com$/i, '');
  const start = new Date();
  start.setDate(start.getDate() - 3);
  start.setHours(0,0,0,0);
  const end = new Date();
  end.setDate(end.getDate() + 365);
  end.setHours(23,59,59,999);
  const events = calendar.getEvents(start, end);

  for (let i = 0; i < events.length; i++) {
    const foundRaw = tmv3_clean_(events[i].getId());
    const foundNormalized = foundRaw.replace(/@google\.com$/i, '');
    if (
      foundRaw === raw ||
      foundRaw === normalized ||
      foundNormalized === raw ||
      foundNormalized === normalized
    ) {
      return { calendar: calendar, event: events[i], match: 'BOUNDED_SCAN' };
    }
  }

  return null;
}










function tmv3_operationWriteCalendarLinks_(bundle, scope) {
  tmv3_assertOperationWrite_(scope);

  const eventRecord = bundle.eventRecord;
  const mirror = tmv3_ensurePreInspectionMirror_(
    eventRecord,
    scope
  );
  const records = (bundle.records || [bundle.resolved]).map(function(record) {
    return Object.assign({}, record || {});
  });

  records.forEach(function(record) {
    tmv3_assertResolvedRecordWritable_(record, { requireTask: true });
  });

  if (
    eventRecord.vertical === 'PreInspection' &&
    bundle.contract &&
    tmv3_clean_(bundle.contract.desiredTaskName)
  ) {
    const expectedTaskId = tmv3_clean_(bundle.contract.taskId);
    records.forEach(function(record) {
      const recordTaskId = tmv3_clean_(record.taskId);
      if (
        !expectedTaskId ||
        recordTaskId === expectedTaskId
      ) {
        record.task = tmv3_clean_(bundle.contract.desiredTaskName);
      }
    });
  }

  let linkEventRecord = eventRecord;

  if (eventRecord.vertical === 'PreInspection') {
    const customer = tmv3_customerFromRefs_(
      bundle.refs,
      bundle.resolved.customerId
    );
    const location = tmv3_locationFromRefs_(
      bundle.refs,
      bundle.resolved.customerId,
      bundle.resolved.locationId
    );

    if (!customer || !location) {
      throw new Error(
        'PreInspection Calendar link write requires verified Customer and Customer-owned Location.'
      );
    }

    linkEventRecord = Object.assign({}, eventRecord, {
      customerNumber:
        tmv3_clean_(eventRecord.customerNumber) ||
        tmv3_clean_(customer['Customer Number']),
      phone:
        tmv3_clean_(eventRecord.phone) ||
        tmv3_clean_(customer['Primary Phone']) ||
        tmv3_clean_(location['Phone']),
      step4:{
        disposition:'VERIFIED',
        customer:customer,
        location:location
      }
    });
  }

  const plan = tmv3_buildCalendarLinkPlan_(linkEventRecord, records);
  const copies = [];

  (plan.calendarIds || []).forEach(function(calendarId) {
    const found = tmv3_findEventCopyRobust_(calendarId, plan.eventId);
    if (found && found.event) {
      copies.push({
        calendarId: calendarId,
        calendarName: found.calendar.getName(),
        event: found.event
      });
    }
  });

  if (!copies.length) {
    throw new Error('No Calendar copy found for Event ID ' + plan.eventId + '.');
  }

  const missingCalendarCopies = (plan.calendarIds || []).filter(function(calendarId) {
    return !copies.some(function(copy) {
      return tmv3_clean_(copy.calendarId) === tmv3_clean_(calendarId);
    });
  });

  if (
    eventRecord.vertical === 'PreInspection' &&
    missingCalendarCopies.length
  ) {
    throw new Error(
      'Required PreInspection Calendar copy/copies were not found: ' +
      missingCalendarCopies.join(', ') +
      '. No Calendar link write was performed.'
    );
  }

  // Non-PreInspection verticals may legitimately exist on only one configured source
  // Calendar. Patch every copy that actually exists and verify it.
  const results = [];
  copies.forEach(function(copy) {
    const before = String(copy.event.getDescription() || '');
    const savedLegacy = eventRecord.vertical === 'PreInspection'
      ? tmv3_titleMigrationStateForEvent_('PreInspection', plan.eventId,
          TMV3.VERTICALS.PreInspection.primaryCalendarId).preservedTitle
      : '';
    const preservedBefore = eventRecord.vertical === 'PreInspection'
      ? tmv3_preInspectionEnsureLegacyTitle_(before, savedLegacy)
      : before;

    const linkUpdate =
      eventRecord.vertical === 'PreInspection'
        ? tmv3_preInspectionClickableManagedDescription_(
            preservedBefore,
            plan
          )
        : {
            description:tmv3_managedCalendarDescription_(
              before,
              plan
            ),
            status:'STANDARD_MANAGED_BLOCK',
            changed:false
          };

    const desired = eventRecord.vertical === 'PreInspection'
      ? tmv3_preInspectionEnsureLegacyTitle_(String(linkUpdate.description || ''), savedLegacy)
      : String(linkUpdate.description || '');
    if (before !== desired) {
      copy.event.setDescription(desired);
    }

    const readBackFound = tmv3_findEventCopyRobust_(copy.calendarId, plan.eventId);
    if (!readBackFound || !readBackFound.event) {
      throw new Error('Calendar copy disappeared during link read-back.');
    }

    const after = String(readBackFound.event.getDescription() || '');
    if (eventRecord.vertical === 'PreInspection') tmv3_assertPreInspectionLinkDescription_(after, plan);
    const missing = (plan.links || []).filter(function(link) {
      if (eventRecord.vertical === 'PreInspection') {
        return !tmv3_preInspectionClickableLinkPresent_(
          after,
          link
        );
      }
      return after.indexOf(link.url) === -1;
    });

    if (missing.length) {
      throw new Error(
        'Calendar link read-back failed on ' + copy.calendarName +
        ': missing clickable ' +
        missing.map(function(item) { return item.key; }).join(', ')
      );
    }

    results.push({
      calendarId: copy.calendarId,
      calendarName: copy.calendarName,
      status:
        before === desired
          ? (
              eventRecord.vertical === 'PreInspection'
                ? 'CLICKABLE_LINK_BLOCK_PRESERVED'
                : 'ALREADY_CORRECT'
            )
          : (
              eventRecord.vertical === 'PreInspection'
                ? 'CLICKABLE_LINK_BLOCK_REPAIRED_AND_VERIFIED'
                : 'WRITTEN_AND_VERIFIED'
            )
    });
  });

  const parity =
    eventRecord.vertical === 'PreInspection'
      ? tmv3_reconcilePreInspectionCalendarParity_(
          eventRecord,
          scope
        )
      : null;

  if(eventRecord.vertical==='PreInspection') copies.forEach(function(copy){
    const finalCopy=tmv3_findEventCopyRobust_(copy.calendarId,plan.eventId);
    if(!finalCopy||!finalCopy.event) throw new Error('PREINSPECTION_FINAL_COPY_MISSING');
    tmv3_assertPreInspectionLinkDescription_(finalCopy.event.getDescription(), plan);
  });

  return {
    status: 'CALENDAR_LINKS_VERIFIED',
    vertical: eventRecord.vertical,
    eventId: plan.eventId,
    copies: results,
    links: plan.links,
    missingCalendarCopies: missingCalendarCopies,
    attention: missingCalendarCopies.length
      ? 'Some actual Calendar copies were not present; existing copies were updated and verified.'
      : '',
    mirror: mirror,
    parity: parity
  };
}

function tmv3_replacementSourceSnapshot_(taskId) {
  const raw = tmv3_rawTaskById_(taskId);
  const assignments = tmv3_taskAssignmentsFromRaw_(raw);
  const type = raw.type || raw.Type || null;
  const priority = raw.priority || raw.Priority || null;
  const customer = raw.customer || raw.Customer || null;
  const location = raw.location || raw.Location || null;
  const salesOrder = raw.salesOrder || raw.SalesOrder || null;
  const requestedBy = raw.requestedBy || raw.RequestedBy || null;

  return {
    sourceTaskId: Number(raw.id || raw.Id || taskId),
    status: tmv3_clean_((raw.status || raw.Status || {}).name || (raw.status || raw.Status || {}).Name),
    title: tmv3_clean_(raw.title || raw.Title || raw.name || raw.Name),
    description: String(raw.description || raw.Description || ''),
    budget: Number(raw.budget !== undefined ? raw.budget : (raw.Budget || 0)),
    type: type ? {
      id: Number(type.id || type.Id || 0),
      name: tmv3_clean_(type.name || type.Name)
    } : null,
    priority: priority ? {
      id: Number(priority.id || priority.Id || 0),
      name: tmv3_clean_(priority.name || priority.Name)
    } : null,
    customer: customer ? {
      id: Number(customer.id || customer.Id || 0),
      number: tmv3_clean_(customer.number || customer.Number),
      name: tmv3_clean_(customer.name || customer.Name)
    } : null,
    location: location ? {
      id: Number(location.id || location.Id || 0),
      name: tmv3_clean_(location.name || location.Name)
    } : null,
    salesOrder: salesOrder ? {
      id: Number(salesOrder.id || salesOrder.Id || 0),
      number: tmv3_clean_(salesOrder.number || salesOrder.Number),
      name: tmv3_clean_(salesOrder.name || salesOrder.Name)
    } : null,
    requestedBy: requestedBy ? {
      id: Number(requestedBy.id || requestedBy.Id || 0),
      name: tmv3_clean_(requestedBy.name || requestedBy.Name),
      type: tmv3_clean_(requestedBy.type || requestedBy.Type) || 'contact'
    } : null,
    startDateTime: raw.startDateTime || raw.StartDateTime || '',
    dueDateTime: raw.dueDateTime || raw.DueDateTime || '',
    employeeAssignmentIds: assignments.filter(function(a) {
      return a.type === 'employee';
    }).map(function(a) { return Number(a.id); }),
    poolAssignmentIds: assignments.filter(function(a) {
      return a.type === 'pool';
    }).map(function(a) { return Number(a.id); }),
    infoCustomFields: JSON.parse(JSON.stringify(tmv3_infoCustomFieldsRaw_(raw))),
    useSubContractor: raw.useSubContractor === true || raw.UseSubContractor === true,
    rawTask: raw
  };
}

function tmv3_safeCreateInfoCustomFields_(sourceFields, calendarNotes) {
  const output = [];

  (sourceFields || []).forEach(function(field) {
    const id = tmv3_fieldId_(field);
    const name = tmv3_clean_(tmv3_first_(field, ['name','Name']));
    if (!id || id === 834 || id === 732) return;
    if (/\bstatus\b|\bcompleted?\b|\bcompletion\b|\bdone\b|\boutcome\b|\bresult\b/i.test(name)) return;

    const value = tmv3_fieldValue_(field);
    if (value === null || value === undefined || value === '') return;
    output.push(JSON.parse(JSON.stringify(field)));
  });

  const notes = String(calendarNotes || '').trim();
  if (notes) {
    output.push({
      Id: 834,
      Name: 'Notes from Google Calendar',
      Value: notes,
      ValueText: null
    });
  }

  return output;
}

function tmv3_buildReplacementCreatePayload_(source, request) {
  if (!source || !source.type || !source.type.id) {
    throw new Error('CREATE source has no Task Type.');
  }
  if (!source.title) {
    throw new Error('CREATE source has no Task title.');
  }

  const payload = {
    Title: source.title,
    Type: {
      Id: Number(source.type.id),
      Name: source.type.name || ''
    },
    Description: source.description || '',
    Budget: Number(source.budget || 0),
    StartDateTime: request.startDateTime,
    DueDateTime: request.dueDateTime,
    UseSubContractor: source.useSubContractor === true
  };

  if (source.priority && source.priority.id) {
    payload.Priority = {
      Id: Number(source.priority.id),
      Name: source.priority.name || ''
    };
  }

  if (source.customer && source.customer.id) {
    payload.Customer = {
      Id: Number(source.customer.id),
      Number: source.customer.number || '',
      Name: source.customer.name || ''
    };
  }

  if (source.location && source.location.id) {
    payload.Location = {
      Id: Number(source.location.id),
      Name: source.location.name || ''
    };
  }

  if (source.salesOrder && source.salesOrder.id) {
    payload.SalesOrder = {
      Id: Number(source.salesOrder.id),
      Number: source.salesOrder.number || '',
      Name: source.salesOrder.name || ''
    };
  }

  if (source.requestedBy && source.requestedBy.id) {
    payload.RequestedBy = {
      Id: Number(source.requestedBy.id),
      Name: source.requestedBy.name || '',
      Type: source.requestedBy.type || 'contact'
    };
  }

  const fields = tmv3_safeCreateInfoCustomFields_(
    source.infoCustomFields || [],
    request.calendarNotes
  );
  if (fields.length) payload.InfoCustomFields = fields;

  return payload;
}

function tmv3_extractCreatedTaskId_(json) {
  if (typeof json === 'number' && json > 0) return Number(json);
  if (typeof json === 'string' && /^\d+$/.test(json.trim())) return Number(json.trim());
  if (!json || typeof json !== 'object') return null;

  const candidates = [
    json.id, json.Id, json.taskId, json.TaskId, json.TaskID,
    json.data && json.data.id,
    json.data && json.data.Id,
    json.task && json.task.id,
    json.task && json.task.Id
  ];

  for (let i = 0; i < candidates.length; i++) {
    const id = Number(candidates[i]);
    if (Number.isFinite(id) && id > 0) return id;
  }

  return null;
}

function tmv3_customerFromRefs_(refs, customerId) {
  return refs.customerById[String(customerId)] || null;
}

function tmv3_locationFromRefs_(refs, customerId, locationId) {
  const matches = (refs.locationsByCustomer[String(customerId)] || []).filter(function(row) {
    return String(row['Location ID'] || '') === String(locationId || '');
  });
  return matches.length === 1 ? matches[0] : null;
}

function tmv3_createTaskTitle_(bundle, action, source) {
  const eventRecord = bundle && bundle.eventRecord;
  const resolved = bundle && bundle.resolved || {};
  const refs = bundle && bundle.refs || {};
  const contract = bundle && bundle.contract || {};

  if (!eventRecord) {
    throw new Error('Task title construction requires an Event record.');
  }

  // CREATE execution starts from the raw Calendar event, while the canonical
  // title builder intentionally requires Step 3 + Step 4 verified identity.
  // Rehydrate only that verified title context from the already-authorized
  // Step 7 relationship. Do not infer or guess identity from Calendar text.
  tmv3_assertResolvedOwnership_(bundle);

  if (
    tmv3_clean_(contract.expectedCustomerId) &&
    tmv3_clean_(contract.expectedCustomerId) !==
      tmv3_clean_(resolved.customerId)
  ) {
    throw new Error(
      'CREATE title context Customer no longer matches the Step 7 contract.'
    );
  }

  if (
    tmv3_clean_(contract.expectedLocationId) &&
    tmv3_clean_(contract.expectedLocationId) !==
      tmv3_clean_(resolved.locationId)
  ) {
    throw new Error(
      'CREATE title context Location no longer matches the Step 7 contract.'
    );
  }

  const customer = tmv3_customerFromRefs_(
    refs,
    resolved.customerId
  );
  const location = tmv3_locationFromRefs_(
    refs,
    resolved.customerId,
    resolved.locationId
  );

  if (!customer || !location) {
    throw new Error(
      'CREATE title context requires the verified Customer and customer-owned Location.'
    );
  }

  const verifiedRecord = Object.assign(
    {},
    eventRecord,
    {
      step3:{
        disposition:'VERIFIED',
        anchor:{
          customerNumber:tmv3_clean_(customer['Customer Number']),
          orderNumber:tmv3_clean_(eventRecord.orderNumber)
        }
      },
      step4:{
        disposition:'VERIFIED',
        customer:customer,
        location:location
      }
    }
  );

  const desired = tmv3_desiredTaskName_(
    verifiedRecord,
    {
      sourceTaskTitle:
        source && source.title
          ? source.title
          : ''
    }
  );

  if (desired.status !== 'READY' || !tmv3_clean_(desired.value)) {
    throw new Error(
      desired.reason ||
      'Canonical Task Name could not be built from verified evidence.'
    );
  }

  return desired.value;
}



function tmv3_findCreateTemplateTask_(bundle) {
  const vertical = bundle.eventRecord.vertical;
  const cfg = TMV3.VERTICALS[vertical];

  if (vertical === 'PreInspection') {
    const source = tmv3_replacementSourceSnapshot_(11138);
    if (!source.type || Number(source.type.id) !== 105) {
      throw new Error('PreInspection Template Task 11138 is no longer Task Type 105.');
    }
    return source;
  }

  let desiredType = '';
  if (vertical === 'Delivery') desiredType = 'BBQ Delivery';
  if (vertical === 'Service') {
    desiredType = tmv3_serviceTaskTypeForOrder_(bundle.resolved.orderId);
    if (!desiredType) {
      throw new Error(
        'Service CREATE blocked: Work Order Category / Task Type could not be proven.'
      );
    }
  }

  const candidates = (bundle.refs.tasks || []).filter(function(task) {
    if (!tmv3_taskFitsVertical_(task, vertical, cfg)) return false;
    if (
      desiredType &&
      tmv3_norm_(task['Task Type']) !== tmv3_norm_(desiredType)
    ) return false;
    return !!tmv3_clean_(task['Task ID']);
  });

  candidates.sort(function(a, b) {
    const ao = tmv3_taskIsOpen_(a['Status']) ? 1 : 0;
    const bo = tmv3_taskIsOpen_(b['Status']) ? 1 : 0;
    if (ao !== bo) return bo - ao;
    return Number(a['Task ID'] || 0) - Number(b['Task ID'] || 0);
  });

  if (!candidates.length) {
    throw new Error(
      vertical + ' CREATE blocked: no proven structural Task template is available.'
    );
  }

  return tmv3_replacementSourceSnapshot_(candidates[0]['Task ID']);
}

function tmv3_findRecreateSourceTask_(bundle) {
  const resolved = bundle.resolved;
  const vertical = bundle.eventRecord.vertical;
  const cfg = TMV3.VERTICALS[vertical];
  let candidates = [];

  (bundle.refs.tasks || []).forEach(function(task) {
    if (!tmv3_taskFitsVertical_(task, vertical, cfg)) return;
    if (!tmv3_step6IsFulfilledStatus_(task['Status'])) return;

    if (
      resolved.orderId &&
      String(task['Order ID'] || '') === String(resolved.orderId)
    ) {
      candidates.push(task);
      return;
    }

    if (
      !resolved.orderId &&
      resolved.customerId &&
      String(task['Customer ID'] || '') === String(resolved.customerId) &&
      String(task['Location ID'] || '') === String(resolved.locationId || '')
    ) {
      candidates.push(task);
    }
  });

  // PreInspection tasks are intentionally resolved on demand and are not
  // guaranteed to exist in the shared Source Tasks cache. For RECREATE, use
  // the same Customer-scoped Type 105 search as Step 5 so completed history
  // can be proven without weakening the duplicate guard.
  if (vertical === 'PreInspection' && candidates.length === 0) {
    const customer = tmv3_customerFromRefs_(bundle.refs, resolved.customerId);
    if (!customer) {
      throw new Error(
        'RECREATE requires verified PreInspection Customer before history lookup.'
      );
    }

    const history = tmv3_searchPreInspectionTasks_(customer);
    if (history.some(function(task){return tmv3_taskIsOpen_(task['Status']);})) {
      throw new Error('RECREATE blocked: live PreInspection history contains an open Task.');
    }
    candidates = history.filter(function(task) {
      return (
        tmv3_step6IsFulfilledStatus_(task['Status']) &&
        String(task['Customer ID'] || '') === String(resolved.customerId || '') &&
        String(task['Location ID'] || '') === String(resolved.locationId || '')
      );
    });
    // A single fulfilled customer-owned task at another location is history,
    // not a structural template for the new verified job site. The caller's
    // fresh Step 6/7 lifecycle and durable duplicate guards still authorize
    // recreation; seed the new Type 105 task from template 11138 instead.
    if (!candidates.length && history.length===1 &&
        tmv3_step6IsFulfilledStatus_(history[0]['Status']) &&
        String(history[0]['Customer ID'] || '')===String(resolved.customerId || '') &&
        tmv3_clean_(history[0]['Location ID']) &&
        String(history[0]['Location ID'])!==String(resolved.locationId || '') &&
        Number(history[0]['Task Type ID'])===105) {
      return tmv3_findCreateTemplateTask_(bundle);
    }
  }

  if (candidates.length !== 1) {
    throw new Error(
      'RECREATE requires exactly one completed source Task; found ' +
      candidates.length + '.'
    );
  }

  return tmv3_replacementSourceSnapshot_(candidates[0]['Task ID']);
}

function tmv3_buildCreateSource_(bundle, action) {
  tmv3_assertResolvedOwnership_(bundle);

  const resolved = bundle.resolved;
  const eventRecord = bundle.eventRecord;
  const customer = tmv3_customerFromRefs_(bundle.refs, resolved.customerId);
  const location = tmv3_locationFromRefs_(bundle.refs, resolved.customerId, resolved.locationId);
  const order = resolved.orderId ? bundle.refs.orderById[String(resolved.orderId)] : null;

  if (!customer || !location) {
    throw new Error('CREATE requires verified Customer and customer-owned Location.');
  }

  const base = action === 'RECREATE'
    ? tmv3_findRecreateSourceTask_(bundle)
    : tmv3_findCreateTemplateTask_(bundle);

  let requestedBy = null;
  if (eventRecord.vertical === 'PreInspection') {
    const employee = tmv3_resolveOrganizerEmployee_(eventRecord);
    if (employee.status !== 'MATCHED') {
      throw new Error(employee.reason || 'PreInspection Requested By unresolved.');
    }
    requestedBy = {
      id: Number(employee.employee.id),
      name: employee.employee.name || '',
      type: 'employee'
    };
  } else if (resolved.contactId) {
    requestedBy = {
      id: Number(resolved.contactId),
      name: '',
      type: 'contact'
    };
  }

  return {
    sourceTaskId: base.sourceTaskId,
    title: tmv3_createTaskTitle_(bundle, action, base),
    description:
      eventRecord.vertical === 'PreInspection'
        ? ''
        : (action === 'RECREATE' ? base.description : tmv3_stripAuthoredCalendarNotes_(eventRecord.description)),
    budget: Number(base.budget || 0),
    type: base.type,
    priority: base.priority,
    customer: {
      id: Number(resolved.customerId),
      number: tmv3_clean_(customer['Customer Number']),
      name: tmv3_clean_(customer['Name'])
    },
    location: {
      id: Number(resolved.locationId),
      name: [location['Address 1'], location['City']].map(tmv3_clean_).filter(Boolean).join(', ')
    },
    salesOrder:
      eventRecord.vertical === 'PreInspection'
        ? null
        : (order ? {
            id: Number(resolved.orderId),
            number: tmv3_clean_(order['Order Number']),
            name: tmv3_clean_(order['Name'])
          } : null),
    requestedBy: requestedBy,
    useSubContractor: base.useSubContractor === true,
    infoCustomFields:
      eventRecord.vertical === 'PreInspection'
        ? []
        : (action === 'RECREATE' ? base.infoCustomFields : [])
  };
}






function tmv3_durableWriteGuardKey_(kind, identity) {
  return (
    'TMV3_' +
    tmv3_clean_(kind).toUpperCase().replace(/[^A-Z0-9_]/g, '_') +
    '_' +
    tmv3_hash_(tmv3_clean_(identity))
  );
}

function tmv3_readDurableWriteGuard_(key) {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return null;

  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object'
      ? parsed
      : { state:'CORRUPT', raw:raw };
  } catch (err) {
    return {
      state:'CORRUPT',
      raw:raw,
      error:String(err && err.message || err)
    };
  }
}

function tmv3_writeDurableWriteGuard_(key, state, data) {
  const value = Object.assign(
    {},
    data || {},
    {
      state:tmv3_clean_(state),
      guardKey:key,
      engineVersion:TMV3.VERSION,
      updatedAt:tmv3_now_()
    }
  );

  PropertiesService.getScriptProperties().setProperty(
    key,
    JSON.stringify(value)
  );

  return value;
}

function tmv3_createGuardStateBlocksRetry_(guard) {
  if (!guard) return false;
  return tmv3_clean_(guard.state).toUpperCase() !== 'CLEARED';
}

function tmv3_taskCreateGuardKey_(contract) {
  const operation =
    (contract.actions || []).indexOf(
      TMV3_STEP7_ACTION.RECREATE_TASK
    ) !== -1
      ? 'RECREATE_TASK'
      : 'CREATE_TASK';

  return tmv3_durableWriteGuardKey_(
    'TASK_CREATE_GUARD',
    [
      contract.vertical,
      contract.eventId,
      operation,
      (contract.sourceTaskIdList || []).slice().sort().join(',')
    ].join('|')
  );
}

function tmv3_locationCreateGuardKey_(contract) {
  return tmv3_durableWriteGuardKey_(
    'LOCATION_CREATE_GUARD',
    [
      contract.vertical,
      contract.eventId,
      contract.expectedCustomerId,
      tmv3_normalizeAddress_(contract.expectedLocationAddress)
    ].join('|')
  );
}

function tmv3_parseCanadianCustomerLocation_(address, customerId) {
  const raw = tmv3_clean_(address);
  const customer = Number(customerId || 0);

  if (!raw || !customer) {
    throw new Error(
      'Customer Location creation requires verified Customer ID and Calendar address.'
    );
  }

  const upper = raw.toUpperCase();
  const postalMatch = upper.match(/\b([A-Z]\d[A-Z])\s?(\d[A-Z]\d)\b/);
  const provinceMatch = upper.match(
    /\b(ON|QC|BC|AB|MB|SK|NS|NB|NL|PE|NT|NU|YT)\b/
  );

  const parts = raw
    .split(',')
    .map(tmv3_clean_)
    .filter(Boolean);

  let provinceIndex = -1;

  for (let i = 0; i < parts.length; i++) {
    if (
      /\b(ON|QC|BC|AB|MB|SK|NS|NB|NL|PE|NT|NU|YT)\b/i.test(parts[i]) ||
      /\b[A-Z]\d[A-Z]\s?\d[A-Z]\d\b/i.test(parts[i])
    ) {
      provinceIndex = i;
      break;
    }
  }

  if (
    !postalMatch ||
    !provinceMatch ||
    provinceIndex < 2
  ) {
    throw new Error(
      'Customer Location creation requires a complete Canadian address with street, city, province and postal code. Address was: ' +
      raw
    );
  }

  const street = tmv3_clean_(parts[0]);
  const city = tmv3_clean_(parts[provinceIndex - 1]);

  if (!street || !city) {
    throw new Error(
      'Customer Location creation could not deterministically separate street and city from: ' +
      raw
    );
  }

  return {
    Id:0,
    Name:'Primary',
    IsPrimary:false,
    Address:{
      Address1:street,
      City:city,
      State:provinceMatch[1],
      PostalCode:postalMatch[1] + ' ' + postalMatch[2],
      Country:'CA'
    }
  };
}

function tmv3_normalizedLocationObject_(normalized) {
  const row = normalized || [];

  return {
    'Location ID':tmv3_clean_(row[0]),
    'Customer ID':tmv3_clean_(row[1]),
    'Address 1':tmv3_clean_(row[2]),
    'Address 2':tmv3_clean_(row[3]),
    'City':tmv3_clean_(row[4]),
    'Province':tmv3_clean_(row[5]),
    'Postal Code':tmv3_clean_(row[6]),
    'Phone':tmv3_clean_(row[7]),
    'Fingerprint':tmv3_clean_(row[8])
  };
}

function tmv3_freshCustomerLocationCandidates_(bundle, customerId, address) {
  const target = tmv3_addressParts_(address);
  const exactKey = tmv3_normalizeAddress_(address);
  const owned = tmv3_preInspectionCustomerLocationsFresh_(customerId);

  const exact = owned.filter(function(location) {
    return (
      tmv3_normalizeAddress_(tmv3_locationFullAddress_(location)) ===
      exactKey
    );
  });

  const matches = exact.length
    ? exact
    : owned.filter(function(location) {
        return tmv3_addressStrongMatch_(
          target,
          tmv3_addressParts_(
            tmv3_locationFullAddress_(location)
          )
        );
      });

  return {
    owned:owned,
    matches:matches
  };
}

function tmv3_extractCreatedLocationId_(json) {
  if (typeof json === 'number' && json > 0) {
    return Number(json);
  }

  if (typeof json === 'string' && /^\d+$/.test(json.trim())) {
    return Number(json.trim());
  }

  if (!json || typeof json !== 'object') {
    return 0;
  }

  const candidates = [
    json.id,
    json.Id,
    json.locationId,
    json.LocationId,
    json.LocationID,
    json.customerLocationId,
    json.CustomerLocationId,
    json.data && json.data.id,
    json.data && json.data.Id,
    json.location && json.location.id,
    json.location && json.location.Id
  ];

  for (let i = 0; i < candidates.length; i++) {
    const id = Number(candidates[i]);

    if (Number.isFinite(id) && id > 0) {
      return id;
    }
  }

  return 0;
}

function tmv3_upsertLocationCacheObject_(location) {
  const id = tmv3_clean_(location && location['Location ID']);

  if (!id) {
    throw new Error('Location cache upsert requires Location ID.');
  }

  const sh = tmv3_sheet_(TMV3.SHEETS.LOCATIONS);
  const headers = [
    'Location ID','Customer ID','Address 1','Address 2','City',
    'Province','Postal Code','Phone','Fingerprint'
  ];
  const row = headers.map(function(header) {
    return location[header] === undefined
      ? ''
      : location[header];
  });

  const lastRow = sh.getLastRow();
  let targetRow = 0;

  if (lastRow > 1) {
    const ids = sh.getRange(
      2,
      1,
      lastRow - 1,
      1
    ).getValues();

    for (let i = 0; i < ids.length; i++) {
      if (tmv3_clean_(ids[i][0]) === id) {
        targetRow = i + 2;
        break;
      }
    }
  }

  if (targetRow) {
    sh.getRange(
      targetRow,
      1,
      1,
      row.length
    ).setValues([row]);
  } else {
    sh.getRange(
      Math.max(2, lastRow + 1),
      1,
      1,
      row.length
    ).setValues([row]);
  }

  return location;
}

function tmv3_locationResultFromVerified_(
  bundle,
  contract,
  location,
  guardKey
) {
  tmv3_upsertLocationCacheObject_(location);

  const locationId = tmv3_clean_(
    location['Location ID']
  );

  const updatedContract = Object.assign(
    {},
    contract,
    {
      expectedLocationId:locationId,
      locationStatus:'MATCHED',
      actions:(contract.actions || []).filter(function(action) {
        return action !== TMV3_STEP7_ACTION.CREATE_LOCATION;
      })
    }
  );

  const updatedBundle = tmv3_step7BundleFromContract_(
    updatedContract
  );

  return {
    status:'LOCATION_VERIFIED',
    locationId:locationId,
    guardKey:guardKey,
    location:location,
    contract:updatedContract,
    bundle:updatedBundle
  };
}

function tmv3_ensureStep7LocationUnlocked_(
  bundle,
  scope,
  contract
) {
  tmv3_assertOperationWrite_(scope);
  contract = tmv3_step7ValidateExecutionContract_(contract);

  if (
    (contract.actions || []).indexOf(
      TMV3_STEP7_ACTION.CREATE_LOCATION
    ) === -1
  ) {
    return {
      status:'LOCATION_NOT_REQUIRED',
      contract:contract,
      bundle:bundle
    };
  }

  if (contract.vertical !== 'PreInspection') {
    throw new Error(
      'CREATE_LOCATION is currently authorized only for PreInspection. ' +
      contract.vertical +
      ' requires explicit relationship review to avoid duplicating an Order-linked Location.'
    );
  }

  const payload = tmv3_parseCanadianCustomerLocation_(
    contract.expectedLocationAddress,
    contract.expectedCustomerId
  );

  const guardKey = tmv3_locationCreateGuardKey_(contract);
  const guard = tmv3_readDurableWriteGuard_(guardKey);

  const preflight = tmv3_freshCustomerLocationCandidates_(
    bundle,
    contract.expectedCustomerId,
    contract.expectedLocationAddress
  );

  if (preflight.matches.length > 1) {
    throw new Error(
      'CREATE_LOCATION blocked: multiple Customer-owned Locations match the Calendar address.'
    );
  }

  if (guard) {
    const knownId = Number(guard.locationId || 0);

    if (knownId) {
      const known = preflight.owned.filter(function(location) {
        return Number(location['Location ID'] || 0) === knownId;
      });

      if (
        known.length === 1 &&
        tmv3_addressStrongMatch_(
          tmv3_addressParts_(contract.expectedLocationAddress),
          tmv3_addressParts_(
            tmv3_locationFullAddress_(known[0])
          )
        )
      ) {
        tmv3_writeDurableWriteGuard_(
          guardKey,
          'VERIFIED',
          Object.assign({}, guard, {
            locationId:knownId,
            customerId:contract.expectedCustomerId,
            address:contract.expectedLocationAddress
          })
        );

        return tmv3_locationResultFromVerified_(
          bundle,
          contract,
          known[0],
          guardKey
        );
      }

      throw new Error(
        'CREATE_LOCATION blocked: Location ' +
        knownId +
        ' was previously captured but fresh Striven Location data cannot verify it yet. Reconcile before retry.'
      );
    }

    if (preflight.matches.length === 1) {
      const recoveredId = Number(
        preflight.matches[0]['Location ID'] || 0
      );

      tmv3_writeDurableWriteGuard_(
        guardKey,
        'VERIFIED',
        Object.assign({}, guard, {
          locationId:recoveredId,
          customerId:contract.expectedCustomerId,
          address:contract.expectedLocationAddress,
          reconciledFrom:'FRESH_LOCATION_REPORT'
        })
      );

      return tmv3_locationResultFromVerified_(
        bundle,
        contract,
        preflight.matches[0],
        guardKey
      );
    }

    if (!tmv3_preInspectionRejectedAddressGuard_(guard,preflight)) {
      throw new Error(
        'CREATE_LOCATION blocked: a previous Location write attempt is ' +
        tmv3_clean_(guard.state || 'UNCERTAIN') +
        ' and no unique fresh Location can yet be reconciled. No second POST is allowed.'
      );
    }
    tmv3_audit_(contract.vertical,contract.eventId,'','LOCATION_VALIDATION_RECONCILED','PASS',
      'Prior explicit HTTP 400 Address/Country validation rejection; complete fresh owned-Location read found no match. Corrected nested Address and CA country-code schema authorized once.');
  }

  if (preflight.matches.length === 1) {
    const existingId = Number(
      preflight.matches[0]['Location ID'] || 0
    );

    tmv3_writeDurableWriteGuard_(
      guardKey,
      'VERIFIED',
      {
        vertical:contract.vertical,
        eventId:contract.eventId,
        customerId:contract.expectedCustomerId,
        address:contract.expectedLocationAddress,
        locationId:existingId,
        reconciledFrom:'PREEXISTING_FRESH_LOCATION'
      }
    );

    return tmv3_locationResultFromVerified_(
      bundle,
      contract,
      preflight.matches[0],
      guardKey
    );
  }

  tmv3_writeDurableWriteGuard_(
    guardKey,
    'ATTEMPT_STARTED',
    {
      vertical:contract.vertical,
      eventId:contract.eventId,
      customerId:contract.expectedCustomerId,
      address:contract.expectedLocationAddress,
      addressSchema:'nested-address-country-code-v2'
    }
  );

  let response;
  let locationId = 0;

  try {
    response = tmv3_fetchJson_(
      TMV3.API_BASE +
        '/v1/customers/' +
        encodeURIComponent(
          Number(contract.expectedCustomerId)
        ) +
        '/location',
      {
        method:'post',
        contentType:'application/json',
        payload:JSON.stringify(payload)
      }
    );

    locationId = tmv3_extractCreatedLocationId_(response);

    if (!locationId) {
      throw new Error(
        'Location POST returned success but no durable Location ID was extracted.'
      );
    }
  } catch (err) {
    tmv3_writeDurableWriteGuard_(
      guardKey,
      'UNCERTAIN',
      {
        vertical:contract.vertical,
        eventId:contract.eventId,
        customerId:contract.expectedCustomerId,
        address:contract.expectedLocationAddress,
      addressSchema:'nested-address-country-code-v2',
        error:String(err && err.message || err)
      }
    );

    tmv3_audit_(
      contract.vertical,
      contract.eventId,
      contract.taskId || '',
      'CREATE_LOCATION',
      'CREATE UNCERTAIN',
      'No retry is authorized until the Customer Location is reconciled. ' +
        String(err && err.message || err)
    );

    throw new Error(
      'CREATE LOCATION UNCERTAIN. Do not retry blindly. Reconcile Striven first. ' +
      String(err && err.message || err)
    );
  }

  tmv3_writeDurableWriteGuard_(
    guardKey,
    'ID_CAPTURED',
    {
      vertical:contract.vertical,
      eventId:contract.eventId,
      customerId:contract.expectedCustomerId,
      address:contract.expectedLocationAddress,
      addressSchema:'nested-address-country-code-v2',
      locationId:locationId
    }
  );

  const readback = tmv3_freshCustomerLocationCandidates_(
    bundle,
    contract.expectedCustomerId,
    contract.expectedLocationAddress
  );

  const verified = readback.owned.filter(function(location) {
    return (
      Number(location['Location ID'] || 0) ===
      Number(locationId)
    );
  });

  if (verified.length !== 1) {
    throw new Error(
      'Location ' +
      locationId +
      ' was created and captured, but fresh Location report read-back has not converged. No Task write is authorized yet.'
    );
  }

  if (
    !tmv3_addressStrongMatch_(
      tmv3_addressParts_(contract.expectedLocationAddress),
      tmv3_addressParts_(
        tmv3_locationFullAddress_(verified[0])
      )
    )
  ) {
    throw new Error(
      'Created Location read-back address does not match the verified Calendar job-site address.'
    );
  }

  tmv3_writeDurableWriteGuard_(
    guardKey,
    'VERIFIED',
    {
      vertical:contract.vertical,
      eventId:contract.eventId,
      customerId:contract.expectedCustomerId,
      address:contract.expectedLocationAddress,
      addressSchema:'nested-address-country-code-v2',
      locationId:locationId
    }
  );

  tmv3_audit_(
    contract.vertical,
    contract.eventId,
    contract.taskId || '',
    'CREATE_LOCATION',
    'PASS',
    JSON.stringify({
      locationId:locationId,
      customerId:contract.expectedCustomerId,
      readback:'FRESH_LOCATION_REPORT'
    })
  );

  return tmv3_locationResultFromVerified_(
    bundle,
    contract,
    verified[0],
    guardKey
  );
}

function tmv3_ensureStep7Location_(
  bundle,
  scope,
  contract
) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error(
      'CREATE_LOCATION blocked because another V3 write is already in progress.'
    );
  }

  try {
    return tmv3_ensureStep7LocationUnlocked_(
      bundle,
      scope,
      contract
    );
  } finally {
    lock.releaseLock();
  }
}

function tmv3_existingPersistedTaskForEvent_(
  vertical,
  eventId,
  options
) {
  options = options || {};

  const sourceRows = Array.isArray(options.rows)
    ? options.rows
    : tmv3_rows_(TMV3.SHEETS.STATE);

  const rows = sourceRows.filter(function(row) {
    return (
      tmv3_clean_(row['Vertical']) === tmv3_clean_(vertical) &&
      tmv3_clean_(row['Event ID']) === tmv3_clean_(eventId) &&
      !!tmv3_clean_(row['Task ID'])
    );
  });

  const taskIds = tmv3_unique_(
    rows.map(function(row) {
      return tmv3_clean_(row['Task ID']);
    }).filter(Boolean)
  );

  if (!taskIds.length) {
    return {
      status:'NONE',
      taskIds:[],
      taskId:0,
      task:null,
      state:null
    };
  }

  if (taskIds.length > 1) {
    return {
      status:'MULTIPLE_KNOWN',
      taskIds:taskIds,
      taskId:0,
      task:null,
      state:null,
      reason:
        'Multiple persisted Task IDs exist for this Calendar event: ' +
        taskIds.join(', ') +
        '.'
    };
  }

  const taskId = Number(taskIds[0] || 0);
  const state = rows.filter(function(row) {
    return Number(row['Task ID'] || 0) === taskId;
  })[0] || null;

  const readTask =
    typeof options.readTask === 'function'
      ? options.readTask
      : tmv3_getTaskById_;

  try {
    const task = readTask(taskId);

    if (!task || !tmv3_clean_(task['Task ID'])) {
      return {
        status:'KNOWN_READ_FAILED',
        taskIds:taskIds,
        taskId:taskId,
        task:null,
        state:state,
        reason:
          'Persisted Task ' +
          taskId +
          ' did not return an authoritative Task record.'
      };
    }

    return {
      status:'FOUND_VERIFIED',
      taskIds:taskIds,
      taskId:taskId,
      task:task,
      state:state
    };
  } catch (err) {
    return {
      status:'KNOWN_READ_FAILED',
      taskIds:taskIds,
      taskId:taskId,
      task:null,
      state:state,
      reason:
        'Persisted Task ' +
        taskId +
        ' could not be freshly read: ' +
        String(err && err.message || err)
    };
  }
}

function tmv3_persistCreatedTaskIdState_(bundle, taskId) {
  tmv3_upsertState_([{
    vertical: bundle.eventRecord.vertical,
    eventId: bundle.eventRecord.eventId,
    calendarId: bundle.eventRecord.calendarId,
    customerId: bundle.resolved.customerId,
    locationId: bundle.resolved.locationId,
    contactId: bundle.resolved.contactId,
    orderId: bundle.resolved.orderId,
    taskId: Number(taskId),
    fingerprint: bundle.eventRecord.fingerprint,
    calendarUpdatedAt: bundle.eventRecord.calendarUpdatedAt,
    status: 'READY',
    nextAction: 'VERIFY CREATED TASK',
    errorCode: '',
    lastVerified: ''
  }]);
}

function tmv3_createOrRecreateFromBundle_(bundle, scope, contract) {
  // The dedicated lane holds this same script lock across the entire event
  // transaction. Reacquiring/releasing it here would invalidate that guard.
  if (['preinspectionauto','verticalauto'].indexOf(tmv3_norm_(scope))>=0) {
    return tmv3_createOrRecreateFromBundleUnlocked_(bundle, scope, contract);
  }
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error(
      'CREATE / RECREATE blocked because another V3 write is already in progress.'
    );
  }

  try {
    return tmv3_createOrRecreateFromBundleUnlocked_(
      bundle,
      scope,
      contract
    );
  } finally {
    lock.releaseLock();
  }
}

function tmv3_createOrRecreateFromBundleUnlocked_(
  bundle,
  scope,
  contract
) {
  if(contract&&contract.vertical==='PreInspection'&&tmv3_preInspectionApprovalHold_(contract.eventId))
    throw new Error('OPERATOR_APPROVED_INTAKE_TASK_CREATION_HELD');
  contract = tmv3_step7ValidateExecutionContract_(contract);
  tmv3_assertOperationWrite_(scope);

  if (tmv3_clean_(contract.blocker)) {
    throw new Error(
      'CREATE / RECREATE blocked by Step 7: ' + contract.blocker
    );
  }

  if (
    contract.actions.indexOf(TMV3_STEP7_ACTION.CREATE_LOCATION) !== -1
  ) {
    throw new Error(
      'CREATE / RECREATE ordering error: Customer Location must be freshly verified before Task POST.'
    );
  }

  const isRecreate =
    contract.actions.indexOf(TMV3_STEP7_ACTION.RECREATE_TASK) !== -1;
  const isCreate =
    contract.actions.indexOf(TMV3_STEP7_ACTION.CREATE_TASK) !== -1;

  if (!isCreate && !isRecreate) {
    throw new Error(
      'Step 7 contract does not authorize CREATE / RECREATE.'
    );
  }

  if (bundle.resolved.taskId) {
    throw new Error(
      'CREATE blocked because a Task ID is already resolved.'
    );
  }

  if (
    isCreate &&
    (contract.sourceTaskIdList || []).length
  ) {
    throw new Error(
      'CREATE blocked because Step 7 still exposes existing source Task IDs: ' +
      contract.sourceTaskIdList.join(',') + '.'
    );
  }

  const persisted = tmv3_existingPersistedTaskForEvent_(
    contract.vertical,
    contract.eventId
  );

  if (persisted.status === 'KNOWN_READ_FAILED') {
    throw new Error(
      'CREATE blocked: ' +
      persisted.reason +
      ' A failed read is not evidence that the Task does not exist.'
    );
  }

  if (persisted.status === 'MULTIPLE_KNOWN') {
    throw new Error(
      'CREATE blocked: ' +
      persisted.reason +
      ' Reconcile the event before any new Task POST.'
    );
  }

  if (persisted.status === 'FOUND_VERIFIED') {
    const sourceIds = (contract.sourceTaskIdList || [])
      .map(String);

    const authorizedRecreateSource =
      isRecreate &&
      sourceIds.length === 1 &&
      sourceIds[0] === String(persisted.taskId) &&
      tmv3_taskIsCompleted_(persisted.task['Status']);

    if (!authorizedRecreateSource) {
      throw new Error(
        'CREATE blocked: V3 previously persisted Task ' +
        persisted.taskId +
        ' for this Calendar event. Reconcile that Task first.'
      );
    }
  }

  const taskGuardKey = tmv3_taskCreateGuardKey_(contract);
  const priorGuard = tmv3_readDurableWriteGuard_(
    taskGuardKey
  );

  if (tmv3_createGuardStateBlocksRetry_(priorGuard)) {
    const priorTaskId = Number(
      priorGuard.taskId || 0
    );

    if (priorTaskId) {
      try {
        const priorTask = tmv3_getTaskById_(
          priorTaskId
        );

        throw new Error(
          'CREATE blocked: durable guard already captured Task ' +
          priorTaskId +
          ' (' +
          tmv3_clean_(priorTask['Status']) +
          '). Reconcile that Task instead of POSTing again.'
        );
      } catch (err) {
        if (
          String(err && err.message || err).indexOf(
            'durable guard already captured Task'
          ) !== -1
        ) {
          throw err;
        }

        throw new Error(
          'CREATE blocked: durable guard captured Task ' +
          priorTaskId +
          ' but the fresh read failed. No second POST is allowed. ' +
          String(err && err.message || err)
        );
      }
    }

    throw new Error(
      'CREATE blocked: prior write guard is ' +
      tmv3_clean_(priorGuard.state || 'UNCERTAIN') +
      '. Reconcile Striven before any retry.'
    );
  }

  const action = isRecreate ? 'RECREATE' : 'CREATE';
  const source = tmv3_buildCreateSource_(bundle, action);
  const request = {
    startDateTime:tmv3_strivenTaskDateTime_(
      new Date(contract.expectedStart)
    ),
    dueDateTime:tmv3_strivenTaskDateTime_(
      new Date(contract.expectedDue)
    ),
    calendarNotes:
      contract.vertical === 'PreInspection'
        ? null
        : tmv3_stripAuthoredCalendarNotes_(
            bundle.eventRecord.description
          )
  };
  const payload = tmv3_buildReplacementCreatePayload_(
    source,
    request
  );

  if (contract.vertical === 'PreInspection') {
    tmv3_enforcePreInspectionCreatePayloadHardRules_(payload);
    tmv3_assertPreInspectionCreatePayloadHardRules_(payload);
  }

  let json;
  let newTaskId = 0;

  tmv3_writeDurableWriteGuard_(
    taskGuardKey,
    'ATTEMPT_STARTED',
    {
      vertical:contract.vertical,
      eventId:contract.eventId,
      action:action,
      sourceTaskIds:(contract.sourceTaskIdList || []).slice()
    }
  );

  try {
    json = tmv3_fetchJson_(
      TMV3.API_BASE + '/v2/tasks',
      {
        method:'post',
        contentType:'application/json',
        payload:JSON.stringify(payload)
      }
    );
    newTaskId = Number(
      tmv3_extractCreatedTaskId_(json) || 0
    );

    if (!newTaskId) {
      throw new Error(
        'POST succeeded but created Task ID could not be extracted.'
      );
    }
  } catch (err) {
    tmv3_writeDurableWriteGuard_(
      taskGuardKey,
      'UNCERTAIN',
      {
        vertical:contract.vertical,
        eventId:contract.eventId,
        action:action,
        sourceTaskIds:(contract.sourceTaskIdList || []).slice(),
        error:String(err && err.message || err)
      }
    );

    tmv3_audit_(
      contract.vertical,
      contract.eventId,
      '',
      action + '_TASK',
      'CREATE UNCERTAIN',
      'Do not retry blindly. ' +
        String(err && err.message || err)
    );

    throw new Error(
      'CREATE UNCERTAIN. Do not retry blindly. ' +
      'Reconcile Striven first. ' +
      String(err && err.message || err)
    );
  }

  tmv3_writeDurableWriteGuard_(
    taskGuardKey,
    'ID_CAPTURED',
    {
      vertical:contract.vertical,
      eventId:contract.eventId,
      action:action,
      sourceTaskIds:(contract.sourceTaskIdList || []).slice(),
      taskId:newTaskId
    }
  );

  tmv3_persistCreatedTaskIdState_(
    bundle,
    newTaskId
  );

  const createdResolved = Object.assign(
    {},
    bundle.resolved,
    {
      taskId:newTaskId,
      task:String(newTaskId),
      status:'READY'
    }
  );

  const createdBundle = Object.assign({}, bundle, {
    resolved:createdResolved,
    records:[createdResolved]
  });

  let assignment = null;
  if (
    contract.actions.indexOf(
      TMV3_STEP7_ACTION.PATCH_ASSIGNMENTS
    ) !== -1
  ) {
    assignment = tmv3_reconcileTaskAssignments_(
      bundle.eventRecord,
      newTaskId,
      scope,
      {
        newTask:true,
        desiredAssignment:
          tmv3_step7DesiredAssignmentFromContract_(contract)
      }
    );
  }

  const readback = tmv3_getTaskById_(newTaskId);
  tmv3_upsertTaskCacheRow_(readback);

  if (!tmv3_taskIsOpen_(readback['Status'])) {
    throw new Error(
      'Created Task read-back is not OPEN. Task ID ' +
      newTaskId + '.'
    );
  }

  if (
    String(readback['Customer ID'] || '') !==
    String(contract.expectedCustomerId)
  ) {
    throw new Error(
      'Created Task Customer read-back mismatch.'
    );
  }

  if (
    String(readback['Location ID'] || '') !==
    String(contract.expectedLocationId)
  ) {
    throw new Error(
      'Created Task Location read-back mismatch.'
    );
  }

  if (
    contract.vertical === 'PreInspection' &&
    (
      Number(readback['Task Type ID'] || 0) !== 105 ||
      readback['Order ID']
    )
  ) {
    throw new Error(
      'Created PreInspection Task violated Type 105 / no-SO invariant.'
    );
  }

  let calendarLinks = null;
  if (
    contract.actions.indexOf(
      TMV3_STEP7_ACTION.VERIFY_CALENDAR_LINKS
    ) !== -1
  ) {
    calendarLinks = tmv3_operationWriteCalendarLinks_(
      createdBundle,
      scope
    );
  }

  let calendarTitle = null;
  if (
    contract.actions.indexOf(
      TMV3_STEP7_ACTION.PATCH_CALENDAR_TITLE
    ) !== -1
  ) {
    calendarTitle = tmv3_operationNormalizeCalendarTitle_(
      createdBundle,
      contract,
      scope
    );
  }

  const after = tmv3_step7FreshExecutionContract_(
    contract.vertical,
    contract.eventId,
    newTaskId
  );

  if (
    after.plan !== 'NO_CHANGE' ||
    tmv3_clean_(after.blocker)
  ) {
    tmv3_audit_(
      contract.vertical,
      contract.eventId,
      newTaskId,
      action + '_TASK',
      'ATTENTION',
      JSON.stringify({
        expectedTerminalPlan:'NO_CHANGE',
        freshPlan:after.plan,
        blocker:after.blocker || '',
        createdTaskId:newTaskId
      })
    );
    throw new Error(
      'CREATE read-back did not converge to Step 7 NO_CHANGE; ' +
      'fresh plan is ' + after.plan + '.'
    );
  }

  tmv3_writeDurableWriteGuard_(
    taskGuardKey,
    'VERIFIED',
    {
      vertical:contract.vertical,
      eventId:contract.eventId,
      action:action,
      sourceTaskIds:(contract.sourceTaskIdList || []).slice(),
      taskId:newTaskId,
      terminalPlan:after.plan
    }
  );

  tmv3_audit_(
    contract.vertical,
    contract.eventId,
    newTaskId,
    action + '_TASK',
    'PASS',
    JSON.stringify({
      assignment:assignment,
      calendarLinks:calendarLinks,
      calendarTitle:calendarTitle
    })
  );

  return {
    status:'CREATED_VERIFIED_AND_CONVERGED',
    action:action,
    taskId:newTaskId,
    assignment:assignment,
    calendarLinks:calendarLinks,
    calendarTitle:calendarTitle,
    readback:readback,
    readbackPlan:after
  };
}

function tmv3_executeExistingTaskSync_(
  bundle,
  scope,
  mode,
  contract
) {
  contract = tmv3_step7ValidateExecutionContract_(contract);
  if (contract.vertical === 'PreInspection') tmv3_piAssertLockedStandard_();
  tmv3_assertOperationWrite_(scope);

  if (!contract.taskId) {
    throw new Error(
      'Existing Task execution requires a Step 7 Task ID.'
    );
  }

  if (tmv3_clean_(contract.blocker)) {
    throw new Error(
      'Existing Task execution blocked by Step 7: ' +
      contract.blocker
    );
  }

  const actions = tmv3_step7ActionsForMode_(
    contract,
    mode || 'ALL'
  );

  const taskId = Number(contract.taskId);
  let patch = null;
  let assignment = null;
  let calendarLinks = null;
  let calendarTitle = null;
  let taskName = null;

  const relationshipActions = [
    TMV3_STEP7_ACTION.PATCH_LOCATION,
    TMV3_STEP7_ACTION.PATCH_ORDER,
    TMV3_STEP7_ACTION.PATCH_REQUESTED_BY
  ];

  if (relationshipActions.some(function(action) {
    return actions.indexOf(action) !== -1;
  })) {
    patch = tmv3_operationPatchTask_(
      taskId,
      tmv3_step7RelationshipPayload_(contract),
      scope
    );
  }

  if (
    actions.indexOf(TMV3_STEP7_ACTION.PATCH_DATES) !== -1
  ) {
    const datePatch = tmv3_operationPatchTask_(
      taskId,
      tmv3_step7DatePayload_(contract),
      scope
    );
    patch = patch || datePatch;
  }

  if (
    actions.indexOf(
      TMV3_STEP7_ACTION.PATCH_ASSIGNMENTS
    ) !== -1
  ) {
    assignment = tmv3_reconcileTaskAssignments_(
      bundle.eventRecord,
      taskId,
      scope,
      {
        desiredAssignment:
          tmv3_step7DesiredAssignmentFromContract_(contract)
      }
    );
  }

  if (
    actions.indexOf(
      TMV3_STEP7_ACTION.VERIFY_CALENDAR_LINKS
    ) !== -1
  ) {
    calendarLinks = tmv3_operationWriteCalendarLinks_(
      bundle,
      scope
    );
  }

  // Title mutation is intentionally last among existing-record mutations.
  // Within the title pair, Calendar is always written and verified first,
  // then the canonical OPEN Task Name is patched independently.
  if (
    actions.indexOf(
      TMV3_STEP7_ACTION.PATCH_CALENDAR_TITLE
    ) !== -1
  ) {
    calendarTitle = tmv3_operationNormalizeCalendarTitle_(
      bundle,
      contract,
      scope
    );
  }

  if (
    actions.indexOf(
      TMV3_STEP7_ACTION.PATCH_TASK_NAME
    ) !== -1
  ) {
    try {
      taskName = tmv3_operationNormalizeTaskName_(
        taskId,
        contract.desiredTaskName,
        contract,
        scope
      );
    } catch (err) {
      const afterFailure = tmv3_step7FreshExecutionContract_(
        contract.vertical,
        contract.eventId,
        taskId
      );

      const calendarWasWritten = !!(
        calendarTitle &&
        (calendarTitle.copies || []).some(function(copy) {
          return copy && copy.writePerformed === true;
        })
      );

      const partial = {
        status:
          calendarWasWritten
            ? 'PARTIAL_RECONCILIATION_REQUIRED'
            : 'TASK_NAME_WRITE_FAILED_NO_CALENDAR_MUTATION',
        vertical:contract.vertical,
        eventId:contract.eventId,
        taskId:taskId,
        failedAction:TMV3_STEP7_ACTION.PATCH_TASK_NAME,
        error:String(err && err.message || err),
        calendarTitle:calendarTitle,
        freshPlan:afterFailure.plan,
        freshActions:(afterFailure.actions || []).slice(),
        readback:afterFailure
      };

      tmv3_audit_(
        contract.vertical,
        contract.eventId,
        taskId,
        'STEP7_TITLE',
        calendarWasWritten ? 'PARTIAL' : 'ATTENTION',
        JSON.stringify(partial)
      );

      return partial;
    }
  }

  const after = tmv3_step7FreshExecutionContract_(
    contract.vertical,
    contract.eventId,
    taskId
  );

  const fullExecution = (mode || 'ALL') === 'ALL';
  const titleExecution = (mode || 'ALL') === 'TITLE';
  const titleActionsRemaining = (after.actions || []).filter(function(action) {
    return [
      TMV3_STEP7_ACTION.PATCH_CALENDAR_TITLE,
      TMV3_STEP7_ACTION.PATCH_TASK_NAME
    ].indexOf(action) !== -1;
  });

  if (
    fullExecution &&
    (
      after.plan !== 'NO_CHANGE' ||
      tmv3_clean_(after.blocker)
    )
  ) {
    tmv3_audit_(
      contract.vertical,
      contract.eventId,
      taskId,
      'STEP7_' + String(mode || 'ALL'),
      'ATTENTION',
      JSON.stringify({
        expectedTerminalPlan:'NO_CHANGE',
        freshPlan:after.plan,
        blocker:after.blocker || '',
        executedActions:actions
      })
    );
    throw new Error(
      'Step 7 read-back did not converge to NO_CHANGE; ' +
      'fresh plan is ' + after.plan + '.'
    );
  }

  if (
    titleExecution &&
    (
      titleActionsRemaining.length ||
      tmv3_clean_(after.calendarTitleBlocker) ||
      tmv3_clean_(after.taskNameBlocker)
    )
  ) {
    const detail = {
      expectedTerminalTitleActions:[],
      freshPlan:after.plan,
      remainingTitleActions:titleActionsRemaining,
      calendarTitleCheck:after.calendarTitleCheck || '',
      taskNameCheck:after.taskNameCheck || '',
      calendarTitleBlocker:after.calendarTitleBlocker || '',
      taskNameBlocker:after.taskNameBlocker || ''
    };

    tmv3_audit_(
      contract.vertical,
      contract.eventId,
      taskId,
      'STEP7_TITLE',
      'ATTENTION',
      JSON.stringify(detail)
    );

    throw new Error(
      'Title read-back did not converge; remaining title actions: ' +
      titleActionsRemaining.join(',') + '.'
    );
  }

  tmv3_audit_(
    contract.vertical,
    contract.eventId,
    taskId,
    'STEP7_' + String(mode || 'ALL'),
    'PASS',
    JSON.stringify({
      plan:contract.plan,
      actions:actions,
      patch:patch && patch.status,
      assignment:assignment && assignment.status,
      calendarLinks:calendarLinks && calendarLinks.status,
      calendarTitle:calendarTitle && calendarTitle.status,
      taskName:taskName && taskName.status,
      calendarTitleCheck:after.calendarTitleCheck || '',
      taskNameCheck:after.taskNameCheck || '',
      readbackPlan:after.plan
    })
  );

  return {
    status:
      fullExecution
        ? 'VERIFIED_CONVERGENCE'
        : (
            titleExecution
              ? 'TITLE_NORMALIZATION_VERIFIED'
              : 'VERIFIED_PARTIAL'
          ),
    mode:mode || 'ALL',
    vertical:contract.vertical,
    eventId:contract.eventId,
    taskId:taskId,
    plan:contract.plan,
    actions:actions,
    patch:patch,
    assignment:assignment,
    calendarLinks:calendarLinks,
    calendarTitle:calendarTitle,
    taskName:taskName,
    readbackPlan:after
  };
}

function tmv3_refreshOperatorAfterManual_() {
  try {
    return tmv3_shadowMapFromCache();
  } catch (err) {
    tmv3_audit_('SYSTEM','','','POST_MANUAL_MAP_REFRESH','ATTENTION',String(err && err.message || err));
    return { status: 'REFRESH_ATTENTION', error: String(err && err.message || err) };
  }
}

function tmv3_executeVerifiedStep7ExistingPlan(
  vertical,
  eventId,
  taskId,
  expectedPlan
) {
  const contract = tmv3_step7FreshExecutionContract_(
    vertical,
    eventId,
    taskId
  );

  const previousPlan = tmv3_clean_(expectedPlan);

  if (previousPlan && previousPlan !== contract.plan) {
    const changed = {
      status:'PLAN_CHANGED_NO_WRITE',
      previousPlan:previousPlan,
      freshPlan:contract.plan,
      reason:
        'Fresh canonical Step 7 plan differs from the expected plan.',
      eventId:contract.eventId,
      taskId:contract.taskId
    };
    tmv3_audit_(
      contract.vertical,
      contract.eventId,
      contract.taskId || '',
      'STEP7_PLAN_GUARD',
      'NO_WRITE',
      JSON.stringify(changed)
    );
    return changed;
  }

  return tmv3_executeFreshStep7Selection_(
    {
      previousPlan:previousPlan,
      contract:contract
    },
    'MANUAL',
    'ALL'
  );
}



function tmv3_previewTitleNormalizationForEvent(
  vertical,
  eventId,
  taskId
) {
  const contract = tmv3_step7FreshExecutionContract_(
    vertical,
    eventId,
    taskId
  );

  const result = {
    mode:'READ_ONLY',
    authority:'FRESH_STEP7_CANONICAL_CONTRACT',
    vertical:contract.vertical,
    eventId:contract.eventId,
    taskId:contract.taskId || '',
    plan:contract.plan,
    blocker:contract.blocker || '',
    actions:(contract.actions || []).slice(),
    titleNormalization:tmv3_titleNormalizationPreview_(contract),
    manualWritesEnabled:tmv3_operationWritesEnabled_('MANUAL'),
    automationWritesEnabled:tmv3_operationWritesEnabled_('AUTO')
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function tmv3_executeVerifiedTitleNormalization(
  vertical,
  eventId,
  taskId,
  expectedPlan
) {
  const expected = tmv3_clean_(expectedPlan);
  if (!expected) {
    throw new Error(
      'Title normalization execution requires the exact plan returned by the prior read-only preview.'
    );
  }

  const contract = tmv3_step7FreshExecutionContract_(
    vertical,
    eventId,
    taskId
  );

  const titleActions = tmv3_step7ActionsForMode_(
    contract,
    'TITLE'
  );

  if (expected !== tmv3_clean_(contract.plan)) {
    return {
      status:'PLAN_CHANGED_NO_WRITE',
      previousPlan:expected,
      freshPlan:contract.plan,
      vertical:contract.vertical,
      eventId:contract.eventId,
      taskId:contract.taskId || ''
    };
  }

  if (tmv3_clean_(contract.blocker)) {
    return {
      status:'BLOCKED_NO_WRITE',
      blocker:contract.blocker,
      vertical:contract.vertical,
      eventId:contract.eventId,
      taskId:contract.taskId || '',
      plan:contract.plan
    };
  }

  if (!titleActions.length) {
    return {
      status:'NO_TITLE_CHANGE_REQUIRED',
      vertical:contract.vertical,
      eventId:contract.eventId,
      taskId:contract.taskId || '',
      plan:contract.plan,
      calendarTitleCheck:contract.calendarTitleCheck || '',
      taskNameCheck:contract.taskNameCheck || ''
    };
  }

  return tmv3_executeFreshStep7Selection_(
    {
      previousPlan:expected,
      contract:contract
    },
    'MANUAL',
    'TITLE'
  );
}

function tmv3_titleNormalizationPreview_(contract) {
  contract = tmv3_step7ValidateExecutionContract_(contract);

  const bundle = tmv3_step7BundleFromContract_(contract);
  const customer = tmv3_customerFromRefs_(
    bundle.refs,
    contract.expectedCustomerId
  );
  const location = tmv3_locationFromRefs_(
    bundle.refs,
    contract.expectedCustomerId,
    contract.expectedLocationId
  );
  const order = contract.expectedOrderId
    ? bundle.refs.orderById[String(contract.expectedOrderId)] || null
    : null;

  const previewRecord = Object.assign(
    {},
    bundle.eventRecord,
    {
      step3:{
        disposition:'VERIFIED',
        anchor:{
          orderId:contract.expectedOrderId || '',
          orderNumber:tmv3_clean_(order && order['Order Number']),
          customerNumber:tmv3_clean_(customer && customer['Customer Number'])
        }
      },
      step4:{
        disposition:
          customer && location
            ? 'VERIFIED'
            : 'NOT_VERIFIED',
        customer:customer || null,
        location:location || null
      }
    }
  );

  const desiredCalendar = tmv3_desiredCalendarTitle_(previewRecord);

  let currentTaskName = '';
  let currentTaskReadStatus = contract.taskId
    ? 'NOT_READ'
    : 'NO_TASK';

  if (contract.taskId) {
    try {
      const task = tmv3_getTaskById_(contract.taskId);
      currentTaskName = tmv3_clean_(task && task['Name']);
      currentTaskReadStatus = 'FRESH_TASK_GET';
    } catch (err) {
      currentTaskReadStatus =
        'TASK_READ_FAILED: ' +
        String(err && err.message || err);
    }
  }

  const desiredTask = tmv3_desiredTaskName_(
    previewRecord,
    { sourceTaskTitle:currentTaskName }
  );

  const currentCalendarTitle = tmv3_clean_(
    bundle.eventRecord && bundle.eventRecord.title
  );
  const rawDescription = String(
    bundle.eventRecord &&
    (
      bundle.eventRecord.rawDescription !== undefined
        ? bundle.eventRecord.rawDescription
        : bundle.eventRecord.description
    ) || ''
  );

  const descriptionPlan =
    desiredCalendar.status === 'READY'
      ? tmv3_titleDescriptionPlan_(
          currentCalendarTitle,
          desiredCalendar.value,
          rawDescription,
          {
            preserveOldTitle:true
          }
        )
      : {
          status:'NOT_PLANNED',
          before:rawDescription,
          after:rawDescription,
          titleChange:false,
          descriptionChange:false,
          alreadyPreserved:false,
          reason:desiredCalendar.reason || ''
        };

  return {
    jev:'JEV NOT USED — DETERMINISTIC',
    calendar:{
      currentTitle:currentCalendarTitle,
      desiredTitle:desiredCalendar.value || '',
      desiredStatus:desiredCalendar.status,
      desiredSource:desiredCalendar.source || '',
      check:
        desiredCalendar.status === 'READY'
          ? (
              currentCalendarTitle === tmv3_clean_(desiredCalendar.value)
                ? 'MATCH'
                : 'MISMATCH'
            )
          : desiredCalendar.status,
      blocker:
        desiredCalendar.status === 'BLOCKED'
          ? desiredCalendar.reason
          : '',
      description:{
        preservationStatus:descriptionPlan.status,
        action:
          descriptionPlan.descriptionChange
            ? 'PREPEND_OLD_TITLE_PLUS_ONE_BLANK_LINE'
            : 'NO_DESCRIPTION_CHANGE',
        alreadyPreserved:descriptionPlan.alreadyPreserved === true,
        currentLength:rawDescription.length,
        proposedLength:String(descriptionPlan.after || '').length,
        currentHash:tmv3_hash_(rawDescription),
        proposedHash:tmv3_hash_(String(descriptionPlan.after || ''))
      }
    },
    task:{
      taskId:contract.taskId || '',
      taskStatus:contract.taskStatus || '',
      currentName:currentTaskName,
      currentReadStatus:currentTaskReadStatus,
      desiredName:desiredTask.value || '',
      desiredStatus:desiredTask.status,
      desiredSource:desiredTask.source || '',
      check:
        !contract.taskId
          ? 'CREATE_USES_CANONICAL_NAME'
          : (
              desiredTask.status === 'READY'
                ? (
                    currentTaskName === tmv3_clean_(desiredTask.value)
                      ? 'MATCH'
                      : 'MISMATCH'
                  )
                : desiredTask.status
            ),
      blocker:
        desiredTask.status === 'BLOCKED'
          ? desiredTask.reason
          : ''
    },
    identitySources:{
      customerId:contract.expectedCustomerId || '',
      locationId:contract.expectedLocationId || '',
      orderId:contract.expectedOrderId || '',
      customerName:tmv3_clean_(customer && customer['Name']),
      customerNumber:tmv3_clean_(customer && customer['Customer Number']),
      locationAddress:tmv3_titleLocationDisplay_(location),
      authority:'STEP7_VERIFIED_IDS + CURRENT_SOURCE_TABLES'
    }
  };
}

function tmv3_previewSelectedAction() {
  const selection = tmv3_step7FreshSelectedContract_();
  const contract = tmv3_step7ValidateExecutionContract_(
    selection.contract
  );
  const publishedPlan = tmv3_clean_(selection.previousPlan);
  const planChanged =
    !publishedPlan ||
    publishedPlan !== tmv3_clean_(contract.plan);
  const locationCreatePending =
    (contract.actions || []).indexOf(
      TMV3_STEP7_ACTION.CREATE_LOCATION
    ) !== -1;

  const locationCreateSupported =
    !locationCreatePending ||
    (
      contract.vertical === 'PreInspection' &&
      !!tmv3_clean_(contract.expectedLocationAddress)
    );

  const titleNormalization =
    tmv3_titleNormalizationPreview_(contract);

  const output = {
    mode:'READ_ONLY',
    authority:'STEP7_CANONICAL_CONTRACT',
    vertical:contract.vertical,
    eventId:contract.eventId,
    taskId:contract.taskId || '',
    publishedPlan:publishedPlan,
    plan:contract.plan,
    planChanged:planChanged,
    blocker:contract.blocker || '',
    actions:(contract.actions || []).slice(),
    identifiers:{
      customerId:contract.expectedCustomerId || '',
      locationId:contract.expectedLocationId || '',
      orderId:contract.expectedOrderId || '',
      requestedById:contract.expectedRequestedById || '',
      requestedByType:contract.expectedRequestedByType || ''
    },
    schedule:{
      start:contract.expectedStart || '',
      due:contract.expectedDue || ''
    },
    desiredAssignment:{
      employeeIds:(contract.desiredAssignmentEmployeeIds || []).slice(),
      poolIds:(contract.desiredAssignmentPoolIds || []).slice()
    },
    titleNormalization:titleNormalization,
    mutationPreview:{
      authorizedByPlan:
        !planChanged &&
        !tmv3_clean_(contract.blocker) &&
        (contract.actions || []).length > 0 &&
        locationCreateSupported,
      externalWriteCurrentlyEnabled:
        tmv3_operationWritesEnabled_('MANUAL'),
      locationCreatePending:locationCreatePending,
      locationCreateSupported:locationCreateSupported,
      expectedLocationAddress:contract.expectedLocationAddress || '',
      expectedTerminalPlan:'NO_CHANGE'
    },
    manualWritesEnabled:tmv3_operationWritesEnabled_('MANUAL'),
    automationWritesEnabled:tmv3_operationWritesEnabled_('AUTO')
  };
  Logger.log(JSON.stringify(output, null, 2));
  return output;
}

function tmv3_syncSelectedAppointment() {
  const selection = tmv3_step7FreshSelectedContract_();
  const result = tmv3_executeFreshStep7Selection_(
    selection,
    'MANUAL',
    'ALL'
  );
  tmv3_refreshOperatorAfterManual_();
  return result;
}

function tmv3_pushSelectedDates() {
  const selection = tmv3_step7FreshSelectedContract_();
  const result = tmv3_executeFreshStep7Selection_(
    selection,
    'MANUAL',
    'DATES'
  );
  tmv3_refreshOperatorAfterManual_();
  return result;
}

function tmv3_pushSelectedRelationships() {
  const selection = tmv3_step7FreshSelectedContract_();
  const result = tmv3_executeFreshStep7Selection_(
    selection,
    'MANUAL',
    'RELATIONSHIPS'
  );
  tmv3_refreshOperatorAfterManual_();
  return result;
}

function tmv3_reconcileSelectedAssignee() {
  const selection = tmv3_step7FreshSelectedContract_();
  const result = tmv3_executeFreshStep7Selection_(
    selection,
    'MANUAL',
    'ASSIGNEE'
  );
  tmv3_refreshOperatorAfterManual_();
  return result;
}

function tmv3_fixSelectedCalendarLinks() {
  const selection = tmv3_step7FreshSelectedContract_();
  const result = tmv3_executeFreshStep7Selection_(
    selection,
    'MANUAL',
    'LINKS'
  );
  tmv3_refreshOperatorAfterManual_();
  return result;
}

function tmv3_createOrRecreateSelectedTask() {
  const selection = tmv3_step7FreshSelectedContract_();
  const contract = selection.contract;
  const isCreate =
    contract.actions.indexOf(
      TMV3_STEP7_ACTION.CREATE_TASK
    ) !== -1 ||
    contract.actions.indexOf(
      TMV3_STEP7_ACTION.RECREATE_TASK
    ) !== -1;

  if (!isCreate) {
    const notAuthorized = {
      status:'STEP7_ACTION_NOT_AUTHORIZED_NO_WRITE',
      plan:contract.plan,
      eventId:contract.eventId,
      taskId:contract.taskId || '',
      reason:
        'Fresh Step 7 does not authorize CREATE / RECREATE for this row.'
    };
    tmv3_audit_(
      contract.vertical,
      contract.eventId,
      contract.taskId || '',
      'STEP7_CREATE_GUARD',
      'NO_WRITE',
      JSON.stringify(notAuthorized)
    );
    return notAuthorized;
  }

  const result = tmv3_executeFreshStep7Selection_(
    selection,
    'MANUAL',
    'ALL'
  );
  tmv3_refreshOperatorAfterManual_();
  return result;
}

function tmv3_canaryWriteEventAllowed_(eventRecord) {
  const policy = tmv3_operationPolicy_();

  if (TMV3.MODE !== 'CANARY_WRITE') return true;

  const eventId = tmv3_clean_(policy.canaryEventId);
  const vertical = tmv3_clean_(policy.canaryVertical);
  const allowedDate = tmv3_clean_(policy.canaryDate);
  const eventStart = eventRecord && eventRecord.start
    ? new Date(eventRecord.start)
    : null;
  const eventDate =
    eventStart && !isNaN(eventStart.getTime())
      ? Utilities.formatDate(
          eventStart,
          TMV3_TIMEZONE,
          'yyyy-MM-dd'
        )
      : '';

  if (!eventId || !allowedDate) return false;

  return (
    tmv3_clean_(eventRecord && eventRecord.eventId) === eventId &&
    (!vertical || tmv3_clean_(eventRecord && eventRecord.vertical) === vertical) &&
    eventDate === allowedDate
  );
}

function tmv3_runSafeReadyRows_(scope) {
  const policy = tmv3_operationPolicy_();
  const maxWrites = Number(policy.autoMaxWritesPerRun || 10);
  const isAuto = tmv3_norm_(scope) === 'auto';

  if (isAuto && !tmv3_operationWritesEnabled_('AUTO')) {
    tmv3_audit_(
      'SYSTEM','','','AUTO_WRITE_CYCLE','GATED',
      'Scheduled external writes remain disabled.'
    );
    return {
      status:'AUTO_WRITES_GATED',
      writes:0
    };
  }

  const events = tmv3_calendarRecords_().filter(function(eventRecord) {
    return tmv3_canaryWriteEventAllowed_(eventRecord);
  });
  const results = [];

  for (
    let i = 0;
    i < events.length && results.length < maxWrites;
    i++
  ) {
    const eventRecord = events[i];
    let contracts = [];

    try {
      contracts = tmv3_step7FreshPlansForEvent_(
        eventRecord.vertical,
        eventRecord.eventId
      ).filter(function(contract) {
        return (
          !tmv3_clean_(contract.blocker) &&
          (contract.actions || []).length > 0
        );
      });
    } catch (err) {
      results.push({
        vertical:eventRecord.vertical,
        eventId:eventRecord.eventId,
        status:'ATTENTION',
        error:String(err && err.message || err)
      });
      continue;
    }

    for (
      let j = 0;
      j < contracts.length && results.length < maxWrites;
      j++
    ) {
      const contract = contracts[j];

      try {
        const one = tmv3_executeFreshStep7Selection_(
          {
            previousPlan:contract.plan,
            contract:contract
          },
          scope,
          'ALL'
        );

        results.push({
          vertical:contract.vertical,
          eventId:contract.eventId,
          taskId:one.taskId || contract.taskId || '',
          status:one.status
        });
      } catch (err) {
        tmv3_audit_(
          contract.vertical,
          contract.eventId,
          contract.taskId || '',
          'SAFE_READY_ROW',
          'ATTENTION',
          String(err && err.message || err)
        );

        results.push({
          vertical:contract.vertical,
          eventId:contract.eventId,
          taskId:contract.taskId || '',
          status:'ATTENTION',
          error:String(err && err.message || err)
        });
      }
    }
  }

  tmv3_shadowMapFromCache();

  return {
    status:'COMPLETE',
    scope:scope,
    writes:results.filter(function(item) {
      return /VERIFIED|CREATED/.test(
        String(item.status || '')
      );
    }).length,
    attempts:results.length,
    results:results
  };
}

function tmv3_runSafeReadyRows_MANUAL() {
  return tmv3_runSafeReadyRows_('MANUAL');
}

function tmv3_configuredCanaryContractDecision_(contract, policy) {
  contract = tmv3_step7ValidateExecutionContract_(contract);
  policy = policy || {};

  const expectedCustomerId = tmv3_clean_(
    policy.canaryExpectedCustomerId
  );
  const expectedLocationId = tmv3_clean_(
    policy.canaryExpectedLocationId
  );
  const expectedAction = tmv3_clean_(
    policy.canaryExpectedAction
  );
  const actualCustomerId = tmv3_clean_(
    contract.expectedCustomerId
  );
  const actualLocationId = tmv3_clean_(
    contract.expectedLocationId
  );
  const actions = Array.isArray(contract.actions)
    ? contract.actions.slice()
    : [];
  const taskId = Number(contract.taskId || 0);

  if (actualCustomerId !== expectedCustomerId) {
    throw new Error(
      'Configured canary Customer mismatch. Expected ' +
      expectedCustomerId + ', fresh Step 7 resolved ' +
      actualCustomerId + '.'
    );
  }

  if (actualLocationId !== expectedLocationId) {
    throw new Error(
      'Configured canary Location mismatch. Expected ' +
      expectedLocationId + ', fresh Step 7 resolved ' +
      actualLocationId + '.'
    );
  }

  if (!actions.length) {
    return {
      status:'CONVERGED_NO_CHANGE',
      execute:false,
      taskId:taskId || '',
      actions:actions
    };
  }

  if (taskId) {
    if (
      actions.indexOf(TMV3_STEP7_ACTION.CREATE_TASK) !== -1 ||
      actions.indexOf(TMV3_STEP7_ACTION.RECREATE_TASK) !== -1
    ) {
      throw new Error(
        'Configured canary duplicate protection stopped CREATE/RECREATE because fresh Step 7 already resolved Task ' +
        taskId + '.'
      );
    }

    return {
      status:'RECONCILE_EXISTING_TASK',
      execute:true,
      taskId:taskId,
      actions:actions
    };
  }

  if (actions.indexOf(expectedAction) === -1) {
    throw new Error(
      'Configured canary expected initial action ' +
      expectedAction + ' is not authorized by fresh Step 7. Fresh plan: ' +
      contract.plan + '.'
    );
  }

  return {
    status:'EXECUTE_INITIAL_CANARY',
    execute:true,
    taskId:'',
    actions:actions
  };
}

function tmv3_runConfiguredCanary_(scope) {
  const normalizedScope = tmv3_norm_(scope || '');

  if (normalizedScope !== 'manual' && normalizedScope !== 'auto') {
    throw new Error(
      'Configured canary scope must be MANUAL or AUTO.'
    );
  }

  tmv3_assertOperationWrite_(scope);

  if (TMV3.MODE !== 'CANARY_WRITE') {
    throw new Error(
      'Configured canary runner is available only in CANARY_WRITE mode.'
    );
  }

  TMV3_CANARY_API_ALLOWANCE_ACTIVE = true;

  try {
    return tmv3_runConfiguredCanaryWithAllowance_(scope);
  } finally {
    TMV3_CANARY_API_ALLOWANCE_ACTIVE = false;
  }
}

function tmv3_runConfiguredCanary_MANUAL() {
  return tmv3_runConfiguredCanary_('MANUAL');
}

function tmv3_runConfiguredCanary_AUTO() {
  return tmv3_runConfiguredCanary_('AUTO');
}

function tmv3_runConfiguredCanaryWithAllowance_(scope) {
  const policy = tmv3_operationPolicy_();
  const vertical = tmv3_clean_(policy.canaryVertical);
  const eventId = tmv3_clean_(policy.canaryEventId);
  const expectedCustomerId = tmv3_clean_(
    policy.canaryExpectedCustomerId
  );
  const expectedLocationId = tmv3_clean_(
    policy.canaryExpectedLocationId
  );
  const expectedAction = tmv3_clean_(
    policy.canaryExpectedAction
  );

  if (
    !TMV3.VERTICALS[vertical] ||
    !eventId ||
    !expectedCustomerId ||
    !expectedLocationId ||
    !expectedAction
  ) {
    throw new Error(
      'Configured canary is incomplete. Vertical, Event ID, Customer ID, Location ID and expected action are required.'
    );
  }

  const eventRecord = tmv3_findFreshEventRecord_(
    vertical,
    eventId
  );

  if (!tmv3_canaryWriteEventAllowed_(eventRecord)) {
    throw new Error(
      'Configured canary event/date/vertical guard rejected the fresh Calendar event.'
    );
  }

  const plans = tmv3_step7FreshPlansForEvent_(
    vertical,
    eventId
  );
  const eligible = plans.filter(function(contract) {
    return !tmv3_clean_(contract.blocker);
  });

  if (eligible.length !== 1) {
    throw new Error(
      'Configured canary requires exactly one unblocked fresh Step 7 contract; found ' +
      eligible.length + '.'
    );
  }

  const contract = tmv3_step7ValidateExecutionContract_(
    eligible[0]
  );
  const decision = tmv3_configuredCanaryContractDecision_(
    contract,
    policy
  );

  tmv3_audit_(
    vertical,
    eventId,
    contract.taskId || '',
    'CONFIGURED_CANARY_GUARD',
    'PASS',
    JSON.stringify({
      scope:tmv3_clean_(scope),
      plan:contract.plan,
      expectedCustomerId:expectedCustomerId,
      expectedLocationId:expectedLocationId,
      expectedAction:expectedAction,
      decision:decision.status,
      actions:decision.actions
    })
  );

  if (!decision.execute) {
    return {
      status:'CONVERGED_NO_CHANGE',
      scope:tmv3_clean_(scope),
      vertical:vertical,
      eventId:eventId,
      taskId:decision.taskId || '',
      writes:0,
      attempts:1
    };
  }

  return tmv3_executeFreshStep7Selection_(
    {
      previousPlan:contract.plan,
      contract:contract
    },
    scope,
    'ALL'
  );
}


function tmv3_runSafeReadyRows_AUTO() {
  return tmv3_runSafeReadyRows_('AUTO');
}

function tmv3_emailListFromEvent_(eventRecord) {
  const source = [
    eventRecord.organizer,
    eventRecord.guests,
    tmv3_stripAuthoredCalendarNotes_(eventRecord.description)
  ].join(' ');

  return tmv3_unique_(
    tmv3_extractEmails_(source)
      .filter(function(email) { return !!email; })
  );
}

function tmv3_businessDaysUntil_(date) {
  const target = new Date(date);
  if (isNaN(target.getTime())) return null;
  target.setHours(0,0,0,0);

  const today = new Date();
  today.setHours(0,0,0,0);
  if (target < today) return -1;

  let count = 0;
  const cursor = new Date(today);
  while (cursor < target && count <= 10) {
    cursor.setDate(cursor.getDate() + 1);
    const day = cursor.getDay();
    if (day !== 0 && day !== 6) count++;
  }
  return count;
}









function tmv3_runCalendarLinksForReady_AUTO() {
  if (!tmv3_operationWritesEnabled_('AUTO')) {
    tmv3_audit_(
      'SYSTEM','','','AUTO_CALENDAR_LINK_CYCLE','GATED',
      'Scheduled Calendar writes remain disabled.'
    );
    return {
      status:'AUTO_WRITES_GATED',
      writes:0
    };
  }

  const events = tmv3_calendarRecords_().filter(function(eventRecord) {
    return tmv3_canaryWriteEventAllowed_(eventRecord);
  });
  const results = [];
  const limit = Number(
    tmv3_operationPolicy_().autoMaxWritesPerRun || 10
  );

  for (
    let i = 0;
    i < events.length && results.length < limit;
    i++
  ) {
    const eventRecord = events[i];
    let contracts = [];

    try {
      contracts = tmv3_step7FreshPlansForEvent_(
        eventRecord.vertical,
        eventRecord.eventId
      ).filter(function(contract) {
        return (
          !tmv3_clean_(contract.blocker) &&
          Number(contract.taskId || 0) > 0 &&
          (contract.actions || []).indexOf(
            TMV3_STEP7_ACTION.VERIFY_CALENDAR_LINKS
          ) !== -1
        );
      });
    } catch (err) {
      results.push({
        eventId:eventRecord.eventId,
        vertical:eventRecord.vertical,
        status:'ATTENTION',
        error:String(err && err.message || err)
      });
      continue;
    }

    for (
      let j = 0;
      j < contracts.length && results.length < limit;
      j++
    ) {
      const contract = contracts[j];

      try {
        const write = tmv3_executeFreshStep7Selection_(
          {
            previousPlan:contract.plan,
            contract:contract
          },
          'AUTO',
          'LINKS'
        );

        results.push({
          eventId:contract.eventId,
          vertical:contract.vertical,
          taskId:contract.taskId,
          status:write.status
        });
      } catch (err) {
        tmv3_audit_(
          contract.vertical,
          contract.eventId,
          contract.taskId || '',
          'AUTO_CALENDAR_LINK',
          'ATTENTION',
          String(err && err.message || err)
        );

        results.push({
          eventId:contract.eventId,
          vertical:contract.vertical,
          taskId:contract.taskId || '',
          status:'ATTENTION',
          error:String(err && err.message || err)
        });
      }
    }
  }

  return {
    status:'COMPLETE',
    writes:results.length,
    results:results
  };
}



/************************************************************
 * TM V3 — PRE-INSPECTION MISSING-DETAIL NOTIFICATION PREVIEW
 *
 * Current policy:
 * - Missing-data checks only. Formatting differences alone never notify.
 * - At least one Sales Order reference anywhere in the event is sufficient.
 * - Recipient is the Calendar event creator, not the Calendar owner.
 * - Delivery is separately gated; send at most once per logical Event ID.
 * - Missing SO is a notification issue, never a task-creation blocker.
 ************************************************************/









// Notification-only evidence check. Does not authorize identity or task writes.





















/**
 * V3-001: Independently repair verified relationships on an existing
 * unscheduled Service/Delivery Task. Never derives schedule or assignments.
 *
 * Explicit expected IDs are mandatory: a Task title is evidence to review,
 * NOT authoritative identity. The existing operation write gates apply.
 */
function tmv3_repairExistingTaskRelationships(taskId, vertical, expected, permission) {
  const id = Number(taskId);
  const kind = tmv3_clean_(vertical);
  expected = expected || {};
  if (!Number.isSafeInteger(id) || id <= 0 ||
      (kind !== 'Service' && kind !== 'Delivery')) {
    throw new Error('Only positive existing Service/Delivery Task IDs are permitted.');
  }
  const scope = permission === 'TASK_FIRST_PRODUCTION'
    ? 'taskfirstproduction'
    : (permission === 'TASK_FIRST_CANARY' &&
      id === Number(TMV3.OPERATIONS.taskFirstCanaryTaskId)
      ? 'taskfirstcanary19083' : 'manual');
  tmv3_assertOperationWrite_(scope, 'TASK_RELATIONSHIP_REPAIR');

  const customerId = tmv3_clean_(expected.customerId);
  const orderId = tmv3_clean_(expected.orderId);
  const locationId = tmv3_clean_(expected.locationId);
  const contactId = tmv3_clean_(expected.contactId);
  if (scope === 'taskfirstcanary19083' && (
    customerId !== String(TMV3.OPERATIONS.taskFirstCanaryExpectedCustomerId) ||
    orderId !== String(TMV3.OPERATIONS.taskFirstCanaryExpectedOrderId) ||
    contactId !== String(TMV3.OPERATIONS.taskFirstCanaryExpectedContactId) ||
    locationId
  )) throw new Error('Canary relationship IDs are not the approved exact tuple.');

  const validId = function(v) { return /^[1-9]\d*$/.test(v); };
  if (!validId(customerId) ||
      [orderId, locationId, contactId].some(function(v) {
        return Boolean(v) && !validId(v);
      }) || !(orderId || locationId || contactId)) {
    throw new Error('Verified Customer ID and at least one valid relationship ID are required.');
  }

  const source = tmv3_rows_(TMV3.SHEETS.TASKS).filter(function(row) {
    return Number(row['Task ID']) === id;
  });
  if (source.length !== 1 ||
      (kind === 'Service' && tmv3_norm_(source[0]['Task Type']).indexOf('service') < 0) ||
      (kind === 'Delivery' && tmv3_norm_(source[0]['Task Type']).indexOf('deliver') < 0)) {
    throw new Error('Source Tasks does not show exactly one correct-vertical Task.');
  }

  const beforeRaw = tmv3_rawTaskById_(id);
  const before = tmv3_normalizeV2TaskModel_(beforeRaw);
  const beforeBy = tmv3_step7RequestedByFromRaw_(beforeRaw);
  const beforeType = tmv3_norm_(before['Task Type']);
  if (tmv3_clean_(before['Task ID']) !== String(id) ||
      (kind === 'Service' && beforeType.indexOf('service') < 0) ||
      (kind === 'Delivery' && beforeType.indexOf('deliver') < 0) ||
      !tmv3_taskIsOpen_(before['Status']) ||
      tmv3_clean_(before['Customer ID']) !== customerId) {
    throw new Error('Task is not OPEN or Task Customer is unverified/mismatched.');
  }

  const fields = [
    {label:'ORDER', id:orderId, actual:tmv3_clean_(before['Order ID'])},
    {label:'LOCATION', id:locationId, actual:tmv3_clean_(before['Location ID'])},
    {label:'REQUESTED_BY', id:contactId, actual:tmv3_clean_(beforeBy.id)}
  ];
  const review = [];
  const changes = [];
  fields.forEach(function(field) {
    if (!field.id) {
      review.push(field.label + (field.actual
        ? '_EXPECTED_ID_UNRESOLVED' : '_UNRESOLVED'));
      return; // Presence alone does not verify identity or ownership.
    }
    if (field.actual && field.actual !== field.id) {
      review.push(field.label + '_CONFLICT');
      return;
    }
    if (field.label === 'REQUESTED_BY' && field.actual &&
        tmv3_norm_(beforeBy.type) !== 'contact') {
      review.push('REQUESTED_BY_TYPE_CONFLICT');
      return;
    }
    try {
      if (field.label === 'ORDER') {
        tmv3_taskFirstFreshOrderOwnershipById_(
          field.id,customerId,tmv3_clean_(before['Name']),kind
        );
      } else if (field.label === 'LOCATION') {
        tmv3_step7FreshLocationOwnership_({
          expectedLocationId:field.id,expectedCustomerId:customerId
        });
      } else {
        const c = tmv3_getContactById_(field.id,customerId);
        if (!c || c.__ownershipVerified !== true) {
          throw new Error('Contact ownership unverified');
        }
      }
      if (!field.actual) changes.push(field);
    } catch (err) {
      if (field.label === 'ORDER' &&
          String(err && err.message || err).indexOf('ORDER_NOT_VERIFIED_ACTIVE') !== -1) {
        throw err; // Contact inferred from this Order must not bypass its lifecycle guard.
      }
      review.push(field.label + '_OWNERSHIP_NOT_VERIFIED');
    }
  });
  if (!changes.length) {
    const complete = !review.length &&
      Boolean(orderId) && Boolean(contactId) && Boolean(locationId) &&
      fields.filter(function(field) { return Boolean(field.id); })
        .every(function(field) { return field.actual === field.id; });
    const result = {
      status:complete ? 'NO_CHANGE_VERIFIED' : 'REVIEW_NO_SAFE_PATCH',
      taskId:id,vertical:kind,writes:0,review:review
    };
    tmv3_audit_(kind,'',String(id),'TASK_FIRST_RELATIONSHIP_VERIFY',
      complete ? 'PASS' : 'REVIEW',JSON.stringify(result));
    return result;
  }

  const payload = {Id:id};
  changes.forEach(function(field) {
    if (field.label === 'ORDER') payload.SalesOrder = {Id:Number(field.id)};
    if (field.label === 'LOCATION') payload.Location = {Id:Number(field.id)};
    if (field.label === 'REQUESTED_BY') {
      payload.RequestedBy = {Id:Number(field.id),Type:'contact'};
    }
  });

  // Guard against concurrent task changes; no blind overwrite or POST.
  const nowRaw = tmv3_rawTaskById_(id);
  const now = tmv3_normalizeV2TaskModel_(nowRaw);
  const nowBy = tmv3_step7RequestedByFromRaw_(nowRaw);
  if (tmv3_clean_(now['Task ID']) !== String(id) ||
      tmv3_clean_(now['Task Type']) !== tmv3_clean_(before['Task Type']) ||
      !tmv3_taskIsOpen_(now['Status']) ||
      tmv3_clean_(now['Customer ID']) !== customerId ||
      tmv3_clean_(now['Order ID']) !== tmv3_clean_(before['Order ID']) ||
      tmv3_clean_(now['Location ID']) !== tmv3_clean_(before['Location ID']) ||
      tmv3_clean_(nowBy.id) !== tmv3_clean_(beforeBy.id) ||
      tmv3_norm_(nowBy.type) !== tmv3_norm_(beforeBy.type)) {
    throw new Error('Task changed after ownership verification. No mutation.');
  }

  tmv3_operationPatchTask_(id,payload,scope);
  const afterRaw = tmv3_rawTaskById_(id);
  const after = tmv3_normalizeV2TaskModel_(afterRaw);
  const afterBy = tmv3_step7RequestedByFromRaw_(afterRaw);
  const verified = changes.every(function(field) {
    if (field.label === 'ORDER') {
      return tmv3_clean_(after['Order ID']) === field.id;
    }
    if (field.label === 'LOCATION') {
      return tmv3_clean_(after['Location ID']) === field.id;
    }
    return tmv3_clean_(afterBy.id) === field.id &&
      tmv3_norm_(afterBy.type) === 'contact';
  }) && fields.filter(function(field) {
    return !changes.some(function(change) { return change.label === field.label; });
  }).every(function(field) {
    const actual = field.label === 'ORDER' ? after['Order ID']
      : field.label === 'LOCATION' ? after['Location ID'] : afterBy.id;
    return tmv3_clean_(actual) === field.actual &&
      (field.label !== 'REQUESTED_BY' ||
        tmv3_norm_(afterBy.type) === tmv3_norm_(beforeBy.type));
  }) && tmv3_clean_(after['Customer ID']) === customerId &&
    tmv3_clean_(after['Task ID']) === String(id) &&
    tmv3_clean_(after['Task Type']) === tmv3_clean_(before['Task Type']) &&
    tmv3_clean_(after['Status']) === tmv3_clean_(before['Status']);
  const result = {
    status:verified ? (review.length ? 'PARTIAL_PATCH_VERIFIED' : 'PATCHED_AND_VERIFIED')
      : 'READBACK_MISMATCH',
    taskId:id,vertical:kind,writes:1,
    patchedFields:changes.map(function(c) {return c.label;}),
    review:review,readbackVerified:verified
  };
  tmv3_audit_(kind,'',String(id),'TASK_FIRST_RELATIONSHIP_PATCH',
    verified ? (review.length ? 'REVIEW' : 'PASS') : 'REVIEW',
    JSON.stringify(result));
  if (!verified) {
    throw new Error('Task-first PATCH read-back failed; do not blindly retry.');
  }
  return result;
}


/** One authoritative GET instead of downloading the entire Work Order report.
 * The order ID, customer ownership and SO number must all agree.
 */
function tmv3_taskFirstFreshOrderOwnershipById_(orderId, customerId, taskName, vertical) {
  if (['Install','Delivery','Service'].indexOf(vertical)<0) throw new Error('ORDER_NOT_VERIFIED_ACTIVE: vertical status policy required.');
  const requiredStatus=vertical==='Service'?'in progress':'approved';
  const id = Number(orderId);
  const cid = tmv3_clean_(customerId);
  if (!Number.isSafeInteger(id) || id <= 0 || !/^[1-9]\d*$/.test(cid)) {
    throw new Error('Invalid task-first Order ownership target.');
  }
  const order = tmv3_fetchJson_(
    TMV3.API_BASE + '/v1/sales-orders/' + encodeURIComponent(id),
    {method:'get'}
  ) || {};
  const fetchedId = tmv3_clean_(tmv3_first_(order,['id','Id']));
  const customer = order.customer || order.Customer || {};
  const fetchedCustomerId = tmv3_clean_(tmv3_first_(customer,['id','Id']));
  const orderNumber = tmv3_clean_(tmv3_first_(order,
    ['orderNumber','OrderNumber','number','Number']));
  // Striven official Order Status IDs: Approved=22, In Progress=25.
  // Only fresh active-order evidence authorizes new relationships.
  const rawStatus = order.status !== undefined ? order.status : order.Status;
  const statusObject = rawStatus && typeof rawStatus === 'object' ? rawStatus : {};
  const statusId = tmv3_clean_(tmv3_first_(statusObject,['id','Id']) ||
    tmv3_first_(order,['statusId','StatusId']) ||
    (/^\d+$/.test(String(rawStatus)) ? rawStatus : ''));
  const statusName = tmv3_norm_(tmv3_first_(statusObject,['name','Name']) ||
    tmv3_first_(order,['statusName','StatusName']) ||
    (typeof rawStatus === 'string' && !/^\d+$/.test(rawStatus) ? rawStatus : ''));
  const known = {'18':'incomplete','19':'quoted','20':'pending approval',
    '21':'declined','22':'approved','23':'canceled','24':'lost','25':'in progress','27':'completed'};
  const normalizedName = statusName === 'cancelled' ? 'canceled' : statusName;
  const statusFromId = known[statusId] || '';
  if ((statusId && !statusFromId) ||
      (statusFromId && normalizedName && statusFromId !== normalizedName) ||
      (statusFromId || normalizedName) !== requiredStatus) {
    throw new Error('ORDER_NOT_VERIFIED_ACTIVE: canceled, closed, unapproved, unknown or conflicting Order status.');
  }
  const titleOrder = tmv3_taskFirstTitleOrderNumber_(taskName);
  if (fetchedId !== String(id) || fetchedCustomerId !== cid ||
      titleOrder.status !== 'UNIQUE' || !orderNumber || titleOrder.number !== orderNumber) {
    throw new Error('Order ID, Customer ownership or Task SO number disagrees with fresh Striven Order.');
  }
  const contact = order.contact || order.Contact || {};
  const contactId = tmv3_clean_(tmv3_first_(contact,['id','Id']));
  return {verified:true,orderId:String(id),orderNumber:orderNumber,customerId:cid,
    status:statusFromId || normalizedName,contactId:/^[1-9]\d*$/.test(contactId) ? contactId : '',
    locationId:tmv3_clean_((order.ShipToLocation || order.shipToLocation || {}).Id || (order.ShipToLocation || order.shipToLocation || {}).id)};
}


// Google Apps Script: Task title is the Order-number key, never the Customer key.
function tmv3_taskFirstTitleOrderNumber_(title) {
  const text=String(title || '').replace(/https?:\/\/\S+/gi,' ');
  const pattern=/\bSO\s*#\s*(\d{4,9}(?!\d)(?:\s*(?:[/&,]|\band\b)\s*(?:SO\s*#\s*)?\d{4,9}(?!\d))*)/gi;
  const values=[];let match;
  while((match=pattern.exec(text))!==null) {
    (match[1].match(/\d{4,9}/g)||[]).forEach(function(n){if(values.indexOf(n)<0)values.push(n);});
  }
  return {status:values.length===1?'UNIQUE':values.length?'AMBIGUOUS':'MISSING',
    number:values.length===1?values[0]:'',numbers:values};
}

// Google Apps Script: one complete, customer-scoped read instead of the whole Location report.


// Only an old, explicit validation rejection can be recovered; ambiguous results remain blocked.




// Cache the authoritative schedule; never infer a twelve-hour shift or write Task dates.



// Original: 60_Automation
/************************************************************
 * TM V3 — MENU + AUTOMATION PARITY
 ************************************************************/

function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const morning = ss.getSheetByName(TMV3.SHEETS.MORNING);

  if (morning) ss.setActiveSheet(morning);

  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Task Mapping V3')
    .addSubMenu(
      ui.createMenu('Operations')
        .addItem('Step 1 — Refresh Calendar Intake', 'tmv3_step1CalendarRun')
        .addItem('Step 2 — Refresh Eligibility', 'tmv3_step2CalendarRun')
        .addItem('Step 3 — Resolve Business Anchors', 'tmv3_step3BusinessAnchorRunCached')
        .addItem('Step 4 — Resolve Customer / Location', 'tmv3_step4IdentityRunCached')
        .addItem('Step 5 — Resolve Tasks', 'tmv3_step5TaskResolutionRunCached')
        .addItem('Step 6 — Decide Task Action', 'tmv3_step6TaskDecisionRunCached')
        .addItem('Refresh Sources + Mapping', 'tmv3_shadowRun')
        .addItem('Refresh Mapping From Cache', 'tmv3_shadowMapFromCache')
        .addItem('Refresh Morning Ops', 'tmv3_refreshMorningOps')
        .addSeparator()
        .addItem('Run Safe READY Rows Now', 'tmv3_runSafeReadyRows_MANUAL')
    )
    .addSubMenu(
      ui.createMenu('Selected Appointment')
        .addItem('Preview Selected Action', 'tmv3_previewSelectedAction')
        .addSeparator()
        .addItem('Sync ALL Safe Fields', 'tmv3_syncSelectedAppointment')
        .addItem('Push Date / Time', 'tmv3_pushSelectedDates')
        .addItem('Push Relationships / IDs', 'tmv3_pushSelectedRelationships')
        .addItem('Reconcile Assignee', 'tmv3_reconcileSelectedAssignee')
        .addItem('Fix Calendar Links', 'tmv3_fixSelectedCalendarLinks')
        .addSeparator()
        .addItem('Create / Recreate Task', 'tmv3_createOrRecreateSelectedTask')
    )
    .addSubMenu(
      ui.createMenu('Reminders')
        .addItem('Send Missing SO Reminder Emails Now', 'tmv3_sendInstallMissingSoRemindersNow')
        .addItem('Preview Pre-Inspection Missing Detail Emails', 'tmv3_previewPreInspectionMissingDetailEmails')
        .addItem('Run Pre-Inspection Automation Now', 'tmv3_runPreInspectionAutomation')
    )
    .addSubMenu(
      ui.createMenu('Automation')
        .addItem('Install Managed Triggers', 'tmv3_installTriggers')
        .addItem('Remove Managed Triggers', 'tmv3_removeTriggers')
        .addItem('List Managed Triggers', 'tmv3_logTriggers')
    )
    .addSubMenu(
      ui.createMenu('Diagnostics')
        .addItem('Run Full Shadow Verification', 'tmv3_shadowRun')
        .addItem('Health Check', 'tmv3_healthCheck')
        .addItem('Run Live Hardening Regression', 'tmv3_liveHardeningRegression')
        .addItem('Acknowledge Selected Exception', 'tmv3_acknowledgeSelectedMorningOps')
        .addItem('Clear Selected Acknowledgement', 'tmv3_clearAcknowledgementSelectedMorningOps')
    )
    .addToUi();
}

function tmv3_managedTriggerHandlers_() {
  return [
    'tmv3_allVerticalAutomationTick',
    'tmv3_dailySourceRefresh',
    'tmv3_scheduledSheetRefresh',
    'tmv3_scheduledOperations',
    'tmv3_refreshLinksSlot_0800',
    'tmv3_refreshLinksSlot_1000',
    'tmv3_refreshLinksSlot_1200',
    'tmv3_refreshLinksSlot_1400',
    'tmv3_refreshLinksSlot_1600',
    'tmv3_refreshLinksSlot_1800',
    'tmv3_installReminderCheck',
    'tmv3_calendarEventUpdated',
    'tmv3_calendarReconciliationFallback',
    // Older V3 handler retained for cleanup during upgrade.
    'tmv3_scheduledShadow'
  ];
}

function tmv3_assertCalendarTriggerAuthorization_(calendarIds) {
  const ids = (calendarIds || []).map(tmv3_clean_).filter(Boolean);
  if (!ids.length) return true;

  // Explicitly require the scopes needed by this installation flow.
  // When the Calendar scope is missing, Apps Script stops here and presents
  // the user authorization flow before any managed trigger is deleted.
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/script.scriptapp'
  ]);

  try {
    const probe = CalendarApp.getCalendarById(ids[0]);
    if (!probe) {
      throw new Error(
        'Configured Calendar is not accessible: ' + ids[0]
      );
    }
  } catch (err) {
    const message =
      'Managed trigger install blocked before deleting existing triggers. ' +
      'Calendar authorization/access is required. ' +
      String(err && err.message || err);

    tmv3_audit_(
      'SYSTEM','','','INSTALL_TRIGGER_PREFLIGHT','BLOCKED',
      message
    );

    throw new Error(message);
  }

  return true;
}





function tmv3_calendarStageRefresh_(reason) {
  const stage = tmv3_executionStage_();

  if (stage <= 1) {
    return tmv3_step1CalendarRun(reason || 'STAGE1_REFRESH');
  }

  if (stage === 2) {
    return tmv3_step2CalendarRun(reason || 'STAGE2_REFRESH');
  }

  if (stage === 3) {
    return tmv3_step3BusinessAnchorRunCached(
      reason || 'STAGE3_ANCHOR_REFRESH'
    );
  }

  if (stage === 4) {
    return tmv3_step4IdentityRunCached(
      reason || 'STAGE4_IDENTITY_REFRESH'
    );
  }

  if (stage === 5) {
    return tmv3_step5TaskResolutionRunCached(
      reason || 'STAGE5_TASK_REFRESH'
    );
  }

  if (stage === 6) {
    return tmv3_step6TaskDecisionRunCached(
      reason || 'STAGE6_TASK_DECISION_REFRESH'
    );
  }

  if (stage === 7) {
    // Keep routine Calendar-driven refreshes on the proven Step 6 view.
    // Full Step 7 performs many fresh Task/Contact reads and is deliberately
    // invoked through the guarded verification runner during this stage.
    return tmv3_step6TaskDecisionRunCached(
      reason || 'STAGE7_RECONCILIATION_VIEW_REFRESH'
    );
  }

  return tmv3_shadowMapFromCache();
}

function tmv3_sourceRefreshState_(propertyKey) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(propertyKey);
  const ms = raw ? new Date(raw).getTime() : 0;

  return {
    lastRefresh:raw || '',
    ageMinutes:ms ? Math.max(0, (Date.now() - ms) / 60000) : null
  };
}

function tmv3_operationalSourceRefreshState_() {
  return tmv3_sourceRefreshState_(
    'TMV3_LAST_OPERATIONAL_SOURCE_REFRESH'
  );
}

function tmv3_masterSourceRefreshState_() {
  return tmv3_sourceRefreshState_(
    'TMV3_LAST_MASTER_SOURCE_REFRESH'
  );
}

function tmv3_operationalSourceMinAgeMinutes_() {
  return Math.max(
    1,
    Number(
      TMV3.OPERATIONS &&
      TMV3.OPERATIONS.sourceRefreshMinMinutes ||
      120
    )
  );
}

function tmv3_masterSourceMinAgeMinutes_() {
  return Math.max(
    tmv3_operationalSourceMinAgeMinutes_(),
    Number(
      TMV3.OPERATIONS &&
      TMV3.OPERATIONS.masterSourceRefreshMinMinutes ||
      360
    )
  );
}

function tmv3_refreshOperationalSourcesIfNeeded_(maxAgeMinutes, force) {
  const operationalState = tmv3_operationalSourceRefreshState_();
  const masterState = tmv3_masterSourceRefreshState_();
  const operationalMinimum = tmv3_operationalSourceMinAgeMinutes_();
  const masterMinimum = tmv3_masterSourceMinAgeMinutes_();
  const requestedOperationalAge = Math.max(
    1,
    Number(maxAgeMinutes || operationalMinimum)
  );
  const operationalMaxAge =
    force === true
      ? requestedOperationalAge
      : Math.max(operationalMinimum, requestedOperationalAge);

  const operationalStale =
    force === true ||
    operationalState.ageMinutes === null ||
    operationalState.ageMinutes >= operationalMaxAge;

  const masterStale =
    force === true ||
    masterState.ageMinutes === null ||
    masterState.ageMinutes >= masterMinimum;

  if (!operationalStale && !masterStale) {
    return {
      status:'CACHE_FRESH',
      refreshed:false,
      master:{
        refreshed:false,
        lastRefresh:masterState.lastRefresh,
        ageMinutes:masterState.ageMinutes,
        effectiveMaxAgeMinutes:masterMinimum
      },
      transactions:{
        refreshed:false,
        lastRefresh:operationalState.lastRefresh,
        ageMinutes:operationalState.ageMinutes,
        requestedMaxAgeMinutes:requestedOperationalAge,
        effectiveMaxAgeMinutes:operationalMaxAge
      }
    };
  }

  tmv3_resetRuntimeMetrics_();

  const props = PropertiesService.getScriptProperties();
  let masterResult = null;
  let transactionResult = null;
  let masterRefreshedAt = masterState.lastRefresh;
  let operationalRefreshedAt = operationalState.lastRefresh;

  if (masterStale) {
    masterResult = tmv3_refreshMasterSources_();
    masterRefreshedAt = new Date().toISOString();
    props.setProperty(
      'TMV3_LAST_MASTER_SOURCE_REFRESH',
      masterRefreshedAt
    );
  }

  if (operationalStale) {
    transactionResult = tmv3_refreshTransactionSources_();
    operationalRefreshedAt = new Date().toISOString();
    props.setProperty(
      'TMV3_LAST_OPERATIONAL_SOURCE_REFRESH',
      operationalRefreshedAt
    );
  }

  return {
    status:'REFRESHED',
    refreshed:true,
    master:{
      refreshed:masterStale,
      refreshedAt:masterRefreshedAt,
      effectiveMaxAgeMinutes:masterMinimum,
      result:masterResult
    },
    transactions:{
      refreshed:operationalStale,
      refreshedAt:operationalRefreshedAt,
      requestedMaxAgeMinutes:requestedOperationalAge,
      effectiveMaxAgeMinutes:operationalMaxAge,
      result:transactionResult
    },
    api:tmv3_runtimeMetrics_()
  };
}

function tmv3_stage7ConfiguredCanaryEnabled_(
  stage,
  mode,
  automationWritesEnabled
) {
  return (
    Number(stage) === 7 &&
    String(mode || '') === 'CANARY_WRITE' &&
    automationWritesEnabled === true
  );
}

function tmv3_stage7AutomaticCanaryCycle_(reason) {
  const stage = tmv3_executionStage_();
  const policy = tmv3_operationPolicy_();
  let sources = null;
  let sourceError = '';
  let mapped = null;

  try {
    sources = tmv3_refreshOperationalSourcesIfNeeded_(
      tmv3_operationalSourceMinAgeMinutes_(),
      false
    );
  } catch (err) {
    sourceError = String(err && err.message || err);
    sources = {
      status:'FAILED_KEEP_LAST_GOOD',
      refreshed:false,
      error:sourceError
    };

    tmv3_audit_(
      'SYSTEM','','','STAGE7_SOURCE_REFRESH','ATTENTION',
      JSON.stringify({
        reason:tmv3_clean_(reason),
        error:sourceError
      })
    );
  }

  try {
    mapped = tmv3_calendarStageRefresh_(
      reason || 'STAGE7_AUTOMATIC_CANARY'
    );
  } catch (err) {
    mapped = {
      status:'ATTENTION',
      error:String(err && err.message || err)
    };

    tmv3_audit_(
      'SYSTEM','','','STAGE7_CANARY_MAPPING','ATTENTION',
      JSON.stringify({
        reason:tmv3_clean_(reason),
        error:mapped.error
      })
    );
  }

  if (sourceError) {
    return {
      status:'SOURCE_REFRESH_ATTENTION',
      stage:stage,
      sources:sources,
      mapped:mapped,
      writes:{
        status:'AUTO_WRITES_GATED_SOURCE_REFRESH_FAILED',
        writes:0
      }
    };
  }

  if (
    !tmv3_stage7ConfiguredCanaryEnabled_(
      stage,
      TMV3.MODE,
      policy.automationWritesEnabled === true
    )
  ) {
    return {
      status:'CALENDAR_STAGE_GATED',
      stage:stage,
      sources:sources,
      mapped:mapped,
      writes:{
        status:'AUTO_WRITES_GATED',
        writes:0
      }
    };
  }

  let writes;

  try {
    writes = tmv3_runConfiguredCanary_AUTO();

    tmv3_audit_(
      'SYSTEM','','','STAGE7_AUTOMATIC_CANARY','PASS',
      JSON.stringify({
        reason:tmv3_clean_(reason),
        status:tmv3_clean_(writes && writes.status),
        taskId:tmv3_clean_(writes && writes.taskId)
      })
    );
  } catch (err) {
    writes = {
      status:'ATTENTION',
      writes:0,
      error:String(err && err.message || err)
    };

    tmv3_audit_(
      'SYSTEM','','','STAGE7_AUTOMATIC_CANARY','ATTENTION',
      JSON.stringify({
        reason:tmv3_clean_(reason),
        error:writes.error
      })
    );
  }

  return {
    status:'STAGE7_AUTOMATIC_CANARY_CYCLE',
    stage:stage,
    sources:sources,
    mapped:mapped,
    writes:writes
  };
}

function tmv3_runProductionAutomationCycle_(reason, options) {
  options = options || {};

  if (tmv3_executionStage_() <= 7) {
    return {
      status:'CALENDAR_STAGE_GATED',
      stage:tmv3_executionStage_(),
      mapped:tmv3_calendarStageRefresh_(reason || 'AUTOMATION_STAGE_GATED'),
      writes:{ status:'CALENDAR_STAGE_GATED', writes:0 }
    };
  }

  if (!tmv3_operationWritesEnabled_('AUTO')) {
    return {
      status:'AUTO_WRITES_GATED',
      stage:tmv3_executionStage_(),
      sources:tmv3_refreshOperationalSourcesIfNeeded_(
        Number(options.maxSourceAgeMinutes || 15),
        options.forceSourceRefresh === true
      ),
      mapped:tmv3_shadowMapFromCache(),
      writes:{ status:'AUTO_WRITES_GATED', writes:0 }
    };
  }

  const sources = tmv3_refreshOperationalSourcesIfNeeded_(
    Number(options.maxSourceAgeMinutes || 15),
    options.forceSourceRefresh === true
  );
  const mapped = tmv3_shadowMapFromCache();
  const writes = tmv3_runSafeReadyRows_AUTO();

  const result = {
    status:'AUTOMATION_CYCLE_COMPLETE',
    reason:tmv3_clean_(reason || 'AUTO'),
    stage:tmv3_executionStage_(),
    sources:sources,
    mapped:mapped,
    writes:writes
  };

  tmv3_audit_(
    'SYSTEM','','','PRODUCTION_AUTOMATION_CYCLE',
    'PASS',
    JSON.stringify({
      reason:result.reason,
      sourceStatus:sources.status,
      writeStatus:writes.status,
      writes:Number(writes.writes || 0),
      attempts:Number(writes.attempts || 0)
    })
  );

  return result;
}





function tmv3_calendarReconciliationFallback() { return {status:'RETIRED_USE_STANDARD_DAILY_SLOTS',businessWrites:0}; }

function tmv3_dailySourceRefresh() {
  tmv3_assertShadow_();

  if (tmv3_executionStage_() <= 2) {
    return tmv3_calendarStageRefresh_('DAILY_CALENDAR_STAGE_REFRESH');
  }

  if (tmv3_executionStage_() === 3) {
    const sources = tmv3_step3RefreshAnchorSources_();
    const mapped = tmv3_step3BusinessAnchorRunCached(
      'DAILY_STAGE3_ANCHOR_REFRESH'
    );
    return { sources: sources, mapped: mapped };
  }

  if (tmv3_executionStage_() === 4) {
    const sources = tmv3_step4RefreshIdentitySources_();
    const mapped = tmv3_step4IdentityRunCached(
      'DAILY_STAGE4_IDENTITY_REFRESH'
    );
    return { sources: sources, mapped: mapped };
  }

  if (tmv3_executionStage_() === 5) {
    const sources = tmv3_step5RefreshTaskSources_();
    const mapped = tmv3_step5TaskResolutionRunCached(
      'DAILY_STAGE5_TASK_REFRESH'
    );
    return { sources: sources, mapped: mapped };
  }

  if (tmv3_executionStage_() === 6) {
    const sources = tmv3_step5RefreshTaskSources_();
    const mapped = tmv3_step6TaskDecisionRunCached(
      'DAILY_STAGE6_TASK_DECISION_REFRESH'
    );
    return { sources: sources, mapped: mapped };
  }

  if (tmv3_executionStage_() === 7) {
    // Stage 7 still depends on Stage 3/4 source integrity. Refresh the complete
    // shared snapshot here instead of refreshing Task reports alone.
    const sources = tmv3_refreshOperationalSourcesIfNeeded_(
      tmv3_operationalSourceMinAgeMinutes_(),
      true
    );
    const mapped = tmv3_step6TaskDecisionRunCached(
      'DAILY_STAGE7_RECONCILIATION_VIEW_REFRESH'
    );
    return { sources: sources, mapped: mapped };
  }

  const result = tmv3_refreshSources();
  tmv3_audit_(
    'SYSTEM','','','DAILY_SOURCE_REFRESH','PASS',JSON.stringify(result)
  );
  tmv3_shadowMapFromCache_({
    customers: result.customers,
    locations: result.locations,
    contacts: result.contacts,
    orders: result.orders,
    tasks: result.tasks,
    source: 'FRESH_DAILY'
  });
  return result;
}

function tmv3_scheduledShadow() {
  // Compatibility alias. Scheduled production uses tmv3_scheduledOperations.
  return tmv3_scheduledOperations();
}


/**
 * V3-001/P0 relationship-only lane. Runs independently of all Calendar
 * stages and cannot enable other V3 mutation scopes.
 *
 * API budget: one candidate/task per 60-minute slot, one fresh Task GET,
 * one fresh Work Order GET, and only when applicable a verified Contact GET.
 * The existing relationship writer independently repeats identity checks,
 * performs a pre-mutation drift GET and requires field-level GET readback.
 * No creates, no dates, no assignments, no Calendar writes.
 */
function tmv3_taskFirstProductionCycle_() {
  const policy = tmv3_operationPolicy_();
  if (!tmv3_operationWritesEnabled_('taskfirstproduction')) {
    return {status:'TASK_FIRST_PRODUCTION_GATED',writes:0};
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return {status:'LOCK_BUSY',writes:0};

  try {
    const props = PropertiesService.getScriptProperties();
    // Explicit one-time backlog: independent of the routine hourly allowance.
    const backfill = tmv3_serviceDeliveryOrderBackfill_(props);
    if (backfill) return backfill;
    const lastKey = 'TMV3_TASK_FIRST_PRODUCTION_LAST_ATTEMPT_MS';
    const cursorKey = 'TMV3_TASK_FIRST_PRODUCTION_CURSOR';
    const last = Number(props.getProperty(lastKey) || 0);
    const now = Date.now();
    const interval = Math.max(60,
      Number(policy.taskFirstProductionIntervalMinutes || 60)) * 60000;
    if (last && now - last < interval) {
      return {status:'THROTTLED',writes:0};
    }

    let detailWriteStarted=false;
    // The existing native timer owns both lanes and their shared one-hour budget.
    // Calendar-backed existing Tasks get full safe details; unscheduled Tasks
    // continue through the relationship-only fallback below.
    function detailActions(plan) {
      return (plan.actions || []).filter(function(action) {
        return ['PATCH_LOCATION','PATCH_ORDER','PATCH_REQUESTED_BY',
          'PATCH_DATES','PATCH_ASSIGNMENTS'].indexOf(action) >= 0;
      });
    }
    function executeDetails(plan, record, guardKey) {
      plan = tmv3_step7ValidateExecutionContract_(plan);
      if (['Service','Delivery'].indexOf(plan.vertical)<0 || !plan.taskId || plan.blocker ||
          (plan.actions || []).some(function(a){return /CREATE|RECREATE/.test(a);})) {
        throw new Error('EXISTING_SERVICE_DELIVERY_DETAILS_ONLY');
      }
      if (!/^[1-9]\d*$/.test(String(plan.expectedOrderId || '')) ||
          !/^[1-9]\d*$/.test(String(plan.expectedLocationId || '')) ||
          !/^[1-9]\d*$/.test(String(plan.expectedRequestedById || '')) ||
          tmv3_norm_(plan.expectedRequestedByType)!=='contact' ||
          !/^\d+$/.test(String(record.orderNumber || ''))) {
        throw new Error('TASK_DETAILS_REQUIRED_RELATIONSHIPS_UNRESOLVED');
      }
      tmv3_taskFirstFreshOrderOwnershipById_(plan.expectedOrderId,
        plan.expectedCustomerId,'SO#'+record.orderNumber,plan.vertical);
      tmv3_step7FreshLocationOwnership_(plan);
      const contact=tmv3_getContactById_(plan.expectedRequestedById,plan.expectedCustomerId);
      if (!contact || contact.__ownershipVerified!==true) throw new Error('TASK_DETAILS_CONTACT_OWNERSHIP');
      const beforeRaw=tmv3_rawTaskById_(plan.taskId);
      const before=tmv3_normalizeV2TaskModel_(beforeRaw);
      const requested=tmv3_step7RequestedByFromRaw_(beforeRaw);
      if (String(before['Task ID'])!==String(plan.taskId) ||
          !tmv3_taskIsOpen_(before['Status']) ||
          tmv3_clean_(before['Customer ID'])!==String(plan.expectedCustomerId) ||
          (plan.vertical==='Service' && tmv3_norm_(before['Task Type']).indexOf('service')<0) ||
          (plan.vertical==='Delivery' && tmv3_norm_(before['Task Type']).indexOf('deliver')<0)) {
        throw new Error('TASK_DETAILS_TASK_IDENTITY_CHANGED');
      }
      [['Order ID',plan.expectedOrderId],['Location ID',plan.expectedLocationId]].forEach(function(pair){
        if (tmv3_clean_(before[pair[0]]) && tmv3_clean_(before[pair[0]])!==String(pair[1])) {
          throw new Error('TASK_DETAILS_EXISTING_RELATIONSHIP_CONFLICT');
        }
      });
      if (tmv3_clean_(requested.id) && (tmv3_clean_(requested.id)!==String(plan.expectedRequestedById) ||
          tmv3_norm_(requested.type)!=='contact')) throw new Error('TASK_DETAILS_REQUESTED_BY_CONFLICT');
      const actions=detailActions(plan);
      const scoped=Object.assign({},plan,{actions:actions});
      const payload=tmv3_step7RelationshipPayload_(scoped);
      if (actions.indexOf('PATCH_DATES')>=0) {
        const start=new Date(plan.expectedStart),due=new Date(plan.expectedDue);
        if (!Number.isFinite(start.getTime()) || !Number.isFinite(due.getTime()) || due<=start ||
            !tmv3_sameMinute_(start,record.start) || !tmv3_sameMinute_(due,record.end)) {
          throw new Error('TASK_DETAILS_CALENDAR_DATES_UNVERIFIED');
        }
        Object.assign(payload,tmv3_step7DatePayload_(scoped));
      }
      const desired=tmv3_step7DesiredAssignmentFromContract_(plan);
      const freshRecord=tmv3_findFreshEventRecord_(plan.vertical,plan.eventId);
      if (!record.fingerprint || freshRecord.fingerprint!==record.fingerprint) {
        throw new Error('TASK_DETAILS_CALENDAR_CHANGED_NO_WRITE');
      }
      const assignmentsBefore=tmv3_taskAssignmentsFromRaw_(beforeRaw);
      props.setProperty(guardKey,JSON.stringify({state:'INTENT',taskId:plan.taskId,
        fingerprint:record.fingerprint,at:now}));
      tmv3_taskFirstProductionCycle_.detailContext={vertical:plan.vertical,
        taskId:Number(plan.taskId),payload:payload,desiredAssignment:desired};
      try {
        if (Object.keys(payload).length>1) {detailWriteStarted=true;tmv3_operationPatchTask_(plan.taskId,payload,'SERVICEDELIVERYDETAILS');}
        if (actions.indexOf('PATCH_ASSIGNMENTS')>=0) {
          detailWriteStarted=true;
          tmv3_reconcileTaskAssignments_(record,plan.taskId,'SERVICEDELIVERYDETAILS',{desiredAssignment:desired});
        }
        const afterRaw=tmv3_rawTaskById_(plan.taskId);
        const after=tmv3_normalizeV2TaskModel_(afterRaw);
        const assignmentsAfter=tmv3_taskAssignmentsFromRaw_(afterRaw);
        if (String(after['Task ID'])!==String(plan.taskId) ||
            tmv3_clean_(after['Customer ID'])!==String(plan.expectedCustomerId) ||
            after['Task Type']!==before['Task Type'] || after['Status']!==before['Status'] ||
            after['Name']!==before['Name'] ||
            JSON.stringify(afterRaw.description || afterRaw.Description || '')!==
              JSON.stringify(beforeRaw.description || beforeRaw.Description || '') ||
            !assignmentsBefore.filter(function(a){return !(a.type==='pool' && Number(a.id)===4 && desired.employeeIds.length);})
              .every(function(a){return assignmentsAfter.some(function(b){return a.type===b.type && Number(a.id)===Number(b.id);});})) {
          throw new Error('TASK_DETAILS_UNRELATED_FIELDS_READBACK_MISMATCH');
        }
        const afterPlan=tmv3_step7FreshExecutionContract_(plan.vertical,plan.eventId,plan.taskId);
        if (afterPlan.blocker || detailActions(afterPlan).length ||
            String(afterPlan.expectedOrderId)!==String(plan.expectedOrderId) ||
            String(afterPlan.expectedLocationId)!==String(plan.expectedLocationId) ||
            String(afterPlan.expectedRequestedById)!==String(plan.expectedRequestedById) ||
            String(afterPlan.expectedStart)!==String(plan.expectedStart) ||
            String(afterPlan.expectedDue)!==String(plan.expectedDue)) {
          throw new Error('TASK_DETAILS_FRESH_PLAN_READBACK_NOT_VERIFIED');
        }
        props.setProperty(guardKey,JSON.stringify({state:'VERIFIED',taskId:plan.taskId,
          fingerprint:record.fingerprint,at:Date.now()}));
        tmv3_upsertTaskCacheRow_(after);
        return {status:'TASK_DETAILS_PATCHED_AND_VERIFIED',taskId:plan.taskId,
          vertical:plan.vertical,eventId:plan.eventId,actions:actions,writes:1,readbackVerified:true};
      } finally {delete tmv3_taskFirstProductionCycle_.detailContext;}
    }
    function detailCycle() {
    const eventQueues=['Service','Delivery'].map(function(vertical){
      const cfg=TMV3.VERTICALS && TMV3.VERTICALS[vertical];
      return cfg ? tmv3_rows_(cfg.sheet).filter(function(row){return !!tmv3_clean_(row['Event ID']);})
        .map(function(row){return {vertical:vertical,eventId:tmv3_clean_(row['Event ID'])};}) : [];
    });
    const detailEvents=[],seenEvents={};
    for(let q=0;q<Math.max(eventQueues[0].length,eventQueues[1].length);q++) {
      eventQueues.forEach(function(queue){
        const event=queue[q];if(!event)return;
        const key=event.vertical+'|'+event.eventId;
        if(!seenEvents[key]){seenEvents[key]=true;detailEvents.push(event);}
      });
    }
    const detailCursorKey='TMV3_SD_DETAILS_EVENT_CURSOR';
    let detailCursor=Number(props.getProperty(detailCursorKey)||0);
    if(!Number.isSafeInteger(detailCursor)||detailCursor<0)detailCursor=0;
    for(let j=0;j<Math.min(4,detailEvents.length);j++) {
      const item=detailEvents[(detailCursor+j)%detailEvents.length];
      props.setProperty(detailCursorKey,String((detailCursor+j+1)%detailEvents.length));
      try {
        const record=tmv3_findFreshEventRecord_(item.vertical,item.eventId);
        const plans=tmv3_step7FreshPlansForEvent_(item.vertical,item.eventId);
        for(let k=0;k<plans.length;k++) {
          const plan=plans[k];
          if(!plan.taskId || plan.blocker || (plan.actions||[]).some(function(a){return /CREATE|RECREATE/.test(a);}))continue;
          const guardKey='TMV3_SD_DETAILS_'+tmv3_hash_(item.vertical+'|'+item.eventId+'|'+plan.taskId);
          const guard=JSON.parse(props.getProperty(guardKey)||'null');
          if(!detailActions(plan).length) {
            if(guard && guard.state==='INTENT')props.setProperty(guardKey,JSON.stringify({state:'VERIFIED_AFTER_RECONCILIATION',taskId:plan.taskId,at:now}));
            continue;
          }
          if(guard && guard.state==='INTENT') {
            tmv3_audit_(item.vertical,item.eventId,String(plan.taskId),'TASK_DETAILS_AUTO','REVIEW','UNCERTAIN_PRIOR_WRITE_RECONCILE_BEFORE_RETRY');
            continue;
          }
          props.setProperty(lastKey,String(now));
          const result=executeDetails(plan,record,guardKey);
          tmv3_audit_(item.vertical,item.eventId,String(plan.taskId),'TASK_DETAILS_AUTO','PASS',JSON.stringify(result));
          return result;
        }
      } catch(err) {
        tmv3_audit_(item.vertical,item.eventId,'','TASK_DETAILS_AUTO',detailWriteStarted?'ATTENTION':'REVIEW',String(err && err.message || err));
        // Once an event transaction was reserved, do not spend another write slot.
        if(Number(props.getProperty(lastKey)||0)===now)return {status:'TASK_DETAILS_ATTENTION',writes:detailWriteStarted?null:0,requiresReadback:detailWriteStarted};
      }
    }

      return null;
    }

    // Read cached business anchors at zero Striven API cost.
    const orderMap = {};
    tmv3_rows_(TMV3.SHEETS.ORDERS).forEach(function(row) {
      const number = tmv3_clean_(row['Order Number']);
      const id = tmv3_clean_(row['Order ID']);
      if (!number || !/^[1-9]\d*$/.test(id)) return;
      if (!orderMap[number]) orderMap[number] = [];
      orderMap[number].push(row);
    });
    const candidates = tmv3_rows_(TMV3.SHEETS.TASKS)
      .filter(function(row) {
        const type = tmv3_norm_(row['Task Type']);
        const status = tmv3_norm_(row['Status']);
        const backfillState=JSON.parse(props.getProperty('TMV3_SD_ORDER_BACKFILL_20261020_'+row['Task ID']) || 'null');
        if (backfillState && ['INTENT','UNCERTAIN'].indexOf(backfillState.state)>=0) return false;
        return status === 'open' &&
          (type.indexOf('service') !== -1 || type.indexOf('delivery') !== -1) &&
          Number.isSafeInteger(Number(row['Task ID'])) && Number(row['Task ID']) > 0 &&
          (!tmv3_clean_(row['Order ID']) || !tmv3_clean_(row['Contact ID'])) &&
          /\bSO\s*#\s*\d{4,9}\b/i.test(tmv3_clean_(row['Name']));
      })
      .sort(function(a,b) {
        return Number(b['Task ID'] || 0)-Number(a['Task ID'] || 0);
      });
    const unlinked=candidates.filter(function(row){
      if(tmv3_clean_(row['Order ID']))return false;
      const title=tmv3_taskFirstTitleOrderNumber_(row.Name);
      if(title.status!=='UNIQUE')return false;
      const vertical=tmv3_norm_(row['Task Type']).indexOf('deliver')>=0?'Delivery':'Service';
      const matches=(orderMap[title.number] || []).filter(function(order){return tmv3_taskFirstOrderMatchesVertical_(order,vertical);});
      const ids=[];matches.forEach(function(order){const id=tmv3_clean_(order['Order ID']);if(ids.indexOf(id)<0)ids.push(id);});
      // Unresolved or known-closed links cannot starve Calendar-backed details.
      // A refreshed active Order automatically becomes eligible again.
      return ids.length===1 && matches.some(function(order){
        const status=tmv3_norm_(order.Status);
        return !status || status==='approved' || status==='in progress';
      });
    });
    if (unlinked.length) candidates.splice(0,candidates.length,...unlinked);
    else {const details=detailCycle();if(details)return details;}
    if (!candidates.length) {
      props.setProperty(lastKey,String(now));
      return {status:'NO_CACHED_CANDIDATES',writes:0};
    }
    let cursor = Number(props.getProperty(cursorKey) || 0);
    if (!Number.isSafeInteger(cursor) || cursor < 0) cursor = 0;
    cursor = cursor % candidates.length;
    const source = candidates[cursor];
    props.setProperty(cursorKey,String((cursor + 1) % candidates.length));
    // Throttle even after failure; repeated exceptions must not burn API calls.
    props.setProperty(lastKey,String(now));

    const taskId = Number(source['Task ID']);
    const title = tmv3_clean_(source['Name']);
    const titleOrder = tmv3_taskFirstTitleOrderNumber_(title);
    if (titleOrder.status !== 'UNIQUE') {
      tmv3_audit_('SYSTEM','',String(taskId),'TASK_FIRST_AUTO','REVIEW','TASK_TITLE_SO_'+titleOrder.status);
      return {status:'REVIEW_TASK_TITLE_SO',taskId:taskId,writes:0};
    }
    const vertical = tmv3_norm_(source['Task Type']).indexOf('delivery') !== -1
      ? 'Delivery' : 'Service';
    const match = (orderMap[titleOrder.number] || []).filter(function(order) {
      return tmv3_taskFirstOrderMatchesVertical_(order, vertical);
    }).filter(function(order,index,all){return all.findIndex(function(other){
      return tmv3_clean_(other['Order ID'])===tmv3_clean_(order['Order ID']);
    })===index;});
    if (match.length !== 1) {
      const reason = match.length ? 'AMBIGUOUS_WORK_ORDER' : 'WORK_ORDER_NOT_IN_CACHE';
      tmv3_audit_(vertical,'',String(taskId),'TASK_FIRST_AUTO','REVIEW',reason);
      return {status:'REVIEW',taskId:taskId,reason:reason,writes:0};
    }

    let repairStarted = false;
    try {
      const live = tmv3_rawTaskById_(taskId);
      const normalized = tmv3_normalizeV2TaskModel_(live);
      const liveCustomerId = tmv3_clean_(normalized['Customer ID']);
      const liveOrderId = tmv3_clean_(normalized['Order ID']);
      const liveType = tmv3_norm_(normalized['Task Type']);
      if (tmv3_clean_(normalized['Task ID']) !== String(taskId) ||
          (vertical === 'Service' && liveType.indexOf('service') < 0) ||
          (vertical === 'Delivery' && liveType.indexOf('deliver') < 0) ||
          !tmv3_taskIsOpen_(normalized['Status']) ||
          !/^[1-9]\d*$/.test(liveCustomerId)) {
        return {status:'REVIEW_TASK_STATUS_OR_CUSTOMER',taskId:taskId,writes:0};
      }
      const freshTitleOrder=tmv3_taskFirstTitleOrderNumber_(normalized['Name']);
      if (freshTitleOrder.status!=='UNIQUE' || freshTitleOrder.number!==titleOrder.number) {
        return {status:'REVIEW_TASK_TITLE_CHANGED',taskId:taskId,writes:0};
      }
      // Persist the fresh GET locally so report omissions do not cause
      // repeated per-task reads on already-correct relationships.
      tmv3_upsertTaskCacheRow_(normalized);

      // A live relationship that appeared since cache refresh takes priority.
      if (liveOrderId && liveOrderId !== tmv3_clean_(match[0]['Order ID'])) {
        return {status:'REVIEW_EXISTING_ORDER_CONFLICT',taskId:taskId,writes:0};
      }

      // Work Order identity must be independently proven from Striven, never
      // inferred solely from the task name or stale Source Orders cache.
      const id = Number(match[0]['Order ID']);
      let proof;
      try {
        proof = tmv3_taskFirstFreshOrderOwnershipById_(id,liveCustomerId,normalized['Name'],vertical);
      } catch (err) {
        const reason = String(err && err.message || err);
        tmv3_audit_(vertical,'',String(taskId),'TASK_FIRST_AUTO','REVIEW',reason);
        return {status:'REVIEW_ORDER_NOT_VERIFIED_ACTIVE',taskId:taskId,writes:0,reason:reason,requiresReadback:false,writeOutcome:'NOT_ATTEMPTED'};
      }
      const contactId = proof.contactId;
      // Missing/conflicting Contact remains REVIEW in the guarded writer;
      // it must not prevent an independently safe Order relationship repair.

      // Keep an existing location; do not invent or update location fields.
      repairStarted = true;
      const result = tmv3_repairExistingTaskRelationships(
        taskId,vertical,
        {customerId:liveCustomerId,orderId:String(id),
          contactId:contactId,locationId:''},
        'TASK_FIRST_PRODUCTION'
      );
      tmv3_audit_(vertical,'',String(taskId),'TASK_FIRST_AUTO',
        (result.status === 'PATCHED_AND_VERIFIED' || result.status === 'NO_CHANGE_VERIFIED')
          ? 'PASS' : 'REVIEW',JSON.stringify(result));
      return result;
    } catch(err) {
      const message = String(err && err.message || err);
      tmv3_audit_(vertical,'',String(taskId),'TASK_FIRST_AUTO','ATTENTION',message);
      return {
        status:'ATTENTION',taskId:taskId,error:message,
        writes:repairStarted ? null : 0,
        writeOutcome:repairStarted ? 'UNKNOWN_RECONCILE_BEFORE_RETRY' : 'NOT_ATTEMPTED',
        requiresReadback:repairStarted
      };
    }
  } finally {
    lock.releaseLock();
  }
}

/** Google Apps Script: dedicated resumable continuation, never a general write gate. */
function tmv3_serviceDeliveryOrderBackfill() {
  const result=tmv3_taskFirstProductionCycle_();
  if (result && result.status==='LOCK_BUSY') result.continuation='DEFERRED_TO_NEXT_STANDARD_SLOT';
  return result;
}


function tmv3_serviceDeliveryOrderBackfill_(props) {
  const ids = [18374,18471,18733,18735,18754,18774,18801,18806,18831,18896,18897,18908,18913,18914,18979,19037,19064,19098,19109,19126,19137,19190,19195,19198,19201,19203,19204,19205,19214,19218,19221,19222,19223,19224,19225,19226,19227,19228,19229,19230,19231,19232,19233,19234,19235,19240,19241,19242,19248,19249,19250,19251,19252,19253,19255,19256,19257,19258,19259,19260,19261,19262,19263,19264,19265,19266,19267,19268,19269,19270,19272,19273,19274,19275,19276,19279,19280,19281,19282,19283,19284];
  const prefix = 'TMV3_SD_ORDER_BACKFILL_20261020_';
  const handler = 'tmv3_serviceDeliveryOrderBackfill';
  const tasks = tmv3_rows_(TMV3.SHEETS.TASKS);
  // Isolated test projects and other source sets never enroll extra Tasks.
  if (!ids.some(function(id){return tasks.some(function(row){return String(row['Task ID'])===String(id);});})) return null;
  const orders=tmv3_rows_(TMV3.SHEETS.ORDERS);
  function needsWork(id) {
    const value=props.getProperty(prefix+id);
    if (!value) return true;
    const saved=JSON.parse(value);
    if (['INTENT','RETRY'].indexOf(saved.state)>=0) return true;
    // A newly available compatible cached Order may resolve a prior cache hold.
    if (saved.state==='BLOCKED' && ['ORDER_NOT_IN_CACHE','ORDER_SEARCH_NO_EXACT_CUSTOMER_MATCH'].indexOf(saved.reason)>=0 && saved.soNumber) {
      if (saved.reason==='ORDER_NOT_IN_CACHE' && !saved.searchAttempted) return true;
      return orders.some(function(order){return tmv3_clean_(order['Order Number'])===saved.soNumber &&
        tmv3_taskFirstOrderMatchesVertical_(order,saved.vertical);});
    }
    return false;
  }
  const pending=ids.filter(needsWork);
  if (!pending.length) return null;
  const started=Date.now();
  const initialBudget=tmv3_strivenApiBudgetStatus_();
  let processed=0,writes=0,paused='';
  function save(id,value) {
    const result=Object.assign({taskId:id,at:new Date().toISOString()},value);
    props.setProperty(prefix+id,JSON.stringify(result));
    tmv3_audit_(result.vertical || 'SYSTEM','',String(id),'SD_ORDER_BACKFILL',
      result.state==='VERIFIED'?'PASS':'REVIEW',JSON.stringify(result));
    return result;
  }
  function continueLater(delay) {
    // Persisted queue state remains available; no off-slot trigger is created.
    return 'DEFERRED_TO_NEXT_STANDARD_SLOT';
  }
  // Queue the next slice before work so a hard timeout cannot strand the backlog.
  continueLater(180000);
  for (let n=0;n<pending.length && processed<8;n++) {
    const budget=tmv3_strivenApiBudgetStatus_();
    if (budget.exhausted || budget.remaining<12) {paused='API_DAILY_BUDGET';break;}
    if ((budget.enabled !== false && budget.count-initialBudget.count>=32) || Date.now()-started>=120000) {paused='SLICE_BUDGET';break;}
    const id=pending[n];
    processed++;
    let intent=false,vertical='',expectedOrderId='',customerId='',soNumber='',searchAttempted=false;
    try {
      const source=tasks.filter(function(row){return String(row['Task ID'])===String(id);});
      if (source.length!==1) {save(id,{state:'BLOCKED',reason:'SOURCE_TASK_NOT_UNIQUE'});continue;}
      const sourceType=tmv3_norm_(source[0]['Task Type']);
      vertical=sourceType.indexOf('service')>=0?'Service':sourceType.indexOf('deliver')>=0?'Delivery':'';
      if (!vertical) {save(id,{state:'BLOCKED',reason:'WRONG_VERTICAL'});continue;}
      const live=tmv3_normalizeV2TaskModel_(tmv3_rawTaskById_(id));
      const type=tmv3_norm_(live['Task Type']);
      customerId=tmv3_clean_(live['Customer ID']);
      if (String(live['Task ID'])!==String(id) || !tmv3_taskIsOpen_(live.Status) ||
          (vertical==='Service'?type.indexOf('service')<0:type.indexOf('deliver')<0) ||
          !/^[1-9]\d*$/.test(customerId)) {
        save(id,{state:'BLOCKED',vertical:vertical,reason:'FRESH_TASK_IDENTITY_STATUS_OR_CUSTOMER'});continue;
      }
      tmv3_upsertTaskCacheRow_(live);
      const title=tmv3_taskFirstTitleOrderNumber_(live.Name);
      if (title.status!=='UNIQUE') {save(id,{state:'BLOCKED',vertical:vertical,reason:'TASK_TITLE_SO_'+title.status});continue;}
      soNumber=title.number;
      const matches=orders.filter(function(order){return tmv3_clean_(order['Order Number'])===title.number &&
        tmv3_taskFirstOrderMatchesVertical_(order,vertical);});
      const orderIds=[];
      matches.forEach(function(order){const value=tmv3_clean_(order['Order ID']);
        if (/^[1-9]\d*$/.test(value) && orderIds.indexOf(value)<0) orderIds.push(value);});
      // Explicit user-provided identity; fresh API proof still must accept the lifecycle.
      if (!orderIds.length && id===19198 && title.number==='586061') orderIds.push('27795');
      if (!orderIds.length) {searchAttempted=true;orderIds.push(tmv3_taskFirstFindOrderByTitle_(title.number,customerId));}
      if (orderIds.length!==1) {save(id,{state:'BLOCKED',vertical:vertical,soNumber:title.number,
        reason:orderIds.length?'AMBIGUOUS_ORDER':'ORDER_NOT_IN_CACHE'});continue;}
      expectedOrderId=orderIds[0];
      const actual=tmv3_clean_(live['Order ID']);
      if (actual && actual!==expectedOrderId) {save(id,{state:'BLOCKED',vertical:vertical,reason:'EXISTING_ORDER_CONFLICT'});continue;}
      tmv3_taskFirstFreshOrderOwnershipById_(expectedOrderId,customerId,live.Name,vertical);
      const prior=JSON.parse(props.getProperty(prefix+id) || 'null');
      if (prior && prior.state==='INTENT') {
        if (prior.orderId===actual && prior.customerId===customerId) {
          save(id,{state:'VERIFIED',vertical:vertical,orderId:actual,customerId:customerId,
            reason:'INTENT_RECONCILED_BY_FRESH_READ',writes:0});
        } else save(id,{state:'UNCERTAIN',vertical:vertical,orderId:expectedOrderId,
          reason:'PREVIOUS_WRITE_NOT_CONFIRMED_NO_RETRY'});
        continue;
      }
      if (actual===expectedOrderId) {
        save(id,{state:'VERIFIED',vertical:vertical,orderId:actual,customerId:customerId,
          reason:'ALREADY_LINKED_FRESH_VERIFIED',writes:0});continue;
      }
      // Persist intent before the existing, independently guarded, Order-only PATCH.
      props.setProperty(prefix+id,JSON.stringify({state:'INTENT',taskId:id,vertical:vertical,
        orderId:expectedOrderId,customerId:customerId,at:new Date().toISOString()}));
      intent=true;
      const result=tmv3_repairExistingTaskRelationships(id,vertical,
        {customerId:customerId,orderId:expectedOrderId,contactId:'',locationId:''},'TASK_FIRST_PRODUCTION');
      writes+=Number(result.writes || 0);
      const after=tmv3_normalizeV2TaskModel_(tmv3_rawTaskById_(id));
      if (String(after['Task ID'])!==String(id) || tmv3_clean_(after['Order ID'])!==expectedOrderId ||
          tmv3_clean_(after['Customer ID'])!==customerId || !tmv3_taskIsOpen_(after.Status) ||
          tmv3_clean_(after['Task Type'])!==tmv3_clean_(live['Task Type']) ||
          tmv3_clean_(after.Name)!==tmv3_clean_(live.Name)) throw new Error('BACKFILL_READBACK_MISMATCH');
      tmv3_upsertTaskCacheRow_(after);
      save(id,{state:'VERIFIED',vertical:vertical,orderId:expectedOrderId,customerId:customerId,
        soNumber:title.number,writes:Number(result.writes || 0),readbackVerified:true});
    } catch(err) {
      const reason=String(err && err.message || err).slice(0,600);
      if (intent) {
        // Keep INTENT for one automatic fresh read on continuation; never blind retry.
        tmv3_audit_(vertical,'',String(id),'SD_ORDER_BACKFILL','ATTENTION',reason);
      } else {
        const prior=JSON.parse(props.getProperty(prefix+id) || 'null');
        const attempts=Number(prior && prior.attempts || 0)+1;
        const transient=/HTTP (429|5\d\d)|DAILY_GUARD|GET failure|timed out|timeout|temporar|Service invoked too many/i.test(reason);
        save(id,{state:transient && attempts<3?'RETRY':'BLOCKED',vertical:vertical,reason:reason,attempts:attempts,soNumber:soNumber,searchAttempted:searchAttempted});
      }
      if (/429|DAILY_GUARD/.test(reason)) {paused='API_RATE_OR_DAILY_GUARD';break;}
    }
  }
  const tally={VERIFIED:0,BLOCKED:0,UNCERTAIN:0,PENDING:0};
  ids.forEach(function(id){const value=JSON.parse(props.getProperty(prefix+id) || 'null');
    const state=value && value.state;
    if (needsWork(id)) {tally.PENDING++;return;}
    if (state && Object.prototype.hasOwnProperty.call(tally,state)) tally[state]++;
    else tally.PENDING++;});
  ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()===handler;})
    .forEach(function(t){ScriptApp.deleteTrigger(t);});
  if (tally.PENDING) continueLater(paused==='API_DAILY_BUDGET'?3600000:
    Math.max(1000,60000-(Date.now()-started))); // At least 60 seconds between slice starts; no extra idle minute after a slow slice.
  const summary={status:tally.PENDING?'ORDER_BACKFILL_CONTINUING':'ORDER_BACKFILL_COMPLETE',
    total:ids.length,processed:processed,writes:writes,tally:tally,paused:paused,apiBudget:tmv3_strivenApiBudgetStatus_()};
  props.setProperty(prefix+'SUMMARY',JSON.stringify(summary));
  tmv3_audit_('SYSTEM','','','SD_ORDER_BACKFILL_SUMMARY','INFO',JSON.stringify(summary));
  return summary;
}

/** Existing Delivery Task identity does not require a Calendar approval-report row. */

function tmv3_taskFirstOrderMatchesVertical_(order,vertical) {
  if (vertical==='Delivery' && tmv3_clean_(order && order['Order Type']).split('|')
      .some(function(type){return tmv3_clean_(type).toUpperCase()==='SALES_ORDER';})) return true;
  return tmv3_step3OrderMatchesVertical_(order,vertical);
}

/** Official Striven OpenAPI: SalesOrderSearchParameters Number/CustomerId, paged Data/TotalCount. */
function tmv3_taskFirstFindOrderByTitle_(number,customerId) {
  if (!/^\d{4,9}$/.test(String(number)) || !/^[1-9]\d*$/.test(String(customerId))) throw new Error('ORDER_SEARCH_INVALID_IDENTITY');
  const result=tmv3_fetchJson_(TMV3.API_BASE+'/v1/sales-orders/search',{
    method:'post',contentType:'application/json',
    payload:JSON.stringify({Number:String(number),CustomerId:Number(customerId),PageIndex:0,PageSize:100})
  }) || {};
  const data=result.Data!==undefined?result.Data:result.data;
  const total=Number(result.TotalCount!==undefined?result.TotalCount:result.totalCount);
  if (!Array.isArray(data) || !Number.isSafeInteger(total) || total<0 || total!==data.length || total>100) {
    throw new Error('ORDER_SEARCH_INCOMPLETE_RESULT_NO_PATCH');
  }
  const ids=[];
  data.forEach(function(row){
    const id=tmv3_clean_(row.Id!==undefined?row.Id:row.id);
    const candidate=tmv3_clean_(row.Number!==undefined?row.Number:row.number);
    const customer=row.Customer || row.customer || {};
    const owner=tmv3_clean_(customer.Id!==undefined?customer.Id:customer.id);
    if (candidate===String(number) && owner===String(customerId) && /^[1-9]\d*$/.test(id) && ids.indexOf(id)<0) ids.push(id);
  });
  if (ids.length!==1) throw new Error(ids.length?'ORDER_SEARCH_AMBIGUOUS_NO_PATCH':'ORDER_SEARCH_NO_EXACT_CUSTOMER_MATCH');
  return ids[0]; // The caller must still freshly GET this Order to prove ownership and active status.
}

function tmv3_scheduledOperations() { return {status:'RETIRED_USE_STANDARD_DAILY_SLOTS',businessWrites:0}; }

function tmv3_refreshLinksSlot_(label) {
  if (tmv3_executionStage_() <= 7) {
    const mapped = tmv3_calendarStageRefresh_(
      'CALENDAR_STAGE_SLOT_' + String(label || '')
    );
    const links = { status: 'CALENDAR_STAGE_GATED', writes: 0 };
    return { slot: label, mapped: mapped, links: links };
  }

  const mapped = tmv3_shadowMapFromCache();
  const links = tmv3_runCalendarLinksForReady_AUTO();
  tmv3_audit_(
    'SYSTEM','','','REFRESH_LINK_SLOT','PASS',
    String(label || '') + ' · ' + JSON.stringify(links)
  );
  return { slot: label, mapped: mapped, links: links };
}

function tmv3_refreshLinksSlot_0800() { return tmv3_refreshLinksSlot_('8:00 AM'); }
function tmv3_refreshLinksSlot_1000() { return tmv3_refreshLinksSlot_('10:00 AM'); }
function tmv3_refreshLinksSlot_1200() { return tmv3_refreshLinksSlot_('12:00 PM'); }
function tmv3_refreshLinksSlot_1400() { return tmv3_refreshLinksSlot_('2:00 PM'); }
function tmv3_refreshLinksSlot_1600() { return tmv3_refreshLinksSlot_('4:00 PM'); }
function tmv3_refreshLinksSlot_1800() { return tmv3_refreshLinksSlot_('6:00 PM'); }



function tmv3_calendarEventUpdated() { return {status:'RETIRED_USE_STANDARD_DAILY_SLOTS',businessWrites:0}; }






/** Google Apps Script: bounded PreInspection lane. Other verticals stay gated. */


// Google Apps Script: discovery is metadata-only. Identity evidence never proves visit type.

// Google Apps Script: classify likely visits using current Order ownership and job site.
var TMV3_PI_LIKELY_REFS = null;
var TMV3_PI_LIKELY_DECISIONS = {};


// Recheck the current Order after asynchronous Calendar intake, immediately before CREATE.
// An inferred visit keeps its verified Order evidence even when no SO was authored on Calendar.





// Same executor as the scheduled slots; no second repair implementation.






// Toronto wall-clock dates keep business-day deadlines stable across DST.





// Stable key lets successive time-driven executions cover the complete eligible
// calendar snapshot without holding a large list in Script Properties.


// A one-event request must not scan or classify the entire calendar.













// Runs one bounded batch after the preceding batch finishes. Trigger
// installation for the 08/13/17 slots remains unchanged.


















// Google Apps Script: restore only the recorded legacy-title prefix. No Task writes.




// Google Apps Script: verification-only actions must not keep a converged event pending.




// Google Apps Script: classify fresh same-day completed coverage without changing history.


// Metadata-only reconciliation. Every removal requires a new canonical Task read.

// Google Apps Script: targeted evidence refresh; no Striven or Calendar mutation.
function tmv3_verticalEnrichEvidence_(vertical, record, plans) {
  if (['Install','Delivery','Service'].indexOf(vertical)<0) throw new Error('VERTICAL_REQUIRED');
  const one=(plans||[])[0] || {};
  const customerId=tmv3_clean_(one.expectedCustomerId);
  const number=tmv3_clean_(record.orderNumber);
  if (!/^[1-9]\d*$/.test(customerId) || !/^\d+$/.test(number)) return {status:'IDENTITY_OR_ORDER_REVIEW'};
  let id=tmv3_clean_(one.expectedOrderId);
  if (!id) id=tmv3_taskFirstFindOrderByTitle_(number,customerId);
  const proof=tmv3_taskFirstFreshOrderOwnershipById_(id,customerId,'SO#'+number,vertical);
  const sheet=tmv3_sheet_(TMV3.SHEETS.ORDERS),values=sheet.getDataRange().getValues();
  const headers=['Order ID','Order Number','Customer ID','Location ID','Contact ID','Status','Name','Order Type','URL','Fingerprint'];
  if (!values.length || JSON.stringify(values[0].slice(0,10))!==JSON.stringify(headers)) throw new Error('ORDER_CACHE_SCHEMA_DRIFT');
  const matches=[];
  values.slice(1).forEach(function(row,index){if(String(row[0])===String(id))matches.push({row:row,index:index+2});});
  if (matches.length>1 || matches.some(function(x){return String(x.row[1])!==number ||
    (tmv3_clean_(x.row[2]) && String(x.row[2])!==customerId);})) throw new Error('ORDER_CACHE_IDENTITY_CONFLICT');
  const current=matches.length?matches[0].row.slice(0,10):[id,number,customerId,'','','','','SALES_ORDER',TMV3.ORDER_URL_BASE+id,''];
  // In Progress alone does not establish a Service Work Order classification.
  if (vertical==='Service' && String(current[7]).split('|').indexOf('WORK_ORDER')<0) return {status:'WORK_ORDER_SOURCE_REQUIRED'};
  if(vertical==='Install') {
    proof.contactId=tmv3_installResolveContactFresh_(record,customerId,proof.contactId || current[4],one);
  }
  if (proof.contactId) {
    const contact=tmv3_getContactById_(proof.contactId,customerId);
    if (!contact || contact.__ownershipVerified!==true) throw new Error('ORDER_CONTACT_OWNERSHIP_UNPROVEN');
    current[4]=proof.contactId;
  }
  const owned=tmv3_preInspectionCustomerLocationsFresh_(Number(customerId));
  const locations=owned.filter(function(location){return tmv3_addressStrongMatch_(
    tmv3_addressParts_(record.location),tmv3_addressParts_(tmv3_locationFullAddress_(location)));});
  if (locations.length===1) tmv3_upsertLocationCacheObject_(locations[0]);
  if (proof.locationId) current[3]=proof.locationId;
  current[2]=customerId;current[5]=proof.status;
  current[9]=tmv3_hash_(current.slice(0,9).map(String).join('|'));
  sheet.getRange(matches.length?matches[0].index:values.length+1,1,1,10).setValues([current]);
  const after=sheet.getRange(matches.length?matches[0].index:values.length+1,1,1,10).getValues()[0];
  if (String(after[0])!==String(id) || String(after[2])!==customerId || String(after[4])!==String(current[4])) throw new Error('TARGETED_ORDER_CACHE_READBACK_FAILED');
  return {status:'EVIDENCE_REFRESHED',orderId:id,orderStatus:proof.status,contactVerified:!!proof.contactId,locationMatches:locations.length};
}

function tmv3_verticalProductionCycle_(vertical, options) {
  options=options || {};
  const weekScope=tmv3_verticalRepairScope_(vertical,options);
  const targetIds=options.eventIds||null;
  if(targetIds && (!Array.isArray(targetIds)||!targetIds.length||targetIds.some(function(id){return typeof id!=='string'||!id;})))throw Error('TARGET_IDS_REQUIRED');
  const policy=TMV3.OPERATIONS;
  if (TMV3.MODE!=='PRODUCTION_WRITE' || policy.verticalProductionEnabled!==true ||
      (policy.verticalProductionVerticals||[]).indexOf(vertical)<0) return {status:'VERTICAL_DISABLED',vertical:vertical,writes:0};
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(1000)) return {status:'LOCK_BUSY',vertical:vertical,writes:0};
  const result={status:'ACTIVE_NO_READY_RECORD',vertical:vertical,writes:0,scanned:0,results:[]};
  try {
    const props=PropertiesService.getScriptProperties(),now=Date.now();
    const stateKey='TMV3_VERTICAL_SWEEP_'+vertical+(weekScope ? '_RANGE_'+weekScope.fromDay+'_'+weekScope.toDay : '')+(targetIds?'_TARGET_'+tmv3_hash_(targetIds.slice().sort().join('|')):'');
    let state=JSON.parse(props.getProperty(stateKey)||'null');
    if(!state || state.complete) {
      if(vertical==='Install') {
        const budget=tmv3_strivenApiBudgetStatus_();
        if(budget.exhausted || budget.remaining<Number(policy.verticalProductionMinApiRemaining||60))return Object.assign(result,{status:'API_BUDGET_WAIT'});
        result.sources=tmv3_refreshOperationalSourcesIfNeeded_(120,false);
        result.intake=tmv3_installRefreshView_('INSTALL_SWEEP_START');
      }
      // Consume the existing intake queue; each event is freshly verified below.
      // Rebuilding every operator sheet here can consume the entire execution.
      const today=tmv3_step6LocalDay_(new Date()),seen={};
      const events=tmv3_rows_(TMV3.VERTICALS[vertical].sheet).filter(function(row){
        const id=tmv3_clean_(row['Event ID']),day=tmv3_step6LocalDay_(row['Calendar Start']);
        if(!id||seen[id]||targetIds && targetIds.indexOf(id)<0||!day||day<(weekScope ? weekScope.fromDay : today)||weekScope && day>=weekScope.toDay)return false;seen[id]=true;return true;
      }).sort(function(a,b){return new Date(a['Calendar Start'])-new Date(b['Calendar Start']);});
      const ids=events.map(function(row){return tmv3_clean_(row['Event ID']);});
      const chunks=Math.ceil(ids.length/50);
      for(let n=0;n<chunks;n++)props.setProperty(stateKey+'_IDS_'+n,JSON.stringify(ids.slice(n*50,n*50+50)));
      state={id:String(now),cursor:0,total:ids.length,chunks:chunks,complete:false,scanned:0,verified:0,review:0,uncertain:0};
      props.setProperty(stateKey,JSON.stringify(state));
    }
    result.total=state.total;result.sweepId=state.id;
    const events=[];for(let n=0;n<state.chunks;n++)events.push.apply(events,JSON.parse(props.getProperty(stateKey+'_IDS_'+n)||'[]'));
    if(events.length!==state.total)throw Error('SWEEP_SNAPSHOT_INCOMPLETE_NO_WRITE');
    for(let i=state.cursor;i<events.length;i++) {
      const budget=tmv3_strivenApiBudgetStatus_();
      if(Date.now()-now>120000){result.status='RUNTIME_CONTINUATION_REQUIRED';break;}
      if(budget.exhausted||budget.remaining<Number(policy.verticalProductionMinApiRemaining||60)){result.status='API_BUDGET_WAIT';break;}
      const id=events[i];result.scanned++;
      const key='TMV3_VERTICAL_EVENT_'+tmv3_hash_(vertical+'|'+id);
      let writeStarted=false;
      try {
        const record=tmv3_findFreshEventRecord_(vertical,id);
        if(weekScope){const freshDay=tmv3_step6LocalDay_(record.start);if(!freshDay||freshDay<weekScope.fromDay||freshDay>=weekScope.toDay){result.results.push({eventId:id,status:'SKIPPED_MOVED_OUTSIDE_REPAIR_RANGE'});continue;}}
        if(vertical==='Install' && !tmv3_installCreatorAllowed_(record)){result.results.push({eventId:id,status:'SKIPPED_CREATOR_UNVERIFIED'});continue;}
        let plans=tmv3_step7FreshPlansForEvent_(vertical,id);
        const alreadyFulfilled=plans.length>0 && plans.every(function(p){return p.disposition==='FULFILLED_NO_RECREATE' && p.plan==='VERIFY_ONLY_HISTORY_NO_WRITE' && p.readStatus==='FRESH_TASK_GET' && p.actions.length===0 && p.customerCheck==='MATCH' && tmv3_step6IsFulfilledStatus_(p.taskStatus) && tmv3_step6TaskTouchesEventDay_({start:p.calendarStart},{Start:p.taskStart,Due:p.taskDue});});
        if(alreadyFulfilled){result.results.push({eventId:id,status:'FULFILLED_NO_RECREATE'});continue;}
        const held=JSON.parse(props.getProperty(key)||'null');
        if(weekScope && !options.forceLinkFormat && !(held && ['INTENT','UNCERTAIN'].indexOf(held.state)>=0) && tmv3_weeklyAlreadyAligned_(record,plans)) {
          result.results.push({eventId:id,status:'ALREADY_ALIGNED_READ_ONLY',taskIds:plans.map(function(p){return p.taskId;}),freshTaskReads:true,clickableLinksVerified:true,businessWrites:0});continue;
        }
        const enriched=tmv3_verticalEnrichEvidence_(vertical,record,plans);
        if(enriched.status==='EVIDENCE_REFRESHED') plans=tmv3_step7FreshPlansForEvent_(vertical,id);
        const historical=plans.length>0 && plans.every(function(p){return p.disposition==='FULFILLED_NO_RECREATE' && p.plan==='VERIFY_ONLY_HISTORY_NO_WRITE' && p.readStatus==='FRESH_TASK_GET' && p.actions.length===0 && p.customerCheck==='MATCH' && tmv3_step6IsFulfilledStatus_(p.taskStatus);});
        if(historical){result.results.push({eventId:id,status:'FULFILLED_NO_RECREATE'});continue;}
        if(vertical==='Install' && plans.some(function(p){return !p.expectedRequestedById || p.expectedRequestedByType!=='contact';})) {
          result.results.push({eventId:id,status:'REVIEW',error:'INSTALL_REQUESTED_BY_UNRESOLVED'});continue;
        }
        if (!plans.length || plans.some(function(p){return !!p.blocker;}) ||
            (plans.length>1 && (vertical!=='Service' || !plans.every(function(p){return p.disposition==='MATCH_EXISTING_MULTI' && p.taskId;})))) {
          result.results.push({eventId:id,status:'REVIEW',plans:plans.map(function(p){return {taskId:p.taskId,plan:p.plan,blocker:p.blocker};})});continue;
        }
        if(plans.some(function(p){return p.actualRequestedById && p.expectedRequestedById && (String(p.actualRequestedById)!==String(p.expectedRequestedById) || p.actualRequestedByType!==p.expectedRequestedByType);})) {
          result.results.push({eventId:id,status:'REVIEW_EXISTING_REQUESTED_BY_CONFLICT'});continue;
        }
        let prior=JSON.parse(props.getProperty(key)||'null');
        // This exact local validation error occurs before any business mutation
        // in tmv3_step7FreshCriticalOwnership_. It is a proven rejection, not
        // an unknown network/write outcome. Other uncertain writes stay held.
        if(prior && prior.state==='UNCERTAIN' &&
          /^Fresh Customer existence verification expected one Customer \d+; found \d+\.$/.test(prior.error||'') &&
          Number((prior.error.match(/found (\d+)/)||[])[1])>1) {
          props.deleteProperty(key);prior=null;
          tmv3_audit_(vertical,id,'','KNOWN_PREWRITE_REJECTION_RECOVERED','PASS','Duplicate report rows for one Customer ID; canonical ownership checks remain required.');
        }
        const noTaskActions=plans.every(function(p){return p.plan==='NO_CHANGE' && (p.actions||[]).every(function(a){return a==='VERIFY_CALENDAR_LINKS';});});
        if(prior && ['INTENT','UNCERTAIN'].indexOf(prior.state)>=0 && !noTaskActions) {
          result.results.push({eventId:id,status:'UNCERTAIN_RECONCILIATION_REQUIRED'});continue;
        }
        if(options.forceLinkFormat && plans.some(function(p){return !p.taskId||(p.actions||[]).some(function(a){return a!=='VERIFY_CALENDAR_LINKS';});})){result.results.push({eventId:id,status:'REVIEW',error:'LINK_FORMAT_ONLY_OTHER_CHANGES_REQUIRE_SEPARATE_REPAIR'});continue;}
        const taskIds=[];
        if(plans.some(function(p){const duplicate=p.taskId && taskIds.indexOf(String(p.taskId))>=0;if(p.taskId)taskIds.push(String(p.taskId));return duplicate;})) throw new Error('DUPLICATE_TASK_CONTRACTS');
        const executions=[];
        for(let j=0;j<plans.length;j++) {
          const plan=j ? tmv3_step7FreshExecutionContract_(vertical,id,plans[j].taskId) : plans[j];
          if(plan.blocker)throw new Error('FRESH_SIBLING_CONTRACT_BLOCKED');
          if(vertical==='Service' && (plan.actions||[]).some(function(a){return /CREATE|RECREATE/.test(a);})) throw new Error('SERVICE_TASK_CREATION_OWNED_BY_STRIVEN');
          // Canonical executor rechecks identity/lifecycle/duplicate evidence and reads every write back.
          props.setProperty(key,JSON.stringify({state:'INTENT',at:Date.now(),fingerprint:record.fingerprint,taskId:plan.taskId||'',plan:plan.plan}));
          tmv3_verticalProductionCycle_.context={vertical:vertical,eventId:id,contractJson:JSON.stringify(plan)};
          writeStarted=true;
          try {
            const execution=tmv3_executeFreshStep7Selection_({previousPlan:plan.plan,contract:plan},'VERTICALAUTO','ALL');
            if(['VERIFIED_CONVERGENCE','CREATED_VERIFIED_AND_CONVERGED'].indexOf(execution.status)<0) throw new Error('VERTICAL_CONVERGENCE_NOT_VERIFIED:'+execution.status);
            executions.push({taskId:execution.taskId||plan.taskId,status:execution.status,actions:plan.actions,readbackVerified:true});
          }finally{delete tmv3_verticalProductionCycle_.context;}
        }
        props.setProperty(key,JSON.stringify({state:'VERIFIED',at:Date.now(),fingerprint:record.fingerprint,taskIds:executions.map(function(x){return x.taskId;})}));
        result.status='VERIFIED_CONVERGENCE';result.transactionsVerified=Number(result.transactionsVerified||0)+1;result.writes+=noTaskActions?0:1;result.results.push({eventId:id,status:'VERIFIED_CONVERGENCE',executions:executions});
        tmv3_audit_(vertical,id,'','VERTICAL_AUTO','PASS',JSON.stringify(result.results[result.results.length-1]));
      } catch(err) {
        const message=String(err && err.message || err);
        if(writeStarted)props.setProperty(key,JSON.stringify({state:'UNCERTAIN',at:Date.now(),error:message.slice(0,500)}));
        result.results.push({eventId:id,status:writeStarted?'UNCERTAIN_RECONCILIATION_REQUIRED':'REVIEW',error:message});
        tmv3_audit_(vertical,id,'','VERTICAL_AUTO',writeStarted?'ATTENTION':'REVIEW',message);
        if(writeStarted)tmv3_verticalAutomationFailureAlert_(vertical,id,message);
      } finally {
        const outcome=result.results[result.results.length-1]||{status:'UNKNOWN'};
        state.cursor=i+1;state.scanned++;
        if(outcome.status==='VERIFIED_CONVERGENCE'||outcome.status==='FULFILLED_NO_RECREATE'||outcome.status==='ALREADY_ALIGNED_READ_ONLY')state.verified++;
        else if(outcome.status==='UNCERTAIN_RECONCILIATION_REQUIRED')state.uncertain++;
        else state.review++;
        props.setProperty(stateKey,JSON.stringify(state));
        tmv3_audit_(vertical,id,'','VERTICAL_SWEEP_EVENT',outcome.status,JSON.stringify(outcome));
      }
    }
    state.complete=state.cursor>=state.total;props.setProperty(stateKey,JSON.stringify(state));
    result.sweep=state;result.pending=state.total-state.cursor;
    if(state.complete)result.status='SWEEP_COMPLETE';
    return result;
  } finally {
    delete tmv3_verticalProductionCycle_.context;
    PropertiesService.getScriptProperties().setProperty('TMV3_VERTICAL_RESULT_'+vertical,JSON.stringify(Object.assign({},result,{results:result.results.slice(-8)})));
    lock.releaseLock();
  }
}









function tmv3_verticalAutomationFailureAlert_(vertical,eventId,error) {
  const props=PropertiesService.getScriptProperties();
  const key='TMV3_VERTICAL_ALERT_'+tmv3_hash_(tmv3_step6LocalDay_(new Date())+'|'+vertical+'|'+eventId);
  if(props.getProperty(key))return {status:'ALREADY_ALERTED_OR_UNCERTAIN'};
  if(MailApp.getRemainingDailyQuota()<1)return {status:'QUOTA_WAIT'};
  props.setProperty(key,'SENDING');
  try{
    MailApp.sendEmail({to:TMV3.OPERATIONS.verticalFailureAlertEmail,
      subject:vertical+' automation: action required',
      body:'This is an automated email notification from Classic Fireplace & BBQ Store.\n\nHi Pramodh,\n\n'+vertical+' could not verify an automatic transaction.\nEvent: '+eventId+'\n'+error+'\n\nThe record remains held. An uncertain write must be reconciled before retrying.',
      name:'Classic Fireplace & BBQ Store'});
    props.setProperty(key,'SENT');return {status:'SENT'};
  }catch(err){tmv3_audit_(vertical,eventId,'','VERTICAL_FAILURE_ALERT','ATTENTION',String(err));return {status:'UNCERTAIN_SEND_NO_RETRY'};}
}

function tmv3_fourVerticalAutomationStatus() {
  const props=PropertiesService.getScriptProperties(),policy=TMV3.OPERATIONS;
  const verticals={PreInspection:{enabled:tmv3_operationWritesEnabled_('PREINSPECTIONAUTO'),hours:[8,13,17]}};
  ['Service','Delivery','Install'].forEach(function(v){
    let recent=null;try{recent=JSON.parse(props.getProperty('TMV3_VERTICAL_RESULT_'+v)||'null');}catch(ignored){}
    verticals[v]={enabled:TMV3.MODE==='PRODUCTION_WRITE' && policy.verticalProductionEnabled===true &&
      (policy.verticalProductionVerticals||[]).indexOf(v)>=0,requiredSalesOrderStatus:v==='Service'?'In Progress':'Approved',recent:recent};
  });
  return {version:TMV3.VERSION,mode:TMV3.MODE,verticals:verticals,triggers:tmv3_listTriggers(),
    generalAutomationWritesEnabled:policy.automationWritesEnabled,manualWritesEnabled:policy.manualWritesEnabled,
    installReminderEnabled:TMV3.MODE==='PRODUCTION_WRITE'&&policy.installRemindersEnabled===true,
    apiBudget:tmv3_strivenApiBudgetStatus_()};
}


// Google Apps Script: cache-backed operator publication, with no Striven/Calendar writes.
function tmv3_refreshAutomationSheetsAfter_(startedAt) {
  try {
    PropertiesService.getScriptProperties().setProperty('TMV3_SHEET_REFRESH_PENDING','true');
    if(Date.now()-Number(startedAt)>240000) return {status:'DEFERRED_TO_NEXT_STANDARD_SLOT'};
    return tmv3_refreshAutomationSheets();
  } catch(err) {
    // A view-publication failure must not replay or relabel a completed business transaction.
    return {status:'SHEET_REFRESH_ATTENTION',error:String(err && err.message || err)};
  }
}



function tmv3_refreshAutomationSheets() {
  if(tmv3_verticalBusinessQueuePending_())return {status:'DEFERRED_ACTIVE_BUSINESS_SWEEP',strivenWrites:0,calendarWrites:0};
  const lock=LockService.getScriptLock(),props=PropertiesService.getScriptProperties();
  props.setProperty('TMV3_SHEET_REFRESH_PENDING','true');
  if(!lock.tryLock(1000))return {status:'LOCK_BUSY'};
  const startedAt=new Date().toISOString();
  try {
    const mapped=tmv3_step6TaskDecisionRunCached('AUTOMATIC_SHEET_REFRESH');
    if(!mapped.verification || mapped.verification.pass!==true)throw Error('SHEET_PROJECTION_VERIFICATION_FAILED');
    const counts={};
    Object.keys(TMV3.VERTICALS).forEach(function(v){
      const rows=tmv3_rows_(TMV3.VERTICALS[v].sheet);
      counts[v]=rows.filter(function(r){return !!tmv3_clean_(r['Event ID']);}).length;
      if(counts[v]!==Number(mapped.write[v]))throw Error('SHEET_ROW_READBACK_MISMATCH:'+v);
    });
    const morning=tmv3_refreshMorningOps();
    const stamp=Utilities.formatDate(new Date(),TMV3_TIMEZONE,'yyyy-MM-dd h:mm a');
    const overview=tmv3_sheet_(TMV3.SHEETS.OVERVIEW);
    const title='V3 '+TMV3.VERSION+' — Sheet refreshed '+stamp+' Toronto';
    overview.getRange('B2').setValue(title);
    overview.getRange('B4').setValue(TMV3.MODE+' — dedicated four-vertical automation; general AUTO/MANUAL remain gated.');
    overview.getRange('B5').setValue('Current Calendar intake + latest local Striven cache; fresh write verification stays in TM Audit.');
    overview.getRange('B6').setValue('Mapping Status shows the current cached decision; refresh does not certify business writes.');
    overview.getRange('B16').setValue('Automatic Sheet refresh: '+stamp+' Toronto. Operator tabs and Morning Ops refreshed.');
    SpreadsheetApp.flush();
    if(overview.getRange('B2').getValue()!==title)throw Error('SHEET_TIMESTAMP_READBACK_MISMATCH');
    const result={status:'AUTOMATION_SHEETS_REFRESHED',version:TMV3.VERSION,startedAt:startedAt,
      completedAt:new Date().toISOString(),timezone:TMV3_TIMEZONE,rows:counts,
      projectionVerified:true,timestampReadbackVerified:true,sourceMode:'CURRENT_CALENDAR_AND_LOCAL_STRIVEN_CACHE',
      strivenWrites:0,calendarWrites:0,morningOps:morning};
    props.setProperty('TMV3_SHEET_REFRESH_LAST_SUCCESS',JSON.stringify(result));
    props.deleteProperty('TMV3_SHEET_REFRESH_PENDING');props.deleteProperty('TMV3_SHEET_REFRESH_LAST_ERROR');
    tmv3_audit_('SYSTEM','','','AUTOMATIC_SHEET_REFRESH','PASS',JSON.stringify(result));
    Logger.log(JSON.stringify(result));return result;
  } catch(err) {
    props.setProperty('TMV3_SHEET_REFRESH_LAST_ERROR',JSON.stringify({at:new Date().toISOString(),error:String(err && err.message || err)}));
    tmv3_audit_('SYSTEM','','','AUTOMATIC_SHEET_REFRESH','ATTENTION',String(err && err.message || err));
    throw err;
  } finally {lock.releaseLock();}
}


function tmv3_automationSheetRefreshStatus() {
  const p=PropertiesService.getScriptProperties();
  function json(key){const value=p.getProperty(key);return value?JSON.parse(value):null;}
  return {version:TMV3.VERSION,lastSuccess:json('TMV3_SHEET_REFRESH_LAST_SUCCESS'),
    lastError:json('TMV3_SHEET_REFRESH_LAST_ERROR'),pending:p.getProperty('TMV3_SHEET_REFRESH_PENDING')==='true',
    triggers:tmv3_listTriggers()};
}


// Google Apps Script: independent daily execution windows for each vertical.














// Google Apps Script: resume the same daily sweep; no additional recurring window.

function tmv3_runFullVerticalSweep_(vertical) {
  const began=Date.now();let result=tmv3_verticalProductionCycle_(vertical);
  let totalWrites=Number(result.writes||0),totalScanned=Number(result.scanned||0),allResults=(result.results||[]).slice(),batchStatuses=[result.status];
  while (result.pending>0 && result.status!=='API_BUDGET_WAIT' && result.status!=='LOCK_BUSY' &&
      result.scanned>0 && Date.now()-began<240000) {
    const before=result.pending,next=tmv3_verticalProductionCycle_(vertical);
    batchStatuses.push(next.status);
    totalWrites+=Number(next.writes||0);totalScanned+=Number(next.scanned||0);allResults.push.apply(allResults,next.results||[]);
    result=next;if(!next.scanned || next.pending>=before)break;
  }
  result.batchStatuses=batchStatuses;
  result.writes=totalWrites;result.scanned=totalScanned;result.results=allResults;
  if(result.status==='LOCK_BUSY'||result.pending>0) result.continuation='DEFERRED_TO_NEXT_STANDARD_SLOT';
  PropertiesService.getScriptProperties().setProperty('TMV3_SHEET_REFRESH_PENDING','true');
  result.sheetRefresh={status:'DEFERRED_TO_NEXT_STANDARD_SLOT'};
  if(vertical==='Install' && result.sweep && result.sweep.complete)result.sheetRefresh=tmv3_installSchedulePublication_();
  Logger.log(JSON.stringify(result));return result;
}








function tmv3_verticalPatchDiagnostics(vertical,eventIds) {
  if(['Install','Delivery','Service'].indexOf(vertical)<0)throw Error('VERTICAL_REQUIRED');
  return (eventIds||[]).map(function(id){try{
    const plans=tmv3_step7FreshPlansForEvent_(vertical,id);
    return {eventId:id,plans:plans.map(function(p){return {taskId:p.taskId,plan:p.plan,blocker:p.blocker,actions:p.actions,requestedBy:p.expectedRequestedById,actualRequestedBy:p.actualRequestedById,order:p.expectedOrderId,location:p.expectedLocationId};})};
  }catch(e){return {eventId:id,error:String(e)};}});
}

function tmv3_fullVerticalSweepStatus() {
  const p=PropertiesService.getScriptProperties(),out={version:TMV3.VERSION,budget:tmv3_strivenApiBudgetStatus_(),verticals:{},triggers:tmv3_listTriggers()};
  ['Install','Delivery','Service'].forEach(function(v){
    const state=JSON.parse(p.getProperty('TMV3_VERTICAL_SWEEP_'+v)||'null');
    const recent=JSON.parse(p.getProperty('TMV3_VERTICAL_RESULT_'+v)||'null');
    out.verticals[v]={sweep:state,recent:recent};
  });return out;
}
function tmv3_resumeFullVerticalQueues() {
  return {status:'QUEUES_RETAINED_FOR_NEXT_STANDARD_SLOT',snapshot:tmv3_fullVerticalSweepStatus()};
}

function tmv3_verticalLinkedTaskDiagnostics(vertical,eventIds) {
  return (eventIds||[]).map(function(id){try{
    const record=tmv3_findFreshEventRecord_(vertical,id);
    const records=tmv3_step4IdentityRecords_(tmv3_step3BusinessAnchorRecords_(tmv3_step2CalendarRecords_([record]),tmv3_step3AnchorIndex_()),tmv3_step4IdentityIndex_());
    const identity=records[0].step4||{},linkedId=record.existingTaskId;
    const fresh=linkedId?tmv3_step7FreshTask_(linkedId,{taskById:{}}):null;
    return {eventId:id,linkedId:linkedId,identity:identity.disposition,expectedCustomer:identity.customer&&identity.customer['Customer ID'],expectedLocation:identity.location&&identity.location['Location ID'],locationStatus:identity.locationStatus,fresh:fresh&&{ok:fresh.ok,error:fresh.error,task:fresh.task}};
  }catch(e){return {eventId:id,error:String(e)};}});
}

// Publication yields while runnable business work remains; budget-paused queues
// can still publish their latest verified results.
function tmv3_verticalBusinessQueuePending_() {
  const props=PropertiesService.getScriptProperties();
  if(tmv3_strivenApiBudgetStatus_().remaining<Number(TMV3.OPERATIONS.verticalProductionMinApiRemaining||60))return false;
  return ['Install','Delivery','Service'].some(function(v){
    const state=JSON.parse(props.getProperty('TMV3_VERTICAL_SWEEP_'+v)||'null');
    return state && !state.complete && state.cursor<state.total;
  });
}

// Google Apps Script: Install-only contact resolution; no external writes.

// Caller owns the script lock. Rebuild only Install, never another vertical.




// Read-only readiness inspection. It never installs triggers or executes tasks.




// Google Apps Script: date-scoped owner repair through the unchanged canonical worker.
function tmv3_verticalRepairScope_(vertical,options) {
  if(!options || !options.fromDay && !options.toDay)return null;
  if(['Install','Delivery','Service'].indexOf(vertical)<0)throw Error('REPAIR_VERTICAL_INVALID');
  const valid=function(s){return typeof s==='string' && /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(new Date(s+'T12:00:00Z').getTime()) && new Date(s+'T12:00:00Z').toISOString().slice(0,10)===s;};
  if(!valid(options.fromDay)||!valid(options.toDay)||options.fromDay>=options.toDay||new Date(options.toDay)-new Date(options.fromDay)>7*86400000)throw Error('REPAIR_RANGE_INVALID_MAX_SEVEN_DAYS');
  return {fromDay:options.fromDay,toDay:options.toDay};
}
function tmv3_runWeeklyVerticalRepair(vertical,fromDay,toDay) {
  const scope=tmv3_verticalRepairScope_(vertical,{fromDay:fromDay,toDay:toDay});
  if(!scope)throw Error('REPAIR_RANGE_REQUIRED');
  const result=tmv3_verticalProductionCycle_(vertical,scope);
  PropertiesService.getScriptProperties().setProperty('TMV3_SHEET_REFRESH_PENDING','true');
  return result;
}

// Google Apps Script: zero business writes; preserve correct clickable Calendar blocks.
function tmv3_weeklyAlreadyAligned_(record,plans) {
  if(!record||['Install','Delivery','Service'].indexOf(record.vertical)<0||!plans||!plans.length)return false;
  if(!plans.every(function(p){return p.taskId && p.readStatus==='FRESH_TASK_GET' && p.plan==='NO_CHANGE' && !p.blocker && p.inputFingerprint===record.fingerprint && p.customerCheck==='MATCH' && p.orderCheck==='MATCH' && p.locationCheck==='MATCH' && (p.actions||[]).every(function(a){return a==='VERIFY_CALENDAR_LINKS';});}))return false;
  const cfg=TMV3.VERTICALS[record.vertical],taskIds=plans.map(function(p){return String(p.taskId);}).sort();
  if(taskIds.some(function(id,i){return i>0&&id===taskIds[i-1];}))return false;
  const required=taskIds.map(function(id){return TMV3.TASK_URL_BASE+id;});
  if(cfg.calendarOrderLinkRequired){const orderIds=plans.map(function(p){return String(p.expectedOrderId||'');});if(!orderIds[0]||!orderIds.every(function(id){return id===orderIds[0];}))return false;required.push(TMV3.ORDER_URL_BASE+orderIds[0]);}
  const ids=tmv3_requiredCalendarCopiesForEvent_(record);if(!ids.length)return false;
  return ids.every(function(calendarId){
    const found=tmv3_findEventCopyRobust_(calendarId,record.eventId);if(!found||!found.event)return false;
    const description=String(found.event.getDescription()||''),hrefs=[];let m;const rx=/<a\b[^>]*\bhref\s*=\s*["']([^"']+)["'][^>]*>/gi;
    while((m=rx.exec(description))!==null)hrefs.push(m[1].replace(/&amp;/g,'&'));
    if(!required.every(function(url){return hrefs.filter(function(h){return h===url;}).length===1;}))return false;
    const actualTasks=hrefs.filter(function(h){return h.indexOf(TMV3.TASK_URL_BASE)===0;}).map(function(h){return h.slice(TMV3.TASK_URL_BASE.length);}).sort();
    return JSON.stringify(actualTasks)===JSON.stringify(taskIds);
  });
}
function tmv3_weeklyVerticalRepairStatus(vertical,fromDay,toDay) {
  const s=tmv3_verticalRepairScope_(vertical,{fromDay:fromDay,toDay:toDay});
  return JSON.parse(PropertiesService.getScriptProperties().getProperty('TMV3_VERTICAL_SWEEP_'+vertical+'_RANGE_'+s.fromDay+'_'+s.toDay)||'null');
}

// Google Apps Script: restore presentation only on records touched by this repair session.
function tmv3_linkFormatRepairTargets() {
  const rows=tmv3_sheet_(TMV3.SHEETS.AUDIT).getDataRange().getValues(),out={Install:[],Delivery:[],Service:[]};
  rows.slice(1).forEach(function(row){const at=new Date(row[0]).getTime(),v=String(row[1]||''),id=String(row[2]||'');if(at>=Date.parse('2026-10-26T20:26:00Z')&&at<Date.parse('2026-10-26T20:50:00Z')&&out[v]&&row[4]==='STEP7_ALL'&&row[5]==='PASS'&&id&&out[v].indexOf(id)<0)out[v].push(id);});
  return out;
}
function tmv3_runLinkFormatRepair(vertical,eventIds) {
  if(['Install','Delivery','Service'].indexOf(vertical)<0||!Array.isArray(eventIds)||!eventIds.length)throw Error('FORMAT_REPAIR_SCOPE_REQUIRED');
  return tmv3_verticalProductionCycle_(vertical,{fromDay:'2026-10-26',toDay:'2026-11-02',eventIds:eventIds,forceLinkFormat:true});
}


// Original: 80_MorningOps
/************************************************************
 * TM V3 — MORNING OPERATIONS CONSOLE
 *
 * Purpose:
 * Answer, in one place:
 *   1) Is the system working?
 *   2) What needs attention?
 *   3) What should the operator do next?
 *
 * This module writes only to the V3 workbook. It never mutates
 * Striven or Google Calendar.
 ************************************************************/

function tmv3_refreshMorningOps() {
  const sh = tmv3_sheet_(TMV3.SHEETS.MORNING);
  const verticals = ['Install','Delivery','Service','PreInspection'];
  const today = new Date();
  const todayKey = Utilities.formatDate(today, TMV3_TIMEZONE, 'yyyy-MM-dd');
  const scheduleByTime = {};
  const attentionRows = [];
  const summaries = [];
  let totalAppointments = 0;
  let totalMapped = 0;
  let totalAttention = 0;
  let totalNoTask = 0;

  function customerName_(row) {
    const calendarCustomer = tmv3_clean_(row['Calendar Customer']);
    const nameMatch = calendarCustomer.match(/(?:^|\n)Name:\s*([^\n]+)/i);
    return (
      (nameMatch && tmv3_clean_(nameMatch[1])) ||
      tmv3_clean_(row['Calendar Event']).split('\n')[0] ||
      '(unnamed event)'
    );
  }

  function taskLabel_(row) {
    const taskText = tmv3_clean_(row['Task']);
    const taskMatch = taskText.match(/Task\s*#?\s*(\d+)/i);
    if (taskMatch) return '#' + taskMatch[1];
    if (!taskText || /^NOT RUN/i.test(taskText)) return 'NO TASK';
    return taskText.split('\n')[0];
  }

  function rowNeedsAttention_(row) {
    const mapping = tmv3_clean_(row['Mapping Status']).toUpperCase();
    const evidence = tmv3_clean_(row['Task Customer / IDs (Evidence)']);
    return /NOT_RUN|REVIEW|BLOCKED/.test(mapping) || /CREATE REQUIRED/i.test(evidence);
  }

  verticals.forEach(function(vertical) {
    const rows = tmv3_rows_(TMV3.VERTICALS[vertical].sheet).filter(function(row) {
      const d = tmv3_parseDateTime_(row['Calendar Start']);
      return d && Utilities.formatDate(d, TMV3_TIMEZONE, 'yyyy-MM-dd') === todayKey;
    });

    const unique = {};
    let mapped = 0;
    let attention = 0;

    rows.forEach(function(row) {
      const eventId = tmv3_clean_(row['Event ID']);
      if (eventId && unique[eventId]) return;
      if (eventId) unique[eventId] = true;

      const start = tmv3_parseDateTime_(row['Calendar Start']);
      const timeKey = start ? Utilities.formatDate(start, TMV3_TIMEZONE, 'HH:mm') : '99:99';
      const timeLabel = start ? Utilities.formatDate(start, TMV3_TIMEZONE, 'h:mm a') : 'Unscheduled';
      const customer = customerName_(row);
      const task = taskLabel_(row);
      const needsAttention = rowNeedsAttention_(row);
      const hasTask = task !== 'NO TASK';

      if (!scheduleByTime[timeKey]) {
        scheduleByTime[timeKey] = {
          label: timeLabel,
          Install: [],
          Delivery: [],
          Service: [],
          PreInspection: [],
          flags: []
        };
      }

      const taskUrl = tmv3_clean_(row['Existing Task Link']);
      const calendarUrl = tmv3_clean_(row['Calendar Event Link']);
      const linkUrl = taskUrl || calendarUrl;
      const evidence = tmv3_clean_(row['Task Customer / IDs (Evidence)']);
      const issue = tmv3_clean_(row['Issue / Next Action']);
      const locationReview = /CREATE REQUIRED/i.test(evidence);
      const reviewNote = needsAttention
        ? [
            'REVIEW NEEDED',
            '',
            taskUrl ? ('Task: ' + task) : 'No Striven Task exists yet.',
            'Issue: ' + (
              locationReview
                ? 'Customer-owned Location is unresolved.'
                : (issue || 'Appointment requires review.')
            ),
            '',
            'Next action: ' + (
              locationReview
                ? 'Verify the correct customer-owned Location before any relationship write.'
                : (issue || 'Review and correct the source appointment, then rerun V3.')
            )
          ].join('\n')
        : '';

      scheduleByTime[timeKey][vertical].push({
        display: customer + (task ? '\n' + task : ''),
        url: linkUrl,
        note: reviewNote
      });

      if (needsAttention) {
        scheduleByTime[timeKey].flags.push('ATTENTION');
        attention++;

        attentionRows.push({
          values: [
            locationReview ? 'CHECK' : 'ATTENTION',
            timeLabel,
            vertical,
            customer,
            locationReview
              ? 'Customer-owned Location unresolved'
              : (issue || 'Appointment requires attention'),
            locationReview
              ? 'Verify Location before any relationship write'
              : (issue || 'Review appointment')
          ],
          url: linkUrl,
          note: reviewNote,
          taskUrl: taskUrl,
          calendarUrl: calendarUrl,
          eventId: eventId
        });
      }

      if (hasTask) mapped++;
      else totalNoTask++;
    });

    const appointments = Object.keys(unique).length;
    totalAppointments += appointments;
    totalMapped += mapped;
    totalAttention += attention;

    summaries.push({
      vertical: vertical,
      appointments: appointments,
      mapped: mapped,
      attention: attention
    });
  });

  const discovery=tmv3_preInspectionDiscoveryStatus();
  (discovery.records || []).filter(function(record){return record.status==='REVIEW_VISIT_TYPE' || record.status==='INSTALLATION';}).forEach(function(record){
    attentionRows.push({values:['CHECK',tmv3_compactDateTime_(record.start),record.status==='INSTALLATION'?'Install':'PreInspection?',record.title,
      record.reason,record.status==='INSTALLATION'?'Route to Installation workflow; no PreInspection Task.':'Review customer, site and current Sales Order evidence.'],
      url:tmv3_calendarEventUrl_(record.eventId,record.calendarId),calendarUrl:tmv3_calendarEventUrl_(record.eventId,record.calendarId),
      taskUrl:'',eventId:record.eventId,note:record.status==='INSTALLATION' ? 'Verified approved/in-progress Order at this job site. Route to Install; no PreInspection Task.' : 'Likely visit awaiting verified customer, job site and current Order classification.'});
  });
  if(discovery.truncated)attentionRows.push({values:['CHECK','','PreInspection?','Additional candidates',
    'Discovery has '+discovery.total+' candidates; compact display is limited.','Use tmv3_refreshPreInspectionDiscovery / audit source calendars.'],url:'',note:'',taskUrl:'',calendarUrl:'',eventId:''});
  const values = [];
  values.push(['MORNING OPS — TODAY','','','','','']);
  values.push([
    Utilities.formatDate(today, TMV3_TIMEZONE, 'EEEE, MMMM d, yyyy'),
    '','','','',TMV3.MODE
  ]);
  values.push(['','','','','','']);
  values.push(['TODAY',totalAppointments,'MAPPED',totalMapped,'ATTENTION',totalAttention]);
  values.push(['NO TASK',totalNoTask,'INSTALL',(summaries[0] || {}).appointments || 0,'DELIVERY',(summaries[1] || {}).appointments || 0]);
  values.push(['SERVICE',(summaries[2] || {}).appointments || 0,'PREINSPECTION',(summaries[3] || {}).appointments || 0,'','']);
  values.push(['DISCOVERY',discovery.total || 0,'VISIT TYPE REVIEW',discovery.review || 0,'','']);
  values.push(["TODAY'S SCHEDULE",'','','','','']);
  values.push(['Time','Install','Delivery','Service','PreInspection','Flag']);

  Object.keys(scheduleByTime).sort().forEach(function(timeKey) {
    const slot = scheduleByTime[timeKey];
    values.push([
      slot.label,
      slot.Install.map(function(x){ return x.display; }).join('\n\n'),
      slot.Delivery.map(function(x){ return x.display; }).join('\n\n'),
      slot.Service.map(function(x){ return x.display; }).join('\n\n'),
      slot.PreInspection.map(function(x){ return x.display; }).join('\n\n'),
      slot.flags.length ? 'ATTENTION' : ''
    ]);
  });

  values.push(['','','','','','']);
  values.push(['NEEDS ATTENTION','','','','','']);
  values.push(['Priority','Time','Vertical','Customer / Event','Issue','Next Action']);

  if (!attentionRows.length) {
    values.push(['OK','','','','No attention items for today.','']);
  } else {
    attentionRows.forEach(function(item) {
      values.push(item.values.slice(0,6));
    });
  }

  // Full-grid reset: Morning Ops must never retain stale rows/columns from an older layout.
  sh.getRange(1,1,sh.getMaxRows(),sh.getMaxColumns())
    .breakApart()
    .clearContent()
    .clearFormat();

  if (sh.getMaxColumns() < 7) {
    sh.insertColumnsAfter(sh.getMaxColumns(), 7 - sh.getMaxColumns());
  }
  if (sh.getMaxRows() < values.length) {
    sh.insertRowsAfter(sh.getMaxRows(), values.length - sh.getMaxRows());
  }

  sh.getRange(1,1,values.length,6).setValues(values);

  // Visual snapshot layout.
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(2);
  sh.showColumns(1,6);
  if (sh.getMaxColumns() > 6) {
    sh.hideColumns(7,sh.getMaxColumns() - 6);
  }

  sh.getRange(1,1,1,6).merge()
    .setBackground('#1f1f1f')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sh.getRange(2,1,1,5).merge()
    .setBackground('#1f1f1f')
    .setFontColor('#e6e6e6')
    .setFontWeight('bold');
  sh.getRange(2,6)
    .setBackground('#1f1f1f')
    .setFontColor('#e6e6e6')
    .setFontWeight('bold')
    .setHorizontalAlignment('right');

  const sectionRows = [8, values.findIndex(function(r){ return r[0] === 'NEEDS ATTENTION'; }) + 1]
    .filter(function(r){ return r > 0; });
  sectionRows.forEach(function(rowNum) {
    sh.getRange(rowNum,1,1,6).merge()
      .setBackground('#e9edf2')
      .setFontWeight('bold')
      .setFontSize(11);
  });

  // KPI blocks.
  sh.getRange(4,1,1,2).setBackground('#dceeff');
  sh.getRange(4,3,1,2).setBackground('#def3e4');
  sh.getRange(4,5,1,2).setBackground('#ffe5bd');
  sh.getRange(5,1,1,2).setBackground(totalNoTask ? '#ffd9d9' : '#def3e4');
  sh.getRange(5,3,2,4).setBackground('#f3f4f6');
  sh.getRange(4,1,3,6)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  const scheduleHeaderRow = 9;
  sh.getRange(scheduleHeaderRow,1,1,6)
    .setBackground('#2f3742')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const scheduleStart = 10;
  const scheduleEnd = scheduleStart + Object.keys(scheduleByTime).length - 1;
  if (scheduleEnd >= scheduleStart) {
    sh.getRange(scheduleStart,2,scheduleEnd-scheduleStart+1,1).setBackground('#e8f1ff');
    sh.getRange(scheduleStart,3,scheduleEnd-scheduleStart+1,1).setBackground('#eaf7ec');
    sh.getRange(scheduleStart,4,scheduleEnd-scheduleStart+1,1).setBackground('#f4ecff');
    sh.getRange(scheduleStart,5,scheduleEnd-scheduleStart+1,1).setBackground('#fff5dc');
    sh.getRange(scheduleStart,1,scheduleEnd-scheduleStart+1,6)
      .setVerticalAlignment('middle')
      .setWrap(true);

    for (let r = scheduleStart; r <= scheduleEnd; r++) {
      const flag = tmv3_clean_(sh.getRange(r,6).getValue());
      if (flag === 'ATTENTION') {
        sh.getRange(r,1,1,6).setFontWeight('bold');
        sh.getRange(r,6).setBackground('#ffd59a');
      }
    }
  }

  // Make the snapshot actionable: task cells link to Striven; no-task items link to Calendar.
  const verticalColumn = { Install: 2, Delivery: 3, Service: 4, PreInspection: 5 };
  Object.keys(scheduleByTime).sort().forEach(function(timeKey, slotIndex) {
    const slot = scheduleByTime[timeKey];
    const rowNum = scheduleStart + slotIndex;

    ['Install','Delivery','Service','PreInspection'].forEach(function(vertical) {
      const entries = slot[vertical] || [];
      if (!entries.length) return;

      const cell = sh.getRange(rowNum, verticalColumn[vertical]);
      const text = entries.map(function(x){ return x.display; }).join('\n\n');
      const builder = SpreadsheetApp.newRichTextValue().setText(text);
      let offset = 0;

      entries.forEach(function(entry, idx) {
        if (entry.url) {
          builder.setLinkUrl(offset, offset + entry.display.length, entry.url);
        }
        offset += entry.display.length + (idx < entries.length - 1 ? 2 : 0);
      });

      cell.setRichTextValue(builder.build());

      const notes = entries
        .map(function(x){ return x.note; })
        .filter(Boolean);
      if (notes.length) {
        cell.setNote(notes.join('\n\n--------------------\n\n'));
      } else {
        cell.setNote('');
      }
    });
  });

  const attentionTitleRow = sectionRows[1];
  const attentionHeaderRow = attentionTitleRow + 1;
  const attentionDataStart = attentionHeaderRow + 1;
  sh.getRange(attentionHeaderRow,1,1,6)
    .setBackground('#2f3742')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  for (let i = 0; i < Math.max(1, attentionRows.length); i++) {
    const rowNum = attentionDataStart + i;
    const priority = tmv3_clean_(sh.getRange(rowNum,1).getValue());
    sh.getRange(rowNum,1,1,6)
      .setBackground(priority === 'ATTENTION' ? '#ffe0c2' : '#fff2bf')
      .setWrap(true)
      .setVerticalAlignment('middle');
  }

  attentionRows.forEach(function(item, index) {
    const rowNum = attentionDataStart + index;
    const customerCell = sh.getRange(rowNum, 4);
    if (item.url) {
      customerCell.setRichTextValue(
        SpreadsheetApp.newRichTextValue()
          .setText(item.values[3])
          .setLinkUrl(item.url)
          .build()
      );
    }
    customerCell.setNote(item.note || '');
    sh.getRange(rowNum, 5).setNote(item.note || '');
  });

  [90,180,180,180,180,300].forEach(function(px,index) {
    sh.setColumnWidth(index + 1, px);
  });

  sh.setRowHeight(1, 30);
  sh.setRowHeight(2, 24);
  sh.autoResizeRows(3, Math.max(1, values.length - 2));

  return {
    status: 'TODAY_SNAPSHOT',
    date: todayKey,
    mode: TMV3.MODE,
    appointments: totalAppointments,
    mapped: totalMapped,
    attention: totalAttention,
    noTask: totalNoTask,
    verticals: summaries
  };
}

function tmv3_morningStateByEvent_(rows) {
  const out = {};

  (rows || []).forEach(function(row) {
    const key =
      tmv3_clean_(row['Vertical']) +
      '|' +
      tmv3_clean_(row['Event ID']);

    if (!out[key]) {
      out[key] = row;
      return;
    }

    const current = out[key];
    const currentDetected = tmv3_parseDate_(current['First Detected At']);
    const nextDetected = tmv3_parseDate_(row['First Detected At']);

    if (
      nextDetected &&
      (!currentDetected || nextDetected < currentDetected)
    ) {
      out[key] = row;
    }
  });

  return out;
}

function tmv3_morningExceptionRecord_(vertical, row, state) {
  const status = tmv3_clean_(row['Status']).toUpperCase();
  const verification = tmv3_clean_(row['Verification']);
  const issue = tmv3_clean_(row['Issue']);
  const errorCode = tmv3_clean_(state['Last Error Class'] || state['Error Code']);
  const eventId = tmv3_clean_(row['Event ID']);

  const needsAttention =
    ['MATCHED','IGNORED'].indexOf(status) === -1;

  let priority = 'INFO';
  let rank = 9;

  if (
    /failed|error|uncertain|readback|technical|exception/i.test(issue + ' ' + errorCode)
  ) {
    priority = 'RED';
    rank = 0;
  } else if (status === 'BLOCKED' || status === 'NOT MATCHED') {
    priority = 'ORANGE';
    rank = 1;
  } else if (status === 'REVIEW') {
    priority = 'ORANGE';
    rank = 2;
  } else if (status.indexOf('READY') === 0) {
    priority = 'YELLOW';
    rank = 3;
  } else if (status === 'CONFIRMED') {
    priority = 'YELLOW';
    rank = 4;
  }

  const ageMinutes = tmv3_minutesSince_(state['First Detected At']);
  const acknowledgedBy = tmv3_clean_(state['Acknowledged By']);
  const acknowledgedAt = tmv3_clean_(state['Acknowledged At']);

  return {
    needsAttention: needsAttention,
    priority: priority,
    rank: rank,
    vertical: vertical,
    date: tmv3_morningDateValue_(row['Date']),
    time: tmv3_morningTimeValue_(row['Time']),
    appointment: tmv3_clean_(row['Calendar Title']),
    customer: tmv3_clean_(row['Customer']),
    task: tmv3_clean_(row['Task']),
    data: tmv3_clean_(row['Data']),
    status: status,
    issue: issue || verification || 'Needs review.',
    ageMinutes: ageMinutes,
    attempts: Number(state['Attempt Count'] || 0),
    nextAction: tmv3_clean_(row['Action'] || row['Next Action']) || 'REVIEW',
    acknowledged:
      acknowledgedBy
        ? (
            acknowledgedBy +
            (acknowledgedAt ? ' · ' + tmv3_compactDateTime_(acknowledgedAt) : '')
          )
        : 'Not acknowledged',
    eventId: eventId,
    source: vertical,
    operatorRow: Number(row.__operatorRow || 0)
  };
}

function tmv3_morningDateValue_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return tmv3_date_(value);
  }
  return tmv3_clean_(value);
}

function tmv3_morningTimeValue_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return tmv3_time_(value);
  }
  return tmv3_clean_(value);
}

function tmv3_morningExceptionSort_(a, b) {
  // Technical failures always come first.
  if (a.rank === 0 || b.rank === 0) {
    if (a.rank !== b.rank) return a.rank - b.rank;
  }

  // Then prioritize the appointment that is operationally closest.
  const au = tmv3_morningUrgencyRank_(a.date);
  const bu = tmv3_morningUrgencyRank_(b.date);
  if (au !== bu) return au - bu;

  // Within the same urgency window, blocked/review outrank planned actions.
  if (a.rank !== b.rank) return a.rank - b.rank;

  const aa = a.ageMinutes === null ? -1 : a.ageMinutes;
  const bb = b.ageMinutes === null ? -1 : b.ageMinutes;
  if (aa !== bb) return bb - aa;

  return (
    (a.date + ' ' + a.time)
      .localeCompare(b.date + ' ' + b.time)
  );
}

function tmv3_morningUrgencyRank_(dateValue) {
  const text = tmv3_clean_(dateValue);
  const d = new Date(text + 'T00:00:00');
  if (isNaN(d.getTime())) return 9;

  const today = new Date();
  today.setHours(0,0,0,0);
  d.setHours(0,0,0,0);

  const days = Math.floor((d.getTime() - today.getTime()) / 86400000);

  if (days <= 0) return 0;
  if (days === 1) return 1;
  if (days <= 3) return 2;
  if (days <= 7) return 3;
  return 4;
}

function tmv3_lastSuccessfulRun_(auditRows) {
  let latest = null;

  (auditRows || []).forEach(function(row) {
    if (
      tmv3_clean_(row['Action']) !== 'SHADOW_RUN' ||
      tmv3_clean_(row['Result']).toUpperCase() !== 'PASS'
    ) {
      return;
    }

    const dt = tmv3_parseDate_(row['Timestamp']);
    if (dt && (!latest || dt > latest)) latest = dt;
  });

  return latest;
}

function tmv3_recentTechnicalFailures_(auditRows, hours) {
  const cutoff = new Date(Date.now() - Number(hours || 24) * 3600000);

  return (auditRows || []).filter(function(row) {
    const dt = tmv3_parseDate_(row['Timestamp']);
    if (!dt || dt < cutoff) return false;

    const result = tmv3_clean_(row['Result']).toUpperCase();
    const detail = tmv3_clean_(row['Detail']);

    return (
      ['FAIL','FAILED','ERROR','CREATE UNCERTAIN'].indexOf(result) !== -1 ||
      /readback failed|exception|technical failure|rollback failed/i.test(detail)
    );
  });
}

function tmv3_regressionHealth_() {
  const rows = tmv3_rows_('Regression');
  const counts = {
    agree: 0,
    intentional: 0,
    investigate: 0,
    notRun: 0,
    newEvents: 0,
    total: rows.length
  };

  rows.forEach(function(row) {
    const value = tmv3_clean_(row['Comparison']).toUpperCase();

    if (value === 'AGREE') counts.agree++;
    else if (value === 'INTENTIONAL IMPROVEMENT') counts.intentional++;
    else if (value === 'INVESTIGATE') counts.investigate++;
    else if (value === 'NEW V3 EVENT') counts.newEvents++;
    else counts.notRun++;
  });

  return counts;
}

function tmv3_triggerHealth_() {
  const installed = tmv3_listTriggers();
  const calendarTriggerCount =
    typeof tmv3_step1CalendarIds_ === 'function'
      ? tmv3_step1CalendarIds_().length
      : 0;

  const expectedCounts = {
    tmv3_dailySourceRefresh: 1,
    tmv3_scheduledOperations: 5,
    tmv3_refreshLinksSlot_0800: 1,
    tmv3_refreshLinksSlot_1000: 1,
    tmv3_refreshLinksSlot_1200: 1,
    tmv3_refreshLinksSlot_1400: 1,
    tmv3_refreshLinksSlot_1600: 1,
    tmv3_refreshLinksSlot_1800: 1,
    tmv3_installReminderCheck: 1,
    tmv3_calendarEventUpdated: calendarTriggerCount,
    tmv3_calendarReconciliationFallback: 1
  };

  const actualCounts = {};
  installed.forEach(function(t) {
    const handler = tmv3_clean_(t.handler);
    actualCounts[handler] = (actualCounts[handler] || 0) + 1;
  });

  const missing = [];
  const extra = [];

  Object.keys(expectedCounts).forEach(function(handler) {
    const expected = expectedCounts[handler];
    const actual = actualCounts[handler] || 0;
    if (actual < expected) {
      missing.push(handler + ' ' + actual + '/' + expected);
    } else if (actual > expected) {
      extra.push(handler + ' ' + actual + '/' + expected);
    }
  });

  return {
    installed: installed.length,
    missing: missing,
    duplicates: extra,
    enabled: Object.keys(actualCounts).some(function(handler) {
      return expectedCounts[handler] !== undefined;
    }),
    expectedCount: Object.keys(expectedCounts).reduce(function(total, handler) {
      return total + Number(expectedCounts[handler] || 0);
    }, 0),
    manualWritesEnabled: tmv3_operationWritesEnabled_('MANUAL'),
    automationWritesEnabled: tmv3_operationWritesEnabled_('AUTO')
  };
}

function tmv3_morningReadiness_(input) {
  const checks = [];

  checks.push({
    name: 'System heartbeat',
    status: input.lastRun ? 'PASS' : 'PENDING',
    meaning: input.lastRun
      ? 'Last successful V3 run ' + tmv3_ageLabel_(tmv3_minutesSince_(input.lastRun)) + ' ago.'
      : 'No independent V3 run has completed yet.'
  });

  checks.push({
    name: 'Calendar sources fresh',
    status: input.lastRun ? 'PASS' : 'PENDING',
    meaning: input.lastRun
      ? 'Calendar intake completed in the last successful V3 run.'
      : 'Waiting for first successful V3 Calendar intake.'
  });

  checks.push({
    name: 'Striven sources fresh',
    status:
      input.missingConfig.length
        ? 'ATTENTION'
        : (input.lastRun ? 'PASS' : 'PENDING'),
    meaning:
      input.missingConfig.length
        ? 'Missing configuration: ' + input.missingConfig.join(', ')
        : (
            input.lastRun
              ? 'Striven source refresh completed in the last successful V3 run.'
              : 'Configuration is present; waiting for first successful V3 run.'
          )
  });

  const triggerProblem =
    input.triggers.duplicates.length > 0 ||
    (
      input.triggers.enabled &&
      input.triggers.missing.length > 0
    );

  checks.push({
    name: 'Managed triggers healthy',
    status:
      TMV3.MODE === 'SHADOW_READ_ONLY' && !input.triggers.enabled
        ? 'PENDING'
        : (triggerProblem ? 'ATTENTION' : 'PASS'),
    meaning:
      TMV3.MODE === 'SHADOW_READ_ONLY' && !input.triggers.enabled
        ? 'Scheduled production automation is intentionally not enabled.'
        : (
            triggerProblem
              ? 'Missing/duplicate trigger(s): ' +
                input.triggers.missing.concat(input.triggers.duplicates).join(', ')
              : 'Expected managed triggers are installed once each.'
          )
  });

  checks.push({
    name: 'No failed writes',
    status:
      input.failures.length
        ? 'ATTENTION'
        : 'PASS',
    meaning:
      input.failures.length
        ? input.failures.length + ' technical failure(s) in the last 24 hours.'
        : (
            tmv3_writesEnabled_()
              ? 'No failed V3 mutations in the last 24 hours.'
              : 'External V3 writes remain disabled.'
          )
  });

  let overdue = 0;
  input.summaries.forEach(function(summary) {
    if (
      summary.oldestMinutes !== null &&
      summary.oldestMinutes > 10
    ) {
      overdue++;
    }
  });

  checks.push({
    name: 'No overdue exceptions',
    status:
      input.lastRun
        ? (overdue ? 'ATTENTION' : 'PASS')
        : 'UNKNOWN',
    meaning:
      input.lastRun
        ? (
            overdue
              ? overdue + ' vertical(s) have unresolved items older than the 10-minute operating SLO.'
              : 'No unresolved V3 exception is older than 10 minutes.'
          )
        : 'Age tracking begins with the first independent V3 run.'
  });

  checks.push({
    name: 'Regression safety gate',
    status:
      input.regression.investigate
        ? 'ATTENTION'
        : (
            input.regression.notRun
              ? 'PENDING'
              : 'PASS'
          ),
    meaning:
      input.regression.investigate
        ? input.regression.investigate + ' old-vs-V3 difference(s) need investigation.'
        : (
            input.regression.notRun
              ? input.regression.notRun + ' baseline row(s) have not been compared yet.'
              : 'No unexplained old-vs-V3 differences remain.'
          )
  });

  const attention = checks.filter(function(c) {
    return c.status === 'ATTENTION';
  }).length;

  const pending = checks.filter(function(c) {
    return c.status === 'PENDING' || c.status === 'UNKNOWN';
  }).length;

  let overallStatus = 'READY FOR OPERATIONS';
  if (attention) overallStatus = 'ATTENTION REQUIRED — ' + attention + ' CHECK(S)';
  else if (pending) overallStatus = 'NOT YET VERIFIED — ' + pending + ' CHECK(S) PENDING';

  return {
    overallStatus: overallStatus,
    checks: checks,
    passed: checks.filter(function(c) { return c.status === 'PASS'; }).length,
    total: checks.length
  };
}

function tmv3_writeMorningHeader_(sh, readiness, lastRun) {
  sh.getRange('A1').setValue('TASK MAPPING — MORNING OPERATIONS');
  sh.getRange('A3').setValue(readiness.overallStatus);
  sh.getRange('A4').setValue(
    TMV3.MODE +
    ' · ' +
    (
      lastRun
        ? 'Last verified ' + tmv3_compactDateTime_(lastRun)
        : 'No successful V3 run yet'
    )
  );
}

function tmv3_writeMorningWorkflowHealth_(sh, summaries) {
  const headers = [
    'Vertical','Appointments','Task rows','V3 verified',
    'Needs attention','Blocked','Morning status'
  ];

  sh.getRange(8, 1, 1, headers.length).setValues([headers]);

  const values = summaries.map(function(s) {
    return [
      s.vertical,
      s.appointments,
      s.taskRows,
      s.verified,
      s.needsAttention,
      s.blocked,
      s.status
    ];
  });

  sh.getRange(9, 1, 4, headers.length).clearContent();
  sh.getRange(9, 1, values.length, headers.length).setValues(values);
}

function tmv3_writeMorningExceptions_(sh, exceptions) {
  const headers = [
    'Status','Data','Date','Time','Customer','Task','Issue','Action',
    'Vertical','Event ID','Priority','Age','Attempts',
    'Acknowledged / Owner'
  ];

  const headerRow = 16;
  const firstRow = 17;
  const maxRows = 15;
  const visible = 8;

  sh.getRange(headerRow, 1, 1, headers.length).setValues([headers]);
  sh.getRange(firstRow, 1, maxRows, headers.length).clearContent();

  if (!exceptions.length) {
    sh.getRange(firstRow, 1, 1, visible).setValues([[
      'MATCHED',
      '✅ No operator action required',
      '',
      '',
      '',
      '',
      '',
      'NONE'
    ]]);
  } else {
    const values = exceptions.map(function(x) {
      return [
        x.status,
        x.data,
        x.date,
        x.time,
        x.customer,
        x.task,
        x.issue,
        x.nextAction,
        x.vertical,
        x.eventId,
        x.priority,
        x.ageMinutes === null ? 'Baseline' : tmv3_ageLabel_(x.ageMinutes),
        x.attempts || '',
        x.acknowledged
      ];
    });

    sh.getRange(firstRow, 1, values.length, headers.length).setValues(values);
  }

  const maxColumns = sh.getMaxColumns();
  sh.showColumns(1, Math.min(visible, maxColumns));
  if (maxColumns > visible) {
    sh.hideColumns(visible + 1, maxColumns - visible);
  }

  const widths = [110, 185, 95, 85, 180, 220, 320, 145];
  widths.forEach(function(width, index) {
    if (index + 1 <= maxColumns) sh.setColumnWidth(index + 1, width);
  });

  const visibleRows = Math.max(1, exceptions.length);
  sh.getRange(headerRow, 1, visibleRows + 1, visible)
    .setWrap(true)
    .setVerticalAlignment('middle');
  sh.getRange(headerRow, 1, 1, visible).setFontWeight('bold');
  sh.autoResizeRows(firstRow, visibleRows);

  if (sh.getMaxRows() >= 35) {
    sh.hideRows(35, sh.getMaxRows() - 34);
  }
}

function tmv3_writeMorningAutomation_(sh, input) {
  const regressionText =
    input.regression.investigate
      ? input.regression.investigate + ' INVESTIGATE'
      : (
          input.regression.notRun
            ? input.regression.notRun + ' NOT RUN'
            : (
                input.regression.agree +
                ' agree · ' +
                input.regression.intentional +
                ' intentional improvements'
              )
        );

  const triggerText =
    TMV3.MODE === 'SHADOW_READ_ONLY' && !input.triggers.enabled
      ? 'Not enabled in shadow'
      : (
          input.triggers.missing.length || input.triggers.duplicates.length
            ? 'Missing/duplicate managed triggers'
            : input.triggers.installed + ' installed'
        );

  const rows = [
    ['Check','Current state','What good looks like','Action'],
    ['V3 mode',TMV3.MODE,'Explicit safe mode','OK'],
    ['System heartbeat',
      input.lastRun ? tmv3_compactDateTime_(input.lastRun) : 'No successful run',
      'Recent successful end-to-end run',
      input.lastRun ? 'OK' : 'ATTENTION'
    ],
    ['Striven configuration',
      input.missingConfig.length ? 'Missing: ' + input.missingConfig.join(', ') : 'Complete',
      'All required capabilities present',
      input.missingConfig.length ? 'ATTENTION' : 'OK'
    ],
    ['Managed triggers',triggerText,'Expected = installed = recently ran',
      (
        TMV3.MODE === 'SHADOW_READ_ONLY' && !input.triggers.enabled
          ? 'PENDING'
          : (
              input.triggers.missing.length || input.triggers.duplicates.length
                ? 'ATTENTION'
                : 'OK'
            )
      )
    ],
    ['Regression',regressionText,'No unexplained differences',
      input.regression.investigate ? 'ATTENTION' : (input.regression.notRun ? 'PENDING' : 'OK')
    ],
    ['Technical failures',
      input.failures.length + ' in last 24h',
      '0',
      input.failures.length ? 'ATTENTION' : 'OK'
    ],
    ['Calendar link writeback',
      tmv3_writesEnabled_() ? 'Enabled with readback' : 'Writer ready · writes disabled',
      'Write + readback verified',
      tmv3_writesEnabled_() ? 'OK' : 'PENDING'
    ],
    ['Business writes',
      tmv3_writesEnabled_() ? 'Enabled' : 'Disabled',
      TMV3.MODE === 'SHADOW_READ_ONLY' ? 'Disabled' : 'Guarded',
      TMV3.MODE === 'SHADOW_READ_ONLY' ? 'OK' : 'CHECK'
    ]
  ];

  sh.getRange(35, 1, 9, 4).clearContent();
  sh.getRange(35, 1, rows.length, 4).setValues(rows);
}

function tmv3_writeMorningReadiness_(sh, readiness) {
  const rows = [['Check','Status','Meaning']].concat(
    readiness.checks.map(function(c) {
      return [c.name, c.status, c.meaning];
    })
  );

  sh.getRange(46, 1, 8, 3).clearContent();
  sh.getRange(46, 1, rows.length, 3).setValues(rows);
  sh.getRange('A45').setValue(
    'MORNING READINESS — ' + readiness.passed + '/' + readiness.total + ' PASSED'
  );
}

function tmv3_writeMorningSinceRun_(sh, stateRows, lastRun, summaries, failures) {
  let newAppointments = 0;
  let resolved = 0;
  let review = 0;
  let blocked = 0;
  let oldest = null;

  (stateRows || []).forEach(function(row) {
    const firstSeen = tmv3_parseDate_(row['First Seen At']);
    const resolvedAt = tmv3_parseDate_(row['Resolved At']);
    const status = tmv3_clean_(row['State']).toUpperCase();

    if (lastRun && firstSeen && firstSeen > lastRun) newAppointments++;
    if (lastRun && resolvedAt && resolvedAt > lastRun) resolved++;
    if (status === 'REVIEW') review++;
    if (status === 'BLOCKED' || status === 'NOT MATCHED') blocked++;

    const age = tmv3_minutesSince_(row['First Detected At']);
    if (age !== null) oldest = oldest === null ? age : Math.max(oldest, age);
  });

  const rows = [
    ['Metric','Value'],
    ['Last successful V3 run',lastRun ? tmv3_compactDateTime_(lastRun) : 'Not yet completed'],
    ['New appointments',lastRun ? newAppointments : 'Pending'],
    ['Automatically resolved',lastRun ? resolved : 'Pending'],
    ['Operator reviews',review],
    ['Blocked',blocked],
    ['Technical failures',failures.length],
    ['Oldest unresolved',oldest === null ? 'None / pending' : tmv3_ageLabel_(oldest)]
  ];

  sh.getRange(57, 1, 8, 2).clearContent();
  sh.getRange(57, 1, rows.length, 2).setValues(rows);
}

function tmv3_acknowledgeSelectedMorningOps() {
  const ss = tmv3_ss_();
  const sh = ss.getActiveSheet();

  if (sh.getName() !== TMV3.SHEETS.MORNING) {
    throw new Error('Select an exception row on Morning Ops first.');
  }

  const row = sh.getActiveRange().getRow();
  const vertical = tmv3_clean_(sh.getRange(row, 3).getValue());
  const customer = tmv3_clean_(sh.getRange(row, 4).getValue());

  if (!vertical || !customer) {
    throw new Error('Select a row in NEEDS ATTENTION.');
  }

  const sourceRows = tmv3_rows_(TMV3.VERTICALS[vertical].sheet);
  const todayKey = Utilities.formatDate(new Date(), TMV3_TIMEZONE, 'yyyy-MM-dd');
  const matches = sourceRows.filter(function(sourceRow) {
    const d = tmv3_parseDateTime_(sourceRow['Calendar Start']);
    const sameDay = d && Utilities.formatDate(d, TMV3_TIMEZONE, 'yyyy-MM-dd') === todayKey;
    const calendarCustomer = tmv3_clean_(sourceRow['Calendar Customer']);
    const nameMatch = calendarCustomer.match(/(?:^|\n)Name:\s*([^\n]+)/i);
    const sourceCustomer =
      (nameMatch && tmv3_clean_(nameMatch[1])) ||
      tmv3_clean_(sourceRow['Calendar Event']).split('\n')[0];
    return sameDay && sourceCustomer === customer;
  });

  if (matches.length !== 1) {
    throw new Error('Could not resolve this attention row to exactly one today event.');
  }

  const eventId = tmv3_clean_(matches[0]['Event ID']);

  if (!vertical || !eventId) {
    throw new Error('Selected row has no V3 Event identity.');
  }

  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt(
    'Acknowledge exception',
    'Optional note for this appointment:',
    ui.ButtonSet.OK_CANCEL
  );

  if (prompt.getSelectedButton() !== ui.Button.OK) {
    return { status: 'CANCELLED' };
  }

  const activeEmail = tmv3_clean_(Session.getActiveUser().getEmail());
  const owner = activeEmail || 'Operator';
  const note = tmv3_clean_(prompt.getResponseText());
  const now = tmv3_now_();

  const stateSheet = tmv3_sheet_(TMV3.SHEETS.STATE);
  const values = stateSheet.getDataRange().getValues();
  const headers = values[0].map(tmv3_clean_);
  const ix = {};
  headers.forEach(function(h, i) { ix[h] = i; });

  let changed = 0;

  for (let i = 1; i < values.length; i++) {
    if (
      tmv3_clean_(values[i][ix['Vertical']]) === vertical &&
      tmv3_clean_(values[i][ix['Event ID']]) === eventId
    ) {
      values[i][ix['Acknowledged By']] = owner;
      values[i][ix['Acknowledged At']] = now;
      values[i][ix['Acknowledgement Note']] = note;
      changed++;
    }
  }

  if (!changed) {
    throw new Error('No durable V3 state row was found for the selected Event.');
  }

  stateSheet
    .getRange(1, 1, values.length, values[0].length)
    .setValues(values);

  tmv3_audit_(
    vertical,
    eventId,
    '',
    'ACKNOWLEDGE_EXCEPTION',
    'PASS',
    owner + (note ? ': ' + note : '')
  );

  tmv3_refreshMorningOps();

  return {
    status: 'ACKNOWLEDGED',
    vertical: vertical,
    eventId: eventId,
    owner: owner
  };
}

function tmv3_clearAcknowledgementSelectedMorningOps() {
  const ss = tmv3_ss_();
  const sh = ss.getActiveSheet();

  if (sh.getName() !== TMV3.SHEETS.MORNING) {
    throw new Error('Select an exception row on Morning Ops first.');
  }

  const row = sh.getActiveRange().getRow();
  const vertical = tmv3_clean_(sh.getRange(row, 3).getValue());
  const customer = tmv3_clean_(sh.getRange(row, 4).getValue());

  if (!vertical || !customer) {
    throw new Error('Select a row in NEEDS ATTENTION.');
  }

  const sourceRows = tmv3_rows_(TMV3.VERTICALS[vertical].sheet);
  const todayKey = Utilities.formatDate(new Date(), TMV3_TIMEZONE, 'yyyy-MM-dd');
  const matches = sourceRows.filter(function(sourceRow) {
    const d = tmv3_parseDateTime_(sourceRow['Calendar Start']);
    const sameDay = d && Utilities.formatDate(d, TMV3_TIMEZONE, 'yyyy-MM-dd') === todayKey;
    const calendarCustomer = tmv3_clean_(sourceRow['Calendar Customer']);
    const nameMatch = calendarCustomer.match(/(?:^|\n)Name:\s*([^\n]+)/i);
    const sourceCustomer =
      (nameMatch && tmv3_clean_(nameMatch[1])) ||
      tmv3_clean_(sourceRow['Calendar Event']).split('\n')[0];
    return sameDay && sourceCustomer === customer;
  });

  if (matches.length !== 1) {
    throw new Error('Could not resolve this attention row to exactly one today event.');
  }

  const eventId = tmv3_clean_(matches[0]['Event ID']);

  const stateSheet = tmv3_sheet_(TMV3.SHEETS.STATE);
  const values = stateSheet.getDataRange().getValues();
  const headers = values[0].map(tmv3_clean_);
  const ix = {};
  headers.forEach(function(h, i) { ix[h] = i; });

  let changed = 0;

  for (let i = 1; i < values.length; i++) {
    if (
      tmv3_clean_(values[i][ix['Vertical']]) === vertical &&
      tmv3_clean_(values[i][ix['Event ID']]) === eventId
    ) {
      values[i][ix['Acknowledged By']] = '';
      values[i][ix['Acknowledged At']] = '';
      values[i][ix['Acknowledgement Note']] = '';
      changed++;
    }
  }

  if (changed) {
    stateSheet
      .getRange(1, 1, values.length, values[0].length)
      .setValues(values);
  }

  tmv3_refreshMorningOps();

  return {
    status: changed ? 'CLEARED' : 'NOT_FOUND',
    vertical: vertical,
    eventId: eventId
  };
}

function tmv3_parseDate_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function tmv3_minutesSince_(value) {
  const d = tmv3_parseDate_(value);
  if (!d) return null;

  return Math.max(
    0,
    Math.floor((Date.now() - d.getTime()) / 60000)
  );
}

function tmv3_ageLabel_(minutes) {
  if (minutes === null || minutes === undefined || isNaN(minutes)) return 'Unknown';

  const m = Math.max(0, Math.floor(Number(minutes)));
  if (m < 60) return m + 'm';

  const h = Math.floor(m / 60);
  const rem = m % 60;

  if (h < 24) return h + 'h ' + rem + 'm';

  const d = Math.floor(h / 24);
  const hr = h % 24;
  return d + 'd ' + hr + 'h';
}

function tmv3_compactDateTime_(value) {
  const d = tmv3_parseDate_(value);
  if (!d) return tmv3_clean_(value);

  return Utilities.formatDate(
    d,
    'America/Toronto',
    'MMM d, h:mm a'
  );
}


function tmv3_writeMorningPipeline_(sh, auditRows, lastRun) {
  const refresh = tmv3_latestAuditAction_(auditRows, 'REFRESH_SOURCES');
  const regression = tmv3_latestAuditAction_(auditRows, 'REGRESSION_COMPARE');

  const rows = [
    ['Stage','Last success','Freshness','Status'],
    [
      'Calendar read',
      lastRun ? tmv3_compactDateTime_(lastRun) : 'Not run',
      lastRun ? tmv3_ageLabel_(tmv3_minutesSince_(lastRun)) : 'Pending',
      lastRun ? 'PASS' : 'PENDING'
    ],
    [
      'Striven source refresh',
      refresh ? tmv3_compactDateTime_(refresh.date) : 'Not run',
      refresh ? tmv3_ageLabel_(tmv3_minutesSince_(refresh.date)) : 'Pending',
      refresh ? 'PASS' : 'PENDING'
    ],
    [
      'Mapping + identity resolution',
      lastRun ? tmv3_compactDateTime_(lastRun) : 'Not run',
      lastRun ? tmv3_ageLabel_(tmv3_minutesSince_(lastRun)) : 'Pending',
      lastRun ? 'PASS' : 'PENDING'
    ],
    [
      'Regression comparison',
      regression ? tmv3_compactDateTime_(regression.date) : 'Not run',
      regression ? tmv3_ageLabel_(tmv3_minutesSince_(regression.date)) : 'Pending',
      regression
        ? (tmv3_clean_(regression.row['Result']).toUpperCase() === 'PASS' ? 'PASS' : 'ATTENTION')
        : 'PENDING'
    ],
    [
      'Calendar link writeback',
      tmv3_writesEnabled_() ? 'Production/canary runs' : 'Disabled in shadow',
      tmv3_writesEnabled_() ? 'Guarded' : 'N/A',
      tmv3_writesEnabled_() ? 'CHECK' : 'SAFE'
    ]
  ];

  sh.getRange('A66').setValue('PIPELINE FRESHNESS');
  sh.getRange(67, 1, 6, 4).clearContent();
  sh.getRange(67, 1, rows.length, 4).setValues(rows);
}

function tmv3_latestAuditAction_(auditRows, action) {
  let latest = null;

  (auditRows || []).forEach(function(row) {
    if (tmv3_clean_(row['Action']) !== action) return;

    const d = tmv3_parseDate_(row['Timestamp']);
    if (!d) return;

    if (!latest || d > latest.date) {
      latest = {
        date: d,
        row: row
      };
    }
  });

  return latest;
}


function tmv3_operatorLinkFormula_(exceptionRecord) {
  const vertical = tmv3_clean_(exceptionRecord && exceptionRecord.vertical);
  const row = Number(exceptionRecord && exceptionRecord.operatorRow || 0);
  const label = tmv3_clean_(exceptionRecord && exceptionRecord.appointment) || 'Open record';

  if (!vertical || !row || !TMV3.VERTICALS[vertical]) {
    return label;
  }

  const sheet = tmv3_sheet_(TMV3.VERTICALS[vertical].sheet);
  const gid = sheet.getSheetId();
  const safeLabel = label.replace(/"/g, '""');

  return '=HYPERLINK("#gid=' + gid + '&range=A' + row + '","' + safeLabel + '")';
}


// Original: 90_Admin
function tmv3_healthCheck() {
  const requiredSheets = [
    TMV3.SHEETS.MORNING,TMV3.SHEETS.OVERVIEW,TMV3.SHEETS.INSTALL,TMV3.SHEETS.DELIVERY,TMV3.SHEETS.SERVICE,
    TMV3.SHEETS.PREINSPECTION,TMV3.SHEETS.STATE,TMV3.SHEETS.AUDIT,TMV3.SHEETS.CUSTOMERS,
    TMV3.SHEETS.LOCATIONS,TMV3.SHEETS.CONTACTS,TMV3.SHEETS.ORDERS,TMV3.SHEETS.TASKS
  ];
  const ss = tmv3_ss_();
  const missingSheets = requiredSheets.filter(function(n) { return !ss.getSheetByName(n); });
  const caps = tmv3_configCapabilities_();
  const missingCaps = Object.keys(caps).filter(function(k) { return !caps[k].present; });
  const result = {
    version:TMV3.VERSION,
    mode:TMV3.MODE,
    scriptId:ScriptApp.getScriptId(),
    spreadsheetId:ss.getId(),
    missingSheets:missingSheets,
    missingPropertyCapabilities:missingCaps,
    ready:missingSheets.length===0 && missingCaps.length===0
  };
  Logger.log(JSON.stringify(result));
  return result;
}

function tmv3_configKeysOnly() {
  return Object.keys(PropertiesService.getScriptProperties().getProperties()).sort();
}

function tmv3_importConfigBundle(bundle) {
  if (!bundle || typeof bundle !== 'object') throw new Error('Configuration bundle required.');
  const allowed = [
    'CLIENT_ID','CLIENT_SECRET','Striven_Customers_ReportAPI','Striven_CustomerLocations_ReportAPI',
    'Striven_ApprovedSalesOrders_ReportAPI','Striven_DeliveryApprovedOrders_ReportAPI',
    'Striven_InstallTasks_ReportAPI','Striven_DeliveryTasks_ReportAPI','Striven_ServiceTasks_ReportAPI','Striven_ServiceWorkOrders_ReportAPI'
  ];
  const safe = {};
  allowed.forEach(function(k) {
    if (bundle[k] !== undefined && bundle[k] !== null && tmv3_clean_(bundle[k]) !== '') safe[k] = String(bundle[k]);
  });
  if (Object.keys(safe).length !== allowed.length) {
    throw new Error('Config bundle incomplete. Expected ' + allowed.length + ' keys, received ' + Object.keys(safe).length + '.');
  }
  PropertiesService.getScriptProperties().setProperties(safe, false);
  return { status:'IMPORTED', keys:Object.keys(safe).sort() };
}

function tmv3_assertNoBusinessWrites() {
  const legacyWrites = tmv3_writesEnabled_();
  const manualWrites = tmv3_operationWritesEnabled_('MANUAL');
  const automationWrites = tmv3_operationWritesEnabled_('AUTO');
  const anyWrites = legacyWrites || manualWrites || automationWrites;

  return {
    mode: TMV3.MODE,
    businessWritesEnabled: anyWrites,
    calendarWritesEnabled: anyWrites,
    taskCreateEnabled: manualWrites || automationWrites,
    taskPatchEnabled: anyWrites,
    manualWritesEnabled: manualWrites,
    automationWritesEnabled: automationWrites,
    assertionPass: !anyWrites
  };
}


// Google Apps Script: current-day operator controls; each run executes at most one event.



// Google Apps Script: prioritize cached eligible PreInspection events.
// These cached values are only a low-cost scheduling hint. The production
// transaction always revalidates the customer, location, task history and
// duplicate guards using tmv3_preInspectionAutomationCycle_.




// Google Apps Script: bounded operator batch (maximum 3 event transactions).
// Every selected event gets its own lock, API guard, fresh identity check,
// duplicate check and readback. Do not parallelize task creations.


// Resolve only customer-owned location evidence; never creates or patches a Task.


// Retry one current-day event lacking a Task link. Cached rows select only;
// native fresh identity, history, API and durable create guards authorize writes.


// Read-only response metadata for current-day events without Task links.


// Read-only Striven reconciliation of current-day cached task schedules.


// Google Apps Script: date-scoped native controls; existing transaction guards remain mandatory.




// Google Apps Script: retrospective audit only. Never executes proposed actions or sends mail.