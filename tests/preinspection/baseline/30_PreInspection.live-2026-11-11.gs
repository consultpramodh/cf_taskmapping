// Google Apps Script — PreInspection V3 guarded cleanup, staging revision 2026-11-11-r5.
// IMPORTANT: This is ONE vertical module, not a standalone replacement for TMV3 shared utilities.
// The original 2026-11-10-r1 standard is an external dependency and must be verified.
// Live Striven and Calendar Task mutations, guest invites and email are OFF.
// Do not enable until full V3 integration and specific owner/RequestedBy rules pass.
var TMV3_PI_CLEANED_RELEASE_APPROVED = false;

function tmv3_preInspectionTaskRowArray_(task) {
  return [
    task['Task ID'],
    task['Task Number'],
    task['Task Type ID'],
    task['Task Type'],
    task['Status'],
    task['Name'],
    task['Customer ID'],
    task['Location ID'],
    task['Contact ID'],
    task['Order ID'],
    task['Start'],
    task['Due'],
    task['Assignees'],
    task['Pools'],
    task['URL'],
    task.Fingerprint
  ];
}

// Original: 20_Striven; Function: tmv3_preInspectionTaskRefreshTargets_
function tmv3_preInspectionTaskRefreshTargets_() {
  const step2Records = tmv3_step2CalendarRecords_()
    .filter(function(record) {
      return record.vertical === 'PreInspection';
    });
  const step3Records = tmv3_step3BusinessAnchorRecords_(
    step2Records,
    tmv3_step3AnchorIndex_()
  );
  const step4Records = tmv3_step4IdentityRecords_(
    step3Records,
    tmv3_step4IdentityIndex_()
  );

  return step4Records
    .filter(function(record) {
      return (
        record.step4 &&
        record.step4.disposition === 'VERIFIED' &&
        record.step4.customer
      );
    })
    .map(function(record) {
      return {
        eventId:tmv3_clean_(record.eventId),
        existingTaskId:tmv3_clean_(record.existingTaskId),
        customer:record.step4.customer
      };
    });
}

// Original: 20_Striven; Function: tmv3_searchPreInspectionTaskIdsForCustomer_
function tmv3_searchPreInspectionTaskIdsForCustomer_(customer) {
  const customerId = Number(
    tmv3_clean_(
      customer &&
      customer['Customer ID']
    ) || 0
  );

  if (!customerId) {
    throw new Error(
      'PreInspection task search requires Customer ID.'
    );
  }

  const customerPayload = { Id: customerId };
  const customerNumber = tmv3_clean_(
    customer['Customer Number']
  );
  const customerName = tmv3_clean_(
    customer['Name']
  );

  if (customerNumber) {
    customerPayload.Number = customerNumber;
  }
  if (customerName) {
    customerPayload.Name = customerName;
  }

  const maxPages = Math.max(
    1,
    Math.min(
      3,
      Number(
        TMV3.OPERATIONS &&
        TMV3.OPERATIONS.preInspectionTaskSearchMaxPages ||
        3
      )
    )
  );

  const ids = [];
  let previousSignature = '';
  let terminatedNaturally = false;

  for (
    let pageIndex = 0;
    pageIndex < maxPages;
    pageIndex++
  ) {
    const json = tmv3_fetchJson_(
      TMV3.API_BASE + '/v2/tasks/search',
      {
        method:'post',
        contentType:'application/json',
        payload:JSON.stringify({
          Customer:customerPayload,
          Type:[105],
          PageIndex:pageIndex,
          PageSize:100
        })
      }
    ) || {};

    const rows =
      tmv3_extractTaskSearchRows_(json);

    const signature = rows
      .slice(0, 5)
      .map(function(row) {
        return tmv3_clean_(
          tmv3_first_(
            row,
            ['id','Id','taskId','TaskId']
          )
        );
      })
      .join('|');

    if (
      pageIndex > 0 &&
      signature &&
      signature === previousSignature
    ) {
      throw new Error(
        'TMV3_PREINSPECTION_CUSTOMER_TASK_PAGINATION_REPEATED: Customer ' +
        customerId +
        ' repeated page ' +
        pageIndex +
        '.'
      );
    }

    previousSignature = signature;

    rows.forEach(function(row) {
      const id = tmv3_clean_(
        tmv3_first_(
          row,
          ['id','Id','taskId','TaskId']
        )
      );
      if (id) ids.push(id);
    });

    if (
      !rows.length ||
      rows.length < 100
    ) {
      terminatedNaturally = true;
      break;
    }
  }

  if (!terminatedNaturally) {
    throw new Error(
      'TMV3_PREINSPECTION_CUSTOMER_TASK_PAGINATION_TRUNCATED: Customer ' +
      customerId +
      ' reached the page cap without proving completion.'
    );
  }

  return tmv3_unique_(ids);
}

// Original: 20_Striven; Function: tmv3_preInspectionTaskSourceRows_
function tmv3_preInspectionTaskSourceRows_(options) {
  options = options || {};
  const prior = tmv3_rows_(TMV3.SHEETS.TASKS).filter(function(t) {
    return Number(t['Task Type ID']) === 105;
  });
  const props = PropertiesService.getScriptProperties();
  const key = 'TMV3_PREINSPECTION_TASK_CACHE_REFRESH_AT';
  const lastMs = Number(props.getProperty(key) || 0);
  const ttl = Math.max(60, Number((TMV3.OPERATIONS || {}).preInspectionTaskCacheMinMinutes || 360));
  const age = lastMs ? (Date.now() - lastMs) / 60000 : Infinity;
  const summary = {reused:true,targetCustomers:0,searchedCustomers:0,
    hydratedTasks:0,priorTasks:prior.length,
    ageMinutes:isFinite(age) ? Math.round(age) : null};
  if (options.force !== true && prior.length && age < ttl) {
    return Object.assign({rows:prior.map(tmv3_preInspectionTaskRowArray_)},summary);
  }
  const targets = tmv3_preInspectionTaskRefreshTargets_();
  if (!targets.length) return Object.assign({rows:prior.map(tmv3_preInspectionTaskRowArray_)},summary);
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
  const configuredCap=Number((TMV3.OPERATIONS || {}).preInspectionTaskHydrationMaxTasks || 60);
  const cap=Number.isSafeInteger(configuredCap)&&configuredCap>0?Math.min(100,configuredCap):60;
  if (ids.length > cap) throw new Error('PI_CACHE_HYDRATION_CAP_NO_REPLACE');
  const byId = {};
  // Retain prior tasks outside the targeted customer set. This refresh is not
  // an all-customer snapshot and may never erase unrelated cached task history.
  prior.forEach(function(task) {
    const id = String(tmv3_clean_(task['Task ID']));
    if (!id || byId[id]) throw new Error('PI_PRIOR_CACHE_DUPLICATE_OR_BLANK_TASK_ID');
    byId[id] = task;
  });
  ids.forEach(function(id) {
    const task = tmv3_getTaskById_(id);
    const owner = String(tmv3_clean_(task && task['Customer ID']));
    if (!task || String(tmv3_clean_(task['Task ID'])) !== id ||
        Number(task['Task Type ID']) !== 105 || !customers[owner]) {
      throw new Error('PI_CACHE_TASK_IDENTITY_CONFLICT_NO_REPLACE:' + id);
    }
    byId[id] = task;
  });
  // Commit freshness only after the entire bounded retrieval succeeds.
  props.setProperty(key,String(Date.now()));
  return {rows:Object.keys(byId).sort(function(a,b){return Number(a)-Number(b);})
      .map(function(id){return tmv3_preInspectionTaskRowArray_(byId[id]);}),
    reused:false,targetCustomers:customerIds.length,
    searchedCustomers:customerIds.length,hydratedTasks:ids.length,
    priorTasks:prior.length,ageMinutes:isFinite(age)?Math.round(age):null};
}

function tmv3_searchPreInspectionTasks_(customer) {
  // One exhaustive search path: never silently discard the 11th task or assume
  // a capped search is complete. The ID search throws on ambiguous pagination.
  const customerId = String(tmv3_clean_(customer && customer['Customer ID']));
  if (!/^[1-9]\d*$/.test(customerId)) throw new Error('PI_CUSTOMER_ID_REQUIRED');
  const ids = tmv3_searchPreInspectionTaskIdsForCustomer_(customer);
  const configuredCap=Number((TMV3.OPERATIONS || {}).preInspectionTaskHydrationMaxTasks || 60);
  const cap=Number.isSafeInteger(configuredCap) && configuredCap>0 ? Math.min(100,configuredCap) : 60;
  if (ids.length > cap) throw new Error('PI_HISTORY_HYDRATION_CAP_REVIEW_NO_CREATE');
  const seen = {};
  return ids.map(function(id) {
    const key = String(id);
    if (seen[key]) throw new Error('PI_HISTORY_DUPLICATE_TASK_ID');
    seen[key] = true;
    const task = tmv3_getTaskById_(id);
    if (!task || String(task['Task ID']) !== key ||
        Number(task['Task Type ID']) !== 105 ||
        String(task['Customer ID']) !== customerId) {
      throw new Error('PI_HISTORY_TASK_OWNERSHIP_CONFLICT_NO_CREATE:' + key);
    }
    return task;
  });
}

function tmv3_preInspectionStep1SourceAllowed_(cfg, calCfg, creator) {
  if (calCfg.role === 'PRIMARY_SHARED') return true;

  tmv3_audit_(
    'PreInspection','','','STEP1_SOURCE_SCOPE','SKIP',
    'PreInspection intake is authoritative only from CF Preinspects. Secondary Calendar copies are mirror-only.'
  );
  return false;
}

// Original: 30_Calendar; Function: tmv3_extractLegacyPreInspectionOrderNumber_
function tmv3_extractLegacyPreInspectionOrderNumber_(description) {
  const clean = tmv3_stripPhonesForOrderParsing_(description);
  if (!clean) return '';

  const match = clean.match(
    /\b(?:SO|S\/O|Sales\s*Order|Order)\s*(?:#|No\.?|Number)?\s*[:\-]?\s*(\d{5,8})\b/i
  );

  return match ? match[1] : '';
}

// Original: 30_Calendar; Function: tmv3_preInspectionCalendarCustomerName_
function tmv3_preInspectionCalendarCustomerName_(title, customerNumber) {
  let value = tmv3_clean_(title);
  if (!value) return '';

  value = value.replace(
    /(?:\+?1[\s.\-]?)?\(?[2-9]\d{2}\)?[\s.\-]?[2-9]\d{2}[\s.\-]?\d{4}/g,
    ' '
  );

  if (customerNumber) {
    const digits = String(customerNumber).replace(/[^0-9]/g, '');
    value = value
      .replace(
        new RegExp('(^|\\s)(?:C|Cust(?:omer)?)\\s*#?\\s*' + digits + '(?=\\s|[-–—:]|$)', 'ig'),
        ' '
      )
      .replace(
        new RegExp('(^|\\s)#?' + digits + '(?=\\s|[-–—:]|$)', 'g'),
        ' '
      );
  }

  return value
    .replace(/\bpre\s*-?inspect(?:ion)?\b/ig, ' ')
    .replace(/\s*[-–—:+&]+\s*/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// Original: 30_Calendar; Function: tmv3_preInspectionProvenanceCreator_
function tmv3_preInspectionProvenanceCreator_(vertical,cfg,calCfg,event,fallback) {
  if(vertical!=='PreInspection'||calCfg.calendarId!==cfg.primaryCalendarId)return fallback;
  const raw=PropertiesService.getScriptProperties().getProperty('TMV3_PI_ORIGIN_'+tmv3_hash_(event.getId()));
  if(!raw)return fallback;
  let origin;try{origin=JSON.parse(raw);}catch(e){throw new Error('PREINSPECTION_ORIGIN_METADATA_INVALID');}
  if(origin.eventId!==event.getId()||(cfg.secondaryCalendarIds||[]).indexOf(origin.calendarId)<0)throw new Error('PREINSPECTION_ORIGIN_SCOPE_INVALID');
  const calendar=CalendarApp.getCalendarById(origin.calendarId);
  const source=calendar&&calendar.getEventById(origin.eventId);
  const creators=source&&source.getCreators?source.getCreators().map(function(x){return String(x).toLowerCase();}):[];
  if(!source||creators.length!==1||creators[0]!==origin.creator)throw new Error('PREINSPECTION_ORIGINAL_CREATOR_CHANGED_REVIEW');
  return origin.creator;
}

// Original: 34_Step2_Eligibility; Function: tmv3_step2ClassifyPreInspection_
function tmv3_step2ClassifyPreInspection_(record) {
  if (tmv3_step2PreInspectionNonCustomer_(record)) {
    return tmv3_step2Decision_('SKIP', 'PREINSPECTION_NON_JOB',
      'Non-customer appointment excluded from PreInspection automation.', false);
  }
  const roles = record.sourceCalendarRoles || [
    record.calendarRole || ''
  ];
  const fromShared = roles.indexOf('PRIMARY_SHARED') !== -1;
  const fromStephen =
    roles.indexOf('SECONDARY_STEPHEN') !== -1;

  const warnings =
    tmv3_step2PreInspectionFormatWarnings_(record);

  if (fromShared) {
    return tmv3_step2Decision_(
      'ELIGIBLE',
      warnings.length
        ? 'PREINSPECTION_SHARED_FORMAT_ATTENTION'
        : 'PREINSPECTION_SHARED_ELIGIBLE',
      'CF Preinspects is the authoritative PreInspection source.',
      false,
      warnings,
      'PRIMARY_SOURCE'
    );
  }

  if (fromStephen) {
    return tmv3_step2Decision_(
      'SKIP',
      'PREINSPECTION_SECONDARY_MIRROR_ONLY',
      'Stephen Calendar is mirror-only. PreInspection intake must originate from CF Preinspects.',
      false,
      warnings,
      'SECONDARY_NOT_AUTHORITATIVE'
    );
  }

  return tmv3_step2Decision_(
    'REVIEW',
    'PREINSPECTION_UNKNOWN_SOURCE',
    'PreInspection event is not traceable to the authoritative CF Preinspects source.',
    false,
    warnings,
    'UNKNOWN_SOURCE'
  );
}

// Original: 34_Step2_Eligibility; Function: tmv3_step2PreInspectionNonCustomer_
function tmv3_step2PreInspectionNonCustomer_(record) {
  if (record.isAllDay) return true;

  const title = tmv3_clean_(record.title);

  const patterns = [
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
  ];

  return patterns.some(function(pattern) {
    return pattern.test(title);
  });
}

// Original: 34_Step2_Eligibility; Function: tmv3_step2PreInspectionFormatWarnings_
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

// Original: 35_CalendarLinks; Function: tmv3_preInspectionClickableLinkPresent_
function tmv3_preInspectionClickableLinkPresent_(description, link) {
  const text = String(description || '');
  const url = tmv3_regexEscape_(String(link && link.url || ''));
  if (!url) return false;

  return new RegExp(
    '<a\\b[^>]*href=["\\\']' + url + '["\\\'][^>]*>',
    'i'
  ).test(text);
}

// Original: 35_CalendarLinks; Function: tmv3_preInspectionHtmlLink_
function tmv3_preInspectionHtmlLink_(link) {
  return '<a href="' +
    tmv3_calendarHtmlEscape_(String(link.url || '')) +
    '">' +
    tmv3_calendarHtmlEscape_(String(link.label || '')) +
    '</a>';
}

// Original: 35_CalendarLinks; Function: tmv3_preInspectionEnsureLegacyTitle_
function tmv3_preInspectionEnsureLegacyTitle_(description,legacyTitle) {
  const authored=String(description||'');
  const title=String(legacyTitle||'').trim();
  if(!title)return authored;
  const readable=tmv3_calendarReadableDescription_(authored);
  if(String(readable||'').trim().indexOf(title)===0)return authored;
  const html=/<[a-z][^>]*>/i.test(authored);
  return html ? tmv3_calendarHtmlEscape_(title)+'<br>*<br>'+authored :
    title+(authored?'\n*\n'+authored:'');
}

function tmv3_preInspectionAuthoredMarkup_(description) {
  // Strip ONLY the known automation-managed blocks. Do not rewrite headings,
  // author-created blank lines, notes, paragraphs, photos or HTML markup.
  let raw=tmv3_stripV3ManagedBlock_(String(description||''));
  raw=raw.replace(/<!--\s*PREINSPECT_STRIVEN_TASK_LINK_START\s*-->[\s\S]*?<!--\s*PREINSPECT_STRIVEN_TASK_LINK_END\s*-->/gi,'');
  const offset=tmv3_managedCalendarBlockStart_(raw);
  if(offset>=0){
    raw=raw.slice(0,offset);
    // Remove only the separator immediately before our managed footer.
    raw=raw.replace(/(?:<br\s*\/?>|\r?\n)\s*$/i,'');
  }
  if(/<[a-z][^>]*>/i.test(raw))return raw.replace(/\s+$/,'');
  // Calendar HTML hyperlinks require markup; escape only plain-text contents
  // and turn each existing newline into a single <br>, never collapse paragraphs.
  return tmv3_calendarHtmlEscape_(raw.replace(/\r\n?/g,'\n').replace(/\s+$/,''))
    .replace(/\n/g,'<br>');
}

function tmv3_preInspectionClickableManagedDescription_(existingDescription,plan) {
  const before=String(existingDescription||'');
  const links=(plan&&plan.links)||[];
  const seen={};
  links.forEach(function(link){
    if(!link||!link.key||!/^https:\/\//i.test(String(link.url||''))||seen[link.key])
      throw new Error('PI_MANAGED_LINK_INVALID_OR_DUPLICATE');
    seen[link.key]=true;
  });
  const authored=tmv3_preInspectionAuthoredMarkup_(before);
  const header=tmv3_calendarHtmlEscape_(TMV3_FINAL_LINK_HEADING);
  const footer=header+(links.length?'<br>'+links.map(tmv3_preInspectionHtmlLink_).join('<br>'):'');
  const out=authored+(authored?'<br>':'')+footer;
  tmv3_assertPreInspectionLinkDescription_(out,plan);
  return {description:out,status:out===before?'LINKS_ALREADY_CANONICAL':'LINKS_REQUIRE_UPDATE',changed:out!==before};
}

function tmv3_preInspectionTitleVariants_(plan) {
  plan = plan || {};
  const number = tmv3_clean_(plan.customerNumber);
  const name = tmv3_clean_(plan.customerName);
  const phone = tmv3_clean_(plan.phone);
  const current = tmv3_clean_(plan.currentTitle);
  const variants = [current];

  if (number && name && phone) {
    variants.push(number + ' - ' + name + ' - ' + phone);
    variants.push('C#' + number + ' - ' + name + ' - ' + phone);
    variants.push('Cust#' + number + ' - ' + name + ' - ' + phone);
  }

  return tmv3_unique_(variants.map(tmv3_clean_).filter(Boolean));
}

// Original: 35_CalendarLinks; Function: tmv3_stripPreInspectionPreservedTitle_
function tmv3_stripPreInspectionPreservedTitle_(value, plan) {
  let text = tmv3_decodeCalendarHtmlEntities_(
    String(value || '')
  );

  tmv3_preInspectionTitleVariants_(plan).forEach(function(title) {
    const esc = tmv3_regexEscape_(title);
    text = text.replace(
      new RegExp(
        '^\\s*(?:<p[^>]*>\\s*)?' +
        esc +
        '\\s*(?:<\\/p>)?\\s*(?:(?:<br\\s*\\/?>|\\r?\\n)\\s*){0,2}',
        'i'
      ),
      ''
    );
  });

  return text.replace(/^\s+/, '');
}

// Original: 35_CalendarLinks; Function: tmv3_normalizePreInspectionAuthoredWraps_
function tmv3_normalizePreInspectionAuthoredWraps_(value) {
  let text = tmv3_decodeCalendarHtmlEntities_(
    String(value || '')
  )
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
    .replace(new RegExp(marker, 'g'), '\n\n')
    .replace(
      /(^|\n\n)\s*(?:\*\*)?Sales Order:(?:\*\*)?\s*/ig,
      function(match, prefix) { return prefix || ''; }
    )
    .replace(
      /(^|\n\n)\s*(?:\*\*)?Notes:(?:\*\*)?\s*/ig,
      function(match, prefix) { return prefix || ''; }
    )
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  return text;
}

// Original: 35_CalendarLinks; Function: tmv3_preInspectionCalendarNotesText_
function tmv3_preInspectionCalendarNotesText_(description, plan) {
  let authored = tmv3_stripManagedLinkBlocks_(description || '');

  // The exact previous Calendar title is intentional business history.
  // Preserve it as the first Description line when title normalization
  // placed it there. Only presentation-only headings are removed.
  authored = tmv3_normalizePreInspectionAuthoredWraps_(authored);

  if (!authored) return '';

  authored = authored
    .replace(
      /(^|\n\n)\s*(?:\*\*)?Sales Order:(?:\*\*)?\s*/ig,
      function(match, prefix) { return prefix || ''; }
    )
    .replace(
      /(^|\n\n)\s*(?:\*\*)?Notes:(?:\*\*)?\s*/ig,
      function(match, prefix) { return prefix || ''; }
    );

  return authored.replace(/\n{3,}/g, '\n\n').trim();
}

// Original: 35_CalendarLinks; Function: tmv3_assertPreInspectionLinkDescription_
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
    // Reject raw/Markdown copies outside the single expected anchor as well.
    if (text.replace(anchors[0], '').indexOf(String(link.url || '')) >= 0) {
      throw new Error('PREINSPECTION_DUPLICATE_RAW_LINK: ' + link.key);
    }
  });
  return true;
}

// Original: 36_Step3_Business_Anchor; Function: tmv3_step3ResolvePreInspectionAnchor_
function tmv3_step3ResolvePreInspectionAnchor_(record, refs) {
  const calendarCustomerNumber = tmv3_clean_(record.customerNumber);
  const evidence = [];
  const warnings = [];

  let customer = null;
  let order = null;
  let unresolvedCalendarCustomerNumber = '';

  if (calendarCustomerNumber) {
    customer = refs.customerByNumber[calendarCustomerNumber] || null;

    if (customer) {
      evidence.push('CALENDAR_CUSTOMER_NUMBER_EXACT');
    } else {
      unresolvedCalendarCustomerNumber = calendarCustomerNumber;
      warnings.push(
        'Calendar number ' + calendarCustomerNumber +
        ' did not match a Customer #; it will be tested only as deterministic SO evidence.'
      );
    }
  }

  const orderResolution = tmv3_step3ResolvePreInspectionOrderEvidence_(
    record,
    refs,
    unresolvedCalendarCustomerNumber
  );

  if (orderResolution.ambiguous) {
    return tmv3_step3Decision_(
      'REVIEW',
      'PREINSPECTION_SO_AMBIGUOUS',
      'Multiple distinct Striven orders are supported by the Calendar numeric evidence.',
      customer
        ? tmv3_step3AnchorFromPreInspection_(record, customer, null)
        : null,
      evidence.concat(orderResolution.evidence),
      warnings.concat(orderResolution.warnings)
    );
  }

  order = orderResolution.order;
  Array.prototype.push.apply(evidence, orderResolution.evidence);
  Array.prototype.push.apply(warnings, orderResolution.warnings);

  if (customer && order) {
    const orderCustomerId = tmv3_clean_(order['Customer ID']);
    const customerId = tmv3_clean_(customer['Customer ID']);

    if (orderCustomerId && customerId && orderCustomerId !== customerId) {
      return tmv3_step3Decision_(
        'REVIEW',
        'PREINSPECTION_CUSTOMER_SO_CONFLICT',
        'Calendar Customer # and Calendar SO resolve to different Striven Customers.',
        tmv3_step3AnchorFromPreInspection_(record, customer, order),
        evidence,
        warnings
      );
    }

    if (orderCustomerId && customerId) {
      evidence.push('SO_CORROBORATES_CUSTOMER');
    }
  }

  if (!customer && order) {
    const orderCustomerId = tmv3_clean_(order['Customer ID']);
    customer = orderCustomerId ? (refs.customerById[orderCustomerId] || null) : null;

    if (customer) {
      evidence.push('UNIQUE_SO_RESOLVES_CUSTOMER');
    }
  }

  if (!customer) {
    const identityBundleReady =
      !unresolvedCalendarCustomerNumber &&
      !!tmv3_clean_(record.calendarCustomerName) &&
      !!tmv3_phone10_(record.phone) &&
      !!tmv3_clean_(record.location);

    if (identityBundleReady) {
      return tmv3_step3Decision_(
        'VERIFIED',
        'PREINSPECTION_IDENTITY_EVIDENCE_READY',
        'No deterministic Customer/Order anchor is available, but Calendar name + phone + address are sufficient to proceed to Step 4 identity proof.',
        {
          anchorType: 'IDENTITY_EVIDENCE',
          customerId: '',
          customerNumber: '',
          customerName: tmv3_clean_(record.calendarCustomerName),
          orderId: tmv3_clean_(order && order['Order ID']),
          orderNumber: tmv3_clean_(order && order['Order Number']),
          orderType: tmv3_clean_(order && order['Order Type']),
          orderStatus: tmv3_clean_(order && order['Status']),
          orderUrl: tmv3_clean_(order && order['URL']),
          orderLocationIdEvidence: tmv3_clean_(order && order['Location ID']),
          orderContactIdEvidence: tmv3_clean_(order && order['Contact ID']),
          orderEvidenceOnly: true
        },
        evidence.concat(['CALENDAR_IDENTITY_BUNDLE_READY']),
        warnings
      );
    }

    return tmv3_step3Decision_(
      'REVIEW',
      'PREINSPECTION_BUSINESS_ANCHOR_UNRESOLVED',
      'No exact Customer #, unique SO relationship, or complete identity bundle is available.',
      tmv3_step3AnchorFromPreInspection_(record, null, order),
      evidence,
      warnings
    );
  }

  const anchor = tmv3_step3AnchorFromPreInspection_(record, customer, order);

  // Exact Customer # is already a deterministic business anchor.
  // SO is corroboration/data quality, not a prerequisite for Step 3.
  return tmv3_step3Decision_(
    'VERIFIED',
    order
      ? 'PREINSPECTION_CUSTOMER_SO_VERIFIED'
      : 'PREINSPECTION_CUSTOMER_VERIFIED',
    order
      ? 'PreInspection Customer resolved and the Calendar SO evidence points to the same Customer.'
      : 'PreInspection Customer resolved exactly; SO corroboration is unavailable or unresolved.',
    anchor,
    evidence,
    warnings
  );
}

// Original: 36_Step3_Business_Anchor; Function: tmv3_step3ResolvePreInspectionOrderEvidence_
function tmv3_step3ResolvePreInspectionOrderEvidence_(record, refs, unresolvedCalendarCustomerNumber) {
  const candidateNumbers = [];
  const evidence = [];
  const warnings = [];

  const parsed = tmv3_clean_(record.orderNumber);
  if (parsed) candidateNumbers.push(parsed);

  // A six-digit leading number that failed Customer lookup is a common
  // legacy pattern where the title actually carried the Sales Order.
  if (
    unresolvedCalendarCustomerNumber &&
    /^\d{6}$/.test(unresolvedCalendarCustomerNumber)
  ) {
    candidateNumbers.push(unresolvedCalendarCustomerNumber);
  }

  if (!parsed) {
    tmv3_step3BareSixDigitOrderCandidates_(record).forEach(function(number) {
      candidateNumbers.push(number);
    });
  }

  const uniqueNumbers = tmv3_unique_(
    candidateNumbers.map(tmv3_clean_).filter(Boolean)
  );

  const matchedOrders = [];

  uniqueNumbers.forEach(function(number) {
    const matches = tmv3_step3UniqueOrders_(
      (refs.ordersByNumber[number] || [])
        .filter(function(order) {
          return tmv3_step3OrderMatchesVertical_(
            order,
            'PreInspection'
          );
        })
        .slice()
    );

    if (matches.length === 1) {
      matchedOrders.push(matches[0]);

      if (number === parsed) {
        evidence.push('CALENDAR_SO_NUMBER_EXACT');
      } else if (number === unresolvedCalendarCustomerNumber) {
        evidence.push('LEADING_NUMBER_RECOVERED_AS_SO');
      } else {
        evidence.push('LEGACY_BARE_SIX_DIGIT_SO_EXACT');
      }
    }
  });

  const orders = tmv3_step3UniqueOrders_(matchedOrders);

  if (orders.length > 1) {
    return {
      ambiguous: true,
      order: null,
      evidence: tmv3_unique_(evidence),
      warnings: warnings
    };
  }

  if (orders.length === 1) {
    return {
      ambiguous: false,
      order: orders[0],
      evidence: tmv3_unique_(evidence),
      warnings: warnings
    };
  }

  if (parsed) {
    warnings.push(
      'Calendar SO #' + parsed +
      ' was not found in the current Order source.'
    );
  } else {
    warnings.push('No deterministic SO corroboration was found.');
  }

  return {
    ambiguous: false,
    order: null,
    evidence: tmv3_unique_(evidence),
    warnings: warnings
  };
}

// Original: 36_Step3_Business_Anchor; Function: tmv3_step3AnchorFromPreInspection_
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

// Original: 38_Step4_Identity; Function: tmv3_step4ResolvePreInspectionIdentity_
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
    const recovery = tmv3_step4ResolveCustomerFromCalendarEvidence_(
      record,
      refs
    );

    if (recovery.status !== 'MATCHED') {
      return tmv3_step4Decision_(
        'REVIEW',
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
    Array.prototype.push.apply(evidence, recovery.evidence || []);
  }

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
      'REVIEW',
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

  // Customer Contact is optional identity evidence, NOT the Striven Task
  // Requested By. Requested By is the verified Calendar-owner EMPLOYEE.
  // Missing/ambiguous Contact does not block a pre-inspection solely for that.
  const contactResult = tmv3_step4ResolveContact_(record, customer, '');
  if (contactResult.status === 'REVIEW') {
    warnings.push(contactResult.reason || 'Customer Contact requires review; not used as Requested By.');
  } else if (contactResult.status === 'NO_MATCH') {
    warnings.push('No unique Contact match. Requested By uses the verified Calendar owner employee, not Contact.');
  }

  return tmv3_step4Decision_(
    'VERIFIED',
    'PREINSPECTION_IDENTITY_VERIFIED',
    'PreInspection Customer and Customer-owned Location are verified.',
    customer,
    locationResult.location,
    contactResult.status === 'MATCHED' ? contactResult.contact : null,
    contactResult.status,
    evidence.concat(locationResult.evidence || [], contactResult.evidence || []),
    warnings
  );
}

// Original: 40_Engine; Function: tmv3_preInspectionTaskDecision_
function tmv3_preInspectionTaskDecision_(eventRecord, customer, location, cachedCandidates) {
  // Customer ID alone is NEVER evidence that an OPEN task belongs to this
  // appointment. Strong association requires an explicit linked task or
  // corroborated date, job site and phone/name.
  const cid = String(tmv3_clean_(customer && customer['Customer ID']));
  const lid = String(tmv3_clean_(location && location['Location ID']));
  const linkedId = String(tmv3_clean_(eventRecord && eventRecord.existingTaskId));
  if (!/^[1-9]\d*$/.test(cid)) return {status:'REVIEW',errorCode:'PI_CUSTOMER_UNVERIFIED'};
  let candidates;
  if (Array.isArray(cachedCandidates)) candidates=cachedCandidates.slice();
  else if (linkedId) candidates=[tmv3_getTaskById_(linkedId)];
  else candidates=tmv3_searchPreInspectionTasks_(customer);
  if (linkedId && !candidates.some(function(t){return String(t && t['Task ID'])===linkedId;})) {
    return {status:'REVIEW',errorCode:'PI_EXPLICIT_TASK_LINK_NOT_RESOLVED',
      reason:'Linked Task is absent; no fallback to another Task and no create.'};
  }
  const all=(candidates||[]).filter(function(t){
    return t && Number(t['Task Type ID'])===105 &&
      String(t['Customer ID'])===cid;
  });
  if (linkedId && !all.some(function(t){return String(t['Task ID'])===linkedId;})) {
    return {status:'REVIEW',errorCode:'PI_LINKED_TASK_OWNERSHIP_OR_TYPE_CONFLICT'};
  }
  const evaluated=all.map(function(task){
    const taskId=String(task['Task ID']);
    const sameStart=!!task.Start && tmv3_sameMinute_(eventRecord.start,task.Start);
    const sameDue=!!task.Due && tmv3_sameMinute_(eventRecord.end,task.Due);
    const sameDay=!!task.Start && tmv3_sameLocalDay_(eventRecord.start,task.Start);
    const sameLocation=!!lid && String(task['Location ID'])===lid;
    const taskLocation=String(tmv3_clean_(task['Location ID']));
    const phone=tmv3_phone10_(eventRecord.phone);
    const titlePhoneMatch=!!phone && tmv3_allPhones_(String(task.Name||'')).indexOf(phone)>=0;
    const name=tmv3_norm_(customer.Name);
    const titleNameMatch=!!name && tmv3_norm_(task.Name).indexOf(name)>=0;
    const explicit=!!linkedId && taskId===linkedId;
    const locationConflict=!!lid && !!taskLocation && taskLocation!==lid;
    const corroborated=titlePhoneMatch||titleNameMatch;
    const strong=(!locationConflict) && (
      (explicit && (sameDay || (sameLocation && corroborated))) ||
      (sameStart && sameLocation && corroborated) ||
      (sameDay && sameDue && sameLocation && corroborated)
    );
    return {task:task,open:tmv3_taskIsOpen_(task.Status),
      completed:tmv3_taskIsCompleted_(task.Status),
      sameStart:sameStart,sameDue:sameDue,sameDay:sameDay,
      sameLocation:sameLocation,titlePhoneMatch:titlePhoneMatch,
      titleNameMatch:titleNameMatch,identityCorroborated:corroborated,
      exactCalendarTaskLink:explicit,strongMatch:strong,
      locationConflict:locationConflict};
  });
  const history=evaluated.filter(function(x){return x.completed;}).map(function(x){return x.task;});
  const active=evaluated.filter(function(x){return !x.open && !x.completed;});
  const open=evaluated.filter(function(x){return x.open;});
  if (active.length) return {status:'REVIEW',historyTasks:history,
    errorCode:'PI_NONOPEN_ACTIVE_TASK_REQUIRES_REVIEW',
    reason:'A non-completed, non-open Task exists; lifecycle must be inspected.'};
  if (linkedId && !open.some(function(x){return x.exactCalendarTaskLink;})) {
    return {status:'REVIEW',historyTasks:history,
      errorCode:'PI_EXPLICIT_TASK_LINK_NOT_OPEN_REVIEW',
      reason:'Calendar-linked Task is not open; verify fulfilled history before any create.'};
  }
  if (!open.length) return {status:'CLEAR',task:null,historyTasks:history,
    reason:'No open customer Task; completed Task history remains subject to fresh verification.'};
  if (open.length>1) return {status:'REVIEW',historyTasks:history,
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

// Original: 42_Step5_Task_Resolution; Function: tmv3_step5ResolvePreInspection_
function tmv3_step5ResolvePreInspection_(record, refs) {
  const identity=record.step4||{};
  const cid=String(tmv3_clean_(identity.customer && identity.customer['Customer ID']));
  const cached=((refs.tasksByCustomer||{})[cid]||[]).filter(function(task){
    return Number(task['Task Type ID'])===105;
  });
  const linked=String(tmv3_clean_(record.existingTaskId));
  const evidence=['SHARED_CACHE_ONLY_NOT_FULL_FRESH_HISTORY'];
  if (linked && !(refs.taskById||{})[linked]) {
    return tmv3_step5Decision_('REVIEW','PI_LINKED_TASK_NOT_CACHED_NO_FALLBACK',
      'Calendar carries a Task ID missing from the cache; hydrate exact ID before proceeding.',
      [],0,evidence,['Task ID not certified from fresh Striven.']);
  }
  try {
    const d=tmv3_preInspectionTaskDecision_(record,identity.customer,identity.location,cached);
    if (d.status==='MATCHED')return tmv3_step5Decision_('MATCHED','PREINSPECTION_TASK_MATCHED',
      d.reason,[d.task],(d.historyTasks||[]).length,evidence.concat(d.evidence||[]),[],d.historyTasks||[]);
    if (d.status==='CLEAR')return tmv3_step5Decision_('NO_TASK','PREINSPECTION_NO_OPEN_TASK',
      d.reason,[],(d.historyTasks||[]).length,evidence,[],d.historyTasks||[]);
    return tmv3_step5Decision_('REVIEW',d.errorCode||'PI_TASK_MATCH_UNCERTAIN',
      d.reason||'Task relationship requires review.',[],(d.historyTasks||[]).length,evidence,[],d.historyTasks||[]);
  }catch(error){
    return tmv3_step5Decision_('REVIEW','PI_TASK_RESOLUTION_FAILED',
      String(error && error.message || error),[],0,evidence,[]);
  }
}

function tmv3_step5NormalizePreInspectionTaskTimesAcrossRecords_(records) {
  // Never invent a corrected Striven timestamp by interpreting a 12-hour
  // discrepancy. The exact mismatch must remain visible in the reconciliation.
  return records;
}

function tmv3_step5AlignPreInspectionTaskClock_(task, record) {
  // Evidence-preserving, read-only copy. Do not adjust AM/PM without a verified
  // server response and explicit normalization contract.
  return Object.assign({},task||{});
}

function tmv3_step7HydrateLinkedPreInspectionTasks_(records,refs,runtime) {
  (records||[]).forEach(function(record){
    if (!record || record.vertical!=='PreInspection') return;
    const step4=record.step4||{};
    const cid=String(tmv3_clean_(step4.customer && step4.customer['Customer ID']));
    const id=String(tmv3_clean_(record.existingTaskId));
    if (step4.disposition!=='VERIFIED'||!cid||!id||(refs.taskById||{})[id])return;
    if(!/^[1-9]\d*$/.test(id))throw new Error('PI_LINKED_TASK_ID_INVALID_NO_CREATE');
    const fresh=tmv3_step7FreshTask_(id,runtime);
    if(!fresh.ok||!fresh.task||String(fresh.task['Task ID'])!==id||
        Number(fresh.task['Task Type ID'])!==105||
        String(fresh.task['Customer ID'])!==cid) {
      throw new Error('PI_LINKED_TASK_FRESH_OWNERSHIP_UNVERIFIED_NO_CREATE:'+id);
    }
    refs.tasksByCustomer=refs.tasksByCustomer||{};
    refs.tasksByOrder=refs.tasksByOrder||{};
    refs.taskById[id]=fresh.task;
    refs.tasks=(refs.tasks||[]).filter(function(t){return String(t['Task ID'])!==id;});
    refs.tasks.push(fresh.task);
    Object.keys(refs.tasksByCustomer||{}).forEach(function(k){
      refs.tasksByCustomer[k]=refs.tasksByCustomer[k].filter(function(t){return String(t['Task ID'])!==id;});
    });
    if(!refs.tasksByCustomer[cid])refs.tasksByCustomer[cid]=[];
    refs.tasksByCustomer[cid].push(fresh.task);
    Object.keys(refs.tasksByOrder||{}).forEach(function(k){
      refs.tasksByOrder[k]=refs.tasksByOrder[k].filter(function(t){return String(t['Task ID'])!==id;});
    });
    const orderId=String(tmv3_clean_(fresh.task['Order ID']));
    if(orderId){if(!refs.tasksByOrder[orderId])refs.tasksByOrder[orderId]=[];
      refs.tasksByOrder[orderId].push(fresh.task);}
    tmv3_upsertTaskCacheRow_(fresh.task);
  });
}

function tmv3_preInspectionAssignmentOverride_(record) {
  if(!record || !record.eventId) throw new Error('PI_EVENT_ID_REQUIRED_FOR_ASSIGNMENT');
  // An operator override never bypasses owner identity review. Overrides may
  // only reaffirm the exact verified owner mapping; otherwise block writes.
  const owner=tmv3_piVerifiedOwnerEmployee_(record);
  if(owner.status!=='MATCHED')throw new Error(owner.blocker);
  const key='TMV3_PI_ASSIGNMENT_OVERRIDE_'+tmv3_hash_(record.eventId);
  const raw=PropertiesService.getScriptProperties().getProperty(key);
  if(!raw)return [owner.employeeId];
  let saved;try{saved=JSON.parse(raw);}catch(error){throw new Error('PI_ASSIGNMENT_OVERRIDE_INVALID');}
  if(String(saved.eventId)!==String(record.eventId)||
     String(saved.customerNumber)!==String(record.customerNumber)||
     String(saved.organizerCalendarId)!==String(record.organizerCalendarId)||
     Number(saved.employeeId)!==owner.employeeId)
    throw new Error('PI_OVERRIDE_CONFLICTS_WITH_VERIFIED_OWNER_NO_WRITE');
  return [owner.employeeId];
}

function tmv3_configurePreInspectionAssignmentOverride(eventId,customerNumber,organizerCalendarId,employeeId) {
  // Preserve the existing public function but refuse unverified arbitrary IDs.
  if(!/^[a-z0-9_]+@google\.com$/i.test(String(eventId||'')))throw new Error('PI_EVENT_ID_INVALID');
  const record=tmv3_findFreshEventRecord_('PreInspection',eventId);
  if(!record||String(record.customerNumber)!==String(customerNumber)||
      String(record.organizerCalendarId)!==String(organizerCalendarId))
    throw new Error('PI_ASSIGNMENT_SOURCE_EVIDENCE_CHANGED');
  const approved=tmv3_piVerifiedOwnerEmployee_(record);
  if(approved.status!=='MATCHED'||Number(employeeId)!==approved.employeeId)
    throw new Error('PI_OPERATOR_OVERRIDE_MUST_MATCH_VERIFIED_OWNER');
  const value={eventId:String(eventId),customerNumber:String(customerNumber),
    organizerCalendarId:String(organizerCalendarId),employeeId:approved.employeeId};
  const key='TMV3_PI_ASSIGNMENT_OVERRIDE_'+tmv3_hash_(eventId);
  PropertiesService.getScriptProperties().setProperty(key,JSON.stringify(value));
  return {status:'VERIFIED_OWNER_OVERRIDE_SAVED',employeeId:approved.employeeId,
    businessWrites:0};
}

function tmv3_preInspectionApprovalHold_(eventId) {
  const normalized=String(eventId||'').trim().toLowerCase().replace(/@google[.]com$/,'');
  const entries=TMV3_HARD_RULES.PreInspection.operatorApprovedIntakeHolds||[];
  const matches=entries.filter(function(x) {
    return String(x.eventId||'').toLowerCase().replace(/@google[.]com$/,'')===normalized;
  });
  if(matches.length>1)throw new Error('DUPLICATE_PI_OPERATOR_APPROVAL_NO_WRITE');
  return matches.length?matches[0]:null;
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

function tmv3_ensurePreInspectionMirror_(record,scope) {
  // Disallow calendar guest mutations in unattended automation. Invites
  // propagate notifications and must be a separately authorized operator act.
  return {status:'PI_GUEST_CHANGE_OPERATOR_APPROVAL_REQUIRED',
    eventId:record&&record.eventId||'',writePerformed:false,
    calendarWrites:0,guestInvites:0};
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
      const clean = tmv3_normEmail_(
        guest && guest.getEmail ? guest.getEmail() : ''
      );
      if (clean) participants[clean] = true;
    });
  } catch (ignored) {}

  const stephen = tmv3_normEmail_(
    (TMV3.VERTICALS.PreInspection || {}).secondaryOwnerEmail ||
    'stephen@classicfireplace.ca'
  );

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

// Original: 58_Operations; Function: tmv3_preInspectionCalendarParityEqual_
function tmv3_preInspectionCalendarParityEqual_(expected, actual) {
  expected = expected || {};
  actual = actual || {};
  return (
    expected.title === actual.title &&
    expected.description === actual.description &&
    expected.location === actual.location &&
    expected.start === actual.start &&
    expected.end === actual.end &&
    expected.allDay === actual.allDay &&
    expected.requiredStephenPresent === actual.requiredStephenPresent
  );
}

// Original: 58_Operations; Function: tmv3_reconcilePreInspectionCalendarParity_
function tmv3_reconcilePreInspectionCalendarParity_(record,scope) {
  // Audit copies but never blindly synchronize, rename, reschedule, or add
  // attendees to the representative's private appointment.
  if(!record||record.vertical!=='PreInspection')return {status:'NOT_NEEDED',copies:[],writes:0};
  const cfg=TMV3.VERTICALS.PreInspection||{};
  const primaryId=String(cfg.primaryCalendarId||'');
  if(!primaryId)throw new Error('PI_PRIMARY_CALENDAR_NOT_CONFIGURED');
  const master=tmv3_findEventCopyRobust_(primaryId,record.eventId);
  if(!master||!master.event)throw new Error('PI_AUTHORITY_EVENT_MISSING');
  const authority=tmv3_preInspectionCalendarBusinessSnapshot_(master.event);
  const ids=tmv3_unique_(tmv3_requiredCalendarCopiesForEvent_(record)||[]);
  const differences=[];
  ids.forEach(function(calendarId){
    if(String(calendarId)===primaryId)return;
    const found=tmv3_findEventCopyRobust_(calendarId,record.eventId);
    if(!found||!found.event){differences.push({calendarId:calendarId,issue:'COPY_NOT_ACCESSIBLE_OR_MISSING'});return;}
    const state=tmv3_preInspectionCalendarBusinessSnapshot_(found.event);
    const fields=['title','description','location','start','end','allDay'];
    const changed=fields.filter(function(key){return authority[key]!==state[key];});
    if(changed.length)differences.push({calendarId:calendarId,issue:'PARITY_MISMATCH',fields:changed});
  });
  return {status:differences.length?'PI_PARITY_REVIEW_REQUIRED':'PI_PARITY_READ_ONLY_MATCH',
    authoritativeCalendarId:primaryId,eventId:record.eventId,copies:ids.length,
    differences:differences,writes:0,guestInvites:0};
}

function tmv3_enforcePreInspectionCreatePayloadHardRules_(payload) {
  payload=payload||{};
  Object.keys(payload).forEach(function(k){
    if(/^(?:SalesOrder|SalesOrderId|SalesOrderID|SOId|OrderId|OrderID|InfoCustomFields|CustomFields|CustomFieldValues|TaskFields|InfoFields|Fields)$/i.test(k))
      delete payload[k];
  });
  payload.Description='';
  return payload;
}

function tmv3_assertPreInspectionCreatePayloadHardRules_(payload) {
  const p=payload||{},type=p.Type||{};
  if(Number(type.Id||type.id)!==105)throw new Error('PI_CREATE_TYPE_MUST_BE_105');
  if(Object.keys(p).some(function(k){
    return /^(?:SalesOrder|SalesOrderId|SalesOrderID|SOId|OrderId|OrderID|InfoCustomFields|CustomFields|CustomFieldValues|TaskFields|InfoFields|Fields)$/i.test(k);
  }))throw new Error('PI_CREATE_FORBIDDEN_ORDER_OR_CUSTOM_FIELDS');
  if(p.Description!=='' )throw new Error('PI_CREATE_DESCRIPTION_NOT_BLANK');
  // Requested By must be a real Striven Employee, not a Customer Contact.
  // Organizer ownership is enforced at the canonical plan boundary.
  tmv3_piAssertRequestedByEmployeeShape_(p.RequestedBy);
  return true;
}

function tmv3_preInspectionGuestEmails_(eventRecord) {
  return tmv3_unique_(
    String(eventRecord && eventRecord.guests || '')
      .split(',')
      .map(function(value) { return tmv3_norm_(value); })
      .filter(Boolean)
  );
}

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

  // Accept labelled SO / Sales Order / Order references, including compact
  // slash notation such as SO#5853/67/80/85. We only care that the event
  // carries at least one Sales Order reference; no specific formatting is
  // required.
  const clean = tmv3_stripPhonesForOrderParsing_(haystack)
    .replace(/https?:\/\/\S+/gi, ' ')
    .replace(/(?:Customer|Cust|Task)\s*(?:#|Number|No\.?)?\s*[:\-]?\s*\d+/gi, ' ')
    .replace(/^\s*#?\d{4,8}\s*[-–—]/, ' ');
  return /\b\d{6}\b/.test(clean) || /\b(?:SO|S\/O|Sales\s*Order|Order)\s*(?:#|No\.?|Number)?\s*[:\-]?\s*\d{4,8}(?:\s*\/\s*\d{2,8})*/i.test(
    String(haystack || '')
  );
}

// Original: 58_Operations; Function: tmv3_preInspectionCreatorEmail_
function tmv3_preInspectionCreatorEmail_(eventRecord) {
  const sources = []
    .concat(eventRecord && eventRecord.sourceCreators || [])
    .concat(String(eventRecord && eventRecord.creator || '').split(','))
    .map(function(value) { return tmv3_clean_(value).toLowerCase(); })
    .filter(Boolean);

  const company = sources.filter(function(email) {
    return /@classicfireplace\.ca$/i.test(email);
  });

  const unique = tmv3_unique_(company);
  return unique.length === 1 ? unique[0] : ''; // Missing/ambiguous creator: no send.
}

// Original: 58_Operations; Function: tmv3_preInspectionCreatorFirstName_
function tmv3_preInspectionCreatorFirstName_(email) {
  const local = tmv3_clean_(email).split('@')[0] || '';
  if (!local) return 'there';

  const token = local
    .split(/[._\-]+/)
    .filter(Boolean)[0] || '';

  return token
    ? token.charAt(0).toUpperCase() + token.slice(1).toLowerCase()
    : 'there';
}

// Original: 58_Operations; Function: tmv3_preInspectionNotificationCustomerNumber_
function tmv3_preInspectionNotificationCustomerNumber_(record,customers) {
  const title=tmv3_clean_(record && record.title);
  const stripped=tmv3_stripPhonesForOrderParsing_(title)
    .replace(/https?:\/\/\S+/gi,' ')
    .replace(/\b(?:SO|S\/O|Sales\s*Order|Order)\s*(?:#|No\.?|Number)?\s*[:\-]?\s*\d{4,8}(?:\s*\/\s*\d{2,8})*/ig,' ');
  const numbers=stripped.match(/\b\d{4,8}\b/g)||[];
  const phone=tmv3_phone10_(record && record.phone || tmv3_extractPhone_(title));
  const words=title.toLowerCase().replace(/[^a-z0-9]+/g,' ').split(' ').filter(Boolean);
  const matches=(customers||[]).filter(function(customer){
    const number=tmv3_clean_(customer['Customer Number']);
    if(numbers.indexOf(number)<0) return false;
    const nameWords=tmv3_clean_(customer.Name).toLowerCase().replace(/[^a-z0-9]+/g,' ').split(' ').filter(Boolean);
    const nameMatches=nameWords.length>=2 && nameWords.every(function(word){return words.indexOf(word)>=0;});
    const profilePhone=tmv3_phone10_(customer['Primary Phone']);
    return (phone && profilePhone && phone===profilePhone) || nameMatches;
  });
  if(matches.length!==1) return '';
  const number=tmv3_clean_(matches[0]['Customer Number']);
  if((customers||[]).filter(function(c){return tmv3_clean_(c['Customer Number'])===number;}).length!==1) return '';
  return number;
}

// Original: 58_Operations; Function: tmv3_preInspectionMissingDetailAssessment_
function tmv3_preInspectionMissingDetailAssessment_(eventRecord) {
  const record = eventRecord || {};
  const text = [record.title || '', record.identityDescriptionClean ||
    record.descriptionClean || record.rawDescription || ''].join(' ');
  const customerNumber = tmv3_extractCustomerNumberForVertical_('PreInspection', text) ||
    tmv3_preInspectionNotificationCustomerNumber_(record,tmv3_rows_(TMV3.SHEETS.CUSTOMERS));
  const missing = [];
  if (!customerNumber) missing.push({code:'CUSTOMER_NUMBER',field:'Title',
    message:'Customer Number missing'});
  // The Sales Order is optional. Never email the creator for its absence.
  if (!tmv3_clean_(record.location)) missing.push({code:'LOCATION',field:'Location',
    message:'Location missing — please add the full job-site address'});
  return {eventId:record.eventId || '', title:tmv3_clean_(record.title),
    creatorEmail:tmv3_preInspectionCreatorEmail_(record), missing:missing,
    shouldNotify:missing.length > 0};
}

// Original: 58_Operations; Function: tmv3_preInspectionHtmlEscape_
function tmv3_preInspectionHtmlEscape_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// Original: 58_Operations; Function: tmv3_preInspectionDisplayGuests_
function tmv3_preInspectionDisplayGuests_(eventRecord) {
  const guests = tmv3_preInspectionGuestEmails_(eventRecord);
  if (!guests.length) return 'No guests added';

  return guests.map(function(email) {
    if (email === tmv3_norm_(TMV3.VERTICALS.PreInspection.secondaryOwnerEmail)) {
      return 'Stephen Foley';
    }
    if (email === tmv3_norm_(TMV3.VERTICALS.PreInspection.primaryCalendarId)) {
      return 'CF Preinspects';
    }
    return email;
  }).join('<br>');
}

// Original: 58_Operations; Function: tmv3_preInspectionMissingFields_
function tmv3_preInspectionMissingFields_(assessment) {
  const fields = {};
  (assessment.missing || []).forEach(function(item) {
    fields[item.field] = true;
  });
  return fields;
}

// Original: 58_Operations; Function: tmv3_preInspectionComparisonRowsHtml_
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

// Original: 58_Operations; Function: tmv3_preInspectionNotificationCc_
function tmv3_preInspectionNotificationCc_() {
  return [
    'jay@classicfireplace.ca',
    'thang@classicfireplace.ca'
  ];
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

// Original: 58_Operations; Function: tmv3_preInspectionMissingDetailPreviewFromRecords_
function tmv3_preInspectionMissingDetailPreviewFromRecords_(records) {
  const preInspection = (records || []).filter(function(record) {
    return record && record.vertical === 'PreInspection';
  });

  const previews = preInspection
    .map(tmv3_preInspectionMissingDetailEmailPreview_)
    .filter(Boolean);

  return {
    status:'PREVIEW_ONLY_NO_EMAILS',
    checked:preInspection.length,
    wouldNotify:previews.length,
    previews:previews
  };
}

// Original: 58_Operations; Function: tmv3_previewPreInspectionMissingDetailEmails
function tmv3_previewPreInspectionMissingDetailEmails() {
  const result = tmv3_preInspectionMissingDetailPreviewFromRecords_(
    tmv3_step1CalendarRecords_()
  );

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// Original: 58_Operations; Function: tmv3_preInspectionCustomerLocationsFresh_
function tmv3_preInspectionCustomerLocationsFresh_(customerId) {
  const id=Number(customerId);
  if (!Number.isInteger(id) || id<=0) throw new Error('LOCATION_CUSTOMER_ID_INVALID');
  const response=tmv3_fetchJson_(TMV3.API_BASE+'/v1/customers/'+encodeURIComponent(id)+'/locations',{method:'get'});
  const rawData=response && (response.Data!==undefined?response.Data:response.data);
  const count=response && (response.TotalCount!==undefined?response.TotalCount:response.totalCount);
  // Live Striven represents an explicitly empty collection as count=0, data=null.
  const data=count===0 && rawData===null ? [] : rawData;
  if (!response || !Array.isArray(data) ||
      !Number.isInteger(count) || count!==data.length) {
    throw new Error('LOCATION_RESPONSE_INCOMPLETE_NO_CREATE');
  }
  const seen={};
  return data.map(function(raw) {
    const customer=raw && (raw.Customer || raw.customer);
    const address=raw && (raw.Address || raw.address);
    const row={Id:raw && (raw.Id!==undefined?raw.Id:raw.id),
      Customer:customer && {Id:customer.Id!==undefined?customer.Id:customer.id},
      Address:address && {Address1:address.Address1 || address.address1,FullAddress:address.FullAddress || address.fullAddress,
        Address2:address.Address2 || address.address2,City:address.City || address.city,
        State:address.State || address.state,PostalCode:address.PostalCode || address.postalCode}};
    if (!raw || !Number.isInteger(row.Id) || row.Id<=0 || seen[row.Id] ||
        !row.Customer || Number(row.Customer.Id)!==id || !row.Address ||
        !tmv3_clean_(row.Address.Address1 || row.Address.FullAddress)) {
      throw new Error('LOCATION_RESPONSE_IDENTITY_INVALID_NO_CREATE');
    }
    seen[row.Id]=true;
    const a=row.Address;
    return tmv3_normalizedLocationObject_(tmv3_normalizeLocation_({
      Id:row.Id,CustomerId:id,Address1:a.Address1 || a.FullAddress,Address2:a.Address2 || '',
      City:a.City || '',State:a.State || '',PostalCode:a.PostalCode || ''
    },{}));
  });
}

// Original: 58_Operations; Function: tmv3_preInspectionRejectedAddressGuard_
function tmv3_preInspectionRejectedAddressGuard_(guard,preflight) {
  if(!guard || guard.state!=='UNCERTAIN' || guard.locationId ||
      !preflight || !Array.isArray(preflight.owned) || !Array.isArray(preflight.matches) || preflight.matches.length!==0) return false;
  const error=String(guard.error||'');
  if(!/HTTP 400\b/.test(error)) return false;
  const oldAddress=!guard.addressSchema && /Address object is required/i.test(error);
  const oldCountry=guard.addressSchema==='nested-address-v1' &&
    /country_Invalid/.test(error) && /country code Canada could not be read/i.test(error);
  return oldAddress || oldCountry;
}

// Original: 58_Operations; Function: tmv3_preInspectionOptionalOrderLocationEvidence_
function tmv3_preInspectionOptionalOrderLocationEvidence_(record) {
  if(!TMV3_PI_CLEANED_RELEASE_APPROVED)return {status:'PI_STAGING_OPERATOR_PATH_DISABLED',writes:0,emailsSent:0,guestInvites:0};
  return tmv3_piOriginal_preInspectionOptionalOrderLocationEvidence_(record);
}

function tmv3_piOriginal_preInspectionOptionalOrderLocationEvidence_(record) {
  if (!record || record.vertical!=='PreInspection' || !/^\d+$/.test(String(record.orderNumber||''))) return {status:'NOT_REQUIRED'};
  const s2=tmv3_step2CalendarRecords_([record]);
  const s3=tmv3_step3BusinessAnchorRecords_(s2,tmv3_step3AnchorIndex_());
  const s4=tmv3_step4IdentityRecords_(s3,tmv3_step4IdentityIndex_());
  if (s4.length!==1 || !s4[0].step4 || s4[0].step4.code!=='AMBIGUOUS_STRONG_LOCATION') return {status:'NOT_REQUIRED'};
  const customerId=Number(s4[0].step4.customer && s4[0].step4.customer['Customer ID']);
  if (!Number.isInteger(customerId) || customerId<=0) return {status:'CUSTOMER_UNRESOLVED'};
  const number=String(record.orderNumber);
  const found=tmv3_fetchJson_(TMV3.API_BASE+'/v1/sales-orders/search',{method:'post',contentType:'application/json',
    payload:JSON.stringify({PageIndex:0,PageSize:2,Number:number,CustomerId:customerId})});
  const data=found && (found.Data || found.data);
  const count=found && (found.TotalCount!==undefined?found.TotalCount:found.totalCount);
  if (!found || !Array.isArray(data) || count!==data.length || data.length!==1) return {status:'ORDER_NOT_UNIQUE'};
  const rawCandidate=data[0];
  const candidate={Id:tmv3_first_(rawCandidate,['Id']),Number:tmv3_first_(rawCandidate,['Number']),
    Customer:{Id:tmv3_first_(tmv3_first_(rawCandidate,['Customer']),['Id'])}};
  if (String(candidate.Number)!==number || !candidate.Customer || Number(candidate.Customer.Id)!==customerId || !Number.isInteger(candidate.Id) || candidate.Id<=0) return {status:'ORDER_IDENTITY_CONFLICT'};
  const rawOrder=tmv3_fetchJson_(TMV3.API_BASE+'/v1/sales-orders/'+candidate.Id,{method:'get'});
  const order={Id:tmv3_first_(rawOrder,['Id']),OrderNumber:tmv3_first_(rawOrder,['OrderNumber']),OrderName:tmv3_first_(rawOrder,['OrderName']),
    Customer:{Id:tmv3_first_(tmv3_first_(rawOrder,['Customer']),['Id'])},
    ShipToLocation:{Id:tmv3_first_(tmv3_first_(rawOrder,['ShipToLocation']),['Id'])},
    Contact:{Id:tmv3_first_(tmv3_first_(rawOrder,['Contact']),['Id'])},
    Status:{Name:tmv3_first_(tmv3_first_(rawOrder,['Status']),['Name'])}};
  if (!order || order.Id!==candidate.Id || String(order.OrderNumber)!==number || !order.Customer || Number(order.Customer.Id)!==customerId) return {status:'ORDER_READBACK_CONFLICT'};
  const locationId=Number(order.ShipToLocation && order.ShipToLocation.Id);
  const owned=tmv3_preInspectionCustomerLocationsFresh_(customerId);
  const matches=owned.filter(function(location){return Number(location['Location ID'])===locationId &&
    tmv3_addressStrongMatch_(tmv3_addressParts_(record.location),tmv3_addressParts_(tmv3_locationFullAddress_(location)));});
  if (matches.length!==1) return {status:'ORDER_LOCATION_CONFLICT'};
  const row=tmv3_normalizeOrder_({Id:order.Id,OrderNumber:number,CustomerId:customerId,LocationId:locationId,
    ContactId:order.Contact&&order.Contact.Id,Status:order.Status&&order.Status.Name,Name:order.OrderName},'SALES_ORDER');
  const sheet=tmv3_sheet_(TMV3.SHEETS.ORDERS),values=sheet.getDataRange().getValues();
  if (!values.length || values[0][0]!=='Order ID' || values[0][1]!=='Order Number') throw new Error('ORDER_CACHE_SCHEMA_NO_WRITE');
  const existing=[];
  values.slice(1).forEach(function(value,index){if(String(value[1])===number && value[7]==='SALES_ORDER')existing.push({value:value,row:index+2});});
  if (existing.length>1 || existing.some(function(item){return Number(item.value[0])!==order.Id || Number(item.value[2])!==customerId;})) return {status:'ORDER_CACHE_CONFLICT'};
  const cachedLocations=tmv3_rows_(TMV3.SHEETS.LOCATIONS).filter(function(item){return Number(item['Location ID'])===locationId;});
  if(cachedLocations.length>1 || cachedLocations.some(function(item){return Number(item['Customer ID'])!==customerId;})) return {status:'LOCATION_CACHE_CONFLICT'};
  tmv3_upsertLocationCacheObject_(matches[0]);
  sheet.getRange(existing.length?existing[0].row:values.length+1,1,1,row.length).setValues([row]);
  tmv3_audit_('PreInspection',record.eventId,'','OPTIONAL_ORDER_LOCATION_EVIDENCE','PASS',
    JSON.stringify({orderId:order.Id,customerId:customerId,locationId:locationId,orderAttachedToTask:false}));
  return {status:'CACHE_REPAIRED',orderId:order.Id,locationId:locationId};
}

// Original: 58_Operations; Function: tmv3_preInspectionCanonicalCacheTask_
function tmv3_preInspectionCanonicalCacheTask_(task,canonical) {
  if(Number(task&&task['Task Type ID'])!==105) return task;
  if(!canonical || Number(canonical.taskId)!==Number(task['Task ID']) ||
      !canonical.startDateTime || !canonical.dueDateTime ||
      !Number.isFinite(new Date(canonical.startDateTime).getTime()) ||
      !Number.isFinite(new Date(canonical.dueDateTime).getTime()) ||
      new Date(canonical.dueDateTime).getTime()<new Date(canonical.startDateTime).getTime()) throw new Error('CANONICAL_PI_CACHE_SCHEDULE_UNVERIFIED_NO_CACHE_WRITE');
  const verified=Object.assign({},task,{'Start':canonical.startDateTime,'Due':canonical.dueDateTime});
  verified.Fingerprint=tmv3_hash_(['Task ID','Task Number','Task Type ID','Task Type','Status','Name','Customer ID','Location ID','Contact ID','Order ID','Start','Due','Assignees','Pools'].map(function(key){return verified[key]||'';}).join('|'));
  return verified;
}

// Original: 60_Automation; Function: tmv3_preInspectionGuestSyncDecision_
function tmv3_preInspectionGuestSyncDecision_(guestEmails,stephenEmail,ownerCalendarId) {
  const owner=String(ownerCalendarId||'').trim();
  if(!owner)return {status:'OWNER_NOT_VERIFIED',action:'NONE'};
  return {status:'REVIEW_ONLY_NO_UNATTENDED_GUEST_CHANGES',action:'NONE'};
}

function tmv3_reconcilePreInspectionSharedGuest_AUTO_() { return {status:'RETIRED_UNCONDITIONAL_GUEST_SYNC',writes:0,emailsSent:0,guestInvites:0}; }

function tmv3_preInspectionExplicitCandidate_(record) {
  return !!record && tmv3_preInspectionClassifyLikely_(record,TMV3.VERTICALS.PreInspection).status==='EXPLICIT_PREINSPECTION';
}

// Original: 60_Automation; Function: tmv3_preInspectionDiscoveryDecision_
function tmv3_preInspectionDiscoveryDecision_(record,cfg) {
  if (!record || tmv3_step2PreInspectionNonCustomer_(record) || record.status==='cancelled') return {status:'IGNORE',reason:'NON_CUSTOMER_OR_CANCELLED'};
  const creator=tmv3_preInspectionCreatorEmail_(record);
  const excluded=(cfg.secondaryIgnoredCreatorEmails || []).map(tmv3_normEmail_);
  if (creator && excluded.indexOf(creator)>=0) return {status:'IGNORE',reason:'EXCLUDED_CREATOR'};
  const text=[record.title || '',record.descriptionClean || record.rawDescription || ''].join(' ');
  if (/\b(?:not|no)\s+(?:a\s+)?pre[\s-]*inspect(?:ion)?\b/i.test(text)) return {status:'IGNORE',reason:'EXPLICIT_NOT_PREINSPECTION'};
  if (/\bpre[\s-]*inspect(?:ion)?\s+(?:(?:is|was)\s+)?(?:not\b|required later|already completed)/i.test(text)) return {status:'IGNORE',reason:'EXPLICIT_NOT_PREINSPECTION'};
  const instruction=/\bpre[\s-]*inspect(?:ion)?\b/i.test(record.title || '') ||
    /^\s*(?:(?:this is (?:a )?|appointment(?: type)?[: -]+))?pre[\s-]*inspect(?:ion)?\b/i.test(record.descriptionClean || record.rawDescription || '');
  const hasPhone=!!tmv3_phone10_(record.phone), hasLocation=!!tmv3_clean_(record.location);
  if (!creator) return instruction || hasPhone || hasLocation ? {status:'REVIEW_VISIT_TYPE',reason:'CREATOR_UNVERIFIED'} : {status:'IGNORE',reason:'NO_CUSTOMER_VISIT_EVIDENCE'};
  if (instruction) return {status:'EXPLICIT_PREINSPECTION',reason:'EXPLICIT_SOURCE_INSTRUCTION'};
  if (!hasPhone || !hasLocation) return hasPhone || hasLocation ? {status:'REVIEW_VISIT_TYPE',reason:'INSUFFICIENT_CUSTOMER_VISIT_EVIDENCE'} : {status:'IGNORE',reason:'INSUFFICIENT_CUSTOMER_VISIT_EVIDENCE'};
  return {status:'REVIEW_VISIT_TYPE',reason:/\b(?:return|go back|measure|filler|trim)\b/i.test(text)?'RETURN_OR_MEASURING_VISIT_UNCONFIRMED':'CUSTOMER_VISIT_TYPE_UNCONFIRMED'};
}

// Original: 60_Automation; Function: tmv3_preInspectionClassifyLikely_
function tmv3_preInspectionClassifyLikely_(record,cfg,fresh) {
  const base=tmv3_preInspectionDiscoveryDecision_(record,cfg);
  if(base.status!=='REVIEW_VISIT_TYPE' || base.reason==='CREATOR_UNVERIFIED' || base.reason==='INSUFFICIENT_CUSTOMER_VISIT_EVIDENCE')return base;
  if(/\b(?:service|delivery|repair|maintenance|troubleshoot(?:ing)?)\b/i.test(record.title || '') ||
      /^\s*(?:service|delivery|repair|maintenance)\b/i.test(record.descriptionClean || record.rawDescription || ''))return {status:'REVIEW_VISIT_TYPE',reason:'OTHER_VERTICAL_INTENT_REQUIRES_REVIEW'};
  const key=record.eventId+'|'+record.fingerprint;
  if(!fresh && TMV3_PI_LIKELY_DECISIONS[key])return TMV3_PI_LIKELY_DECISIONS[key];
  let result=base;
  try {
    const refs=TMV3_PI_LIKELY_REFS || (TMV3_PI_LIKELY_REFS=tmv3_step4IdentityIndex_());
    const identity=tmv3_step4ResolveCustomerFromCalendarEvidence_(record,refs);
    const customerId=Number(identity.customer && identity.customer['Customer ID']);
    if(!Number.isSafeInteger(customerId)||customerId<=0)return {status:'REVIEW_VISIT_TYPE',reason:'CUSTOMER_IDENTITY_UNRESOLVED'};
    const explicit=String(record.orderNumber || '');
    // Never silently choose the first of multiple authored SO references.
    const text=[record.title,record.descriptionClean || record.rawDescription || ''].join(' ');
    const mentions=tmv3_unique_((text.match(/\b(?:SO|Sales\s+Order)\s*#?\s*\d{5,}/ig)||[]).map(function(v){return v.match(/\d+/)[0];}));
    if(mentions.length>1)return {status:'REVIEW_VISIT_TYPE',reason:'MULTIPLE_ORDER_REFERENCES'};
    const query={PageIndex:0,PageSize:10,CustomerId:customerId};
    if(explicit)query.Number=explicit;
    const search=tmv3_fetchJson_(TMV3.API_BASE+'/v1/sales-orders/search',{method:'post',contentType:'application/json',payload:JSON.stringify(query)});
    const rows=tmv3_first_(search,['Data']),total=Number(tmv3_first_(search,['TotalCount']));
    if(!Array.isArray(rows)||!Number.isSafeInteger(total)||total!==rows.length||total>10 || (explicit && rows.length!==1))return {status:'REVIEW_VISIT_TYPE',reason:'ORDER_SEARCH_INCOMPLETE_OR_AMBIGUOUS'};
    const locations=tmv3_preInspectionCustomerLocationsFresh_(customerId);
    const sites=locations.filter(function(l){return Number(l['Customer ID'])===customerId && tmv3_addressStrongMatch_(tmv3_addressParts_(record.location),tmv3_addressParts_(tmv3_locationFullAddress_(l)));});
    if(sites.length!==1)return {status:'REVIEW_VISIT_TYPE',reason:'JOB_SITE_NOT_UNIQUE_FOR_CUSTOMER'};
    const relevant=[];
    rows.forEach(function(row){
      const id=Number(tmv3_first_(row,['Id'])),number=String(tmv3_first_(row,['Number']));
      if(!Number.isSafeInteger(id)||id<=0||(explicit&&number!==explicit))throw Error('ORDER_SEARCH_IDENTITY_CONFLICT');
      const raw=tmv3_fetchJson_(TMV3.API_BASE+'/v1/sales-orders/'+id,{method:'get'});
      const customer=tmv3_first_(raw,['Customer']),location=tmv3_first_(raw,['ShipToLocation']),status=tmv3_first_(raw,['Status']);
      if(Number(tmv3_first_(raw,['Id']))!==id||String(tmv3_first_(raw,['OrderNumber','Number']))!==number||Number(tmv3_first_(customer,['Id']))!==customerId)throw Error('ORDER_READBACK_IDENTITY_CONFLICT');
      const statusId=String(tmv3_first_(status,['Id'])),statusName=tmv3_norm_(tmv3_first_(status,['Name']));
      const known={'19':'quoted','22':'approved','25':'in progress'};
      if(!known[statusId]||known[statusId]!==statusName)return;
      if(Number(tmv3_first_(location,['Id']))!==Number(sites[0]['Location ID']))return;
      relevant.push({orderId:id,orderNumber:number,customerId:customerId,locationId:Number(sites[0]['Location ID']),orderStatus:statusName});
    });
    const installations=relevant.filter(function(o){return o.orderStatus==='approved';});
    const inProgress=relevant.filter(function(o){return o.orderStatus==='in progress';});
    if(inProgress.length)result={status:'REVIEW_VISIT_TYPE',reason:'IN_PROGRESS_ORDER_TYPE_UNVERIFIED',orders:inProgress};
    else if(installations.length)result={status:'INSTALLATION',reason:'APPROVED_ORDER_AT_JOB_SITE',orders:installations};
    else if(relevant.length===1 && relevant[0].orderStatus==='quoted') {
      // Measuring return visits already booked for install are not initial inspections.
      if(/\binstall(?:ation)?\b/i.test(record.title || '') || (/\b(?:return|go back)\b/i.test(text) && /\binstall(?:ation)?\b/i.test(text)))result={status:'REVIEW_VISIT_TYPE',reason:'INSTALLATION_RETURN_VISIT_ORDER_NOT_APPROVED'};
      else result={status:'EXPLICIT_PREINSPECTION',reason:'VERIFIED_QUOTED_ORDER_LIKELY_PREINSPECTION',orders:relevant};
    } else result={status:'REVIEW_VISIT_TYPE',reason:'NO_UNIQUE_ACTIVE_ORDER_AT_JOB_SITE'};
  }catch(error){result={status:'REVIEW_VISIT_TYPE',reason:'LIVE_ORDER_CLASSIFICATION_FAILED',error:String(error).slice(0,240)};}
  TMV3_PI_LIKELY_DECISIONS[key]=result;
  return result;
}

// Original: 60_Automation; Function: tmv3_preInspectionOrderBeforeCreate_
function tmv3_preInspectionOrderBeforeCreate_(record,plan) {
  // Sales Order is optional and never attached at create. If an explicit SO
  // exists, however, a known Approved order must not be misclassified as PI.
  if(!record || record.vertical!=='PreInspection')throw new Error('PI_RECORD_REQUIRED');
  const number=String(tmv3_clean_(record.orderNumber));
  if(!number)return {status:'OPTIONAL_ORDER_MISSING_NOT_AN_ERROR'};
  if(!/^\d{4,8}$/.test(number))throw new Error('PI_SO_NUMBER_INVALID_REVIEW');
  const cid=Number(plan.expectedCustomerId);
  if(!Number.isSafeInteger(cid)||cid<=0)throw new Error('PI_CUSTOMER_NOT_VERIFIED_NO_CREATE');
  const response=tmv3_fetchJson_(TMV3.API_BASE+'/v1/sales-orders/search',{
    method:'post',contentType:'application/json',
    payload:JSON.stringify({PageIndex:0,PageSize:2,Number:number,CustomerId:cid})});
  const rows=response && (response.Data||response.data);
  const count=Number(response&&(response.TotalCount!==undefined?response.TotalCount:response.totalCount));
  if(!Array.isArray(rows)||count!==1||rows.length!==1||
      String(tmv3_first_(rows[0],['Number','OrderNumber']))!==number)
    throw new Error('PI_SO_NOT_UNIQUE_OR_SOURCE_INCOMPLETE_NO_CREATE');
  const orderId=Number(tmv3_first_(rows[0],['Id','id']));
  if(!Number.isSafeInteger(orderId)||orderId<=0)throw new Error('PI_SO_ID_UNVERIFIED');
  const order=tmv3_fetchJson_(TMV3.API_BASE+'/v1/sales-orders/'+orderId,{method:'get'});
  const customer=tmv3_first_(order,['Customer','customer']);
  const status=tmv3_first_(order,['Status','status']);
  if(Number(tmv3_first_(order,['Id','id']))!==orderId||
      String(tmv3_first_(order,['OrderNumber','Number']))!==number||
      Number(tmv3_first_(customer,['Id','id']))!==cid)
    throw new Error('PI_SO_FRESH_IDENTITY_CONFLICT_NO_CREATE');
  const name=tmv3_norm_(tmv3_first_(status,['Name','name']));
  if(name==='approved')throw new Error('APPROVED_ORDER_INSTALLATION_REVIEW_NO_PI_CREATE');
  if(name==='in progress')throw new Error('IN_PROGRESS_ORDER_VISIT_TYPE_REVIEW_NO_PI_CREATE');
  if(!['quoted','incomplete','pending approval'].includes(name))
    throw new Error('PI_SO_STATUS_UNVERIFIED_NO_CREATE');
  return {status:'SO_CORROBORATION_ONLY_NOT_ATTACHED',orderId:orderId,
    orderNumber:number,orderStatus:name};
}

function tmv3_preInspectionStoreDiscovery_(records) {
  const sorted=records.slice().sort(function(a,b){return new Date(a.start)-new Date(b.start);});
  const state={checkedAt:new Date().toISOString(),total:sorted.length,
    review:sorted.filter(function(r){return r.status==='REVIEW_VISIT_TYPE';}).length,
    explicit:sorted.filter(function(r){return r.status==='EXPLICIT_PREINSPECTION';}).length,
    installation:sorted.filter(function(r){return r.status==='INSTALLATION';}).length,
    records:sorted.slice(0,15),truncated:sorted.length>15};
  // A bounded property is only the operator projection; every sweep rescans the source.
  while(JSON.stringify(state).length>2700 && state.records.length){state.records.pop();state.truncated=true;}
  const props=PropertiesService.getScriptProperties();
  const previous=tmv3_preInspectionDiscoveryStatus();
  props.setProperty('TMV3_PI_DISCOVERY',JSON.stringify(state));
  if(JSON.stringify(previous.records)!==JSON.stringify(state.records) || previous.total!==state.total) {
    tmv3_audit_('PreInspection','','','PREINSPECTION_DISCOVERY','REVIEW',JSON.stringify({total:state.total,review:state.review,explicit:state.explicit,truncated:state.truncated}));
  }
  return state;
}

// Original: 60_Automation; Function: tmv3_preInspectionDiscoveryStatus
function tmv3_preInspectionDiscoveryStatus() {
  const raw=PropertiesService.getScriptProperties().getProperty('TMV3_PI_DISCOVERY');
  if(!raw)return {status:'NOT_SCANNED',total:0,records:[]};
  try{return JSON.parse(raw);}catch(error){return {status:'INVALID_DISCOVERY_STATE',total:0,records:[]};}
}

// Original: 60_Automation; Function: tmv3_refreshPreInspectionDiscovery
function tmv3_refreshPreInspectionDiscovery() {
  // Search the four-calendar review directly; no live SO inference or imports.
  return tmv3_piCalendarReviewRefresh_();
}

function tmv3_reconcilePreInspectionEvent(eventId) {
  if(!TMV3_PI_CLEANED_RELEASE_APPROVED)return {status:'PI_STAGING_OPERATOR_PATH_DISABLED',writes:0,emailsSent:0,guestInvites:0};
  return tmv3_piOriginal_reconcilePreInspectionEvent(eventId);
}

function tmv3_piOriginal_reconcilePreInspectionEvent(eventId) {
  if(!/^[a-z0-9_]+@google\.com$/.test(String(eventId || '')))throw new Error('INVALID_PREINSPECTION_EVENT_ID');
  const started=Date.now();
  const result=tmv3_preInspectionAutomationCycle_({force:true,eventId:eventId,reason:'TARGETED_RECONCILIATION'});
  result.sheetRefresh=tmv3_refreshAutomationSheetsAfter_(started);
  return result;
}

// Original: 60_Automation; Function: tmv3_preInspectionIntakeRecords_
function tmv3_preInspectionIntakeRecords_(options) {
  options=options||{};
  const cfg=TMV3.VERTICALS.PreInspection;
  const sources=tmv3_verticalCalendars_('PreInspection',cfg).filter(function(s){
    return s.role==='PRIMARY_SHARED' && String(s.calendarId)===String(cfg.primaryCalendarId);
  });
  if(sources.length!==1)throw new Error('PI_AUTHORITATIVE_SOURCE_NOT_UNIQUE');
  const today=new Date();today.setHours(0,0,0,0);
  const end=new Date(today);end.setDate(end.getDate()+Math.max(1,Number(cfg.lookaheadDays||365)));
  const primary=CalendarApp.getCalendarById(cfg.primaryCalendarId);
  if(!primary)throw new Error('PI_PRIMARY_CALENDAR_INACCESSIBLE');
  const records=primary.getEvents(today,end).map(function(ev){
    return tmv3_calendarEvent_('PreInspection',cfg,sources[0],ev,{applyEligibility:false});
  }).filter(function(r){return r&&!tmv3_step2PreInspectionNonCustomer_(r)&&r.status!=='cancelled';});
  const seen={};
  records.forEach(function(r){
    if(!r.eventId||seen[r.eventId])throw new Error('PI_PRIMARY_EVENT_ID_DUPLICATE_OR_BLANK');
    seen[r.eventId]=true;
  });
  // This helper remains authoritative-only. Cross-calendar discovery is
  // published by tmv3_piCalendarReviewRefresh_, never silently imported.
  return records.sort(function(a,b){return new Date(a.start)-new Date(b.start);});
}

function tmv3_preInspectionPromoteCandidate_(record) {
  if(!record||record.vertical!=='PreInspection')throw new Error('PI_EVENT_REQUIRED');
  if(tmv3_preInspectionApprovalHold_(record.eventId))return {
    status:'OPERATOR_INTAKE_HOLD',record:record,calendarWrites:0,guestInvites:0};
  if(record.calendarRole==='PRIMARY_SHARED' &&
      String(record.calendarId)===String(TMV3.VERTICALS.PreInspection.primaryCalendarId))
    return {status:'PRIMARY',record:record,calendarWrites:0,guestInvites:0};
  // Private calendar candidates are DISCOVERY ONLY. A quoted order, phone,
  // address, or creator cannot by itself authorize Calendar sharing.
  return {status:'SECONDARY_REVIEW_ONLY',record:record,calendarWrites:0,guestInvites:0};
}

function tmv3_preInspectionNotificationDateParts_(value) {
  const date=new Date(value);
  if (!isFinite(date.getTime())) return null;
  const text=Utilities.formatDate(date,'America/Toronto','yyyy-MM-dd HH:mm:ss');
  const day=text.slice(0,10);
  const utc=new Date(day+'T12:00:00Z');
  return {text:text,day:day,weekday:utc.getUTCDay(),utc:utc};
}

// Original: 60_Automation; Function: tmv3_preInspectionNotificationWeek_
function tmv3_preInspectionNotificationWeek_(parts) {
  const day=new Date(parts.utc.getTime());
  day.setUTCDate(day.getUTCDate()-((day.getUTCDay()+6)%7));
  return day.toISOString().slice(0,10);
}

// Original: 60_Automation; Function: tmv3_preInspectionNotificationDecision_
function tmv3_preInspectionNotificationDecision_(record,state,nowValue) {
  const now=new Date(nowValue || Date.now());
  const start=new Date(record.start);
  if (!isFinite(start.getTime())) return {status:'EVENT_START_UNKNOWN',send:false};
  if (record.status==='cancelled' || now.getTime()>=start.getTime()) return {status:'APPOINTMENT_ENDED',send:false};
  if (state && state.status!=='SENT') return {status:'UNCERTAIN_SEND_NO_RETRY',send:false};
  const today=tmv3_preInspectionNotificationDateParts_(now);
  if (today.weekday===0 || today.weekday===6) return {status:'WEEKEND_DEFERRED',send:false};
  if (state) {
    const last=new Date(state.lastSentAt);
    if (!isFinite(last.getTime())) return {status:'SEND_TIME_UNKNOWN_NO_RETRY',send:false};
    if (now.getTime()-last.getTime()<86400000 ||
        today.day===tmv3_preInspectionNotificationDateParts_(last).day) return {status:'ALREADY_SENT',send:false};
    return {status:'FOLLOW_UP_DUE',send:true,followupNumber:Number(state.sentCount || 1)};
  }
  const created=tmv3_preInspectionNotificationDateParts_(record.calendarCreatedAt);
  if (!created) return {status:'EVENT_CREATION_TIME_UNKNOWN',send:false};
  const appointment=tmv3_preInspectionNotificationDateParts_(start);
  if (tmv3_preInspectionNotificationWeek_(created)===tmv3_preInspectionNotificationWeek_(appointment)) {
    return {status:'SAME_WEEK_DUE',send:true,followupNumber:0};
  }
  const due=new Date(created.utc.getTime());let remaining=2;
  while (remaining) {due.setUTCDate(due.getUTCDate()+1);if(due.getUTCDay()!==0 && due.getUTCDay()!==6) remaining--;}
  const eligibleAt=due.toISOString().slice(0,10)+created.text.slice(10);
  return today.text>=eligibleAt ? {status:'FIRST_DUE',send:true,followupNumber:0} :
    {status:'GRACE_PERIOD',send:false,eligibleAt:eligibleAt,timezone:'America/Toronto'};
}

// Original: 60_Automation; Function: tmv3_preInspectionSendMissingDetails_
function tmv3_preInspectionSendMissingDetails_(record) {
  if(!TMV3_PI_CLEANED_RELEASE_APPROVED)return {status:'PI_EMAILS_DISABLED_PENDING_RELEASE',sent:0};
  tmv3_piAssertLockedStandard_();
  if(!record||record.vertical!=='PreInspection'||record.calendarRole!=='PRIMARY_SHARED'||
     String(record.calendarId)!==String(TMV3.VERTICALS.PreInspection.primaryCalendarId))
    return {status:'NONAUTHORITATIVE_SOURCE_NO_SEND',sent:0};
  if(tmv3_preInspectionApprovalHold_(record.eventId))
    return {status:'OPERATOR_APPROVAL_HOLD_NO_SEND',sent:0};
  return tmv3_piSendMissingDetailsAfterRelease_(record);
}

function tmv3_piSendMissingDetailsAfterRelease_(record) {
  tmv3_piAssertLockedStandard_();
  if(tmv3_preInspectionApprovalHold_(record&&record.eventId))return {status:'APPROVAL_NOTIFICATION_HOLD',sent:0};
  const assessment=tmv3_preInspectionMissingDetailAssessment_(record);
  if (!assessment.shouldNotify) return {status:'NO_MISSING_DETAILS',sent:0};
  if (!tmv3_normEmail_(assessment.creatorEmail)) return {status:'UNVERIFIED_CREATOR_NO_SEND',sent:0};
  if (tmv3_piCreatorEmailExcluded_(assessment.creatorEmail)) return {status:'EXCLUDED_CREATOR_NO_SEND',sent:0};
  if (TMV3.MODE!=='PRODUCTION_WRITE' || TMV3.OPERATIONS.preInspectionMissingDetailEmailsEnabled!==true) return {status:'DELIVERY_GATED',sent:0};
  const props=PropertiesService.getScriptProperties();
  const key='TMV3_PI_NOTIFICATION_'+tmv3_hash_(record.eventId);
  const prior=props.getProperty(key);let state=null;
  // Old SENT markers lack timestamps. Start a conservative 24-hour waiting
  // period at migration instead of resending all historic notifications.
  if (prior==='SENT') {
    state={status:'SENT',sentCount:1,lastSentAt:new Date().toISOString(),migrated:true};
    props.setProperty(key,JSON.stringify(state));
    return {status:'ALREADY_SENT',sent:0};
  }
  if (prior) {
    try {state=JSON.parse(prior);} catch(err) {return {status:'UNCERTAIN_SEND_NO_RETRY',sent:0};}
  }
  const decision=tmv3_preInspectionNotificationDecision_(record,state);
  if (!decision.send) return Object.assign({sent:0},decision);
  const email=tmv3_preInspectionMissingDetailEmailPreview_(Object.assign({},record,{notificationFollowupNumber:decision.followupNumber}));
  if (!email.recipient) return {status:'CREATOR_MISSING_OR_AMBIGUOUS_NO_SEND',sent:0};
  if (MailApp.getRemainingDailyQuota()<3) return {status:'EMAIL_QUOTA_DEFERRED',sent:0};
  const next={status:'SENDING',sentCount:Number(state && state.sentCount || 0)+1,lastSentAt:new Date().toISOString()};
  props.setProperty(key,JSON.stringify(next));
  MailApp.sendEmail({to:email.recipient,cc:email.cc.join(','),subject:email.subject,
    body:(decision.followupNumber ? 'FOLLOW-UP #'+decision.followupNumber+': ' : '')+'Please update the missing Pre-Inspection details: '+assessment.missing.map(function(x){return x.message;}).join('; '),
    htmlBody:email.htmlBody,name:'Classic Fireplace & BBQ Store'});
  next.status='SENT';props.setProperty(key,JSON.stringify(next));
  tmv3_audit_('PreInspection',record.eventId,'','MISSING_DETAILS_EMAIL','SENT',JSON.stringify({to:email.recipient,followupNumber:decision.followupNumber,missing:assessment.missing}));
  return {status:'SENT',sent:1,to:email.recipient,followupNumber:decision.followupNumber,missing:assessment.missing};
}

// Original: 60_Automation; Function: tmv3_preInspectionSweepKey_

function tmv3_preInspectionSweepKey_(record) {
  const start = new Date(record.start);
  if (!isFinite(start.getTime())) throw new Error('PREINSPECTION_SWEEP_INVALID_START');
  return start.toISOString() + '|' + String(record.eventId || '');
}

// Original: 60_Automation; Function: tmv3_preInspectionTargetedIntake_
function tmv3_preInspectionTargetedIntake_(eventId,props) {
  const pending=JSON.parse(props.getProperty('TMV3_PI_PENDING_EVENTS') || '[]');
  const record=tmv3_preInspectionRecoveryRecord_(eventId);
  const start=new Date(record.start);
  if (!isFinite(start.getTime()) || record.status==='cancelled') throw new Error('TARGET_EVENT_INVALID_OR_CANCELLED');
  const first=new Date();first.setHours(0,0,0,0);
  const last=new Date(first);last.setDate(last.getDate()+Number(TMV3.VERTICALS.PreInspection.lookaheadDays || 365));
  if (pending.indexOf(eventId)<0 && (start<first || start>=last)) return [];
  return [record];
}

// Original: 60_Automation; Function: tmv3_preInspectionAutomationCycle_
function tmv3_preInspectionAutomationCycle_(options) {
  options=options||{};
  if(!TMV3_PI_CLEANED_RELEASE_APPROVED) {
    return {status:'PI_STAGING_RELEASE_GATE_NO_BUSINESS_WRITES',
      eventId:options.eventId||'',processed:0,writes:0,notifications:0,
      nextAction:'RUN_PI_CLEANED_VALIDATION_AND_VERIFY_FULL_V3_DEPENDENCIES'};
  }

  tmv3_piAssertLockedStandard_();
  options = options || {};
  if (!tmv3_operationWritesEnabled_('PREINSPECTIONAUTO')) return {status:'PREINSPECTION_GATED',writes:0};
  if (TMV3.OPERATIONS.preInspectionAutomationScheduleHours && options.force !== true) return {status:'SCHEDULE_ONLY',writes:0};
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(options.force === true ? 30000 : 1000)) return {status:'LOCK_BUSY',writes:0};
  const props = PropertiesService.getScriptProperties();
  const result = {status:'COMPLETE',processed:0,writes:0,notifications:0,results:[]};
  const previousApiContext=TMV3_PREINSPECTION_API_CONTEXT_ACTIVE;
  TMV3_PREINSPECTION_API_CONTEXT_ACTIVE=true;
  try {
    const scopeEventId = tmv3_clean_(TMV3.OPERATIONS.preInspectionAutomationCanaryEventId || options.eventId);
    const sweepRequested = options.sweepStart === true || options.sweepContinue === true;
    if (sweepRequested && scopeEventId) return {status:'SWEEP_BLOCKED_BY_CANARY_SCOPE',processed:0,writes:0};
    const records=scopeEventId ? tmv3_preInspectionTargetedIntake_(scopeEventId,props) : tmv3_preInspectionIntakeRecords_({recordDiscovery:true,resolveLikely:true});
    result.discovery=tmv3_preInspectionDiscoveryStatus();
    const budget=tmv3_strivenApiBudgetStatus_();
    if(budget.exhausted||budget.remaining<30) {
      result.status='API_BUDGET_WAIT';result.apiBudget=budget;
      result.resume='NEXT_SCHEDULED_SLOT_AFTER_DAILY_RESET';
      const pauseKey='TMV3_PI_BUDGET_PAUSE_DATE';
      if(props.getProperty(pauseKey)!==budget.date) {
        props.setProperty(pauseKey,budget.date);
        tmv3_audit_('PreInspection','','','PREINSPECTION_API_BUDGET_WAIT','WAIT',JSON.stringify({date:budget.date,count:budget.count,softLimit:budget.softLimit,resume:result.resume,triggers:tmv3_listTriggers()}));
      }
      props.setProperty('TMV3_PI_LAST_RESULT',JSON.stringify(result));
      return result;
    }
    const now = Date.now();
    const last = Number(props.getProperty('TMV3_PI_LAST_CYCLE') || 0);
    if (options.force !== true && now-last < Math.max(1,Number(TMV3.OPERATIONS.preInspectionAutomationIntervalMinutes || 5))*60000) {
      return {status:'THROTTLED',writes:0};
    }
    props.setProperty('TMV3_PI_LAST_CYCLE',String(now));
    // Reuse the cache-age gates and last-good source policy; refresh failure
    // stops this cycle before task creation or misleading missing-data emails.
    tmv3_refreshOperationalSourcesIfNeeded_(120,false);
    let pending = JSON.parse(props.getProperty('TMV3_PI_PENDING_EVENTS') || '[]');
    if (!scopeEventId) pending.forEach(function(id) {
      if (records.some(function(r) {return r.eventId===id;})) return;
      try {
        const recovered=tmv3_preInspectionRecoveryRecord_(id);
        if(recovered.calendarRole==='PRIMARY_SHARED' || tmv3_preInspectionClassifyLikely_(recovered,TMV3.VERTICALS.PreInspection).status==='EXPLICIT_PREINSPECTION') records.push(recovered);
        else pending=pending.filter(function(pendingId){return pendingId!==id;});
      }
      catch(err) {tmv3_audit_('PreInspection',id,'','PREINSPECTION_RECOVERY','ATTENTION',String(err));}
    });
    let selected = scopeEventId ? records.filter(function(r){return r.eventId===scopeEventId;}) : records;
    let sweep = null;
    if (sweepRequested && scopeEventId) {
      return {status:'SWEEP_BLOCKED_BY_CANARY_SCOPE',processed:0,writes:0,eligible:records.length};
    }
    if (sweepRequested) {
      const raw = props.getProperty('TMV3_PI_SWEEP_STATE') || '';
      let prior = null;
      try {if (raw) prior=JSON.parse(raw);} catch (ignored) {}
      const valid = prior && prior.status==='IN_PROGRESS' &&
        isFinite(new Date(prior.startedAt).getTime()) &&
        now-new Date(prior.startedAt).getTime() < 86400000;
      if (options.sweepContinue === true && !valid) {
        return {status:'NO_ACTIVE_PREINSPECTION_SWEEP',processed:0,writes:0};
      }
      sweep = valid ? prior : {status:'IN_PROGRESS',startedAt:new Date(now).toISOString(),
        lastKey:'',processed:0};
      // Deterministic position survives batch boundaries without a 9-KB
      // Script Properties queue; new earlier events enter the next sweep.
      selected.sort(function(a,b) {
        const aa=tmv3_preInspectionSweepKey_(a),bb=tmv3_preInspectionSweepKey_(b);
        return aa<bb?-1:aa>bb?1:0;
      });
      selected=selected.filter(function(r){return tmv3_preInspectionSweepKey_(r)>sweep.lastKey;});
      result.eligible=records.length;
    }
    // Scheduled slots still evaluate notification deadlines for all
    // authoritative events; continuations do not repeat the full mail scan.
    const notificationRecords=options.sweepContinue===true ? [] :
      (scopeEventId ? selected : records);
    notificationRecords.filter(function(r){return r.calendarRole==='PRIMARY_SHARED'&&!tmv3_preInspectionApprovalHold_(r.eventId);}).forEach(function(r){
      try {const notice=tmv3_preInspectionSendMissingDetails_(r);result.notifications+=Number(notice.sent || 0);}
      catch(err) {result.status='ATTENTION';tmv3_preInspectionFailureAlert_(r.eventId,String(err));}
    });
    let cursor = sweep ? 0 : Number(props.getProperty('TMV3_PI_CURSOR') || 0);
    // Preserve the tested three-event and one-mutation safety budgets.
    const cap = Math.max(1,Math.min(3,Number(TMV3.OPERATIONS.preInspectionAutomationMaxEventsPerRun || 3)));
    if (selected.length) cursor = cursor % selected.length;
    for (let i=0;i<Math.min(cap,selected.length);i++) {
      let record = selected[(cursor+i)%selected.length];
      const one = {eventId:record.eventId};
      result.processed++;
      try {
        if(tmv3_preInspectionApprovalHold_(record.eventId)) {
          one.status='APPROVED_INTAKE_TASK_AND_NOTIFICATION_HELD';
          one.notification={status:'APPROVAL_NOTIFICATION_HOLD',sent:0};
          result.results.push(one);continue;
        }
        const doneKey='TMV3_PI_DONE_'+tmv3_hash_(record.eventId);
        const done=JSON.parse(props.getProperty(doneKey) || 'null');
        if (!options.eventId && done && done.fingerprint===record.fingerprint && now-Number(done.at || 0)<3600000) {
          one.status='VERIFIED_UNCHANGED_CACHE';result.results.push(one);continue;
        }
        const intake = tmv3_preInspectionPromoteCandidate_(record);
        one.intake = intake.status;
        record = intake.record;
        if (intake.status !== 'PRIMARY') {
          if (pending.indexOf(record.eventId)<0) pending.push(record.eventId);
          // Wait until authoritative intake and matching have been attempted.
          one.notification = {status:'WAITING_FOR_INTAKE',sent:0};
        } else {
          const preservation=tmv3_restorePreInspectionLegacyTitle_(record);
          if (preservation.writes) record=tmv3_preInspectionRecoveryRecord_(record.eventId);
          let plans = tmv3_step7FreshPlansForEvent_('PreInspection',record.eventId);
          if (plans.length===1 && plans[0].blocker &&
              !tmv3_preInspectionHistoricalCoverageVerified_(plans[0])) {
            one.optionalOrderEvidence=tmv3_preInspectionOptionalOrderLocationEvidence_(record);
            if (one.optionalOrderEvidence.status==='CACHE_REPAIRED') plans=tmv3_step7FreshPlansForEvent_('PreInspection',record.eventId);
          }
          if (plans.length === 1 && !plans[0].blocker && (plans[0].actions || []).some(function(a) {
            return a===TMV3_STEP7_ACTION.CREATE_TASK || a===TMV3_STEP7_ACTION.RECREATE_TASK;
          })) {
            plans = tmv3_preInspectionFreshHistoryBeforeCreate_(record,plans[0]);
          }
          if (plans.length !== 1) throw new Error('MULTIPLE_CANONICAL_PLANS_NO_WRITE');
          const plan = plans[0];
           if (!plan.blocker && Array.isArray(plan.actions) &&
               plan.actions.indexOf(TMV3_STEP7_ACTION.RECREATE_TASK)>=0)
             throw new Error('PI_AUTOMATIC_RECREATE_DISABLED_REVIEW_HISTORY_FIRST');
           // Before any Task mutation, enforce the verified Calendar-owner
           // Employee as Requested By. Link verification alone does not patch Task.
           if (!plan.blocker && Array.isArray(plan.actions) &&
               plan.actions.some(function(action){
                 return action!==TMV3_STEP7_ACTION.VERIFY_CALENDAR_LINKS;
               })) tmv3_piAssertRequestedByPlan_(plan,record);
          if(!plan.blocker && plan.actions.some(function(a){return a===TMV3_STEP7_ACTION.CREATE_TASK || a===TMV3_STEP7_ACTION.RECREATE_TASK;}))one.orderStatusBeforeCreate=tmv3_preInspectionOrderBeforeCreate_(record,plan);
          one.plan=plan.plan;one.blocker=plan.blocker || '';one.taskId=plan.taskId || '';
          // Preserve the single business-write transaction cap. Do not
          // advance the durable sweep cursor past a record whose update has
          // been deferred solely because this batch used its write allowance.
          if (!plan.blocker && plan.actions.length && result.writes>=1) {
            result.processed--;
            result.deferredWriteBudget=true;
            break;
          }
          if (!plan.blocker && plan.actions.length && result.writes<1) {
            // One event transaction per run, using the existing fresh ownership,
            // lifecycle, durable create guards and field-level readbacks.
            if (plan.actions.some(function(action){return action!==TMV3_STEP7_ACTION.VERIFY_CALENDAR_LINKS;})) result.writes++;
            one.execution = tmv3_executeFreshStep7Selection_({previousPlan:plan.plan,contract:plan},'PREINSPECTIONAUTO','ALL');
            if (!/VERIFIED|CONVERGED|NO_CHANGE/.test(String(one.execution.status || ''))) {
              throw new Error('PREINSPECTION_RECONCILIATION_NOT_VERIFIED: '+one.execution.status);
            }
          }
          const fresh = tmv3_findFreshEventRecord_('PreInspection',record.eventId);
          one.notification = tmv3_preInspectionSendMissingDetails_(fresh);
          if (/CREATOR_MISSING|UNCERTAIN_SEND/.test(one.notification.status || '')) {
            result.status='ATTENTION';
            tmv3_preInspectionFailureAlert_(record.eventId,one.notification.status);
          }
          const after = tmv3_step7FreshPlansForEvent_('PreInspection',record.eventId);
          one.afterPlans=after.map(function(p){return {plan:p.plan,blocker:p.blocker,taskId:p.taskId,actions:p.actions};});
          const complete = tmv3_preInspectionCompletionVerified_(after,one.execution);
          if (complete) {
            if (tmv3_preInspectionHistoricalCoverageVerified_(after[0])) {
              one.status='FULFILLED_NO_RECREATE';one.historicalEvidence=one.blocker;one.blocker='';
            }
            pending=pending.filter(function(id){return id!==record.eventId;});
            const inferredKey='TMV3_PI_INFERRED_'+tmv3_hash_(record.eventId);
            if(props.getProperty(inferredKey))props.deleteProperty(inferredKey);
            props.setProperty(doneKey,JSON.stringify({fingerprint:fresh.fingerprint,at:Date.now()}));
          }
          else {
            one.status='ATTENTION';
            result.status='ATTENTION';
            if (pending.indexOf(record.eventId)<0) pending.push(record.eventId);
          }
        }
        result.notifications += Number(one.notification && one.notification.sent || 0);
      } catch(err) {
        one.status='ATTENTION';one.error=String(err && err.message || err);result.status='ATTENTION';
        if (pending.indexOf(record.eventId)<0) pending.push(record.eventId);
        tmv3_audit_('PreInspection',record.eventId,'','PREINSPECTIONAUTO','ATTENTION',one.error);
        tmv3_preInspectionFailureAlert_(record.eventId,one.error);
        if(/TMV3_STRIVEN_API_DAILY_GUARD|daily api usage limit/i.test(one.error)) {
          // Preserve this event's cursor position for a fresh reconciliation after reset.
          result.status='API_BUDGET_WAIT';result.processed--;
          result.results.push(one);break;
        }
      }
      result.results.push(one);
    }
    if (sweep) {
      if (result.processed) sweep.lastKey=tmv3_preInspectionSweepKey_(selected[result.processed-1]);
      sweep.processed=Number(sweep.processed || 0)+result.processed;
      const remaining=Math.max(0,selected.length-result.processed);
      sweep.status=remaining ? 'IN_PROGRESS' :
        (pending.length ? 'SCANNED_WITH_PENDING_REVIEW' : 'COMPLETE');
      props.setProperty('TMV3_PI_SWEEP_STATE',JSON.stringify(sweep));
      result.sweep={status:sweep.status,eligible:result.eligible,attempted:sweep.processed,
        remaining:remaining,pendingReview:pending.length,lastKey:sweep.lastKey};
      // A successful batch is not a completed queue; retain honest backlog state.
      if(result.status==='COMPLETE' && remaining>0) result.status='IN_PROGRESS';
      else if(result.status==='COMPLETE' && pending.length>0) result.status='ATTENTION_PENDING_REVIEW';
    } else if (!options.eventId && selected.length) {
      props.setProperty('TMV3_PI_CURSOR',String((cursor+result.processed)%selected.length));
    }
    props.setProperty('TMV3_PI_PENDING_EVENTS',JSON.stringify(pending));
    props.setProperty('TMV3_PI_LAST_RESULT',JSON.stringify({status:result.status,
      processed:result.processed,writes:result.writes,notifications:result.notifications,
      events:result.results.map(function(x){return {eventId:x.eventId,status:x.status || x.plan || x.intake,
        taskId:x.taskId,error:String(x.error || '').slice(0,500)};})}));
    return result;
  } catch(err) {
    result.status='ATTENTION';result.error=String(err && err.message || err);
    tmv3_audit_('PreInspection','','','PREINSPECTION_CYCLE','ATTENTION',result.error);
    tmv3_preInspectionFailureAlert_('',result.error);
    return result;
  } finally {TMV3_PREINSPECTION_API_CONTEXT_ACTIVE=previousApiContext;lock.releaseLock();}
}

// Original: 60_Automation; Function: tmv3_runPreInspectionAutomation
function tmv3_runPreInspectionAutomation() {
  if(!TMV3_PI_CLEANED_RELEASE_APPROVED)return {status:'PI_STAGING_OPERATOR_PATH_DISABLED',writes:0,emailsSent:0,guestInvites:0};
  return tmv3_piOriginal_runPreInspectionAutomation();
}

function tmv3_piOriginal_runPreInspectionAutomation() {
  const sheetRefreshStarted=Date.now();
  const result=tmv3_preInspectionAutomationCycle_({force:true,reason:'MANUAL_KICK'});
  result.sheetRefresh=tmv3_refreshAutomationSheetsAfter_(sheetRefreshStarted);
  Logger.log(JSON.stringify(result));return result;
}

// Original: 60_Automation; Function: tmv3_verifyPreInspectionAutomation
function tmv3_verifyPreInspectionAutomation(eventId) {
  // Fast cache-only diagnostic: three events per invocation, no business/API writes.
  // Fresh Task history, Requested By and field 854 are NOT certified by this function.
  const startedMs=Date.now();
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(1500)) {
    const busy={status:'PI_VERIFY_LOCK_BUSY',businessWrites:0};
    Logger.log('PI_VERIFICATION_BATCH='+JSON.stringify(busy));
    return busy;
  }
  try {
    const today=Utilities.formatDate(new Date(),TMV3_TIMEZONE,'yyyy-MM-dd');
    const rows=tmv3_rows_(TMV3.VERTICALS.PreInspection.sheet);
    function startText(row) {
      const value=row['Calendar Start'];
      return value instanceof Date
        ? Utilities.formatDate(value,TMV3_TIMEZONE,'yyyy-MM-dd HH:mm:ss')
        : String(value||'');
    }
    const queue=rows.filter(function(row) {
      if(String(row['Vertical']||'')!=='PreInspection' ||
         String(row['Calendar Role']||'')!=='PRIMARY_SHARED') return false;
      const start=startText(row);
      return !/^\d{4}-\d{2}-\d{2}/.test(start) || start.slice(0,10)>=today;
    }).sort(function(a,b) {
      return startText(a).localeCompare(startText(b)) ||
        String(a['Event ID']||'').localeCompare(String(b['Event ID']||''));
    });
    const seen={};
    queue.forEach(function(row) {
      const id=String(row['Event ID']||'');
      if(!id||seen[id]) throw new Error('PI_DIAGNOSTIC_EVENT_ID_EMPTY_OR_DUPLICATE');
      seen[id]=true;
    });
    const target=String(eventId||'').trim();
    const selectedQueue=target
      ? queue.filter(function(row){return String(row['Event ID'])===target;})
      : queue;
    const fingerprint=tmv3_hash_(today+'|'+queue.map(function(row) {
      return String(row['Event ID'])+'|'+startText(row)+'|'+String(row['Fingerprint']||'');
    }).join('|'));
    const props=PropertiesService.getScriptProperties();
    const cursorKey='TMV3_PI_VERIFY_CURSOR';
    let offset=0;
    if(!target) {
      let state=null;
      try{state=JSON.parse(props.getProperty(cursorKey)||'null');}catch(ignored){}
      if(state&&state.fingerprint===fingerprint&&
         Number.isInteger(state.offset)&&state.offset>=0&&state.offset<queue.length) {
        offset=state.offset;
      }
    }
    const batch=target?selectedQueue.slice(0,1):selectedQueue.slice(offset,offset+3);
    const records=batch.map(function(row) {
      const linked=String(row['Existing Task Link']||'');
      const detail=String(row['Task']||'');
      const linkedId=(/[?&]TaskID=(\d+)/i.exec(linked)||[])[1]||'';
      const detailId=(/\bTask\s*#\s*(\d+)/i.exec(detail)||[])[1]||'';
      const taskId=linkedId||detailId;
      const mapping=String(row['Mapping Status']||'');
      const checklist=String(row['Data Checklist']||'');
      const warnings=checklist.split(/\r?\n/).filter(function(line) {
        return line.indexOf('⚠')>=0;
      }).slice(0,5);
      const pending=checklist.split(/\r?\n/).filter(function(line) {
        return line.indexOf('⏳')>=0;
      }).slice(0,5);
      let disposition='CACHE_SNAPSHOT_UNRESOLVED';
      if(linkedId&&detailId&&linkedId!==detailId)
        disposition='TASK_ID_CONFLICT_REVIEW_NO_WRITE';
      else if(/\b(?:NOT_RUN|REVIEW|BLOCKED)\b/i.test(mapping))
        disposition='OPERATOR_REVIEW_REQUIRED';
      else if(/\b(?:CREATE_TASK|RECREATE_TASK)\b/i.test(mapping))
        disposition='FRESH_TASK_HISTORY_REQUIRED_NO_CREATE';
      else if(!linkedId&&taskId)
        disposition='TASK_LINK_MISSING_LIVE_CHECK_REQUIRED';
      else if(!taskId)
        disposition='TASK_ID_NOT_IN_OPERATOR_SNAPSHOT';
      else if(warnings.length)
        disposition='OPERATOR_WARNING_LIVE_CHECK_REQUIRED';
      else disposition='CACHED_MATCH_LIVE_NOT_VERIFIED';
      return {
        eventId:String(row['Event ID']||''),
        customerNumber:String(row['Customer # (Calendar)']||''),
        start:startText(row),
        taskId:taskId,
        taskLinkPresent:!!linkedId,
        mapping:mapping,
        disposition:disposition,
        warnings:warnings,
        pending:pending,
        nextAction:String(row['Issue / Next Action']||'').slice(0,180)
      };
    });
    const hasMore=!target&&offset+batch.length<queue.length;
    const result={
      status:target&&!batch.length?'TARGET_NOT_IN_OPERATOR_VIEW':
        !queue.length?'NO_CURRENT_OPERATOR_EVENTS_REVIEW_REFRESH':'CACHE_ONLY_BATCH',
      version:TMV3.VERSION,
      source:'PREINSPECTION_OPERATOR_SNAPSHOT',
      evidence:'CACHE_ONLY_NOT_LIVE_STRIVEN_OR_CALENDAR_VERIFICATION',
      automationEnabled:TMV3.OPERATIONS.preInspectionAutomationEnabled,
      emailEnabled:TMV3.OPERATIONS.preInspectionMissingDetailEmailsEnabled,
      total:queue.length,
      processed:batch.length,
      startIndex:target?null:offset,
      nextIndex:hasMore?offset+batch.length:null,
      remaining:target?null:Math.max(0,queue.length-offset-batch.length),
      hasMore:hasMore,
      records:records,
      elapsedMs:Date.now()-startedMs,
      strivenCalls:0,calendarReads:0,calendarWrites:0,taskWrites:0,emailsSent:0,
      diagnosticCursorPropertyOnly:!target
    };
    if(!target) {
      if(hasMore)props.setProperty(cursorKey,JSON.stringify({fingerprint:fingerprint,offset:offset+batch.length}));
      else props.deleteProperty(cursorKey);
    }
    Logger.log('PI_VERIFICATION_BATCH='+JSON.stringify(result));
    return result;
  }catch(error) {
    Logger.log('PI_VERIFICATION_ERROR='+String(error&&error.message||error).slice(0,300));
    throw error;
  }finally{lock.releaseLock();}
}

// Original: 60_Automation; Function: tmv3_preInspectionScheduledCycle
function tmv3_preInspectionScheduledCycle() {
  return {status:'RETIRED_FIVE_MINUTE_SCHEDULE',writes:0};
}

// Original: 60_Automation; Function: tmv3_preInspectionCheck0800
function tmv3_preInspectionCheck0800() { return tmv3_runPreInspectionDailySlot_('08:00'); }

// Original: 60_Automation; Function: tmv3_preInspectionCheck1300
function tmv3_preInspectionCheck1300() { return tmv3_runPreInspectionDailySlot_('13:00'); }

// Original: 60_Automation; Function: tmv3_preInspectionCheck1700
function tmv3_preInspectionCheck1700() { return tmv3_runPreInspectionDailySlot_('17:00'); }

// Original: 60_Automation; Function: tmv3_preInspectionScheduleNextSweep_
function tmv3_preInspectionScheduleNextSweep_() {
  return 'DEFERRED_TO_NEXT_STANDARD_SLOT';
}

// Original: 60_Automation; Function: tmv3_preInspectionClearContinuationTrigger_
function tmv3_preInspectionClearContinuationTrigger_() {
  if(!TMV3_PI_CLEANED_RELEASE_APPROVED)return {status:'PI_STAGING_OPERATOR_PATH_DISABLED',writes:0,emailsSent:0,guestInvites:0};
  return tmv3_piOriginal_preInspectionClearContinuationTrigger_();
}

function tmv3_piOriginal_preInspectionClearContinuationTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction()==='tmv3_preInspectionContinueEligibleSweep') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

// Original: 60_Automation; Function: tmv3_preInspectionLogSweep_
function tmv3_preInspectionLogSweep_(slot,result) {
  const summary={
    slot:slot,status:result.status,processed:result.processed,
    writes:result.writes,notifications:result.notifications,
    sweep:result.sweep || null,continuation:result.continuation || null,
    error:result.error || '',
    calendarReview:result.calendarReview && result.calendarReview.status || 'NOT_RUN',
    operatorQueue:result.operatorQueue && result.operatorQueue.status || 'NOT_RUN',
    attentionEvents:(result.results || []).filter(function(r){
      return r.status==='ATTENTION'||!!r.blocker||!!r.error;
    }).map(function(r){return {eventId:r.eventId,status:r.status||'',blocker:r.blocker||'',error:r.error||''};})
  };
  Logger.log(JSON.stringify(summary));
  tmv3_audit_('PreInspection','','','PREINSPECTION_SWEEP',result.status||'UNKNOWN',JSON.stringify(summary));
}

// Original: 60_Automation; Function: tmv3_preInspectionMaybeContinueSweep_
function tmv3_preInspectionMaybeContinueSweep_(result) {
  const errors=[result.error||''].concat((result.results||[]).map(function(r){return r.error||'';})).join(' ');
  if(result.status==='API_BUDGET_WAIT'||/TMV3_STRIVEN_API_DAILY_GUARD|daily api usage limit/i.test(errors)) {
    result.continuation='WAIT_FOR_DAILY_API_BUDGET';
    return result;
  }
  // A reviewed historical relationship is not an infrastructure failure.
  const allTechnicalFailures=(result.results||[]).length>0 &&
    result.results.every(function(r){return !!r.error;});
  if(allTechnicalFailures) {
    result.continuation='PAUSED_TECHNICAL_FAILURE_AUTOMATIC_RECOVERY';
    return result;
  }
  if((result.sweep&&result.sweep.remaining>0)||result.status==='LOCK_BUSY') {
    try {result.continuation=tmv3_preInspectionScheduleNextSweep_();}
    catch(err) {
      result.status='ATTENTION';result.continuation='SCHEDULING_FAILED';
      result.error=String(err&&err.message||err);
      tmv3_audit_('PreInspection','','','ELIGIBLE_SWEEP','ATTENTION',result.error);
      tmv3_preInspectionFailureAlert_('',result.error);
    }
  }
  return result;
}

// Original: 60_Automation; Function: tmv3_preInspectionContinueEligibleSweep
function tmv3_preInspectionContinueEligibleSweep() {
  return {status:'RETIRED_OUTSIDE_STANDARD_SLOTS',writes:0,notifications:0};
}

// Original: 60_Automation; Function: tmv3_runPreInspectionDailySlot_
function tmv3_runPreInspectionDailySlot_(slot) {
  if(!TMV3_PI_CLEANED_RELEASE_APPROVED)return {status:'PI_STAGING_OPERATOR_PATH_DISABLED',writes:0,emailsSent:0,guestInvites:0};
  return tmv3_piOriginal_runPreInspectionDailySlot_(slot);
}

function tmv3_piOriginal_runPreInspectionDailySlot_(slot) {
  tmv3_piAssertLockedStandard_();
  const sheetRefreshStarted=Date.now();
  // Report-only Calendar synchronization; failure preserves last-good view and never blocks Striven work.
  let calendarReview=null;
  try { calendarReview=tmv3_piCalendarReviewRefresh_(); }
  catch(err) {
    calendarReview={status:'PI_REVIEW_ATTENTION',error:String(err&&err.message||err).slice(0,400)};
    try {tmv3_audit_('PreInspection','','','PI_CALENDAR_REVIEW','ATTENTION',calendarReview.error);}catch(ignored){}
  }
  let result=tmv3_preInspectionAutomationCycle_({force:true,sweepStart:true,reason:'DAILY_SLOT_'+slot});
  const totals={processed:Number(result.processed||0),writes:Number(result.writes||0),notifications:Number(result.notifications||0),results:(result.results||[]).slice()};
  let priorRemaining=result.sweep && result.sweep.remaining;
  const batchStatuses=[result.status];
  while (result.sweep && result.sweep.remaining>0 && result.processed>0 &&
      result.status!=='API_BUDGET_WAIT' && result.status!=='LOCK_BUSY' &&
      Date.now()-sheetRefreshStarted<240000) {
    const next=tmv3_preInspectionAutomationCycle_({force:true,sweepContinue:true,reason:'DAILY_SLOT_'+slot});
    batchStatuses.push(next.status);
    totals.processed+=Number(next.processed||0);totals.writes+=Number(next.writes||0);totals.notifications+=Number(next.notifications||0);totals.results.push.apply(totals.results,next.results||[]);
    if (!next.sweep || next.sweep.remaining>=priorRemaining || !next.processed) {result=next;break;}
    priorRemaining=next.sweep.remaining;result=next;
  }
  result.batchStatuses=batchStatuses;
  result.processed=totals.processed;result.writes=totals.writes;result.notifications=totals.notifications;result.results=totals.results;
  if (result.status==='LOCK_BUSY') PropertiesService.getScriptProperties().setProperty('TMV3_PI_RETRY_SWEEP_START','1');
  tmv3_preInspectionMaybeContinueSweep_(result);
  if(result.sweep && result.sweep.remaining>0 && result.status==='COMPLETE')result.status='IN_PROGRESS';
  if(result.sweep && result.sweep.remaining===0 && result.sweep.pendingReview>0 && result.status==='COMPLETE')
    result.status='ATTENTION_PENDING_REVIEW';
  if(calendarReview && calendarReview.status==='PI_REVIEW_PARTIAL_SOURCE_ACCESS' && result.status==='COMPLETE')
    result.status='ATTENTION_CALENDAR_COVERAGE';
  result.sheetRefresh=tmv3_refreshAutomationSheetsAfter_(sheetRefreshStarted);
  try{result.operatorQueue=tmv3_preInspectionRefreshOperatorView_();}
  catch(err){
    result.operatorQueue={status:'PI_OPERATOR_QUEUE_ATTENTION',error:String(err&&err.message||err).slice(0,350)};
    try{tmv3_audit_('PreInspection','','','PI_OPERATOR_QUEUE','ATTENTION',result.operatorQueue.error);}catch(ignore){}
  }
  result.calendarReview=calendarReview;
  if(result.status==='COMPLETE' && (!result.operatorQueue || result.operatorQueue.status!=='PI_OPERATOR_VIEW_REFRESHED'))
    result.status='ATTENTION_OPERATOR_QUEUE';
  if(result.status==='COMPLETE' && (!calendarReview || !/^(?:PI_REVIEW_REFRESHED|PI_REVIEW_UNCHANGED)$/.test(String(calendarReview.status||''))))
    result.status='ATTENTION_CALENDAR_COVERAGE';
  tmv3_preInspectionLogSweep_(slot,result); // Audit only AFTER both publication outcomes are known.
  return result;
}

// Original: 60_Automation; Function: tmv3_installPreInspectionAutomationTrigger
function tmv3_installPreInspectionAutomationTrigger() {
  if(!TMV3_PI_CLEANED_RELEASE_APPROVED)return {status:'PI_STAGING_OPERATOR_PATH_DISABLED',writes:0,emailsSent:0,guestInvites:0};
  return tmv3_piOriginal_installPreInspectionAutomationTrigger();
}

function tmv3_piOriginal_installPreInspectionAutomationTrigger() { return tmv3_installStandardDailyTriggers(); }

// Original: 60_Automation; Function: tmv3_preInspectionFreshHistoryBeforeCreate_
function tmv3_preInspectionFreshHistoryBeforeCreate_(record,plan) {
  const customer=tmv3_step4IdentityIndex_().customerById[String(plan.expectedCustomerId)];
  if (!customer) throw new Error('CREATE_CUSTOMER_NOT_IN_VERIFIED_CACHE');
  const ids=tmv3_searchPreInspectionTaskIdsForCustomer_(customer);
  if (ids.length>10) throw new Error('CREATE_HISTORY_API_LIMIT_REVIEW_NO_CREATE');
  ids.forEach(function(id){
    const task=tmv3_getTaskById_(Number(id));
    if (String(task['Task ID'])!==String(id) ||
        String(task['Customer ID'])!==String(plan.expectedCustomerId) ||
        Number(task['Task Type ID'])!==105) throw new Error('CREATE_HISTORY_OWNERSHIP_FAILED_NO_CREATE');
    tmv3_upsertTaskCacheRow_(task);
  });
  // Recompute the same lifecycle gate on fresh complete Customer task history.
  return tmv3_step7FreshPlansForEvent_('PreInspection',record.eventId);
}

// Original: 60_Automation; Function: tmv3_preInspectionRecoveryRecord_
function tmv3_preInspectionRecoveryRecord_(eventId) {
  const cfg=TMV3.VERTICALS.PreInspection;
  const sources=tmv3_verticalCalendars_('PreInspection',cfg).concat(
    (cfg.secondaryCalendarIds || []).map(function(id){return {calendarId:id,calendarName:'Stephen',role:'SECONDARY_STEPHEN'};})
  );
  for(let i=0;i<sources.length;i++) {
    const found=tmv3_findCalendarEventById_(sources[i].calendarId,eventId);
    if (!found || !found.event) continue;
    const record=tmv3_calendarEvent_('PreInspection',cfg,sources[i],found.event,{applyEligibility:false});
    if (record && !tmv3_step2PreInspectionNonCustomer_(record) &&
        (sources[i].role==='PRIMARY_SHARED' || tmv3_preInspectionExplicitCandidate_(record))) return record;
  }
  throw new Error('RECOVERY_EVENT_MISSING_OR_NO_LONGER_PREINSPECTION');
}

// Original: 60_Automation; Function: tmv3_preInspectionFailureAlert_
function tmv3_preInspectionFailureAlert_(eventId,error) {
  if(!TMV3_PI_CLEANED_RELEASE_APPROVED)return {status:'PI_FAILURE_EMAIL_DISABLED_PENDING_RELEASE',sent:0};
  if (TMV3.MODE!=='PRODUCTION_WRITE' || TMV3.OPERATIONS.preInspectionFailureAlertsEnabled!==true) return {status:'ALERT_GATED'};
  const day=Utilities.formatDate(new Date(),TMV3_TIMEZONE,'yyyy-MM-dd');
  const key='TMV3_PI_FAILURE_'+tmv3_hash_(day+'|'+eventId+'|'+error);
  const props=PropertiesService.getScriptProperties();
  if (props.getProperty(key)) return {status:'ALREADY_ALERTED'};
  try {
    if (MailApp.getRemainingDailyQuota()<1) return {status:'QUOTA_DEFERRED'};
    props.setProperty(key,'SENDING');
    const esc=tmv3_preInspectionHtmlEscape_;
    MailApp.sendEmail({to:TMV3.OPERATIONS.preInspectionFailureAlertEmail,
      subject:'Pre-Inspection Automation: action required',
      body:'Pre-Inspection automation needs review. Event: '+(eventId || 'Source refresh')+'\n'+error,
      htmlBody:'<p><em>This is an automated email notification from Classic Fireplace &amp; BBQ Store.</em></p>'+
        '<p>Hi Pramodh,</p><p style="color:#B42318;font-weight:700;">Pre-Inspection automation requires review.</p>'+
        '<p>Event: '+esc(eventId || 'Source refresh')+'</p><p>'+esc(error)+'</p>'+
        '<p>The affected action remains queued. Verify any uncertain write before retrying.</p>',
      name:'Classic Fireplace & BBQ Store'});
    props.setProperty(key,'SENT');return {status:'SENT'};
  } catch(err) {
    tmv3_audit_('PreInspection',eventId,'','PREINSPECTION_FAILURE_ALERT','ATTENTION',String(err));
    return {status:'ALERT_UNCERTAIN_NO_RETRY'};
  }
}

// Original: 60_Automation; Function: tmv3_restorePreInspectionLegacyTitle_
function tmv3_restorePreInspectionLegacyTitle_(record) {
  // Automatic title-history reconstruction was unsafe: it wrote to every
  // Calendar copy before a stable original-title proof was always available.
  return {status:'PI_TITLE_HISTORY_READ_ONLY_REQUIRE_APPROVAL',
    eventId:record&&record.eventId||'',writes:0};
}

function tmv3_restoreRecordedPreInspectionLegacyTitles() {
  return {status:'RETIRED_BULK_TITLE_MUTATION_NO_WRITE',writes:0};
}

function tmv3_preInspectionCompletionVerified_(plans,execution) {
  if(!Array.isArray(plans)||plans.length!==1)return false;
  const p=plans[0];
  if(tmv3_preInspectionHistoricalCoverageVerified_(p))return true;
  if(!p||p.blocker||!Array.isArray(p.actions)||p.actions.length!==0||
     p.plan!=='NO_CHANGE'||p.readStatus!=='FRESH_TASK_GET'||
     !/^[1-9]\d*$/.test(String(p.taskId||'')))return false;
  // Cache-only/no-change plans must not be certified as fully converged.
  const checks=['customerCheck','locationCheck','requestedByCheck',
    'startCheck','endCheck','assignmentCheck','calendarTitleCheck','taskNameCheck'];
  return checks.every(function(key){return p[key]==='MATCH';});
}

function tmv3_preInspectionAutomaticRecovery_(reason) {
  if(!TMV3_PI_CLEANED_RELEASE_APPROVED)return {status:'PI_STAGING_OPERATOR_PATH_DISABLED',writes:0,emailsSent:0,guestInvites:0};
  return tmv3_piOriginal_preInspectionAutomaticRecovery_(reason);
}

function tmv3_piOriginal_preInspectionAutomaticRecovery_(reason) {
  const hour=Number(Utilities.formatDate(new Date(),TMV3_TIMEZONE,'H'));
  if(hour<8||hour>18) return {status:'OUTSIDE_PREINSPECTION_RECOVERY_HOURS',writes:0};
  const props=PropertiesService.getScriptProperties();
  let state=null;
  try {state=JSON.parse(props.getProperty('TMV3_PI_SWEEP_STATE')||'null');} catch(ignored) {}
  if(!state||state.status!=='IN_PROGRESS') return {status:'NO_STALLED_PREINSPECTION_SWEEP',writes:0};
  const last=Number(props.getProperty('TMV3_PI_LAST_CYCLE')||0);
  if(Date.now()-last<20*60000) return {status:'RECENT_PREINSPECTION_BATCH',writes:0};
  // Existing scheduled jobs provide recovery; no API executable or manual run needed.
  tmv3_preInspectionClearContinuationTrigger_();
  const result=tmv3_preInspectionAutomationCycle_({force:true,sweepContinue:true,reason:'AUTOMATIC_RECOVERY_'+reason});
  tmv3_preInspectionMaybeContinueSweep_(result);
  tmv3_preInspectionLogSweep_('RECOVERY',result);
  return result;
}

// Original: 60_Automation; Function: tmv3_preInspectionHistoricalCoverageVerified_
function tmv3_preInspectionHistoricalCoverageVerified_(plan) {
  if(!plan||plan.vertical!=='PreInspection'||
     plan.disposition!=='FULFILLED_NO_RECREATE'||
     plan.plan!=='VERIFY_ONLY_HISTORY_NO_WRITE'||
     plan.readStatus!=='FRESH_TASK_GET'||
     !/^[1-9]\d*$/.test(String(plan.taskId||''))||
     !Array.isArray(plan.actions)||plan.actions.length!==0||
     !tmv3_step6IsFulfilledStatus_(plan.taskStatus)||
     plan.customerCheck!=='MATCH'||
     String(plan.expectedCustomerId)!==String(plan.actualCustomerId))return false;
  // Completed Task by itself does not establish that THIS Calendar event was
  // fulfilled. A verified link or exact site/date agreement is also needed.
  const exactLink=String(plan.existingTaskId||plan.calendarTaskId||'')===String(plan.taskId);
  const sameSite=plan.locationCheck==='MATCH' && plan.calendarTitleCheck!=='CONFLICT';
  const day=tmv3_step6TaskTouchesEventDay_({start:plan.calendarStart},
    {Start:plan.taskStart,Due:plan.taskDue});
  return !!day && sameSite && (exactLink||plan.customerCheck==='MATCH');
}

function tmv3_reconcilePreInspectionPendingReview() {
  if(!TMV3_PI_CLEANED_RELEASE_APPROVED)return {status:'PI_STAGING_OPERATOR_PATH_DISABLED',writes:0,emailsSent:0,guestInvites:0};
  return tmv3_piOriginal_reconcilePreInspectionPendingReview();
}

function tmv3_piOriginal_reconcilePreInspectionPendingReview() {
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(1000)) return {status:'LOCK_BUSY',businessWrites:0};
  try {
    const props=PropertiesService.getScriptProperties();
    const stored=JSON.parse(props.getProperty('TMV3_PI_PENDING_EVENTS') || '[]');
    if (!Array.isArray(stored) || stored.some(function(id){return typeof id!=='string' || !id;})) {
      throw new Error('INVALID_PENDING_QUEUE_NO_CHANGE');
    }
    let pending=stored.slice();
    const results=[];
    const ids=stored.filter(function(id,index){return stored.indexOf(id)===index;}).slice(0,5);
    ids.forEach(function(id) {
      try {
        const plans=tmv3_step7FreshPlansForEvent_('PreInspection',id);
        if (plans.length===1 && tmv3_preInspectionHistoricalCoverageVerified_(plans[0])) {
          pending=pending.filter(function(candidate){return candidate!==id;});
          results.push({eventId:id,taskId:plans[0].taskId,status:'FULFILLED_NO_RECREATE',
            readbackVerified:true,businessWrites:0});
        } else {
          results.push({eventId:id,status:'REVIEW_RETAINED',plans:plans.map(function(p){
            return {plan:p.plan,blocker:p.blocker,taskId:p.taskId};
          }),businessWrites:0});
        }
      } catch(err) {results.push({eventId:id,status:'READ_FAILED_REVIEW_RETAINED',error:String(err),businessWrites:0});}
    });
    props.setProperty('TMV3_PI_PENDING_EVENTS',JSON.stringify(pending));
    const readback=JSON.parse(props.getProperty('TMV3_PI_PENDING_EVENTS') || '[]');
    if (JSON.stringify(readback)!==JSON.stringify(pending)) throw new Error('PENDING_QUEUE_READBACK_FAILED');
    const result={status:pending.length?'REVIEW_REMAINS':'COMPLETE',
      before:stored.length,remaining:pending.length,cleared:stored.length-pending.length,
      results:results,businessWrites:0};
    props.setProperty('TMV3_PI_LAST_PENDING_RECONCILIATION',JSON.stringify(result));
    tmv3_audit_('PreInspection','','','PENDING_HISTORY_RECONCILE',result.status,JSON.stringify(result));
    Logger.log(JSON.stringify(result));return result;
  } finally {lock.releaseLock();}
}

// Original: 90_Admin; Function: tmv3_preInspectionTodayRecords_
function tmv3_preInspectionTodayRecords_(targetDate) {
  const today=targetDate || Utilities.formatDate(new Date(),TMV3_TIMEZONE,'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(today)) throw new Error('INVALID_PI_TARGET_DATE');
  return tmv3_preInspectionIntakeRecords_().filter(function(record){
    return record.calendarRole==='PRIMARY_SHARED' &&
      Utilities.formatDate(new Date(record.start),TMV3_TIMEZONE,'yyyy-MM-dd')===today;
  });
}

// Original: 90_Admin; Function: tmv3_preInspectionCompactPlan_
function tmv3_preInspectionCompactPlan_(plan) {
  return {eventId:plan.eventId,taskId:plan.taskId,taskStatus:plan.taskStatus,plan:plan.plan,blocker:plan.blocker,
    customer:plan.customerCheck,location:plan.locationCheck,requestedBy:plan.requestedByCheck,
    start:plan.startCheck,end:plan.endCheck,assignment:plan.assignmentCheck,
    calendarTitle:plan.calendarTitleCheck,taskName:plan.taskNameCheck};
}

// Original: 90_Admin; Function: tmv3_preInspectionTodayPreview
function tmv3_preInspectionTodayPreview(targetDate) {
  const before=tmv3_strivenApiBudgetStatus_();
  if (before.exhausted || !Number.isFinite(before.remaining) || before.remaining<100) throw new Error('INSUFFICIENT_READBACK_HEADROOM');
  const records=tmv3_preInspectionTodayRecords_(targetDate);
  records.forEach(function(record){
    const plans=tmv3_step7FreshPlansForEvent_('PreInspection',record.eventId);
    Logger.log(JSON.stringify({title:record.title,plans:plans.map(tmv3_preInspectionCompactPlan_)}));
  });
  Logger.log(JSON.stringify({status:'PREVIEW_ONLY',events:records.length,before:before.count,
    after:tmv3_strivenApiBudgetStatus_().count}));
}

// Original: 90_Admin; Function: tmv3_preInspectionTodayEligibleQueue_
function tmv3_preInspectionTodayEligibleQueue_(records, sheetRows) {
  const byId={};
  (sheetRows||[]).forEach(function(row){
    const id=tmv3_clean_(row[9]); // PreInspection!O, from read of F:O.
    if (!id) return;
    if (byId[id]) { byId[id].duplicate=true;return; }
    byId[id]={evidence:String(row[0]||''),status:String(row[4]||''),duplicate:false};
  });
  const eligible=[],excluded=[];
  (records||[]).forEach(function(record){
    const entry=byId[record.eventId];
    const status=entry ? entry.status : '';
    const evidence=entry ? entry.evidence : '';
    const decision=/STEP\s*6\s*[\u2014-]\s*(CREATE_TASK|RECREATE_TASK|MATCH_EXISTING)\b/i.exec(status);
    let reason='';
    if (!entry) reason='NO_CACHED_STEP6_DECISION';
    else if (entry.duplicate) reason='DUPLICATE_OPERATOR_ROWS';
    else if (!decision) reason='STEP6_NOT_ELIGIBLE: '+status;
    else if (!/Customer:\s*#\d+\b/i.test(evidence)) reason='CUSTOMER_NOT_RESOLVED_IN_OPERATOR_VIEW';
    else if (!/Location ID:\s*\d+\b/i.test(evidence) && decision[1].toUpperCase()==='MATCH_EXISTING') reason='LOCATION_NOT_RESOLVED_IN_OPERATOR_VIEW';
    if (reason) {excluded.push({eventId:record.eventId,title:record.title,reason:reason});return;}
    const disposition=decision[1].toUpperCase();
    eligible.push({record:record,disposition:disposition,
      priority:disposition==='CREATE_TASK'?0:disposition==='RECREATE_TASK'?1:2});
  });
  eligible.sort(function(a,b){
    return a.priority-b.priority ||
      new Date(a.record.start).getTime()-new Date(b.record.start).getTime();
  });
  return {eligible:eligible,excluded:excluded};
}

// Original: 90_Admin; Function: tmv3_preInspectionTodayRunNext
function tmv3_preInspectionTodayRunNext() { return {status:'RETIRED_OPERATOR_SINGLE_EVENT_WRITE_USE_APPROVED_CYCLE',writes:0,emailsSent:0,guestInvites:0}; }

function tmv3_preInspectionTodayRunEligibleBatch() { return {status:'RETIRED_OPERATOR_BATCH_WRITE_USE_APPROVED_CYCLE',writes:0,emailsSent:0,guestInvites:0}; }

function tmv3_preInspectionTodayLocationEvidence() {
  const before=tmv3_strivenApiBudgetStatus_();
  if (before.exhausted || !Number.isFinite(before.remaining) || before.remaining<100) throw new Error('INSUFFICIENT_READBACK_HEADROOM');
  const results=tmv3_preInspectionTodayRecords_().map(function(record){
    return {eventId:record.eventId,title:record.title,evidence:tmv3_preInspectionOptionalOrderLocationEvidence_(record)};
  });
  Logger.log(JSON.stringify({results:results,before:before.count,afterCalls:tmv3_strivenApiBudgetStatus_().count}));
}

// Original: 90_Admin; Function: tmv3_preInspectionTodayRunUnlinked
function tmv3_preInspectionTodayRunUnlinked() { return {status:'RETIRED_UNLINKED_REPAIR_USE_APPROVED_CYCLE',writes:0,emailsSent:0,guestInvites:0}; }

function tmv3_preInspectionTodayLocationResponseShape() {
  const records=tmv3_preInspectionTodayRecords_();
  const rows=tmv3_sheet_(TMV3.SHEETS.PREINSPECTION).getDataRange().getDisplayValues().slice(1);
  records.forEach(function(record){
    const matches=rows.filter(function(row){return row[14]===record.eventId;});
    if(matches.length!==1 || matches[0][24]) return;
    const s2=tmv3_step2CalendarRecords_([record]);
    const s3=tmv3_step3BusinessAnchorRecords_(s2,tmv3_step3AnchorIndex_());
    const s4=tmv3_step4IdentityRecords_(s3,tmv3_step4IdentityIndex_());
    const id=Number(s4[0]&&s4[0].step4&&s4[0].step4.customer&&s4[0].step4.customer['Customer ID']);
    if(!Number.isInteger(id)||id<=0) return;
    const response=tmv3_fetchJson_(TMV3.API_BASE+'/v1/customers/'+id+'/locations',{method:'get'});
    Logger.log(JSON.stringify({eventId:record.eventId,customerId:id,keys:Object.keys(response||{}),responseArray:Array.isArray(response),count:response&&(response.TotalCount!==undefined?response.TotalCount:response.totalCount),dataType:response&&(response.Data===null||response.data===null?'null':Array.isArray(response.Data||response.data)?'array':typeof(response.Data||response.data)),length:response&&Array.isArray(response.Data||response.data)?(response.Data||response.data).length:null}));
  });
}

// Original: 90_Admin; Function: tmv3_preInspectionTodayRefreshTaskCacheSchedules
function tmv3_preInspectionTodayRefreshTaskCacheSchedules() { return {status:'RETIRED_CACHE_SCHEDULE_REWRITE',writes:0,emailsSent:0,guestInvites:0}; }

function tmv3_preInspectionOct21Preview() { return tmv3_preInspectionTodayPreview("2026-10-21"); }

// Original: 90_Admin; Function: tmv3_preInspectionOct21RunNext
function tmv3_preInspectionOct21RunNext() { return {status:'RETIRED_HISTORIC_OCT21_WRITE',writes:0,emailsSent:0,guestInvites:0}; }

function tmv3_preInspectionOct21RefreshCache() { return {status:'RETIRED_HISTORIC_OCT21_CACHE',writes:0,emailsSent:0,guestInvites:0}; }

function tmv3_preInspectionLastWeekAudit() {
  const before=tmv3_strivenApiBudgetStatus_();
  if(before.exhausted || !Number.isFinite(before.remaining) || before.remaining<100) throw new Error('INSUFFICIENT_READBACK_HEADROOM');
  const cfg=TMV3.VERTICALS.PreInspection;
  const source=tmv3_verticalCalendars_('PreInspection',cfg).filter(function(c){return c.role==='PRIMARY_SHARED';})[0];
  if(!source) throw new Error('AUTHORITATIVE_PI_SOURCE_REQUIRED');
  const calendar=CalendarApp.getCalendarById(source.calendarId);
  if(!calendar) throw new Error('AUTHORITATIVE_PI_CALENDAR_UNAVAILABLE');
  const records=calendar.getEvents(new Date('2026-10-12T00:00:00-04:00'),new Date('2026-10-19T00:00:00-04:00'))
    .map(function(event){return tmv3_calendarEvent_('PreInspection',cfg,source,event,{applyEligibility:false});})
    .filter(function(record){return record && !tmv3_step2PreInspectionNonCustomer_(record);});
  if(records.length>20) throw new Error('AUDIT_EVENT_CAP_EXCEEDED');
  const cache=tmv3_rows_(TMV3.SHEETS.TASKS);
  records.forEach(function(record){
    const taskId=tmv3_clean_(record.existingTaskId||record.taskNumber);
    const cached=cache.filter(function(task){return String(task['Task ID'])===taskId;});
    let plans=[],error='';
    try {plans=tmv3_step7FreshPlansForEvent_('PreInspection',record.eventId);}
    catch(err){error=String(err&&err.message||err);}
    Logger.log(JSON.stringify({eventId:record.eventId,title:record.title,start:tmv3_iso_(record.start),end:tmv3_iso_(record.end),
      sourceClassification:tmv3_step2ClassifyPreInspection_(record),linkedTaskId:taskId,cacheRows:cached.length,
      notificationAssessment:tmv3_preInspectionMissingDetailAssessment_(record),
      plans:plans.map(function(plan){return Object.assign(tmv3_preInspectionCompactPlan_(plan),{
        expectedStart:plan.expectedStart,taskStart:plan.taskStart,expectedDue:plan.expectedDue,taskDue:plan.taskDue,
        expectedRequestedById:plan.expectedRequestedById,expectedRequestedByType:plan.expectedRequestedByType,
        actualRequestedById:plan.actualRequestedById,actualRequestedByType:plan.actualRequestedByType,actions:plan.actions});}),error:error}));
  });
  Logger.log(JSON.stringify({status:'READ_ONLY_RETROSPECTIVE_AUDIT',from:'2026-10-12',untilExclusive:'2026-10-19',
    events:records.length,before:before.count,afterCalls:tmv3_strivenApiBudgetStatus_().count,
    strivenWrites:0,calendarWrites:0,notifications:0}));
}
// Preinspection authorization diagnostic: READ ONLY. Never sends mail or modifies Calendar, Striven, or Sheets.
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
  const handlers = ScriptApp.getProjectTriggers().map(function(t) {
    return t.getHandlerFunction();
  }).filter(function(name) {
    return /^tmv3_preInspectionCheck(0800|1300|1700)$/.test(name);
  });
  const result = {
    status: 'CALENDAR_MAIL_AUTHORIZATION_CHECK_PASSED',
    calendarName: calendar.getName(),
    calendarReadable: true,
    mailQuotaReadable: Number.isInteger(mailQuota),
    mailQuotaRecipientsRemaining: mailQuota,
    currentUserPreinspectionTriggerHandlers: handlers.sort(),
    calendarWrites: 0,
    strivenCalls: 0,
    emailsSent: 0,
    triggerChanges: 0
  };
  if (!result.mailQuotaReadable) throw new Error('MAIL_QUOTA_NOT_INTEGER');
  Logger.log(JSON.stringify(result));
  return result;
}

/**
 * PREINSPECTION / INTAKE AUDIT — Nov 2026
 * Business owner: Preinspection vertical. No Striven reads/writes, no Calendar mutations,
 * no emails, no trigger installation. Snapshot only; operator decisions persist by event+start.
 * Source of truth: CF Preinspects. Representative calendars are discovery-only.
 * Do not automatically copy or promote private-calendar events into the shared Calendar.
 */
function tmv3_piCalendarReviewRefresh_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1200)) return {status:'PI_REVIEW_LOCK_BUSY',calendarWrites:0,strivenCalls:0};
  try {
    const spreadsheet = SpreadsheetApp.openById(TMV3_SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('PI Calendar Review');
    if (!sheet) throw new Error('PI_REVIEW_SHEET_NOT_FOUND_NO_WRITE');
    const timezone = 'America/Toronto';
    const today = new Date(); today.setHours(0,0,0,0);
    const until = new Date(today.getTime()); until.setDate(until.getDate()+31);
    const checkedDate = Utilities.formatDate(today,timezone,'yyyy-MM-dd');
    const calendarSources = [
      {name:'CF Preinspects',id:TMV3.VERTICALS.PreInspection.primaryCalendarId},
      {name:'Stephen',id:'classicfireplace.ca_c20qcqfhvjbv784asn9pvuiaf4@group.calendar.google.com'},
      {name:'Jay',id:'jay@classicfireplace.ca'},
      {name:'Karen',id:'karen@classicfireplace.ca'}
    ];
    const copiesByKey = {};
    const sourceCounts = {};
    const coverageErrors = [];
    calendarSources.forEach(function(source) {
      let calendar=null,events=null;
      try {
        calendar=CalendarApp.getCalendarById(source.id);
        if(calendar) events=calendar.getEvents(today,until);
      }catch(error){
        coverageErrors.push({source:source.name,reason:'CALENDAR_ACCESS_DENIED'});
      }
      if (!calendar || !events) {
        if (!coverageErrors.some(function(x){return x.source===source.name;}))
          coverageErrors.push({source:source.name,reason:'CALENDAR_UNAVAILABLE'});
        if(source.name==='CF Preinspects') throw new Error('PI_PRIMARY_CALENDAR_REQUIRED_NO_WRITE');
        sourceCounts[source.name]=null;
        return;
      }
      if(events.length>750) throw new Error('PI_CALENDAR_EVENT_LIMIT_REVIEW_NO_WRITE:'+source.name);
      sourceCounts[source.name] = events.length;
      events.forEach(function(event) {
        if(event.isAllDayEvent() && event.getStartTime()<today) return;
        const rawId = String(event.getId()||'');
        const id = rawId.replace(/@google\.com$/i,'');
        const start = event.getStartTime();
        const key = id+'|'+start.getTime();
        const record = {
          id:id,
          rawId:rawId,
          start:start.toISOString(),
          end:event.getEndTime().toISOString(),
          title:String(event.getTitle()||''),
          location:String(event.getLocation()||''),
          description:String(event.getDescription()||''),
          creators:event.getCreators().map(String),
          calendarId:source.id,
          calendarName:source.name
        };
        if (!copiesByKey[key]) copiesByKey[key]={};
        if (copiesByKey[key][source.name]) throw new Error('PI_DUPLICATE_SOURCE_EVENT_NO_WRITE:'+key);
        copiesByKey[key][source.name]=record;
      });
    });
    const old = sheet.getLastRow()>3 ? sheet.getRange(4,1,Math.min(sheet.getLastRow()-3,1500),20).getValues() : [];
    const oldByKey = {};
    old.forEach(function(row){
      const id=String(row[8]||'');
      const start=String(row[1]||'');
      if (!id || !start) return;
      const key=id+'|'+start;
      if (!oldByKey[key]) oldByKey[key]={decision:String(row[18]||''),notes:String(row[19]||''),row:row};
    });
    const headers = [
      'Status','Start (Toronto)','End (Toronto)','Event / Customer','Calendar Copies',
      'Organizer / Origin','Location','Customer #','Google Event ID','Shared CF?',
      'Stephen Copy','Jay Copy','Karen Copy','Task Link?','Review Reason',
      'Original Notes Preview','Calendar URL','Checked Date','Operator Decision','Operator Notes'
    ];
    const rows = [];
    const norm=function(s){return String(s||'').replace(/\s+/g,' ').trim().toLowerCase();};
    const statusOrder={REVIEW_MIRROR_MISMATCH:0,SHARED_UNLINKED:1,SHARED_LINKED:2,
      REVIEW_POSSIBLE_PI:3,REVIEW_UNCLEAR_VISIT:4,REVIEW_IDENTITY_MISSING:5,EXCLUDED_SERVICE:6,
      EXCLUDED_OTHER_WORKFLOW:7,EXCLUDED_NON_CUSTOMER:8,EXCLUDED_UNVERIFIED:9};
    Object.keys(copiesByKey).forEach(function(key){
      const copies=copiesByKey[key];
      const shared=copies['CF Preinspects'];
      const ev=shared||copies.Stephen||copies.Karen||copies.Jay;
      const creators=(ev.creators||[]).join(', ');
      const knownOther=(ev.creators||[]).some(function(email){return /(?:vpg05s56lp9f7kaegsiqask34o|2pbtbkjl1g2sos1fspgfq029h0|2h90kflud3re121hga9bvsi33o|service@classicfireplace\.ca|steve\.jensen83@gmail\.com)/i.test(email);});
      const internal=/do not book|newmarket|team meeting|flash|rylander|layout meeting|follow up appointment|meet with steve|vacation|personal|holiday|golf club/i.test(ev.title);
      const serviceOnly=/\bservice\b/i.test(ev.title)&&!/\bpre[\s-]?inspect/i.test(ev.title+' '+ev.description);
      const ownerVisit=Boolean(copies.Stephen||copies.Jay||copies.Karen);
      const visitClue=/\bpre[\s-]?inspect|\bon[\s-]?site\b|\bsite visit\b|\bmeasure\b|\binspect\b/i.test(ev.title+' '+ev.description);
      const productClue=Boolean(ev.location)&&/\b(fireplace|insert|gas unit|stove|mantle|chimney|remove|replace)\b/i.test(ev.title+' '+ev.description);
      const customerClue=/(?:\b\d{5}\b|(?:\d{3}[-. ]?\d{3}[-. ]?\d{4}))/i.test(ev.title+' '+ev.description);
      const linked=/Tasks\/TaskInfo\.aspx\?TaskID|Task\s*#\s*\d+/i.test(ev.description);
      const number=ev.title.match(/(?:^|#)\s*(\d{5})\b/);
      let status='',reason='',mirror='NOT_EXPECTED';
      if(shared){
        const st=copies.Stephen;
        const exact=st&&['title','location','description','start','end'].every(function(field){return String(shared[field]||'')===String(st[field]||'');});
        mirror=st?(exact?'EXACT':'MISMATCH'):
          coverageErrors.some(function(x){return x.source==='Stephen';})?'ACCESS_UNVERIFIED':'NOT_PRESENT';
        status=mirror==='MISMATCH'?'REVIEW_MIRROR_MISMATCH':
          mirror==='ACCESS_UNVERIFIED'?'REVIEW_MIRROR_ACCESS':
          linked?'SHARED_LINKED':'SHARED_UNLINKED';
        reason=linked?'Task link is present; verify Striven before marking fully complete':'Task link absent; search existing Striven Task before any create';
      }else if(knownOther){
        status='EXCLUDED_OTHER_WORKFLOW';
        reason='Creator identifies Installation/Delivery/Service or contractor workflow; no automatic promotion';
      }else if(internal){status='EXCLUDED_NON_CUSTOMER';reason='Internal block/meeting/non-customer';}
      else if(serviceOnly){status='EXCLUDED_SERVICE';reason='Service-specific appointment';}
      else if(ownerVisit&&(visitClue||productClue)){
        status='REVIEW_POSSIBLE_PI';reason='Possible on-site preinspection omitted from shared Calendar; confirm visit type and ownership';
      }else if(ownerVisit&&customerClue){status='REVIEW_UNCLEAR_VISIT';reason='Customer/sales visit with no confirmed preinspection intention or adequate location';}
      else if(ownerVisit && (ev.creators||[]).some(function(email){
        const id=String(email||'').toLowerCase();
        return (copies.Stephen && (id==='classicfireplace.ca_c20qcqfhvjbv784asn9pvuiaf4@group.calendar.google.com'||id==='stephen@classicfireplace.ca')) ||
          (copies.Jay && id==='jay@classicfireplace.ca') ||
          (copies.Karen && id==='karen@classicfireplace.ca');
      })){
        status='REVIEW_IDENTITY_MISSING';reason='Owner-created appointment missing evidence to classify; verify identity, location and visit type, do not silently exclude';
      }
      else{status='EXCLUDED_UNVERIFIED';reason='Insufficient evidence to treat as Preinspection';}
      const stamp=function(iso){return Utilities.formatDate(new Date(iso),timezone,"yyyy-MM-dd'T'HH:mm:ssXXX");};
      const startText=stamp(ev.start),endText=stamp(ev.end);
      const orig=oldByKey[ev.id+'|'+startText];
      const link='https://www.google.com/calendar/event?eid='+Utilities.base64EncodeWebSafe(ev.id+' '+ev.calendarId).replace(/=+$/,'');
      rows.push([status,startText,endText,ev.title,Object.keys(copies).join(', '),creators,
        ev.location,number?number[1]:'',ev.id,shared?'YES':'NO',shared?mirror:copies.Stephen?'PRIVATE':'—',
        copies.Jay?'PRESENT':'—',copies.Karen?'PRESENT':'—',linked?'YES':'NO',reason,
        norm(ev.description).slice(0,330),link,checkedDate,orig?orig.decision:'',orig?orig.notes:'']);
    });
    // Preserve last-known rows for inaccessible secondary calendars. Never report
    // absence as a fact when an access failure prevented the scan.
    Object.keys(oldByKey).forEach(function(key){
      const prev=oldByKey[key];
      const sourceNames=String(prev.row[4]||'');
      if(!coverageErrors.some(function(x){return sourceNames.split(', ').indexOf(x.source)>=0;}))return;
      if(rows.some(function(r){return r[8]+'|'+r[1]===key;}))return;
      const preserved=prev.row.slice(0,20);
      preserved[0]='SOURCE_ACCESS_UNVERIFIED';
      preserved[14]='Calendar source access missing; last-known event retained, no auto-promotion or Task creation';
      rows.push(preserved);
    });
    // Keep operator decisions for events that moved beyond the review horizon; never erase human decisions.
    Object.keys(oldByKey).forEach(function(key){
      const prev=oldByKey[key];
      if(!prev.decision && !prev.notes)return;
      const already=rows.some(function(row){return row[8]+'|'+row[1]===key;});
      if(already)return;
      const preserved=prev.row.slice(0,20);
      preserved[0]='OUT_OF_WINDOW_REVIEW';
      preserved[14]='Retained due to operator decision; event not in current 31-day Calendar window';
      rows.push(preserved);
    });
    rows.sort(function(a,b){
      const aa=typeof statusOrder[a[0]]==='number'?statusOrder[a[0]]:10;
      const bb=typeof statusOrder[b[0]]==='number'?statusOrder[b[0]]:10;
      return aa-bb||String(a[1]).localeCompare(String(b[1]));
    });
    const statuses={};rows.forEach(function(r){statuses[r[0]]=(statuses[r[0]]||0)+1;});
    const summary='Calendar-only review '+checkedDate+' | '+sourceCounts['CF Preinspects']+' shared | '+
      (statuses.REVIEW_POSSIBLE_PI||0)+' possible | '+(statuses.REVIEW_UNCLEAR_VISIT||0)+' unclear | '+
      (coverageErrors.length?'PARTIAL COVERAGE: '+coverageErrors.map(function(x){return x.source;}).join(', '):'all sources accessible')+
      ' | zero Striven writes';
    const report=[
      ['PI CALENDAR INTAKE REVIEW — READ ONLY'].concat(Array(19).fill('')),
      [summary].concat(Array(19).fill('')),
      headers
    ].concat(rows);
    const lastRow=sheet.getLastRow();
    if(report.length>sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(),report.length-sheet.getMaxRows());
    let changed=false;
    // A1:T1 and A2:T2 are merged. Set only the anchor cell to protect the merge.
    if(sheet.getRange(1,1).getDisplayValue()!==report[0][0]){sheet.getRange(1,1).setValue(report[0][0]);changed=true;}
    if(sheet.getRange(2,1).getDisplayValue()!==report[1][0]){sheet.getRange(2,1).setValue(report[1][0]);changed=true;}
    const headerNow=sheet.getRange(3,1,1,20).getDisplayValues()[0];
    if(!headerNow.every(function(v,i){return String(v||'')===String(headers[i]||'');})){
      sheet.getRange(3,1,1,20).setValues([headers]);changed=true;
    }
    if(rows.length){
      const bodyNow=sheet.getRange(4,1,rows.length,20).getDisplayValues();
      const bodyMatches=bodyNow.every(function(row,i){return row.every(function(v,j){return String(v||'')===String(rows[i][j]||'');});});
      if(!bodyMatches){sheet.getRange(4,1,rows.length,20).setValues(rows);changed=true;}
    }
    if(lastRow>report.length){sheet.getRange(report.length+1,1,lastRow-report.length,20).clearContent();changed=true;}
    if(changed) SpreadsheetApp.flush();
    // Row-by-row independent event keys make the publication idempotent.
    if(sheet.getRange(3,1).getDisplayValue()!=='Status'||sheet.getRange(1,1).getDisplayValue()!=='PI CALENDAR INTAKE REVIEW — READ ONLY')
      throw new Error('PI_REVIEW_SHEET_READBACK_FAILED');
    return {status:coverageErrors.length?'PI_REVIEW_PARTIAL_SOURCE_ACCESS':
        (changed?'PI_REVIEW_REFRESHED':'PI_REVIEW_UNCHANGED'),
      calendarCounts:sourceCounts,coverageErrors:coverageErrors,
      logicalEvents:rows.length,statuses:statuses,
      scope:'TODAY_FORWARD_31_DAYS',calendarWrites:0,strivenCalls:0,emailsSent:0};
  } finally {lock.releaseLock();}
}

/** Manual read/write-to-review-sheet-only entry point; NEVER alters Calendar/Striven. */
function tmv3_preInspectionCalendarReviewRefresh() {
  return tmv3_piCalendarReviewRefresh_();
}


/**
 * Google Apps Script — Nov 9 2026 targeted production repair.
 * Explicit 2-event scope, fresh Task history guard, one existing transaction per event,
 * no force recreation, stop on uncertain outcome. Remove after acceptance.
 */
function tmv3_preInspectionRepairNov9Stephen() { return {status:'RETIRED_HISTORIC_NOV9_WRITE',writes:0,emailsSent:0,guestInvites:0}; }

function tmv3_preInspectionRepairThreeAddedByOthersNov9() { return {status:'RETIRED_HISTORIC_NOV9_WRITE',writes:0,emailsSent:0,guestInvites:0}; }

function tmv3_preInspectionRefreshOperatorView_() {
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1500))return {status:'PI_OPERATOR_LOCK_BUSY',strivenCalls:0};
  try{
    const cfg=TMV3.VERTICALS.PreInspection;
    const source=tmv3_verticalCalendars_('PreInspection',cfg)[0];
    const calendar=CalendarApp.getCalendarById(cfg.primaryCalendarId);
    if(!calendar)throw new Error('PI_OPERATOR_PRIMARY_CALENDAR_NOT_ACCESSIBLE');
    const start=new Date();start.setHours(0,0,0,0);
    const end=new Date(start);end.setDate(end.getDate()+Math.max(1,Number(cfg.lookaheadDays||365)));
    const events=calendar.getEvents(start,end);
    const step1=events.map(function(event){
      return tmv3_calendarEvent_('PreInspection',cfg,source,event,{applyEligibility:false});
    }).filter(function(x){return x && !tmv3_step2PreInspectionNonCustomer_(x);});
    const step2=tmv3_step2CalendarRecords_(step1);
    const step3=tmv3_step3BusinessAnchorRecords_(step2,tmv3_step3AnchorIndex_());
    const step4=tmv3_step4IdentityRecords_(step3,tmv3_step4IdentityIndex_());
    const step5=tmv3_step5TaskRecords_(step4,tmv3_step5TaskIndex_());
    const step6=tmv3_step6DecisionRecords_(step5).filter(function(r){return r.vertical==='PreInspection';});
    step6.sort(function(a,b){return new Date(a.start)-new Date(b.start);});
    const sheet=tmv3_sheet_(cfg.sheet),oldLast=sheet.getLastRow();
    const rows=step6.map(tmv3_step6OperatorRow_);
    tmv3_replaceRows_(cfg.sheet,TMV3_STEP1_HEADERS.slice(),rows);
    tmv3_step1FormatOperatorView_(cfg.sheet,'PreInspection',step6,oldLast);
    const actual=tmv3_rows_(cfg.sheet),seen={};
    actual.forEach(function(row){
      const id=String(row['Event ID']||'');
      if(!id||seen[id])throw new Error('PI_OPERATOR_EVENT_ID_EMPTY_OR_DUPLICATE');
      seen[id]=true;
    });
    const expected=step6.map(function(r){return String(r.eventId||'');});
    if(actual.length!==expected.length ||
       expected.some(function(id){return !seen[id];}))
      throw new Error('PI_OPERATOR_ROW_READBACK_MISMATCH');
    const stamp=Utilities.formatDate(new Date(),TMV3_TIMEZONE,'yyyy-MM-dd HH:mm:ss');
    const result={status:'PI_OPERATOR_VIEW_REFRESHED',at:stamp,events:actual.length,
      mode:'CURRENT_SHARED_CALENDAR_PLUS_CACHED_STRIVEN_SOURCES',
      strivenCalls:0,calendarWrites:0,taskWrites:0};
    tmv3_audit_('PreInspection','','','PI_OPERATOR_QUEUE','PASS',JSON.stringify(result));
    return result;
  }finally{lock.releaseLock();}
}

/** Explicit manual read-only cache projection for operator acceptance. */
function tmv3_preInspectionRefreshOperatorView() {
  return tmv3_preInspectionRefreshOperatorView_();
}



// Google Apps Script: approved Terry #53480 no-invitation private-copy intake.
// This function NEVER creates or patches Striven Tasks, sends email or adds guests.
function tmv3_terryQuietCalendarIntake_(mode) {
  return {status:'PI_TERRY_MANUAL_INTAKE_ARCHIVED_NO_MUTATION',
    mode:mode,eventId:'5mafnqht16o2hrqv4232rofmm3@google.com',
    calendarWrites:0,guestInvites:0,emailsSent:0,strivenWrites:0};
}

function tmv3_previewTerryQuietCalendarIntake() {
  return {status:'PI_TERRY_REQUIRES_FRESH_READ_ONLY_CHECK',calendarWrites:0,
    taskWrites:0,emailsSent:0};
}

function tmv3_importTerryQuietCalendarIntake() { return {status:'REQUIRES_FRESH_OPERATOR_APPROVAL_SEPARATE_TOOL',writes:0,emailsSent:0,guestInvites:0}; }

function tmv3_piRendererFingerprint_() {
  const source = Function.prototype.toString.call(tmv3_preInspectionMissingDetailEmailPreview_);
  let hash = 2166136261;
  for (let i = 0; i < source.length; i++) {
    hash = Math.imul(hash ^ source.charCodeAt(i), 16777619) >>> 0;
  }
  return ('00000000' + hash.toString(16)).slice(-8);
}

function tmv3_piAssertLockedStandard_() {
  if(typeof TMV3_HARD_RULES==='undefined'||!TMV3_HARD_RULES ||
      typeof TMV3==='undefined'||!TMV3.VERTICALS||!TMV3.OPERATIONS)
    throw new Error('PI_STANDARD_DEPENDENCY_MISSING_NO_WRITE');
  const root=TMV3_HARD_RULES,r=root.PreInspection||{};
  const live=TMV3.VERTICALS.PreInspection||{};
  if(root.version!=='2026-11-10-r1'||r.lockedStandardId!=='2026-11-10-r1'||
     r.lockedStandardReference!=='docs/workflows/preinspection/STANDARD_2026_11_10.md')
    throw new Error('PI_STANDARD_ID_OR_REFERENCE_DRIFT_NO_WRITE');
  if(Number(r.taskTypeId)!==105||Number(live.defaultPoolId)!==8||
     r.orderRequired!==false||r.attachOrderToTask!==false||
     r.field854Policy!=='DO_NOT_MANAGE'||r.infoCustomFieldsAtCreate!=='NONE'||
     r.calendarNotesPolicy!=='CALENDAR_ONLY'||live.descriptionMustBeBlankAtCreate!==true)
    throw new Error('PI_TASK_HARD_RULE_DRIFT_NO_WRITE');
  if(r.ownerGuestPolicy!=='VERIFIED_EVENT_OWNER_ONLY')
    throw new Error('PI_GUEST_RULE_DRIFT_NO_WRITE');
  if(JSON.stringify((r.inspectorOwnedFields||[]).map(Number).sort(function(a,b){return a-b;}))!==
     JSON.stringify([852,853,854,860]))
    throw new Error('PI_INSPECTOR_OWNED_FIELDS_DRIFT_NO_WRITE');
  if(JSON.stringify((TMV3.OPERATIONS.preInspectionAutomationScheduleHours||[]))!==JSON.stringify([8,13,17])||
     TMV3.OPERATIONS.preInspectionAutomationScheduleTimezone!=='America/Toronto')
    throw new Error('PI_DAILY_SLOTS_DRIFT_NO_WRITE');
  const presentation=r.presentation||{};
  if(presentation.linkHeading!=='-----Striven Links-----'||
     (presentation.calendarTitleParts||[]).join('|')!=='customerNumber|customerName|phone'||
     (presentation.taskTitleParts||[]).join('|')!=='customerName|address|phone')
    throw new Error('PI_PRESENTATION_DRIFT_NO_WRITE');
  if(r.notificationTemplatePolicy!=='EXACT_EXISTING_RENDERER_NO_REDRAFT'||
     r.notificationRendererFingerprint!==tmv3_piRendererFingerprint_())
    throw new Error('PI_NOTIFICATION_TEMPLATE_DRIFT_NO_SEND');
  return {status:'PASS',standard:root.version,reference:r.lockedStandardReference,
    requestedByTypeContract:tmv3_piRequestedByContract_()};
}

/** Canonical rule confirmed by operator: RequestedBy.Type = employee.
 * The Employee ID must equal the verified owning Calendar employee. An
 * explicitly conflicting shared configuration remains a deployment blocker. */
/** Pure contract validation; never mutate the shared frozen policy in tests. */
function tmv3_piRequestedByContractFromRule_(rule) {
  const configured=String(rule && rule.requestedByType || '').trim().toLowerCase();
  if(configured && configured!=='employee')
    throw new Error('PI_REQUESTED_BY_SHARED_STANDARD_CONFLICT_NO_WRITE');
  return 'employee';
}

function tmv3_piRequestedByContract_() {
  const rule=(typeof TMV3_HARD_RULES!=='undefined' && TMV3_HARD_RULES &&
    TMV3_HARD_RULES.PreInspection) || {};
  return tmv3_piRequestedByContractFromRule_(rule);
}

/** Shape only: a positive integer Employee ID; provenance checked at Step 7. */
function tmv3_piAssertRequestedByEmployeeShape_(requestedBy) {
  tmv3_piRequestedByContract_();
  const type=String(requestedBy && (requestedBy.Type||requestedBy.type)||'').trim().toLowerCase();
  const rawId=requestedBy && (requestedBy.Id!==undefined?requestedBy.Id:requestedBy.id);
  const id=Number(rawId);
  if(type!=='employee'||!/^[1-9]\d*$/.test(String(rawId||''))||
      !Number.isSafeInteger(id)||id<=0)
    throw new Error('PI_REQUESTED_BY_VERIFIED_EMPLOYEE_REQUIRED_NO_WRITE');
  return true;
}

/** Plan requester must equal the employee mapped from this event's verified
 * owning Calendar. A customer Contact or arbitrary salesperson is not proof. */
function tmv3_piAssertRequestedByPlan_(plan, record) {
  tmv3_piRequestedByContract_();
  const owner=tmv3_piVerifiedOwnerEmployee_(record);
  if(owner.status!=='MATCHED'||!Number.isSafeInteger(owner.employeeId))
    throw new Error(owner.blocker||'PI_REQUESTED_BY_OWNER_UNVERIFIED_NO_WRITE');
  const type=String(plan&&plan.expectedRequestedByType||'').trim().toLowerCase();
  const rawId=plan&&plan.expectedRequestedById;
  const id=Number(rawId);
  if(type!=='employee'||!/^[1-9]\d*$/.test(String(rawId||''))||
      !Number.isSafeInteger(id)||id<=0)
    throw new Error('PI_EXPECTED_REQUESTED_BY_EMPLOYEE_NOT_PROVEN_NO_WRITE');
  if(id!==owner.employeeId)
    throw new Error('PI_REQUESTED_BY_MUST_MATCH_VERIFIED_EVENT_OWNER_NO_WRITE');
  return true;
}

function tmv3_piCreatorEmailExcluded_(creatorEmail) {
  const creator = String(creatorEmail || '').trim().toLowerCase();
  const blocked = TMV3_HARD_RULES.PreInspection.notificationExcludedCreatorEmails || [];
  return !!creator && blocked.indexOf(creator) >= 0;
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

function tmv3_piAssertSafeTaskPatch_(payload) {
  tmv3_piAssertLockedStandard_();
  const p=payload||{};
  const protectedNames=/^(?:InfoCustomFields|CustomFields|Description|SalesOrder|SalesOrderID|SalesOrderId|SOId|Order|OrderId|Type)$/i;
  Object.keys(p).forEach(function(key){
    if(protectedNames.test(key))throw new Error('PI_PROTECTED_TASK_FIELD_NO_WRITE:'+key);
    if(key.toLowerCase()==='requestedby' && key!=='RequestedBy')
      throw new Error('PI_REQUESTED_BY_FIELD_CASE_INVALID_NO_WRITE');
  });
  if(p.RequestedBy!==undefined)tmv3_piAssertRequestedByEmployeeShape_(p.RequestedBy);
  if(p.Title!==undefined)tmv3_piAssertTaskTitle_(p.Title);
  return true;
}

function tmv3_preInspectionLockedStandardStatus() {
  return tmv3_piAssertLockedStandard_();
}


// Only positively verified owner -> employee mappings may authorize writes.
// Jay and Karen calendar IDs are not yet certified in this staging source.
/** Resolve Requested By strictly from a known owning calendar identity.
 * Calendar source array order MUST NEVER confer employee identity.
 * Stephen/Jay are grounded in the approved employee table. Karen remains REVIEW.
 * An alternate Stephen calendar ID is accepted only if the live configuration
 * explicitly associates a *single* secondary calendar with Stephen's email. */
function tmv3_piVerifiedOwnerEmployee_(record) {
  const cfg=(typeof TMV3!=='undefined' && TMV3.VERTICALS &&
    TMV3.VERTICALS.PreInspection)||{};
  const owner=String(record && record.organizerCalendarId || '').trim().toLowerCase();
  const stephenEmail='stephen@classicfireplace.ca';
  const jayEmail='jay@classicfireplace.ca';
  const karenEmail='karen@classicfireplace.ca';
  const aliases=(cfg.secondaryCalendarIds||[]).map(function(v){
    return String(v||'').trim().toLowerCase();
  }).filter(Boolean);
  const primaryId=String(cfg.primaryCalendarId||'').trim().toLowerCase();
  const configuredStephen=String(cfg.secondaryOwnerEmail||'').trim().toLowerCase();
  if(!owner)return {status:'REVIEW',employeeId:null,blocker:'PI_OWNER_MISSING_NO_WRITE'};
  if(owner===stephenEmail)
    return {status:'MATCHED',employeeId:15,employeeName:'Stephen Foley',evidence:'EXACT_VERIFIED_OWNER_EMAIL'};
  if(owner===jayEmail)
    return {status:'MATCHED',employeeId:1,employeeName:'Jay Scott',evidence:'EXACT_VERIFIED_OWNER_EMAIL'};
  if(owner===karenEmail)
    return {status:'REVIEW',employeeId:null,blocker:'PI_KAREN_EMPLOYEE_ID_UNVERIFIED_NO_WRITE'};
  // Never assign an employee by guessing which entry of a multi-calendar list is theirs.
  if(aliases.length===1 && configuredStephen===stephenEmail &&
     aliases[0]===owner && owner!==primaryId &&
     owner!==jayEmail && owner!==karenEmail) {
    return {status:'MATCHED',employeeId:15,employeeName:'Stephen Foley',
      evidence:'SINGLE_SECONDARY_CALENDAR_EXPLICIT_STEPHEN_OWNER'};
  }
  return {status:'REVIEW',employeeId:null,
    blocker:'PI_OWNER_EMPLOYEE_MAPPING_UNVERIFIED_NO_WRITE'};
}


function tmv3_preInspectionCleanedReleaseStatus() {
  return {revision:'2026-11-11-r5',releaseApproved:TMV3_PI_CLEANED_RELEASE_APPROVED,
    productionWritesEnabledByFile:TMV3_PI_CLEANED_RELEASE_APPROVED,
    outstanding:['VERIFY_SHARED_STEP7_EMITS_OWNER_EMPLOYEE_REQUESTER','VERIFY_OWNER_TO_EMPLOYEE_MAPPING_KAREN',
      'RUN_FULL_V3_SHARED_DEPENDENCIES_AND_LIVE_READBACK']};
}

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
  const failed=results.filter(function(x){return x.status==='FAIL';});
  const out={status:failed.length?'ATTENTION_UNRESOLVED_RULES':'SYNTHETIC_SELF_TEST_PASS',
    revision:'2026-11-11-r5',checks:results.length,passed:results.length-failed.length,
    failures:failed,results:results,calendarWrites:0,strivenCalls:0,emailsSent:0,
    evidence:'PURE_FUNCTIONS_AND_CONFIG_ONLY_NOT_LIVE_READBACK'};
  Logger.log('PI_CLEANED_SELF_TEST='+JSON.stringify(out));
  return out;
}








/**
 * GOOGLE APPS SCRIPT — CF PreInspection V3 r6 targeted live readback (time-identity fix).
 * Add as ONE new file; do NOT replace shared functions or flip release flags.
 * No Task POST/PATCH, Calendar writes, guest mutations, notifications,
 * Sheets writes, or Script Properties writes from the functions in this file.
 * Invokes existing V3 Calendar/source and Striven GET/search helpers.
 * Date and event IDs intentionally scoped to the two authorized events.
 */
function tmv3_piNov11TerryJakobReadOnly() {
  const targets = [
    {customerNumber:'53480',customerId:53480,name:'Terry Beettam',
     eventId:'5mafnqht16o2hrqv4232rofmm3',orderNumber:'586828',
     start:'2026-11-11T15:00:00-05:00'},
    {customerNumber:'62883',customerId:62883,name:'Jakob Pizzera',
     eventId:'0aa25j196lol5imushv6mrpmv1',orderNumber:'586702',
     start:'2026-11-11T17:00:00-05:00',preferredLocationId:58653,
     rejectedAlternativeLocationId:58654}
  ];
  const before=tmv3_strivenApiBudgetStatus_();
  if(before.exhausted || !Number.isFinite(before.remaining) || before.remaining<50) {
    const blocked={status:'API_BUDGET_INSUFFICIENT_NO_WRITES',before:before.count,
      remaining:before.remaining,events:[],calendarWrites:0,taskWrites:0,emailsSent:0};
    Logger.log('PI_NOV11_TARGET_READBACK='+JSON.stringify(blocked));return blocked;
  }
  let records;
  try {records=tmv3_preInspectionTodayRecords_('2026-11-11');}
  catch(error){
    const blocked={status:'CALENDAR_READ_FAILED_NO_WRITES',error:String(error),
      events:[],calendarWrites:0,taskWrites:0,emailsSent:0};
    Logger.log('PI_NOV11_TARGET_READBACK='+JSON.stringify(blocked));return blocked;
  }
  const events=[];
  targets.forEach(function(target){
    try {events.push(tmv3_piNov11ReadOnlyOne_(target,records));}
    catch(error){events.push({customerNumber:target.customerNumber,eventId:target.eventId,
      status:'REVIEW_NO_WRITES',blockers:['READBACK_EXCEPTION: '+String(error&&error.message||error)]});}
  });
  const output={revision:'2026-11-11-r6-calendar-identity-instant',
    status:events.every(function(e){return e.status==='EVIDENCE_READY_PENDING_STAGING_RELEASE';})
      ? 'EVIDENCE_READY_NO_WRITES':'REVIEW_REMAINS_NO_WRITES',
    events:events,apiBefore:before.count,apiAfter:tmv3_strivenApiBudgetStatus_().count,
    calendarWrites:0,taskWrites:0,emailsSent:0,
    nextAction:'Resolve blockers before guarded Step 5/6 replan. Do not bypass r5 release gate.'};
  Logger.log('PI_NOV11_TARGET_READBACK='+JSON.stringify(output));
  return output;
}

/**
 * Strict business-event identity check. Does not mutate Calendar or normalize
 * either event's stored timestamp; conversion is comparison-only.
 * The blockers separately identify a changed appointment versus a formatting
 * difference. Expects same exact instant (no fuzzy start-time tolerance).
 */
function tmv3_piNov11IdentityReadback_(record,target) {
  const cleanId=function(id){return String(id||'').trim().toLowerCase().replace(/@google\.com$/,'');};
  const actualStart=record && record.start;
  const expectedStart=target && target.start;
  const aMs=actualStart instanceof Date ? actualStart.getTime() : new Date(actualStart).getTime();
  const eMs=new Date(expectedStart).getTime();
  const startParseable=Number.isFinite(aMs) && Number.isFinite(eMs);
  const checks={
    eventId:cleanId(record && record.eventId)===cleanId(target && target.eventId),
    calendarRole:String(record && record.calendarRole||'')==='PRIMARY_SHARED',
    customerNumber:String(record && record.customerNumber||'').trim()===String(target && target.customerNumber||'').trim(),
    startParseable:startParseable,
    startInstant:startParseable && aMs===eMs,
    active:String(record && record.status||'').trim().toLowerCase()!=='cancelled'
  };
  const failures=Object.keys(checks).filter(function(k){return !checks[k];});
  const map={eventId:'CALENDAR_EVENT_ID_CHANGED',calendarRole:'CALENDAR_NOT_PRIMARY_SHARED',
    customerNumber:'CALENDAR_CUSTOMER_NUMBER_MISMATCH',
    startParseable:'CALENDAR_START_UNREADABLE',startInstant:'CALENDAR_START_INSTANT_CHANGED',
    active:'CALENDAR_EVENT_CANCELLED'};
  // Avoid reporting both missing parse and a start mismatch for one invalid date.
  const blockers=failures.filter(function(k){return k!=='startInstant'||startParseable;}).map(function(k){return map[k];});
  return {ok:blockers.length===0,blockers:blockers,evidence:{checks:checks,
    expectedStart:String(expectedStart||''),
    actualStartType:Object.prototype.toString.call(actualStart),
    actualStart:String(actualStart||''),
    actualStartInstant:Number.isFinite(aMs)?new Date(aMs).toISOString():null,
    expectedStartInstant:Number.isFinite(eMs)?new Date(eMs).toISOString():null,
    actualCalendarRole:String(record && record.calendarRole||''),
    actualCustomerNumber:String(record && record.customerNumber||''),
    actualEventId:String(record && record.eventId||'')}};
}

function tmv3_piNov11ReadOnlyOne_(target,records) {
  const events=(records||[]).filter(function(r){
    return String(r.eventId||'').replace(/@google\.com$/i,'')===target.eventId;
  });
  const blockers=[];
  if(events.length!==1) return {customerNumber:target.customerNumber,eventId:target.eventId,
    status:'REVIEW_NO_WRITES',blockers:['AUTHORITATIVE_CALENDAR_EVENT_NOT_UNIQUE']};
  const record=events[0];
  // The shared V3 Calendar mapper may provide start as a Date, ISO UTC,
  // or a timestamp with a Toronto offset. Compare the *instant*, not
  // String(Date) to an ISO literal. Keep identity/source fail-closed.
  const identity=tmv3_piNov11IdentityReadback_(record,target);
  if(!identity.ok) {
    return {customerNumber:target.customerNumber,eventId:target.eventId,
      status:'REVIEW_NO_WRITES',blockers:identity.blockers,
      calendarIdentityReadback:identity.evidence};
  }
  const owner=tmv3_piVerifiedOwnerEmployee_(record);
  if(owner.status!=='MATCHED')blockers.push(owner.blocker||'OWNER_EMPLOYEE_UNVERIFIED');
  const hold=tmv3_preInspectionApprovalHold_(record.eventId);
  if(hold)blockers.push('OPERATOR_APPROVAL_HOLD_ACTIVE');
  if(String(record.orderNumber||'')!==target.orderNumber)blockers.push('CALENDAR_SO_NUMBER_MISMATCH');
  const order=tmv3_piNov11OrderReadOnly_(target);
  if(order.blocker)blockers.push(order.blocker);
  else {
    if(['approved','in progress'].indexOf(order.status)>=0)blockers.push('ORDER_IS_INSTALL_OR_IN_PROGRESS_NOT_PI');
    else if(['quoted','incomplete','pending approval'].indexOf(order.status)<0)
      blockers.push('ORDER_STATUS_NOT_AUTHORIZED_FOR_PI');
  }
  const locations=tmv3_preInspectionCustomerLocationsFresh_(target.customerId);
  const matching=locations.filter(function(l){
    return Number(l['Customer ID'])===target.customerId &&
      tmv3_addressStrongMatch_(tmv3_addressParts_(record.location),
        tmv3_addressParts_(tmv3_locationFullAddress_(l)));
  });
  if(!matching.length)blockers.push('NO_FRESH_CUSTOMER_OWNED_ADDRESS_MATCH');
  const matchedIds=matching.map(function(l){return Number(l['Location ID']);}).sort(function(a,b){return a-b;});
  if(!order.blocker) {
    if(matchedIds.indexOf(order.shipToLocationId)<0)
      blockers.push('SO_SHIP_TO_NOT_A_MATCHING_CUSTOMER_OWNED_LOCATION');
    if(target.preferredLocationId && order.shipToLocationId!==target.preferredLocationId)
      blockers.push('JAKOB_SO_SHIP_TO_DOES_NOT_EQUAL_OPERATOR_APPROVED_58653');
  }
  if(matchedIds.length>1 && !target.preferredLocationId)
    blockers.push('MULTIPLE_CUSTOMER_OWNED_LOCATION_MATCHES');
  if(target.preferredLocationId && matchedIds.indexOf(target.preferredLocationId)<0)
    blockers.push('JAKOB_APPROVED_LOCATION_NOT_FRESHLY_VERIFIED');
  const customer={'Customer ID':String(target.customerId),'Customer Number':target.customerNumber,'Name':target.name};
  const taskIds=tmv3_searchPreInspectionTaskIdsForCustomer_(customer);
  if(taskIds.length>10)blockers.push('TASK_HISTORY_OVER_TEN_BLOCKED_BY_EXISTING_CREATE_GUARD');
  // Existing V3 create guard rejects >10 IDs. Avoid potentially hundreds
  // of unnecessary Task GETs in that case. Never treat truncated GETs as clear.
  const tasks=taskIds.length>10 ? [] : taskIds.map(function(id){return tmv3_getTaskById_(id);});
  if(tasks.some(function(t){return !t || Number(t['Task Type ID'])!==105 ||
     Number(t['Customer ID'])!==target.customerId || !/^\d+$/.test(String(t['Task ID']||''));}))
    blockers.push('FRESH_TASK_HISTORY_OWNERSHIP_OR_TYPE_CONFLICT');
  const opens=tasks.filter(function(t){return t&&tmv3_taskIsOpen_(t.Status);});
  const completed=tasks.filter(function(t){return t&&tmv3_taskIsCompleted_(t.Status);});
  if(opens.length)blockers.push('OPEN_TYPE105_TASK_EXISTS_REQUIRES_EVENT_SPECIFIC_STEP5_MATCH');
  const completedOnVisit=completed.filter(function(t){
    return t.Start && tmv3_sameLocalDay_(record.start,t.Start) &&
      (!t['Location ID'] || matchedIds.indexOf(Number(t['Location ID']))>=0);
  });
  if(completedOnVisit.length)blockers.push('POSSIBLE_COMPLETED_VISIT_TASK_DO_NOT_RECREATE');
  return {customerNumber:target.customerNumber,eventId:target.eventId,
    status:blockers.length?'REVIEW_NO_WRITES':'EVIDENCE_READY_PENDING_STAGING_RELEASE',
    blockers:blockers,requesterEmployeeId:owner.status==='MATCHED'?owner.employeeId:null,
    approvalHoldActive:!!hold,orderId:order.orderId||null,orderNumber:target.orderNumber,
    orderStatus:order.status||'',orderShipToLocationId:order.shipToLocationId||null,
    freshMatchingLocationIds:matchedIds,operatorPreferredLocationId:target.preferredLocationId||null,
    freshType105TaskCount:taskIds.length,freshType105TaskIds:taskIds.slice(0,10),
    openType105TaskIds:opens.map(function(t){return String(t['Task ID']);}),
    completedVisitTaskIds:completedOnVisit.map(function(t){return String(t['Task ID']);}),
    calendarIdentityReadback:identity.evidence,
    note:'No write performed. Step 5/6 must be re-planned through V3 after approved cache/source repair.'};
}

function tmv3_piNov11OrderReadOnly_(target) {
  const query={PageIndex:0,PageSize:2,Number:target.orderNumber,CustomerId:target.customerId};
  const search=tmv3_fetchJson_(TMV3.API_BASE+'/v1/sales-orders/search',
    {method:'post',contentType:'application/json',payload:JSON.stringify(query)});
  const rows=search&&(search.Data||search.data);
  const count=Number(search&&(search.TotalCount!==undefined?search.TotalCount:search.totalCount));
  if(!Array.isArray(rows)||count!==1||rows.length!==1)
    return {blocker:'SALES_ORDER_SEARCH_NOT_UNIQUE_OR_INCOMPLETE'};
  const row=rows[0],orderId=Number(tmv3_first_(row,['Id','id']));
  if(!Number.isSafeInteger(orderId)||orderId<=0 ||
     String(tmv3_first_(row,['Number','OrderNumber']))!==target.orderNumber)
    return {blocker:'SALES_ORDER_SEARCH_IDENTITY_CONFLICT'};
  const raw=tmv3_fetchJson_(TMV3.API_BASE+'/v1/sales-orders/'+orderId,{method:'get'});
  const customer=tmv3_first_(raw,['Customer','customer']);
  const location=tmv3_first_(raw,['ShipToLocation','shipToLocation']);
  const status=tmv3_first_(raw,['Status','status']);
  const customerId=Number(tmv3_first_(customer,['Id','id']));
  const locationId=Number(tmv3_first_(location,['Id','id']));
  const name=String(tmv3_first_(status,['Name','name'])||'').trim().toLowerCase();
  if(Number(tmv3_first_(raw,['Id','id']))!==orderId ||
     String(tmv3_first_(raw,['OrderNumber','Number']))!==target.orderNumber ||
     customerId!==target.customerId || !Number.isSafeInteger(locationId)||locationId<=0)
    return {blocker:'SALES_ORDER_LIVE_IDENTITY_OR_SHIPTO_UNVERIFIED'};
  return {orderId:orderId,shipToLocationId:locationId,status:name};
}

/**
 * Separately callable guarded write runner. Intentionally refuses ALL writes
 * until r5's established release flag AND production mode are authorized,
 * missing-details email/failure-alert switches are OFF, and BOTH preflights
 * return full evidence. Never toggles global safety configuration itself.
 * This does not replace the shared Step 7 checks/readbacks.
 */
function tmv3_piNov11TerryJakobRunIfReleased() {
  if(!TMV3_PI_CLEANED_RELEASE_APPROVED || TMV3.MODE!=='PRODUCTION_WRITE') {
    const result={status:'RELEASE_NOT_APPROVED_NO_WRITES',taskWrites:0,emailsSent:0,
      message:'Run tmv3_piNov11TerryJakobReadOnly first; keep release disabled until Step 7 is verified.'};
    Logger.log('PI_NOV11_TARGET_EXECUTION='+JSON.stringify(result));return result;
  }
  if(TMV3.OPERATIONS.preInspectionMissingDetailEmailsEnabled!==false ||
     TMV3.OPERATIONS.preInspectionFailureAlertsEnabled!==false ||
     tmv3_clean_(TMV3.OPERATIONS.preInspectionAutomationCanaryEventId)) {
    const result={status:'NOTIFICATION_OR_CANARY_CONFIG_UNSAFE_NO_WRITES',taskWrites:0,emailsSent:0};
    Logger.log('PI_NOV11_TARGET_EXECUTION='+JSON.stringify(result));return result;
  }
  const preflight=tmv3_piNov11TerryJakobReadOnly();
  if(preflight.status!=='EVIDENCE_READY_NO_WRITES' || preflight.events.length!==2) {
    const result={status:'PREFLIGHT_BLOCKED_NO_WRITES',taskWrites:0,emailsSent:0,
      preflightStatus:preflight.status,events:preflight.events};
    Logger.log('PI_NOV11_TARGET_EXECUTION='+JSON.stringify(result));return result;
  }
  const ids=['5mafnqht16o2hrqv4232rofmm3@google.com',
    '0aa25j196lol5imushv6mrpmv1@google.com'];
  const results=[];
  for(let i=0;i<ids.length;i++) {
    // Run exactly one target per guarded cycle; never broad queue sweeps.
    const result=tmv3_preInspectionAutomationCycle_({force:true,eventId:ids[i],
      reason:'OPERATOR_AUTHORIZED_TWO_EVENT_REPAIR_20261111'});
    results.push({eventId:ids[i],status:result.status,processed:result.processed,
      writes:result.writes||0,notifications:result.notifications||0,
      results:result.results||[],error:result.error||''});
    const plansVerified=Array.isArray(result.results) && result.results.some(function(one){
      return Array.isArray(one.afterPlans) && one.afterPlans.length===1 &&
        one.afterPlans[0].plan==='NO_CHANGE' && !one.afterPlans[0].blocker;
    });
    if((result.notifications||0)!==0 || result.status!=='COMPLETE' ||
       (result.writes||0)>1 || !plansVerified)break;
  }
  const output={status:results.length===2 && results.every(function(r){
      return r.status==='COMPLETE' && r.notifications===0;
    })?'CYCLES_FINISHED_REQUIRE_LIVE_READBACK':'ATTENTION_REVIEW_BEFORE_MORE_WRITES',
    results:results,notifications:results.reduce(function(n,r){return n+r.notifications;},0),
    note:'Read each returned afterPlans and fresh Task/Calendar links; do not infer success from cycle status alone.'};
  Logger.log('PI_NOV11_TARGET_EXECUTION='+JSON.stringify(output));return output;
}
