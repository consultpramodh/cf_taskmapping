// MIGRATION STAGING ONLY. Source: reconstructed V3 3.12.21. NOT APPROVED FOR DEPLOYMENT.

// Original: 20_Striven; Function: tmv3_preInspectionTaskRowArray_
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

  const prior = tmv3_rows_(
    TMV3.SHEETS.TASKS
  ).filter(function(task) {
    return Number(
      task['Task Type ID'] || 0
    ) === 105;
  });

  const props =
    PropertiesService.getScriptProperties();
  const refreshKey =
    'TMV3_PREINSPECTION_TASK_CACHE_REFRESH_AT';
  const lastMs = Number(
    props.getProperty(refreshKey) || 0
  );
  const ttlMinutes = Math.max(
    60,
    Number(
      TMV3.OPERATIONS &&
      TMV3.OPERATIONS.preInspectionTaskCacheMinMinutes ||
      360
    )
  );
  const ageMinutes = lastMs
    ? (
        new Date().getTime() -
        lastMs
      ) / 60000
    : Infinity;

  if (
    options.force !== true &&
    prior.length &&
    ageMinutes < ttlMinutes
  ) {
    return {
      rows:prior.map(
        tmv3_preInspectionTaskRowArray_
      ),
      reused:true,
      targetCustomers:0,
      searchedCustomers:0,
      hydratedTasks:0,
      priorTasks:prior.length,
      ageMinutes:Math.round(ageMinutes)
    };
  }

  const targets =
    tmv3_preInspectionTaskRefreshTargets_();

  if (!targets.length) {
    return {
      rows:prior.map(
        tmv3_preInspectionTaskRowArray_
      ),
      reused:true,
      targetCustomers:0,
      searchedCustomers:0,
      hydratedTasks:0,
      priorTasks:prior.length,
      ageMinutes:isFinite(ageMinutes)
        ? Math.round(ageMinutes)
        : null
    };
  }

  const customerById = {};
  const wantedTaskIds = {};

  targets.forEach(function(target) {
    const customer = target.customer || {};
    const customerId = tmv3_clean_(
      customer['Customer ID']
    );

    if (customerId) {
      customerById[customerId] = customer;
    }

    if (target.existingTaskId) {
      wantedTaskIds[
        target.existingTaskId
      ] = true;
    }
  });

  const customerIds =
    Object.keys(customerById);

  customerIds.forEach(function(customerId) {
    tmv3_searchPreInspectionTaskIdsForCustomer_(
      customerById[customerId]
    ).forEach(function(taskId) {
      wantedTaskIds[taskId] = true;
    });
  });

  const taskIds =
    Object.keys(wantedTaskIds);
  const maxHydrate = Math.max(
    1,
    Number(
      TMV3.OPERATIONS &&
      TMV3.OPERATIONS.preInspectionTaskHydrationMaxTasks ||
      60
    )
  );

  if (taskIds.length > maxHydrate) {
    throw new Error(
      'TMV3_PREINSPECTION_TASK_HYDRATION_CAP: ' +
      taskIds.length +
      ' Tasks exceed the configured cap of ' +
      maxHydrate +
      '.'
    );
  }

  const hydrated = taskIds
    .map(function(taskId) {
      return tmv3_getTaskById_(taskId);
    })
    .filter(function(task) {
      const customerId = tmv3_clean_(
        task &&
        task['Customer ID']
      );

      return (
        task &&
        Number(
          task['Task Type ID'] || 0
        ) === 105 &&
        !!customerById[customerId]
      );
    });

  const seen = {};
  hydrated.forEach(function(task) {
    const id = tmv3_clean_(
      task['Task ID']
    );
    if (!id) {
      throw new Error(
        'TMV3_PREINSPECTION_TASK_HYDRATION_INVALID: hydrated Task is missing Task ID.'
      );
    }
    if (seen[id]) {
      throw new Error(
        'TMV3_PREINSPECTION_TASK_HYDRATION_DUPLICATE: Task ' +
        id +
        ' was hydrated more than once.'
      );
    }
    seen[id] = true;
  });

  props.setProperty(
    refreshKey,
    String(new Date().getTime())
  );

  return {
    rows:hydrated.map(
      tmv3_preInspectionTaskRowArray_
    ),
    reused:false,
    targetCustomers:customerIds.length,
    searchedCustomers:customerIds.length,
    hydratedTasks:hydrated.length,
    priorTasks:prior.length,
    ageMinutes:isFinite(ageMinutes)
      ? Math.round(ageMinutes)
      : null
  };
}

// Original: 20_Striven; Function: tmv3_searchPreInspectionTasks_
function tmv3_searchPreInspectionTasks_(customer) {
  const customerId = Number(tmv3_clean_(customer && customer['Customer ID']) || 0);
  if (!customerId) throw new Error('PreInspection task search requires Customer ID.');

  const customerPayload = { Id: customerId };
  const customerNumber = tmv3_clean_(customer['Customer Number']);
  const customerName = tmv3_clean_(customer['Name']);
  if (customerNumber) customerPayload.Number = customerNumber;
  if (customerName) customerPayload.Name = customerName;

  const all = [];
  let previousSignature = '';

  for (let pageIndex = 0; pageIndex < 3; pageIndex++) {
    const payload = {
      Customer: customerPayload,
      Type: [105],
      PageIndex: pageIndex,
      PageSize: 100
    };

    const json = tmv3_fetchJson_(
      TMV3.API_BASE + '/v2/tasks/search',
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload)
      }
    );

    const rows = tmv3_extractTaskSearchRows_(json);
    const signature = rows.slice(0, 5).map(function(row) {
      return tmv3_clean_(tmv3_first_(row, ['id','Id','taskId','TaskId']));
    }).join('|');

    if (pageIndex > 0 && signature && signature === previousSignature) {
      throw new Error('PreInspection task search pagination repeated a page.');
    }

    previousSignature = signature;
    Array.prototype.push.apply(all, rows);

    if (!rows.length || rows.length < 100) break;
  }

  const ids = tmv3_unique_(all.map(function(row) {
    return tmv3_first_(row, ['id','Id','taskId','TaskId']);
  })).slice(0, 11);

  if (ids.length > 10) {
    throw new Error('More than 10 PreInspection task candidates were returned for Customer + Type 105.');
  }

  return ids.map(function(id) {
    return tmv3_getTaskById_(id);
  }).filter(function(task) {
    return (
      Number(task['Task Type ID'] || 0) === 105 &&
      String(task['Customer ID'] || '') === String(customerId)
    );
  });
}

// Original: 30_Calendar; Function: tmv3_preInspectionStep1SourceAllowed_
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
function tmv3_preInspectionEnsureLegacyTitle_(description, legacyTitle) {
  const raw = String(description || '');
  const legacy = tmv3_clean_(legacyTitle);
  if (!legacy) return raw;
  const readable = tmv3_calendarReadableDescription_(raw).trim();
  if (readable === legacy) return raw;
  const html = /<[a-z][^>]*>/i.test(raw);
  const separator = html ? '<br><br>' : '\n\n';
  // Normalize only the separator after the preserved title. Authored notes
  // and their inline markup remain intact.
  if (readable.indexOf(legacy) === 0) {
    const escaped = tmv3_calendarHtmlEscape_(legacy);
    const token = raw.indexOf(escaped) >= 0 ? escaped : legacy;
    const ix = raw.indexOf(token);
    if (ix >= 0) {
      const tail = raw.slice(ix+token.length).replace(/^(?:\s|<br\s*\/?>)+/gi, '');
      return raw.slice(0,ix) + token + separator + tail;
    }
  }
  return html
    ? tmv3_calendarHtmlEscape_(legacy) + '<br><br>' + raw
    : legacy + (raw ? '\n\n' + raw : '');
}

// Original: 35_CalendarLinks; Function: tmv3_preInspectionAuthoredMarkup_
function tmv3_preInspectionAuthoredMarkup_(description) {
  let text = tmv3_stripV3ManagedBlock_(String(description || ''));
  text = text.replace(
    /<!--\s*PREINSPECT_STRIVEN_TASK_LINK_START\s*-->[\s\S]*?<!--\s*PREINSPECT_STRIVEN_TASK_LINK_END\s*-->/gi,
    ''
  );
  const start = tmv3_managedCalendarBlockStart_(text);
  if (start >= 0) text = text.slice(0, start);
  // Remove only the empty wrapper/separator belonging to the managed footer.
  text = text.replace(/(?:<p[^>]*>\s*|<br\s*\/?>\s*|\s)+$/gi, '');
  if (!/<[a-z][^>]*>/i.test(text)) {
    text = text.replace(/^\s*#{1,6}\s+/gm, '');
    text = text.replace(/\r\n?/g, '\n')
      .replace(/\n/, '<br>')
      .replace(/\n\s*\n+/g, '<br>')
      .replace(/[ \t]*\n[ \t]*/g, ' ');
    return text.split('<br>').map(tmv3_calendarHtmlEscape_).join('<br>');
  }
  // Imported Markdown headings and paragraph wrappers must not turn the
  // preserved sales-rep title into oversized text or artificial vertical gaps.
  // Real inline formatting and lists stay intact. API/export line wrapping
  // inside HTML text is whitespace, not an authored hard line break.
  return text
    .replace(/<\/(?:p|h[1-6])\s*>/gi, '<br>')
    .replace(/<(?:p|h[1-6])(?:\s[^>]*)?>/gi, '')
    .replace(/[\r\n]+[ \t]*/g, ' ')
    .replace(/(?:<br\s*\/?>\s*){2,}/gi, '<br>')
    .replace(/^\s+|(?:<br\s*\/?>\s*|\s)+$/gi, '');
}

// Original: 35_CalendarLinks; Function: tmv3_preInspectionClickableManagedDescription_
function tmv3_preInspectionClickableManagedDescription_(existingDescription, plan) {
  const before = String(existingDescription || '');
  const links = (plan && plan.links) || [];
  const authored = tmv3_preInspectionAuthoredMarkup_(before);
  const managed = tmv3_calendarHtmlEscape_(TMV3_FINAL_LINK_HEADING) +
    (links.length ? '<br>' + links.map(tmv3_preInspectionHtmlLink_).join('<br>') : '');
  const description = authored + (authored ? '<br><br>' : '') + managed;
  return {
    description:description,
    status:description === before
      ? 'PRESERVED_EXISTING_CLICKABLE_LINK_BLOCK'
      : 'RESTORED_CLICKABLE_HYPERLINK_BLOCK',
    changed:description !== before
  };
}

// Original: 35_CalendarLinks; Function: tmv3_preInspectionTitleVariants_
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

  const contactResult = tmv3_step4ResolveContact_(record, customer, '');

  // PreInspection never invents a Contact. Contact ambiguity is retained as a
  // warning because Requested By comes from the Calendar organizer, not Contact.
  if (contactResult.status === 'REVIEW') {
    warnings.push(contactResult.reason || 'Contact requires review.');
  } else if (contactResult.status === 'NO_MATCH') {
    warnings.push('No deterministic Contact match; this does not block PreInspection.');
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
  const customerId = tmv3_clean_(customer && customer['Customer ID']);
  const locationId = tmv3_clean_(location && location['Location ID']);

  let candidates = [];

  if (Array.isArray(cachedCandidates)) {
    candidates = cachedCandidates.slice();

    if (eventRecord.existingTaskId) {
      const linkedId = tmv3_clean_(eventRecord.existingTaskId);
      const linked = candidates.filter(function(task) {
        return tmv3_clean_(task['Task ID']) === linkedId;
      });

      // Prefer an exact cached Calendar Task link when present. If the link is
      // stale/missing from cache, continue with the verified Customer's cached
      // Type-105 candidates rather than spending a live Task GET.
      if (linked.length === 1) candidates = linked;
    }
  } else if (eventRecord.existingTaskId) {
    const linked = tmv3_getTaskById_(eventRecord.existingTaskId);
    candidates = linked ? [linked] : [];
  } else {
    candidates = tmv3_searchPreInspectionTasks_(customer);
  }

  candidates = (candidates || []).filter(function(task) {
    return (
      Number(task['Task Type ID'] || 0) === 105 &&
      tmv3_clean_(task['Customer ID']) === customerId
    );
  });

  const evaluated = candidates.map(function(task) {
    const taskStart = tmv3_clean_(task['Start']);
    const taskDue = tmv3_clean_(task['Due']);
    const sameStart = tmv3_sameMinute_(eventRecord.start, taskStart);
    const sameDue = tmv3_sameMinute_(eventRecord.end, taskDue);
    const sameDay = tmv3_sameLocalDay_(eventRecord.start, taskStart);
    const sameLocation =
      !!locationId &&
      tmv3_clean_(task['Location ID']) === locationId;

    const taskTitle = tmv3_clean_(task['Name']);
    const titlePhones = tmv3_allPhones_(taskTitle);
    const titlePhoneMatch =
      !!eventRecord.phone &&
      titlePhones.indexOf(tmv3_phone10_(eventRecord.phone)) !== -1;

    const customerName = tmv3_norm_(customer && customer['Name']);
    const titleNameMatch =
      !!customerName &&
      tmv3_norm_(taskTitle).indexOf(customerName) !== -1;

    const identityCorroborated = titlePhoneMatch || titleNameMatch;
    const exactCalendarTaskLink =
      !!eventRecord.existingTaskId &&
      tmv3_clean_(eventRecord.existingTaskId) ===
        tmv3_clean_(task['Task ID']);

    const strongMatch =
      exactCalendarTaskLink ||
      sameStart ||
      (sameDay && sameLocation && identityCorroborated);

    return {
      task: task,
      open: tmv3_taskIsOpen_(task['Status']),
      sameStart: sameStart,
      sameDue: sameDue,
      sameDay: sameDay,
      sameLocation: sameLocation,
      titlePhoneMatch: titlePhoneMatch,
      titleNameMatch: titleNameMatch,
      identityCorroborated: identityCorroborated,
      exactCalendarTaskLink: exactCalendarTaskLink,
      strongMatch: strongMatch,
      hasStart: !!taskStart,
      score:
        (sameStart ? 100 : 0) +
        (sameDue ? 30 : 0) +
        (sameLocation ? 50 : 0) +
        (sameDay ? 20 : 0) +
        (titlePhoneMatch ? 80 : 0) +
        (titleNameMatch ? 40 : 0)
    };
  });

  const open = evaluated.filter(function(c) { return c.open; });
  const historyTasks = evaluated
    .filter(function(c) { return !c.open && tmv3_taskIsCompleted_(c.task['Status']); })
    .map(function(c) { return c.task; });
  const nonOpenActive = evaluated.filter(function(c) {
    return (
      !c.open &&
      !tmv3_taskIsCompleted_(c.task['Status'])
    );
  });

  if (!open.length && nonOpenActive.length) {
    return {
      status: 'REVIEW',
      task: null,
      historyTasks: historyTasks,
      reason: 'Applicable PreInspection Task exists in a non-open active status and must not be duplicated.',
      errorCode: 'NON_OPEN_ACTIVE_PREINSPECTION_TASK'
    };
  }

  // Step 5 resolves only the current OPEN task. Historical tasks are carried
  // forward so Step 6 can decide CREATE vs RECREATE vs FULFILLED safely.
  if (!open.length) {
    return {
      status: 'CLEAR',
      task: null,
      historyTasks: historyTasks,
      reason: 'No OPEN PreInspection task remains for this appointment.'
    };
  }

  open.sort(function(a, b) { return b.score - a.score; });

  const strong = open.filter(function(c) { return c.strongMatch; });

  if (strong.length === 1) {
    return {
      status: 'MATCHED',
      task: strong[0].task,
      reason: 'One OPEN PreInspection task strongly matches this appointment.',
      evidence: tmv3_preInspectionCandidateEvidence_(strong[0])
    };
  }

  if (strong.length > 1) {
    return {
      status: 'REVIEW',
      reason: 'Multiple OPEN PreInspection tasks strongly match this appointment.',
      errorCode: 'MULTIPLE_STRONG_PREINSPECTION_TASKS'
    };
  }

  // Business rule: if this Customer already has exactly one OPEN
  // PreInspection Task (Type 105), reuse it and patch it to the current
  // Calendar appointment rather than creating another Task. The Customer
  // match is already exact because candidates are filtered by Customer ID
  // above. Multiple OPEN PreInspection Tasks for the same Customer remain
  // review-only because choosing between them would be ambiguous.
  if (open.length === 1) {
    return {
      status: 'MATCHED',
      task: open[0].task,
      reason: 'Exactly one OPEN PreInspection Task exists for this Customer. Reuse it and patch the appointment details.',
      evidence: tmv3_preInspectionCandidateEvidence_(open[0]).concat([
        'SINGLE_OPEN_PREINSPECTION_FOR_CUSTOMER_REUSED'
      ])
    };
  }

  const sameDay = open.filter(function(c) { return c.sameDay; });
  if (sameDay.length) {
    return {
      status: 'REVIEW',
      reason: 'Multiple OPEN PreInspection Tasks exist for this Customer and at least one is on the same day. Select the correct Task before writing.',
      errorCode: 'MULTIPLE_OPEN_PREINSPECTION_TASKS_SAME_DAY'
    };
  }

  const sameLocation = open.filter(function(c) { return c.sameLocation; });
  if (sameLocation.length) {
    return {
      status: 'REVIEW',
      reason: 'Multiple OPEN PreInspection Tasks exist for this Customer and at least one uses this Location. Select the correct Task before writing.',
      errorCode: 'MULTIPLE_OPEN_PREINSPECTION_TASKS_SAME_LOCATION'
    };
  }

  const unresolved = open.filter(function(c) { return !c.hasStart; });
  const reusableUnscheduled = unresolved.filter(function(c) {
    const noConflictingLocation =
      !tmv3_clean_(c.task['Location ID']) || c.sameLocation;
    return c.identityCorroborated && noConflictingLocation;
  });

  if (open.length === 1 && reusableUnscheduled.length === 1) {
    return {
      status: 'MATCHED',
      task: reusableUnscheduled[0].task,
      reason: 'Exactly one unscheduled OPEN PreInspection task is identity-corroborated and has no conflicting Location.',
      evidence: tmv3_preInspectionCandidateEvidence_(reusableUnscheduled[0])
    };
  }

  if (unresolved.length) {
    return {
      status: 'REVIEW',
      reason: 'OPEN PreInspection candidate(s) have no usable start date and cannot be safely reused or ruled out.',
      errorCode: 'UNRESOLVED_PREINSPECTION_TASK'
    };
  }

  return {
    status: 'CLEAR',
    task: null,
    historyTasks: historyTasks,
    reason: 'No OPEN existing PreInspection task match was found for this appointment.'
  };
}

// Original: 40_Engine; Function: tmv3_preInspectionCandidateEvidence_
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
  const identity = record.step4 || {};
  const customer = identity.customer;
  const location = identity.location;
  const customerId = tmv3_clean_(
    customer && customer['Customer ID']
  );

  const cachedCandidates = (
    refs.tasksByCustomer[customerId] || []
  ).filter(function(task) {
    return Number(task['Task Type ID'] || 0) === 105;
  });

  const warnings = [];
  const evidence = ['PREINSPECTION_TASKS_FROM_SHARED_CACHE'];

  const linkedTaskMissingFromCache =
    !!record.existingTaskId &&
    !refs.taskById[tmv3_clean_(record.existingTaskId)];

  if (linkedTaskMissingFromCache) {
    warnings.push(
      'Calendar carries Task ' +
      tmv3_clean_(record.existingTaskId) +
      ', but it is not present in the current shared Task cache. Stage 5 did not spend a live Task GET.'
    );
    evidence.push(
      'CALENDAR_TASK_LINK_NOT_IN_CACHE_DEFER_LIVE_READ'
    );

    if (!cachedCandidates.length) {
      return tmv3_step5Decision_(
        'REVIEW',
        'PREINSPECTION_CALENDAR_TASK_NOT_IN_CACHE_NO_SAFE_FALLBACK',
        'Calendar carries a PreInspection Task link that is absent from the current shared Task cache, and no Customer-owned cached Type-105 Task can safely replace it.',
        [],
        0,
        evidence,
        warnings
      );
    }
  }

  try {
    const decision = tmv3_preInspectionTaskDecision_(
      record,
      customer,
      location,
      cachedCandidates
    );

    if (decision.status === 'MATCHED' && decision.task) {
      return tmv3_step5Decision_(
        'MATCHED',
        'PREINSPECTION_TASK_MATCHED',
        decision.reason,
        [decision.task],
        (decision.historyTasks || []).length,
        evidence.concat(decision.evidence || []),
        warnings,
        decision.historyTasks || []
      );
    }

    if (decision.status === 'CLEAR') {
      return tmv3_step5Decision_(
        'NO_TASK',
        'PREINSPECTION_NO_OPEN_TASK',
        decision.reason,
        [],
        (decision.historyTasks || []).length,
        evidence,
        warnings,
        decision.historyTasks || []
      );
    }

    return tmv3_step5Decision_(
      'REVIEW',
      decision.errorCode || 'PREINSPECTION_TASK_REVIEW',
      decision.reason || 'PreInspection Task candidates require review.',
      [],
      0,
      evidence.concat(decision.evidence || []),
      warnings
    );
  } catch (err) {
    return tmv3_step5Decision_(
      'REVIEW',
      'PREINSPECTION_TASK_CACHE_RESOLUTION_ERROR',
      String(err && err.message || err),
      [],
      0,
      evidence,
      warnings
    );
  }
}

// Original: 42_Step5_Task_Resolution; Function: tmv3_step5NormalizePreInspectionTaskTimesAcrossRecords_
function tmv3_step5NormalizePreInspectionTaskTimesAcrossRecords_(records) {
  const corrections = {};

  (records || []).forEach(function(record) {
    if (record.vertical !== 'PreInspection') return;

    const tasks =
      record.step5 && Array.isArray(record.step5.tasks)
        ? record.step5.tasks
        : [];

    tasks.forEach(function(task) {
      const taskId = tmv3_clean_(task && task['Task ID']);
      if (!taskId) return;

      if (!corrections[taskId]) {
        corrections[taskId] = {
          Start:null,
          Due:null,
          conflictStart:false,
          conflictDue:false
        };
      }

      [
        ['Start', record.start],
        ['Due', record.end]
      ].forEach(function(pair) {
        const field = pair[0];
        const calendarValue = pair[1];
        const taskValue = tmv3_clean_(task && task[field]);
        if (!taskValue || !calendarValue) return;

        const taskDate = tmv3_parseDateTime_(taskValue);
        const calendarDate = tmv3_parseDateTime_(calendarValue);
        if (!taskDate || !calendarDate) return;

        if (
          tmv3_step6LocalDay_(taskDate) !==
          tmv3_step6LocalDay_(calendarDate)
        ) {
          return;
        }

        const diff = Math.abs(
          taskDate.getTime() - calendarDate.getTime()
        );

        if (diff !== 12 * 60 * 60 * 1000) return;

        const corrected = Utilities.formatDate(
          calendarDate,
          TMV3_TIMEZONE,
          "yyyy-MM-dd'T'HH:mm:ssXXX"
        );

        if (
          corrections[taskId][field] &&
          corrections[taskId][field] !== corrected
        ) {
          corrections[taskId]['conflict' + field] = true;
          return;
        }

        corrections[taskId][field] = corrected;
      });
    });
  });

  (records || []).forEach(function(record) {
    if (record.vertical !== 'PreInspection' || !record.step5) return;

    ['tasks','historyTasks'].forEach(function(listName) {
      const list = Array.isArray(record.step5[listName])
        ? record.step5[listName]
        : [];

      record.step5[listName] = list.map(function(task) {
        const out = Object.assign({}, task || {});
        const taskId = tmv3_clean_(out['Task ID']);
        const fix = corrections[taskId];
        if (!fix) return out;

        if (fix.Start && !fix.conflictStart) out.Start = fix.Start;
        if (fix.Due && !fix.conflictDue) out.Due = fix.Due;

        return out;
      });
    });
  });

  return records;
}

// Original: 42_Step5_Task_Resolution; Function: tmv3_step5AlignPreInspectionTaskClock_
function tmv3_step5AlignPreInspectionTaskClock_(task, record) {
  const out = Object.assign({}, task || {});

  [
    ['Start', record && record.start],
    ['Due', record && record.end]
  ].forEach(function(pair) {
    const field = pair[0];
    const calendarValue = pair[1];
    const taskValue = tmv3_clean_(out[field]);

    if (!taskValue || !calendarValue) return;

    const taskDate = tmv3_parseDateTime_(taskValue);
    const calendarDate = tmv3_parseDateTime_(calendarValue);

    if (!taskDate || !calendarDate) {
      return;
    }

    if (
      tmv3_step6LocalDay_(taskDate) !==
      tmv3_step6LocalDay_(calendarDate)
    ) {
      return;
    }

    const diff = Math.abs(
      taskDate.getTime() - calendarDate.getTime()
    );

    // Striven V2 occasionally serializes afternoon 1–11 PM as 1–11 AM.
    // Correct only the exact 12-hour case corroborated by this Calendar event.
    if (diff === 12 * 60 * 60 * 1000) {
      out[field] = Utilities.formatDate(
        calendarDate,
        TMV3_TIMEZONE,
        "yyyy-MM-dd'T'HH:mm:ssXXX"
      );
    }
  });

  return out;
}

// Original: 46_Step7_Reconciliation; Function: tmv3_step7HydrateLinkedPreInspectionTasks_
function tmv3_step7HydrateLinkedPreInspectionTasks_(records, refs, runtime) {
  (records || []).forEach(function(record) {
    const identity = record.step4 || {};
    const customerId = tmv3_clean_(identity.customer && identity.customer['Customer ID']);
    const linkedId = tmv3_clean_(record.existingTaskId);
    if (identity.disposition !== 'VERIFIED' ||
        !customerId || !linkedId || (record.vertical==='PreInspection' && refs.taskById[linkedId])) return;
    if (!/^[1-9]\d*$/.test(linkedId)) {
      throw new Error('PREINSPECTION_LINKED_TASK_ID_INVALID_NO_CREATE');
    }
    // Exact linked ID only, one fresh GET shared with the later Step 7 checks.
    // Failed/mismatched reads remain blocked; they never authorize CREATE.
    const fresh = tmv3_step7FreshTask_(linkedId, runtime);
    if (!fresh.ok || !fresh.task ||
        tmv3_clean_(fresh.task['Task ID']) !== linkedId ||
        !tmv3_taskFitsVertical_(fresh.task,record.vertical,TMV3.VERTICALS[record.vertical]) ||
        tmv3_clean_(fresh.task['Customer ID']) !== customerId) {
      throw new Error('PREINSPECTION_LINKED_TASK_UNVERIFIED_NO_CREATE: ' + linkedId);
    }
    const task = fresh.task;
    refs.taskById[linkedId] = task;
    refs.tasks=refs.tasks.filter(function(t){return tmv3_clean_(t['Task ID'])!==linkedId;});
    Object.keys(refs.tasksByCustomer).forEach(function(k){refs.tasksByCustomer[k]=refs.tasksByCustomer[k].filter(function(t){return tmv3_clean_(t['Task ID'])!==linkedId;});});
    Object.keys(refs.tasksByOrder).forEach(function(k){refs.tasksByOrder[k]=refs.tasksByOrder[k].filter(function(t){return tmv3_clean_(t['Task ID'])!==linkedId;});});
    refs.tasks.push(task);
    if (!refs.tasksByCustomer[customerId]) refs.tasksByCustomer[customerId] = [];
    refs.tasksByCustomer[customerId].push(task);
    const orderId = tmv3_clean_(task['Order ID']);
    if (orderId) {
      if (!refs.tasksByOrder[orderId]) refs.tasksByOrder[orderId] = [];
      refs.tasksByOrder[orderId].push(task);
    }
    tmv3_upsertTaskCacheRow_(task);
  });
}

// Original: 50_Verticals; Function: tmv3_preInspectionAssignmentOverride_
function tmv3_preInspectionAssignmentOverride_(record) {
  if (!record || !tmv3_clean_(record.eventId) || typeof PropertiesService === 'undefined') return [];
  const key = 'TMV3_PI_ASSIGNMENT_OVERRIDE_' + tmv3_hash_(record.eventId);
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) {
    // The employee comes from the original verified appointment organizer,
    // never from the sales representative who created the event or a guest.
    const owner=String(record.organizerCalendarId||'').toLowerCase();
    const stephen=String((TMV3.VERTICALS.PreInspection.secondaryCalendarIds||[])[0]||'').toLowerCase();
    if (owner && owner===stephen) return [15]; // Stephen Foley — verified Striven employee ID.
    if (owner==='jay@classicfireplace.ca') return [1]; // Jay Scott — verified Striven employee ID.
    if (owner==='karen@classicfireplace.ca')
      throw new Error('PREINSPECTION_KAREN_STRIVEN_EMPLOYEE_ID_NOT_VERIFIED_NO_WRITE');
    return []; // Unknown organizer: preserve Pool 8, never invent an employee.
  }
  let value;
  try { value = JSON.parse(raw); } catch (error) { throw new Error('PREINSPECTION_ASSIGNMENT_OVERRIDE_INVALID'); }
  const employeeId = Number(value.employeeId);
  if (String(value.eventId || '') !== String(record.eventId) ||
      !value.customerNumber || String(value.customerNumber) !== String(record.customerNumber || '') ||
      !value.organizerCalendarId || String(value.organizerCalendarId) !== String(record.organizerCalendarId || '') ||
      !Number.isSafeInteger(employeeId) || employeeId <= 0) {
    throw new Error('PREINSPECTION_ASSIGNMENT_OVERRIDE_EVIDENCE_MISMATCH');
  }
  return [employeeId];
}

// Original: 50_Verticals; Function: tmv3_configurePreInspectionAssignmentOverride
function tmv3_configurePreInspectionAssignmentOverride(eventId, customerNumber, organizerCalendarId, employeeId) {
  if (!/^[a-z0-9_]+@google\.com$/.test(String(eventId || '')) ||
      !/^\d{3,8}$/.test(String(customerNumber || '')) ||
      !tmv3_clean_(organizerCalendarId) || !Number.isSafeInteger(Number(employeeId)) ||
      Number(employeeId) <= 0 || !tmv3_employeeNameById_(employeeId)) {
    throw new Error('PREINSPECTION_ASSIGNMENT_OVERRIDE_INVALID');
  }
  const record=tmv3_findFreshEventRecord_('PreInspection',eventId);
  if (!record || String(record.eventId)!==String(eventId) ||
      String(record.customerNumber || '')!==String(customerNumber) ||
      String(record.organizerCalendarId || '')!==String(organizerCalendarId)) {
    throw new Error('PREINSPECTION_ASSIGNMENT_OVERRIDE_EVIDENCE_MISMATCH');
  }
  const value={eventId:String(eventId),customerNumber:String(customerNumber),
    organizerCalendarId:String(organizerCalendarId),employeeId:Number(employeeId)};
  const key='TMV3_PI_ASSIGNMENT_OVERRIDE_'+tmv3_hash_(eventId);
  const props=PropertiesService.getScriptProperties();props.setProperty(key,JSON.stringify(value));
  if (props.getProperty(key)!==JSON.stringify(value)) throw new Error('PREINSPECTION_ASSIGNMENT_OVERRIDE_READBACK_FAILED');
  return {status:'ASSIGNMENT_OVERRIDE_VERIFIED',employeeId:Number(employeeId),businessWrites:0};
}

// Explicit operator approval, scoped to immutable iCalUID. No write permission.
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
  return {
    taskTypeId: 105,
    attachSalesOrder: false,
    description: '',
    calendarNotes: 'CALENDAR_ONLY',
    previousCalendarTitlePolicy: 'PRESERVE_AS_FIRST_DESCRIPTION_LINE_NO_HEADER',
    descriptionHeaderPolicy: 'REMOVE_NOTES_AND_SALES_ORDER_LABELS_PRESERVE_CONTENT',
    calendarCopyParityPolicy: 'FULL_BUSINESS_VISIBLE_PARITY_REQUIRED',
    calendarLinksPreservationPolicy: 'PRESERVE_EXISTING_CLICKABLE_BLOCK_REPAIR_NONCLICKABLE_OR_MISSING_ONLY',
    calendarTaskLinkLabelPolicy: 'TASK_ID_PLUS_ACTUAL_STRIVEN_TASK_NAME',
    field854: 'DO_NOT_MANAGE',
    infoCustomFieldsAtCreate: 'NONE',
    requiredPoolId: 8,
    requestedBy: 'CALENDAR_ORGANIZER',
    assignedToFromOrganizer: false,
    calendarLinks: [
      'CUSTOMER_SALES_ORDERS_PAGE',
      'TASK'
    ],
    technicianFields: 'DO_NOT_PREFILL'
  };
}

// Original: 58_Operations; Function: tmv3_ensurePreInspectionMirror_
function tmv3_ensurePreInspectionMirror_(eventRecord, scope) {
  if(tmv3_preInspectionApprovalHold_(eventRecord&&eventRecord.eventId))return {status:'APPROVAL_GUEST_HOLD',writePerformed:false};
  // An appointment owned by another organizer must not invite Stephen.
  const organizer=tmv3_clean_(eventRecord && eventRecord.organizerCalendarId);
  if (!organizer) return {status:'OWNER_UNKNOWN_NO_GUEST_WRITE',writePerformed:false};
  const secondary=((TMV3.VERTICALS.PreInspection || {}).secondaryCalendarIds || []).map(tmv3_clean_);
  if (eventRecord && eventRecord.vertical==='PreInspection' && organizer && secondary.indexOf(organizer)<0) {
    return {status:'OTHER_ORGANIZER_NO_STEPHEN_MIRROR',writePerformed:false};
  }
  if (
    !eventRecord ||
    eventRecord.vertical !== 'PreInspection'
  ) {
    return { status:'NOT_NEEDED', writePerformed:false };
  }

  tmv3_assertOperationWrite_(scope);

  const cfg = TMV3.VERTICALS.PreInspection || {};
  const primaryId = tmv3_clean_(cfg.primaryCalendarId);
  const stephenEmail = tmv3_normEmail_(
    cfg.secondaryOwnerEmail ||
    'stephen@classicfireplace.ca'
  );

  if (!primaryId || !stephenEmail) {
    throw new Error(
      'PreInspection primary Calendar or Stephen guest email is not configured.'
    );
  }

  const primaryCopy = tmv3_findEventCopyRobust_(
    primaryId,
    eventRecord.eventId
  );

  if (!primaryCopy || !primaryCopy.event) {
    throw new Error(
      'Authoritative CF Preinspects event could not be found for Event ' +
      eventRecord.eventId + '.'
    );
  }

  function guestEmails(event) {
    return (event.getGuestList ? event.getGuestList(true) : [])
      .map(function(guest) {
        return tmv3_normEmail_(
          guest && guest.getEmail ? guest.getEmail() : ''
        );
      })
      .filter(Boolean);
  }

  if (
    guestEmails(primaryCopy.event).indexOf(
      stephenEmail
    ) !== -1
  ) {
    return {
      status:'STEPHEN_ALREADY_PRESENT',
      writePerformed:false,
      primaryCalendarId:primaryId,
      stephenEmail:stephenEmail
    };
  }

  primaryCopy.event.addGuest(stephenEmail);

  const readback = tmv3_findEventCopyRobust_(
    primaryId,
    eventRecord.eventId
  );

  if (
    !readback ||
    !readback.event ||
    guestEmails(readback.event).indexOf(
      stephenEmail
    ) === -1
  ) {
    throw new Error(
      'PreInspection Stephen guest read-back failed for Event ' +
      eventRecord.eventId + '.'
    );
  }

  return {
    status:'STEPHEN_GUEST_WRITTEN_AND_VERIFIED',
    writePerformed:true,
    primaryCalendarId:primaryId,
    stephenEmail:stephenEmail
  };
}

// Original: 58_Operations; Function: tmv3_preInspectionCalendarBusinessSnapshot_
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
function tmv3_reconcilePreInspectionCalendarParity_(eventRecord, scope) {
  if(tmv3_preInspectionApprovalHold_(eventRecord&&eventRecord.eventId))return {status:'APPROVAL_CALENDAR_HOLD',copies:[],writes:0};
  if (!eventRecord || eventRecord.vertical !== 'PreInspection') {
    return { status:'NOT_NEEDED', copies:[] };
  }

  tmv3_assertOperationWrite_(scope);

  const cfg = TMV3.VERTICALS.PreInspection || {};
  const primaryId = tmv3_clean_(cfg.primaryCalendarId);
  const ids = tmv3_requiredCalendarCopiesForEvent_(eventRecord);

  if (!primaryId || !ids.length) {
    throw new Error('PreInspection Calendar parity configuration is incomplete.');
  }

  const primaryFound = tmv3_findEventCopyRobust_(
    primaryId,
    eventRecord.eventId
  );
  if (!primaryFound || !primaryFound.event) {
    throw new Error(
      'Authoritative CF Preinspects copy is missing; parity cannot be reconciled.'
    );
  }

  const authority = tmv3_preInspectionCalendarBusinessSnapshot_(
    primaryFound.event
  );
  if (authority.allDay) {
    throw new Error(
      'PreInspection parity refuses to normalize an all-day authoritative event.'
    );
  }

  const results = [];

  ids.forEach(function(calendarId) {
    const found = tmv3_findEventCopyRobust_(
      calendarId,
      eventRecord.eventId
    );
    if (!found || !found.event) {
      throw new Error(
        'Required PreInspection Calendar copy is missing: ' + calendarId
      );
    }

    const event = found.event;
    let before = tmv3_preInspectionCalendarBusinessSnapshot_(event);
    let writePerformed = false;

    if (before.title !== authority.title) {
      event.setTitle(authority.title);
      writePerformed = true;
    }
    if (before.description !== authority.description) {
      event.setDescription(authority.description);
      writePerformed = true;
    }
    if (before.location !== authority.location) {
      event.setLocation(authority.location);
      writePerformed = true;
    }
    if (
      before.start !== authority.start ||
      before.end !== authority.end ||
      before.allDay !== authority.allDay
    ) {
      if (before.allDay || authority.allDay) {
        throw new Error(
          'PreInspection parity will not convert between all-day and timed events.'
        );
      }
      event.setTime(
        new Date(authority.start),
        new Date(authority.end)
      );
      writePerformed = true;
    }

    if (
      authority.requiredStephenPresent &&
      !before.requiredStephenPresent
    ) {
      event.addGuest(
        tmv3_normEmail_(
          cfg.secondaryOwnerEmail ||
          'stephen@classicfireplace.ca'
        )
      );
      writePerformed = true;
    }

    const readbackFound = tmv3_findEventCopyRobust_(
      calendarId,
      eventRecord.eventId
    );
    if (!readbackFound || !readbackFound.event) {
      throw new Error(
        'Required PreInspection Calendar copy disappeared during parity read-back.'
      );
    }

    const after = tmv3_preInspectionCalendarBusinessSnapshot_(
      readbackFound.event
    );

    if (!tmv3_preInspectionCalendarParityEqual_(authority, after)) {
      throw new Error(
        'PreInspection cross-calendar parity read-back failed on ' +
        calendarId +
        '. Expected=' + JSON.stringify(authority) +
        ' Actual=' + JSON.stringify(after)
      );
    }

    results.push({
      calendarId:calendarId,
      status:writePerformed ? 'WRITTEN_AND_VERIFIED' : 'ALREADY_MATCHED',
      businessState:after
    });
  });

  return {
    status:'FULL_BUSINESS_VISIBLE_PARITY_VERIFIED',
    authoritativeCalendarId:primaryId,
    eventId:eventRecord.eventId,
    copies:results
  };
}

// Original: 58_Operations; Function: tmv3_enforcePreInspectionCreatePayloadHardRules_
function tmv3_enforcePreInspectionCreatePayloadHardRules_(payload) {
  payload = payload || {};

  delete payload.SalesOrder;
  delete payload.SalesOrderId;
  delete payload.SalesOrderID;
  delete payload.SOId;
  delete payload.InfoCustomFields;
  payload.Description = '';

  return payload;
}

// Original: 58_Operations; Function: tmv3_assertPreInspectionCreatePayloadHardRules_
function tmv3_assertPreInspectionCreatePayloadHardRules_(payload) {
  const errors = [];
  const type = payload && payload.Type || {};
  const description = String(payload && payload.Description || '');
  const hasOrder = !!(
    payload &&
    (
      payload.SalesOrder ||
      payload.SalesOrderId ||
      payload.SalesOrderID ||
      payload.SOId
    )
  );
  const hasCustomFields = !!(
    payload &&
    Array.isArray(payload.InfoCustomFields) &&
    payload.InfoCustomFields.length
  );

  if (Number(type.Id || type.id || 0) !== 105) {
    errors.push('Task Type must be 105.');
  }
  if (hasOrder) {
    errors.push('Sales Order attachment is forbidden.');
  }
  if (description !== '') {
    errors.push('Task Description must be blank at CREATE.');
  }
  if (hasCustomFields) {
    errors.push('PreInspection CREATE must not prefill InfoCustomFields, including Field 854.');
  }

  if (errors.length) {
    throw new Error(
      'PREINSPECTION HARD RULE VIOLATION: ' +
      errors.join(' | ')
    );
  }

  return true;
}

// Original: 58_Operations; Function: tmv3_preInspectionGuestEmails_
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
function tmv3_preInspectionGuestSyncDecision_(guestEmails, stephenEmail, ownerCalendarId) {
  // Owner is required. Guest presence alone must never add Stephen.
  const cfg = TMV3.VERTICALS.PreInspection || {};
  const verifiedStephenCalendar = String((cfg.secondaryCalendarIds || [])[0] || '').trim();
  const actualOwner = String(ownerCalendarId || '').trim();
  if (!actualOwner) return {status:'OWNER_NOT_VERIFIED',action:'NONE'};
  if (!verifiedStephenCalendar || actualOwner !== verifiedStephenCalendar)
    return {status:'OTHER_OWNER_NO_STEPHEN',action:'NONE'};
  const stephen = tmv3_normEmail_(stephenEmail || cfg.secondaryOwnerEmail);
  if (!stephen) return {status:'BLOCKED_CONFIG',action:'NONE'};
  const guests = tmv3_unique_((guestEmails || []).map(tmv3_normEmail_).filter(Boolean));
  if (guests.indexOf(stephen) >= 0) return {status:'ALREADY_PRESENT',action:'NONE'};
  return {status:'ADD_STEPHEN',action:'ADD'};
}

// Original next-section boundary
function tmv3_reconcilePreInspectionSharedGuest_AUTO_(options) {
  // Legacy global Stephen sync is unsafe for Jay/Karen appointments.
  // Keep the public entrypoint without mutating any Calendar guest.
  tmv3_piAssertLockedStandard_();
  return {status:'RETIRED_UNCONDITIONAL_STEPHEN_GUEST_SYNC',
    standard:'2026-11-10-r1',writes:0,scanned:0,
    reason:'OWNER_SPECIFIC_GUEST_SYNC_ONLY'};
}

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
  const key='TMV3_PI_INFERRED_'+tmv3_hash_(record.eventId);
  const rawEvidence=PropertiesService.getScriptProperties().getProperty(key);
  const evidence=rawEvidence ? JSON.parse(rawEvidence) : null;
  const number=String(record.orderNumber || (evidence && evidence.orderNumber) || '');
  if(!number)return {status:'OPTIONAL_ORDER_NOT_ATTACHED'};
  if(!/^\d+$/.test(number))throw Error('PREINSPECTION_ORDER_STATUS_UNVERIFIED_NO_CREATE');
  const customerId=Number(plan.expectedCustomerId),locationId=Number(plan.expectedLocationId);
  if(!Number.isSafeInteger(customerId)||customerId<=0)throw Error('PREINSPECTION_ORDER_CUSTOMER_UNVERIFIED_NO_CREATE');
  if(evidence && (String(evidence.orderNumber)!==number || Number(evidence.customerId)!==customerId || Number(evidence.locationId)!==locationId))throw Error('PREINSPECTION_INFERRED_ORDER_EVIDENCE_CONFLICT_NO_CREATE');
  let id=Number(evidence && evidence.orderId);
  if(!Number.isSafeInteger(id)||id<=0){
    const found=tmv3_fetchJson_(TMV3.API_BASE+'/v1/sales-orders/search',{method:'post',contentType:'application/json',payload:JSON.stringify({PageIndex:0,PageSize:2,Number:number,CustomerId:customerId})});
    const data=tmv3_first_(found,['Data']),count=Number(tmv3_first_(found,['TotalCount']));
    if(!Array.isArray(data)||count!==1||data.length!==1||String(tmv3_first_(data[0],['Number']))!==number)throw Error('PREINSPECTION_ORDER_NOT_UNIQUE_NO_CREATE');
    id=Number(tmv3_first_(data[0],['Id']));
  }
  if(!Number.isSafeInteger(id)||id<=0)throw Error('PREINSPECTION_ORDER_ID_UNVERIFIED_NO_CREATE');
  const order=tmv3_fetchJson_(TMV3.API_BASE+'/v1/sales-orders/'+id,{method:'get'});
  const customer=tmv3_first_(order,['Customer']),status=tmv3_first_(order,['Status']);
  if(Number(tmv3_first_(order,['Id']))!==id || String(tmv3_first_(order,['OrderNumber','Number']))!==number || Number(tmv3_first_(customer,['Id']))!==customerId)throw Error('PREINSPECTION_ORDER_OWNERSHIP_CONFLICT_NO_CREATE');
  const statusId=String(tmv3_first_(status,['Id'])),statusName=tmv3_norm_(tmv3_first_(status,['Name']));
  const known={'18':'incomplete','19':'quoted','20':'pending approval','22':'approved','25':'in progress'};
  if(!known[statusId]||known[statusId]!==statusName)throw Error('PREINSPECTION_ORDER_STATUS_UNVERIFIED_NO_CREATE');
  if(statusId==='22')throw Error('APPROVED_ORDER_REQUIRES_INSTALLATION_NO_PREINSPECTION_CREATE');
  if(statusId==='25')throw Error('IN_PROGRESS_ORDER_TYPE_UNVERIFIED_NO_PREINSPECTION_CREATE');
  return {status:'FRESH_PREAPPROVAL_ORDER_VERIFIED',orderId:id,orderNumber:number,orderStatus:statusName};
}

// Original: 60_Automation; Function: tmv3_preInspectionStoreDiscovery_
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
  const lock=LockService.getScriptLock();if(!lock.tryLock(30000))return {status:'LOCK_BUSY',writes:0};
  try{tmv3_preInspectionIntakeRecords_({recordDiscovery:true,resolveLikely:true});return tmv3_preInspectionDiscoveryStatus();}
  finally{lock.releaseLock();}
}

// Original: 60_Automation; Function: tmv3_reconcilePreInspectionEvent
function tmv3_reconcilePreInspectionEvent(eventId) {
  if(!/^[a-z0-9_]+@google\.com$/.test(String(eventId || '')))throw new Error('INVALID_PREINSPECTION_EVENT_ID');
  const started=Date.now();
  const result=tmv3_preInspectionAutomationCycle_({force:true,eventId:eventId,reason:'TARGETED_RECONCILIATION'});
  result.sheetRefresh=tmv3_refreshAutomationSheetsAfter_(started);
  return result;
}

// Original: 60_Automation; Function: tmv3_preInspectionIntakeRecords_
function tmv3_preInspectionIntakeRecords_(options) {
  options=options || {};
  const cfg = TMV3.VERTICALS.PreInspection;
  const calendars = tmv3_verticalCalendars_('PreInspection',cfg).concat(
    (cfg.secondaryCalendarIds || []).map(function(id) {
      return {calendarId:id,calendarName:'Stephen',role:'SECONDARY_STEPHEN'};
    })
  );
  const start = new Date(); start.setHours(0,0,0,0);
  const end = new Date(start); end.setDate(end.getDate() + Number(cfg.lookaheadDays || 365));
  const records = [], discovery = [];
  calendars.forEach(function(source) {
    const cal = CalendarApp.getCalendarById(source.calendarId);
    if (!cal) throw new Error('PreInspection source calendar unavailable: ' + source.calendarId);
    cal.getEvents(start,end).forEach(function(event) {
      const record = tmv3_calendarEvent_('PreInspection',cfg,source,event,
        {applyEligibility:false});
      if (!record || tmv3_step2PreInspectionNonCustomer_(record)) return;
      if (source.role === 'SECONDARY_STEPHEN') {
        if(records.some(function(r){return r.calendarRole==='PRIMARY_SHARED' && r.eventId===record.eventId && tmv3_iso_(r.start)===tmv3_iso_(record.start);}))return;
        const decision=options.resolveLikely===true ? tmv3_preInspectionClassifyLikely_(record,cfg) : tmv3_preInspectionDiscoveryDecision_(record,cfg);
        if (decision.status==='IGNORE') return;
        discovery.push({eventId:record.eventId,calendarId:record.calendarId,
          title:String(record.title || '').slice(0,100),start:tmv3_iso_(record.start),
          reason:decision.reason,status:decision.status,orders:decision.orders || [],error:decision.error || ''});
        if (decision.status!=='EXPLICIT_PREINSPECTION') return;
      }
      records.push(record);
    });
  });
  const primaryKeys={};
  records.filter(function(r){return r.calendarRole==='PRIMARY_SHARED';}).forEach(function(r){primaryKeys[r.eventId+'|'+tmv3_iso_(r.start)]=true;});
  if(options.recordDiscovery===true) tmv3_preInspectionStoreDiscovery_(discovery.filter(function(r){return !primaryKeys[r.eventId+'|'+r.start];}));
  const seen = {};
  return records.filter(function(record) {
    // CF source is visited first. Guest mirrors keep the same iCal UID.
    const key = record.eventId + '|' + tmv3_iso_(record.start);
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  }).sort(function(a,b) { return new Date(a.start)-new Date(b.start); });
}

// Original: 60_Automation; Function: tmv3_preInspectionPromoteCandidate_
function tmv3_preInspectionPromoteCandidate_(record) {
  if(tmv3_preInspectionApprovalHold_(record&&record.eventId))return {status:'APPROVED_INTAKE_EXTERNAL_IMPORT_ONLY_NO_INVITES',record:record};
  if (record.calendarRole !== 'SECONDARY_STEPHEN') return {status:'PRIMARY',record:record};
  if (!tmv3_preInspectionExplicitCandidate_(record)) throw new Error('NOT_EXPLICIT_PREINSPECTION');
  if (!tmv3_preInspectionCreatorEmail_(record)) throw new Error('CREATOR_MISSING_OR_AMBIGUOUS_NO_INTAKE_WRITE');
  tmv3_assertOperationWrite_('PREINSPECTIONAUTO','PREINSPECTION_INTAKE');
  const cfg = TMV3.VERTICALS.PreInspection;
  const primary = CalendarApp.getCalendarById(cfg.primaryCalendarId);
  if (!primary) throw new Error('CF Preinspects unavailable.');
  const existing = tmv3_findCalendarEventById_(cfg.primaryCalendarId,record.eventId);
  if (existing && existing.event) return {status:'PRIMARY',record:
    tmv3_calendarEvent_('PreInspection',cfg,tmv3_verticalCalendars_('PreInspection',cfg)[0],existing.event)};
  const duplicates = primary.getEvents(new Date(record.start),new Date(record.end)).filter(function(event) {
    return event.getId() !== record.eventId &&
      event.getStartTime().getTime() === new Date(record.start).getTime() &&
      event.getEndTime().getTime() === new Date(record.end).getTime() &&
      tmv3_phone10_(tmv3_extractPhone_(event.getTitle())) === tmv3_phone10_(record.phone) &&
      !!tmv3_phone10_(record.phone) && tmv3_norm_(event.getLocation()) === tmv3_norm_(record.location);
  });
  if (duplicates.length) throw new Error('POSSIBLE_DUPLICATE_CF_EVENT_NO_INTAKE_WRITE');
  const source = tmv3_findCalendarEventById_(record.calendarId,record.eventId);
  if (!source || !source.event) throw new Error('Stephen source event disappeared.');
  const freshSource=tmv3_calendarEvent_('PreInspection',cfg,{calendarId:record.calendarId,calendarName:'Stephen',role:'SECONDARY_STEPHEN'},source.event,{applyEligibility:false});
  const freshClassification=tmv3_preInspectionClassifyLikely_(freshSource,cfg,true);
  if(freshClassification.status!=='EXPLICIT_PREINSPECTION') throw new Error('SOURCE_INTENT_OR_CREATOR_NOT_AUTHORIZED_NO_INTAKE_WRITE');
  if(freshSource.fingerprint!==record.fingerprint) throw new Error('SOURCE_CHANGED_NO_INTAKE_WRITE');
  if(freshClassification.reason==='VERIFIED_QUOTED_ORDER_LIKELY_PREINSPECTION' && freshClassification.orders.length===1)PropertiesService.getScriptProperties().setProperty('TMV3_PI_INFERRED_'+tmv3_hash_(record.eventId),JSON.stringify(freshClassification.orders[0]));
  const guests = source.event.getGuestList(true).map(function(g) {return tmv3_normEmail_(g.getEmail());});
  [cfg.primaryCalendarId,cfg.secondaryOwnerEmail].forEach(function(email) {
    if (guests.indexOf(tmv3_normEmail_(email)) < 0) source.event.addGuest(email);
  });
  const after = source.event.getGuestList(true).map(function(g) {return tmv3_normEmail_(g.getEmail());});
  if (after.indexOf(tmv3_normEmail_(cfg.primaryCalendarId)) < 0 ||
      after.indexOf(tmv3_normEmail_(cfg.secondaryOwnerEmail)) < 0) throw new Error('INTAKE_GUEST_READBACK_FAILED');
  // Calendar propagation is asynchronous; no task writes until CF copy exists.
  const promoted = tmv3_findCalendarEventById_(cfg.primaryCalendarId,record.eventId);
  return promoted && promoted.event ? {status:'PRIMARY',record:
    tmv3_calendarEvent_('PreInspection',cfg,tmv3_verticalCalendars_('PreInspection',cfg)[0],promoted.event)}
    : {status:'WAITING_FOR_CF_COPY',record:record};
}

// Original: 60_Automation; Function: tmv3_preInspectionNotificationDateParts_
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
function tmv3_installPreInspectionAutomationTrigger() { return tmv3_installStandardDailyTriggers(); }

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
  if(tmv3_preInspectionApprovalHold_(record&&record.eventId))return {status:'APPROVAL_TITLE_HOLD',writes:0};
  tmv3_assertOperationWrite_('PREINSPECTIONAUTO','PRESERVE_LEGACY_TITLE');
  if (!record || record.vertical!=='PreInspection') throw new Error('PreInspection event required.');
  const cfg=TMV3.VERTICALS.PreInspection;
  const state=tmv3_titleMigrationStateForEvent_('PreInspection',record.eventId,cfg.primaryCalendarId);
  let legacy=state.preservedTitle;
  const copies=tmv3_requiredCalendarCopiesForEvent_(record).map(function(id){
    const found=tmv3_findEventCopyRobust_(id,record.eventId);
    if (!found || !found.event) throw new Error('LEGACY_TITLE_REPAIR_MISSING_COPY: '+id);
    return {calendarId:id,event:found.event};
  });
  if (!legacy) {
    const primary=copies.filter(function(copy){return copy.calendarId===cfg.primaryCalendarId;})[0];
    const canonical=primary && String(primary.event.getTitle() || '');
    const oldTitles=tmv3_unique_(copies.map(function(copy){return String(copy.event.getTitle() || '');}).filter(function(title){return title && title!==canonical;}));
    if (/^\d+ - .+ - \(\d{3}\) \d{3}-\d{4}$/.test(canonical) && oldTitles.length===1) {
      legacy=oldTitles[0];
      tmv3_recordTitleMigrationState_('PreInspection',record.eventId,cfg.primaryCalendarId,legacy,canonical);
    } else if (oldTitles.length>1) throw new Error('AMBIGUOUS_PREVIOUS_COPY_TITLES_NO_OVERWRITE');
    else return {status:'NO_RECORDED_LEGACY_TITLE',eventId:record.eventId,writes:0};
  }
  let writes=0;
  copies.forEach(function(copy){
    const before=String(copy.event.getDescription() || '');
    const desired=tmv3_preInspectionEnsureLegacyTitle_(before,legacy);
    if (desired!==before) {copy.event.setDescription(desired);writes++;}
    const found=tmv3_findEventCopyRobust_(copy.calendarId,record.eventId);
    if (!found || !found.event || String(found.event.getDescription() || '')!==desired) {
      throw new Error('LEGACY_TITLE_REPAIR_READBACK_FAILED: '+copy.calendarId);
    }
  });
  return {status:'LEGACY_TITLE_VERIFIED',eventId:record.eventId,legacyTitle:legacy,writes:writes};
}

// Original: 60_Automation; Function: tmv3_restoreRecordedPreInspectionLegacyTitles
function tmv3_restoreRecordedPreInspectionLegacyTitles() {
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(30000)) return {status:'LOCK_BUSY',writes:0};
  try {
    const states=tmv3_titleMigrationStateIndex_();
    const result={status:'COMPLETE',writes:0,results:[]};
    Object.keys(states).filter(function(key){return key.indexOf('PreInspection|')===0;}).forEach(function(key){
      const id=key.slice('PreInspection|'.length);
      const row=states[key];
      if (!tmv3_clean_(row['Title Migration Map'])) return;
      try {
        const record=tmv3_preInspectionRecoveryRecord_(id);
        const one=tmv3_restorePreInspectionLegacyTitle_(record);
        result.results.push(one);result.writes+=one.writes;
      } catch(err) {result.status='ATTENTION';result.results.push({eventId:id,error:String(err)});}
    });
    Logger.log(JSON.stringify(result));return result;
  } finally {lock.releaseLock();}
}

// Original: 60_Automation; Function: tmv3_preInspectionCompletionVerified_
function tmv3_preInspectionCompletionVerified_(plans, execution) {
  if (!Array.isArray(plans) || plans.length !== 1) return false;
  const plan = plans[0];
  if (tmv3_preInspectionHistoricalCoverageVerified_(plan)) return true;
  if (!plan || plan.blocker || !Array.isArray(plan.actions)) return false;
  if (plan.actions.length === 0) return true;
  if (plan.plan !== 'NO_CHANGE' || plan.actions.some(function(action) {
    return action !== TMV3_STEP7_ACTION.VERIFY_CALENDAR_LINKS;
  })) return false;
  return !!(execution &&
    ['VERIFIED_CONVERGENCE','CREATED_VERIFIED_AND_CONVERGED'].indexOf(execution.status) >= 0 &&
    execution.calendarLinks &&
    execution.calendarLinks.status === 'CALENDAR_LINKS_VERIFIED');
}

// Original: 60_Automation; Function: tmv3_preInspectionAutomaticRecovery_
function tmv3_preInspectionAutomaticRecovery_(reason) {
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
  if (!plan || plan.vertical !== 'PreInspection' ||
      plan.disposition !== 'FULFILLED_NO_RECREATE' ||
      plan.plan !== 'VERIFY_ONLY_HISTORY_NO_WRITE' ||
      plan.readStatus !== 'FRESH_TASK_GET' ||
      !/^[1-9]\d*$/.test(String(plan.taskId || '')) ||
      !Array.isArray(plan.actions) || plan.actions.length !== 0 ||
      !tmv3_step6IsFulfilledStatus_(plan.taskStatus) ||
      plan.customerCheck !== 'MATCH' ||
      !/^[1-9]\d*$/.test(String(plan.expectedCustomerId || '')) ||
      String(plan.expectedCustomerId) !== String(plan.actualCustomerId)) return false;
  return tmv3_step6TaskTouchesEventDay_({start:plan.calendarStart},
    {Start:plan.taskStart,Due:plan.taskDue});
}

// Original: 60_Automation; Function: tmv3_reconcilePreInspectionPendingReview
function tmv3_reconcilePreInspectionPendingReview() {
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
function tmv3_preInspectionTodayRunNext(excludeEventIds,targetDate) {
  const before=tmv3_strivenApiBudgetStatus_();
  if (before.exhausted || !Number.isFinite(before.remaining) || before.remaining<100) {
    const gated={status:'API_BUDGET_GATED',remaining:before.remaining,
      reason:'Minimum 100-call readback headroom is required before another event transaction.'};
    Logger.log(JSON.stringify(gated));return gated;
  }
  const records=tmv3_preInspectionTodayRecords_(targetDate);
  const sheet=tmv3_sheet_(TMV3.SHEETS.PREINSPECTION);
  const last=sheet.getLastRow();
  const rows=last>1?sheet.getRange(2,6,last-1,10).getDisplayValues():[];
  const queue=tmv3_preInspectionTodayEligibleQueue_(records,rows);
  const skipIds={};
  (Array.isArray(excludeEventIds)?excludeEventIds:[]).forEach(function(id){skipIds[String(id)]=true;});
  if (!queue.eligible.length) {
    const empty={status:'NO_ELIGIBLE_EVENTS',events:records.length,excluded:queue.excluded,strivenCalls:0};
    Logger.log(JSON.stringify(empty));return empty;
  }
  const props=PropertiesService.getScriptProperties();
  const day=targetDate || Utilities.formatDate(new Date(),TMV3_TIMEZONE,'yyyy-MM-dd');
  const key='TMV3_PI_ELIGIBLE_CURSOR_'+day;
  const cursor=Array.isArray(excludeEventIds)&&excludeEventIds.length
    ? 0 : Number(props.getProperty(key)||0);
  let recent=0;
  for(let i=0;i<queue.eligible.length;i++){
    const index=(cursor+i)%queue.eligible.length;
    const candidate=queue.eligible[index],record=candidate.record;
    if (skipIds[record.eventId]) continue;
    const done=JSON.parse(props.getProperty('TMV3_PI_DONE_'+tmv3_hash_(record.eventId))||'null');
    if (done && done.fingerprint===record.fingerprint &&
        Date.now()-Number(done.at||0)<3600000){recent++;continue;}
    // Advance even for blocked events; do not retry the same blocker forever.
    props.setProperty(key,String((index+1)%queue.eligible.length));
    const result=tmv3_preInspectionAutomationCycle_({
      force:true,eventId:record.eventId,reason:'TODAY_ELIGIBLE_SINGLE_EVENT'});
    const one=(result.results||[]).filter(function(r){
      return r.eventId===record.eventId;
    })[0]||{};
    const after=one.afterPlans||[];
    const taskId=tmv3_clean_(one.taskId||
      (one.execution&&one.execution.taskId)||
      (after[0]&&after[0].taskId));
    const complete=result.status==='COMPLETE'&&!!taskId&&
      tmv3_preInspectionCompletionVerified_(after,one.execution);
    const response={
      status:complete?'EVENT_RECONCILED_AND_VERIFIED':'EVENT_NOT_CONVERGED',
      eventId:record.eventId,title:record.title,
      candidate: candidate.disposition,taskId:taskId,
      processingStatus:result.status,
      executionStatus:one.execution&&one.execution.status||'NO_EXECUTION',
      after:after,optionalOrderEvidence:one.optionalOrderEvidence||null,
      blocker:one.blocker||'',error:one.error||result.error||'',
      before:before.count,afterCalls:tmv3_strivenApiBudgetStatus_().count,
      excluded:queue.excluded
    };
    Logger.log(JSON.stringify(response));return response;
  }
  const doneResult={status:'ALL_ELIGIBLE_RECENTLY_VERIFIED',
    events:queue.eligible.length,recent:recent,excluded:queue.excluded,strivenCalls:0};
  Logger.log(JSON.stringify(doneResult));return doneResult;
}

// Original: 90_Admin; Function: tmv3_preInspectionTodayRunEligibleBatch
function tmv3_preInspectionTodayRunEligibleBatch() {
  const results=[],seen={};
  for(let i=0;i<3;i++){
    const r=tmv3_preInspectionTodayRunNext(Object.keys(seen));
    if (!r.eventId){if (!results.length) results.push(r);break;}
    if (seen[r.eventId]){results.push({status:'CURSOR_DID_NOT_ADVANCE',eventId:r.eventId});break;}
    seen[r.eventId]=true;
    results.push(r);
  }
  const verified=results.filter(function(r){return r.status==='EVENT_RECONCILED_AND_VERIFIED';}).length;
  const result={status:verified===results.length?'ALL_SELECTED_VERIFIED':'PARTIAL_OR_BLOCKED',
    verified:verified,attempted:results.length,results:results};
  Logger.log(JSON.stringify(result));return result;
}

// Original: 90_Admin; Function: tmv3_preInspectionTodayLocationEvidence
function tmv3_preInspectionTodayLocationEvidence() {
  const before=tmv3_strivenApiBudgetStatus_();
  if (before.exhausted || !Number.isFinite(before.remaining) || before.remaining<100) throw new Error('INSUFFICIENT_READBACK_HEADROOM');
  const results=tmv3_preInspectionTodayRecords_().map(function(record){
    return {eventId:record.eventId,title:record.title,evidence:tmv3_preInspectionOptionalOrderLocationEvidence_(record)};
  });
  Logger.log(JSON.stringify({results:results,before:before.count,afterCalls:tmv3_strivenApiBudgetStatus_().count}));
}

// Original: 90_Admin; Function: tmv3_preInspectionTodayRunUnlinked
function tmv3_preInspectionTodayRunUnlinked() {
  const before=tmv3_strivenApiBudgetStatus_();
  if (before.exhausted || !Number.isFinite(before.remaining) || before.remaining<100) throw new Error('INSUFFICIENT_READBACK_HEADROOM');
  const records=tmv3_preInspectionTodayRecords_();
  const rows=tmv3_sheet_(TMV3.SHEETS.PREINSPECTION).getDataRange().getDisplayValues().slice(1);
  const props=PropertiesService.getScriptProperties();
  const candidates=records.map(function(record){
    const matches=rows.filter(function(row){return row[14]===record.eventId;});
    if(matches.length!==1 || matches[0][24]) return null;
    const done=JSON.parse(props.getProperty('TMV3_PI_DONE_'+tmv3_hash_(record.eventId))||'null');
    if(done && done.fingerprint===record.fingerprint && Date.now()-Number(done.at||0)<3600000) return null;
    return {record:record,priority:/CREATE_TASK/.test(matches[0][9])?0:1};
  }).filter(Boolean).sort(function(a,b){return a.priority-b.priority || new Date(a.record.start)-new Date(b.record.start);});
  if(!candidates.length){Logger.log(JSON.stringify({status:'NO_UNLINKED_EVENTS',strivenCalls:0}));return;}
  const record=candidates[0].record;
  const result=tmv3_preInspectionAutomationCycle_({force:true,eventId:record.eventId,reason:'TODAY_UNLINKED_EVENT'});
  const one=(result.results||[])[0]||{},execution=one.execution||{};
  Logger.log(JSON.stringify({status:result.status,eventId:record.eventId,title:record.title,
    taskId:execution.taskId||one.taskId||'',executionStatus:execution.status||'NO_EXECUTION',
    after:one.afterPlans||[],checks:execution.readbackPlan?tmv3_preInspectionCompactPlan_(execution.readbackPlan):null,
    optionalOrderEvidence:one.optionalOrderEvidence||null,error:one.error||result.error||'',
    before:before.count,afterCalls:tmv3_strivenApiBudgetStatus_().count}));
  return result;
}

// Original: 90_Admin; Function: tmv3_preInspectionTodayLocationResponseShape
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
function tmv3_preInspectionTodayRefreshTaskCacheSchedules(targetDate) {
  const before=tmv3_strivenApiBudgetStatus_();
  if(before.exhausted || !Number.isFinite(before.remaining) || before.remaining<100) throw new Error('INSUFFICIENT_READBACK_HEADROOM');
  const records=tmv3_preInspectionTodayRecords_(targetDate),rows=tmv3_sheet_(TMV3.SHEETS.PREINSPECTION).getDataRange().getDisplayValues().slice(1);
  const results=[],cachedTasks=tmv3_rows_(TMV3.SHEETS.TASKS);
  records.forEach(function(record){
    const matches=rows.filter(function(row){return row[14]===record.eventId;});
    if(matches.length!==1) return;
    const found=String(matches[0][24]||'').match(/[?&]TaskID=([1-9][0-9]*)/i);
    const id=tmv3_clean_(record.existingTaskId||record.taskNumber||(found&&found[1]));
    if(!/^[1-9][0-9]*$/.test(id)) return;
    if(found && found[1]!==id) throw new Error('TASK_LINK_CACHE_CONFLICT');
    const cached=cachedTasks.filter(function(task){return String(task['Task ID'])===id;});
    if(cached.length>1) throw new Error('DUPLICATE_TASK_CACHE_ROWS');
    if(cached.length===1 && tmv3_sameMinute_(record.start,cached[0].Start) && tmv3_sameMinute_(record.end,cached[0].Due)) {
      results.push({taskId:id,status:'ALREADY_CACHED_MATCH'});return;
    }
    const task=tmv3_getTaskById_(Number(id));
    if(Number(task['Task Type ID'])!==105 || String(task['Task ID'])!==id) throw new Error('PI_TASK_CACHE_IDENTITY_UNVERIFIED');
    tmv3_upsertTaskCacheRow_(task);
    results.push({taskId:id,status:'CANONICAL_V1_SCHEDULE_CACHED'});
  });
  tmv3_step6TaskDecisionRunCached('PREINSPECTION_CANONICAL_CACHE_READBACK');
  Logger.log(JSON.stringify({results:results,before:before.count,afterCalls:tmv3_strivenApiBudgetStatus_().count,strivenWrites:0}));
}

// Original: 90_Admin; Function: tmv3_preInspectionOct21Preview
function tmv3_preInspectionOct21Preview() { return tmv3_preInspectionTodayPreview("2026-10-21"); }

// Original: 90_Admin; Function: tmv3_preInspectionOct21RunNext
function tmv3_preInspectionOct21RunNext() { return tmv3_preInspectionTodayRunNext([],"2026-10-21"); }

// Original: 90_Admin; Function: tmv3_preInspectionOct21RefreshCache
function tmv3_preInspectionOct21RefreshCache() { return tmv3_preInspectionTodayRefreshTaskCacheSchedules("2026-10-21"); }

// Original: 90_Admin; Function: tmv3_preInspectionLastWeekAudit
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
function tmv3_preInspectionRepairNov9Stephen() {
  const today=Utilities.formatDate(new Date(),TMV3_TIMEZONE,'yyyy-MM-dd');
  if(today!=='2026-11-09')return {status:'DATE_GUARD_NO_WRITES',today:today};
  const targets=[
    {id:'1h35gn1jffpmllhof1467gi00a@google.com',number:'63199',name:'Stephanie & Kyle Tomlin'},
    {id:'3n8bshit4qn9mfjtahdbivk94u@google.com',number:'63202',name:'Emilie & John Stamers'}
  ];
  const out={status:'IN_PROGRESS',date:today,events:[],writesAttempted:0};
  for(let i=0;i<targets.length;i++){
    const t=targets[i],entry={eventId:t.id,name:t.name};
    try{
      const record=tmv3_preInspectionRecoveryRecord_(t.id);
      if(record.calendarRole!=='PRIMARY_SHARED' ||
          Utilities.formatDate(new Date(record.start),TMV3_TIMEZONE,'yyyy-MM-dd')!==today ||
          String(record.title||'').indexOf(t.number+' - ')!==0){
        entry.status='SOURCE_IDENTITY_CHANGED_NO_WRITE';out.events.push(entry);break;
      }
      const plans=tmv3_step7FreshPlansForEvent_('PreInspection',t.id);
      if(plans.length!==1 || plans[0].blocker){
        entry.status='FRESH_PLAN_BLOCKED_NO_WRITE';entry.plan=plans[0]?plans[0].plan:'NONE';
        entry.blocker=plans[0]?String(plans[0].blocker||''):'NO_UNIQUE_PLAN';
        out.events.push(entry);break;
      }
      const p=plans[0];
      if(p.plan==='NO_CHANGE' && /^[1-9]\d*$/.test(String(p.taskId||''))){
        entry.status='ALREADY_CONVERGED';entry.taskId=String(p.taskId);
        out.events.push(entry);continue;
      }
      const allowed=Array.isArray(p.actions) && p.actions.indexOf(TMV3_STEP7_ACTION.CREATE_TASK)>=0 &&
        p.actions.indexOf(TMV3_STEP7_ACTION.RECREATE_TASK)<0 &&
        !p.taskId;
      if(!allowed){
        entry.status='NOT_SAFE_CREATE_PLAN';entry.plan=p.plan;
        out.events.push(entry);break;
      }
      out.writesAttempted++;
      // Existing cycle rechecks live Striven task history immediately before create.
      // It also guards the write with the canonical event lock and readback.
      const run=tmv3_preInspectionAutomationCycle_({force:true,eventId:t.id,reason:'NOV9_STEPPHEN_TARGETED_VERIFIED_CREATE'});
      const one=(run.results||[]).filter(function(x){return x.eventId===t.id;})[0];
      entry.cycleStatus=run.status;
      entry.executionStatus=one&&one.execution&&one.execution.status||'NO_VERIFIED_EXECUTION';
      if(!one || run.status!=='COMPLETE'){
        entry.status='UNCERTAIN_STOP_REVIEW_BEFORE_RETRY';
        entry.reason=String(one&&one.blocker||one&&one.error||run.error||'').slice(0,350);
        out.events.push(entry);break;
      }
      const fresh=tmv3_step7FreshPlansForEvent_('PreInspection',t.id);
      entry.taskId=fresh.length===1?String(fresh[0].taskId||''):'';
      entry.status=fresh.length===1 && /^[1-9]\d*$/.test(entry.taskId) &&
        tmv3_preInspectionCompletionVerified_(fresh,one.execution)
        ?'TASK_CREATED_AND_READBACK_VERIFIED':'INCOMPLETE_STOP_REVIEW_BEFORE_RETRY';
      out.events.push(entry);
      if(entry.status!=='TASK_CREATED_AND_READBACK_VERIFIED')break;
    }catch(error){
      entry.status='ERROR_STOP_REVIEW_BEFORE_RETRY';
      entry.reason=String(error&&error.message||error).slice(0,400);
      out.events.push(entry);break;
    }
  }
  out.status=out.events.length===targets.length &&
    out.events.every(function(x){return x.status==='ALREADY_CONVERGED'||x.status==='TASK_CREATED_AND_READBACK_VERIFIED';})
    ?'TWO_TARGETS_VERIFIED':'PARTIAL_ATTENTION';
  Logger.log('PI_NOV9_TARGETED_REPAIR='+JSON.stringify(out));
  return out;
}


/**
 * Targeted, explicit Nov 9 three-event reconciliation.
 * Existing production V3 guards own all Striven mutations; this wrapper is
 * fail-closed, idempotent, and never sends a generic missing-details email.
 * Creator notifications must use tmv3_preInspectionMissingDetailEmailPreview_.
 */
function tmv3_preInspectionRepairThreeAddedByOthersNov9() {
  const today=Utilities.formatDate(new Date(),TMV3_TIMEZONE,'yyyy-MM-dd');
  if(today!=='2026-11-09')return {status:'DATE_GUARD_NO_WRITES',today:today};
  const cfg=TMV3.VERTICALS.PreInspection;
  const ownerCalendar=(cfg.secondaryCalendarIds||[])[0];
  const targets=[
    {id:'1cmb7i5r3i16esfk2q05betmd4@google.com',number:'63196',title:'63196 - Nick Rogers - (416) 988-4212',start:'12:00'},
    {id:'1p8inlluovcrlm1ei32mjafocr@google.com',number:'63200',title:'63200 - Laura Menard - (416) 876-4411',start:'13:00'},
    {id:'3hv5sv78a5ub2vsmgl245llqsl@google.com',number:'63123',title:'63123 - Bill and Zena Klement - (416) 286-6700',start:'14:00'}
  ];
  const out={status:'IN_PROGRESS',date:today,events:[],writeCyclesAttempted:0,notificationsSent:0};
  for(let i=0;i<targets.length;i++){
    const t=targets[i],line={eventId:t.id,customerNumber:t.number};
    try{
      const record=tmv3_findFreshEventRecord_('PreInspection',t.id);
      const start=Utilities.formatDate(new Date(record.start),TMV3_TIMEZONE,'yyyy-MM-dd HH:mm');
      if(record.calendarRole!=='PRIMARY_SHARED'||
         String(record.eventId)!==t.id||
         String(record.customerNumber||'')!==t.number||
         String(record.title)!==t.title||
         start!==today+' '+t.start||
         String(record.organizerCalendarId||'')!==String(ownerCalendar)||
         !tmv3_clean_(record.location)){
        line.status='IDENTITY_OWNER_SCHEDULE_OR_LOCATION_NOT_VERIFIED_NO_WRITE';
        line.found={title:record.title,start:start,owner:record.organizerCalendarId,hasLocation:!!record.location};
        out.events.push(line);break;
      }
      const budget=tmv3_strivenApiBudgetStatus_();
      if(budget.exhausted||!Number.isFinite(budget.remaining)||budget.remaining<100){
        line.status='API_READBACK_BUDGET_GATED';out.events.push(line);break;
      }
      const initial=tmv3_step7FreshPlansForEvent_('PreInspection',t.id);
      if(initial.length!==1||initial[0].blocker){
        line.status='FRESH_TASK_OR_LOCATION_REVIEW_REQUIRED';
        line.plan=initial[0]&&initial[0].plan||'NONE';
        line.blocker=String(initial[0]&&initial[0].blocker||'NO_UNIQUE_PLAN').slice(0,300);
        out.events.push(line);
        // Independent appointments continue; no Task write on this one.
        continue;
      }
      const plan=initial[0];
      if(plan.actions&&plan.actions.indexOf(TMV3_STEP7_ACTION.RECREATE_TASK)>=0){
        line.status='RECREATE_NOT_ALLOWED';out.events.push(line);continue;
      }
      const permitted=(
        plan.plan==='NO_CHANGE'&&/^[1-9]\d*$/.test(String(plan.taskId||'')))||
        (!plan.taskId&&Array.isArray(plan.actions)&&
         plan.actions.indexOf(TMV3_STEP7_ACTION.CREATE_TASK)>=0);
      if(!permitted){
        line.status='UNSAFE_OR_UNSUPPORTED_TASK_PLAN';
        line.plan=plan.plan;out.events.push(line);continue;
      }
      // Stephen's verified owner-calendar identity is anchored to each
      // immutable Google event ID before entering the existing Striven pipeline.
      tmv3_configurePreInspectionAssignmentOverride(t.id,t.number,ownerCalendar,15);
      const confirmed=tmv3_step7FreshPlansForEvent_('PreInspection',t.id);
      if(confirmed.length!==1||confirmed[0].blocker||
         (plan.taskId&&String(confirmed[0].taskId||'')!==String(plan.taskId))){
        line.status='ASSIGNMENT_PLAN_CHANGED_NO_WRITE';out.events.push(line);continue;
      }
      out.writeCyclesAttempted++;
      const result=tmv3_preInspectionAutomationCycle_({
        force:true,eventId:t.id,reason:'NOV9_THREE_ADDED_CONFIRMED_PI'
      });
      const item=(result.results||[]).filter(function(r){return r.eventId===t.id;})[0];
      if(result.status!=='COMPLETE'||!item){
        line.status='EXECUTION_UNCERTAIN_DO_NOT_RETRY';
        line.cycleStatus=result.status;
        out.events.push(line);break;
      }
      const finalPlans=tmv3_step7FreshPlansForEvent_('PreInspection',t.id);
      const finalPlan=finalPlans.length===1?finalPlans[0]:null;
      const taskId=String(finalPlan&&finalPlan.taskId||'');
      const found=tmv3_findCalendarEventById_(cfg.primaryCalendarId,t.id);
      const description=String(found&&found.event&&found.event.getDescription()||'');
      line.taskId=taskId;
      line.status=/^[1-9]\d*$/.test(taskId)&&
        description.indexOf('TaskID='+taskId)>=0&&
        finalPlan&&!finalPlan.blocker&&
        tmv3_preInspectionCompletionVerified_(finalPlans,item.execution)
        ?'TASK_AND_CALENDAR_VERIFIED':'READBACK_INCOMPLETE_DO_NOT_RETRY';
      out.events.push(line);
      if(line.status!=='TASK_AND_CALENDAR_VERIFIED')break;
    }catch(error){
      line.status='ERROR_DO_NOT_RETRY';
      line.error=String(error&&error.message||error).slice(0,350);
      out.events.push(line);break;
    }
  }
  out.status=out.events.length===targets.length&&
    out.events.every(function(x){return x.status==='TASK_AND_CALENDAR_VERIFIED';})
    ?'THREE_EVENTS_VERIFIED':'PARTIAL_ATTENTION';
  Logger.log('PI_NOV9_THREE_REPAIR='+JSON.stringify(out));
  return out;
}


/**
 * Google Apps Script — read-only Striven cache, current CF Preinspects Calendar
 * and ONLY the PreInspection operator tab. This publishes a fresh Stage-6
 * projection independently of pending Install/Delivery/Service sweeps.
 * Does NOT refresh Striven source reports or mutate Striven/Calendar records.
 */
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
  if(mode!=='PREVIEW'&&mode!=='IMPORT')throw new Error('TERRY_INTAKE_MODE_REJECTED');
  const uid='5mafnqht16o2hrqv4232rofmm3@google.com';
  const approved=tmv3_preInspectionApprovalHold_(uid);
  if(!approved||approved.customerNumber!=='53480'||approved.authorizeTaskWrites!==false||
      approved.authorizeGuestChanges!==false||approved.authorizeEmail!==false)
    throw new Error('TERRY_APPROVAL_HOLD_NOT_VERIFIED');
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(2000))return {status:'LOCK_BUSY',writes:0};
  try {
    const token=ScriptApp.getOAuthToken(),base='https://www.googleapis.com/calendar/v3/calendars/';
    const endpoint=function(cal,path){return base+encodeURIComponent(cal)+'/events'+path;};
    const call=function(method,url,data){
      const options={method:method,headers:{Authorization:'Bearer '+token},
        muteHttpExceptions:true,followRedirects:false};
      if(data){options.contentType='application/json';options.payload=JSON.stringify(data);}
      const response=UrlFetchApp.fetch(url,options),code=response.getResponseCode();
      if(code<200||code>299)throw new Error('GOOGLE_CALENDAR_'+method+'_HTTP_'+code+'_NO_RETRY');
      return JSON.parse(response.getContentText());
    };
    const source=approved.sourceCalendarId,target=approved.primaryCalendarId;
    if(target!==TMV3.VERTICALS.PreInspection.primaryCalendarId)throw new Error('CF_CALENDAR_MISMATCH');
    const epoch=function(e){return Date.parse(e&&e.dateTime||'');};
    const getOriginal=function(){
      const e=call('get',endpoint(source,'/5mafnqht16o2hrqv4232rofmm3'));
      if(e.status!=='confirmed'||e.iCalUID!==uid||e.id!=='5mafnqht16o2hrqv4232rofmm3'||
        e.summary!==approved.title||e.location!==approved.location||
        epoch(e.start)!==Date.parse(approved.start)||epoch(e.end)!==Date.parse(approved.end)||
        String(e.creator&&e.creator.email||'').toLowerCase()!==approved.creator||
        String(e.organizer&&e.organizer.email||'').toLowerCase()!==source.toLowerCase()||
        e.recurrence||e.recurringEventId||(e.attendees||[]).length||(e.attachments||[]).length||
        typeof e.description!=='string'||e.description.indexOf('H3 install')===-1||
        e.description.toLowerCase().indexOf('g3.5')===-1)
        throw new Error('TERRY_SOURCE_IDENTITY_CONTENT_OR_GUEST_DRIFT');
      return e;
    };
    const readCopy=function(e){
      const found=call('get',endpoint(target,'?iCalUID='+encodeURIComponent(uid)+'&maxResults=20'));
      const items=found.items||[];
      if(found.nextPageToken||items.length>1)throw new Error('TERRY_ICALUID_DUPLICATES');
      const copy=items.length?items[0]:null;
      if(copy){
        const marker=copy.extendedProperties&&copy.extendedProperties.private||{};
        if(copy.iCalUID!==uid||copy.status==='cancelled'||copy.summary!==e.summary||
          copy.description!==e.description||copy.location!==e.location||
          epoch(copy.start)!==epoch(e.start)||epoch(copy.end)!==epoch(e.end)||
          (copy.attendees||[]).length||copy.privateCopy!==true||
          marker.tmv3TaskWriteHold!=='YES'||marker.tmv3SourceEventUid!==uid)
          throw new Error('TERRY_SHARED_COPY_DIFFERENT_OR_UNVERIFIED_NO_RETRY');
      } else {
        const near=call('get',endpoint(target,'?timeMin='+encodeURIComponent('2026-11-11T19:50:00Z')+
          '&timeMax='+encodeURIComponent('2026-11-11T21:10:00Z')+'&singleEvents=true&maxResults=250'));
        if(near.nextPageToken||(near.items||[]).some(function(x){
          return x.status!=='cancelled'&&x.iCalUID!==uid&&
            (String(x.summary||'').indexOf('53480')>=0||
             String(x.summary||'').toLowerCase().indexOf('terry beettam')>=0||
             String(x.summary||'').indexOf('523-9364')>=0);
        }))throw new Error('TERRY_POSSIBLE_CF_DUPLICATE_NO_IMPORT');
      }
      return copy;
    };
    let origin=getOriginal(),copy=readCopy(origin);
    if(mode==='PREVIEW'||copy){
      const status=copy?'TERRY_COPY_ALREADY_VERIFIED':'TERRY_SAFE_IMPORT_CANDIDATE';
      const result={status:status,eventId:uid,copyId:copy?copy.id:'',
        guestWrites:0,emailSends:0,strivenWrites:0,taskCreates:0};
      Logger.log('PI_TERRY_INTAKE='+JSON.stringify(result));return result;
    }
    // Revalidate immediately before the one permitted import.
    origin=getOriginal();if(readCopy(origin))throw new Error('TERRY_COPY_RACE_NO_RETRY');
    const body={iCalUID:uid,summary:origin.summary,location:origin.location,
      description:origin.description,start:origin.start,end:origin.end,
      organizer:origin.organizer,visibility:origin.visibility||'private',privateCopy:true,
      reminders:{useDefault:false,overrides:[]},
      extendedProperties:{private:{tmv3CustomerNumber:'53480',tmv3TaskWriteHold:'YES',
        tmv3SourceEventUid:uid,tmv3SourceCalendarId:source}}};
    // Import copies the existing iCalUID without inviting attendees.
    // No attendees property, CalendarApp.addGuest, insert, move or MailApp.
    let response;
    try{response=call('post',endpoint(target,'/import'),body);}
    catch(err){throw new Error('TERRY_IMPORT_OUTCOME_UNKNOWN_DO_NOT_RETRY_'+err.message);}
    if(response.iCalUID!==uid)throw new Error('TERRY_IMPORT_IDENTITY_UNCERTAIN_NO_RETRY');
    copy=readCopy(origin);
    if(!copy)throw new Error('TERRY_IMPORT_READBACK_NOT_VISIBLE_NO_RETRY');
    const after=getOriginal();
    if(after.summary!==origin.summary||after.description!==origin.description||
       after.location!==origin.location||JSON.stringify(after.attendees||[])!==JSON.stringify(origin.attendees||[]))
      throw new Error('TERRY_ORIGINAL_CHANGED_STOP');
    const result={status:'TERRY_PRIVATE_CF_COPY_VERIFIED',eventId:uid,copyId:copy.id,
      sourceUnchanged:true,notesPreserved:true,calendarWrites:1,
      guestInvites:0,emailSends:0,strivenWrites:0,taskCreates:0,taskWriteHold:true};
    Logger.log('PI_TERRY_INTAKE='+JSON.stringify(result));return result;
  }finally{lock.releaseLock();}
}
function tmv3_previewTerryQuietCalendarIntake() {
  return tmv3_terryQuietCalendarIntake_('PREVIEW');
}
function tmv3_importTerryQuietCalendarIntake() {
  return tmv3_terryQuietCalendarIntake_('IMPORT');
}


// PREINSPECTION LOCK 2026-11-10-r1.
// Canonical rules: docs/workflows/preinspection/STANDARD_2026_11_10.md
// Never fetch a rule from GitHub or ask an agent to interpret it at runtime.
function tmv3_piRendererFingerprint_() {
  const source = Function.prototype.toString.call(tmv3_preInspectionMissingDetailEmailPreview_);
  let hash = 2166136261;
  for (let i = 0; i < source.length; i++) {
    hash = Math.imul(hash ^ source.charCodeAt(i), 16777619) >>> 0;
  }
  return ('00000000' + hash.toString(16)).slice(-8);
}

function tmv3_piAssertLockedStandard_() {
  const root = TMV3_HARD_RULES || {};
  const r = root.PreInspection || {};
  const live = (TMV3.VERTICALS || {}).PreInspection || {};
  if (root.version !== '2026-11-10-r1' || r.lockedStandardId !== '2026-11-10-r1') {
    throw new Error('PI_STANDARD_VERSION_MISMATCH_NO_WRITE');
  }
  if (r.lockedStandardReference !== 'docs/workflows/preinspection/STANDARD_2026_11_10.md') {
    throw new Error('PI_STANDARD_REFERENCE_MISMATCH_NO_WRITE');
  }
  if (Number(r.taskTypeId) !== 105 || Number(live.defaultPoolId) !== 8 ||
      r.orderRequired !== false || r.attachOrderToTask !== false ||
      r.field854Policy !== 'DO_NOT_MANAGE' || r.infoCustomFieldsAtCreate !== 'NONE' ||
      r.calendarNotesPolicy !== 'CALENDAR_ONLY' || live.descriptionMustBeBlankAtCreate !== true ||
      live.requestedByFromOrganizer !== true) {
    throw new Error('PI_TASK_RULES_MISMATCH_NO_WRITE');
  }
  if (r.ownerGuestPolicy !== 'VERIFIED_EVENT_OWNER_ONLY') {
    throw new Error('PI_OWNER_GUEST_POLICY_DRIFT_NO_WRITE');
  }
  if (JSON.stringify((r.inspectorOwnedFields || []).slice().sort()) !== JSON.stringify([852,853,854,860])) {
    throw new Error('PI_INSPECTOR_FIELDS_POLICY_DRIFT_NO_WRITE');
  }
  if (JSON.stringify((TMV3.OPERATIONS.preInspectionAutomationScheduleHours || [])) !== JSON.stringify([8,13,17]) ||
      TMV3.OPERATIONS.preInspectionAutomationScheduleTimezone !== 'America/Toronto') {
    throw new Error('PI_NOTIFICATION_SCHEDULE_DRIFT_NO_WRITE');
  }
  if (r.notificationTemplatePolicy !== 'EXACT_EXISTING_RENDERER_NO_REDRAFT' ||
      r.notificationRendererFingerprint !== tmv3_piRendererFingerprint_()) {
    throw new Error('PI_NOTIFICATION_TEMPLATE_DRIFT_NO_SEND');
  }
  if (r.presentation.linkHeading !== '-----Striven Links-----' ||
      r.presentation.calendarTitleParts.join('|') !== 'customerNumber|customerName|phone' ||
      r.presentation.taskTitleParts.join('|') !== 'customerName|address|phone') {
    throw new Error('PI_CALENDAR_OR_TASK_FORMAT_DRIFT_NO_WRITE');
  }
  return {status:'PASS',standard:'2026-11-10-r1',
    reference:r.lockedStandardReference,notificationFingerprint:r.notificationRendererFingerprint};
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
  const p = payload || {};
  const forbidden = /^(?:InfoCustomFields|Description|SalesOrder|SalesOrderID|SalesOrderId|SOId|Type)$/i;
  Object.keys(p).forEach(function(k) {
    if (forbidden.test(k)) throw new Error('PI_PROTECTED_TASK_FIELD_NO_WRITE:' + k);
  });
  if (p.RequestedBy && String(p.RequestedBy.Type || '').toLowerCase() !== 'employee') {
    throw new Error('PI_REQUESTED_BY_MUST_BE_VERIFIED_EMPLOYEE_NO_WRITE');
  }
  if (p.Title !== undefined) tmv3_piAssertTaskTitle_(p.Title);
  return true;
}

function tmv3_preInspectionLockedStandardStatus() {
  return tmv3_piAssertLockedStandard_();
}


// Only positively verified owner -> employee mappings may authorize writes.
// Jay and Karen calendar IDs are not yet certified in this staging source.
function tmv3_piVerifiedOwnerEmployee_(eventRecord) {
  const record = eventRecord || {};
  const cfg = TMV3.VERTICALS.PreInspection || {};
  const ownerCalendar = String(record.organizerCalendarId || '').trim();
  const stephenCalendar = String((cfg.secondaryCalendarIds || [])[0] || '').trim();
  if (ownerCalendar && stephenCalendar && ownerCalendar === stephenCalendar) {
    return {status:'MATCHED',employeeId:15,employeeName:'Stephen Foley',
      evidence:'VERIFIED_STEPHEN_CALENDAR_ID'};
  }
  return {status:'REVIEW',employeeId:null,
    blocker:ownerCalendar ? 'PI_CALENDAR_OWNER_NOT_VERIFIED_IN_EXECUTION_CONFIG' :
      'PI_CALENDAR_OWNER_MISSING_NO_WRITE'};
}

/**
 * Pre-Inspection rebuild: pure decision functions, not wired to the live worker.
 * Facts must come from verified source reads; these helpers never perform writes.
 */
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
    if (candidate.addressMatch === true) matches.push(candidate);
    else if (candidate.addressMatch !== false) return {decision:'REVIEW_ADDRESS_MATCH_UNVERIFIED'};
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
      if (mode==='CREATE' && /^(salesorder|salesorderid|soid|orderid)$/.test(normalized)) return 'REJECT_ORDER_ATTACHMENT';
      const nested=visit(value[key]);if(nested)return nested;
    }
    seen.pop();return '';
  }
  const violation=visit(payload);return {decision:violation||'ACCEPT'};
}
