// MIGRATION STAGING ONLY. Source: reconstructed V3 3.12.21. NOT APPROVED FOR DEPLOYMENT.

// Original: 30_Calendar; Function: tmv3_extractLegacyDeliveryOrderNumber_
// Fix: labelled Sales Order forms (SO#, SO #, SO-, SO:, S/O, Sales Order No.)
// win outright. Customer / Task / Work Order / PO / Invoice / Quote numbers and
// URLs are removed before any unlabelled fallback, and the unlabelled "#12345"
// and bare 6-digit fallbacks only resolve when exactly one distinct candidate
// exists. An ambiguous title returns '' so the event goes to REVIEW instead of
// being matched to a guessed Sales Order.
function tmv3_extractLegacyDeliveryOrderNumber_(text) {
  let clean = tmv3_stripPhonesForOrderParsing_(text);
  if (!clean) return '';

  // Striven URLs carry internal record IDs, never the Sales Order number.
  clean = clean.replace(/\bhttps?:\/\/\S+/gi, ' ');

  const labelled = clean.match(
    /\b(?:SO|S\/O|Sales\s*Order)\s*(?:#|No\.?|Number)?\s*[#:\-]?\s*(\d{4,8})(?!\d)/i
  );
  if (labelled) return labelled[1];

  clean = clean
    .replace(
      /\b(?:Customer|Cust|C(?=\s*#)|Task|Work\s*Order|W\/?O|Purchase\s*Order|P\/?O|Invoice|Inv|Quote|Estimate)\s*(?:#|No\.?|Number)?\s*[#:\-]?\s*\d{3,10}(?!\d)/gi,
      ' '
    )
    .replace(/\s+/g, ' ')
    .trim();

  const order = clean.match(/\bOrder\s*(?:#|No\.?|Number)?\s*[#:\-]?\s*(\d{5,8})(?!\d)/i);
  if (order) return order[1];

  const unique = function(regex) {
    const found = [];
    let m;
    while ((m = regex.exec(clean)) !== null) {
      if (found.indexOf(m[1]) < 0) found.push(m[1]);
    }
    return found.length === 1 ? found[0] : (found.length ? null : '');
  };

  const hashed = unique(/#\s*(\d{5,8})(?!\d)/g);
  if (hashed === null) return '';
  if (hashed) return hashed;

  const bare = unique(/(?:^|[^\d#])(\d{6})(?!\d)/g);
  return bare || '';
}

// Original: 30_Calendar; Function: tmv3_deliveryLegacyCustomerName_
function tmv3_deliveryLegacyCustomerName_(title) {
  let clean = tmv3_clean_(title);
  if (!clean) return '';

  clean = clean
    .replace(
      /(?:\+?1[\s\-.]?)?(?:\(?\d{3}\)?[\s\-.]?\d{3}[\s\-.]?\d{4})(?:\s*(?:x|ext\.?)\s*\d+)?/gi,
      ' '
    )
    .replace(/\b(?:SO|S\/O|Sales\s*Order)\s*[#:\-]?\s*\d{4,8}\b/gi, ' ')
    .replace(/#\s*\d{4,8}\b/g, ' ')
    .replace(
      /\b\d{1,2}(?::\d{2})?\s*(?:am|pm)?\s*[-–]\s*\d{1,2}(?::\d{2})?\s*(?:am|pm)?\b/gi,
      ' '
    )
    .replace(/\b\d{1,2}(?::\d{2})?\s*(?:am|pm)\b/gi, ' ')
    .replace(/\bjohn\s+hoang\b/gi, ' ')
    .replace(/\bjohn\b/gi, ' ')
    .replace(/\bjay\s+scott\b/gi, ' ')
    .replace(/\bjay\b/gi, ' ')
    .replace(/\bstephen\s+foley\b/gi, ' ')
    .replace(/\bsteven\s+foley\b/gi, ' ')
    .replace(/\bstephen\b/gi, ' ')
    .replace(/\bsteven\b/gi, ' ')
    .replace(/\bsf\b/gi, ' ')
    .replace(/\bspencer\b/gi, ' ')
    .replace(/\bthang\b/gi, ' ')
    .replace(/\bpramodh\b/gi, ' ')
    .replace(/\bmatthew\s+thompson\b/gi, ' ')
    .replace(/\bmatthew\b/gi, ' ')
    .replace(/\bmatt\b/gi, ' ')
    .replace(/\baiden\b/gi, ' ')
    .replace(/\bdelivery\b/gi, ' ')
    .replace(/\bdeliveries\b/gi, ' ')
    .replace(/\binstall\b/gi, ' ')
    .replace(/\binstallation\b/gi, ' ')
    .replace(/\bfireplace\b/gi, ' ')
    .replace(/\bbbq\b/gi, ' ')
    .replace(/[&+/,]/g, ' ')
    .replace(/[-–]/g, ' ')
    .replace(/\band\b/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  return clean;
}

// Original: 60_Automation; Function: tmv3_runDeliveryAutomation
function tmv3_runDeliveryAutomation() {const r=tmv3_runFullVerticalSweep_('Delivery');Logger.log(JSON.stringify(r));return r;}

// Original: 60_Automation; Function: tmv3_deliveryCheck0800
function tmv3_deliveryCheck0800() { return tmv3_standardVerticalSlot_('Delivery','08:00'); }

// Original: 60_Automation; Function: tmv3_deliveryCheck1300
function tmv3_deliveryCheck1300() { return tmv3_standardVerticalSlot_('Delivery','13:00'); }

// Original: 60_Automation; Function: tmv3_deliveryCheck1700
function tmv3_deliveryCheck1700() { return tmv3_standardVerticalSlot_('Delivery','17:00'); }

// Original: 60_Automation; Function: tmv3_deliverySweepContinue
// Fix: Delivery now uses the same follow-on mechanism as Install
// (TMV3_VERTICAL_FOLLOW_ON). When a Delivery slot finds the script lock busy,
// tmv3_runFullVerticalSweep_ keeps asking for up to 2 minutes, and an
// unfinished sweep continues a minute later (max 12 follow-ons a day) instead
// of being dropped until the next 08:00 / 13:00 / 17:00 slot.
function tmv3_deliverySweepContinue(){
  // Follow-on run: continues an unfinished Delivery sweep a minute after a slot ran out of time.
  tmv3_verticalClearFollowOn_('Delivery');
  if(!TMV3_VERTICAL_FOLLOW_ON.Delivery)return {status:'RETIRED_OUTSIDE_STANDARD_SLOTS',writes:0};
  const result=tmv3_runFullVerticalSweep_('Delivery',{followOn:true});
  tmv3_audit_('Delivery','','','STANDARD_DAILY_SLOT',result.status,JSON.stringify({slot:'FOLLOW_ON',result:result}));
  return result;
}

// Fix: Delivery used to sweep whatever was already on the Delivery tab. That
// tab is only rebuilt by the shared sheet refresh, which defers while any
// vertical queue is pending, so new Delivery appointments could be missed for
// several slots. Called by tmv3_verticalProductionCycle_ only when a new
// Delivery sweep starts. Rebuilds the Delivery tab only, never another
// vertical, and performs no Striven or Calendar writes.
function tmv3_deliveryRefreshView_(reason) {
  const cfg=TMV3.VERTICALS.Delivery,start=new Date(),end=new Date();
  start.setDate(start.getDate()-Number(cfg.lookbackDays||0));start.setHours(0,0,0,0);
  end.setDate(end.getDate()+Number(cfg.lookaheadDays||365));end.setHours(23,59,59,999);
  const input=[];
  tmv3_verticalCalendars_('Delivery',cfg).forEach(function(calCfg){
    const calendar=CalendarApp.getCalendarById(calCfg.calendarId);if(!calendar)throw Error('DELIVERY_CALENDAR_UNAVAILABLE');
    const resolved=Object.assign({},calCfg,{calendarName:calendar.getName()});
    calendar.getEvents(start,end).forEach(function(event){const record=tmv3_calendarEvent_('Delivery',cfg,resolved,event,{applyEligibility:false});if(record)input.push(record);});
  });
  const records=tmv3_mergeLogicalCalendarRecords_(input);
  const stage2=tmv3_step2CalendarRecords_(records);
  const stage3=tmv3_step3BusinessAnchorRecords_(stage2,tmv3_step3AnchorIndex_());
  const stage4=tmv3_step4IdentityRecords_(stage3,tmv3_step4IdentityIndex_());
  const stage5=tmv3_step5TaskRecords_(stage4,tmv3_step5TaskIndex_());
  const decisions=tmv3_step6DecisionRecords_(stage5).sort(function(a,b){return new Date(a.start)-new Date(b.start);});
  const prior=tmv3_sheet_(cfg.sheet).getLastRow();
  tmv3_replaceRows_(cfg.sheet,TMV3_STEP1_HEADERS.slice(),decisions.map(tmv3_step6OperatorRow_));
  tmv3_step1FormatOperatorView_(cfg.sheet,'Delivery',decisions,prior);
  const count=tmv3_rows_(cfg.sheet).filter(function(r){return !!tmv3_clean_(r['Event ID']);}).length;
  if(count!==decisions.length)throw Error('DELIVERY_VIEW_ROW_READBACK_FAILED');
  const result={status:'DELIVERY_VIEW_REFRESHED',at:new Date().toISOString(),reason:reason,rows:count,strivenWrites:0,calendarWrites:0};
  PropertiesService.getScriptProperties().setProperty('TMV3_DELIVERY_VIEW_LAST_SUCCESS',JSON.stringify(result));return result;
}

// Refresh the Striven source reports, then rebuild the Delivery tab. A failure
// in either step must not stop the sweep: every queued event is re-read fresh
// from Calendar and re-verified before any write, so the sweep falls back to
// the existing cache and Delivery tab, exactly as it ran before this fix.
function tmv3_deliverySweepIntake_(reason) {
  const out={};
  try { out.sources=tmv3_refreshOperationalSourcesIfNeeded_(120,false); }
  catch(err) { out.sources=tmv3_deliveryIntakeAttention_('DELIVERY_SOURCE_REFRESH',err,'EXISTING_STRIVEN_CACHE'); }
  try { out.intake=tmv3_deliveryRefreshView_(reason); }
  catch(err) { out.intake=tmv3_deliveryIntakeAttention_('DELIVERY_INTAKE_REFRESH',err,'EXISTING_DELIVERY_TAB'); }
  return out;
}

function tmv3_deliveryIntakeAttention_(action,err,fallback) {
  const message=String(err && err.message || err);
  tmv3_audit_('Delivery','','',action,'ATTENTION',message);
  return {status:action+'_ATTENTION',error:message,fallback:fallback};
}
