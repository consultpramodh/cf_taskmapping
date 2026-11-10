// MIGRATION STAGING ONLY. Source: reconstructed V3 3.12.21. NOT APPROVED FOR DEPLOYMENT.

// Original: 30_Calendar; Function: tmv3_extractLegacyDeliveryOrderNumber_
function tmv3_extractLegacyDeliveryOrderNumber_(text) {
  const clean = tmv3_stripPhonesForOrderParsing_(text);
  if (!clean) return '';

  const patterns = [
    /\bSO\s*#?\s*(\d{5,8})\b/i,
    /\bS\/O\s*#?\s*(\d{5,8})\b/i,
    /\bSales\s*Order\s*#?\s*(\d{5,8})\b/i,
    /\bOrder\s*#?\s*(\d{5,8})\b/i,
    /#\s*(\d{5,8})\b/i,
    /\b(\d{6})\b/
  ];

  for (let i = 0; i < patterns.length; i++) {
    const match = clean.match(patterns[i]);
    if (match && match[1]) return match[1];
  }

  return '';
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
function tmv3_deliverySweepContinue(){return {status:'RETIRED_OUTSIDE_STANDARD_SLOTS',writes:0};}