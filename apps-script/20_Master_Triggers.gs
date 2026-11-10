// MIGRATION STAGING ONLY. Source: reconstructed V3 3.12.21. NOT APPROVED FOR DEPLOYMENT.

// Original: 60_Automation; Function: tmv3_removeTriggers
function tmv3_removeTriggers() {
  const managed = tmv3_managedTriggerHandlers_();
  let deleted = 0;

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (managed.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
      deleted++;
    }
  });

  tmv3_audit_(
    'SYSTEM','','','REMOVE_MANAGED_TRIGGERS','PASS',
    'Deleted ' + deleted + ' managed trigger(s).'
  );
  return tmv3_listTriggers();
}

// Original: 60_Automation; Function: tmv3_listTriggers
function tmv3_listTriggers() {
  return ScriptApp.getProjectTriggers().map(function(t) {
    return {
      handler: t.getHandlerFunction(),
      source: String(t.getTriggerSource()),
      eventType: String(t.getEventType()),
      sourceId: tmv3_safeCalendar_(function() { return t.getTriggerSourceId(); }, '')
    };
  });
}

// Original: 60_Automation; Function: tmv3_logTriggers
function tmv3_logTriggers() {
  const result = tmv3_listTriggers();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// Original: 60_Automation; Function: tmv3_allVerticalAutomationTick
function tmv3_allVerticalAutomationTick() {
  const hour=Number(Utilities.formatDate(new Date(),TMV3_TIMEZONE,'H'));
  if(hour<8 || hour>18) return {status:'OUTSIDE_BUSINESS_HOURS',writes:0};
  const props=PropertiesService.getScriptProperties(),all=['Service','Delivery','Install'];
  const start=Number(props.getProperty('TMV3_VERTICAL_ROTATION')||0)%all.length;
  props.setProperty('TMV3_VERTICAL_ROTATION',String((start+1)%all.length));
  const result={status:'ACTIVE',results:[]},began=Date.now();
  for(let i=0;i<all.length;i++){
    if(Date.now()-began>210000)break;
    try{result.results.push(tmv3_verticalProductionCycle_(all[(start+i)%all.length]));}
    catch(err){result.results.push({vertical:all[(start+i)%all.length],status:'ATTENTION',error:String(err)});}
  }
  result.sheetRefresh=tmv3_refreshAutomationSheetsAfter_(began);
  Logger.log(JSON.stringify(result));return result;
}

// Original: 60_Automation; Function: tmv3_scheduledSheetRefresh
function tmv3_scheduledSheetRefresh() {
  const hour=Number(Utilities.formatDate(new Date(),TMV3_TIMEZONE,'H'));
  if([8,13,17].indexOf(hour)<0)return {status:'OUTSIDE_STANDARD_SLOTS'};
  return tmv3_refreshAutomationSheets();
}

// Original: 60_Automation; Function: tmv3_standardDailySchedule_
function tmv3_standardDailySchedule_() {
  return {hours:[8,13,17],timezone:'America/Toronto',verticals:['PreInspection','Install','Delivery','Service']};
}

// Original: 60_Automation; Function: tmv3_standardVerticalSlot_
function tmv3_standardVerticalSlot_(vertical,slot) {
  const result=tmv3_runFullVerticalSweep_(vertical);
  if(vertical==='Install') {
    result.reminders=tmv3_installReminderCheck();
  }
  tmv3_audit_(vertical,'','','STANDARD_DAILY_SLOT',result.status,JSON.stringify({slot:slot,result:result}));
  return result;
}