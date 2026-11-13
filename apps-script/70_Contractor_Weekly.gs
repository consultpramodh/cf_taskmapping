/**
 * Contractor Weekly Check — Google Apps Script V8
 * Read-only toward Gmail, Calendar and Striven; writes a NEW tracker sheet per run.
 * Searches Kawartha first, then Steve. It does not create/update tasks, upload
 * files, send messages, or install triggers. Each tracker row is an EMAIL, not
 * a confirmed job. Mixed-job emails and PDFs require manual review.
 * Paste into a spreadsheet-bound Apps Script project and run contractorWeeklyCheck.
 */
const CWC = Object.freeze({
  timezone: 'America/Toronto',
  lookbackDays: 7, // rolling period ending at execution time
  calendarWindowDays: 30,
  pageSize: 100,
  maxThreadsPerContractor: 1000,
  maxRuntimeMs: 270000,
  contractors: [
    {name:'Kawartha / Jamie', calendarId:'classicfireplace.ca_2pbtbkjl1g2sos1fspgfq029h0@group.calendar.google.com', query:'{"Kawartha Home and Hearth" khhltd@gmail.com} -in:trash -in:spam'},
    {name:'Steve Jensen', calendarId:'steve.jensen83@gmail.com', query:'{from:steve.jensen83@gmail.com "Steve Jensen"} -in:trash -in:spam'}
  ]
});

/** Main entry point; safe to run repeatedly. */
function contractorWeeklyCheck() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('A contractor check is already running.');
  try {
    const started = new Date();
    const from = new Date(started.getTime() - CWC.lookbackDays * 86400000);
    const book = SpreadsheetApp.getActiveSpreadsheet();
    if (!book) throw new Error('Run from a spreadsheet-bound Apps Script project.');
    const base = 'Contractors ' + Utilities.formatDate(started,CWC.timezone,'yyyy-MM-dd HHmmss');
    let title = base, suffix = 2;
    while (book.getSheetByName(title)) title = base + ' ' + suffix++;
    const sheet = book.insertSheet(title);
    const headers = ['Contractor','Email date (Toronto)','Subject','From','Email link','Photos found','PDFs found','Other files','Attachment names','SO numbers found','Address hints','Calendar candidates','Calendar dates (Toronto)','Review status','Notes'];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    const totals = [];
    for (const contractor of CWC.contractors) {
      try {
        const result = cwcCheckContractor_(contractor, from, started, sheet, started.getTime());
        totals.push({contractor:contractor.name, emailRows:result.count, status:result.status});
      } catch (err) {
        cwcAppend_(sheet, [contractor.name,'','','','','','','','','','','','', 'ERROR / INCOMPLETE',String(err && err.message || err)]);
        totals.push({contractor:contractor.name, status:'ERROR / INCOMPLETE', error:String(err && err.message || err)});
      }
    }
    sheet.autoResizeColumns(1, 15);
    Logger.log(JSON.stringify({tracker:sheet.getName(), url:book.getUrl()+'#gid='+sheet.getSheetId(), from:from.toISOString(), to:started.toISOString(), totals:totals}));
  } finally { lock.releaseLock(); }
}

function cwcCheckContractor_(contractor, from, to, sheet, startedMs) {
  const calendar = CalendarApp.getCalendarById(contractor.calendarId);
  if (!calendar) throw new Error('Calendar unavailable: ' + contractor.name);
  const span = CWC.calendarWindowDays * 86400000;
  const events = calendar.getEvents(new Date(from.getTime()-span), new Date(to.getTime()+span));
  const query = contractor.query + ' after:' + Utilities.formatDate(new Date(from.getTime()-86400000),CWC.timezone,'yyyy/MM/dd') + ' before:' + Utilities.formatDate(new Date(to.getTime()+86400000),CWC.timezone,'yyyy/MM/dd');
  const seen = Object.create(null);
  let count=0, offset=0, incomplete=false;
  while (offset < CWC.maxThreadsPerContractor) {
    if (Date.now()-startedMs > CWC.maxRuntimeMs) { incomplete=true; break; }
    const threads=GmailApp.search(query,offset,CWC.pageSize);
    if (!threads.length) break;
    for (const thread of threads) {
      for (const message of thread.getMessages()) {
        if (Date.now()-startedMs > CWC.maxRuntimeMs) { incomplete=true; break; }
        const date=message.getDate(), id=message.getId();
        if (seen[id] || date < from || date > to) continue;
        seen[id]=true;
        // Gmail search returns threads: independently check each message's sender,
        // recipients, subject and body, since unrelated thread messages can match.
        const body=message.getPlainBody() || '';
        const subject=message.getSubject() || '';
        const identity=[message.getFrom(),message.getTo(),message.getCc(),subject,body].join(' ');
        const isK=contractor.name.indexOf('Kawartha')===0;
        if (isK ? !/kawartha home and hearth|khhltd@gmail\.com/i.test(identity) : !/steve\s+jensen|steve\.jensen83@gmail\.com/i.test(identity)) continue;
        const attachments=message.getAttachments({includeInlineImages:false,includeAttachments:true});
        const names=attachments.map(a=>a.getName() || '(unnamed)');
        const photos=attachments.filter(a=>/^image\//i.test(a.getContentType()) || /\.(jpe?g|png|heic|webp)$/i.test(a.getName()||'')).length;
        const pdfs=attachments.filter(a=>/\.pdf$/i.test(a.getName()||'') || a.getContentType()==='application/pdf').length;
        const other=attachments.length-photos-pdfs;
        const full=[subject,body,names.join(' ')].join('\n');
        const orderNumbers=cwcUnique_((full.match(/\b(?:SO\s*#?\s*|sales\s*order\s*#?\s*)\d{4,8}\b/gi)||[]).map(x=>x.replace(/\D/g,'')));
        const addresses=cwcAddresses_(subject+'\n'+names.join('\n'));
        const matched=events.filter(ev=>cwcMatches_(ev,orderNumbers,addresses));
        const candidates=matched.slice(0,10).map(ev=>ev.getTitle()+' | '+(ev.getLocation()||''));
        const dates=matched.slice(0,10).map(ev=>cwcEventTime_(ev));
        const receipt=/payment receipt|statement from/i.test(subject);
        const notes=[];
        if (receipt) notes.push('Receipt/statement; normally exclude from job entry.');
        if (pdfs) notes.push('PDF contents were not extracted; inspect invoice and exact work wording.');
        if (addresses.length>1) notes.push('Multiple address hints: may contain multiple jobs.');
        if (matched.length>10) notes.push('More than 10 calendar candidates; inspect calendar.');
        notes.push('Striven task lookup, job eligibility, completion, attachments and notifications NOT verified.');
        cwcAppend_(sheet,[contractor.name,cwcFormat_(date),subject,message.getFrom(),'https://mail.google.com/mail/u/0/#all/'+id,photos,pdfs,other,names.join('\n'),orderNumbers.join(', '),addresses.join('\n'),candidates.join('\n'),dates.join('\n'),receipt?'EXCLUDE / REVIEW':'REVIEW REQUIRED',notes.join(' ')]);
        count++;
      }
      if (incomplete) break;
    }
    offset+=threads.length;
    if (incomplete || threads.length<CWC.pageSize) break;
  }
  if (offset>=CWC.maxThreadsPerContractor) incomplete=true;
  if (incomplete) cwcAppend_(sheet,[contractor.name,'','','','','','','','','','','','','INCOMPLETE','Search stopped at runtime or thread limit; some emails may be missing.']);
  if (!count && !incomplete) cwcAppend_(sheet,[contractor.name,'','','','','','','','','','','','','NO MATCHING EMAILS','Search completed for this configured query; other aliases may not be covered.']);
  return {count:count,status:incomplete?'INCOMPLETE':'SEARCH COMPLETE'};
}
function cwcMatches_(event, orders, addresses) {
  const text=[event.getTitle(),event.getDescription()||'',event.getLocation()||''].join(' ');
  const evOrders=cwcUnique_((text.match(/\b(?:SO\s*#?\s*|sales\s*order\s*#?\s*)\d{4,8}\b/gi)||[]).map(x=>x.replace(/\D/g,'')));
  if (orders.some(n=>evOrders.indexOf(n)>=0)) return true;
  const ev=cwcNormalize_(event.getLocation()||'');
  return !!ev && addresses.some(a=>ev.indexOf(cwcNormalize_(a))>=0);
}
function cwcAddresses_(text) {
  const re=/\b\d{1,5}\s+[A-Za-z0-9][A-Za-z0-9 .'-]{1,60}?\s+(?:road|rd|street|st|drive|dr|lane|ln|avenue|ave|grove|court|ct|crescent|cres|boulevard|blvd|line|trail|trl)\b/gi;
  return cwcUnique_((text.match(re)||[]).map(s=>s.replace(/\s+/g,' ').trim()));
}
function cwcNormalize_(s) {
  return String(s).toLowerCase().replace(/\broad\b/g,'rd').replace(/\bstreet\b/g,'st').replace(/\bdrive\b/g,'dr').replace(/\blane\b/g,'ln').replace(/\bavenue\b/g,'ave').replace(/\bcourt\b/g,'ct').replace(/\bcrescent\b/g,'cres').replace(/\bboulevard\b/g,'blvd').replace(/\btrail\b/g,'trl').replace(/[^a-z0-9]+/g,' ').trim();
}
function cwcEventTime_(event) {
  if (event.isAllDayEvent()) return 'All-day: '+Utilities.formatDate(event.getAllDayStartDate(),CWC.timezone,'yyyy-MM-dd')+' to '+Utilities.formatDate(event.getAllDayEndDate(),CWC.timezone,'yyyy-MM-dd')+' (end exclusive)';
  return cwcFormat_(event.getStartTime())+' to '+cwcFormat_(event.getEndTime());
}
function cwcFormat_(date) { return Utilities.formatDate(date,CWC.timezone,'yyyy-MM-dd h:mm a'); }
function cwcUnique_(items) { return [...new Set(items.filter(Boolean))]; }
function cwcAppend_(sheet,row) { sheet.appendRow(row.map(v=>typeof v==='string' && /^[=+\-@]/.test(v)?"'"+v:v)); }

/**
 * STRIVEN CANDIDATE AUDIT — READ ONLY PROTOTYPE
 * Run contractorWeeklyCheck first, open its new tracker tab, then run
 * contractorAuditStrivenCandidates. Set Script Property CWC_STRIVEN_TOKEN to
 * a current Striven OAuth bearer token; no token or customer details are logged.
 * This is NOT an end-to-end automation. It does not create or modify tasks,
 * attach files, mark jobs done, send notifications, or install triggers.
 * Title matching finds candidates, not proof of a unique existing job.
 */
function contractorAuditStrivenCandidates() {
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('Another contractor check is running.');
  try {
    const sh=SpreadsheetApp.getActiveSheet();
    if (!sh || sh.getRange(1,1).getValue() !== 'Contractor' || sh.getRange(1,10).getValue() !== 'SO numbers found') throw new Error('Open a tracker produced by contractorWeeklyCheck.');
    const token=PropertiesService.getScriptProperties().getProperty('CWC_STRIVEN_TOKEN');
    if (!token) throw new Error('Set script property CWC_STRIVEN_TOKEN to a current Striven access token.');
    const n=sh.getLastRow()-1;
    if (n<=0) return;
    const headers=['Striven candidate task IDs','Striven candidate titles','Striven lookup status'];
    const width=sh.getLastColumn();
    const existing=sh.getRange(1,1,1,width).getValues()[0];
    let start=existing.indexOf(headers[0])+1;
    if (!start) { start=width+1; sh.getRange(1,start,1,headers.length).setValues([headers]); }
    else if (existing.slice(start-1,start+2).join('|')!==headers.join('|')) throw new Error('Candidate columns differ from expected layout.');
    const rows=sh.getRange(2,1,n,15).getValues();
    const deadline=Date.now()+240000, cache=Object.create(null);
    for (let i=0;i<rows.length;i++) {
      if (Date.now()>deadline) { sh.getRange(i+2,start+2).setValue('INCOMPLETE: run again'); break; }
      const raw=String(rows[i][9]||'').trim();
      const nums=[...new Set(raw.split(/[,\s]+/).filter(x=>/^\d{4,8}$/.test(x)))];
      let status='', candidates=[];
      if (!nums.length) status='REVIEW: no SO number; title-only lookup cannot run';
      else if (nums.length>1) status='REVIEW: multiple SO numbers; split jobs before matching';
      else try {
        const num=nums[0];
        if (!cache[num]) cache[num]=cwcSearchStrivenByTitle_(num,token);
        candidates=cache[num];
        status=candidates.length ? 'CANDIDATES ONLY: verify customer, location, task type and job' : 'NO TITLE CANDIDATE: not proof no task exists';
      } catch(e) { status='ERROR: '+String(e && e.message || e); }
      const safe=s=>String(s==null?'':s).replace(/^[=+\-@]/, "'$&");
      sh.getRange(i+2,start,1,3).setValues([[safe(candidates.map(x=>x.id).join(', ')),safe(candidates.map(x=>x.title).join('\n')),safe(status)]]);
    }
    Logger.log('Striven candidate audit finished. Candidate results require review; no Striven changes were made.');
  } finally {lock.releaseLock();}
}
function cwcSearchStrivenByTitle_(number,token) {
  let page=0, all=[];
  for (;page<10;page++) {
    const response=UrlFetchApp.fetch('https://api.striven.com/v2/tasks/search',{
      method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+token},
      payload:JSON.stringify({Title:number,PageIndex:page,PageSize:100}),muteHttpExceptions:true
    });
    const code=response.getResponseCode();
    if (code!==200) throw new Error('Striven task search HTTP '+code+'; check token, permissions and rate limits.');
    let result; try {result=JSON.parse(response.getContentText());} catch(e) {throw new Error('Invalid Striven response.');}
    const data=result.Data || result.data;
    if (!Array.isArray(data)) throw new Error('Unexpected Striven search response; no conclusion made.');
    for (const t of data) {
      const id=t.Id===undefined?t.id:t.Id, title=t.Title===undefined?t.title:t.Title;
      if (id && title && String(title).indexOf(number)>=0) all.push({id:String(id),title:String(title)});
    }
    const total=result.TotalCount===undefined?result.totalCount:result.TotalCount;
    if (data.length<100 || (total != null && Number.isFinite(Number(total)) && (page+1)*100>=Number(total))) return all;
  }
  throw new Error('Search exceeded 1000 results; narrow the search.');
}
