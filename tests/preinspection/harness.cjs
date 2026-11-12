'use strict';
// Loads the Apps Script project into one Node VM context with in-memory
// stand-ins for Google services. No network, Calendar, Striven or mail access.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const ORDER = ['00_Configuration.gs','10_Shared_Core.gs','20_Master_Triggers.gs',
  '30_PreInspection.gs','40_Installation.gs','50_Delivery.gs','60_Service.gs','90_Regression_Tests.gs'];

function formatDate(date, tz, pattern) {
  const d = date instanceof Date ? date : new Date(date);
  if (!isFinite(d.getTime())) throw new Error('Invalid date');
  const parts = {};
  new Intl.DateTimeFormat('en-US', {timeZone: tz, year: 'numeric', month: 'short', day: 'numeric',
    hour: 'numeric', minute: '2-digit', second: '2-digit', hour12: false, weekday: 'short'})
    .formatToParts(d).forEach(p => { parts[p.type] = p.value; });
  const num = new Intl.DateTimeFormat('en-US', {timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit'})
    .formatToParts(d).reduce((o, p) => (o[p.type] = p.value, o), {});
  let H = Number(parts.hour) % 24;
  const offsetMin = (() => {
    const local = Date.UTC(Number(num.year), Number(num.month) - 1, Number(num.day), H, Number(parts.minute), Number(parts.second));
    return Math.round((local - Math.floor(d.getTime() / 1000) * 1000) / 60000);
  })();
  const sign = offsetMin >= 0 ? '+' : '-';
  const abs = Math.abs(offsetMin);
  const xxx = sign + String(Math.floor(abs / 60)).padStart(2, '0') + ':' + String(abs % 60).padStart(2, '0');
  const h12 = H % 12 === 0 ? 12 : H % 12;
  const tokens = {
    yyyy: num.year, MMM: parts.month, MM: num.month, dd: num.day, d: String(Number(num.day)),
    HH: String(H).padStart(2, '0'), H: String(H), hh: String(h12).padStart(2, '0'), h: String(h12),
    mm: parts.minute, ss: parts.second, a: H < 12 ? 'AM' : 'PM', XXX: xxx, EEE: parts.weekday
  };
  let out = '';
  for (let i = 0; i < pattern.length;) {
    if (pattern[i] === "'") { const j = pattern.indexOf("'", i + 1); out += pattern.slice(i + 1, j); i = j + 1; continue; }
    const m = /^(yyyy|MMM|MM|dd|d|HH|H|hh|h|mm|ss|a|XXX|EEE)/.exec(pattern.slice(i));
    if (m) { out += tokens[m[1]]; i += m[1].length; } else { out += pattern[i]; i++; }
  }
  return out;
}

function memoryProperties(initial) {
  const store = Object.assign({}, initial || {});
  return {
    store,
    getProperty: k => Object.prototype.hasOwnProperty.call(store, k) ? store[k] : null,
    setProperty(k, v) { store[k] = String(v); return this; },
    deleteProperty(k) { delete store[k]; return this; },
    getProperties: () => Object.assign({}, store)
  };
}

function createServices(log) {
  const props = memoryProperties();
  const record = (kind, detail) => log.push(Object.assign({kind}, detail || {}));
  return {
    props,
    PropertiesService: {getScriptProperties: () => props},
    LockService: {getScriptLock: () => ({tryLock: () => true, waitLock() {}, releaseLock() {}, hasLock: () => true})},
    Logger: {log: () => {}},
    Utilities: {
      formatDate,
      computeDigest: (alg, s) => Array.from(require('node:crypto').createHash('sha256').update(String(s)).digest()).map(b => b > 127 ? b - 256 : b),
      DigestAlgorithm: {SHA_256: 'SHA_256', MD5: 'MD5'},
      Charset: {UTF_8: 'UTF_8'},
      sleep: () => {},
      getUuid: () => require('node:crypto').randomUUID(),
      computeHmacSha256Signature: (v, k) => Array.from(require('node:crypto').createHmac('sha256', String(k)).update(String(v)).digest()).map(b => b > 127 ? b - 256 : b),
      base64EncodeWebSafe: s => Buffer.from(Array.isArray(s) ? s.map(b => (b + 256) % 256) : String(s)).toString('base64').replace(/\+/g, '-').replace(/\//g, '_')
    },
    MailApp: {
      getRemainingDailyQuota: () => 100,
      sendEmail(o) { record('MAIL', {to: o.to, cc: o.cc, subject: o.subject}); }
    },
    CalendarApp: {getCalendarById(id) { record('CALENDAR_READ', {id}); return null; }},
    ScriptApp: {
      getProjectTriggers: () => [], deleteTrigger() { record('TRIGGER_DELETE'); },
      newTrigger() { record('TRIGGER_CREATE'); throw new Error('TRIGGER_CREATE_BLOCKED_IN_HARNESS'); },
      getScriptId: () => 'harness', requireScopes() {}, AuthMode: {FULL: 'FULL'}
    },
    SpreadsheetApp: {openById() { record('SHEET_OPEN'); throw new Error('SHEETS_BLOCKED_IN_HARNESS'); }, flush() {}},
    UrlFetchApp: {fetch() { record('HTTP'); throw new Error('NETWORK_BLOCKED_IN_HARNESS'); }},
    Session: {getActiveUser: () => ({getEmail: () => (globalThis.__HARNESS_USER__ || 'harness@example.invalid')}), getScriptTimeZone: () => 'America/Toronto'},
    CacheService: {getScriptCache: () => ({get: () => null, put() {}, remove() {}})},
    HtmlService: {createHtmlOutput: html => ({html: String(html), setTitle() { return this; }})}, ContentService: {}
  };
}

// dir: folder containing the .gs files. overrides: {fileName: sourceText}.
function load(dir, options) {
  options = options || {};
  const log = [];
  const services = createServices(log);
  const ctx = Object.assign({console}, services);
  delete ctx.props;
  vm.createContext(ctx);
  if (options.fixedNow !== undefined) {
    // Deterministic clock inside the VM: new Date() and Date.now() return fixedNow.
    ctx.__FIXED_NOW__ = options.fixedNow;
    vm.runInContext('(function(){const R=Date;const F=__FIXED_NOW__;' +
      'class D extends R{constructor(...a){if(a.length===0)super(F);else super(...a);}static now(){return F;}}' +
      'globalThis.Date=D;})();', ctx);
  }
  const files = ORDER.filter(f => (options.overrides && options.overrides[f]) || fs.existsSync(path.join(dir, f)));
  const source = files.map(f => {
    let text = (options.overrides && options.overrides[f]) || fs.readFileSync(path.join(dir, f), 'utf8');
    if (options.transform) text = options.transform(f, text);
    return text + '\n';
  }).join('\n');
  vm.runInContext(source, ctx, {filename: 'apps-script-bundle.gs'});
  return {ctx, log, props: services.props, files};
}

module.exports = {load, formatDate, memoryProperties, ORDER};
