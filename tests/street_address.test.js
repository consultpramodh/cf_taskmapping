// New Striven Locations need a street address (2026-11-23). Run: node tests/street_address.test.js
'use strict';
const fs = require('fs'), path = require('path'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'apps-script', '10_Shared_Core.gs'), 'utf8');
const body = src.slice(src.indexOf('function tmv3_isStreetAddress_'));
const fn = new Function('tmv3_clean_', body.slice(0, body.indexOf('\n}\n') + 3) + '; return tmv3_isStreetAddress_;')(
  v => String(v == null ? '' : v).replace(/\s+/g, ' ').trim());
['4722 Old Simcoe St, Oshawa, ON L1H 0M6', '1360 Glenanna Rd', '1360 Glenanna Rd, Unit 9A, Pickering',
 '10 Broadbridge Dr', '300 Kingswood road', '3 McAlpine St, Unit 508 Toronto', 'Unit 5 - 22 King St W', '12B Main St']
  .forEach(a => assert.strictEqual(fn(a), true, a));
['Highway 35 S / Highway 115 South', 'Hwy 7 & Brock Rd', 'Old Forest Hill Rd', 'Costco Ajax', '', 'corner of King and Bay']
  .forEach(a => assert.strictEqual(fn(a), false, a));
console.log('street_address: 14 checks passed');
