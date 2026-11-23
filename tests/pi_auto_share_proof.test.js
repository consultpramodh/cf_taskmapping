// Customer proof for auto-sharing rep-booked pre-inspections (2026-11-23). Run: node tests/pi_auto_share_proof.test.js
'use strict';
const fs = require('fs'), path = require('path'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'apps-script', '30_PreInspection.gs'), 'utf8');
const body = src.slice(src.indexOf('function tmv3_piAutoShareCustomerProof_'));
const digits = v => String(v || '').replace(/\D/g, '');
const fn = new Function('tmv3_phone10_', 'tmv3_clean_', 'tmv3_norm_', 'tmv3_extractPhone_',
  body.slice(0, body.indexOf('\n}\n') + 3) + '; return tmv3_piAutoShareCustomerProof_;')(
  v => { const d = digits(v); return d.length >= 10 ? d.slice(-10) : ''; },
  v => String(v == null ? '' : v).trim(), v => String(v || '').toLowerCase(),
  t => (String(t).match(/\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}/) || [''])[0]);
const ahmed = {'Customer Number': '63351', Name: 'Ahmed Kayka', 'Primary Phone': '(647) 272-8511'};
const cmh = {'Customer Number': '32051', Name: 'CMH INTERIORS', 'Primary Phone': '(905) 243-4165'};
const other = {'Customer Number': '11111', Name: 'Jane Doe', 'Primary Phone': '(416) 000-0000'};
const index = {byNumber: {'63351': ahmed, '32051': cmh, '11111': other},
  byPhone: {'6472728511': [ahmed], '9052434165': [cmh], '4160000000': [other, {Name: 'dup'}]}};
// Customer # + matching phone
assert.strictEqual(fn({title: '63351 - Ahmed Kayka (647) 272-8511', customerNumber: '63351', phone: '6472728511'}, index), ahmed);
// Carrie: Customer # + matching phone even though the title name differs from the Customer name
assert.strictEqual(fn({title: 'Carrie - CUst #32051 (905) 243-4165', customerNumber: '32051', phone: '9052434165'}, index), cmh);
// Customer # + surname, no phone
assert.strictEqual(fn({title: '63351 - Kayka shop stove', customerNumber: '63351'}, index), ahmed);
// Customer # whose phone and name do not match: not proven
assert.strictEqual(fn({title: '63351 - Bob Smith (905) 111-2222', customerNumber: '63351', phone: '9051112222'}, index), null);
// Unknown Customer #
assert.strictEqual(fn({title: '99999 - X', customerNumber: '99999', phone: '6472728511'}, index), null);
// No Customer #: phone matching exactly one Customer
assert.strictEqual(fn({title: 'Ahmed (647) 272-8511 pre-inspect', phone: '6472728511'}, index), ahmed);
// Suresh: phone not in Striven
assert.strictEqual(fn({title: 'Suresh Singh (416)272-2642', phone: '4162722642'}, index), null);
// Phone shared by two Customers: not proven
assert.strictEqual(fn({title: 'Someone (416) 000-0000', phone: '4160000000'}, index), null);
console.log('pi_auto_share_proof: 8 checks passed');
