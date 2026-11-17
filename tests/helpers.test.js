// Run: node --test tests/   (pure helpers only; Gmail/Drive/Sheets calls need Apps Script)
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const fs = require('node:fs');
const vm = require('node:vm');

// Code.gs is plain JS; load it as a CommonJS module without renaming the file.
const src = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
const mod = { exports: {} };
vm.runInNewContext(src, { module: mod, console });
const h = mod.exports;

test('extractPoNumber finds common PO formats', () => {
  assert.equal(h.extractPoNumber('Order confirmation for PO# 45-1180'), '45-1180');
  assert.equal(h.extractPoNumber('P.O. Number: a12345 attached'), 'A12345');
  assert.equal(h.extractPoNumber('Re: purchase order 778812 - revised'), '778812');
  assert.equal(h.extractPoNumber('PO 2026/0042/'), '2026/0042');
  assert.equal(h.extractPoNumber('Weekly newsletter'), '');
  assert.equal(h.extractPoNumber('Please report to the post office'), '');
  assert.equal(h.extractPoNumber(undefined), '');
});

test('senderName handles named, quoted and bare addresses', () => {
  assert.equal(h.senderName('Jane Doe <jane@acme.com>'), 'Jane Doe');
  assert.equal(h.senderName('"Acme Sales" <sales@acme.com>'), 'Acme Sales');
  assert.equal(h.senderName('orders@vendor.co'), 'orders');
  assert.equal(h.senderName('<orders@vendor.co>'), 'orders');
});

test('sanitizeFilename removes characters Drive and Windows reject', () => {
  assert.equal(h.sanitizeFilename('Q3: price/list*final?.xlsx'), 'Q3_ price_list_final_.xlsx');
  assert.equal(h.sanitizeFilename('   '), 'attachment');
  assert.equal(h.sanitizeFilename('a'.repeat(300)).length, 120);
});

test('isAllowed checks the extension case-insensitively', () => {
  const allowed = ['pdf', 'xlsx'];
  assert.equal(h.isAllowed('Quote.PDF', allowed), true);
  assert.equal(h.isAllowed('logo.png', allowed), false);
  assert.equal(h.isAllowed('no-extension', allowed), false);
});

test('buildFileName prefixes date and sender', () => {
  assert.equal(h.buildFileName('2026-11-18', 'Jane Doe <j@x.com>', 'Quote.pdf'),
               '2026-11-18 Jane Doe - Quote.pdf');
});

test('buildDigest groups by sender and lists PO numbers', () => {
  const d = h.buildDigest([
    { from: 'Acme <a@acme.com>', subject: 's', po: '1001', file: 'a.pdf', link: 'L1' },
    { from: 'Acme <a@acme.com>', subject: 's', po: '1002', file: 'b.pdf', link: 'L2' },
    { from: 'Bolt <b@bolt.com>', subject: 's', po: '', file: 'c.xlsx', link: 'L3' },
  ], '2026-11-18');
  assert.equal(d.subject, 'Attachments 2026-11-18: 3 file(s), 2 sender(s)');
  assert.match(d.text, /PO numbers: 1001, 1002/);
  assert.ok(d.text.indexOf('Acme (2)') < d.text.indexOf('Bolt (1)'), 'busiest sender first');
  assert.match(d.text, /a\.pdf \[PO 1001\]/);
});

test('buildDigest says so when nothing arrived', () => {
  assert.equal(h.buildDigest([], '2026-11-18').subject, 'Attachments 2026-11-18: none');
});
