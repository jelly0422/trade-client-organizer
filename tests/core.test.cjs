const {test} = require('node:test');
const assert = require('node:assert/strict');
const C = require('../dist/core.js');
const XLSX = require('../dist/vendor/xlsx.full.min.js');
test('recognizes bilingual aliases and title rows', () => {
  const rows = [['2026 展会名单'],['公司名称','Contact','E-mail','WhatsApp','Country','Product Interest'],['Acme','A','a@example.com','00123','UK','Cup']];
  assert.equal(C.detectHeader(rows),1);
  const mapped = C.mappedRows(rows,1,C.autoMap(rows[1]),'test');
  assert.equal(mapped[0].phone,'00123'); assert.equal(mapped[0].product,'Cup'); assert.equal(mapped[0].sourceRow,3);
});
test('manual mapping and blank rows preserve row provenance', () => {
  const rows = [['客户','需求'],[],['公司甲','纸杯'],['','']];
  const result = C.mappedRows(rows,0,{company:'0',product:'1'},'sheet');
  assert.equal(result.length,1); assert.equal(result[0].sourceRow,3); assert.equal(result[0].email,'');
});
test('validation flags missing identity, contact methods and malformed emails', () => {
  assert.equal(C.issues({}).length,2); assert.deepEqual(C.issues({company:'A',email:'a@b.com; b@c.com'}),[]);
  assert.deepEqual(C.issues({company:'A',email:'invalid'}),['邮箱格式待核对']);
});
test('duplicate detection includes both originals without matching blank identities', () => {
  assert.deepEqual([...C.duplicateIds([{id:'1',email:'A@B.com'},{id:'2',email:'a@b.com'},{id:'3'},{id:'4'}])],['1','2']);
});
test('CSV preserves quoting and mitigates formula prefixes', () => {
  assert.equal(C.csvCell('=HYPERLINK("x")'),'"\'=HYPERLINK(""x"")"'); assert.equal(C.csvCell('Acme, Inc.'),'"Acme, Inc."');
});
for (const type of ['xlsx','xls','ods','csv']) test(type + ' parsing preserves international phone strings and multilingual cells', () => {
  const data = [['Company','Phone','Product'],['公司甲','+44 001234','竹制餐具'],['Beta','001234','Cups']];
  const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(data),'Clients');
  const bytes = XLSX.write(book,{type:'buffer',bookType:type});
  const read = type === 'csv' ? XLSX.read(new TextDecoder().decode(bytes),{type:'string',raw:true}) : XLSX.read(bytes,{type:'buffer'});
  const rows = XLSX.utils.sheet_to_json(read.Sheets[read.SheetNames[0]],{header:1,raw:false});
  const records = C.mappedRows(rows,0,C.autoMap(rows[0]),type);
  assert.equal(records[0].company,'公司甲'); assert.equal(records[0].phone,'+44 001234'); assert.equal(records[1].phone,'001234');
});
test('XLSX export keeps formula-like customer text as strings', () => {
  const sheet = XLSX.utils.aoa_to_sheet([['=1+1','+12345','00123']]);
  assert.equal(sheet.A1.t,'s'); assert.equal(sheet.A1.f,undefined); assert.equal(sheet.C1.v,'00123');
});
