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
test('fixed fields cannot be removed and customer/company headers map distinctly', () => {
  for (const key of ['company','contact','email']) assert.throws(() => C.removeField(C.fields,key), /不能删除/);
  const mapping = C.autoMap(['客户名称','公司名','邮箱']);
  assert.equal(mapping.contact,'0'); assert.equal(mapping.company,'1'); assert.equal(mapping.email,'2');
});
test('custom schema flows through mapping, data, template, XLSX and CSV export', () => {
  let schema = C.removeField(C.fields,'phone');
  schema = C.addField(schema,'采购数量','custom_quantity');
  const rows = [['公司名','客户名称','邮箱','电话','采购数量'],['Acme','Emma','emma@example.com','00123','00800']];
  const records = C.mappedRows(rows,0,C.autoMap(rows[0],schema),'test',schema);
  assert.equal(records[0].custom_quantity,'00800'); assert.equal(Object.hasOwn(records[0],'phone'),false);
  records[0].custom_quantity = '00900';
  const matrix = C.exportMatrix(records,schema), column = matrix[0].indexOf('采购数量');
  assert.equal(matrix[0].includes('电话'),false); assert.equal(matrix[1][column],'00900');
  assert.equal(C.exportMatrix([],schema,true)[0][column],'采购数量');
  const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(matrix),'Clients');
  const exported = XLSX.read(XLSX.write(book,{type:'buffer',bookType:'xlsx'}),{type:'buffer'});
  const back = XLSX.utils.sheet_to_json(exported.Sheets.Clients,{header:1,raw:false});
  assert.equal(back[1][column],'00900');
  assert.ok(matrix.map(row => row.map(C.csvCell).join(',')).join('\n').includes('"00900"'));
  schema = C.removeField(schema,'custom_quantity');
  assert.equal(C.exportMatrix(records,schema)[0].includes('采购数量'),false);
});
test('schema changes preserve manual mappings and free removed columns', () => {
  let schema = C.removeField(C.fields,'phone');
  schema = C.addField(schema,'WhatsApp','custom_whatsapp');
  const mapping = C.autoMap(['Org','Person','E-mail','WhatsApp'],schema,{company:'0',contact:'',email:'2',phone:'3'});
  assert.equal(mapping.company,'0'); assert.equal(mapping.contact,''); assert.equal(mapping.custom_whatsapp,'3'); assert.equal(Object.hasOwn(mapping,'phone'),false);
});
test('field names validate empty, duplicate, reserved and excessive input', () => {
  for (const name of [' ','邮箱',' 邮 箱 ', '来源文件 / 工作表']) {
    assert.throws(() => C.addField(C.fields,name,'custom_new'));
  }
  assert.throws(() => C.addField(C.fields,'a'.repeat(31),'custom_new'));
  assert.throws(() => C.addField(C.fields,'新字段','__proto__'));
  assert.throws(() => C.addField(Array.from({length:40},(_,i) => ({key:'k'+i,label:'字段'+i,aliases:[]})),'新字段','custom_new'));
});
test('deleted phone no longer satisfies contact validation and restoring it retains aliases', () => {
  const schema = C.removeField(C.fields,'phone'), record = {company:'Acme',phone:'00123'};
  assert.deepEqual(C.issues(record,schema),['缺少联系方式']);
  const restored = C.addField(schema,'电话','custom_restore');
  assert.deepEqual(C.issues(record,restored),[]); assert.equal(C.autoMap(['WhatsApp'],restored).phone,'0');
});
