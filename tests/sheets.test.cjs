const { test } = require('node:test');
const assert = require('node:assert/strict');
const XLSX = require('../public/vendor/xlsx.full.min.js');
const C = require('../src/domain/core.ts');
const S = require('../src/domain/sheets.ts');
const W = require('../src/state/workspace.ts');
const L = require('../src/services/localWorkspace.ts');
const B = require('../src/services/workbooks.ts');
const { IDBFactory } = require('fake-indexeddb');
global.window = { XLSX };

function fixture(id = 'file-one') {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    ['Company', 'Contact', 'Email', 'Phone', 'Other'],
    ['Acme', 'Alex', 'alex@example.com', '001234', '=unmapped'],
    [],
    ['Pending Co', '', 'pending@example.com', '', '补充联系人'],
  ]), '欧洲客户');
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    ['亚洲展会客户'],
    ['邮箱', '公司名', '采购经理', '产品需求', '采购数量'],
    ['alex@example.com', '另一家公司', 'Lin', '杯子', '0008'],
  ], { origin: 'C3' }), '亚洲客户');
  book.Sheets['亚洲客户']['!ref'] = 'C3:G5';
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([]), '空表');
  let source = B.withConfig({ id, name: '客户.xlsx', size: 0, book, sheet: '欧洲客户', configs: {}, encoding: 'utf-8' }, C.fields);
  source = B.withConfig({ ...source, sheet: '亚洲客户' }, C.fields);
  source.configs['亚洲客户'].mapping.contact = '2';
  return { ...source, sheet: '欧洲客户' };
}
const importSeparate = (id) => B.importWorkbook({ ...fixture(id), mergeSheets: false }, C.fields);
const append = batch => W.reducer(W.initialWorkspace(), { type: 'append', ...batch });
const rows = (book, name) => XLSX.utils.sheet_to_json(book.Sheets[name], { header: 1, defval: '' });
const roundTrip = book => XLSX.read(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer' });

test('multi-sheet import requires an explicit merge choice and imports all worksheets together', () => {
  assert.throws(() => B.importWorkbook(fixture(), C.fields), /选择是否合并/);
  const batch = importSeparate();
  assert.equal(batch.records.length, 3);
  assert.deepEqual(batch.sheets.map(s => s.name), ['欧洲客户', '亚洲客户', '空表']);
  assert.equal(batch.records[0].phone, '001234');
  assert.equal(batch.records[1].sourceRow, 4);
  assert.equal(batch.records[2].contact, 'Lin');
  assert.equal(batch.records[2].sourceRow, 5);
  assert.equal(batch.records[2].sourceCells[0].column, 3);
  assert.equal(batch.records[0].sourceCells.at(-1).value, '=unmapped');
  assert.equal(batch.records[1].pendingIdentity, true);
  assert.equal(batch.records[0].sheetId, batch.sheets[0].id);
  assert.equal(batch.records[2].sheetId, batch.sheets[1].id);
  assert.equal(new Set(batch.records.map(r => r.id)).size, 3);
});

test('Excel round-trip retains separate sheets, pending rows, empty sheets and global duplicate warnings', () => {
  const state = append(importSeparate());
  const book = roundTrip(B.buildCustomerWorkbook(state.records, state.fields, false, state.sheets));
  assert.deepEqual(book.SheetNames, ['欧洲客户', '亚洲客户', '空表', '导入原始内容']);
  const europe = rows(book, '欧洲客户'), asia = rows(book, '亚洲客户');
  assert.equal(europe.length, 3);
  assert.equal(asia.length, 2);
  assert.equal(rows(book, '空表').length, 1);
  assert.equal(europe[2][europe[0].indexOf('记录分类')], '缺少主体');
  assert.match(europe[1][europe[0].indexOf('校对提示')], /疑似重复/);
  assert.match(asia[1][asia[0].indexOf('校对提示')], /疑似重复/);
  assert.equal(europe[1][europe[0].indexOf('电话')], '001234');
  assert.ok(rows(book, '导入原始内容').some(row => row.includes('=unmapped')));
  for (const sheet of Object.values(book.Sheets))
    for (const [address, cell] of Object.entries(sheet)) if (!address.startsWith('!')) assert.equal(cell.f, undefined);
});

test('merging combines customers while retaining original sheet provenance', () => {
  const batch = B.importWorkbook({ ...fixture(), mergeSheets: true }, C.fields);
  assert.deepEqual(batch.sheets, []);
  assert.ok(batch.records.every(r => r.sheetId === undefined));
  assert.equal(batch.records.length, 3);
  assert.match(batch.records[2].source, /亚洲客户/);
  const book = B.buildCustomerWorkbook(batch.records, C.fields, false, batch.sheets);
  assert.deepEqual(book.SheetNames, ['客户资料', '缺少主体', '导入原始内容']);
  assert.equal(rows(book, '客户资料').length, 3);
  assert.equal(rows(book, '缺少主体').length, 2);
});

test('independent sheet mappings survive integrated schema additions and removals', () => {
  const fields = C.addField(C.fields, '采购数量', 'custom_qty');
  let [source] = B.synchronizeSchema([fixture()], fields);
  assert.equal(source.configs['亚洲客户'].mapping.contact, '2');
  assert.equal(source.configs['亚洲客户'].mapping.custom_qty, '4');
  const batch = B.importWorkbook({ ...source, mergeSheets: false }, fields);
  assert.equal(batch.records[2].custom_qty, '0008');
  assert.equal(batch.records[0].custom_qty, '');
  [source] = B.synchronizeSchema([source], C.removeField(fields, 'phone'));
  assert.equal(source.configs['欧洲客户'].mapping.phone, undefined);
});

test('sheet ownership survives edits, admission, delete undo and local persistence', async () => {
  let state = append(importSeparate());
  const pending = state.records[1], sheetId = pending.sheetId;
  state = W.reducer(state, { type: 'edit', id: pending.id, key: 'contact', value: 'Filled' });
  state = W.reducer(state, { type: 'admit', ids: [pending.id] });
  state = W.reducer(state, { type: 'delete', ids: [pending.id] });
  const factory = new IDBFactory();
  await L.writeLocalWorkspace(state, null, factory);
  state = (await L.readLocalWorkspace(factory)).workspace;
  assert.equal(state.sheets.length, 3);
  assert.equal(state.deleted[0].record.sheetId, sheetId);
  state = W.reducer(state, { type: 'undoDelete' });
  assert.equal(state.records[1].sheetId, sheetId);
  assert.equal(state.records[1].contact, 'Filled');
  assert.equal(C.pendingIdentity(state.records[1]), false);
  assert.equal(rows(B.buildCustomerWorkbook(state.records, state.fields, false, state.sheets), '欧洲客户').length, 3);
});

test('old saved data without a sheet registry migrates without losing original content', () => {
  const legacy = W.initialWorkspace();
  delete legacy.sheets;
  legacy.records = [{ id: 'old', company: 'Old', contact: 'Alex', email: 'old@example.com', sourceCells: [{ column: 1, label: 'X', value: 'kept' }] }];
  const decoded = L.decodeWorkspace({ version: 1, revision: 'r', savedAt: 1, workspace: legacy });
  assert.deepEqual(decoded.workspace.sheets, []);
  assert.deepEqual(decoded.workspace.records, legacy.records);
});

test('invalid sheet references or duplicate registry entries are rejected without overwriting saved progress', () => {
  const state = append(importSeparate());
  const envelope = workspace => ({ version: 1, revision: 'r', savedAt: 1, workspace });
  assert.throws(() => L.decodeWorkspace(envelope({ ...state, sheets: [] })), /未被覆盖/);
  assert.throws(() => L.decodeWorkspace(envelope({ ...state, sheets: [...state.sheets, state.sheets[0]] })), /未被覆盖/);
  assert.throws(() => W.reducer(W.initialWorkspace(), { type: 'append', records: state.records }), /工作表不存在/);
  assert.throws(() => B.buildCustomerWorkbook(state.records, state.fields, false, []), /工作表信息缺失/);
});

test('same-named sheets from different imports are independent and get unique export names', () => {
  let state = append(importSeparate('file-a'));
  state = W.reducer(state, { type: 'append', ...importSeparate('file-b') });
  assert.equal(state.sheets.length, 6);
  assert.equal(S.recordsInSheet(state.records, state.sheets[0].id).length, 2);
  assert.equal(S.recordsInSheet(state.records, state.sheets[3].id).length, 2);
  const book = roundTrip(B.buildCustomerWorkbook(state.records, state.fields, false, state.sheets));
  assert.deepEqual(book.SheetNames, ['欧洲客户', '亚洲客户', '空表', '欧洲客户 (2)', '亚洲客户 (2)', '空表 (2)', '导入原始内容']);
  assert.throws(() => W.reducer(state, { type: 'append', ...importSeparate('file-a') }), /重复添加/);
});

test('sheet filters keep selection, category and pagination membership isolated', () => {
  const state = append(importSeparate());
  const europe = S.recordsInSheet(state.records, state.sheets[0].id);
  const asia = S.recordsInSheet(state.records, state.sheets[1].id);
  assert.equal(S.recordsInSheet(state.records, state.sheets[2].id).length, 0);
  const selected = new Set(europe.map(r => r.id));
  assert.equal(W.visibleSelection(selected, asia).size, 0);
  assert.equal(W.filterRecords(europe, state.fields, 'missing', '', W.summarize(state.records)).length, 1);
  assert.equal(W.filterRecords(asia, state.fields, 'duplicate', '', W.summarize(state.records)).length, 1);
});

test('empty-only imports and a sheet emptied by deletion remain exportable', () => {
  const source = fixture();
  source.book.SheetNames = ['空表'];
  source.sheet = '空表';
  const batch = B.importWorkbook(source, C.fields);
  assert.equal(batch.records.length, 0);
  assert.deepEqual(B.buildCustomerWorkbook([], C.fields, false, batch.sheets).SheetNames, ['空表']);
  let state = append(importSeparate());
  state = W.reducer(state, { type: 'delete', ids: state.records.map(r => r.id) });
  assert.deepEqual(B.buildCustomerWorkbook(state.records, state.fields, false, state.sheets).SheetNames, ['欧洲客户', '亚洲客户', '空表']);
});

test('CSV cannot silently flatten multiple sheets; standard templates stay single-sheet', () => {
  const state = append(importSeparate());
  assert.throws(() => B.exportCustomers(state.records, state.fields, 'csv', false, state.sheets), /CSV 不支持多个工作表/);
  const template = B.buildCustomerWorkbook(state.records, state.fields, true, state.sheets);
  assert.deepEqual(template.SheetNames, ['客户资料']);
  assert.deepEqual(rows(template, '客户资料'), [C.fields.map(f => f.label)]);
});

test('mixed merged records and separate sheets export every record once, and enforce the total record limit', () => {
  let state = append(importSeparate());
  const manual = { id: 'manual', company: 'Manual', contact: 'Alex', email: 'm@example.com' };
  state = W.reducer(state, { type: 'append', records: [manual] });
  assert.equal(S.customerSheetGroups(state.records, state.sheets).length, 4);
  const book = B.buildCustomerWorkbook(state.records, state.fields, false, state.sheets);
  const dataCount = book.SheetNames.filter(n => n !== '导入原始内容').reduce((sum, n) => sum + rows(book, n).length - 1, 0);
  assert.equal(dataCount, state.records.length);
  assert.throws(() => W.reducer({ ...state, records: Array.from({ length: W.MAX_RECORDS }, () => manual) }, { type: 'append', ...importSeparate('extra') }), /50,000/);
});

test('Excel names handle invalid characters, long Unicode names, case collisions and metadata names', () => {
  const used = new Set();
  const names = ['Sales', 'sales', 'Sales', 'a/b:c?d*e[f]g', "'test'", 'History', '😀'.repeat(40), '😀'.repeat(40), '导入原始内容', '导入原始内容'];
  const safe = names.map(name => S.uniqueWorksheetName(name, used));
  assert.equal(new Set(safe.map(n => n.toLowerCase())).size, names.length);
  assert.ok(safe.every(n => n.length <= 31 && !/[\\\/?*\[\]:]/.test(n) && !n.startsWith("'") && !n.endsWith("'")));
  assert.deepEqual(safe.slice(0, 3), ['Sales', 'sales (2)', 'Sales (3)']);
  const book = XLSX.utils.book_new();
  safe.forEach(name => XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['kept']]), name));
  assert.deepEqual(roundTrip(book).SheetNames, safe);
});

test('XLSX, XLS and ODS input retains multiple worksheet names through file parsing', async () => {
  for (const bookType of ['xlsx', 'xls', 'ods']) {
    const fixtureBook = fixture().book;
    const bytes = XLSX.write(fixtureBook, { type: 'buffer', bookType });
    const source = await B.readFile(new File([bytes], 'test.' + bookType), C.fields);
    assert.deepEqual(source.book.SheetNames, fixtureBook.SheetNames);
    const batch = B.importWorkbook({ ...source, mergeSheets: false }, C.fields);
    assert.equal(batch.sheets.length, 3);
    assert.equal(batch.records.length, 3);
  }
});
