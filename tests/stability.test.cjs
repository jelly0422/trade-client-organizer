const { test } = require('node:test');
const assert = require('node:assert/strict');
const C = require('../src/domain/core.ts');
const W = require('../src/state/workspace.ts');
const L = require('../src/services/localWorkspace.ts');
const { IDBFactory } = require('fake-indexeddb');
const customer = (id, extra = {}) => ({ id, company: 'Acme', contact: 'Alex', email: id + '@example.com', ...extra });
const workspace = records => W.reducer(W.initialWorkspace(), { type: 'append', records });
const tick = () => new Promise(resolve => setImmediate(resolve));
async function until(predicate) { for (let i = 0; i < 100; i++) { if (predicate()) return; await tick(); } assert.fail('Async operation did not settle'); }

test('complete records retain unmapped, duplicate-label, offset and zero source cells', () => {
  const rows = [['Company', 'Contact', 'Email', 'Other', 'Other'], ['Acme', 'Alex', 'a@example.com', 'recycled packaging', 0]];
  const [record] = C.mappedRows(rows, 0, C.autoMap(rows[0]), 'source.xlsx', C.fields, 3, 2);
  assert.equal(C.pendingIdentity(record), false);
  assert.equal(record.sourceCells.length, 5);
  assert.deepEqual(record.sourceCells.at(-1), { column: 7, label: 'Other', value: '0' });
  assert.ok(C.recoveryMatrix([record]).some(row => row.includes('recycled packaging')));
  const edited = W.reducer(workspace([{ ...record, id: 'a' }]), { type: 'edit', id: 'a', key: 'product', value: 'recycled packaging' });
  assert.equal(edited.records[0].product, 'recycled packaging');
  assert.deepEqual(edited.records[0].sourceCells, record.sourceCells);
});
test('duplicates preserve local punctuation and case while ignoring domain case', () => {
  const emails = ['ab@example.com', 'a-b@example.com', 'a_b@example.com', 'a.b@example.com', 'a+b@example.com', 'a/b@example.com', 'Ab@example.com'];
  assert.equal(C.duplicateIds(emails.map((email, i) => customer(String(i), { email }))).size, 0);
  assert.deepEqual([...C.duplicateIds([customer('a', { email: 'Alex@EXAMPLE.COM' }), customer('b', { email: ' Alex@example.com ' })])], ['a', 'b']);
});
test('duplicate matching crosses multi-email order and separators, counting each record once', () => {
  const records = [customer('a', { email: 'sales@example.com;buyer@example.com' }), customer('b', { email: 'buyer@example.com' }), customer('c', { email: 'third@example.com。sales@EXAMPLE.COM' })];
  assert.deepEqual([...C.duplicateIds(records)].sort(), ['a', 'b', 'c']);
  assert.equal(C.duplicateIds([customer('a', { email: 'a@example.com;a@example.com' })]).size, 0);
});
test('invalid email parts do not form duplicate identities; valid parts still match', () => {
  const records = [customer('a', { email: 'invalid;valid@example.com' }), customer('b', { email: 'valid@example.com' }), customer('c', { email: 'invalid' }), customer('d', { email: 'invalid' })];
  assert.deepEqual([...C.duplicateIds(records)], ['a', 'b']);
  assert.equal(W.summarize([customer('p', { pendingIdentity: true }), customer('a', { email: 'p@example.com' })]).counts.duplicate, 0);
});
test('selection drops records leaving review/duplicate/search, retaining selections on other pages', () => {
  let state = workspace([customer('a', { email: 'invalid' }), customer('b')]);
  const selected = new Set(['a']);
  state = W.reducer(state, { type: 'edit', id: 'a', key: 'email', value: 'a@example.com' });
  assert.equal(W.visibleSelection(selected, W.filterRecords(state.records, state.fields, 'review', '', W.summarize(state.records))).size, 0);
  const all = Array.from({ length: 30 }, (_, i) => customer(String(i)));
  assert.deepEqual([...W.visibleSelection(new Set(['0', '29', 'deleted']), all)], ['0', '29']);
  assert.equal(W.visibleSelection(new Set(['0']), W.filterRecords(all, C.fields, 'all', 'not found', W.summarize(all))).size, 0);
  let duplicate = workspace([customer('a'), customer('b', { email: 'a@example.com' })]);
  duplicate = W.reducer(duplicate, { type: 'edit', id: 'a', key: 'email', value: 'unique@example.com' });
  assert.equal(W.visibleSelection(new Set(['a', 'b']), W.filterRecords(duplicate.records, C.fields, 'duplicate', '', W.summarize(duplicate.records))).size, 0);
});
test('undo restores order, pending confirmation, custom data and original source without reverting later edits', () => {
  let state = workspace([customer('a'), customer('b', { pendingIdentity: true, sourceCells: [{ column: 1, label: 'Extra', value: 'Original' }] }), customer('c'), customer('d')]);
  state = W.reducer(state, { type: 'delete', ids: ['b', 'd'] });
  state = W.reducer(state, { type: 'edit', id: 'a', key: 'contact', value: 'Updated' });
  state = W.reducer(state, { type: 'undoDelete' });
  assert.deepEqual(state.records.map(r => r.id), ['a', 'b', 'c', 'd']);
  assert.equal(state.records[0].contact, 'Updated');
  assert.equal(state.records[1].pendingIdentity, true);
  assert.equal(state.records[1].sourceCells[0].value, 'Original');
  assert.equal(state.deleted.length, 0);
});
test('only the last nonempty deletion is undoable, respecting current schema and record limit', () => {
  let state = workspace([customer('a', { phone: '123' }), customer('b', { phone: '456' })]);
  state = W.reducer(state, { type: 'delete', ids: ['a'] });
  state = W.reducer(state, { type: 'delete', ids: ['b'] });
  state = W.reducer(state, { type: 'delete', ids: ['missing'] });
  state = W.reducer(state, { type: 'removeField', key: 'phone' });
  state = W.reducer(state, { type: 'addField', name: '数量', key: 'custom_qty' });
  state = W.reducer(state, { type: 'undoDelete' });
  assert.deepEqual(state.records.map(r => r.id), ['b']);
  assert.equal(state.records[0].phone, undefined);
  assert.equal(state.records[0].custom_qty, '');
  assert.throws(() => W.reducer({ ...state, records: Array.from({ length: W.MAX_RECORDS }, (_, i) => customer(String(i))), deleted: [{ index: 0, record: customer('overflow') }] }, { type: 'undoDelete' }), /50,000/);
});
test('IndexedDB restores schema, exact data, pending confirmation and deletion recovery across connections', async () => {
  const factory = new IDBFactory();
  assert.equal(await L.readLocalWorkspace(factory), null);
  let state = workspace([customer('a', { pendingIdentity: true }), customer('b')]);
  state = W.reducer(state, { type: 'addField', key: 'custom_qty', name: '数量' });
  state = W.reducer(state, { type: 'edit', id: 'a', key: 'custom_qty', value: '0005' });
  state = W.reducer(state, { type: 'delete', ids: ['b'] });
  const saved = await L.writeLocalWorkspace(state, null, factory);
  const restored = await L.readLocalWorkspace(factory);
  assert.deepEqual(restored.workspace, state);
  assert.equal(restored.revision, saved.revision);
  assert.equal(W.summarize(restored.workspace.records).counts.missing, 1);
  assert.equal(W.reducer(restored.workspace, { type: 'undoDelete' }).records.length, 2);
});
test('empty library and field-only settings persist; export does not erase the saved workspace', async () => {
  const factory = new IDBFactory();
  let state = W.reducer(W.initialWorkspace(), { type: 'addField', key: 'custom_qty', name: '数量' });
  let saved = await L.writeLocalWorkspace(state, null, factory);
  state = W.reducer(state, { type: 'append', records: [customer('a')] });
  state = W.reducer(state, { type: 'exported' });
  saved = await L.writeLocalWorkspace(state, saved.revision, factory);
  assert.equal((await L.readLocalWorkspace(factory)).workspace.records.length, 1);
  state = W.reducer(state, { type: 'delete', ids: ['a'] });
  await L.writeLocalWorkspace(state, saved.revision, factory);
  const restored = (await L.readLocalWorkspace(factory)).workspace;
  assert.equal(restored.records.length, 0);
  assert.equal(restored.fields.at(-1).key, 'custom_qty');
  assert.equal(restored.deleted.length, 1);
});
test('two simultaneous writers cannot silently overwrite the same revision', async () => {
  const factory = new IDBFactory();
  const base = await L.writeLocalWorkspace(W.initialWorkspace(), null, factory);
  const results = await Promise.allSettled([L.writeLocalWorkspace(workspace([customer('a')]), base.revision, factory), L.writeLocalWorkspace(workspace([customer('b')]), base.revision, factory)]);
  assert.equal(results.filter(r => r.status === 'fulfilled').length, 1);
  assert.ok(results.find(r => r.status === 'rejected').reason instanceof L.SaveConflict);
  assert.equal((await L.readLocalWorkspace(factory)).workspace.records.length, 1);
});
test('unreadable saved versions are rejected, rather than reset to an empty project', () => {
  for (const value of [null, {}, { version: 2 }, { version: 1, revision: 'r', savedAt: 1, workspace: { ...W.initialWorkspace(), fields: [] } }]) assert.throws(() => L.decodeWorkspace(value), /未被覆盖/);
  const invalid = { version: 1, revision: 'r', savedAt: 1, workspace: workspace([customer('a'), customer('a')]) };
  assert.throws(() => L.decodeWorkspace(invalid), /未被覆盖/);
});
test('queued autosaves retain edits made during a write and report saved only after completion', async () => {
  const initial = W.initialWorkspace(), a = workspace([customer('a')]), b = workspace([customer('a'), customer('b')]);
  const writes = [], statuses = [];
  let release;
  const saver = L.createWorkspaceSaver(initial, null, async (state, revision) => {
    writes.push({ state, revision });
    if (writes.length === 1) await new Promise(resolve => { release = resolve; });
    return { version: 1, revision: 'r' + writes.length, savedAt: 1, workspace: state };
  }, value => statuses.push(value.status));
  saver.update(a); saver.update(b);
  assert.equal(saver.hasPending(), true);
  assert.deepEqual(statuses, ['saving']);
  release();
  await until(() => !saver.hasPending());
  assert.deepEqual(writes.map(w => w.state), [a, b]);
  assert.equal(writes[1].revision, 'r1');
  assert.equal(statuses.at(-1), 'saved');
});
test('failed saves retain pending edits, support retry, and conflicts never retry over new progress', async () => {
  let failure = true;
  const states = [];
  const saver = L.createWorkspaceSaver(W.initialWorkspace(), null, async state => {
    if (failure) throw Error('Quota exceeded');
    return { version: 1, revision: 'r', savedAt: 1, workspace: state };
  }, state => states.push(state.status));
  saver.update(workspace([customer('a')]));
  await until(() => states.includes('error'));
  assert.equal(saver.hasPending(), true);
  failure = false; await saver.retry();
  assert.equal(saver.hasPending(), false);
  let calls = 0;
  const conflict = L.createWorkspaceSaver(W.initialWorkspace(), null, async () => { calls++; throw new L.SaveConflict(); }, () => {});
  conflict.update(workspace([customer('a')])); await tick();
  await conflict.retry(); conflict.update(workspace([customer('b')])); await tick();
  assert.equal(calls, 1);
  assert.equal(conflict.hasPending(), true);
});
