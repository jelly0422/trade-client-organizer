const { test } = require("node:test");
const assert = require("node:assert/strict");
const W = require("../src/state/workspace.ts");
const C = require("../src/domain/core.ts");
const customer = (id, extra = {}) => ({
  id,
  company: "Acme",
  contact: "Alex",
  email: id + "@example.com",
  ...extra,
});

test("editing fixed fields preserves the explicit missing-identity admission lifecycle", () => {
  let state = W.reducer(W.initialWorkspace(), {
    type: "append",
    records: [customer("1")],
  });
  state = W.reducer(state, { type: "edit", id: "1", key: "email", value: " " });
  assert.equal(W.summarize(state.records).counts.missing, 1);
  state = W.reducer(state, {
    type: "edit",
    id: "1",
    key: "email",
    value: "valid@example.com。invalid",
  });
  assert.equal(W.summarize(state.records).counts.all, 0);
  state = W.reducer(state, { type: "admit", ids: ["1"] });
  assert.deepEqual(W.summarize(state.records).counts, {
    all: 1,
    review: 1,
    duplicate: 0,
    missing: 0,
  });
  state = W.reducer(state, {
    type: "edit",
    id: "1",
    key: "email",
    value: "valid@example.com。other@example.com",
  });
  assert.equal(W.summarize(state.records).counts.review, 0);
});
test("batch admission retains incomplete records and original provenance", () => {
  const sourceCells = [{ column: 1, label: "Other", value: "original" }];
  let state = W.reducer(W.initialWorkspace(), {
    type: "append",
    records: [
      customer("a", { pendingIdentity: true, sourceCells }),
      customer("b", { contact: "", pendingIdentity: true }),
    ],
  });
  state = W.reducer(state, { type: "admit", ids: ["a", "b"] });
  assert.equal(C.pendingIdentity(state.records[0]), false);
  assert.equal(C.pendingIdentity(state.records[1]), true);
  assert.deepEqual(state.records[0].sourceCells, sourceCells);
});
test("schema changes preserve records and validation while rejecting fixed-field removal", () => {
  let state = W.reducer(W.initialWorkspace(), {
    type: "append",
    records: [customer("a")],
  });
  state = W.reducer(state, {
    type: "addField",
    name: "采购数量",
    key: "custom_quantity",
  });
  state = W.reducer(state, {
    type: "edit",
    id: "a",
    key: "custom_quantity",
    value: "00080",
  });
  assert.equal(state.records[0].custom_quantity, "00080");
  assert.throws(
    () => W.reducer(state, { type: "removeField", key: "email" }),
    /不能删除/,
  );
  state = W.reducer(state, { type: "removeField", key: "custom_quantity" });
  assert.equal(Object.hasOwn(state.records[0], "custom_quantity"), false);
  assert.equal(state.records.length, 1);
});
test("record cap applies to appending imports and single manual records", () => {
  const state = {
    ...W.initialWorkspace(),
    records: Array.from({ length: 50000 }, (_, i) => customer(String(i))),
  };
  assert.throws(
    () => W.reducer(state, { type: "append", records: [customer("overflow")] }),
    /50,000/,
  );
  assert.equal(state.records.length, 50000);
});
test("classification and search include unmapped source cells and numeric zero", () => {
  const records = [
    customer("a", { phone: 0 }),
    customer("b", { email: "bad" }),
    customer("c", { email: "a@example.com" }),
    customer("d", {
      company: "",
      pendingIdentity: true,
      sourceCells: [{ column: 1, label: "Extra", value: "特别需求" }],
    }),
  ];
  const summary = W.summarize(records);
  assert.deepEqual(summary.counts, {
    all: 3,
    review: 1,
    duplicate: 2,
    missing: 1,
  });
  assert.equal(
    W.filterRecords(records, C.fields, "all", "0", summary).length,
    1,
  );
  assert.equal(
    W.filterRecords(records, C.fields, "missing", "特别需求", summary).length,
    1,
  );
  assert.equal(
    W.filterRecords(records, C.fields, "all", "特别需求", summary).length,
    0,
  );
});
test("export marks current state saved; further edits mark it dirty again", () => {
  let state = W.reducer(W.initialWorkspace(), {
    type: "append",
    records: [customer("a")],
  });
  state = W.reducer(state, { type: "exported" });
  assert.equal(state.dirty, false);
  state = W.reducer(state, {
    type: "edit",
    id: "a",
    key: "contact",
    value: "Updated",
  });
  assert.equal(state.dirty, true);
  state = W.reducer(state, { type: "delete", ids: ["a"] });
  assert.equal(state.records.length, 0);
});
