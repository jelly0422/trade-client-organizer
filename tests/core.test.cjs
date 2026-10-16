const { test } = require("node:test");
const assert = require("node:assert/strict");
const C = require("../src/domain/core.ts");
const XLSX = require("../public/vendor/xlsx.full.min.js");
test("recognizes bilingual aliases and title rows", () => {
  const rows = [
    ["2026 展会名单"],
    [
      "公司名称",
      "Contact",
      "E-mail",
      "WhatsApp",
      "Country",
      "Product Interest",
    ],
    ["Acme", "A", "a@example.com", "00123", "UK", "Cup"],
  ];
  assert.equal(C.detectHeader(rows), 1);
  const mapped = C.mappedRows(rows, 1, C.autoMap(rows[1]), "test");
  assert.equal(mapped[0].phone, "00123");
  assert.equal(mapped[0].product, "Cup");
  assert.equal(mapped[0].sourceRow, 3);
});
test("manual mapping and blank rows preserve row provenance", () => {
  const rows = [["客户", "需求"], [], ["公司甲", "纸杯"], ["", ""]];
  const result = C.mappedRows(rows, 0, { company: "0", product: "1" }, "sheet");
  assert.equal(result.length, 1);
  assert.equal(result[0].sourceRow, 3);
  assert.equal(result[0].email, "");
});
test("review checks email format while missing fixed fields use their own classification", () => {
  assert.equal(C.missingIdentity({}), true);
  assert.deepEqual(C.issues({}), []);
  assert.deepEqual(
    C.issues({ company: "A", contact: "Alex", email: "a@b.com; b@c.com" }),
    [],
  );
  assert.deepEqual(C.issues({ company: "A", email: "invalid" }), [
    "邮箱格式待核对",
  ]);
});
test("duplicate detection includes both originals without matching blank identities", () => {
  assert.deepEqual(
    [
      ...C.duplicateIds([
        { id: "1", email: "A@B.com" },
        { id: "2", email: "a@b.com" },
        { id: "3" },
        { id: "4" },
      ]),
    ],
    ["1", "2"],
  );
});
test("CSV preserves quoting and mitigates formula prefixes", () => {
  assert.equal(C.csvCell('=HYPERLINK("x")'), '"\'=HYPERLINK(""x"")"');
  assert.equal(C.csvCell("Acme, Inc."), '"Acme, Inc."');
});
for (const type of ["xlsx", "xls", "ods", "csv"])
  test(
    type +
      " parsing preserves international phone strings and multilingual cells",
    () => {
      const data = [
        ["Company", "Phone", "Product"],
        ["公司甲", "+44 001234", "竹制餐具"],
        ["Beta", "001234", "Cups"],
      ];
      const book = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(
        book,
        XLSX.utils.aoa_to_sheet(data),
        "Clients",
      );
      const bytes = XLSX.write(book, { type: "buffer", bookType: type });
      const read =
        type === "csv"
          ? XLSX.read(new TextDecoder().decode(bytes), {
              type: "string",
              raw: true,
            })
          : XLSX.read(bytes, { type: "buffer" });
      const rows = XLSX.utils.sheet_to_json(read.Sheets[read.SheetNames[0]], {
        header: 1,
        raw: false,
      });
      const records = C.mappedRows(rows, 0, C.autoMap(rows[0]), type);
      assert.equal(records[0].company, "公司甲");
      assert.equal(records[0].phone, "+44 001234");
      assert.equal(records[1].phone, "001234");
    },
  );
test("XLSX export keeps formula-like customer text as strings", () => {
  const sheet = XLSX.utils.aoa_to_sheet([["=1+1", "+12345", "00123"]]);
  assert.equal(sheet.A1.t, "s");
  assert.equal(sheet.A1.f, undefined);
  assert.equal(sheet.C1.v, "00123");
});
test("fixed fields cannot be removed and customer/company headers map distinctly", () => {
  for (const key of ["company", "contact", "email"])
    assert.throws(() => C.removeField(C.fields, key), /不能删除/);
  const mapping = C.autoMap(["客户名称", "公司名", "邮箱"]);
  assert.equal(mapping.contact, "0");
  assert.equal(mapping.company, "1");
  assert.equal(mapping.email, "2");
});
test("custom schema flows through mapping, data, template, XLSX and CSV export", () => {
  let schema = C.removeField(C.fields, "phone");
  schema = C.addField(schema, "采购数量", "custom_quantity");
  const rows = [
    ["公司名", "客户名称", "邮箱", "电话", "采购数量"],
    ["Acme", "Emma", "emma@example.com", "00123", "00800"],
  ];
  const records = C.mappedRows(
    rows,
    0,
    C.autoMap(rows[0], schema),
    "test",
    schema,
  );
  assert.equal(records[0].custom_quantity, "00800");
  assert.equal(Object.hasOwn(records[0], "phone"), false);
  records[0].custom_quantity = "00900";
  const matrix = C.exportMatrix(records, schema),
    column = matrix[0].indexOf("采购数量");
  assert.equal(matrix[0].includes("电话"), false);
  assert.equal(matrix[1][column], "00900");
  assert.equal(C.exportMatrix([], schema, true)[0][column], "采购数量");
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    book,
    XLSX.utils.aoa_to_sheet(matrix),
    "Clients",
  );
  const exported = XLSX.read(
    XLSX.write(book, { type: "buffer", bookType: "xlsx" }),
    { type: "buffer" },
  );
  const back = XLSX.utils.sheet_to_json(exported.Sheets.Clients, {
    header: 1,
    raw: false,
  });
  assert.equal(back[1][column], "00900");
  assert.ok(
    matrix
      .map((row) => row.map(C.csvCell).join(","))
      .join("\n")
      .includes('"00900"'),
  );
  schema = C.removeField(schema, "custom_quantity");
  assert.equal(C.exportMatrix(records, schema)[0].includes("采购数量"), false);
});
test("schema changes preserve manual mappings and free removed columns", () => {
  let schema = C.removeField(C.fields, "phone");
  schema = C.addField(schema, "WhatsApp", "custom_whatsapp");
  const mapping = C.autoMap(["Org", "Person", "E-mail", "WhatsApp"], schema, {
    company: "0",
    contact: "",
    email: "2",
    phone: "3",
  });
  assert.equal(mapping.company, "0");
  assert.equal(mapping.contact, "");
  assert.equal(mapping.custom_whatsapp, "3");
  assert.equal(Object.hasOwn(mapping, "phone"), false);
});
test("field names validate empty, duplicate, reserved and excessive input", () => {
  for (const name of [" ", "邮箱", " 邮 箱 ", "来源文件 / 工作表"]) {
    assert.throws(() => C.addField(C.fields, name, "custom_new"));
  }
  assert.throws(() => C.addField(C.fields, "a".repeat(31), "custom_new"));
  assert.throws(() => C.addField(C.fields, "新字段", "__proto__"));
  assert.throws(() =>
    C.addField(
      Array.from({ length: 40 }, (_, i) => ({
        key: "k" + i,
        label: "字段" + i,
        aliases: [],
      })),
      "新字段",
      "custom_new",
    ),
  );
});
test("phone cannot replace a required email and restoring phone retains aliases", () => {
  const schema = C.removeField(C.fields, "phone"),
    record = { company: "Acme", contact: "Alex", phone: "00123" };
  assert.equal(C.missingIdentity(record), true);
  assert.throws(() => C.admitRecord(record), /邮箱/);
  assert.throws(
    () =>
      C.admitRecord({ ...record, phone: undefined, custom_whatsapp: "00123" }),
    /邮箱/,
  );
  const restored = C.addField(schema, "电话", "custom_restore");
  assert.equal(C.autoMap(["WhatsApp"], restored).phone, "0");
});
test("unmapped nonblank rows are retained with their original values, including zero", () => {
  const rows = [
    ["Company", "Contact", "Email", "Other"],
    ["Acme", "Alex", "a@example.com", ""],
    ["", "", "", "客户乙：需要 10000 件"],
    [],
    ["", "", "", "   "],
    ["", "", "", "0"],
    ["", "", "", "", 0],
  ];
  const records = C.mappedRows(rows, 0, C.autoMap(rows[0]), "source.xlsx");
  assert.equal(records.length, 4);
  assert.deepEqual(
    records.map((r) => r.sourceRow),
    [2, 3, 6, 7],
  );
  assert.equal(C.pendingIdentity(records[0]), false);
  assert.equal(C.pendingIdentity(records[1]), true);
  assert.deepEqual(records[1].sourceCells, [
    { column: 4, label: "Other", value: "客户乙：需要 10000 件" },
  ]);
  assert.deepEqual(records[3].sourceCells, [
    { column: 5, label: "第 5 列", value: "0" },
  ]);
});
test("zero matched columns still preserve all nonblank source rows", () => {
  const rows = [["未知列"], ["仅有未识别内容"], ["第二条"], [""]];
  const records = C.mappedRows(rows, 0, C.autoMap(rows[0]), "unmatched.csv");
  assert.equal(records.length, 2);
  assert.ok(records.every((r) => C.pendingIdentity(r)));
  assert.equal(records[0].sourceCells[0].value, "仅有未识别内容");
});
test("all eight fixed-field combinations require all three values before admission", () => {
  assert.equal(
    C.missingIdentity({
      company: " ",
      contact: "\t",
      email: "\n",
      phone: "123",
      product: "cups",
    }),
    true,
  );
  for (let mask = 0; mask < 8; mask++) {
    const record = Object.fromEntries(
      C.fixedKeys.map((key, i) => [
        key,
        mask & (1 << i) ? "provided" : " \t\n",
      ]),
    );
    assert.equal(C.missingIdentity(record), mask !== 7, "combination " + mask);
    if (mask !== 7)
      assert.throws(() => C.admitRecord(record), /请先补齐固定字段/);
    else assert.equal(C.pendingIdentity(C.admitRecord(record)), false);
  }
  assert.equal(
    C.missingIdentity({ company: 0, contact: "Alex", email: "a@example.com" }),
    false,
  );
  assert.deepEqual(C.missingFixedFields({ company: "Acme" }), [
    "客户名称",
    "邮箱",
  ]);
});
test("admission needs all fixed fields and explicit confirmation without losing source data", () => {
  const pending = C.mappedRows(
    [["Other"], ["待补充的需求"]],
    0,
    {},
    "original",
  )[0];
  assert.throws(() => C.admitRecord(pending), /公司名、客户名称、邮箱/);
  assert.throws(
    () => C.admitRecord({ ...pending, company: "Acme" }),
    /客户名称、邮箱/,
  );
  assert.throws(
    () => C.admitRecord({ ...pending, company: "Acme", contact: "Alex" }),
    /邮箱/,
  );
  const edited = {
    ...pending,
    company: "Acme",
    contact: "Alex",
    email: "a@example.com",
  };
  assert.equal(C.pendingIdentity(edited), true);
  const accepted = C.admitRecord(edited);
  assert.equal(C.pendingIdentity(accepted), false);
  assert.equal(edited.pendingIdentity, true);
  assert.deepEqual(accepted.sourceCells, pending.sourceCells);
  assert.deepEqual(C.issues(accepted), []);
  for (const key of C.fixedKeys)
    assert.equal(C.pendingIdentity({ ...accepted, [key]: " \t" }), true);
});
test("source recovery keeps actual row and column positions for offset worksheets", () => {
  const sheet = {
    "!ref": "C5:D7",
    C5: { t: "s", v: "Company" },
    D5: { t: "s", v: "Other" },
    C6: { t: "s", v: "Acme" },
    D7: { t: "s", v: "未知主体需求" },
  };
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, "Offset");
  const parsed = XLSX.read(
    XLSX.write(book, { type: "buffer", bookType: "xlsx" }),
    { type: "buffer" },
  ).Sheets.Offset;
  const rows = XLSX.utils.sheet_to_json(parsed, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: true,
  });
  const offset = XLSX.utils.decode_range(parsed["!ref"]).s;
  const records = C.mappedRows(
    rows,
    0,
    C.autoMap(rows[0]),
    "offset.xlsx",
    C.fields,
    offset.r,
    offset.c,
  );
  assert.deepEqual(
    records.map((r) => r.sourceRow),
    [6, 7],
  );
  assert.equal(records[1].sourceCells[0].column, 4);
  assert.ok(
    C.recoveryMatrix(records).some((row) => row[1] === "7" && row[2] === "4"),
  );
});
test("review checks every email and ignores optional-field gaps after admission", () => {
  const record = {
    company: "Acme",
    contact: "Alex",
    email: "person@example.com",
    phone: "",
    product: "",
    country: "",
  };
  assert.deepEqual(C.issues(record), []);
  assert.deepEqual(
    C.issues({
      ...record,
      email:
        "a@example.com；b@example.com\nc@example.com,d@example.com，e@example.com",
    }),
    [],
  );
  const invalid = C.admitRecord({
    ...record,
    pendingIdentity: true,
    email: "a@example.com;invalid",
  });
  assert.equal(C.pendingIdentity(invalid), false);
  assert.deepEqual(C.issues(invalid), ["邮箱格式待核对"]);
  assert.deepEqual(C.issues({ ...record, email: "a@example.com;" }), [
    "邮箱格式待核对",
  ]);
  assert.deepEqual(
    C.reviewRules.map((r) => r.label),
    ["邮箱格式待核对"],
  );
});
test("exports distinguish queued records and retain original content after admission", () => {
  const pending = C.mappedRows(
    [["Extra"], ["=untrusted text"]],
    0,
    {},
    "original.xlsx",
  )[0];
  const ready = {
    ...pending,
    company: "Acme",
    contact: "Alex",
    email: "a@example.com",
    id: "pending",
  };
  const accepted = { ...C.admitRecord(ready), id: "accepted" };
  const matrix = C.exportMatrix([ready, accepted]);
  const classification = matrix[0].indexOf("记录分类"),
    hints = matrix[0].indexOf("校对提示");
  assert.equal(matrix[1][classification], "缺少主体");
  assert.equal(matrix[2][classification], "全部客户");
  assert.match(matrix[1][hints], /待确认加入/);
  const raw = C.recoveryMatrix([accepted]);
  assert.equal(raw[1][4], "=untrusted text");
  const sheet = XLSX.utils.aoa_to_sheet(raw);
  assert.equal(sheet.E2.t, "s");
  assert.equal(sheet.E2.f, undefined);
});
test("partially populated imports and exports report only the actual missing fields", () => {
  const rows = [
    ["Company", "Contact", "Email", "Other"],
    ["", "Alex", "a@example.com", "need1"],
    ["Acme", "", "b@example.com", "need2"],
    ["Beta", "Pat", " \t", "need3"],
  ];
  const records = C.mappedRows(rows, 0, C.autoMap(rows[0]), "partial.xlsx");
  assert.equal(records.length, 3);
  assert.ok(records.every(C.pendingIdentity));
  assert.deepEqual(records.map(C.missingFixedFields), [
    ["公司名"],
    ["客户名称"],
    ["邮箱"],
  ]);
  assert.ok(
    records.every((r) => r.sourceCells.some((cell) => cell.label === "Other")),
  );
  const matrix = C.exportMatrix(records),
    hints = matrix[0].indexOf("校对提示"),
    category = matrix[0].indexOf("记录分类");
  assert.deepEqual(
    matrix.slice(1).map((r) => r[hints]),
    ["缺少固定字段：公司名", "缺少固定字段：客户名称", "缺少固定字段：邮箱"],
  );
  assert.ok(matrix.slice(1).every((r) => r[category] === "缺少主体"));
});
