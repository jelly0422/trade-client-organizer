const { test } = require("node:test");
const assert = require("node:assert/strict");
const C = require("../src/domain/core.ts");
const XLSX = require("../public/vendor/xlsx.full.min.js");

test("every supported separator splits addresses without splitting English periods", () => {
  for (const separator of [",", "，", ";", "；", "。", "\n", "\r", "\r\n"]) {
    const value =
      " first.last+tag@sub.example.com " + separator + " other@example.co.uk ";
    assert.deepEqual(C.splitEmails(value), [
      "first.last+tag@sub.example.com",
      "other@example.co.uk",
    ]);
    assert.deepEqual(C.issues({ email: value }), []);
  }
  const mixed =
    "a@example.com,b@example.com，c@example.com;d@example.com；e@example.com。f@example.com\ng@example.com";
  assert.equal(C.splitEmails(mixed).length, 7);
  assert.deepEqual(C.issues({ email: mixed }), []);
});

test("common valid mailboxes retain meaningful punctuation and case", () => {
  const values = [
    "a@b.co",
    "First.Last+tag@Sub.Example.COM",
    "sales_team@my-company.com",
    "o'connor@example.com",
    "a-b@example.com",
    "user@xn--bcher-kva.example",
    "!#$%&'*+/=?^_`{|}~-@example.com",
  ];
  for (const value of values) {
    assert.equal(C.isValidEmail(value), true, value);
    assert.deepEqual(C.issues({ email: value }), [], value);
  }
});

test("TLD format accepts short and long alphabetic suffixes with exact length boundaries", () => {
  for (const email of [
    "sales@example.cn",
    "sales@example.COM",
    "sales@example.technology",
    "sales@sub.example.co.uk",
    "sales@company123.co.uk",
    "sales@example." + "a".repeat(63),
    // Format validation intentionally does not imply an IANA delegation lookup.
    "sales@example.thisisnotaregisteredtld",
  ]) {
    assert.equal(C.isValidEmail(email), true, email);
  }
  for (const tld of [
    "a",
    "a".repeat(64),
    "12",
    "123",
    "c0m",
    "co9",
    "1com",
    "co-m",
    "co_m",
    "com!",
  ]) {
    assert.equal(C.isValidEmail("sales@example." + tld), false, tld);
  }
});

test("internationalized TLD encodings must be valid rather than merely starting with xn--", () => {
  for (const tld of ["xn--fiqs8s", "xn--p1ai", "XN--55QX5D"]) {
    assert.equal(C.isValidEmail("sales@example." + tld), true, tld);
  }
  for (const tld of [
    "xn--",
    "xn--a",
    "xn--abc",
    "xn--0",
    "xn--a-",
    "xn--" + "a".repeat(60),
  ]) {
    assert.equal(C.isValidEmail("sales@example." + tld), false, tld);
  }
});

test("each split address is reviewed when any position has a malformed TLD", () => {
  for (const separator of [",", "，", ";", "；", "。", "\n", "\r", "\r\n"]) {
    for (const invalid of [
      "bad@example.c",
      "bad@example.123",
      "bad@example.c0m",
      "bad@example.co-m",
      "bad@example.xn--abc",
    ]) {
      for (const emails of [
        [invalid, "good@example.com", "last@example.co.uk"],
        ["good@example.com", invalid, "last@example.co.uk"],
        ["good@example.com", "last@example.co.uk", invalid],
      ]) {
        assert.deepEqual(
          C.issues({ email: emails.join(separator) }),
          ["邮箱格式待核对"],
          emails.join(separator),
        );
      }
    }
  }
});

test("malformed addresses are reviewed in any position of a multi-email cell", () => {
  const invalid = [
    "sales@example..com",
    "sales@-example.com",
    "sales@example-.com",
    "sales@exa_mple.com",
    "a..b@example.com",
    ".user@example.com",
    "user.@example.com",
    "user@example.com.",
    "user@.example.com",
    "user@@example.com",
    "user@example",
    "@example.com",
    "user@",
    "first last@example.com",
    "user@exa mple.com",
    "user@exam\u200bple.com",
    "user@exam\u0000ple.com",
  ];
  for (const email of invalid) {
    assert.equal(C.isValidEmail(email), false, email);
    for (const value of [
      email,
      email + "。valid@example.com",
      "valid@example.com，" + email,
      "valid@example.com;" + email + "；last@example.com",
    ]) {
      assert.deepEqual(C.issues({ email: value }), ["邮箱格式待核对"], value);
    }
  }
});

test("email, username and domain label lengths enforce their exact boundaries", () => {
  assert.equal(C.isValidEmail("a".repeat(64) + "@example.com"), true);
  assert.equal(C.isValidEmail("a".repeat(65) + "@example.com"), false);
  assert.equal(C.isValidEmail("a@" + "b".repeat(63) + ".com"), true);
  assert.equal(C.isValidEmail("a@" + "b".repeat(64) + ".com"), false);
  const maxEmail =
    "a".repeat(64) +
    "@" +
    "b".repeat(63) +
    "." +
    "c".repeat(63) +
    "." +
    "d".repeat(61);
  assert.equal(maxEmail.length, 254);
  assert.equal(C.isValidEmail(maxEmail), true);
  assert.equal(C.isValidEmail(maxEmail + "d"), false);
  assert.equal(C.isValidEmail("a".repeat(32768) + "@example.com"), false);
});

test("blank email uses the missing-identity queue while empty split items require review", () => {
  for (const value of [undefined, null, "", " \t\r\n "]) {
    assert.deepEqual(C.splitEmails(value), []);
    assert.deepEqual(C.issues({ email: value }), []);
    assert.equal(
      C.missingIdentity({ company: "Acme", contact: "Alex", email: value }),
      true,
    );
  }
  for (const value of [
    "a@example.com;",
    ",a@example.com",
    "a@example.com。。b@example.com",
    "a@example.com， ;b@example.com",
    "。",
    ";，；",
  ]) {
    assert.deepEqual(C.issues({ email: value }), ["邮箱格式待核对"], value);
  }
  assert.deepEqual(
    C.issues({ email: " \n a@example.com\r\nb@example.com \n" }),
    [],
  );
});

test("unusual address forms remain available for manual review rather than silently accepted", () => {
  for (const email of [
    "客户@example.com",
    "user@例子.公司",
    '"first last"@example.com',
    "user@[127.0.0.1]",
    "Alex <user@example.com>",
    "a@example.com.b@example.com",
  ]) {
    const record = C.admitRecord({
      company: "Acme",
      contact: "Alex",
      email,
      pendingIdentity: true,
    });
    assert.equal(C.pendingIdentity(record), false);
    assert.equal(record.email, email);
    assert.deepEqual(C.issues(record), ["邮箱格式待核对"]);
  }
});

test("import, review and Excel export keep the original multi-email cell intact", () => {
  const email =
    "First.Last+tag@example.com。wrong@exa_mple.com； other@example.com";
  const rows = [
    ["Company", "Contact", "Email"],
    ["Acme", "Alex", email],
  ];
  const [record] = C.mappedRows(rows, 0, C.autoMap(rows[0]), "emails.xlsx");
  assert.equal(record.email, email);
  const matrix = C.exportMatrix([record]);
  const column = matrix[0].indexOf("邮箱"),
    hints = matrix[0].indexOf("校对提示");
  assert.equal(matrix[1][column], email);
  assert.equal(matrix[1][hints], "邮箱格式待核对");
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    book,
    XLSX.utils.aoa_to_sheet(matrix),
    "Clients",
  );
  const parsed = XLSX.read(
    XLSX.write(book, { type: "buffer", bookType: "xlsx" }),
    { type: "buffer" },
  );
  const exported = XLSX.utils.sheet_to_json(parsed.Sheets.Clients, {
    header: 1,
    raw: false,
  });
  assert.equal(exported[1][column], email);
  const corrected = {
    ...record,
    email: email.replace("wrong@exa_mple.com", "correct@example.com"),
  };
  assert.deepEqual(C.issues(corrected), []);
});
