export interface Field {
  key: string;
  label: string;
  aliases: string[];
}
export interface SourceCell {
  column: number;
  label: string;
  value: string;
}
export interface CustomerRecord {
  [key: string]: unknown;
  id?: string;
  source?: string;
  sourceRow?: number | string;
  pendingIdentity?: boolean;
  sourceCells?: SourceCell[];
}
export interface Customer extends CustomerRecord {
  id: string;
}
export type Mapping = Record<string, string>;
interface ReviewRule {
  label: string;
  description: string;
  check: (record: CustomerRecord, schema?: Field[]) => boolean;
}

export const fields: Field[] = [
  {
    key: "company",
    label: "公司名",
    aliases: [
      "公司",
      "公司名",
      "公司名称",
      "企业名称",
      "company",
      "company name",
      "organization",
      "account name",
    ],
  },
  {
    key: "contact",
    label: "客户名称",
    aliases: [
      "客户名称",
      "联系人",
      "姓名",
      "客户姓名",
      "customer",
      "contact",
      "contact name",
      "name",
      "full name",
      "buyer",
    ],
  },
  {
    key: "email",
    label: "邮箱",
    aliases: [
      "邮箱",
      "电子邮箱",
      "电子邮件",
      "email",
      "e-mail",
      "email address",
      "mail",
    ],
  },
  {
    key: "phone",
    label: "电话",
    aliases: [
      "电话",
      "联系电话",
      "手机",
      "手机号",
      "phone",
      "telephone",
      "tel",
      "mobile",
      "whatsapp",
    ],
  },
  {
    key: "country",
    label: "国家 / 地区",
    aliases: [
      "国家",
      "地区",
      "国家地区",
      "country",
      "region",
      "market",
      "country / region",
    ],
  },
  {
    key: "product",
    label: "产品需求",
    aliases: [
      "产品需求",
      "需求",
      "产品",
      "意向产品",
      "采购产品",
      "product",
      "products",
      "product interest",
      "requirements",
      "inquiry",
    ],
  },
  {
    key: "website",
    label: "公司网站",
    aliases: ["网址", "网站", "公司网站", "website", "url", "web"],
  },
  {
    key: "notes",
    label: "备注",
    aliases: ["备注", "说明", "notes", "note", "remarks", "comment"],
  },
];
export const clean = (v: unknown) => String(v ?? "").trim();
const normalize = (v: unknown) =>
  clean(v)
    .toLowerCase()
    .replace(/[\s_\-\/（）()]/g, "");
export const fixedKeys = Object.freeze(["company", "contact", "email"]);
const metadataLabels = [
  "来源",
  "来源文件 / 工作表",
  "原始行号",
  "校对提示",
  "状态",
  "记录分类",
  "原始内容",
];
export const missingIdentity = (record: CustomerRecord) =>
  fixedKeys.some((key) => !clean(record[key]));
export const pendingIdentity = (record: CustomerRecord) =>
  Boolean(record.pendingIdentity) || missingIdentity(record);
export const missingFixedFields = (record: CustomerRecord) =>
  fields
    .filter((f) => fixedKeys.includes(f.key) && !clean(record[f.key]))
    .map((f) => f.label);
export function admitRecord<T extends CustomerRecord>(record: T): T {
  if (missingIdentity(record))
    throw Error(
      "请先补齐固定字段：" + missingFixedFields(record).join("、") + "。",
    );
  return { ...record, pendingIdentity: false };
}
// Common unquoted ASCII mailboxes; unusual address forms remain available for manual review.
// A period is part of a mailbox/domain, so only the Chinese full stop is a separator.
const emailPattern =
  /^[A-Za-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[A-Za-z0-9!#$%&'*+\/=?^_`{|}~-]+)*@(?:[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)+[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?$/;
export const emailTldDescription =
  "每个邮箱最后一个英文句点后的顶级域名，须为 2–63 个英文字母（不区分大小写），如 .com、.cn、.technology；.co.uk 检查最后一段 uk。也支持有效的 xn-- 国际化域名编码。单字母、纯数字、普通后缀混入数字或短横线均列为待完善。";
function hasValidEmailTld(email: string) {
  const tld = email.slice(email.lastIndexOf(".") + 1).toLowerCase();
  if (/^[a-z]{2,63}$/.test(tld)) return true;
  if (!/^xn--[a-z0-9](?:[a-z0-9-]{0,57}[a-z0-9])?$/.test(tld)) return false;
  // WHATWG URL parsing validates the IDNA encoding locally; this makes no request.
  // Only syntax is checked, not IANA delegation, DNS, or mailbox deliverability.
  // A fixed Unicode label forces the IDNA conversion path even in runtimes
  // that fast-path all-ASCII hostnames without validating existing A-labels.
  const hostname = "é.email-tld-check." + tld;
  try {
    return new URL("https://" + hostname).hostname === "xn--9ca.email-tld-check." + tld;
  } catch {
    return false;
  }
}
export function splitEmails(value: unknown) {
  const text = clean(value);
  return text ? text.split(/[,，;；。]|\r\n?|\n/).map(clean) : [];
}
export function isValidEmail(value: unknown) {
  const email = clean(value);
  // Check lengths before the regex so very long imported cells are cheap to reject.
  return (
    email.length <= 254 &&
    email.indexOf("@") <= 64 &&
    emailPattern.test(email) &&
    hasValidEmailTld(email)
  );
}
export const reviewRules: ReviewRule[] = [
  {
    label: "邮箱格式待核对",
    description:
      "多个邮箱按中英文逗号（,，）、分号（;；）、中文句号（。）或换行拆分，再逐个用正则表达式校验，并检查每个邮箱的顶级域名格式；任意一项不通过就列为待完善。英文句点（.）保留在邮箱中。",
    check: (r) => splitEmails(r.email).some((email) => !isValidEmail(email)),
  },
];
export function addField(schema: Field[], name: string, key: string) {
  const label = clean(name);
  if (!normalize(label)) throw Error("请输入字段名称。");
  if (label.length > 30) throw Error("字段名称不能超过 30 个字符。");
  if (schema.length >= 40) throw Error("最多可设置 40 个字段。");
  if (
    schema.some((f) => normalize(f.label) === normalize(label)) ||
    metadataLabels.some((n) => normalize(n) === normalize(label))
  )
    throw Error("字段名称已存在或用于系统信息，请换一个名称。");
  if (!/^custom_[a-z0-9]+$/i.test(key) || schema.some((f) => f.key === key))
    throw Error("自定义字段标识无效。");
  const builtIn = fields.find(
    (f) =>
      normalize(f.label) === normalize(label) &&
      !schema.some((existing) => existing.key === f.key),
  );
  return [
    ...schema,
    builtIn ? { ...builtIn } : { key, label, aliases: [label] },
  ];
}
export function removeField(schema: Field[], key: string) {
  if (fixedKeys.includes(key))
    throw Error("客户名称、公司名和邮箱是固定字段，不能删除。");
  return schema.filter((f) => f.key !== key);
}
export function autoMap(
  headers: unknown[],
  schema = fields,
  previous: Mapping = {},
) {
  const used = new Set<number>();
  const mapping: Mapping = {};
  // Preserve manual choices, including explicitly unmapped fields.
  schema.forEach((f) => {
    if (!Object.hasOwn(previous, f.key)) return;
    const value = previous[f.key];
    if (value === "") mapping[f.key] = "";
    else if (
      /^\d+$/.test(value) &&
      Number(value) < headers.length &&
      !used.has(Number(value))
    ) {
      mapping[f.key] = value;
      used.add(Number(value));
    }
  });
  return Object.fromEntries(
    schema.map((f) => {
      if (Object.hasOwn(mapping, f.key)) return [f.key, mapping[f.key]];
      const index = headers.findIndex(
        (h, i) =>
          !used.has(i) &&
          [f.label, ...f.aliases].some((a) => normalize(a) === normalize(h)),
      );
      if (index >= 0) used.add(index);
      return [f.key, index < 0 ? "" : String(index)];
    }),
  );
}
export function detectHeader(rows: unknown[][], schema = fields) {
  let best = 0,
    score = -1;
  rows.slice(0, 30).forEach((row, i) => {
    const s = Object.values(autoMap(row, schema)).filter(
      (v) => v !== "",
    ).length;
    if (s > score) {
      score = s;
      best = i;
    }
  });
  return best;
}
export function mappedRows(
  rows: unknown[][],
  headerIndex: number,
  mapping: Mapping,
  source: string,
  schema = fields,
  rowOffset = 0,
  columnOffset = 0,
): CustomerRecord[] {
  return rows.slice(headerIndex + 1).flatMap((row, i) => {
    if (!row.some((v) => clean(v))) return [];
    const record = Object.fromEntries(
      schema.map((f) => [
        f.key,
        mapping[f.key] === "" || mapping[f.key] === undefined
          ? ""
          : clean(row[Number(mapping[f.key])]),
      ]),
    );
    const pending = missingIdentity(record);
    const sourceCells = row
          .map((value, column) => ({
            column: columnOffset + column + 1,
            label:
              clean(rows[headerIndex]?.[column]) ||
              `第 ${columnOffset + column + 1} 列`,
            value: String(value ?? ""),
          }))
          .filter((cell) => clean(cell.value));
    return [
      {
        ...record,
        source,
        sourceRow: rowOffset + headerIndex + i + 2,
        pendingIdentity: pending,
        sourceCells,
      },
    ];
  });
}
export function issues(record: CustomerRecord, schema = fields) {
  return reviewRules
    .filter((rule) => rule.check(record, schema))
    .map((rule) => rule.label);
}
export function duplicateIds(records: CustomerRecord[]) {
  const groups = new Map<string, string | undefined>(),
    ids = new Set<string | undefined>();
  records.forEach((r) => {
    // Compare individual valid addresses. Preserve punctuation and local-part case:
    // only domain names are case-insensitive. Repeating an address inside one row
    // does not make that row a duplicate of itself.
    const keys = new Set(splitEmails(r.email).filter(isValidEmail).map((email) => {
      const at = email.lastIndexOf("@");
      return email.slice(0, at) + "@" + email.slice(at + 1).toLowerCase();
    }));
    for (const key of keys) {
      if (groups.has(key)) {
        ids.add(groups.get(key));
        ids.add(r.id);
      } else groups.set(key, r.id);
    }
  });
  return ids;
}
export function csvCell(value: unknown) {
  let v = String(value ?? "");
  if (/^[\s]*[=+\-@\t\r]/.test(v)) v = "'" + v;
  return '"' + v.replace(/"/g, '""') + '"';
}
export function exportMatrix(
  records: CustomerRecord[],
  schema = fields,
  template = false,
) {
  const headers = schema.map((f) => f.label);
  if (template) return [headers];
  const duplicates = duplicateIds(records.filter((r) => !pendingIdentity(r)));
  return [
    [...headers, "来源文件 / 工作表", "原始行号", "校对提示", "记录分类"],
    ...records.map((r) => [
      ...schema.map((f) => String(r[f.key] ?? "")),
      String(r.source ?? ""),
      String(r.sourceRow || ""),
      [
        ...(pendingIdentity(r)
          ? [
              missingIdentity(r)
                ? "缺少固定字段：" + missingFixedFields(r).join("、")
                : "已补齐三项固定字段，待确认加入全部客户",
            ]
          : []),
        ...issues(r, schema),
        ...(duplicates.has(r.id) ? ["疑似重复"] : []),
      ].join("；"),
      pendingIdentity(r) ? "缺少主体" : "全部客户",
    ]),
  ];
}
export function recoveryMatrix(records: CustomerRecord[]) {
  return [
    ["来源文件 / 工作表", "原始行号", "原始列序号", "原始表头", "原始内容"],
    ...records.flatMap((r) =>
      (r.sourceCells || []).map((cell) => [
        String(r.source ?? ""),
        String(r.sourceRow || ""),
        String(cell.column),
        cell.label,
        cell.value,
      ]),
    ),
  ];
}
