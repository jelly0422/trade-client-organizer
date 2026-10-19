import * as C from "../domain/core.ts";
import type { Customer, Field, Mapping } from "../domain/core.ts";

interface Sheet {
  [key: string]: unknown;
  "!ref"?: string;
}
export interface Workbook {
  SheetNames: string[];
  Sheets: Record<string, Sheet>;
}
interface SheetJS {
  read: (data: unknown, options: Record<string, unknown>) => Workbook;
  writeFile: (book: Workbook, filename: string) => void;
  utils: {
    sheet_to_json: (
      sheet: Sheet,
      options: Record<string, unknown>,
    ) => unknown[][];
    decode_range: (ref: string) => {
      s: { r: number; c: number };
      e: { r: number; c: number };
    };
    encode_col: (column: number) => string;
    book_new: () => Workbook;
    book_append_sheet: (book: Workbook, sheet: Sheet, title: string) => void;
    aoa_to_sheet: (rows: unknown[][]) => Sheet;
  };
}
declare global {
  interface Window {
    XLSX?: SheetJS;
  }
}
export interface SheetConfig {
  header: number;
  mapping: Mapping;
}
export interface FileSource {
  id: string;
  name: string;
  size: number;
  book: Workbook;
  bytes?: ArrayBuffer;
  sheet: string;
  configs: Record<string, SheetConfig>;
  encoding: string;
  demo?: boolean;
}
export const uid = () => crypto.randomUUID();
export function xlsx() {
  if (!window.XLSX) throw Error("Excel 解析组件未加载，请刷新页面后重试。");
  return window.XLSX;
}
export function parseBook(
  bytes: ArrayBuffer,
  name: string,
  encoding = "utf-8",
) {
  return /\.(csv|tsv)$/i.test(name)
    ? xlsx().read(new TextDecoder(encoding).decode(bytes), {
        type: "string",
        raw: true,
      })
    : xlsx().read(bytes, { type: "array", cellText: true });
}
export function rowsFor(source: FileSource) {
  return xlsx().utils.sheet_to_json(source.book.Sheets[source.sheet], {
    header: 1,
    defval: "",
    raw: false,
    blankrows: true,
  });
}
export function offsets(source: FileSource) {
  const ref = source.book.Sheets[source.sheet]["!ref"];
  return ref ? xlsx().utils.decode_range(ref).s : { r: 0, c: 0 };
}
export function withConfig(source: FileSource, fields: Field[]): FileSource {
  if (source.configs[source.sheet]) return source;
  const rows = rowsFor(source),
    header = C.detectHeader(rows, fields);
  return {
    ...source,
    configs: {
      ...source.configs,
      [source.sheet]: {
        header,
        mapping: C.autoMap(rows[header] || [], fields),
      },
    },
  };
}
export function synchronizeSchema(sources: FileSource[], fields: Field[]) {
  return sources.map((s) => ({
    ...s,
    configs: Object.fromEntries(
      Object.entries(s.configs).map(([sheet, cfg]) => {
        const rows = rowsFor({ ...s, sheet });
        return [
          sheet,
          {
            ...cfg,
            mapping: C.autoMap(rows[cfg.header] || [], fields, cfg.mapping),
          },
        ];
      }),
    ),
  }));
}
export function mappedSource(source: FileSource, fields: Field[]) {
  const cfg = source.configs[source.sheet],
    offset = offsets(source);
  return C.mappedRows(
    rowsFor(source),
    cfg.header,
    cfg.mapping,
    source.name + " / " + source.sheet,
    fields,
    offset.r,
    offset.c,
  );
}
export async function readFile(file: File, fields: Field[]) {
  if (!/\.(xlsx|xls|csv|tsv|ods)$/i.test(file.name))
    throw Error("不支持的文件格式：" + file.name);
  if (file.size > 20 * 1024 * 1024)
    throw Error(file.name + " 超过 20 MB 限制。");
  const bytes = await file.arrayBuffer();
  let book: Workbook;
  try {
    book = parseBook(bytes, file.name);
  } catch {
    throw Error("无法读取 " + file.name + "，请确认文件未加密且格式正确。");
  }
  if (!book.SheetNames.length) throw Error("没有可读取的工作表。");
  return withConfig(
    {
      id: uid(),
      name: file.name,
      size: file.size,
      book,
      bytes,
      sheet: book.SheetNames[0],
      configs: {},
      encoding: "utf-8",
    },
    fields,
  );
}
export function demoSource(fields: Field[]): FileSource {
  const rows = [
    [
      "Company Name",
      "Contact",
      "E-mail",
      "WhatsApp",
      "Country",
      "Product Interest",
      "Website",
      "Notes",
    ],
    [
      "Northstar Living",
      "Emma Wilson",
      "emma@example.com",
      "+44 7700 900001",
      "United Kingdom",
      "竹纤维家居用品",
      "https://example.com",
      "需提供产品目录",
    ],
    [
      "Hikari Trading",
      "Yuki Tanaka",
      "yuki@example.com",
      "+81 90 0000 0001",
      "Japan",
      "保温杯",
      "",
      "关注 MOQ 和交期",
    ],
    [
      "Verde Home",
      "Lucas Silva",
      "lucas@example",
      "+55 11 0000 0000",
      "Brazil",
      "环保厨房用品",
      "",
      "邮箱待核对",
    ],
    [
      "Northstar Living",
      "Emma Wilson",
      "emma@example.com",
      "",
      "United Kingdom",
      "定制包装",
      "",
      "历史询盘，可能重复",
    ],
    [
      "Atelier Maison",
      "Camille Martin",
      "",
      "",
      "France",
      "棉麻布艺",
      "",
      "需补充邮箱",
    ],
    [
      "Coastline Supply",
      "Alex Chen",
      "alex@example.com",
      "+1 202 555 0100",
      "United States",
      "户外家具",
      "",
      "索取样品报价",
    ],
  ];
  const book = xlsx().utils.book_new();
  xlsx().utils.book_append_sheet(
    book,
    xlsx().utils.aoa_to_sheet(rows),
    "展会客户",
  );
  return withConfig(
    {
      id: uid(),
      name: "示例 · 展会客户.xlsx",
      size: 0,
      book,
      sheet: "展会客户",
      configs: {},
      encoding: "utf-8",
      demo: true,
    },
    fields,
  );
}
function download(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob),
    link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
export function exportCustomers(
  records: Customer[],
  fields: Field[],
  format: "xlsx" | "csv",
  template = false,
) {
  const name = template
    ? "外贸客户标准模板"
    : "外贸客户资料_" + new Date().toLocaleDateString("sv-SE");
  const matrix = C.exportMatrix(template ? [] : records, fields, template);
  if (format === "csv" && !template) {
    const rows = matrix.map((row, i) => [
      ...row,
      i === 0 ? "原始内容" : JSON.stringify(records[i - 1].sourceCells || []),
    ]);
    download(
      new Blob(
        [
          "\ufeff" +
            rows.map((row) => row.map(C.csvCell).join(",")).join("\r\n"),
        ],
        { type: "text/csv;charset=utf-8" },
      ),
      name + ".csv",
    );
    return;
  }
  const book = xlsx().utils.book_new();
  const append = (rows: unknown[][], title: string) => {
    const sheet = xlsx().utils.aoa_to_sheet(rows);
    sheet["!cols"] = rows[0].map((h) => ({
      wch: Math.max(20, Math.min(40, String(h).length * 2 + 8)),
    }));
    if (rows.length > 1) sheet["!autofilter"] = { ref: sheet["!ref"] };
    xlsx().utils.book_append_sheet(book, sheet, title);
  };
  append(
    template
      ? matrix
      : C.exportMatrix(
          records.filter((r) => !C.pendingIdentity(r)),
          fields,
        ),
    "客户资料",
  );
  const pending = records.filter(C.pendingIdentity);
  if (!template && pending.length)
    append(C.exportMatrix(pending, fields), "缺少主体");
  const raw = C.recoveryMatrix(template ? [] : records);
  if (raw.length > 1) append(raw, "导入原始内容");
  xlsx().writeFile(book, name + ".xlsx");
}
