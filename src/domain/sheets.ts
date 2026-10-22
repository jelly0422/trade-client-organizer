import type { Customer } from "./core.ts";

export interface WorkspaceSheet {
  id: string;
  name: string;
  sourceName: string;
}

export const UNGROUPED_SHEET = "__ungrouped__";

export function recordsInSheet(records: Customer[], id: string) {
  if (!id) return records;
  return records.filter(r => id === UNGROUPED_SHEET ? !r.sheetId : r.sheetId === id);
}

export function customerSheetGroups(records: Customer[], sheets: WorkspaceSheet[]) {
  const grouped = new Map<string, Customer[]>(sheets.map(s => [s.id, []]));
  const ungrouped: Customer[] = [];
  for (const record of records) {
    if (record.sheetId) {
      const group = grouped.get(record.sheetId);
      if (!group) throw Error("记录的工作表信息缺失，请先恢复对应工作表再导出。");
      group.push(record);
    } else ungrouped.push(record);
  }
  return [
    ...(ungrouped.length || !sheets.length
      ? [{ id: UNGROUPED_SHEET, name: "客户资料", sourceName: "合并导入 / 手动新增 / 历史记录", records: ungrouped }]
      : []),
    ...sheets.map(s => ({ ...s, records: grouped.get(s.id)! })),
  ];
}

// Excel names are case-insensitively unique, at most 31 UTF-16 units, and
// cannot contain these characters or begin/end with an apostrophe.
export function uniqueWorksheetName(name: string, used: Set<string>) {
  let base = name.replace(/[\\\/?*\[\]:\x00-\x1f]/g, "_").trim().replace(/^'+|'+$/g, "") || "工作表";
  if (base.toLowerCase() === "history") base += "_";
  const shorten = (value: string, length: number) => {
    let result = "";
    for (const character of value) {
      if (result.length + character.length > length) break;
      result += character;
    }
    return result;
  };
  let candidate = shorten(base, 31).replace(/'+$/g, "") || "工作表";
  for (let n = 2; used.has(candidate.toLowerCase()); n++) {
    const suffix = " (" + n + ")";
    candidate = shorten(base, 31 - suffix.length) + suffix;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}
