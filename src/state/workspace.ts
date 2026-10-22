import * as C from "../domain/core.ts";
import type { Customer, Field } from "../domain/core.ts";
import { UNGROUPED_SHEET, type WorkspaceSheet } from "../domain/sheets.ts";

export type Category = "all" | "review" | "duplicate" | "missing";
export const MAX_RECORDS = 50000;
export interface Workspace {
  fields: Field[];
  records: Customer[];
  dirty: boolean;
  deleted: { record: Customer; index: number }[];
  deletedSheet: { sheet: WorkspaceSheet; index: number } | null;
  sheets: WorkspaceSheet[];
}
export const initialWorkspace = (): Workspace => ({
  fields: C.fields.map((f) => ({ ...f })),
  records: [],
  dirty: false,
  deleted: [],
  deletedSheet: null,
  sheets: [],
});
export type Action =
  | { type: "restore"; workspace: Workspace }
  | { type: "append"; records: Customer[]; sheets?: WorkspaceSheet[] }
  | { type: "edit"; id: string; key: string; value: string }
  | { type: "delete"; ids: string[] }
  | { type: "deleteSheet"; id: string }
  | { type: "undoDelete" }
  | { type: "admit"; ids: string[] }
  | { type: "addField"; name: string; key: string }
  | { type: "removeField"; key: string }
  | { type: "exported" };

export function reducer(state: Workspace, action: Action): Workspace {
  if (action.type === "restore") return action.workspace;
  if (action.type === "exported") return { ...state, dirty: false };
  if (action.type === "append") {
    if (state.records.length + action.records.length > MAX_RECORDS)
      throw Error("当前最多保留 50,000 条记录，请分批整理。");
    const sheets = [...state.sheets];
    const ids = new Set(sheets.map(s => s.id));
    for (const sheet of action.sheets || []) {
      if (ids.has(sheet.id)) throw Error("这张工作表已导入，请勿重复添加。");
      if (state.deletedSheet?.sheet.id === sheet.id)
        throw Error("这张工作表可通过撤销恢复，请勿重复添加。");
      sheets.push(sheet);
      ids.add(sheet.id);
    }
    if (action.records.some(r => r.sheetId && !ids.has(r.sheetId)))
      throw Error("无法添加记录：对应的工作表不存在。");
    return {
      ...state,
      sheets,
      records: [...state.records, ...action.records],
      dirty: true,
    };
  }
  if (action.type === "edit") {
    if (!state.fields.some((f) => f.key === action.key)) return state;
    return {
      ...state,
      dirty: true,
      records: state.records.map((r) => {
        if (r.id !== action.id) return r;
        const next = { ...r, [action.key]: C.clean(action.value) };
        // Once a fixed field is cleared, explicit admission is required again.
        return {
          ...next,
          pendingIdentity:
            Boolean(r.pendingIdentity) || C.missingIdentity(next),
        };
      }),
    };
  }
  if (action.type === "delete") {
    const ids = new Set(action.ids);
    const deleted = state.records.flatMap((record, index) =>
      ids.has(record.id) ? [{ record, index }] : [],
    );
    if (!deleted.length) return state;
    return {
      ...state,
      dirty: true,
      records: state.records.filter((r) => !ids.has(r.id)),
      deleted,
      deletedSheet: null,
    };
  }
  if (action.type === "deleteSheet") {
    const index = state.sheets.findIndex(sheet => sheet.id === action.id);
    if (action.id !== UNGROUPED_SHEET && index < 0) return state;
    const belongs = (record: Customer) => action.id === UNGROUPED_SHEET
      ? !record.sheetId : record.sheetId === action.id;
    // Use the whole workspace, including pending and currently hidden records.
    const deleted = state.records.flatMap((record, index) =>
      belongs(record) ? [{ record, index }] : [],
    );
    if (index < 0 && !deleted.length) return state;
    return {
      ...state,
      records: state.records.filter(record => !belongs(record)),
      sheets: state.sheets.filter(sheet => sheet.id !== action.id),
      deleted,
      deletedSheet: index < 0 ? null : { sheet: state.sheets[index], index },
      dirty: true,
    };
  }
  if (action.type === "undoDelete") {
    if (!state.deleted.length && !state.deletedSheet) return state;
    const existing = new Set(state.records.map(r => r.id));
    const deleted = state.deleted.filter(({ record }) => !existing.has(record.id));
    if (state.records.length + deleted.length > MAX_RECORDS)
      throw Error("恢复后将超过 50,000 条上限，请先减少当前记录再撤销。");
    const records = [...state.records];
    const sheets = [...state.sheets];
    if (state.deletedSheet) {
      const { sheet, index } = state.deletedSheet;
      sheets.splice(Math.min(index, sheets.length), 0, sheet);
    }
    for (const { record, index } of deleted) {
      // Restore only fields still present in the current schema, preserving provenance.
      const restored: Customer = {
        id: record.id, source: record.source, sourceRow: record.sourceRow,
        sourceCells: record.sourceCells, pendingIdentity: record.pendingIdentity,
        ...(record.sheetId ? { sheetId: record.sheetId } : {}),
        ...Object.fromEntries(state.fields.map(f => [f.key, record[f.key] ?? ""])),
      };
      records.splice(Math.min(index, records.length), 0, restored);
    }
    return { ...state, records, sheets, deleted: [], deletedSheet: null, dirty: true };
  }
  if (action.type === "admit") {
    const ids = new Set(action.ids);
    return {
      ...state,
      dirty: true,
      records: state.records.map((r) =>
        ids.has(r.id) && C.pendingIdentity(r) && !C.missingIdentity(r)
          ? C.admitRecord(r)
          : r,
      ),
    };
  }
  if (action.type === "addField") {
    const fields = C.addField(state.fields, action.name, action.key),
      added = fields.at(-1)!;
    return {
      ...state,
      fields,
      dirty: state.dirty || !!state.records.length,
      records: state.records.map((r) => ({ ...r, [added.key]: "" })),
    };
  }
  const fields = C.removeField(state.fields, action.key);
  return {
    ...state,
    fields,
    dirty: state.dirty || !!state.records.length,
    records: state.records.map((r) => {
      const next = { ...r };
      delete next[action.key];
      return next;
    }),
  };
}

// Keep selections across pages, but never across a change in category/search membership.
export function visibleSelection(selected: Set<string>, filtered: Customer[]) {
  const allowed = new Set(filtered.map(record => record.id));
  return new Set([...selected].filter(id => allowed.has(id)));
}

export function summarize(records: Customer[]) {
  const accepted = records.filter((r) => !C.pendingIdentity(r));
  const duplicates = C.duplicateIds(accepted);
  const review = new Set(
    accepted.filter((r) => C.issues(r).length).map((r) => r.id),
  );
  return {
    duplicates,
    review,
    counts: {
      all: accepted.length,
      review: review.size,
      duplicate: duplicates.size,
      missing: records.length - accepted.length,
    },
  };
}

export function filterRecords(
  records: Customer[],
  fields: Field[],
  category: Category,
  query: string,
  summary: ReturnType<typeof summarize>,
) {
  const search = query.trim().toLowerCase();
  return records.filter((r) => {
    const pending = C.pendingIdentity(r);
    const matches =
      category === "missing"
        ? pending
        : !pending &&
          (category === "all" ||
            (category === "review" && summary.review.has(r.id)) ||
            (category === "duplicate" && summary.duplicates.has(r.id)));
    return (
      matches &&
      (!search ||
        [
          ...fields.map((f) => r[f.key]),
          r.source,
          ...(r.sourceCells || []).map((c) => c.value),
        ].some((v) =>
          String(v ?? "")
            .toLowerCase()
            .includes(search),
        ))
    );
  });
}
