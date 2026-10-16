import * as C from "../domain/core.ts";
import type { Customer, Field } from "../domain/core.ts";

export type Category = "all" | "review" | "duplicate" | "missing";
export const MAX_RECORDS = 50000;
export interface Workspace {
  fields: Field[];
  records: Customer[];
  dirty: boolean;
}
export const initialWorkspace = (): Workspace => ({
  fields: C.fields.map((f) => ({ ...f })),
  records: [],
  dirty: false,
});
export type Action =
  | { type: "append"; records: Customer[] }
  | { type: "edit"; id: string; key: string; value: string }
  | { type: "delete"; ids: string[] }
  | { type: "admit"; ids: string[] }
  | { type: "addField"; name: string; key: string }
  | { type: "removeField"; key: string }
  | { type: "exported" };

export function reducer(state: Workspace, action: Action): Workspace {
  if (action.type === "exported") return { ...state, dirty: false };
  if (action.type === "append") {
    if (state.records.length + action.records.length > MAX_RECORDS)
      throw Error("当前最多保留 50,000 条记录，请分批整理。");
    return {
      ...state,
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
    return {
      ...state,
      dirty: true,
      records: state.records.filter((r) => !ids.has(r.id)),
    };
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
