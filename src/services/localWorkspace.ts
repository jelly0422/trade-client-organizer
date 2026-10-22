import { fixedKeys, fields as defaultFields } from "../domain/core.ts";
import { MAX_RECORDS } from "../state/workspace.ts";
import type { Workspace } from "../state/workspace.ts";
import { UNGROUPED_SHEET } from "../domain/sheets.ts";

export interface SavedWorkspace {
  version: 1;
  revision: string;
  savedAt: number;
  workspace: Workspace;
}
export class SaveConflict extends Error {
  constructor() { super("另一个页面已保存新的进度。请先导出本页需要保留的修改，再重新打开页面。"); }
}
const object = (value: unknown): value is Record<string, any> =>
  !!value && typeof value === "object" && !Array.isArray(value);
export function decodeWorkspace(value: unknown): SavedWorkspace | null {
  if (value === undefined) return null;
  const fail = () => { throw Error("本地进度的格式无法读取。已有数据未被覆盖，请保留此浏览器的数据并联系维护人员。"); };
  if (!object(value) || value.version !== 1 || typeof value.revision !== "string" ||
      !Number.isFinite(value.savedAt) || !object(value.workspace)) return fail();
  const w = value.workspace;
  const sheets = w.sheets === undefined ? [] : w.sheets;
  if (!Array.isArray(sheets)) return fail();
  const sheetIds = new Set<string>();
  for (const sheet of sheets) {
    if (!object(sheet) || typeof sheet.id !== "string" || !sheet.id ||
        sheet.id === UNGROUPED_SHEET || sheetIds.has(sheet.id) ||
        typeof sheet.name !== "string" || !sheet.name.trim() ||
        typeof sheet.sourceName !== "string") return fail();
    sheetIds.add(sheet.id);
  }
  if (!Array.isArray(w.fields) || w.fields.length < 3 || w.fields.length > 40 ||
      !Array.isArray(w.records) || w.records.length > MAX_RECORDS ||
      !Array.isArray(w.deleted) || w.deleted.length > MAX_RECORDS || typeof w.dirty !== "boolean") return fail();
  const keys = new Set<string>();
  for (const f of w.fields) {
    if (!object(f) || typeof f.key !== "string" || keys.has(f.key) ||
        !(defaultFields.some(field => field.key === f.key) || /^custom_[a-zA-Z0-9_]+$/.test(f.key)) ||
        typeof f.label !== "string" || !f.label.trim() || f.label.length > 30 ||
        !Array.isArray(f.aliases) || !f.aliases.every((a: unknown) => typeof a === "string")) return fail();
    keys.add(f.key);
  }
  if (!fixedKeys.every(key => keys.has(key))) return fail();
  const validateRecords = (records: unknown[]) => {
    const ids = new Set<string>();
    for (const r of records) {
      if (!object(r) || typeof r.id !== "string" || !r.id || ids.has(r.id) ||
          (r.pendingIdentity !== undefined && typeof r.pendingIdentity !== "boolean") ||
          (r.source !== undefined && typeof r.source !== "string") ||
          (r.sourceRow !== undefined && typeof r.sourceRow !== "string" && typeof r.sourceRow !== "number")) return false;
      ids.add(r.id);
      if (r.sheetId !== undefined && (typeof r.sheetId !== "string" || !sheetIds.has(r.sheetId))) return false;
      if ([...keys].some(key => r[key] !== undefined && typeof r[key] !== "string" && typeof r[key] !== "number")) return false;
      if (r.sourceCells !== undefined && (!Array.isArray(r.sourceCells) || !r.sourceCells.every((c: unknown) =>
        object(c) && Number.isInteger(c.column) && c.column > 0 &&
        typeof c.label === "string" && typeof c.value === "string"))) return false;
    }
    return true;
  };
  if (!validateRecords(w.records) || !w.deleted.every((d: unknown) => object(d) &&
      Number.isInteger(d.index) && d.index >= 0 && d.index < MAX_RECORDS) ||
      !validateRecords(w.deleted.map((d: any) => d.record))) return fail();
  // Older saved workspaces have no sheet registry; their records remain intact.
  return { ...value, workspace: { ...w, sheets } } as SavedWorkspace;
}

function openDatabase(factory: IDBFactory) {
  return new Promise<IDBDatabase>((resolve, reject) => {
    let finished = false;
    const request = factory.open("trade-client-organizer", 1);
    const timer = setTimeout(() => {
      finished = true;
      reject(Error("本地存储暂时无法打开，请关闭其他页面后重试。"));
    }, 5000);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains("workspaces"))
        request.result.createObjectStore("workspaces");
    };
    request.onerror = () => { clearTimeout(timer); reject(request.error); };
    request.onsuccess = () => {
      clearTimeout(timer);
      if (finished) { request.result.close(); return; }
      request.result.onversionchange = () => request.result.close();
      resolve(request.result);
    };
  });
}
export async function readLocalWorkspace(factory = globalThis.indexedDB) {
  if (!factory) throw Error("此浏览器无法使用本地存储，请在普通浏览模式下重试。");
  const db = await openDatabase(factory);
  try {
    return await new Promise<SavedWorkspace | null>((resolve, reject) => {
      const transaction = db.transaction("workspaces", "readonly");
      const request = transaction.objectStore("workspaces").get("current");
      transaction.onabort = () => reject(transaction.error || Error("读取本地进度失败。"));
      transaction.oncomplete = () => {
        try { resolve(decodeWorkspace(request.result)); } catch (error) { reject(error); }
      };
    });
  } finally { db.close(); }
}
export async function writeLocalWorkspace(workspace: Workspace, expectedRevision: string | null, factory = globalThis.indexedDB) {
  if (!factory) throw Error("此浏览器无法使用本地存储。");
  const db = await openDatabase(factory);
  try {
    return await new Promise<SavedWorkspace>((resolve, reject) => {
      // A read/compare/write transaction prevents a stale tab from overwriting progress.
      const transaction = db.transaction("workspaces", "readwrite");
      const store = transaction.objectStore("workspaces");
      const request = store.get("current");
      let failure: unknown;
      let saved: SavedWorkspace;
      request.onsuccess = () => {
        try {
          const previous = decodeWorkspace(request.result);
          if ((previous?.revision ?? null) !== expectedRevision) throw new SaveConflict();
          saved = { version: 1, revision: crypto.randomUUID(), savedAt: Date.now(), workspace };
          store.put(saved, "current");
        } catch (error) { failure = error; transaction.abort(); }
      };
      // Request success alone does not establish a committed save.
      transaction.oncomplete = () => resolve(saved);
      transaction.onabort = () => reject(failure || transaction.error || Error("保存未完成，请重试。"));
    });
  } finally { db.close(); }
}

export type SaveState = { status: "saved" | "saving" | "error" | "conflict"; savedAt?: number; error?: string };
// Serialize asynchronous writes and retain edits made while a previous write is pending.
export function createWorkspaceSaver(
  initial: Workspace, revision: string | null,
  write: (workspace: Workspace, revision: string | null) => Promise<SavedWorkspace>,
  notify: (state: SaveState) => void,
) {
  let latest = initial, saved = initial, running = false, blocked = false;
  const flush = async () => {
    if (running || blocked || latest === saved) return;
    running = true;
    notify({ status: "saving" });
    try {
      while (latest !== saved) {
        const snapshot = latest;
        const result = await write(snapshot, revision);
        revision = result.revision;
        saved = snapshot;
        if (latest === saved) notify({ status: "saved", savedAt: result.savedAt });
      }
    } catch (error) {
      blocked = error instanceof SaveConflict;
      notify({ status: blocked ? "conflict" : "error", error: blocked ? (error as Error).message :
        "本地保存失败，可能是存储空间不足或权限受限。修改仍在本页，请重试保存或导出客户表。" });
    } finally { running = false; }
  };
  return {
    update(workspace: Workspace) { latest = workspace; void flush(); },
    retry: flush,
    hasPending: () => latest !== saved,
  };
}
