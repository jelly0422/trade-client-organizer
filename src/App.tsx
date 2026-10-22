import { useCallback, useEffect, useMemo, useState } from "react";
import {
  Alert,
  App as AntApp,
  Button,
  Drawer,
  Modal,
  Select,
  Spin,
  Tag,
  Tooltip,
} from "antd";
import {
  ApartmentOutlined,
  ArrowDownOutlined,
  ArrowRightOutlined,
  CheckCircleOutlined,
  CloudUploadOutlined,
  CopyOutlined,
  DatabaseOutlined,
  DownOutlined,
  DownloadOutlined,
  GlobalOutlined,
  InfoCircleOutlined,
  LockOutlined,
  QuestionCircleOutlined,
  SafetyCertificateOutlined,
  TeamOutlined,
  WarningOutlined,
} from "@ant-design/icons";
import * as C from "./domain/core";
import type { Customer } from "./domain/core";
import { CustomerLibrary } from "./components/CustomerLibrary";
import { ImportDrawer } from "./components/ImportDrawer";
import { customerSheetGroups, recordsInSheet } from "./domain/sheets";
import { useLocalWorkspace } from "./hooks/useLocalWorkspace";
import { useGlassPointer } from "./hooks/useGlassPointer";
import {
  summarize,
  filterRecords,
  MAX_RECORDS,
  visibleSelection,
  type Category,
} from "./state/workspace";
import {
  demoSource,
  exportCustomers,
  importWorkbook,
  parseBook,
  readFile,
  synchronizeSchema,
  uid,
  withConfig,
  type FileSource,
  type SheetConfig,
} from "./services/workbooks";

const categories = [
  {
    key: "all" as const,
    label: "全部客户",
    note: "主体完整，可继续整理",
    icon: <TeamOutlined />,
  },
  {
    key: "review" as const,
    label: "待完善",
    note: "邮箱格式需要核对",
    icon: <WarningOutlined />,
  },
  {
    key: "duplicate" as const,
    label: "疑似重复",
    note: "可能是同一位客户",
    icon: <CopyOutlined />,
  },
  {
    key: "missing" as const,
    label: "缺少主体",
    note: "补齐固定字段后加入",
    icon: <InfoCircleOutlined />,
  },
];
export default function App() {
  useGlassPointer();
  const { message, modal } = AntApp.useApp();
  const [libraryNavOpen, setLibraryNavOpen] = useState(true);
  const { workspace, dispatch, ready, loadError, saveState, retrySave } = useLocalWorkspace();
  const [sources, setSources] = useState<FileSource[]>([]),
    [current, setCurrent] = useState("");
  const [imported, setImported] = useState(new Set<string>()),
    [busy, setBusy] = useState(false);
  const [importOpen, setImportOpen] = useState(false),
    [importPanel, setImportPanel] = useState<"mapping" | "fields">("mapping"),
    [helpOpen, setHelpOpen] = useState(false);
  const [sheetSelection, setSheetSelection] = useState("");
  const openImport = (panel: "mapping" | "fields" = "mapping") => {
    setImportPanel(panel);
    setImportOpen(true);
  };
  const [category, setCategory] = useState<Category>("all"),
    [query, setQuery] = useState("");
  const [page, setPage] = useState(1),
    [pageSize, setPageSize] = useState(25),
    [selected, setSelected] = useState(new Set<string>());
  const [original, setOriginal] = useState<Customer | null>(null),
    [format, setFormat] = useState<"xlsx" | "csv">("xlsx");
  const [originalTarget, setOriginalTarget] = useState<string>();
  const sheetGroups = useMemo(() => customerSheetGroups(workspace.records, workspace.sheets), [workspace.records, workspace.sheets]);
  const multipleSheets = sheetGroups.length > 1;
  useEffect(() => { if (multipleSheets) setFormat("xlsx"); }, [multipleSheets]);
  useEffect(() => {
    if (sheetSelection && !sheetGroups.some(s => s.id === sheetSelection)) setSheetSelection("");
  }, [sheetGroups, sheetSelection]);
  const summary = useMemo(
    () => summarize(workspace.records),
    [workspace.records],
  );
  const filtered = useMemo(
    () =>
      filterRecords(
        recordsInSheet(workspace.records, sheetSelection),
        workspace.fields,
        category,
        query,
        summary,
      ),
    [workspace.records, workspace.fields, category, query, summary, sheetSelection],
  );
  useEffect(() => {
    setSources((previous) => synchronizeSchema(previous, workspace.fields));
  }, [workspace.fields]);
  const activeSelection = useMemo(() => visibleSelection(selected, filtered), [selected, filtered]);
  useEffect(() => {
    setSelected(previous => {
      const next = visibleSelection(previous, filtered);
      return next.size === previous.size ? previous : next;
    });
  }, [filtered]);
  const changeCategory = (value: Category) => {
    setCategory(value);
    setPage(1);
    setSelected(new Set());
  };
  const updateSource = (change: (source: FileSource) => FileSource) =>
    setSources((previous) =>
      previous.map((s) => (s.id === current ? change(s) : s)),
    );
  const onFiles = async (files: File[]) => {
    if (busy) return;
    setBusy(true);
    openImport();
    const added: FileSource[] = [];
    for (const file of files) {
      try {
        added.push(await readFile(file, workspace.fields));
      } catch (e) {
        message.error((e as Error).message);
      }
    }
    if (added.length) {
      setSources((previous) => [...previous, ...added]);
      setCurrent(added.at(-1)!.id);
      message.success(`已读取 ${added.length} 份文件，请核对映射。`);
    }
    setBusy(false);
  };
  const onDemo = () => {
    try {
      const existing = sources.find((s) => s.demo);
      if (existing) setCurrent(existing.id);
      else {
        const source = demoSource(workspace.fields);
        setSources((previous) => [...previous, source]);
        setCurrent(source.id);
      }
      openImport();
    } catch (e) {
      message.error((e as Error).message);
    }
  };
  const onApply = () => {
    const source = sources.find((s) => s.id === current);
    if (!source) return;
    const key = source.id;
    if (imported.has(key)) return;
    let batch: ReturnType<typeof importWorkbook>;
    try { batch = importWorkbook(source, workspace.fields); }
    catch (error) { message.error((error as Error).message); return; }
    const { records, sheets } = batch;
    if (records.length + workspace.records.length > MAX_RECORDS) {
      message.error("当前最多保留 50,000 条记录，请分批整理。");
      return;
    }
    const missing = records.filter(C.pendingIdentity).length;
    dispatch({ type: "append", records, sheets });
    setImported((previous) => new Set(previous).add(key));
    setQuery("");
    setSheetSelection("");
    changeCategory(records.length && missing === records.length ? "missing" : "all");
    setImportOpen(false);
    message.success(
      `已保留 ${records.length} 条记录：全部客户 ${records.length - missing} 条，缺少主体 ${missing} 条。`,
    );
  };
  const onAdd = () => {
    if (workspace.records.length >= MAX_RECORDS) {
      message.warning("已达到 50,000 条记录上限，请分批整理。");
      return;
    }
    const targetSheet = workspace.sheets.find(s => s.id === sheetSelection) ||
      (!sheetSelection && sheetGroups.length === 1 ? workspace.sheets.find(s => s.id === sheetGroups[0].id) : undefined);
    const record: Customer = {
      id: uid(),
      ...Object.fromEntries(workspace.fields.map((f) => [f.key, ""])),
      source: "手动新增",
      sourceRow: "",
      pendingIdentity: true,
      sourceCells: [],
      ...(targetSheet ? { sheetId: targetSheet.id } : {}),
    };
    dispatch({ type: "append", records: [record] });
    changeCategory("missing");
    setQuery("");
    setPage(Math.ceil((recordsInSheet(workspace.records, sheetSelection).filter(C.pendingIdentity).length + 1) / pageSize));
    message.info("已新增记录，请补齐客户名称、公司名和邮箱后勾选加入。");
  };
  const onEdit = useCallback((id: string, key: string, value: string) => {
    dispatch({ type: "edit", id, key, value });
    if (C.fixedKeys.includes(key) && !C.clean(value))
      setSelected((previous) => {
        const next = new Set(previous);
        next.delete(id);
        return next;
      });
  }, []);
  const onDelete = () =>
    modal.confirm({
      title: `删除选中的 ${activeSelection.size} 条客户？`,
      content: "删除后可点击“撤销上次删除”恢复本次记录；再次删除会替换上一次的撤销记录。",
      okText: "删除客户",
      cancelText: "取消",
      okButtonProps: { danger: true },
      onOk: () => {
        dispatch({ type: "delete", ids: [...activeSelection] });
        setSelected(new Set());
        message.success("已删除所选客户，可撤销上次删除。");
      },
    });
  const onDeleteSheet = () => {
    const sheet = sheetGroups.find(s => s.id === sheetSelection);
    if (!sheet) return;
    modal.confirm({
      title: `删除工作表“${sheet.name}”？`,
      content: <>
        <p>来源：{sheet.sourceName}</p>
        <p>将删除这张工作表及其中全部 {sheet.records.length} 条客户（含缺少主体记录），不受搜索、分类或分页影响。</p>
        <p>可通过“撤销上次删除”恢复工作表和数据；再次删除会替换本次撤销记录。电脑上的原文件会保留。</p>
      </>,
      okText: "删除工作表",
      cancelText: "取消",
      okButtonProps: { danger: true },
      onOk: () => {
        dispatch({ type: "deleteSheet", id: sheet.id });
        setSelected(new Set());
        setSheetSelection("");
        setQuery("");
        changeCategory("all");
        message.success(`已删除“${sheet.name}”及 ${sheet.records.length} 条客户，可撤销。`);
      },
    });
  };
  const onAdmit = () => {
    const ready = workspace.records
      .filter(
        (r) =>
          activeSelection.has(r.id) && C.pendingIdentity(r) && !C.missingIdentity(r),
      )
      .map((r) => r.id);
    if (!ready.length) return;
    dispatch({ type: "admit", ids: ready });
    setSelected(
      (previous) => new Set([...previous].filter((id) => !ready.includes(id))),
    );
    message.success(
      `已将 ${ready.length} 条记录加入全部客户，未补齐的记录继续保留。`,
    );
  };
  const onExport = (template = false) => {
    try {
      exportCustomers(workspace.records, workspace.fields, format, template, workspace.sheets);
      if (!template) dispatch({ type: "exported" });
      message.success(
        template
          ? "标准模板已生成。"
          : `已生成全部 ${workspace.records.length} 条记录，请确认文件已保存。`,
      );
    } catch (e) {
      message.error("导出失败：" + (e as Error).message);
    }
  };
  const onUndoDelete = () => {
    if (workspace.records.length + workspace.deleted.length > MAX_RECORDS) {
      message.warning("恢复后将超过 50,000 条上限，请先减少当前记录。");
      return;
    }
    dispatch({ type: "undoDelete" });
    setQuery("");
    setSheetSelection(workspace.deletedSheet?.sheet.id || "");
    changeCategory(workspace.deleted.length && workspace.deleted.every(({ record }) => C.pendingIdentity(record)) ? "missing" : "all");
    message.success(workspace.deletedSheet
      ? `已恢复工作表“${workspace.deletedSheet.sheet.name}”及 ${workspace.deleted.length} 条客户。`
      : `已恢复 ${workspace.deleted.length} 条客户及其原始内容。`);
  };
  const onFillOriginal = (value: string) => {
    const record = workspace.records.find(r => r.id === original?.id);
    const field = workspace.fields.find(f => f.key === originalTarget);
    if (!record || !field) return;
    const apply = () => {
      onEdit(record.id, field.key, value);
      message.success(`已填入“${field.label}”。`);
    };
    if (C.clean(record[field.key]) && C.clean(record[field.key]) !== C.clean(value))
      modal.confirm({ title: `覆盖“${field.label}”中的已有内容？`, content: `当前内容：${C.clean(record[field.key])}`,
        okText: "确认填入", cancelText: "取消", onOk: apply });
    else apply();
  };
  if (!ready) return (
    <div className="workspace-loading">
      {loadError ? <Alert type="error" showIcon title="暂时无法恢复本地进度" description={loadError}
        action={<Button onClick={() => window.location.reload()}>重新读取</Button>} /> :
        <><Spin size="large" /><p>正在恢复本地整理进度…</p></>}
    </div>
  );
  return (
    <div className="workspace-shell">
      <div className="workspace-landscape" aria-hidden="true">
        <div className="landscape-ring landscape-ring-one" />
        <div className="landscape-ring landscape-ring-two" />
        <div className="landscape-ring landscape-ring-three" />
        <div className="landscape-mist" />
      </div>
      <aside className="sidebar">
        <a className="brand" href="#customer-library">
          <span className="brand-mark">
            <GlobalOutlined />
          </span>
          <span>
            外贸客户工具<small>CLIENT WORKSPACE</small>
          </span>
        </a>
        <span className="nav-label">工作空间</span>
        <nav aria-label="主导航">
          <ul className="nav-list">
            <li>
              <button
                className="nav-item active nav-parent"
                aria-label="客户资料库"
                title="客户资料库"
                aria-expanded={libraryNavOpen}
                aria-controls="customer-library-submenu"
                onClick={() => {
                  setLibraryNavOpen(open => !open);
                  document
                    .getElementById("customer-library")
                    ?.scrollIntoView({ behavior: "smooth" });
                }}
              >
                <DatabaseOutlined />
                客户资料库
                <DownOutlined className="nav-chevron" />
              </button>
              <ul
                id="customer-library-submenu"
                className="nav-submenu"
                aria-label="客户资料库子菜单"
                hidden={!libraryNavOpen}
              >
                <li>
                  <button
                    className={`nav-item nav-subitem${importOpen ? " active" : ""}`}
                    aria-label="导入与映射"
                    title="导入与映射"
                    onClick={() => openImport()}
                  >
                    <CloudUploadOutlined />
                    导入与映射
                    {sources.length > 0 && (
                      <span className="nav-count">{sources.length}</span>
                    )}
                  </button>
                </li>
              </ul>
            </li>
          </ul>
        </nav>
        <div className="sidebar-bottom">
          <div className="privacy-card">
            <SafetyCertificateOutlined />
            <strong>资料，留在你的浏览器</strong>
            <p>
              文件仅在本地处理。
              <br />
              进度自动保存在此浏览器。
            </p>
            <span>
              <i />
              本地工作模式
            </span>
          </div>
          <button className="nav-item" onClick={() => setHelpOpen(true)}>
            <QuestionCircleOutlined />
            使用说明
            <ArrowRightOutlined />
          </button>
          <div className="sidebar-foot">让每一次客户整理更从容</div>
        </div>
      </aside>
      <main className="main-workspace">
        <header className="topbar">
          <div className="breadcrumb">
            <span>工作空间</span>
            <span>/</span>
            <b>客户资料库</b>
          </div>
          <div className="topbar-right">
            <Button className="template-button" icon={<DownloadOutlined />} onClick={() => onExport(true)}>
              下载标准模板
            </Button>
            <span className="local-badge">
              <span />
              浏览器本地处理
            </span>
            <Tooltip title="使用说明">
              <Button
                type="text"
                aria-label="打开使用说明"
                icon={<QuestionCircleOutlined />}
                onClick={() => setHelpOpen(true)}
              />
            </Tooltip>
          </div>
        </header>
        <div className="page-content">
          <section className="hero-banner">
            <div className="hero-copy">
              <span className="hero-eyebrow">
                <span />
                A CLEARER VIEW OF YOUR CLIENTS
              </span>
              <h1>外贸客户信息整理工具</h1>
              <p>归集每一份资料，让下一次合作更清晰。</p>
            </div>
            <div className="glass-hero-art" aria-hidden="true">
              <div className="glass-sheet" />
              <div className="glass-sheet glass-sheet-front">
                {[1, 2, 3, 4, 5, 6].map(i => <i key={i} />)}
              </div>
              <div className="glass-sheet-check"><CheckCircleOutlined /></div>
            </div>
          </section>
          <div className="workspace-overline">
            <span>
              <ApartmentOutlined />
              客户概览{workspace.sheets.length > 0 ? " · 全部工作表" : ""}
            </span>
            <span role="status" data-testid="local-save-status">
              {saveState.status !== "saved" ? (
                <>
                  <span className="unsaved-dot" />
                  {saveState.status === "saving" ? "正在保存到本地…" : "本地进度尚未保存"}
                </>
              ) : (
                <>
                  <CheckCircleOutlined />{" "}
                  {saveState.savedAt ? `已保存到本地 · ${new Date(saveState.savedAt).toLocaleTimeString("zh-CN")}` : "本地自动保存已开启"}
                </>
              )}
            </span>
          </div>
          {(saveState.status === "error" || saveState.status === "conflict") && <Alert
            className="save-warning" type="warning" showIcon title={saveState.error}
            action={saveState.status === "error" ? <Button onClick={retrySave}>重试保存</Button> : undefined} />}
          <div className="category-grid" role="group" aria-label="客户分类">
            {categories.map((item) => (
              <button
                key={item.key}
                className={`category-card ${item.key} ${category === item.key ? "selected" : ""}`}
                aria-pressed={category === item.key}
                data-testid={"category-" + item.key}
                onClick={() => changeCategory(item.key)}
              >
                <div className="category-top">
                  <span className="category-label">{item.label}</span>
                  <span className="category-icon">{item.icon}</span>
                </div>
                <div
                  className="category-number"
                  data-testid={"count-" + item.key}
                >
                  {summary.counts[item.key].toLocaleString()}
                </div>
                <div className="category-note"><span>{item.note}</span><ArrowRightOutlined /></div>
              </button>
            ))}
          </div>
          {summary.counts.missing > 0 && category !== "missing" && (
            <Alert
              className="missing-banner"
              type="info"
              showIcon
              title={`${summary.counts.missing} 条记录缺少主体或等待确认，已单独保留。`}
              action={
                <Button
                  type="link"
                  size="small"
                  onClick={() => { setSheetSelection(""); changeCategory("missing"); }}
                >
                  前往补充 <ArrowRightOutlined />
                </Button>
              }
            />
          )}
          <CustomerLibrary
            fields={workspace.fields}
            records={workspace.records}
            sheetGroups={sheetGroups}
            sheetSelection={sheetSelection}
            onSheetSelection={(value) => { setSheetSelection(value); setPage(1); setSelected(new Set()); }}
            onDeleteSheet={onDeleteSheet}
            multipleSheets={multipleSheets}
            filtered={filtered}
            category={category}
            summary={summary}
            query={query}
            onQuery={(value) => {
              setQuery(value);
              setPage(1);
              setSelected(new Set());
            }}
            page={page}
            pageSize={pageSize}
            onPage={(next, size) => {
              setPage(size !== pageSize ? 1 : next);
              setPageSize(size);
            }}
            selected={activeSelection}
            onSelected={setSelected}
            onEdit={onEdit}
            onAdd={onAdd}
            onDelete={onDelete}
            deletedCount={workspace.deleted.length}
            deletedSheetName={workspace.deletedSheet?.sheet.name}
            onUndoDelete={onUndoDelete}
            onAdmit={onAdmit}
            onImport={() => openImport()}
            onDemo={onDemo}
            onFields={() => openImport("fields")}
            onOriginal={(record) => { setOriginal(record); setOriginalTarget(undefined); }}
            format={format}
            onFormat={setFormat}
            onExport={() => onExport()}
          />
          <footer className="page-footer">
            <span>外贸客户信息整理工具</span>
            <span>
              <LockOutlined />
              本地自动保存 · 仅限当前设备、浏览器与网址
            </span>
          </footer>
        </div>
      </main>
      <ImportDrawer
        open={importOpen}
        onClose={() => setImportOpen(false)}
        fields={workspace.fields}
        customers={workspace.records}
        panel={importPanel}
        onPanel={setImportPanel}
        onMerge={(mergeSheets) => updateSource(s => ({ ...s, mergeSheets }))}
        sources={sources}
        current={current}
        imported={imported}
        busy={busy}
        onFiles={onFiles}
        onSelect={setCurrent}
        onSheet={(sheet) =>
          updateSource((s) => withConfig({ ...s, sheet }, workspace.fields))
        }
        onConfig={(config: SheetConfig) =>
          updateSource((s) => ({
            ...s,
            configs: { ...s.configs, [s.sheet]: config },
          }))
        }
        onEncoding={(encoding) => {
          const source = sources.find((s) => s.id === current);
          if (!source?.bytes) return;
          try {
            const book = parseBook(source.bytes, source.name, encoding);
            updateSource((s) =>
              withConfig(
                {
                  ...s,
                  book,
                  encoding,
                  sheet: book.SheetNames[0],
                  configs: {},
                },
                workspace.fields,
              ),
            );
          } catch {
            message.error("无法使用此编码解析文件。");
          }
        }}
        onApply={onApply}
        onDemo={onDemo}
        onAddField={(name) => {
          const key = "custom_" + uid().replaceAll("-", "");
          C.addField(workspace.fields, name, key);
          dispatch({ type: "addField", name, key });
          message.success("已添加字段。");
        }}
        onRemoveField={(key) => {
          dispatch({ type: "removeField", key });
          message.success("已删除字段。");
        }}
      />
      <Drawer
        title="原始内容"
        open={!!original}
        onClose={() => setOriginal(null)}
        size={520}
      >
        <p className="muted-note">
          {original?.source}
          {original?.sourceRow ? ` · 第 ${original.sourceRow} 行` : ""}
        </p>
        <div className="original-fill-toolbar">
          <Select aria-label="原始内容目标字段" placeholder="选择要补充的字段" value={originalTarget}
            onChange={setOriginalTarget} options={workspace.fields.map(f => ({ value: f.key, label: f.label }))} />
          <p>选择字段后，可填入任意原始值；覆盖已有内容前会先确认。</p>
        </div>
        <div className="original-content">
          {original?.sourceCells?.map((cell) => (
            <div key={cell.column}>
              <label>
                {cell.label}
                <span>第 {cell.column} 列</span>
              </label>
              <p>{cell.value}</p>
              <Button size="small" disabled={!originalTarget} onClick={() => onFillOriginal(cell.value)}>填入所选字段</Button>
            </div>
          ))}
        </div>
      </Drawer>
      <Modal
        title="开始整理你的客户资料"
        open={helpOpen}
        onCancel={() => setHelpOpen(false)}
        footer={
          <Button type="primary" onClick={() => setHelpOpen(false)}>
            知道了
          </Button>
        }
        width={620}
      >
        <div className="help-steps">
          <div>
            <b>01</b>
            <section>
              <h3>导入文件，核对字段</h3>
              <p>
                支持
                XLSX、XLS、CSV、TSV、ODS。字段设置已整合在导入窗口内。多张工作表先选择合并或保留独立表，再逐表核对映射，一次添加整个文件。
              </p>
            </section>
          </div>
          <ArrowDownOutlined />
          <div>
            <b>02</b>
            <section>
              <h3>补充主体，校对资料</h3>
              <p>
                公司名、客户名称、邮箱缺少任一项时单独保留。补齐后勾选并点击“加入全部客户”；单击表格编辑，点击表格外保存。各分类会直接展示判定规则。
              </p>
            </section>
          </div>
          <ArrowDownOutlined />
          <div>
            <b>03</b>
            <section>
              <h3>导出全部记录</h3>
              <p>
                导出不受工作表筛选、搜索和分页影响。选择保留工作表时，Excel 按原表分开保存，缺少主体的记录保留在所属表并标注分类；空表也保留，同名表自动加序号。选择具体工作表后，可删除该表及全部记录，支持撤销最近一次删除（包括空表）。原始内容另附一张表。多个工作表须使用 Excel 导出。客户、工作表归属、字段配置与最近一次删除记录会自动保存到当前浏览器。
              </p>
            </section>
          </div>
        </div>
        <Alert
          type="info"
          showIcon
          title="本地保存仅限当前设备、浏览器与网址，不跨设备同步。清理网站数据或结束无痕会话可能移除进度。尚未加入资料库的导入文件需重新选择。"
        />
      </Modal>
    </div>
  );
}
