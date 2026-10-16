import { useCallback, useEffect, useMemo, useReducer, useState } from "react";
import {
  Alert,
  App as AntApp,
  Button,
  Drawer,
  Modal,
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
  DownloadOutlined,
  FileExcelOutlined,
  GlobalOutlined,
  InfoCircleOutlined,
  LockOutlined,
  QuestionCircleOutlined,
  SafetyCertificateOutlined,
  SettingOutlined,
  TeamOutlined,
  WarningOutlined,
} from "@ant-design/icons";
import * as C from "./domain/core";
import type { Customer } from "./domain/core";
import { CustomerLibrary } from "./components/CustomerLibrary";
import { ImportDrawer } from "./components/ImportDrawer";
import { FieldsDrawer } from "./components/FieldsDrawer";
import {
  initialWorkspace,
  reducer,
  summarize,
  filterRecords,
  MAX_RECORDS,
  type Category,
} from "./state/workspace";
import {
  demoSource,
  exportCustomers,
  mappedSource,
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
  const { message, modal } = AntApp.useApp();
  const [workspace, dispatch] = useReducer(
    reducer,
    undefined,
    initialWorkspace,
  );
  const [sources, setSources] = useState<FileSource[]>([]),
    [current, setCurrent] = useState("");
  const [imported, setImported] = useState(new Set<string>()),
    [busy, setBusy] = useState(false);
  const [importOpen, setImportOpen] = useState(false),
    [fieldsOpen, setFieldsOpen] = useState(false),
    [helpOpen, setHelpOpen] = useState(false);
  const [category, setCategory] = useState<Category>("all"),
    [query, setQuery] = useState("");
  const [page, setPage] = useState(1),
    [pageSize, setPageSize] = useState(25),
    [selected, setSelected] = useState(new Set<string>());
  const [original, setOriginal] = useState<Customer | null>(null),
    [format, setFormat] = useState<"xlsx" | "csv">("xlsx");
  const summary = useMemo(
    () => summarize(workspace.records),
    [workspace.records],
  );
  const filtered = useMemo(
    () =>
      filterRecords(
        workspace.records,
        workspace.fields,
        category,
        query,
        summary,
      ),
    [workspace.records, workspace.fields, category, query, summary],
  );
  useEffect(() => {
    setSources((previous) => synchronizeSchema(previous, workspace.fields));
  }, [workspace.fields]);
  useEffect(() => {
    const onLeave = (event: BeforeUnloadEvent) => {
      if (workspace.dirty && workspace.records.length) {
        event.preventDefault();
        event.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", onLeave);
    return () => window.removeEventListener("beforeunload", onLeave);
  }, [workspace.dirty, workspace.records.length]);
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
    setImportOpen(true);
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
      setImportOpen(true);
    } catch (e) {
      message.error((e as Error).message);
    }
  };
  const onApply = () => {
    const source = sources.find((s) => s.id === current);
    if (!source) return;
    const key = source.id + ":" + source.sheet;
    if (imported.has(key)) return;
    const records = mappedSource(source, workspace.fields).map((r) => ({
      ...r,
      id: uid(),
    }));
    if (records.length + workspace.records.length > MAX_RECORDS) {
      message.error("当前最多保留 50,000 条记录，请分批整理。");
      return;
    }
    const missing = records.filter(C.pendingIdentity).length;
    dispatch({ type: "append", records });
    setImported((previous) => new Set(previous).add(key));
    setQuery("");
    changeCategory(missing === records.length ? "missing" : "all");
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
    const record: Customer = {
      id: uid(),
      ...Object.fromEntries(workspace.fields.map((f) => [f.key, ""])),
      source: "手动新增",
      sourceRow: "",
      pendingIdentity: true,
      sourceCells: [],
    };
    dispatch({ type: "append", records: [record] });
    changeCategory("missing");
    setQuery("");
    setPage(Math.ceil((summary.counts.missing + 1) / pageSize));
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
      title: `删除选中的 ${selected.size} 条客户？`,
      content: "删除后无法撤销，请确认已经保留需要的资料。",
      okText: "删除客户",
      cancelText: "取消",
      okButtonProps: { danger: true },
      onOk: () => {
        dispatch({ type: "delete", ids: [...selected] });
        setSelected(new Set());
        message.success("已删除所选客户。");
      },
    });
  const onAdmit = () => {
    const ready = workspace.records
      .filter(
        (r) =>
          selected.has(r.id) && C.pendingIdentity(r) && !C.missingIdentity(r),
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
      exportCustomers(workspace.records, workspace.fields, format, template);
      if (!template) dispatch({ type: "exported" });
      message.success(
        template
          ? "标准模板已生成。"
          : `已生成全部 ${workspace.records.length} 条记录，请确认文件已保存。`,
      );
    } catch (e) {
      message.error("导出失败，请检查过长单元格，或尝试 CSV 格式。");
    }
  };
  return (
    <div className="workspace-shell">
      <aside className="sidebar">
        <a className="brand" href="#customer-library">
          <span className="brand-mark">
            <GlobalOutlined />
          </span>
          <span>
            客户整理工作台<small>TRADE CLIENT STUDIO</small>
          </span>
        </a>
        <span className="nav-label">工作空间</span>
        <nav aria-label="主导航">
          <button
            className="nav-item active"
            onClick={() =>
              document
                .getElementById("customer-library")
                ?.scrollIntoView({ behavior: "smooth" })
            }
          >
            <DatabaseOutlined />
            客户资料库
            <span className="nav-dot" />
          </button>
          <button className="nav-item" onClick={() => setImportOpen(true)}>
            <CloudUploadOutlined />
            导入与映射
            {sources.length > 0 && (
              <span className="nav-count">{sources.length}</span>
            )}
          </button>
          <button className="nav-item" onClick={() => setFieldsOpen(true)}>
            <SettingOutlined />
            字段设置
          </button>
        </nav>
        <div className="sidebar-bottom">
          <div className="privacy-card">
            <SafetyCertificateOutlined />
            <strong>资料，留在你的浏览器</strong>
            <p>
              文件仅在本地处理。
              <br />
              整理完成后，记得导出保存。
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
            <b>客户资料整理</b>
          </div>
          <div className="topbar-right">
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
            <span className="workspace-avatar">外贸</span>
          </div>
        </header>
        <div className="page-content">
          <section className="hero-banner">
            <div className="hero-copy">
              <span className="hero-eyebrow">
                <span />
                从一份名单，到清晰的客户资料
              </span>
              <h1>外贸客户信息整理工具</h1>
              <p>统一不同格式的表格，让映射、校对与客户管理井然有序。</p>
              <div className="hero-actions">
                <Button
                  type="primary"
                  size="large"
                  icon={<CloudUploadOutlined />}
                  onClick={() => setImportOpen(true)}
                >
                  导入客户表
                </Button>
                <Button
                  size="large"
                  icon={<DownloadOutlined />}
                  onClick={() => onExport(true)}
                >
                  下载标准模板
                </Button>
              </div>
            </div>
            <div className="hero-visual" aria-hidden="true">
              <div className="visual-orbit" />
              <div className="visual-file">
                <FileExcelOutlined />
                <span>多种格式</span>
                <small>XLSX / CSV / ODS</small>
              </div>
              <ArrowRightOutlined className="visual-arrow" />
              <div className="visual-table">
                <div>
                  <span />
                  <span />
                  <span />
                  <b>统一客户资料</b>
                </div>
                <section>
                  {[1, 2, 3].map((i) => (
                    <p key={i}>
                      <i />
                      <span />
                      <span />
                      <CheckCircleOutlined />
                    </p>
                  ))}
                </section>
                <footer>
                  <CheckCircleOutlined />
                  整齐，有据可查
                </footer>
              </div>
            </div>
          </section>
          <div className="workspace-overline">
            <span>
              <ApartmentOutlined />
              你的客户工作台
            </span>
            <span>
              {workspace.dirty ? (
                <>
                  <span className="unsaved-dot" />
                  有未导出的修改
                </>
              ) : (
                <>
                  <CheckCircleOutlined />{" "}
                  {workspace.records.length ? "已生成导出文件" : "准备开始整理"}
                </>
              )}
            </span>
          </div>
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
                  <span className="category-icon">{item.icon}</span>
                  <ArrowRightOutlined />
                </div>
                <div
                  className="category-number"
                  data-testid={"count-" + item.key}
                >
                  {summary.counts[item.key].toLocaleString()}
                </div>
                <div className="category-label">{item.label}</div>
                <div className="category-note">{item.note}</div>
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
                  onClick={() => changeCategory("missing")}
                >
                  前往补充 <ArrowRightOutlined />
                </Button>
              }
            />
          )}
          <CustomerLibrary
            fields={workspace.fields}
            records={workspace.records}
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
            selected={selected}
            onSelected={setSelected}
            onEdit={onEdit}
            onAdd={onAdd}
            onDelete={onDelete}
            onAdmit={onAdmit}
            onImport={() => setImportOpen(true)}
            onDemo={onDemo}
            onFields={() => setFieldsOpen(true)}
            onOriginal={setOriginal}
            format={format}
            onFormat={setFormat}
            onExport={() => onExport()}
          />
          <footer className="page-footer">
            <span>外贸客户信息整理工具</span>
            <span>
              <LockOutlined />
              刷新或关闭页面前，请导出资料
            </span>
          </footer>
        </div>
      </main>
      <ImportDrawer
        open={importOpen}
        onClose={() => setImportOpen(false)}
        fields={workspace.fields}
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
      />
      <FieldsDrawer
        open={fieldsOpen}
        onClose={() => setFieldsOpen(false)}
        fields={workspace.fields}
        records={workspace.records}
        onAdd={(name) => {
          const key = "custom_" + uid().replaceAll("-", "");
          C.addField(workspace.fields, name, key);
          dispatch({ type: "addField", name, key });
          message.success("已添加字段。");
        }}
        onRemove={(key) => {
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
        <div className="original-content">
          {original?.sourceCells?.map((cell) => (
            <div key={cell.column}>
              <label>
                {cell.label}
                <span>第 {cell.column} 列</span>
              </label>
              <p>{cell.value}</p>
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
                XLSX、XLS、CSV、TSV、ODS。选择工作表与表头行，核对自动映射；多张工作表分别添加。
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
                导出不受搜索和分页影响。Excel
                分表保存客户、缺少主体和原始内容；CSV
                保留分类与原始内容。数据与字段设置仅在本次会话有效，导出文件不能完整恢复工作项目状态。
              </p>
            </section>
          </div>
        </div>
        <Alert
          type="info"
          showIcon
          title="文件在浏览器内处理，刷新或关闭前请导出。"
        />
      </Modal>
    </div>
  );
}
