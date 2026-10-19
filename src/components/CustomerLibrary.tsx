import { useEffect, useMemo, useRef, useState } from "react";
import {
  Button,
  Checkbox,
  Empty,
  Input,
  Pagination,
  Select,
  Tag,
  Tooltip,
} from "antd";
import {
  CheckCircleOutlined,
  CloudUploadOutlined,
  DeleteOutlined,
  DownloadOutlined,
  FileSearchOutlined,
  PlusOutlined,
  SearchOutlined,
  SettingOutlined,
  WarningOutlined,
  UndoOutlined,
} from "@ant-design/icons";
import { AgGridReact } from "ag-grid-react";
import {
  AllCommunityModule,
  ModuleRegistry,
  themeQuartz,
  type ColDef,
  type ICellRendererParams,
} from "ag-grid-community";
import * as C from "../domain/core";
import type { Customer, Field } from "../domain/core";
import type { Category, summarize } from "../state/workspace";
import { RulesPanel } from "./RulesPanel";

ModuleRegistry.registerModules([AllCommunityModule]);
const gridTheme = themeQuartz.withParams({
  accentColor: "#0588ad",
  backgroundColor: "#ffffff",
  foregroundColor: "#253c51",
  borderColor: "#e5eef3",
  headerBackgroundColor: "#f5f9fb",
  headerTextColor: "#607588",
  fontFamily: 'Inter, "Segoe UI", "Microsoft YaHei", sans-serif',
  fontSize: 13,
  rowHeight: 58,
  headerHeight: 46,
  wrapperBorderRadius: 0,
  oddRowBackgroundColor: "#fbfdfe",
  selectedRowBackgroundColor: "#e5f7fb",
});
interface Props {
  fields: Field[];
  records: Customer[];
  filtered: Customer[];
  category: Category;
  summary: ReturnType<typeof summarize>;
  query: string;
  onQuery: (value: string) => void;
  page: number;
  pageSize: number;
  onPage: (page: number, size: number) => void;
  selected: Set<string>;
  onSelected: (ids: Set<string>) => void;
  onEdit: (id: string, key: string, value: string) => void;
  onAdd: () => void;
  onDelete: () => void;
  deletedCount: number;
  onUndoDelete: () => void;
  onAdmit: () => void;
  onImport: () => void;
  onDemo: () => void;
  onFields: () => void;
  onOriginal: (record: Customer) => void;
  format: "xlsx" | "csv";
  onFormat: (value: "xlsx" | "csv") => void;
  onExport: () => void;
}
export function CustomerLibrary(p: Props) {
  const grid = useRef<AgGridReact<Customer>>(null);
  const [compact, setCompact] = useState(
    () => window.matchMedia("(max-width:680px)").matches,
  );
  useEffect(() => {
    const media = window.matchMedia("(max-width:680px)");
    const update = () => setCompact(media.matches);
    media.addEventListener("change", update);
    return () => media.removeEventListener("change", update);
  }, []);
  useEffect(() => {
    grid.current?.api?.ensureColumnVisible("company", "start");
  }, [p.category, p.query]);
  const page = Math.min(
    p.page,
    Math.max(1, Math.ceil(p.filtered.length / p.pageSize)),
  );
  const visible = useMemo(
    () => p.filtered.slice((page - 1) * p.pageSize, page * p.pageSize),
    [p.filtered, page, p.pageSize],
  );
  const ready = p.records.filter(
    (r) =>
      p.selected.has(r.id) && C.pendingIdentity(r) && !C.missingIdentity(r),
  ).length;
  const columns = useMemo<ColDef<Customer>[]>(
    () => [
      {
        colId: "selection",
        width: 50,
        minWidth: 50,
        maxWidth: 50,
        pinned: "left",
        editable: false,
        resizable: false,
        headerComponent: () => {
          const count = visible.filter((r) => p.selected.has(r.id)).length;
          return (
            <Checkbox
              aria-label="选择当前页全部记录"
              checked={!!visible.length && count === visible.length}
              indeterminate={count > 0 && count < visible.length}
              onChange={(e) => {
                const ids = new Set(p.selected);
                visible.forEach((r) =>
                  e.target.checked ? ids.add(r.id) : ids.delete(r.id),
                );
                p.onSelected(ids);
              }}
            />
          );
        },
        cellRenderer: ({ data }: ICellRendererParams<Customer>) =>
          data && (
            <Checkbox
              aria-label={
                "选择客户 " +
                C.clean(data.company || data.contact || "待补充记录")
              }
              checked={p.selected.has(data.id)}
              onChange={(e) => {
                const ids = new Set(p.selected);
                e.target.checked ? ids.add(data.id) : ids.delete(data.id);
                p.onSelected(ids);
              }}
            />
          ),
      },
      {
        colId: "status",
        headerName: "资料状态",
        width: compact ? 130 : 158,
        minWidth: 125,
        pinned: compact ? undefined : "left",
        editable: false,
        cellRenderer: ({ data }: ICellRendererParams<Customer>) => {
          if (!data) return null;
          const pending = C.pendingIdentity(data),
            issues = C.issues(data),
            duplicate = p.summary.duplicates.has(data.id);
          const title = pending
            ? C.missingIdentity(data)
              ? "缺少：" + C.missingFixedFields(data).join("、")
              : "已补齐，请勾选后加入全部客户"
            : issues.length
              ? issues.join("；")
              : "通过基础格式检查，联系方式未核实";
          return (
            <div className="record-status">
              <Tooltip title={title}>
                <Tag
                  color={pending ? "orange" : issues.length ? "gold" : "cyan"}
                  icon={
                    pending || issues.length ? (
                      <WarningOutlined />
                    ) : (
                      <CheckCircleOutlined />
                    )
                  }
                >
                  {pending
                    ? C.missingIdentity(data)
                      ? "缺少主体"
                      : "待确认加入"
                    : issues.length
                      ? "待完善"
                      : "可用"}
                </Tag>
              </Tooltip>
              {duplicate && <span className="duplicate-note">疑似重复</span>}
            </div>
          );
        },
      },
      ...p.fields.map(
        (f): ColDef<Customer> => ({
          field: f.key,
          headerName: f.label,
          headerTooltip: C.fixedKeys.includes(f.key)
            ? "固定字段 · 不可删除"
            : f.label,
          width:
            f.key === "email"
              ? 270
              : f.key === "company"
                ? 205
                : f.key === "notes" || f.key === "product"
                  ? 230
                  : 165,
          minWidth: 120,
          editable: true,
          cellDataType: "text",
          valueGetter: (params) => C.clean(params.data?.[f.key]),
          tooltipValueGetter: (params) => C.clean(params.value),
          cellClass: (params) =>
            C.pendingIdentity(params.data || {}) &&
            C.fixedKeys.includes(f.key) &&
            !C.clean(params.value)
              ? "missing-value editable-cell"
              : "editable-cell",
          cellEditor: ["email", "notes", "product"].includes(f.key)
            ? "agLargeTextCellEditor"
            : "agTextCellEditor",
          cellEditorPopup: ["email", "notes", "product"].includes(f.key),
          cellEditorParams: { maxLength: 100000, rows: 4, cols: 45 },
        }),
      ),
      {
        colId: "source",
        headerName: "来源 / 原始内容",
        width: 240,
        minWidth: 170,
        editable: false,
        cellRenderer: ({ data }: ICellRendererParams<Customer>) =>
          data && (
            <div className="source-cell">
              <Tooltip title={data.source}>
                <span>
                  {data.source || "手动新增"}
                  {data.sourceRow ? ` · 第 ${data.sourceRow} 行` : ""}
                </span>
              </Tooltip>
              {!!data.sourceCells?.length && (
                <Button
                  type="link"
                  size="small"
                  icon={<FileSearchOutlined />}
                  onClick={() => p.onOriginal(data)}
                >
                  查看原始内容
                </Button>
              )}
            </div>
          ),
      },
    ],
    [
      p.fields,
      p.selected,
      visible,
      p.summary,
      p.onSelected,
      p.onOriginal,
      compact,
    ],
  );
  return (
    <section className="library-panel" id="customer-library">
      <div className="library-heading">
        <div>
          <span className="eyebrow">CUSTOMER LIBRARY</span>
          <h2>
            客户资料库{" "}
            <span className="count-chip">
              {p.records.length.toLocaleString()} 条记录
            </span>
          </h2>
        </div>
        <div className="export-actions">
          <Select
            aria-label="导出格式"
            value={p.format}
            onChange={p.onFormat}
            options={[
              { value: "xlsx", label: "Excel (.xlsx)" },
              { value: "csv", label: "CSV (.csv)" },
            ]}
          />
          <Button
            type="primary"
            icon={<DownloadOutlined />}
            disabled={!p.records.length}
            onClick={p.onExport}
          >
            导出客户表
          </Button>
        </div>
      </div>
      <RulesPanel category={p.category} />
      <div className="library-toolbar">
        <Input
          className="customer-search"
          prefix={<SearchOutlined />}
          allowClear
          aria-label="搜索客户"
          placeholder="搜索公司、联系人、邮箱或原始内容…"
          value={p.query}
          onChange={(e) => p.onQuery(e.target.value)}
        />
        <div className="row-actions">
          {p.category === "missing" && (
            <Button
              className="admit-button"
              type="primary"
              icon={<CheckCircleOutlined />}
              disabled={!ready}
              onClick={p.onAdmit}
            >
              加入全部客户{ready ? ` (${ready})` : ""}
            </Button>
          )}
          <Button icon={<PlusOutlined />} onClick={p.onAdd}>
            新增客户
          </Button>
          <Button
            danger
            icon={<DeleteOutlined />}
            disabled={!p.selected.size}
            onClick={p.onDelete}
          >
            删除{p.selected.size ? ` (${p.selected.size})` : ""}
          </Button>
          <Tooltip title="管理字段">
            <Button
              aria-label="管理字段"
              icon={<SettingOutlined />}
              onClick={p.onFields}
            />
          </Tooltip>
          {!!p.deletedCount && <Button icon={<UndoOutlined />} onClick={p.onUndoDelete}>
            撤销上次删除 ({p.deletedCount})
          </Button>}
        </div>
      </div>
      <div className="table-context">
        <span>
          {p.selected.size
            ? `已选择 ${p.selected.size} 条 · 跨页保留；记录移出筛选结果时自动取消勾选`
            : "单击单元格编辑，点击表格外保存。固定字段缺失时会单独保留。"}
        </span>
        {p.category === "missing" && (
          <span className="pending-hint">
            {ready
              ? `${ready} 条已补齐，可加入全部客户`
              : "补齐三项固定字段后，勾选并加入"}
          </span>
        )}
      </div>
      {visible.length ? (
        <div className="customer-grid" data-testid="customer-grid">
          <AgGridReact<Customer>
            ref={grid}
            theme={gridTheme}
            rowData={visible}
            columnDefs={columns}
            getRowId={(params) => params.data.id}
            defaultColDef={{
              resizable: true,
              sortable: false,
              suppressHeaderMenuButton: true,
            }}
            singleClickEdit
            readOnlyEdit
            stopEditingWhenCellsLoseFocus
            tooltipShowDelay={300}
            suppressCellFocus={false}
            animateRows={false}
            suppressDragLeaveHidesColumns
            onCellEditRequest={(e) => {
              const key = e.colDef.field;
              if (key) p.onEdit(e.data.id, key, String(e.newValue ?? ""));
            }}
            rowClassRules={{
              "selected-customer": (params) =>
                Boolean(params.data && p.selected.has(params.data.id)),
            }}
          />
        </div>
      ) : (
        <div className="library-empty">
          <div className="empty-illustration">
            <div className="empty-grid">
              <span />
              <span />
              <span />
              <span />
              <span />
              <span />
            </div>
            <CloudUploadOutlined />
          </div>
          <Empty image={Empty.PRESENTED_IMAGE_SIMPLE} description={false} />
          <h3>
            {p.records.length ? "没有匹配的记录" : "让客户资料，从这里变得有序"}
          </h3>
          <p>
            {p.records.length
              ? "切换分类或调整搜索条件，继续整理客户资料。"
              : "导入不同格式的客户表，统一字段，轻松校对并导出。"}
          </p>
          {!p.records.length && (
            <div className="empty-actions">
              <Button
                type="primary"
                size="large"
                icon={<PlusOutlined />}
                onClick={p.onImport}
              >
                导入客户表
              </Button>
              <Button size="large" onClick={p.onDemo}>
                体验示例数据
              </Button>
            </div>
          )}
        </div>
      )}
      <div className="library-footer">
        <span>
          共保留 {p.records.length.toLocaleString()} 条 · 当前分类匹配{" "}
          {p.filtered.length.toLocaleString()} 条
        </span>
        <Pagination
          size="small"
          current={page}
          pageSize={p.pageSize}
          total={p.filtered.length}
          showSizeChanger
          pageSizeOptions={[10, 25, 50, 100]}
          onChange={p.onPage}
          showTotal={(total, range) =>
            total ? `${range[0]}–${range[1]} 条` : "0 条"
          }
        />
      </div>
      <div className="export-footnote">
        <LockIcon />
        导出包含全部记录与原始内容，不受当前搜索、分类或分页影响。
      </div>
    </section>
  );
}
function LockIcon() {
  return <CheckCircleOutlined />;
}
