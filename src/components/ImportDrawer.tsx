import { useMemo } from "react";
import {
  Alert,
  App,
  Button,
  Drawer,
  Empty,
  InputNumber,
  Radio,
  Select,
  Steps,
  Table,
  Tabs,
  Tag,
  Tooltip,
  Upload,
} from "antd";
import {
  ArrowRightOutlined,
  CheckCircleOutlined,
  CloudUploadOutlined,
  FileExcelOutlined,
  LockOutlined,
} from "@ant-design/icons";
import * as C from "../domain/core";
import type { Customer, Field } from "../domain/core";
import { FieldSettings } from "./FieldSettings";
import {
  mappedWorkbook,
  offsets,
  rowsFor,
  xlsx,
  type FileSource,
  type SheetConfig,
} from "../services/workbooks";

interface Props {
  open: boolean;
  onClose: () => void;
  fields: Field[];
  customers: Customer[];
  panel: "mapping" | "fields";
  onPanel: (panel: "mapping" | "fields") => void;
  onAddField: (name: string) => void;
  onRemoveField: (key: string) => void;
  onMerge: (merge: boolean) => void;
  sources: FileSource[];
  current: string;
  imported: Set<string>;
  busy: boolean;
  onFiles: (files: File[]) => void;
  onSelect: (id: string) => void;
  onSheet: (name: string) => void;
  onConfig: (config: SheetConfig) => void;
  onEncoding: (encoding: string) => void;
  onApply: () => void;
  onDemo: () => void;
}
export function ImportDrawer(p: Props) {
  const { message } = App.useApp();
  const source = p.sources.find((s) => s.id === p.current),
    cfg = source?.configs[source.sheet];
  const rows = useMemo(() => (source ? rowsFor(source) : []), [source]);
  const preview = useMemo(() => {
    try { return { groups: source ? mappedWorkbook(source, p.fields) : [], error: "" }; }
    catch (error) { return { groups: [], error: (error as Error).message }; }
  }, [source, p.fields]);
  const records = preview.groups.flatMap(group => group.records);
  const missing = records.filter(C.pendingIdentity).length;
  const headers = cfg ? rows[cfg.header] || [] : [];
  const done = source ? p.imported.has(source.id) : false;
  const needsChoice = !!source && source.book.SheetNames.length > 1 && source.mergeSheets === undefined;
  const offset = source ? offsets(source) : { r: 0, c: 0 };
  const mappingContent = (
    <>
      <Steps
        size="small"
        current={source ? (done ? 2 : 1) : 0}
        items={[
          { title: "选择文件" },
          { title: "核对字段映射" },
          { title: "加入客户资料库" },
        ]}
      />
      <Upload.Dragger
        className="import-dropzone"
        accept=".xlsx,.xls,.csv,.tsv,.ods"
        multiple
        showUploadList={false}
        disabled={p.busy}
        beforeUpload={(file, fileList) => {
          if (file === fileList[0]) p.onFiles(fileList);
          return false;
        }}
      >
        <div className="upload-symbol">
          <CloudUploadOutlined />
        </div>
        <h3>{p.busy ? "正在读取文件…" : "拖放客户表到这里，或点击选择"}</h3>
        <p>
          XLSX · XLS · CSV · TSV · ODS<span>单个文件不超过 20 MB</span>
        </p>
      </Upload.Dragger>
      {!source ? (
        <div className="import-empty">
          <p>还没有准备好文件？先用虚构客户数据体验完整流程。</p>
          <Button icon={<FileExcelOutlined />} onClick={p.onDemo}>
            体验示例数据
          </Button>
        </div>
      ) : (
        <>
          <div className="import-files">
            {p.sources.map((s) => (
              <button
                className={"file-chip " + (s.id === p.current ? "active" : "")}
                key={s.id}
                onClick={() => p.onSelect(s.id)}
              >
                <FileExcelOutlined />
                <span>
                  {s.name}
                  <small>
                    {s.book.SheetNames.length} 张工作表
                    {s.demo ? " · 虚构示例" : ""}
                  </small>
                </span>
                {p.imported.has(s.id) && (
                  <CheckCircleOutlined />
                )}
              </button>
            ))}
          </div>
          {source.book.SheetNames.length > 1 && (
            <section className="sheet-mode" aria-label="工作表导入方式">
              <h3>发现 {source.book.SheetNames.length} 张工作表，是否合并？</h3>
              <Radio.Group aria-label="是否合并工作表" value={source.mergeSheets} disabled={done || p.busy}
                onChange={event => p.onMerge(event.target.value)}>
                <Radio value={false}>保留独立工作表</Radio>
                <Radio value={true}>合并为一张客户表</Radio>
              </Radio.Group>
              <p>{needsChoice ? "请选择一种导入方式；可逐表核对表头和映射后，一次导入整个文件。"
                : source.mergeSheets ? "全部工作表的数据统一汇总，仍保留每条记录的来源。"
                : "导入后可按工作表切换，Excel 导出保留各表及空表。缺少主体的记录仍属于原工作表。"}</p>
            </section>
          )}
          <div className="mapping-options">
            <label>
              工作表
              <Select
                aria-label="选择工作表"
                value={source.sheet}
                options={source.book.SheetNames.map((value) => ({
                  value,
                  label: value + " · " + (preview.groups.find(group => group.name === value)?.records.length ?? 0) + " 条",
                }))}
                onChange={p.onSheet}
              />
            </label>
            <label>
              表头所在行
              <InputNumber
                aria-label="表头所在行"
                min={offset.r + 1}
                max={offset.r + Math.max(1, rows.length)}
                disabled={done}
                value={(cfg?.header || 0) + offset.r + 1}
                onChange={(value) => {
                  const header = Math.max(
                    0,
                    Math.min(rows.length - 1, Number(value) - offset.r - 1),
                  );
                  p.onConfig({
                    header,
                    mapping: C.autoMap(rows[header] || [], p.fields),
                  });
                }}
              />
            </label>
            <label>
              文本编码
              <Select
                aria-label="文本编码"
                disabled={done || !/\.(csv|tsv)$/i.test(source.name)}
                value={source.encoding}
                options={[
                  { value: "utf-8", label: "UTF-8" },
                  { value: "gb18030", label: "GB18030 / GBK" },
                  { value: "utf-16le", label: "UTF-16 LE" },
                ]}
                onChange={p.onEncoding}
              />
            </label>
          </div>
          <p className="muted-note">当前核对：{source.sheet}。本次添加整个文件的 {source.book.SheetNames.length} 张工作表；各表可以设置不同的表头和映射。</p>
          <div className="mapping-title">
            <h3>核对字段映射</h3>
            <Button size="small" onClick={() => p.onPanel("fields")}>设置字段</Button>
            <Tag color="cyan">
              已匹配 {Object.values(cfg?.mapping || {}).filter(Boolean).length}{" "}
              / {p.fields.length} 项
            </Tag>
          </div>
          <p className="muted-note">
            左侧是统一字段，右侧选择来源列。未匹配的固定字段会使记录进入“缺少主体”。
          </p>
          <div className="mapping-list">
            {p.fields.map((f) => {
              const value = cfg?.mapping[f.key] ?? "";
              const sample =
                value === ""
                  ? "该字段未匹配"
                  : (rows
                      .slice((cfg?.header || 0) + 1)
                      .map((row) => row[Number(value)])
                      .find((v) => C.clean(v)) ?? "该列暂无数据");
              return (
                <div
                  className={"mapping-row " + (value === "" ? "unmapped" : "")}
                  key={f.key}
                >
                  <div className="mapping-label">
                    {f.label}
                    {C.fixedKeys.includes(f.key) && (
                      <Tooltip title="固定字段">
                        <LockOutlined />
                      </Tooltip>
                    )}
                  </div>
                  <ArrowRightOutlined className="mapping-arrow" />
                  <div className="mapping-value">
                    <Select
                      aria-label={"映射" + f.label}
                      value={value}
                      disabled={done}
                      options={[
                        { value: "", label: "不导入此字段" },
                        ...headers.map((h, i) => ({
                          value: String(i),
                          label: `${C.clean(h) || "未命名列"} · ${xlsx().utils.encode_col(offset.c + i)}`,
                        })),
                      ]}
                      onChange={(next) => {
                        if (
                          next !== "" &&
                          Object.entries(cfg!.mapping).some(
                            ([key, v]) => key !== f.key && v === next,
                          )
                        ) {
                          message.warning(
                            "该来源列已映射到其他字段，请先取消原映射。",
                          );
                          return;
                        }
                        p.onConfig({
                          ...cfg!,
                          mapping: { ...cfg!.mapping, [f.key]: next },
                        });
                      }}
                    />
                    <small title={String(sample)}>{String(sample)}</small>
                  </div>
                </div>
              );
            })}
          </div>
          <div className="preview-heading">
            <h3>来源数据预览</h3>
            <span>前 5 条 · 内容以文本读取</span>
          </div>
          {headers.length ? (
            <Table
              size="small"
              className="source-preview"
              pagination={false}
              scroll={{ x: "max-content" }}
              rowKey="key"
              columns={headers.map((h, i) => ({
                title: C.clean(h) || "未命名列",
                dataIndex: String(i),
                width: 160,
                ellipsis: true,
              }))}
              dataSource={rows
                .slice((cfg?.header || 0) + 1, (cfg?.header || 0) + 6)
                .map((row, i) =>
                  Object.fromEntries([
                    ["key", String(i)],
                    ...headers.map((_, j) => [String(j), C.clean(row[j])]),
                  ]),
                )}
            />
          ) : (
            <Empty description="当前表头行为空，请调整表头所在行" />
          )}
          {preview.error && <Alert showIcon type="error" title={preview.error} />}
          <div className="import-summary">
            <Alert
              type={done ? "success" : "info"}
              showIcon
              title={
                done
                  ? "这份文件已完成导入"
                  : `识别 ${records.length} 条非空记录`
              }
              description={
                done
                  ? "可在资料库整理或删除工作表；已删除的数据可撤销恢复，或重新选择文件导入。"
                  : `${records.length - missing} 条进入全部客户，${missing} 条进入缺少主体；仅跳过整行空白。${source.mergeSheets === false || source.book.SheetNames.length === 1 ? "空工作表也会保留。" : ""}`
              }
            />
            <Button
              type="primary"
              size="large"
              disabled={done || needsChoice || !!preview.error || (!records.length && source.mergeSheets === true) || p.busy}
              icon={<ArrowRightOutlined />}
              onClick={p.onApply}
            >
              {done ? "已添加" : "添加到客户资料库"}
            </Button>
          </div>
        </>
      )}
    </>
  );
  return (
    <Drawer title="导入与映射" open={p.open} onClose={p.onClose} size={1040}
      className="import-drawer" extra={<Tag color="cyan">文件仅在本地处理</Tag>}>
      <Tabs activeKey={p.panel} onChange={key => p.onPanel(key as "mapping" | "fields")}
        items={[
          { key: "mapping", label: "文件与映射", children: mappingContent },
          { key: "fields", label: "字段设置", children: <>
            <FieldSettings fields={p.fields} records={p.customers} onAdd={p.onAddField} onRemove={p.onRemoveField} />
            <Button onClick={() => p.onPanel("mapping")}>返回文件与映射</Button>
          </> },
        ]} />
    </Drawer>
  );
}
