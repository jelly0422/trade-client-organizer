import { useState } from "react";
import { App, Button, Drawer, Input, Tag, Tooltip } from "antd";
import { DeleteOutlined, LockOutlined, PlusOutlined } from "@ant-design/icons";
import * as C from "../domain/core";
import type { Field, Customer } from "../domain/core";

interface Props {
  open: boolean;
  onClose: () => void;
  fields: Field[];
  records: Customer[];
  onAdd: (name: string) => void;
  onRemove: (key: string) => void;
}
export function FieldsDrawer({
  open,
  onClose,
  fields,
  records,
  onAdd,
  onRemove,
}: Props) {
  const [name, setName] = useState("");
  const { modal, message } = App.useApp();
  const add = () => {
    try {
      onAdd(name);
      setName("");
    } catch (e) {
      message.error((e as Error).message);
    }
  };
  const remove = (field: Field) => {
    const count = records.filter((r) => C.clean(r[field.key])).length;
    if (count)
      modal.confirm({
        title: `删除“${field.label}”？`,
        content: `这将移除 ${count} 条客户中的该字段内容及对应映射，无法撤销。`,
        okText: "删除字段",
        okButtonProps: { danger: true },
        cancelText: "取消",
        onOk: () => onRemove(field.key),
      });
    else onRemove(field.key);
  };
  return (
    <Drawer title="管理表格字段" open={open} onClose={onClose} size={460}>
      <div className="drawer-intro">
        <span className="eyebrow">YOUR WORKSPACE</span>
        <h2>让每一列，适合你的业务</h2>
        <p>字段同步用于映射、编辑、搜索和导出。最多设置 40 个字段。</p>
      </div>
      <div className="field-add">
        <Input
          aria-label="自定义字段名称"
          placeholder="例如：采购数量、客户等级"
          maxLength={30}
          value={name}
          onChange={(e) => setName(e.target.value)}
          onPressEnter={add}
        />
        <Button type="primary" icon={<PlusOutlined />} onClick={add}>
          添加
        </Button>
      </div>
      <div className="list-heading">
        <span>当前字段</span>
        <span>{fields.length} / 40</span>
      </div>
      <div className="field-list">
        {fields.map((f, i) => (
          <div className="field-row" key={f.key}>
            <span className="field-order">
              {String(i + 1).padStart(2, "0")}
            </span>
            <strong>{f.label}</strong>
            {C.fixedKeys.includes(f.key) ? (
              <Tag icon={<LockOutlined />} color="cyan">
                固定
              </Tag>
            ) : (
              <Tooltip title={"删除" + f.label}>
                <Button
                  type="text"
                  danger
                  aria-label={"删除字段 " + f.label}
                  icon={<DeleteOutlined />}
                  onClick={() => remove(f)}
                />
              </Tooltip>
            )}
          </div>
        ))}
      </div>
      <p className="muted-note">
        客户名称、公司名和邮箱不可删除。字段设置仅在当前页面会话内有效，刷新前请导出。
      </p>
    </Drawer>
  );
}
