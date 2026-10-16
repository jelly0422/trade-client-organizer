import { InfoCircleOutlined } from "@ant-design/icons";
import type { Category } from "../state/workspace";
import { reviewRules } from "../domain/core";

const rules = {
  review: {
    title: "待完善",
    intro: "检查已加入全部客户的记录；缺少主体不参与。",
    items: [
      ["逐个检查邮箱", reviewRules[0].description],
      [
        "格式与长度",
        "包含一个 @，域名至少含一个英文句点；不能含内部空白、连续句点或域名下划线，域名每段不能以短横线开头或结尾。单个邮箱最长 254 字符，@ 前最长 64 字符，域名每段最长 63 字符；分隔后出现空项也提示核对。",
      ],
      [
        "人工核对范围",
        "中文等国际化邮箱、带引号的邮箱、IP 域名和“姓名 <邮箱>”写法暂不自动通过；原始文本保留。电话、地区、产品需求、网站、备注和自定义字段为空不触发。",
      ],
    ],
    note: "仅校验文本格式，不验证能否收信或客户真实性；同一条记录可同时属于疑似重复。",
  },
  duplicate: {
    title: "疑似重复",
    intro:
      "跨来源、跨工作表比较已加入全部客户的记录；仅提示，不自动合并或删除。",
    items: [
      [
        "优先按邮箱",
        "忽略英文大小写、所有空白、下划线 _、短横线 -、斜杠 / 及中英文圆括号，处理后相同即提示。例如 a-b@example.com 与 ab@example.com 目前会归为疑似重复，需人工核对。",
      ],
      [
        "邮箱无法比较时",
        "邮箱经上述处理后为空，则用相同方式比较公司名和客户名称；两项都非空且分别相同才提示。",
      ],
      [
        "多个邮箱与计数",
        "整段邮箱一起比较，不逐个交叉匹配；顺序或分隔符不同可能不会匹配。同组全部记录（包括首条）均计数，不是重复组数；可能同时属于待完善。",
      ],
    ],
    note: "缺少主体不参与比对。请核对客户和原始邮箱后再处理。",
  },
  missing: {
    title: "缺少主体",
    intro: "客户名称、公司名、邮箱任意一项为空就进入此分类，单独保留。",
    items: [
      [
        "固定字段必须齐全",
        "去除首尾空白后判断是否为空，电话、产品需求或自定义字段不能代替三项固定字段。导入时已齐全的记录直接进入全部客户。",
      ],
      [
        "补齐后明确确认",
        "已在此分类的记录补齐后仍显示“待确认加入”。勾选记录，点击右上角“加入全部客户”；批量仅加入已补齐的部分，其余继续保留。邮箱格式异常不阻止加入，加入后列为待完善。",
      ],
      [
        "保留原始信息",
        "表头后的非空行即使没有匹配字段也会保留，整行空白才跳过。可查看导入时保留的原始内容。已加入的记录再次清空固定字段，会回到此分类，需要重新补齐并确认。",
      ],
    ],
    note: "此分类不计入全部客户、待完善或疑似重复。导出包含这些记录及原始内容。",
  },
};
export function RulesPanel({ category }: { category: Category }) {
  if (category === "all") return null;
  const rule = rules[category];
  return (
    <section
      className={"rules-panel rules-" + category}
      data-testid="classification-rules"
    >
      <div className="rules-heading">
        <InfoCircleOutlined />
        <div>
          <strong>{rule.title} · 判定规则</strong>
          <p>{rule.intro}</p>
        </div>
      </div>
      <div className="rules-grid">
        {rule.items.map(([title, text]) => (
          <div key={title}>
            <b>{title}</b>
            <p>{text}</p>
          </div>
        ))}
      </div>
      <p className="rules-footnote">{rule.note}</p>
    </section>
  );
}
