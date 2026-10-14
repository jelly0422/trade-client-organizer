(function (root) {
  const fields = [
    { key: 'company', label: '公司名', aliases: ['公司','公司名','公司名称','企业名称','company','company name','organization','account name'] },
    { key: 'contact', label: '客户名称', aliases: ['客户名称','联系人','姓名','客户姓名','customer','contact','contact name','name','full name','buyer'] },
    { key: 'email', label: '邮箱', aliases: ['邮箱','电子邮箱','电子邮件','email','e-mail','email address','mail'] },
    { key: 'phone', label: '电话', aliases: ['电话','联系电话','手机','手机号','phone','telephone','tel','mobile','whatsapp'] },
    { key: 'country', label: '国家 / 地区', aliases: ['国家','地区','国家地区','country','region','market','country / region'] },
    { key: 'product', label: '产品需求', aliases: ['产品需求','需求','产品','意向产品','采购产品','product','products','product interest','requirements','inquiry'] },
    { key: 'website', label: '公司网站', aliases: ['网址','网站','公司网站','website','url','web'] },
    { key: 'notes', label: '备注', aliases: ['备注','说明','notes','note','remarks','comment'] }
  ];
  const clean = v => String(v ?? '').trim();
  const normalize = v => clean(v).toLowerCase().replace(/[\s_\-\/（）()]/g, '');
  const fixedKeys = Object.freeze(['company', 'contact', 'email']);
  const metadataLabels = ['来源', '来源文件 / 工作表', '原始行号', '校对提示', '状态', '记录分类', '原始内容'];
  const missingIdentity = record => fixedKeys.some(key => !clean(record[key]));
  const pendingIdentity = record => Boolean(record.pendingIdentity) || missingIdentity(record);
  const missingFixedFields = record => fields.filter(f => fixedKeys.includes(f.key) && !clean(record[f.key])).map(f => f.label);
  function admitRecord(record) {
    if (missingIdentity(record)) throw Error('请先补齐固定字段：' + missingFixedFields(record).join('、') + '。');
    return { ...record, pendingIdentity: false };
  }
  const reviewRules = [
    { label: '邮箱格式待核对', description: '邮箱非空，但不符合基本格式。多个邮箱用逗号、分号或换行分隔，每一项都需符合格式。', check: r => Boolean(clean(r.email)) && !clean(r.email).split(/[;,；，\n]+/).every(e => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e.trim())) }
  ];
  function addField(schema, name, key) {
    const label = clean(name);
    if (!normalize(label)) throw Error('请输入字段名称。');
    if (label.length > 30) throw Error('字段名称不能超过 30 个字符。');
    if (schema.length >= 40) throw Error('最多可设置 40 个字段。');
    if (schema.some(f => normalize(f.label) === normalize(label)) || metadataLabels.some(n => normalize(n) === normalize(label))) throw Error('字段名称已存在或用于系统信息，请换一个名称。');
    if (!/^custom_[a-z0-9]+$/i.test(key) || schema.some(f => f.key === key)) throw Error('自定义字段标识无效。');
    const builtIn = fields.find(f => normalize(f.label) === normalize(label) && !schema.some(existing => existing.key === f.key));
    return [...schema, builtIn ? {...builtIn} : { key, label, aliases: [label] }];
  }
  function removeField(schema, key) {
    if (fixedKeys.includes(key)) throw Error('客户名称、公司名和邮箱是固定字段，不能删除。');
    return schema.filter(f => f.key !== key);
  }
  function autoMap(headers, schema = fields, previous = {}) {
    const used = new Set();
    const mapping = {};
    // Preserve manual choices, including explicitly unmapped fields.
    schema.forEach(f => {
      if (!Object.hasOwn(previous, f.key)) return;
      const value = previous[f.key];
      if (value === '') mapping[f.key] = '';
      else if (/^\d+$/.test(value) && Number(value) < headers.length && !used.has(Number(value))) { mapping[f.key] = value; used.add(Number(value)); }
    });
    return Object.fromEntries(schema.map(f => {
      if (Object.hasOwn(mapping, f.key)) return [f.key, mapping[f.key]];
      const index = headers.findIndex((h,i) => !used.has(i) && [f.label, ...f.aliases].some(a => normalize(a) === normalize(h)));
      if (index >= 0) used.add(index);
      return [f.key, index < 0 ? '' : String(index)];
    }));
  }
  function detectHeader(rows, schema = fields) {
    let best = 0, score = -1;
    rows.slice(0, 30).forEach((row,i) => {
      const s = Object.values(autoMap(row, schema)).filter(v => v !== '').length;
      if (s > score) { score = s; best = i; }
    });
    return best;
  }
  function mappedRows(rows, headerIndex, mapping, source, schema = fields, rowOffset = 0, columnOffset = 0) {
    return rows.slice(headerIndex + 1).flatMap((row, i) => {
      if (!row.some(v => clean(v))) return [];
      const record = Object.fromEntries(schema.map(f => [f.key, mapping[f.key] === '' || mapping[f.key] === undefined ? '' : clean(row[Number(mapping[f.key])])]));
      const pending = missingIdentity(record);
      const sourceCells = pending ? row.map((value, column) => ({ column: columnOffset + column + 1, label: clean(rows[headerIndex]?.[column]) || `第 ${columnOffset + column + 1} 列`, value: String(value ?? '') })).filter(cell => clean(cell.value)) : [];
      return [{ ...record, source, sourceRow: rowOffset + headerIndex + i + 2, pendingIdentity: pending, sourceCells }];
    });
  }
  function issues(record, schema = fields) {
    return reviewRules.filter(rule => rule.check(record, schema)).map(rule => rule.label);
  }
  function duplicateIds(records) {
    const groups = new Map(), ids = new Set();
    records.forEach(r => {
      const email = normalize(r.email);
      const key = email ? 'email:' + email : normalize(r.company) && normalize(r.contact) ? 'name:' + normalize(r.company) + '|' + normalize(r.contact) : '';
      if (!key) return;
      if (groups.has(key)) { ids.add(groups.get(key)); ids.add(r.id); } else groups.set(key, r.id);
    });
    return ids;
  }
  function csvCell(value) {
    let v = String(value ?? '');
    if (/^[\s]*[=+\-@\t\r]/.test(v)) v = "'" + v;
    return '"' + v.replace(/"/g, '""') + '"';
  }
  function exportMatrix(records, schema = fields, template = false) {
    const headers = schema.map(f => f.label);
    if (template) return [headers];
    const duplicates = duplicateIds(records.filter(r => !pendingIdentity(r)));
    return [[...headers, '来源文件 / 工作表', '原始行号', '校对提示', '记录分类'], ...records.map(r => [
      ...schema.map(f => String(r[f.key] ?? '')), String(r.source ?? ''), String(r.sourceRow || ''),
      [...(pendingIdentity(r) ? [missingIdentity(r) ? '缺少固定字段：' + missingFixedFields(r).join('、') : '已补齐三项固定字段，待确认加入全部客户'] : []), ...issues(r, schema), ...(duplicates.has(r.id) ? ['疑似重复'] : [])].join('；'),
      pendingIdentity(r) ? '缺少主体' : '全部客户'
    ])];
  }
  function recoveryMatrix(records) {
    return [['来源文件 / 工作表', '原始行号', '原始列序号', '原始表头', '原始内容'], ...records.flatMap(r => (r.sourceCells || []).map(cell => [String(r.source ?? ''), String(r.sourceRow || ''), String(cell.column), cell.label, cell.value]))];
  }
  const api = {fields, fixedKeys, addField, removeField, exportMatrix, recoveryMatrix, clean, autoMap, detectHeader, mappedRows, issues, reviewRules, missingIdentity, pendingIdentity, missingFixedFields, admitRecord, duplicateIds, csvCell};
  root.TradeCore = api;
  if (typeof module !== 'undefined') module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
