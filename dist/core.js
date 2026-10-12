(function (root) {
  const fields = [
    { key: 'company', label: '公司名称', aliases: ['公司','公司名称','客户名称','企业名称','company','company name','customer','organization','account name'] },
    { key: 'contact', label: '联系人', aliases: ['联系人','姓名','客户姓名','contact','contact name','name','full name','buyer'] },
    { key: 'email', label: '邮箱', aliases: ['邮箱','电子邮箱','电子邮件','email','e-mail','email address','mail'] },
    { key: 'phone', label: '电话', aliases: ['电话','联系电话','手机','手机号','phone','telephone','tel','mobile','whatsapp'] },
    { key: 'country', label: '国家 / 地区', aliases: ['国家','地区','国家地区','country','region','market','country / region'] },
    { key: 'product', label: '产品需求', aliases: ['产品需求','需求','产品','意向产品','采购产品','product','products','product interest','requirements','inquiry'] },
    { key: 'website', label: '公司网站', aliases: ['网址','网站','公司网站','website','url','web'] },
    { key: 'notes', label: '备注', aliases: ['备注','说明','notes','note','remarks','comment'] }
  ];
  const clean = v => String(v ?? '').trim();
  const normalize = v => clean(v).toLowerCase().replace(/[\s_\-\/（）()]/g, '');
  function autoMap(headers) {
    const used = new Set();
    return Object.fromEntries(fields.map(f => {
      const index = headers.findIndex((h,i) => !used.has(i) && f.aliases.some(a => normalize(a) === normalize(h)));
      if (index >= 0) used.add(index);
      return [f.key, index < 0 ? '' : String(index)];
    }));
  }
  function detectHeader(rows) {
    let best = 0, score = -1;
    rows.slice(0, 30).forEach((row,i) => {
      const s = Object.values(autoMap(row)).filter(v => v !== '').length;
      if (s > score) { score = s; best = i; }
    });
    return best;
  }
  function mappedRows(rows, headerIndex, mapping, source) {
    return rows.slice(headerIndex + 1).flatMap((row, i) => {
      if (!row.some(v => clean(v))) return [];
      const record = Object.fromEntries(fields.map(f => [f.key, mapping[f.key] === '' || mapping[f.key] === undefined ? '' : clean(row[Number(mapping[f.key])])]));
      if (!fields.some(f => record[f.key])) return [];
      return [{ ...record, source, sourceRow: headerIndex + i + 2 }];
    });
  }
  function issues(record) {
    const result = [];
    if (!clean(record.company) && !clean(record.contact)) result.push('缺少公司和联系人');
    if (!clean(record.email) && !clean(record.phone)) result.push('缺少联系方式');
    if (clean(record.email) && !clean(record.email).split(/[;,；，\n]+/).every(e => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e.trim()))) result.push('邮箱格式待核对');
    return result;
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
  const api = {fields, clean, autoMap, detectHeader, mappedRows, issues, duplicateIds, csvCell};
  root.TradeCore = api;
  if (typeof module !== 'undefined') module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
