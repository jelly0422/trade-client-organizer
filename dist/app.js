/* Customer records remain in memory. No analytics, storage, or upload requests. */
(() => {
  'use strict';
  const C = window.TradeCore, $ = id => document.getElementById(id);
  const state = { fields: C.fields.map(f => ({...f})), sources: [], current: '', records: [], selected: new Set(), imported: new Set(), filter: 'all', query: '', page: 1, pageSize: 25, dirty: false };
  let sequence = 0, toastTimer;
  const uid = () => String(++sequence);
  const icon = name => `<svg class="ui-icon" aria-hidden="true"><use href="#icon-${name}"/></svg>`;
  const esc = value => String(value ?? '').replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
  function toast(message, error = false) { $('toast').innerHTML = icon(error ? 'alert' : 'check-circle') + `<span>${esc(message)}</span>`; $('toast').className = 'toast' + (error ? ' error' : ''); $('toast').hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => $('toast').hidden = true, 5000); }
  function libraryReady() { if (window.XLSX) return true; toast('Excel 解析库未加载，请确认 vendor 文件完整后刷新。', true); return false; }
  function current() { return state.sources.find(s => s.id === state.current); }
  function settings(source) { return source.configs[source.sheet]; }
  function rowsFor(source) { return XLSX.utils.sheet_to_json(source.book.Sheets[source.sheet], { header: 1, defval: '', raw: false, blankrows: true }); }
  function rowOffset(source) { const ref = source.book.Sheets[source.sheet]['!ref']; return ref ? XLSX.utils.decode_range(ref).s.r : 0; }
  function columnOffset(source) { const ref = source.book.Sheets[source.sheet]['!ref']; return ref ? XLSX.utils.decode_range(ref).s.c : 0; }
  function mappedSource(source) { const cfg = settings(source); return C.mappedRows(rowsFor(source), cfg.header, cfg.mapping, source.name + ' / ' + source.sheet, state.fields, rowOffset(source), columnOffset(source)); }
  const customers = () => state.records.filter(r => !C.pendingIdentity(r));
  const duplicatesForCustomers = () => C.duplicateIds(customers());
  function initConfig(source) { if (!source.configs[source.sheet]) { const rows = rowsFor(source), header = C.detectHeader(rows, state.fields); source.configs[source.sheet] = { header, mapping: C.autoMap(rows[header] || [], state.fields) }; } }
  function renderFields() {
    $('fieldList').innerHTML = state.fields.map(f => `<span class="managed-field ${C.fixedKeys.includes(f.key) ? 'fixed-field' : ''}"><span>${esc(f.label)}</span>${C.fixedKeys.includes(f.key) ? `<small title="此字段不可删除">${icon('lock')}固定</small>` : `<button type="button" data-remove-field="${f.key}" aria-label="删除字段 ${esc(f.label)}" title="删除字段 ${esc(f.label)}">${icon('close')}</button>`}</span>`).join('');
    $('fieldCount').textContent = `${state.fields.length} / 40`;
    $('guideFields').innerHTML = state.fields.map(f => `<span>${esc(f.label)}</span>`).join('');
    document.querySelectorAll('#dataHeader th[data-field]').forEach(th => th.remove());
    $('sourceHeader').insertAdjacentHTML('beforebegin', state.fields.map(f => `<th data-field="${f.key}">${esc(f.label)}</th>`).join(''));
  }
  function syncSchema() {
    state.sources.forEach(source => Object.entries(source.configs).forEach(([sheet, cfg]) => {
      const rows = XLSX.utils.sheet_to_json(source.book.Sheets[sheet], {header:1,defval:'',raw:false,blankrows:true});
      cfg.mapping = C.autoMap(rows[cfg.header] || [], state.fields, cfg.mapping);
    }));
    if (state.records.length) state.dirty = true;
    renderFields(); if (current()) renderMapping(); renderData();
  }
  function addCustomField(event) {
    event.preventDefault();
    try {
      state.fields = C.addField(state.fields, $('fieldName').value, 'custom_' + uid());
      const field = state.fields[state.fields.length - 1];
      state.records.forEach(record => { record[field.key] = ''; });
      $('fieldName').value = ''; syncSchema();
      toast(`已添加“${field.label}”，可映射待导入数据或在客户列表中填写。`);
      $('fieldName').focus();
    } catch (error) { toast(error.message, true); }
  }
  function removeCustomField(key) {
    const field = state.fields.find(f => f.key === key); if (!field) return;
    try {
      const next = C.removeField(state.fields, key);
      const populated = state.records.filter(r => C.clean(r[key])).length;
      if (populated && !confirm(`删除“${field.label}”字段将移除 ${populated} 条客户中的该字段内容，无法撤销。是否继续？`)) return;
      state.fields = next; state.records.forEach(record => { delete record[key]; });
      syncSchema(); toast(`已删除“${field.label}”字段`);
    } catch (error) { toast(error.message, true); }
  }
  function parseBook(bytes, name, encoding = 'utf-8') {
    if (/\.(csv|tsv)$/i.test(name)) return XLSX.read(new TextDecoder(encoding).decode(bytes), { type: 'string', raw: true });
    return XLSX.read(bytes, { type: 'array', cellText: true });
  }
  async function importFiles(files) {
    if (!libraryReady()) return;
    $('fileInput').disabled = true;
    for (const file of files) {
      if (!/\.(xlsx|xls|csv|tsv|ods)$/i.test(file.name)) { toast('不支持的文件格式：' + file.name, true); continue; }
      if (file.size > 20 * 1024 * 1024) { toast(file.name + ' 超过 20 MB 限制。', true); continue; }
      try {
        const bytes = await file.arrayBuffer();
        const book = parseBook(bytes, file.name);
        if (!book.SheetNames.length) throw Error('没有可读取的工作表');
        const source = { id: uid(), name: file.name, bytes, book, sheet: book.SheetNames[0], configs: {}, encoding: 'utf-8' };
        initConfig(source); state.sources.push(source); state.current = source.id;
      } catch (error) { toast('无法读取 ' + file.name + '，请确认文件未加密且格式正确。', true); }
    }
    $('fileInput').disabled = false; $('fileInput').value = ''; renderSources();
    if (current()) $('mappingPanel').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  function renderSources() {
    $('fileList').innerHTML = state.sources.map(s => `<div class="file-item"><span>${icon('file')}${esc(s.name)}</span><small>${s.book.SheetNames.length} 张工作表</small><button data-source="${s.id}" aria-label="设置 ${esc(s.name)}">设置 ${icon('arrow-right')}</button></div>`).join('');
    $('mappingPanel').hidden = !state.sources.length;
    $('sourceSelect').innerHTML = state.sources.map(s => `<option value="${s.id}">${esc(s.name)}</option>`).join('');
    $('sourceSelect').value = state.current;
    $('step2').classList.toggle('active', !!state.sources.length);
    if (current()) renderMapping();
  }
  function renderMapping() {
    const s = current(); if (!s) return; initConfig(s);
    const rows = rowsFor(s), cfg = settings(s), headers = rows[cfg.header] || [], count = Object.values(cfg.mapping).filter(v => v !== '').length;
    $('sheetSelect').innerHTML = s.book.SheetNames.map(n => `<option>${esc(n)}</option>`).join(''); $('sheetSelect').value = s.sheet;
    $('headerRow').value = rowOffset(s) + cfg.header + 1; $('headerRow').max = rowOffset(s) + Math.max(1, rows.length);
    $('encodingSelect').value = s.encoding; $('encodingSelect').disabled = !/\.(csv|tsv)$/i.test(s.name);
    $('mappingNote').textContent = `已匹配 ${count} / ${state.fields.length} 个字段。请核对下方样例；不需要的列可以不导入。`;
    $('mappingGrid').innerHTML = state.fields.map(f => {
      const mapped = cfg.mapping[f.key];
      const options = headers.map((h,i) => `<option value="${i}" ${String(i) === mapped ? 'selected' : ''}>${esc(h || '未命名列')} · ${XLSX.utils.encode_col(columnOffset(s) + i)}</option>`).join('');
      const sample = mapped === '' ? '未匹配，不导入此字段' : rows.slice(cfg.header + 1).map(r => r[Number(mapped)]).find(v => C.clean(v)) || '该列暂无数据';
      return `<label class="mapping-item ${mapped === '' ? 'unmapped' : ''}"><strong>${esc(f.label)}</strong><select data-field="${f.key}" aria-label="映射${esc(f.label)}"><option value="">— 不导入 —</option>${options}</select><small title="${esc(sample)}">${esc(sample)}</small></label>`;
    }).join('');
    $('rawPreview').innerHTML = `<table><thead><tr>${headers.map((h,i) => `<th>${esc(h || '未命名列')} (${XLSX.utils.encode_col(columnOffset(s) + i)})</th>`).join('')}</tr></thead><tbody>${rows.slice(cfg.header + 1, cfg.header + 6).map(r => `<tr>${headers.map((_,i) => `<td>${esc(r[i])}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
    const already = state.imported.has(s.id + ':' + s.sheet);
    const candidates = mappedSource(s), missing = candidates.filter(C.pendingIdentity).length;
    $('mappingSummary').textContent = already ? '这张工作表已添加，可在客户资料库中继续补充。' : `识别 ${candidates.length} 条非空记录：${candidates.length - missing} 条进入全部客户，${missing} 条进入缺少主体。仅跳过整行空白。`;
    $('applyMapping').disabled = already || !candidates.length;
  }
  function addMapped() {
    const s = current(); if (!s) return;
    const key = s.id + ':' + s.sheet; if (state.imported.has(key)) return;
    const records = mappedSource(s), missing = records.filter(C.pendingIdentity).length;
    if (state.records.length + records.length > 50000) { toast('当前最多整理 50,000 条记录，请分批导出。', true); return; }
    state.records.push(...records.map(r => ({ ...r, id: uid() }))); state.imported.add(key); state.dirty = true; state.filter = missing === records.length ? 'missing' : 'all'; state.selected.clear(); state.query = ''; $('searchInput').value = ''; state.page = 1;
    renderMapping(); renderData(); toast(`已保留 ${records.length} 条记录：全部客户 ${records.length - missing} 条，缺少主体 ${missing} 条。`); document.querySelector('.data-panel').scrollIntoView({ behavior: 'smooth' });
  }
  function filtered(duplicates) { return state.records.filter(r => (state.filter === 'missing' ? C.pendingIdentity(r) : !C.pendingIdentity(r) && (state.filter === 'all' || state.filter === 'review' && C.issues(r, state.fields).length || state.filter === 'duplicate' && duplicates.has(r.id))) && (!state.query || [...state.fields.map(f => r[f.key]),r.source,...(r.sourceCells || []).map(cell => cell.value)].some(v => String(v || '').toLowerCase().includes(state.query)))); }
  function updateMetrics(duplicates) {
    const accepted = customers(), missing = state.records.length - accepted.length;
    $('totalCount').textContent = `${state.records.length} 条记录`; $('statAll').textContent = accepted.length;
    $('statMissing').textContent = missing;
    $('statReview').textContent = accepted.filter(r => C.issues(r, state.fields).length).length; $('statDuplicate').textContent = duplicates.size;
    $('exportBtn').disabled = !state.records.length;
    $('deleteBtn').disabled = !state.selected.size; $('selectedCount').textContent = state.selected.size ? `(${state.selected.size})` : '';
    $('step3').classList.toggle('active', !!state.records.length);
    document.querySelectorAll('.stat[data-filter]').forEach(b => { const active = b.dataset.filter === state.filter; b.classList.toggle('active', active); b.setAttribute('aria-pressed', String(active)); });
    $('missingNotice').hidden = !missing || state.filter === 'missing';
    $('missingNoticeText').textContent = `${missing} 条记录暂未加入全部客户，请到“缺少主体”补充并确认。`;
    $('missingGuide').hidden = state.filter !== 'missing';
  }
  function statusHTML(r, duplicates) {
    if (C.pendingIdentity(r)) return `<span class="status-pill missing">${icon('info')}${C.missingIdentity(r) ? '缺少主体' : '待确认加入'}</span><small class="status-detail">${C.missingIdentity(r) ? '缺少固定字段：' + esc(C.missingFixedFields(r).join('、')) : '已补齐三项固定字段，确认后加入全部客户。'}</small>`;
    const issues = C.issues(r, state.fields);
    return `<span class="status-pill ${issues.length ? 'review' : ''}" title="${esc(issues.join('；') || '通过基础检查，联系方式未核实')}">${icon(issues.length ? 'alert' : 'check-circle')}${issues.length ? '待完善' : '可用'}</span>${issues.length ? `<small class="status-detail">${esc(issues.join('；'))}</small>` : ''}${duplicates.has(r.id) ? `<br><span class="status-pill duplicate" title="疑似重复仅提示，请人工核对。">${icon('copy')}疑似重复</span>` : ''}`;
  }
  function actionHTML(r) { return (C.pendingIdentity(r) ? `<button class="secondary-btn admit-btn" data-admit="${r.id}" ${C.missingIdentity(r) ? 'disabled' : ''} title="客户名称、公司名、邮箱三项全部填写后，才能确认加入全部客户">${icon('plus')}加入全部客户</button>` : '') + (r.sourceCells?.length ? `<button class="text-btn" data-original="${r.id}">${icon('file')}查看原始内容</button>` : ''); }
  function sourceHTML(r) { return `${esc(r.source)}${r.sourceRow ? '<br>第 ' + r.sourceRow + ' 行' : ''}`; }
  function showOriginal(id) {
    const r = state.records.find(record => record.id === id); if (!r) return;
    $('originalMeta').textContent = `${r.source}${r.sourceRow ? ' · 第 ' + r.sourceRow + ' 行' : ''}`;
    $('originalContent').innerHTML = (r.sourceCells || []).map(cell => `<div><dt>${esc(cell.label)} · 第 ${cell.column} 列</dt><dd>${esc(cell.value)}</dd></div>`).join('');
    $('originalDialog').showModal();
  }
  function admitCustomer(id) {
    const index = state.records.findIndex(r => r.id === id); if (index < 0) return;
    try { state.records[index] = C.admitRecord(state.records[index]); state.selected.delete(id); state.dirty = true; renderData(); toast('已加入全部客户，可继续查看待完善提示。'); }
    catch (error) { toast(error.message, true); }
  }
  function renderData() {
    const duplicates = duplicatesForCustomers(), list = filtered(duplicates), pages = Math.max(1, Math.ceil(list.length / state.pageSize));
    state.page = Math.max(1, Math.min(state.page, pages)); const visible = list.slice((state.page - 1) * state.pageSize, state.page * state.pageSize);
    updateMetrics(duplicates);
    $('dataBody').innerHTML = visible.map(r => `<tr data-id="${r.id}"><td><input type="checkbox" data-select="${r.id}" aria-label="选择 ${esc(r.company || r.contact || '待补充记录')}" ${state.selected.has(r.id) ? 'checked' : ''}></td><td class="status-cell">${statusHTML(r,duplicates)}<div class="record-actions">${actionHTML(r)}</div></td>${state.fields.map(f => `<td contenteditable="true" role="textbox" spellcheck="false" aria-label="${esc(f.label)}" data-key="${f.key}" class="${C.pendingIdentity(r) && C.fixedKeys.includes(f.key) && !C.clean(r[f.key]) ? 'missing-field' : ''}">${esc(r[f.key])}</td>`).join('')}<td class="muted source-cell">${sourceHTML(r)}</td></tr>`).join('');
    $('emptyState').hidden = !!visible.length;
    $('emptyState').querySelector('h3').textContent = state.filter === 'missing' ? '没有待补充的主体记录' : state.records.length ? '没有匹配的客户' : '一份整齐的客户表，从这里开始';
    $('emptyState').querySelector('p').textContent = state.records.length ? '可切换分类或调整搜索词；缺少主体的记录需要补充并确认后才会加入全部客户。' : '导入你的第一份表格，或体验示例数据。';
    $('emptyDemoBtn').hidden = !!state.records.length;
    const first = list.length ? (state.page - 1) * state.pageSize + 1 : 0, last = Math.min(state.page * state.pageSize, list.length);
    $('resultInfo').textContent = `共保留 ${state.records.length} 条记录 · 当前分类匹配 ${list.length} 条 · 显示 ${first}–${last} 条`;
    $('pageSize').value = String(state.pageSize);
    $('pageInfo').textContent = `第 ${state.page} / ${pages} 页`; $('prevPage').disabled = state.page <= 1; $('nextPage').disabled = state.page >= pages;
    $('selectAll').checked = !!visible.length && visible.every(r => state.selected.has(r.id)); $('selectAll').indeterminate = visible.some(r => state.selected.has(r.id)) && !$('selectAll').checked;
  }
  function editCell(target) {
    const key = target.dataset.key; if (!key) return;
    const record = state.records.find(r => r.id === target.closest('tr').dataset.id); if (!record) return;
    const value = C.clean(target.textContent); if (record[key] === value) return;
    record[key] = value; state.dirty = true;
    const duplicates = duplicatesForCustomers(); updateMetrics(duplicates);
    document.querySelectorAll('#dataBody tr').forEach(tr => { const r = state.records.find(r => r.id === tr.dataset.id); tr.querySelector('.status-cell').innerHTML = statusHTML(r,duplicates) + `<div class="record-actions">${actionHTML(r)}</div>`; tr.querySelectorAll('[data-key]').forEach(cell => cell.classList.toggle('missing-field', C.pendingIdentity(r) && C.fixedKeys.includes(cell.dataset.key) && !C.clean(r[cell.dataset.key]))); });
  }
  function downloadBlob(blob, name) { const url = URL.createObjectURL(blob), a = document.createElement('a'); a.href = url; a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(url),1000); }
  function exportRows(template = false) {
    if (!libraryReady()) return;
    const records = template ? [] : state.records;
    const rows = C.exportMatrix(records, state.fields, template);
    const name = template ? '外贸客户标准模板' : '外贸客户资料_' + new Date().toLocaleDateString('sv-SE');
    try {
      if (!template && $('exportFormat').value === 'csv') {
        const csvRows = rows.map((row, i) => [...row, i === 0 ? '原始内容' : JSON.stringify(records[i - 1].sourceCells || [])]);
        downloadBlob(new Blob(['\ufeff' + csvRows.map(row => row.map(C.csvCell).join(',')).join('\r\n')], { type: 'text/csv;charset=utf-8' }), name + '.csv');
      } else {
        const book = XLSX.utils.book_new();
        const append = (matrix, title) => { const sheet = XLSX.utils.aoa_to_sheet(matrix); sheet['!cols'] = matrix[0].map(h => ({ wch: Math.max(20, Math.min(40, h.length * 2 + 8)) })); if (matrix.length > 1) sheet['!autofilter'] = {ref:sheet['!ref']}; XLSX.utils.book_append_sheet(book,sheet,title); };
        append(template ? rows : C.exportMatrix(records.filter(r => !C.pendingIdentity(r)), state.fields), '客户资料');
        const pending = records.filter(C.pendingIdentity);
        if (pending.length) append(C.exportMatrix(pending, state.fields), '缺少主体');
        const raw = C.recoveryMatrix(records);
        if (raw.length > 1) append(raw, '主体补充原始内容');
        XLSX.writeFile(book,name + '.xlsx');
      }
      if (!template) { state.dirty = false; toast(`已生成全部 ${records.length} 条记录（含缺少主体及其原始内容），请确认文件已保存。`); }
    } catch (e) { toast('导出失败，请稍后重试。', true); }
  }
  function demo() {
    if (!libraryReady()) return;
    const existing = state.sources.find(s => s.demo); if (existing) { state.current = existing.id; renderSources(); $('mappingPanel').scrollIntoView({behavior:'smooth'}); return; }
    const data = [['Company Name','Contact','E-mail','WhatsApp','Country','Product Interest','Website','Notes'],['Northstar Living','Emma Wilson','emma@example.com','+44 7700 900001','United Kingdom','竹纤维家居用品','https://example.com','展会认识；需提供产品目录'],['Hikari Trading','Yuki Tanaka','yuki@example.com','+81 90 0000 0001','Japan','可重复使用的保温杯','','关注 MOQ 和交期'],['Verde Home','Lucas Silva','lucas@example','+55 11 0000 0000','Brazil','环保厨房用品','','邮箱待核对'],['Northstar Living','Emma Wilson','emma@example.com','','United Kingdom','定制包装','','历史询盘，可能重复'],['Atelier Maison','Camille Martin','','','France','棉麻布艺','','需补充联系方式'],['Coastline Supply','Alex Chen','alex@example.com','+1 202 555 0100','United States','户外折叠家具','','索取样品报价']];
    const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(data),'展会客户');
    const s = {id:uid(),name:'示例 · 展会客户.xlsx',book,sheet:'展会客户',configs:{},encoding:'utf-8',demo:true}; initConfig(s); state.sources.push(s); state.current = s.id; renderSources(); $('mappingPanel').scrollIntoView({behavior:'smooth'}); toast('已载入示例文件，请确认映射后添加到客户列表。');
  }
  $('fileInput').addEventListener('change', e => importFiles([...e.target.files]));
  $('dropzone').addEventListener('dragover', e => { e.preventDefault(); $('dropzone').classList.add('dragover'); });
  $('dropzone').addEventListener('dragleave', () => $('dropzone').classList.remove('dragover'));
  $('dropzone').addEventListener('drop', e => { e.preventDefault(); $('dropzone').classList.remove('dragover'); importFiles([...e.dataTransfer.files]); });
  $('sourceSelect').onchange = e => { state.current = e.target.value; renderMapping(); };
  $('fileList').onclick = e => { const b = e.target.closest('[data-source]'); if (b) { state.current = b.dataset.source; renderSources(); $('mappingPanel').scrollIntoView({behavior:'smooth'}); } };
  $('sheetSelect').onchange = e => { current().sheet = e.target.value; renderMapping(); };
  $('headerRow').onchange = e => { const s = current(), rows = rowsFor(s), cfg = settings(s); cfg.header = Math.max(0, Math.min(rows.length - 1, (parseInt(e.target.value,10) || 1) - rowOffset(s) - 1)); cfg.mapping = C.autoMap(rows[cfg.header] || [], state.fields); renderMapping(); };
  $('encodingSelect').onchange = e => { const s = current(); try { const book = parseBook(s.bytes,s.name,e.target.value); s.book = book; s.sheet = book.SheetNames[0]; s.encoding = e.target.value; s.configs = {}; renderMapping(); } catch { toast('无法使用该编码解析文件。',true); } };
  $('mappingGrid').onchange = e => { if (!e.target.dataset.field) return; const mapping = settings(current()).mapping, value = e.target.value; if (value !== '' && Object.entries(mapping).some(([k,v]) => k !== e.target.dataset.field && v === value)) { toast('该来源列已匹配其他字段，请先取消原映射。',true); renderMapping(); return; } mapping[e.target.dataset.field] = value; renderMapping(); };
  $('applyMapping').onclick = addMapped;
  $('demoBtn').onclick = $('emptyDemoBtn').onclick = demo;
  $('templateBtn').onclick = () => exportRows(true); $('exportBtn').onclick = () => exportRows();
  $('searchInput').oninput = e => { state.query = e.target.value.trim().toLowerCase(); state.page = 1; renderData(); };
  document.querySelectorAll('[data-filter]').forEach(b => b.onclick = () => {state.filter = b.dataset.filter; state.selected.clear(); state.page = 1; renderData();});
  $('dataBody').addEventListener('input', e => editCell(e.target));
  $('dataBody').addEventListener('click', e => { const button = e.target.closest('[data-admit]'), original = e.target.closest('[data-original]'); if (button) admitCustomer(button.dataset.admit); if (original) showOriginal(original.dataset.original); });
  $('dataBody').addEventListener('focusout', e => {
    if (!e.target.dataset.key) return;
    const record = state.records.find(r => r.id === e.target.closest('tr')?.dataset.id);
    if (record && C.missingIdentity(record)) { record.pendingIdentity = true; state.selected.delete(record.id); }
    setTimeout(() => { if (!$('dataBody').contains(document.activeElement)) renderData(); }, 0);
  });
  $('dataBody').addEventListener('paste', e => { if (!e.target.dataset.key) return; e.preventDefault(); const text = e.clipboardData.getData('text/plain'); const selection = window.getSelection(); if (!selection.rangeCount) return; const range = selection.getRangeAt(0); range.deleteContents(); const node = document.createTextNode(text); range.insertNode(node); range.setStartAfter(node); range.collapse(true); selection.removeAllRanges(); selection.addRange(range); editCell(e.target); });
  $('dataBody').addEventListener('keydown', e => { if (e.key === 'Enter' && e.target.dataset.key) {e.preventDefault(); e.target.blur();} });
  $('dataBody').addEventListener('change', e => { const id = e.target.dataset.select; if (id) { e.target.checked ? state.selected.add(id) : state.selected.delete(id); renderData(); } });
  $('selectAll').onchange = e => { filtered(duplicatesForCustomers()).slice((state.page - 1) * state.pageSize,state.page * state.pageSize).forEach(r => e.target.checked ? state.selected.add(r.id) : state.selected.delete(r.id)); renderData(); };
  $('addRowBtn').onclick = () => { state.records.unshift({ id:uid(),...Object.fromEntries(state.fields.map(f => [f.key,''])),source:'手动新增',sourceRow:'',pendingIdentity:true,sourceCells:[] }); state.dirty = true; state.page = 1; state.filter = 'missing'; state.selected.clear(); state.query = ''; $('searchInput').value = ''; renderData(); $('dataBody').querySelector('[data-key=company]').focus(); toast('请先补齐客户名称、公司名、邮箱，再点击“加入全部客户”。'); };
  $('deleteBtn').onclick = () => { if (!confirm(`删除选中的 ${state.selected.size} 条客户？该操作无法撤销。`)) return; state.records = state.records.filter(r => !state.selected.has(r.id)); state.selected.clear(); state.dirty = true; renderData(); };
  $('prevPage').onclick = () => {state.page--;renderData();}; $('nextPage').onclick = () => {state.page++;renderData();};
  $('pageSize').onchange = e => { const size = Number(e.target.value); if (![10,25,50,100].includes(size)) return; state.pageSize = size; state.page = 1; renderData(); };
  $('helpBtn').onclick = () => $('helpDialog').showModal(); $('closeHelp').onclick = () => $('helpDialog').close();
  $('closeOriginal').onclick = () => $('originalDialog').close();
  window.addEventListener('beforeunload', e => { if (state.dirty && state.records.length) { e.preventDefault(); e.returnValue = ''; } });
  $('fieldForm').onsubmit = addCustomField;
  $('fieldList').onclick = e => { const button = e.target.closest('[data-remove-field]'); if (button) removeCustomField(button.dataset.removeField); };
  $('reviewRuleList').innerHTML = C.reviewRules.map(rule => `<li><strong>${esc(rule.label)}</strong><span>${esc(rule.description)}</span></li>`).join('');
  renderFields(); renderData();
})();
