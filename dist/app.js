/* Customer records remain in memory. No analytics, storage, or upload requests. */
(() => {
  'use strict';
  const C = window.TradeCore, $ = id => document.getElementById(id);
  const state = { sources: [], current: '', records: [], selected: new Set(), imported: new Set(), filter: 'all', query: '', page: 1, pageSize: 25, dirty: false };
  let sequence = 0, toastTimer;
  const uid = () => String(++sequence);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
  function toast(message, error = false) { $('toast').textContent = message; $('toast').className = 'toast' + (error ? ' error' : ''); $('toast').hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => $('toast').hidden = true, 5000); }
  function libraryReady() { if (window.XLSX) return true; toast('Excel 解析库未加载，请确认 vendor 文件完整后刷新。', true); return false; }
  function current() { return state.sources.find(s => s.id === state.current); }
  function settings(source) { return source.configs[source.sheet]; }
  function rowsFor(source) { return XLSX.utils.sheet_to_json(source.book.Sheets[source.sheet], { header: 1, defval: '', raw: false, blankrows: true }); }
  function initConfig(source) { if (!source.configs[source.sheet]) { const rows = rowsFor(source), header = C.detectHeader(rows); source.configs[source.sheet] = { header, mapping: C.autoMap(rows[header] || []) }; } }
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
    $('fileList').innerHTML = state.sources.map(s => `<div class="file-item"><span>▤ ${esc(s.name)}</span><small>${s.book.SheetNames.length} 张工作表</small><button data-source="${s.id}" aria-label="设置 ${esc(s.name)}">设置 →</button></div>`).join('');
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
    $('headerRow').value = cfg.header + 1; $('headerRow').max = Math.max(1, rows.length);
    $('encodingSelect').value = s.encoding; $('encodingSelect').disabled = !/\.(csv|tsv)$/i.test(s.name);
    $('mappingNote').textContent = `已匹配 ${count} / ${C.fields.length} 个字段。请核对下方样例；不需要的列可以不导入。`;
    $('mappingGrid').innerHTML = C.fields.map(f => {
      const mapped = cfg.mapping[f.key];
      const options = headers.map((h,i) => `<option value="${i}" ${String(i) === mapped ? 'selected' : ''}>${esc(h || '未命名列')} · ${XLSX.utils.encode_col(i)}</option>`).join('');
      const sample = mapped === '' ? '未匹配，不导入此字段' : rows.slice(cfg.header + 1).map(r => r[Number(mapped)]).find(v => C.clean(v)) || '该列暂无数据';
      return `<label class="mapping-item ${mapped === '' ? 'unmapped' : ''}"><strong>${f.label}</strong><select data-field="${f.key}" aria-label="映射${f.label}"><option value="">— 不导入 —</option>${options}</select><small title="${esc(sample)}">${esc(sample)}</small></label>`;
    }).join('');
    $('rawPreview').innerHTML = `<table><thead><tr>${headers.map((h,i) => `<th>${esc(h || '未命名列')} (${XLSX.utils.encode_col(i)})</th>`).join('')}</tr></thead><tbody>${rows.slice(cfg.header + 1, cfg.header + 6).map(r => `<tr>${headers.map((_,i) => `<td>${esc(r[i])}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
    const already = state.imported.has(s.id + ':' + s.sheet);
    const candidates = C.mappedRows(rows, cfg.header, cfg.mapping, s.name + ' / ' + s.sheet);
    $('mappingSummary').textContent = already ? '这张工作表已添加，可在客户列表中继续修改。' : `将添加 ${candidates.length} 条非空客户记录`;
    $('applyMapping').disabled = already || !count || !candidates.length;
  }
  function addMapped() {
    const s = current(); if (!s) return;
    const key = s.id + ':' + s.sheet; if (state.imported.has(key)) return;
    const cfg = settings(s);
    const records = C.mappedRows(rowsFor(s), cfg.header, cfg.mapping, s.name + ' / ' + s.sheet);
    if (state.records.length + records.length > 50000) { toast('当前最多整理 50,000 条记录，请分批导出。', true); return; }
    state.records.push(...records.map(r => ({ ...r, id: uid() }))); state.imported.add(key); state.dirty = true; state.filter = 'all'; state.query = ''; $('searchInput').value = ''; state.page = 1;
    renderMapping(); renderData(); toast(`已添加 ${records.length} 条客户记录`); document.querySelector('.data-panel').scrollIntoView({ behavior: 'smooth' });
  }
  function filtered(duplicates) { return state.records.filter(r => (state.filter === 'all' || state.filter === 'review' && C.issues(r).length || state.filter === 'duplicate' && duplicates.has(r.id)) && (!state.query || [...C.fields.map(f => r[f.key]),r.source].some(v => String(v || '').toLowerCase().includes(state.query)))); }
  function updateMetrics(duplicates) {
    $('totalCount').textContent = $('statAll').textContent = state.records.length;
    $('statReview').textContent = state.records.filter(r => C.issues(r).length).length; $('statDuplicate').textContent = duplicates.size;
    $('exportBtn').disabled = !state.records.length;
    $('deleteBtn').disabled = !state.selected.size; $('selectedCount').textContent = state.selected.size ? `(${state.selected.size})` : '';
    $('step3').classList.toggle('active', !!state.records.length);
    document.querySelectorAll('.stat[data-filter]').forEach(b => b.classList.toggle('active', b.dataset.filter === state.filter));
  }
  function statusHTML(r, duplicates) { const issues = C.issues(r); return `<span class="status-pill ${issues.length ? 'review' : ''}" title="${esc(issues.join('；') || '资料可用')}">${issues.length ? '待完善' : '可用'}</span>${duplicates.has(r.id) ? '<br><span class="status-pill duplicate" title="邮箱相同，或公司与联系人相同；请人工核对。">疑似重复</span>' : ''}`; }
  function renderData() {
    const duplicates = C.duplicateIds(state.records), list = filtered(duplicates), pages = Math.max(1, Math.ceil(list.length / state.pageSize));
    state.page = Math.min(state.page, pages); const visible = list.slice((state.page - 1) * state.pageSize, state.page * state.pageSize);
    updateMetrics(duplicates);
    $('dataBody').innerHTML = visible.map(r => `<tr data-id="${r.id}"><td><input type="checkbox" data-select="${r.id}" aria-label="选择 ${esc(r.company || r.contact || '客户')}" ${state.selected.has(r.id) ? 'checked' : ''}></td><td class="status-cell">${statusHTML(r,duplicates)}</td>${C.fields.map(f => `<td contenteditable="true" role="textbox" spellcheck="false" aria-label="${f.label}" data-key="${f.key}">${esc(r[f.key])}</td>`).join('')}<td class="muted">${esc(r.source)}${r.sourceRow ? '<br>第 ' + r.sourceRow + ' 行' : ''}</td></tr>`).join('');
    $('emptyState').hidden = !!visible.length;
    $('emptyState').querySelector('h3').textContent = state.records.length ? '没有匹配的客户' : '一份整齐的客户表，从这里开始';
    $('emptyState').querySelector('p').textContent = state.records.length ? '试试其他搜索词，或切换到全部客户。' : '导入你的第一份表格，或体验示例数据。';
    $('emptyDemoBtn').hidden = !!state.records.length;
    $('resultInfo').textContent = `共 ${state.records.length} 条客户 · 当前匹配 ${list.length} 条 · 每页 ${state.pageSize} 条`;
    $('pageInfo').textContent = `${state.page} / ${pages}`; $('prevPage').disabled = state.page <= 1; $('nextPage').disabled = state.page >= pages;
    $('selectAll').checked = !!visible.length && visible.every(r => state.selected.has(r.id)); $('selectAll').indeterminate = visible.some(r => state.selected.has(r.id)) && !$('selectAll').checked;
  }
  function editCell(target) {
    const key = target.dataset.key; if (!key) return;
    const record = state.records.find(r => r.id === target.closest('tr').dataset.id); if (!record) return;
    const value = C.clean(target.textContent); if (record[key] === value) return;
    record[key] = value; state.dirty = true;
    const duplicates = C.duplicateIds(state.records); updateMetrics(duplicates);
    document.querySelectorAll('#dataBody tr').forEach(tr => { const r = state.records.find(r => r.id === tr.dataset.id); tr.querySelector('.status-cell').innerHTML = statusHTML(r,duplicates); });
  }
  function downloadBlob(blob, name) { const url = URL.createObjectURL(blob), a = document.createElement('a'); a.href = url; a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(url),1000); }
  function exportRows(template = false) {
    if (!libraryReady()) return;
    const headers = C.fields.map(f => f.label), records = template ? [] : state.records;
    if (!template) headers.push('来源文件 / 工作表','原始行号','校对提示');
    const duplicates = C.duplicateIds(records);
    const rows = [headers, ...records.map(r => [...C.fields.map(f => String(r[f.key] ?? '')),r.source,String(r.sourceRow || ''),[...C.issues(r),...(duplicates.has(r.id) ? ['疑似重复'] : [])].join('；')])];
    const name = template ? '外贸客户标准模板' : '外贸客户资料_' + new Date().toLocaleDateString('sv-SE');
    try {
      if (!template && $('exportFormat').value === 'csv') downloadBlob(new Blob(['\ufeff' + rows.map(row => row.map(C.csvCell).join(',')).join('\r\n')], { type: 'text/csv;charset=utf-8' }), name + '.csv');
      else { const book = XLSX.utils.book_new(), sheet = XLSX.utils.aoa_to_sheet(rows); sheet['!cols'] = headers.map((h,i) => ({ wch: [24,18,32,22,18,32,30,32,35,12,30][i] })); if (records.length) sheet['!autofilter'] = {ref:sheet['!ref']}; XLSX.utils.book_append_sheet(book,sheet,'客户资料'); XLSX.writeFile(book,name + '.xlsx'); }
      if (!template) { state.dirty = false; toast(`已生成全部 ${records.length} 条客户记录，请确认下载文件已保存。`); }
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
  $('headerRow').onchange = e => { const s = current(), rows = rowsFor(s), cfg = settings(s); cfg.header = Math.max(0, Math.min(rows.length - 1, (parseInt(e.target.value,10) || 1) - 1)); cfg.mapping = C.autoMap(rows[cfg.header] || []); renderMapping(); };
  $('encodingSelect').onchange = e => { const s = current(); try { const book = parseBook(s.bytes,s.name,e.target.value); s.book = book; s.sheet = book.SheetNames[0]; s.encoding = e.target.value; s.configs = {}; renderMapping(); } catch { toast('无法使用该编码解析文件。',true); } };
  $('mappingGrid').onchange = e => { if (!e.target.dataset.field) return; const mapping = settings(current()).mapping, value = e.target.value; if (value !== '' && Object.entries(mapping).some(([k,v]) => k !== e.target.dataset.field && v === value)) { toast('该来源列已匹配其他字段，请先取消原映射。',true); renderMapping(); return; } mapping[e.target.dataset.field] = value; renderMapping(); };
  $('applyMapping').onclick = addMapped;
  $('demoBtn').onclick = $('emptyDemoBtn').onclick = demo;
  $('templateBtn').onclick = () => exportRows(true); $('exportBtn').onclick = () => exportRows();
  $('searchInput').oninput = e => { state.query = e.target.value.trim().toLowerCase(); state.page = 1; renderData(); };
  document.querySelectorAll('[data-filter]').forEach(b => b.onclick = () => {state.filter = b.dataset.filter; state.page = 1; renderData();});
  $('dataBody').addEventListener('input', e => editCell(e.target));
  $('dataBody').addEventListener('paste', e => { if (!e.target.dataset.key) return; e.preventDefault(); const text = e.clipboardData.getData('text/plain'); const selection = window.getSelection(); if (!selection.rangeCount) return; const range = selection.getRangeAt(0); range.deleteContents(); const node = document.createTextNode(text); range.insertNode(node); range.setStartAfter(node); range.collapse(true); selection.removeAllRanges(); selection.addRange(range); editCell(e.target); });
  $('dataBody').addEventListener('keydown', e => { if (e.key === 'Enter' && e.target.dataset.key) {e.preventDefault(); e.target.blur();} });
  $('dataBody').addEventListener('change', e => { const id = e.target.dataset.select; if (id) { e.target.checked ? state.selected.add(id) : state.selected.delete(id); renderData(); } });
  $('selectAll').onchange = e => { filtered(C.duplicateIds(state.records)).slice((state.page - 1) * state.pageSize,state.page * state.pageSize).forEach(r => e.target.checked ? state.selected.add(r.id) : state.selected.delete(r.id)); renderData(); };
  $('addRowBtn').onclick = () => { state.records.unshift({ id:uid(),...Object.fromEntries(C.fields.map(f => [f.key,''])),source:'手动新增',sourceRow:'' }); state.dirty = true; state.page = 1; state.filter = 'all'; state.query = ''; $('searchInput').value = ''; renderData(); $('dataBody').querySelector('[data-key=company]').focus(); };
  $('deleteBtn').onclick = () => { if (!confirm(`删除选中的 ${state.selected.size} 条客户？该操作无法撤销。`)) return; state.records = state.records.filter(r => !state.selected.has(r.id)); state.selected.clear(); state.dirty = true; renderData(); };
  $('prevPage').onclick = () => {state.page--;renderData();}; $('nextPage').onclick = () => {state.page++;renderData();};
  $('helpBtn').onclick = () => $('helpDialog').showModal(); $('closeHelp').onclick = () => $('helpDialog').close();
  window.addEventListener('beforeunload', e => { if (state.dirty && state.records.length) { e.preventDefault(); e.returnValue = ''; } });
  renderData();
})();
