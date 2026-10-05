/* Seif Dashboard 2026 - v2 */
const API = 'https://script.google.com/macros/s/AKfycbz7BA8nNTZPp2cQSudw1mVpFcqrLwJc1NQpys6G3wIXNFAt5dHsjjU_eViJf_rqFZ2U/exec';
const $ = (s, r = document) => r.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = s => String(s ?? '').trim().toLowerCase();
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const CFG = {
  'Codes SHR':    { start: 3, blocks: [['Shrinkage', 1, 2], ['Scheduled SHR', 4, 5]] },
  'Codes absent': { start: 3, blocks: [['Absent', 1, 2], ['Scheduled Absent', 5, 6]] },
  'STR':          { start: 2, blocks: [['STR', 1, 9]] }
};
const PERMS = { run: 'Run', export: 'Export', codes: 'Edit Codes', str: 'Edit STR', users: 'Manage Users' };
let USER = '', PWD = '', ME = null, S = {}, FINAL = [], busyN = 0, LAST_AOA = null;
const can = p => ME && ME.perms.includes(p);
const isAdmin = () => ME && ME.role === 'admin';

/* ---------- helpers ---------- */
function toast(m, ok = true) {
  const d = document.createElement('div'); d.textContent = m; if (!ok) d.className = 'err';
  $('#toast').append(d); setTimeout(() => d.remove(), 3500);
}
async function api(b) {
  $('#busy').hidden = false; busyN++;
  try {
    const r = await fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ user: USER, pwd: PWD, ...b }) });
    const j = await r.json();
    if (!j.ok) throw new Error(j.error || 'Error');
    return j;
  } finally { if (--busyN <= 0) $('#busy').hidden = true; }
}
const loadSheet = async n => { S[n] = (await api({ action: 'read', sheet: n })).values; };
function pickFile(cb) {
  const i = document.createElement('input'); i.type = 'file'; i.accept = '.xlsx,.xls,.csv';
  i.onchange = async () => { if (!i.files[0]) return; cb(XLSX.read(await i.files[0].arrayBuffer(), { type: 'array' })); };
  i.click();
}
const aoaOf = wb => {
  const nm = wb.SheetNames.find(n => norm(n) === 'rd') || wb.SheetNames[0];
  const ws = wb.Sheets[nm], rg = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  rg.s.c = 0; rg.s.r = 0;
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true, range: rg });
};
const toDate = v => {
  if (typeof v === 'number') return Math.trunc(v);
  if (v instanceof Date) return Math.trunc(v / 864e5 + 25569);
  const t = String(v ?? '').trim();
  if (/^\d+(\.\d+)?$/.test(t)) return Math.trunc(+t);
  if (/^\d{1,4}[\/\-.]\d{1,2}[\/\-.]\d{1,4}/.test(t)) {
    const p = Date.parse(t.replace(/\./g, '/'));
    if (!isNaN(p)) { const d = new Date(p); return Math.round(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 864e5 + 25569); }
  }
  return null;
};
const pad = n => String(n).padStart(2, '0');
const fmtT = s => `${Math.floor(s / 3600)}:${pad(Math.floor(s % 3600 / 60))}:${pad(s % 60)}`;
const serialDate = n => new Date((n - 25569) * 864e5);
const fmtD = n => { const d = serialDate(n); return `${d.getUTCDate()}-${MON[d.getUTCMonth()]}`; };
const colIdx = l => l.toUpperCase().split('').reduce((a, c) => a * 26 + c.charCodeAt(0) - 64, 0) - 1;
function toSec(v) {
  if (typeof v === 'number') return Math.round(v * 86400);
  const m = String(v).trim().match(/^(\d+):(\d{1,2})(?::(\d{1,2}))?$/);
  return m ? +m[1] * 3600 + +m[2] * 60 + +(m[3] || 0) : null;
}

/* ---------- دخول ---------- */
async function login(auto) {
  if (!auto) { USER = $('#usr').value.trim(); PWD = $('#pwd').value; }
  try {
    ME = await api({ action: 'login' });
    sessionStorage.setItem('c', JSON.stringify([USER, PWD]));
    await start();
  } catch (e) { sessionStorage.clear(); $('#login').hidden = false; $('#app').hidden = true; toast(e.message, false); }
}
async function start() {
  $('#who').textContent = `${ME.user} • ${ME.role}`;
  const t = [['dash', 'Dashboard']];
  if (can('codes')) t.push(['Codes SHR', 'Codes SHR'], ['Codes absent', 'Codes absent']);
  if (can('str')) t.push(['STR', 'STR']);
  if (can('users')) t.push(['users', 'Users']);
  $('#tabs').innerHTML = t.map(([k, v]) => `<button class="btn sm b-ref" data-t="${k}">${v}</button>`).join('');
  $('#upSch').hidden = !can('run'); $('#gear').hidden = !isAdmin();
  $('#expView').hidden = !can('export'); $('#expAll').hidden = !(isAdmin() && can('export'));
  $('#tlLbl').hidden = !isAdmin(); $('#settings').hidden = true;
  await Promise.all(Object.keys(CFG).map(loadSheet));
  $('#login').hidden = true; $('#app').hidden = false; show('dash');
}
$('#loginBtn').onclick = () => login();
['#usr', '#pwd'].forEach(s => $(s).addEventListener('keydown', e => { if (e.key === 'Enter') login(); }));
$('#logout').onclick = () => { sessionStorage.clear(); location.reload(); };
const saved = sessionStorage.getItem('c');
if (saved) { [USER, PWD] = JSON.parse(saved); login(true); }

$('#tabs').onclick = e => { const t = e.target.dataset.t; if (t) show(t); };
function show(t) {
  document.querySelectorAll('#tabs [data-t]').forEach(b => b.classList.toggle('active', b.dataset.t === t));
  $('#v-dash').hidden = t !== 'dash'; $('#v-ed').hidden = t === 'dash';
  if (t === 'users') renderUsers().catch(e => toast(e.message, false));
  else if (t !== 'dash') renderEditor(t);
}
$('#gear').onclick = () => { $('#settings').hidden = !$('#settings').hidden; };

/* ---------- معالجة الاسكدول (منطق الـ VBA) ---------- */
function process(aoa) {
  const hdr = +$('#hdr').value, dc = colIdx($('#dcol').value || 'J');
  const mk = (sh, ci, si) => { const m = new Map(); S[sh].slice(2).forEach(r => { const k = norm(r[ci]); if (k && !m.has(k)) m.set(k, String(r[si] || '').trim()); }); return m; };
  const shrA = mk('Codes SHR', 0, 1), shrB = mk('Codes SHR', 3, 4), absA = mk('Codes absent', 0, 1), absB = mk('Codes absent', 4, 5);
  const look = (a, b, k) => a.get(k) || b.get(k) || '#N/A';
  const str = new Map(); S.STR.slice(1).forEach(r => { const k = norm(r[0]); if (k) str.set(k, { agent: r[2], tl: r[4], ss: r[6], spv: r[8] }); });
  const out = []; let id = '', date = '';
  for (let i = hdr + 1; i < aoa.length; i++) {
    const r = aoa[i] || [], b = String(r[1] ?? '').trim(), c = r[2];
    if (b) { id = b.slice(0, 6).trim(); continue; }
    const dt = toDate(c);
    if (dt !== null) { date = dt; continue; }
    const code = String(c ?? '').trim(); if (!code || !id || !date) continue;
    const sec = toSec(r[dc]); if (sec == null) continue;
    const u = str.get(norm(id)) || {}, k = norm(code), d = serialDate(date);
    out.push({ id, date, sec, code, agent: u.agent ?? '#N/A', tl: u.tl ?? '#N/A', ss: u.ss ?? '#N/A', spv: u.spv ?? '#N/A',
      shr: look(shrA, shrB, k), abs: look(absA, absB, k), month: `${MON[d.getUTCMonth()]}-${String(d.getUTCFullYear()).slice(2)}` });
  }
  FINAL = out;
  const miss = [...new Set(out.filter(x => x.abs === '#N/A' || x.shr === '#N/A').map(x => x.code))];
  $('#info').innerHTML = `Processed <b>${out.length}</b> rows.` + (miss.length ? ` <span style="color:#ffb4bd">Codes not found (${miss.length}): ${esc(miss.join(', '))}</span>` : ' ✔');
  if (!out.length) {
    toast('No rows matched the expected layout', false);
    const s0 = Math.max(hdr - 2, 0);
    $('#info').innerHTML += `<br>No rows matched. First rows of the file:<div class="wrap" dir="ltr" style="margin-top:8px"><table><thead><tr><th>row</th><th>A</th><th>B</th><th>C</th><th>D</th><th>${esc($('#dcol').value.toUpperCase())}</th></tr></thead><tbody>` +
      aoa.slice(s0, hdr + 14).map((r, k) => `<tr><td>${s0 + k + 1}</td>${[0, 1, 2, 3, dc].map(j => `<td>${esc(String(r[j] ?? '').slice(0, 28))}</td>`).join('')}</tr>`).join('') + '</tbody></table></div>';
  }
  const prev = $('#tlf').value, tls = [...new Set(out.filter(x => x.abs === 'Absent').map(x => x.tl))].sort();
  $('#tlf').innerHTML = '<option value="all">All Leaders</option>' + tls.map(t => `<option>${esc(t)}</option>`).join('');
  if (tls.includes(prev)) $('#tlf').value = prev;
  renderFinal(); renderPivot();
}
const vis = () => FINAL.filter(x => x.abs === 'Absent' && ($('#tlf').value === 'all' || x.tl === $('#tlf').value));
function renderFinal() {
  const rows = vis(), H = ['ID', 'Date', 'Duration', 'Agent Name', 'TL', 'Code', 'Month', 'SPV'];
  const cell = v => `<td class="${v === '#N/A' ? 'na' : ''}">${esc(v)}</td>`;
  $('#cnt').textContent = rows.length ? `(${rows.length} rows)` : '';
  $('#final').innerHTML = `<table><thead><tr>${H.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>` +
    rows.slice(0, 500).map(x => `<tr><td>${esc(x.id)}</td><td>${fmtD(x.date)}</td><td>${fmtT(x.sec)}</td>${cell(x.agent)}${cell(x.tl)}<td>${esc(x.code)}</td><td>${x.month}</td>${cell(x.spv)}</tr>`).join('') +
    `</tbody></table>` + (rows.length > 500 ? `<p class="muted" style="padding:8px">Showing first 500 of ${rows.length} - export includes all.</p>` : '');
}
const ORDER = ['yasmin khaled', 'support sls'];
const spvOrder = (a, b) => { const r = x => { const i = ORDER.indexOf(norm(x)); return x === '#N/A' ? 99 : i < 0 ? 50 : i; }; return r(a) - r(b) || String(a).localeCompare(b); };
function renderPivot() {
  const act = $('#stf').value === 'Active', g = new Map(), ag = new Set();
  FINAL.filter(x => !act || x.ss === 'Active').forEach(x => {
    if (!g.has(x.spv)) g.set(x.spv, new Map());
    const m = g.get(x.spv); if (!m.has(x.tl)) m.set(x.tl, { a: 0, s: 0 });
    const o = m.get(x.tl);
    if (x.abs === 'Absent') { o.a += x.sec; ag.add(x.id); } else if (x.abs === 'Scheduled') o.s += x.sec;
  });
  const pc = (a, t) => t ? (a / t * 100).toFixed(2) + '%' : '-';
  const tr = (n, o, cls) => `<tr class="${cls}"><td>${esc(n)}</td><td>${fmtT(o.a)}</td><td>${fmtT(o.s)}</td><td>${fmtT(o.a + o.s)}</td><td>${pc(o.a, o.a + o.s)}</td></tr>`;
  let h = '', G = { a: 0, s: 0 };
  [...g.keys()].sort(spvOrder).forEach(spv => {
    const m = g.get(spv), t = { a: 0, s: 0 };
    m.forEach(o => { t.a += o.a; t.s += o.s; }); G.a += t.a; G.s += t.s;
    h += tr(spv, t, 'spv') + [...m.keys()].sort().map(tl => tr(tl, m.get(tl), '')).join('');
  });
  h += tr('Total Scheduled Time', G, 'tot');
  $('#pivot').innerHTML = `<table><thead><tr><th>TL</th><th>Absent</th><th>Scheduled</th><th>Total Scheduled Time</th><th>%</th></tr></thead><tbody>${h}</tbody></table>`;
}
$('#upSch').onclick = () => pickFile(wb => { LAST_AOA = aoaOf(wb); process(LAST_AOA); toast('Processed'); });
$('#stf').onchange = () => FINAL.length && renderPivot();
$('#tlf').onchange = renderFinal;
['#hdr', '#dcol'].forEach(s => $(s).onchange = () => LAST_AOA && process(LAST_AOA));
function exportRows(rows, name) {
  if (!rows.length) return toast('No data to export', false);
  const aoa = [['ID', 'Date', 'Duration', 'Agent Name', 'TL', 'Code'], ...rows.map(x => [/^\d+$/.test(x.id) ? +x.id : x.id, x.date, x.sec / 86400, x.agent, x.tl, x.code])];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  for (let i = 2; i <= aoa.length; i++) { ws['B' + i].z = 'd-mmm'; ws['C' + i].z = '[h]:mm:ss'; }
  ws['!cols'] = [10, 10, 11, 34, 18, 24].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Final'); XLSX.writeFile(wb, name);
}
$('#expView').onclick = () => exportRows(vis(), 'Final_Absent.xlsx');
$('#expAll').onclick = () => exportRows(FINAL, 'Final_All.xlsx');

/* ---------- محرر الشيتات (ميرور لجوجل شيت) ---------- */
function renderEditor(name) {
  const cfg = CFG[name], root = $('#v-ed'); root.innerHTML = '';
  cfg.blocks.forEach(([title, c1, c2]) => {
    const head = (S[name][cfg.start - 2] || []).slice(c1 - 1, c2);
    const card = document.createElement('div'); card.className = 'card';
    card.innerHTML = `<div class="bar"><h3>${esc(name)} - ${esc(title)}</h3>
      <div class="row"><input class="q" placeholder="Search..."><button class="btn sm b-up up">Upload & Replace</button><button class="btn sm b-ref rf">Refresh</button></div></div>
      <div class="add">${head.map(h => `<input placeholder="${esc(h)}">`).join('')}<button class="btn sm b-add ad">Add</button></div><div class="wrap" dir="ltr"></div>`;
    root.append(card);
    const wrap = $('.wrap', card);
    const draw = () => {
      const q = norm($('.q', card).value); let n = 0, rows = '';
      for (let r = cfg.start - 1; r < S[name].length; r++) {
        const v = S[name][r].slice(c1 - 1, c2);
        if (v.every(x => x === '') || (q && !v.some(x => norm(x).includes(q)))) continue;
        if (++n > 300) break;
        rows += `<tr>${v.map((x, i) => `<td><input value="${esc(x)}" data-r="${r + 1}" data-c="${c1 + i}"></td>`).join('')}<td><button class="btn sm b-del" data-del="${r + 1}">Delete</button></td></tr>`;
      }
      wrap.innerHTML = `<table><thead><tr>${head.map(h => `<th>${esc(h)}</th>`).join('')}<th></th></tr></thead><tbody>${rows}</tbody></table>` + (n > 300 ? '<p class="muted" style="padding:8px">Showing first 300 - use search.</p>' : '');
    };
    draw();
    const reload = async () => { await loadSheet(name); renderEditor(name); };
    $('.q', card).oninput = draw;
    $('.rf', card).onclick = () => reload().then(() => toast('Refreshed')).catch(e => toast(e.message, false));
    wrap.onchange = async e => {
      const t = e.target; if (!t.dataset.r) return;
      const r = +t.dataset.r; S[name][r - 1][+t.dataset.c - 1] = t.value;
      try { await api({ action: 'setCells', sheet: name, row: r, c1, values: S[name][r - 1].slice(c1 - 1, c2) }); toast('Saved'); } catch (er) { toast(er.message, false); }
    };
    wrap.onclick = async e => {
      const r = e.target.dataset.del; if (!r || !confirm('Delete this row?')) return;
      try { await api({ action: 'delRow', sheet: name, row: +r, c1, c2 }); await reload(); toast('Deleted'); } catch (er) { toast(er.message, false); }
    };
    $('.ad', card).onclick = async () => {
      const vals = [...card.querySelectorAll('.add input')].map(i => i.value.trim());
      if (!vals[0]) return toast('Fill at least the first column', false);
      try { await api({ action: 'addRow', sheet: name, start: cfg.start, c1, values: vals }); await reload(); toast('Added'); } catch (er) { toast(er.message, false); }
    };
    $('.up', card).onclick = () => pickFile(async wb => {
      const w = c2 - c1 + 1;
      const vals = aoaOf(wb).slice(1).map(r => Array.from({ length: w }, (_, i) => r[i] ?? '')).filter(r => r.some(x => x !== ''));
      if (!vals.length) return toast('File is empty', false);
      if (!confirm(`Replace ${title} with ${vals.length} rows. Continue?`)) return;
      try { await api({ action: 'replace', sheet: name, start: cfg.start, c1, c2, values: vals }); await reload(); toast('Uploaded'); } catch (er) { toast(er.message, false); }
    });
  });
}

/* ---------- المستخدمين والصلاحيات ---------- */
async function renderUsers() {
  const { users } = await api({ action: 'listUsers' }), root = $('#v-ed');
  root.innerHTML = `<div class="card slim"><h3>User</h3>
    <div class="add"><input id="u-n" placeholder="Username"><input id="u-p" type="password" placeholder="Password (blank = keep)">
      <select id="u-r"><option value="user">User</option><option value="admin">Admin</option></select>
      <select id="u-a"><option value="YES">Active</option><option value="NO">Disabled</option></select></div>
    <div class="row c" id="u-perms">${Object.entries(PERMS).map(([k, v]) => `<label><input type="checkbox" value="${k}" ${k === 'run' || k === 'export' ? 'checked' : ''}> ${v}</label>`).join('')}</div>
    <p class="muted c">Admin automatically gets all permissions.</p>
    <div class="row c"><button class="btn b-save" id="u-s">Save</button><button class="btn b-ref" id="u-c">New</button></div></div>
    <div class="card slim"><h3>All Users</h3><div class="wrap" dir="ltr"><table><thead><tr><th>User</th><th>Role</th><th>Permissions</th><th>Active</th><th></th></tr></thead><tbody>${
      users.map(u => `<tr><td>${esc(u.username)}</td><td>${esc(u.role)}</td><td>${esc(u.role === 'admin' ? 'all' : u.perms.join(', '))}</td><td>${u.active}</td>
      <td><button class="btn sm b-save" data-e="${esc(u.username)}">Edit</button> <button class="btn sm b-del" data-d="${esc(u.username)}">Delete</button></td></tr>`).join('') || '<tr><td colspan="5">No users yet</td></tr>'}</tbody></table></div></div>`;
  const fill = u => {
    $('#u-n').value = u ? u.username : ''; $('#u-p').value = ''; $('#u-r').value = u ? u.role : 'user'; $('#u-a').value = u ? u.active : 'YES';
    document.querySelectorAll('#u-perms input').forEach(c => c.checked = u ? u.perms.includes(c.value) : (c.value === 'run' || c.value === 'export'));
  };
  $('#u-c').onclick = () => fill(null);
  $('#u-s').onclick = async () => {
    const u = { username: $('#u-n').value.trim(), password: $('#u-p').value, role: $('#u-r').value, active: $('#u-a').value,
      perms: [...document.querySelectorAll('#u-perms input:checked')].map(c => c.value) };
    try { await api({ action: 'saveUser', u }); toast('Saved'); renderUsers(); } catch (e) { toast(e.message, false); }
  };
  root.onclick = async e => {
    const en = e.target.dataset.e, dn = e.target.dataset.d;
    if (en) fill(users.find(u => u.username === en));
    if (dn && confirm('Delete ' + dn + '?')) { try { await api({ action: 'delUser', username: dn }); toast('Deleted'); renderUsers(); } catch (er) { toast(er.message, false); } }
  };
}

/* Copy Dash: image of the Absenteeism % card -> clipboard */
$('#copyDash').onclick = async () => {
  if (!FINAL.length) return toast('Nothing to copy yet', false);
  try {
    const cv = await html2canvas($('#pivotCard'), { backgroundColor: '#27104d', scale: 2, onclone: d => {
      const w = d.querySelector('#pivot'); w.style.maxHeight = 'none'; w.style.overflow = 'visible'; d.querySelector('#copyDash').style.display = 'none'; } });
    const blob = await new Promise(r => cv.toBlob(r, 'image/png'));
    try { await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]); toast('Dashboard copied - paste it anywhere'); }
    catch { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'Absenteeism.png'; a.click(); toast('Clipboard blocked - image downloaded instead'); }
  } catch (e) { toast('Copy failed: ' + e.message, false); }
};
