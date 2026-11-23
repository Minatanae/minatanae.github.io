(() => {
'use strict';
const KEY = 'fallkvist.timeline.v1';
const $ = s => document.querySelector(s);
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));
const pad = n => String(n).padStart(2, '0');
const todayISO = d => `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;

/* ---------- config ---------- */
// Entra ID app registration (SPA, redirect URI brk-multihub://fallkvist.se, delegated Files.ReadWrite.All)
const TEAMS_CLIENT_ID = '';
const GRAPH_SCOPES = ['Files.ReadWrite.All'];
const FOLDER = 'Timeline';
const IN_TEAMS = new URLSearchParams(location.search).get('host') === 'teams';

/* ---------- dialogs ---------- */
function ask(text, okLabel = 'OK', danger = false, noLabel = 'Avbryt'){
  return new Promise(res => {
    $('#dlgText').textContent = text;
    const yes = $('#dlgYes'), no = $('#dlgNo');
    yes.textContent = okLabel; no.textContent = noLabel; yes.style.background = danger ? 'var(--danger)' : ''; yes.style.borderColor = danger ? 'var(--danger)' : '';
    $('#dialog').classList.remove('hidden'); setTimeout(() => yes.focus(), 20);
    const done = v => { $('#dialog').classList.add('hidden'); yes.onclick = no.onclick = null; res(v); };
    yes.onclick = () => done(true); no.onclick = () => done(false);
  });
}
let toastTimer;
function notify(text){
  const t = $('#toast'); t.textContent = text; t.classList.remove('hidden');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.add('hidden'), 5000);
}
function setStatus(s){
  const el = $('#saveState'); if (!IN_TEAMS) { el.textContent = ''; return; }
  el.className = s === 'error' ? 'error' : '';
  el.textContent = s === 'saving' ? 'Sparar…' : s === 'saved' ? 'Sparat' : s === 'error' ? 'Kunde inte spara' : '';
}
function normalize(p){
  p.steps = Array.isArray(p.steps) ? p.steps : [];
  // details only stack upward: older projects may have a 'down' chain, fold it into 'up'
  p.steps.forEach(s => { s.up = [...(s.up || []), ...(s.down || [])]; s.down = []; });
  // checklists were removed (detail cards cover that), drop any leftovers
  p.steps.forEach(s => [s, ...s.up].forEach(x => { delete x.checklist; }));
  return p;
}

/* ---------- storage: this browser ---------- */
const LocalStore = {
  async init(){
    let d = null;
    try { const raw = localStorage.getItem(KEY); if (raw) d = JSON.parse(raw); } catch(e) {}
    $('#storeNote').textContent = 'Projekten sparas bara i den här webbläsaren. Använd Exportera för att flytta dem.';
    return (d && Array.isArray(d.projects) ? d.projects : []).map(normalize);
  },
  write(){ try { localStorage.setItem(KEY, JSON.stringify(db)); } catch(e) { notify('Kunde inte spara i webbläsaren.'); } },
  saveProject(){ this.write(); },
  deleteProject(){ this.write(); },
  pending: () => false,
  refresh: async () => false
};

/* ---------- storage: Teams channel files (SharePoint via Microsoft Graph) ---------- */
function loadScript(src){
  return new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.crossOrigin = 'anonymous'; s.onload = res; s.onerror = () => rej(new Error('Kunde inte ladda ' + src)); document.head.appendChild(s); });
}
function applyTheme(t){ document.documentElement.dataset.theme = (t === 'dark' || t === 'contrast') ? 'dark' : 'light'; }

const TeamsStore = {
  meta: new Map(),     // projectId -> { itemId, eTag }
  timers: new Map(), inflight: new Map(),
  async init(){
    await Promise.all([
      loadScript('https://cdn.jsdelivr.net/npm/@microsoft/teams-js@2.59.0/dist/umd/MicrosoftTeams.min.js'),
      loadScript('https://cdn.jsdelivr.net/npm/@azure/msal-browser@5.27.0/lib/msal-browser.min.js')
    ]);
    try { await microsoftTeams.app.initialize(); }
    catch(e){ throw new Error('Den här adressen fungerar bara inuti Microsoft Teams.\nI webbläsaren: öppna fallkvist.se/timeline utan ?host=teams.'); }
    const c = await microsoftTeams.app.getContext();
    applyTheme(c.app && c.app.theme);
    microsoftTeams.app.registerOnThemeChangeHandler(applyTheme);
    microsoftTeams.app.notifySuccess();
    if (!TEAMS_CLIENT_ID) throw new Error('Timeline är inte kopplad till Microsoft ännu (app-ID saknas).');
    if (!c.team || !c.team.groupId || !c.channel || !c.channel.id) throw new Error('Lägg till Timeline som en flik i en kanal i ett team.');
    this.pca = await msal.createNestablePublicClientApplication({
      auth: { clientId: TEAMS_CLIENT_ID, authority: `https://login.microsoftonline.com/${c.user.tenant.id}` }
    });
    const ff = await this.graph(`/teams/${c.team.groupId}/channels/${encodeURIComponent(c.channel.id)}/filesFolder`);
    this.drive = ff.parentReference.driveId; this.root = ff.id;
    $('#storeNote').textContent = `Projekten sparas i kanalens Filer, i mappen ${FOLDER}. Alla i kanalen ser samma projekt.`;
    $('#exportBtn').classList.add('hidden');
    setInterval(() => this.refresh(), 20000);
    document.addEventListener('visibilitychange', () => { if (!document.hidden) this.refresh(); });
    window.addEventListener('beforeunload', e => { if (this.pending()) { e.preventDefault(); e.returnValue = ''; } });
    const raw = await this.listRaw();
    return (await Promise.all(raw.map(i => this.download(i).catch(() => null)))).filter(Boolean);
  },
  async token(){
    const req = { scopes: GRAPH_SCOPES };
    try { return (await this.pca.acquireTokenSilent(req)).accessToken; }
    catch(e){ return (await this.pca.acquireTokenPopup(req)).accessToken; }
  },
  async graph(path, opts = {}){
    const t = await this.token();
    const r = await fetch('https://graph.microsoft.com/v1.0' + path, { ...opts, headers: { Authorization: 'Bearer ' + t, ...(opts.headers || {}) } });
    if (!r.ok){ const e = new Error(`Microsoft Graph svarade ${r.status}`); e.status = r.status; throw e; }
    return r.status === 204 ? null : r.json();
  },
  path(rest){ return `/drives/${this.drive}/items/${this.root}:/${FOLDER}${rest}`; },
  async listRaw(){
    try {
      const res = await this.graph(this.path(':/children?$top=999'));
      return res.value.filter(i => i.file && i.name.endsWith('.json'));
    } catch(e){ if (e.status === 404) return []; throw e; }
  },
  async download(item){
    const r = await fetch(item['@microsoft.graph.downloadUrl']);
    if (!r.ok) throw new Error('download ' + r.status);
    const p = normalize(await r.json());
    if (!p.id || typeof p.name !== 'string') return null;
    this.meta.set(p.id, { itemId: item.id, eTag: item.eTag });
    return p;
  },
  pending(){ return this.timers.size > 0 || this.inflight.size > 0; },
  saveProject(p){
    clearTimeout(this.timers.get(p.id)); setStatus('saving');
    this.timers.set(p.id, setTimeout(() => this.flush(p.id), 600));
  },
  async flush(id){
    this.timers.delete(id);
    if (this.inflight.has(id)){ this.inflight.set(id, 'again'); return; }
    const p = db.projects.find(x => x.id === id); if (!p) return;
    this.inflight.set(id, true);
    const m = this.meta.get(id);
    try {
      const item = await this.graph(this.path(`/${id}.json:/content`) + (m ? '' : '?@microsoft.graph.conflictBehavior=fail'), {
        method: 'PUT', body: JSON.stringify(p),
        headers: { 'Content-Type': 'application/json', ...(m ? { 'If-Match': m.eTag } : {}) }
      });
      this.meta.set(id, { itemId: item.id, eTag: item.eTag });
      if (!this.timers.size && this.inflight.size <= 1) setStatus('saved');
    } catch(e){
      if (e.status === 412 || e.status === 409){
        await this.reloadOne(id);
        notify('Någon annan ändrade projektet samtidigt. Den senaste versionen visas nu, gör om din ändring.');
        setStatus('');
      } else {
        setStatus('error');
        this.timers.set(id, setTimeout(() => this.flush(id), 8000));
      }
    } finally {
      const again = this.inflight.get(id) === 'again';
      this.inflight.delete(id);
      if (again) this.flush(id);
    }
  },
  async reloadOne(id){
    try {
      const item = await this.graph(this.path(`/${id}.json`));
      const p = await this.download(item);
      const i = db.projects.findIndex(x => x.id === id);
      if (p && i >= 0) db.projects[i] = p;
    } catch(e){}
    rerender();
  },
  async deleteProject(p){
    clearTimeout(this.timers.get(p.id)); this.timers.delete(p.id);
    const m = this.meta.get(p.id); if (!m) return;
    try { await this.graph(`/drives/${this.drive}/items/${m.itemId}`, { method: 'DELETE' }); this.meta.delete(p.id); }
    catch(e){ notify('Kunde inte ta bort projektet i Teams.'); }
  },
  async refresh(){
    if (this.pending() || ctx || !$('#modal').classList.contains('hidden')) return;
    let raw; try { raw = await this.listRaw(); } catch(e){ return; }
    if (this.pending() || ctx) return;
    let changed = false;
    const seen = new Set();
    for (const item of raw){
      const id = item.name.slice(0, -5); seen.add(id);
      const m = this.meta.get(id);
      if (m && m.eTag === item.eTag) continue;
      try {
        const p = await this.download(item); if (!p) continue;
        const i = db.projects.findIndex(x => x.id === p.id);
        if (i >= 0) db.projects[i] = p; else db.projects.push(p);
        changed = true;
      } catch(e){}
    }
    for (const id of [...this.meta.keys()]){
      if (!seen.has(id)){ this.meta.delete(id); db.projects = db.projects.filter(x => x.id !== id); changed = true; }
    }
    if (changed && !this.pending() && !ctx) rerender();
  }
};

const store = IN_TEAMS ? TeamsStore : LocalStore;
let db = { projects: [] };
function save(p = current()){
  if (!p) return;
  if (p.id === hist.pid && hist.snap){ hist.undo.push(hist.snap); if (hist.undo.length > 50) hist.undo.shift(); hist.redo = []; }
  p.updated = Date.now();
  store.saveProject(p);
  if (p.id === hist.pid) hist.snap = JSON.stringify(p);
  updateUndoButtons();
}
function rerender(){
  if (currentId && !current()){ location.hash = ''; return; }
  resetHistory();   // someone else changed the data, old undo steps no longer apply
  current() ? renderBoard() : renderHome();
}

/* ---------- undo / redo ---------- */
const hist = { undo: [], redo: [], snap: null, pid: null };
function resetHistory(){
  const p = current();
  hist.undo = []; hist.redo = []; hist.pid = p ? p.id : null; hist.snap = p ? JSON.stringify(p) : null;
  updateUndoButtons();
}
function updateUndoButtons(){ $('#undoBtn').disabled = !hist.undo.length; $('#redoBtn').disabled = !hist.redo.length; }
function restoreSnap(json){
  const np = normalize(JSON.parse(json));
  const i = db.projects.findIndex(x => x.id === np.id); if (i < 0) return;
  np.updated = Date.now(); db.projects[i] = np;
  store.saveProject(np); hist.snap = JSON.stringify(np);
  closeDrawer(); renderBoard(); updateUndoButtons();
}
function undo(){ const p = current(); if (!p || !hist.undo.length) return; hist.redo.push(JSON.stringify(p)); restoreSnap(hist.undo.pop()); notify('Ångrat'); }
function redo(){ const p = current(); if (!p || !hist.redo.length) return; hist.undo.push(JSON.stringify(p)); restoreSnap(hist.redo.pop()); notify('Gjort om'); }

let currentId = null;
const current = () => db.projects.find(p => p.id === currentId) || null;

/* ---------- helpers ---------- */
const MONTHS = ['jan','feb','mar','apr','maj','jun','jul','aug','sep','okt','nov','dec'];
function fmtDate(s){
  if (!s.date) return '';
  const [y,m,d] = s.date.split('-').map(Number);
  const thisYear = new Date().getFullYear();
  let out = `${d} ${MONTHS[m-1]}` + (y !== thisYear ? ` ${y}` : '');
  return out;
}
function lateCheck(s){
  if (!s.date || s.done) return false;
  const due = new Date(`${s.date}T23:59`);
  return due < new Date();
}
const people = s => (s.who || '').split(',').map(x => x.trim()).filter(Boolean);
const addDays = (iso, n) => { const [y,m,d] = iso.split('-').map(Number); return todayISO(new Date(y, m-1, d + n)); };
const setRadio = (name, v) => { const r = document.querySelector(`input[name="${name}"][value="${v}"]`); if (r) r.checked = true; };
const getRadio = name => { const r = document.querySelector(`input[name="${name}"]:checked`); return r ? r.value : null; };
// all items except the start node (used for counts)
const allItems = p => { const out = []; p.steps.forEach((s, i) => { if (i > 0) out.push(s); out.push(...s.up); }); return out; };

// Dates should run left to right, and a detail shouldn't end after its main step
function dateWarnings(p){
  const w = new Map(); let maxPrev = null;
  p.steps.forEach(s => {
    if (s.date){
      if (maxPrev && s.date < maxPrev) w.set(s, 'Datumet ligger före ett tidigare steg på tidslinjen.');
      if (!maxPrev || s.date > maxPrev) maxPrev = s.date;
    }
    s.up.forEach(d => { if (d.date && s.date && d.date > s.date) w.set(d, 'Detaljsteget har ett senare datum än sitt huvudsteg.'); });
  });
  return w;
}

/* ---------- per-viewer UI prefs (collapsed stacks) ---------- */
const UIKEY = 'fallkvist.timeline.ui';
let ui = {}; try { ui = JSON.parse(localStorage.getItem(UIKEY)) || {}; } catch(e){}
ui.collapsed = ui.collapsed || {};
const isCollapsed = s => { const o = ui.collapsed[s.id]; return o === undefined ? s.up.length > 4 : o; };
function setCollapsed(s, v){ ui.collapsed[s.id] = v; try { localStorage.setItem(UIKEY, JSON.stringify(ui)); } catch(e){} }

let filterPerson = null;
function colorFor(name){
  let h = 0; for (const c of name) h = (h * 31 + c.charCodeAt(0)) % 360;
  return `hsl(${h} 55% 48%)`;
}
const initials = n => n.split(/\s+/).map(w => w[0]).join('').slice(0,2).toUpperCase();
function el(tag, cls, text){ const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }

/* ---------- routing ---------- */
function route(){
  const m = location.hash.match(/^#p=(.+)$/);
  currentId = m ? decodeURIComponent(m[1]) : null;
  if (currentId && !current()) { location.hash = ''; return; }
  closeDrawer(true);
  filterPerson = null; resetHistory();
  if (current()) { $('#home').classList.add('hidden'); $('#board').classList.remove('hidden'); renderBoard(true); }
  else { $('#board').classList.add('hidden'); $('#home').classList.remove('hidden'); renderHome(); }
}
window.addEventListener('hashchange', route);

/* ---------- home ---------- */
function renderHome(){
  const list = $('#plist'); list.innerHTML = '';
  if (!db.projects.length){
    const e = el('div','empty-home'); e.style.gridColumn = '1/-1';
    e.textContent = 'Inga projekt än. Skapa ditt första med knappen ovan.';
    list.appendChild(e); return;
  }
  [...db.projects].sort((a,b) => (b.updated||0) - (a.updated||0)).forEach(p => {
    const c = el('button','pcard');
    c.appendChild(el('h3', null, p.name));
    if (p.description) c.appendChild(el('p', null, p.description));
    const total = p.steps.length + p.steps.reduce((n,s) => n + s.up.length + s.down.length, 0);
    const done = p.steps.filter(s => s.done).length + p.steps.reduce((n,s) => n + s.up.filter(d=>d.done).length + s.down.filter(d=>d.done).length, 0);
    const meta = el('div','meta');
    meta.appendChild(el('span', null, `${p.steps.length} steg`));
    if (total) meta.appendChild(el('span', null, `${done}/${total} klara`));
    if (p.updated) meta.appendChild(el('span', null, 'Ändrad ' + new Date(p.updated).toLocaleDateString('sv-SE')));
    if (total){ const pr = el('div','progress'); const i = el('i'); i.style.width = (done/total*100) + '%'; pr.appendChild(i); c.appendChild(pr); }
    c.appendChild(meta);
    c.onclick = () => { location.hash = 'p=' + encodeURIComponent(p.id); };
    list.appendChild(c);
  });
}

/* ---------- project modal ---------- */
let modalMode = 'new';
function openModal(mode){
  modalMode = mode;
  const p = current();
  $('#mTitle').textContent = mode === 'new' ? 'Nytt projekt' : 'Redigera projekt';
  $('#mSave').textContent = mode === 'new' ? 'Skapa' : 'Spara';
  $('#mDelete').classList.toggle('hidden', mode === 'new');
  $('#mDup').classList.toggle('hidden', mode === 'new');
  $('#mName').value = mode === 'edit' ? p.name : '';
  $('#mDesc').value = mode === 'edit' ? (p.description || '') : '';
  $('#modal').classList.remove('hidden');
  setTimeout(() => $('#mName').focus(), 20);
}
const closeModal = () => $('#modal').classList.add('hidden');
$('#newProject').onclick = () => openModal('new');
$('#ptitle').onclick = () => openModal('edit');
$('#mCancel').onclick = closeModal;
$('#modal').addEventListener('pointerdown', e => { if (e.target.id === 'modal') closeModal(); });
$('#mForm').onsubmit = e => {
  e.preventDefault();
  const name = $('#mName').value.trim(); if (!name) return;
  const description = $('#mDesc').value.trim();
  if (modalMode === 'new'){
    const p = { id: uid(), name, description, created: Date.now(), updated: Date.now(), steps: [] };
    db.projects.push(p); save(p); closeModal();
    location.hash = 'p=' + encodeURIComponent(p.id);
  } else {
    Object.assign(current(), { name, description }); save(); closeModal(); renderBoard();
  }
};
$('#mDelete').onclick = async () => {
  const p = current();
  const msg = IN_TEAMS ? `Ta bort projektet "${p.name}"? Filen hamnar i SharePoints papperskorg.` : `Ta bort projektet "${p.name}"? Det går inte att ångra.`;
  if (!await ask(msg, 'Ta bort', true)) return;
  db.projects = db.projects.filter(x => x.id !== p.id); store.deleteProject(p); closeModal(); location.hash = '';
};

$('#mDup').onclick = async () => {
  const src = current(); closeModal();
  const start = src.steps[0] && src.steps[0].date;
  const shift = start ? await ask('Ska datumen i kopian flyttas så att projektet startar idag?', 'Flytta datum', false, 'Behåll datum') : false;
  const cp = JSON.parse(JSON.stringify(src));
  cp.id = uid(); cp.name = src.name + ' (kopia)'; cp.created = Date.now();
  let delta = 0;
  if (shift){ const [y,m,d] = start.split('-').map(Number); const t = new Date(); t.setHours(0,0,0,0); delta = Math.round((t - new Date(y, m-1, d)) / 86400000); }
  const fresh = (x, isStart) => {
    x.id = uid(); x.done = !!isStart; x.doing = false;
    if (x.date && delta) x.date = addDays(x.date, delta);
  };
  cp.steps.forEach((s, i) => { fresh(s, i === 0); s.up.forEach(d => fresh(d, false)); });
  normalize(cp); db.projects.push(cp); save(cp);
  location.hash = 'p=' + encodeURIComponent(cp.id);
  notify('Kopian är skapad och alla steg är nollställda. Klicka på namnet för att byta namn.');
};

/* ---------- import / export ---------- */
$('#exportBtn').onclick = () => {
  const p = current();
  const blob = new Blob([JSON.stringify({ app:'timeline', version:1, projects:[p] }, null, 2)], { type:'application/json' });
  const a = el('a'); a.href = URL.createObjectURL(blob);
  a.download = (p.name.replace(/[^\wåäöÅÄÖ-]+/g,'-') || 'projekt') + '.json';
  a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 1000);
};
$('#importBtn').onclick = () => $('#importFile').click();
$('#importFile').onchange = async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try {
    const data = JSON.parse(await f.text());
    const list = Array.isArray(data.projects) ? data.projects : [data];
    let n = 0;
    for (const p of list){
      if (!p || typeof p.name !== 'string' || !Array.isArray(p.steps)) continue;
      if (!p.id || db.projects.some(x => x.id === p.id)) p.id = uid();
      normalize(p); db.projects.push(p); save(p); n++;
    }
    renderHome(); notify(`${n} projekt importerade.`);
  } catch(err){ notify('Filen kunde inte läsas.'); }
};

/* ---------- layout constants ---------- */
const SPACING = 310, MAIN_W = 232, MAIN_H = 96, DET_W = 186, DET_H = 76, GAP = 34;
const BIG = 52, SMALL = 26, INLINE = 22, RULER_H = 64, PILL_H = 34;

/* ---------- board render ---------- */
let bounds = { x0:0, y0:0, x1:0, y1:0 };
let curWarns = new Map();
function renderBoard(fit){
  const p = current(); if (!p) return;
  $('#ptitle strong').textContent = p.name;
  $('#ptitle span').textContent = p.description || '';
  document.title = p.name + ' · Timeline';
  renderPersonFilter(p);
  renderSummary(p);

  const nodes = $('#nodes'), svg = $('#lines');
  nodes.innerHTML = ''; svg.innerHTML = '';
  const NS = 'http://www.w3.org/2000/svg';
  const line = (x1,y1,x2,y2,cls) => { const l = document.createElementNS(NS,'line'); l.setAttribute('x1',x1); l.setAttribute('y1',y1); l.setAttribute('x2',x2); l.setAttribute('y2',y2); if (cls) l.setAttribute('class',cls); svg.appendChild(l); };
  const plus = (x,y,size,cls,title,fn) => { const b = el('button','plus '+cls,'+'); b.style.cssText = `left:${x-size/2}px;top:${y-size/2}px;width:${size}px;height:${size}px`; b.title = title; b.onclick = e => { e.stopPropagation(); fn(); }; nodes.appendChild(b); return b; };
  bounds = { x0:-150, y0:-120, x1:150, y1:120 };
  const grow = (x,y) => { bounds.x0 = Math.min(bounds.x0,x); bounds.x1 = Math.max(bounds.x1,x); bounds.y0 = Math.min(bounds.y0,y); bounds.y1 = Math.max(bounds.y1,y); };

  if (!p.steps.length){
    plus(0, 0, 72, 'big', 'Starta projektet', () => openDrawer({ mode:'new', kind:'start' }));
    const lab = el('div','start-label','Klicka för att starta projektet'); lab.style.left = '-100px'; lab.style.top = '50px'; nodes.appendChild(lab);
    $('#ruler').classList.add('hidden'); rulerItems = []; todayEl = null;
    if (fit) fitView(); else apply(); return;
  }

  curWarns = dateWarnings(p);
  p.steps.forEach((s, i) => {
    const x = i * SPACING;
    const goal = isGoal(p, i), ms = isMilestone(p, i);
    nodes.appendChild(card(s, x, 0, MAIN_W, MAIN_H, i === 0 ? 'start' : goal ? 'goal' : ms ? 'milestone' : '', { type:'main', i }));
    grow(x - MAIN_W/2, -MAIN_H/2); grow(x + MAIN_W/2, MAIN_H/2);

    // connector to next / end plus
    if (i < p.steps.length - 1){
      const xa = x + MAIN_W/2, xb = (i+1)*SPACING - MAIN_W/2;
      line(xa, 0, xb, 0, s.done ? 'done' : '');
      const hit = el('div','seg-hit'); hit.style.cssText = `left:${xa}px;top:-15px;width:${xb-xa}px`; nodes.appendChild(hit);
      const b = plus((xa+xb)/2 - xa, 15, INLINE, 'inline', 'Lägg till steg här emellan', () => openDrawer({ mode:'new', kind:'main', at:i+1 }));
      hit.appendChild(b);
    } else if (!goal) {
      const xe = x + SPACING;
      line(x + MAIN_W/2, 0, xe - BIG/2, 0, 'dash');
      plus(xe, 0, BIG, 'big', 'Lägg till nästa steg', () => openDrawer({ mode:'new', kind:'main', at:i+1 }));
      grow(xe + BIG/2, 0);
    }

    // detail stack (upward only, collapsible)
    let cur = -MAIN_H/2;
    if (s.up.length && isCollapsed(s)){
      const cy = cur - GAP - PILL_H/2;
      line(x, cur, x, cy + PILL_H/2, s.up.every(d => d.done) ? 'done' : '');
      nodes.appendChild(stackPill(s, x, cy, i));
      grow(x - DET_W/2, cy - PILL_H/2);
      cur = cy - PILL_H/2;
    } else {
      s.up.forEach((d, di) => {
        const cy = cur - GAP - DET_H/2;
        line(x, cur, x, cy + DET_H/2, d.done ? 'done' : '');
        nodes.appendChild(card(d, x, cy, DET_W, DET_H, 'detail', { type:'detail', i, dir:'up', di }));
        grow(x - DET_W/2, cy - DET_H/2); grow(x + DET_W/2, cy + DET_H/2);
        cur = cy - DET_H/2;
      });
    }
    const py = cur - (GAP - 4 + SMALL/2);
    line(x, cur, x, py + SMALL/2, 'dash');
    plus(x, py, SMALL, 'small', 'Lägg till detaljsteg', () => openDrawer({ mode:'new', kind:'detail', i, dir:'up' }));
    if (s.up.length >= 3 && !isCollapsed(s)){
      const b = el('button','stack-toggle','Fäll ihop');
      b.style.cssText = `left:${x + SMALL/2 + 8}px;top:${py - 11}px`;
      b.onclick = e => { e.stopPropagation(); setCollapsed(s, true); renderBoard(); };
      nodes.appendChild(b);
    }
    grow(x, py - SMALL/2);
  });
  $('#ruler').classList.remove('hidden');
  buildRuler(p);
  if (todayX !== null) line(todayX, bounds.y0 - 30, todayX, bounds.y1 + 30, 'today');
  if (fit) fitView(); else apply();
}

function stackPill(s, x, y, i){
  const n = s.up.length, done = s.up.filter(d => d.done).length;
  const late = s.up.some(lateCheck), crit = s.up.some(d => d.critical && !d.done);
  const faded = filterPerson && !s.up.some(d => people(d).includes(filterPerson));
  const b = el('button', 'stack-pill' + (late ? ' late' : crit ? ' crit' : '') + (faded ? ' faded' : ''));
  b.style.cssText = `left:${x - DET_W/2}px;top:${y - PILL_H/2}px;width:${DET_W}px;height:${PILL_H}px`;
  b.textContent = `+${n} detaljsteg · ${done}/${n} klara` + (late ? ' · försenat' : '');
  b.title = 'Visa detaljstegen'; b.dataset.pill = i;
  b.onclick = e => { e.stopPropagation(); setCollapsed(s, false); renderBoard(); };
  return b;
}

function renderSummary(p){
  const box = $('#summary'); box.innerHTML = '';
  const items = allItems(p); if (!items.length) return;
  const chip = (t, c) => box.appendChild(el('span', 'sum ' + c, t));
  const done = items.filter(x => x.done).length, late = items.filter(lateCheck).length, doing = items.filter(x => x.doing && !x.done).length;
  chip(`${done}/${items.length} klara`, done === items.length ? 'ok all' : 'ok');
  if (late) chip(`${late} försenade`, 'bad');
  if (doing) chip(`${doing} pågår`, 'doing');
  const today = todayISO(new Date());
  const next = items.filter(x => !x.done && x.date && x.date >= today).sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : 0)[0];
  if (next) chip(`Nästa: ${next.title}, ${fmtDate(next)}`, 'next');
}

function renderPersonFilter(p){
  const sel = $('#personSel'), names = new Set();
  p.steps.forEach(s => [s, ...s.up].forEach(x => people(x).forEach(n => names.add(n))));
  if (filterPerson && !names.has(filterPerson)) filterPerson = null;
  sel.innerHTML = '';
  const o0 = el('option', null, 'Alla personer'); o0.value = ''; sel.appendChild(o0);
  [...names].sort((a, b) => a.localeCompare(b, 'sv')).forEach(n => { const o = el('option', null, n); o.value = n; sel.appendChild(o); });
  sel.value = filterPerson || '';
  sel.classList.toggle('active', !!filterPerson);
  sel.classList.toggle('hidden', !names.size);
}
function setFilter(n){ filterPerson = n || null; renderBoard(); }
$('#personSel').onchange = e => setFilter(e.target.value);

/* ---------- date ruler (follows the nodes horizontally) ---------- */
let rulerItems = [], todayEl = null, todayX = null;
function buildRuler(p){
  const tr = $('#rulerTrack'); tr.innerHTML = ''; rulerItems = []; todayEl = null; todayX = null;
  const thisYear = new Date().getFullYear();
  p.steps.forEach((s, i) => {
    const x = i * SPACING, late = i > 0 && lateCheck(s);
    const b = el('button', 'rmark' + (s.date ? '' : ' none') + (s.done ? ' done' : '') + (late ? ' late' : '') + (s.critical && !s.done ? ' crit' : '') + (isGoal(p, i) ? ' goal' : '') + (isMilestone(p, i) ? ' ms' : '') + (s.doing && !s.done ? ' doing' : '') + (curWarns.has(s) ? ' rwarn' : ''));
    if (s.date){
      const [y,m,d] = s.date.split('-').map(Number);
      b.appendChild(el('span','rd', `${d} ${MONTHS[m-1]}` + (y !== thisYear ? ` ${y}` : '')));
      b.appendChild(el('span','rw', WEEKDAYS[new Date(y, m-1, d).getDay()]));
    } else b.appendChild(el('span','rd','inget datum'));
    b.title = s.title;
    b.onclick = () => centerOn(x, 0);
    tr.appendChild(b); rulerItems.push({ el: b, x });
  });
  // "today" marker, interpolated between the dated steps around now
  const now = Date.now();
  const dated = p.steps.map((s, i) => s.date ? { i, t: new Date(`${s.date}T12:00`).getTime() } : null).filter(Boolean);
  for (let k = 0; k < dated.length - 1; k++){
    const a = dated[k], b = dated[k+1];
    if (b.t > a.t && a.t <= now && now <= b.t){ todayX = (a.i + (b.i - a.i) * (now - a.t) / (b.t - a.t)) * SPACING; break; }
  }
  if (todayX !== null){ todayEl = el('div','rtoday'); todayEl.appendChild(el('span', null, 'Idag')); tr.appendChild(todayEl); }
}
function placeRuler(){
  for (const r of rulerItems) r.el.style.left = (view.x + r.x * view.s) + 'px';
  if (todayEl) todayEl.style.left = (view.x + todayX * view.s) + 'px';
  $('#ruler').classList.toggle('compact', SPACING * view.s < 120);
}
function centerOn(x, y){
  view.x = vp.clientWidth / 2 - x * view.s;
  view.y = (vp.clientHeight - RULER_H) / 2 - y * view.s;
  apply();
}

const TROPHY = '<svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M7 2h10v2h3.5v3.5A4.5 4.5 0 0 1 16.3 12a5 5 0 0 1-3.3 2.8V17h3v2.5H8V17h3v-2.2A5 5 0 0 1 7.7 12 4.5 4.5 0 0 1 3.5 7.5V4H7V2zm0 4H5.5v1.5A2.5 2.5 0 0 0 7 9.8V6zm10 0v3.8a2.5 2.5 0 0 0 1.5-2.3V6H17z"/><rect x="6.5" y="20.5" width="11" height="2" rx="1"/></svg>';
// the goal can only be the last main step (and never the start node)
const isGoal = (p, i) => !!p.steps[i].goal && i === p.steps.length - 1 && i > 0;
// milestones: any main step except the start node and the goal
const isMilestone = (p, i) => !!p.steps[i].milestone && i > 0 && !isGoal(p, i);
const SPARK = '<svg viewBox="0 0 10 10" fill="currentColor"><path d="M5 0l1.2 3.8L10 5 6.2 6.2 5 10 3.8 6.2 0 5l3.8-1.2z"/></svg>';
function addSparkles(c, spots){
  // spots: [left%, top%, size px, delay s]
  spots.forEach(([l, t, sz, dl]) => {
    const sp = el('span','spark'); sp.innerHTML = SPARK;
    sp.style.cssText = `left:${l}%;top:${t}%;width:${sz}px;height:${sz}px;animation-delay:${dl}s`;
    c.appendChild(sp);
  });
}

const WEEKDAYS = ['sön','mån','tis','ons','tor','fre','lör'];
function relDays(s){
  const [y,m,d] = s.date.split('-').map(Number);
  const t = new Date(); t.setHours(0,0,0,0);
  const n = Math.round((new Date(y, m-1, d) - t) / 86400000);
  if (n === 0) return 'Idag'; if (n === 1) return 'Imorgon'; if (n === -1) return 'Igår';
  return n > 0 ? `om ${n} d` : `${-n} d sen`;
}
function card(s, cx, cy, w, h, cls, ref){
  const late = cls !== 'start' && lateCheck(s);
  const crit = s.critical && !s.done;
  const doing = s.doing && !s.done;
  const warn = curWarns.get(s);
  const faded = filterPerson && !people(s).includes(filterPerson);
  const c = el('button', ['card', cls, s.done && 'done', doing && 'doing', crit && 'crit', late && 'late', faded && 'faded'].filter(Boolean).join(' '));
  c.style.cssText = `left:${cx-w/2}px;top:${cy-h/2}px;width:${w}px;height:${h}px`;

  const body = el('div','cbody');
  const meta = el('div','meta');
  if (cls === 'goal'){
    const g = el('span','goal-tag'); g.innerHTML = TROPHY; g.title = 'Mål'; meta.appendChild(g);
    addSparkles(c, [[86,10,15,0],[68,66,11,.8],[92,56,12,1.5],[54,8,10,1.1],[78,36,9,.4]]);
  } else if (cls === 'milestone'){
    const g = el('span','ms-tag'); g.innerHTML = TROPHY; g.title = 'Milstolpe'; meta.appendChild(g);
    addSparkles(c, [[84,14,11,.3],[70,64,8,1.2],[93,54,9,.7],[60,10,7,1.6]]);
  }
  if (crit){ const wr = el('span','warn','⚠'); wr.title = 'Kritiskt'; meta.appendChild(wr); }
  if (s.date){
    const [y,m,d] = s.date.split('-').map(Number);
    meta.appendChild(el('span', 'dt' + (warn ? ' dt-warn' : ''), `${d} ${MONTHS[m-1]}` + (y !== new Date().getFullYear() ? ` ${y}` : '')));
    if (warn){ const dw = el('span','dwarn','!'); dw.title = warn; meta.appendChild(dw); }
    if (late) meta.appendChild(el('span','late-tag','Försenad'));
    else if (!s.done && cls !== 'start') meta.appendChild(el('span','rel', relDays(s)));
  } else if (ref.type === 'main'){
    meta.appendChild(el('span','nodate','Inget datum'));
  }
  if (meta.childNodes.length) body.appendChild(meta);
  if (s.done){
    const b = el('span','done-badge'); b.title = 'Klar';
    b.innerHTML = '<svg viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2.5 6.2l2.3 2.3 4.7-5"/></svg>';
    c.appendChild(b);
  } else if (doing){
    const b = el('span','doing-badge'); b.title = 'Pågår'; b.appendChild(el('i')); c.appendChild(b);
  }
  body.appendChild(el('div','title', s.title));
  const ppl = people(s);
  if (ppl.length){
    const wr = el('div','who');
    ppl.slice(0,3).forEach(n => {
      const ch = el('span','chip' + (filterPerson === n ? ' on' : '')); const b = el('b', null, initials(n)); b.style.background = colorFor(n); ch.append(b, n);
      ch.title = filterPerson === n ? 'Visa alla igen' : `Visa bara ${n}s steg`;
      ch.onclick = e => { e.stopPropagation(); setFilter(filterPerson === n ? null : n); };
      wr.appendChild(ch);
    });
    if (ppl.length > 3) wr.appendChild(el('span','chip', '+' + (ppl.length - 3)));
    body.appendChild(wr);
  }
  c.appendChild(body);
  c.title = (s.date ? fmtDate(s) + '\n' : '') + s.title + (warn ? `\n⚠ ${warn}` : '') + (s.note ? `\n\n${s.note}` : '');
  c.onclick = e => { e.stopPropagation(); if (suppressClick){ suppressClick = false; return; } openDrawer({ mode:'edit', ...ref }); };
  if (ref.type === 'main') c.dataset.main = ref.i;
  // every node except the start node can be dragged
  if (!(ref.type === 'main' && ref.i === 0)){ c.classList.add('draggable'); c.addEventListener('pointerdown', e => startDrag(e, c, ref)); }
  return c;
}

/* ---------- drag & drop ---------- */
// Drop rules: above a step's column -> detail of that step (position by height),
// onto the middle of a step's card -> detail on top of its stack,
// anywhere else along the main line -> main step in that gap.
let drag = null, suppressClick = false, autoPanRaf = null;
const toWorld = (cx, cy) => ({ x: (cx - view.x) / view.s, y: (cy - view.y) / view.s });
const detailCenterY = k => -(MAIN_H/2 + GAP + DET_H/2 + k * (GAP + DET_H));

function startDrag(e, c, ref){
  if (e.button !== 0) return;
  suppressClick = false;
  const p0 = { x: e.clientX, y: e.clientY };
  const onMove = ev => {
    if (!drag){
      if (Math.hypot(ev.clientX - p0.x, ev.clientY - p0.y) < 6) return;
      drag = { c, ref, w0: toWorld(p0.x, p0.y), ind: el('div','drop-ind hidden') };
      c.classList.add('dragging'); $('#nodes').appendChild(drag.ind); vp.classList.add('dnd');
    }
    drag.last = { x: ev.clientX, y: ev.clientY };
    updateDrag();
  };
  const onUp = () => {
    window.removeEventListener('pointermove', onMove); window.removeEventListener('pointerup', onUp); window.removeEventListener('pointercancel', onUp);
    if (!drag) return;
    suppressClick = true;
    const d = drag; drag = null;
    cancelAnimationFrame(autoPanRaf); autoPanRaf = null; vp.classList.remove('dnd');
    if (d.target) applyMove(d.ref, d.target); else renderBoard();
  };
  window.addEventListener('pointermove', onMove); window.addEventListener('pointerup', onUp); window.addEventListener('pointercancel', onUp);
}

function updateDrag(fromPan){
  const d = drag, p = current();
  const w = toWorld(d.last.x, d.last.y);
  d.c.style.transform = `translate(${w.x - d.w0.x}px,${w.y - d.w0.y}px) rotate(1.5deg)`;
  d.target = dropTarget(p, d.ref, w);
  showDropIndicator(p, d.target);
  if (!fromPan) autoPan(d.last.x);
}

function dropTarget(p, ref, w){
  const n = p.steps.length, mainIdx = ref.type === 'main' ? ref.i : -1;
  const col = Math.max(0, Math.min(n - 1, Math.round(w.x / SPACING)));
  const dxc = w.x - col * SPACING;
  if (col !== mainIdx && Math.abs(dxc) < SPACING / 2){
    const moving = ref.type === 'detail' ? p.steps[ref.i].up[ref.di] : null;
    const chain = p.steps[col].up.filter(x => x !== moving);
    let k = null;
    if (w.y < -MAIN_H/2 - 6 && isCollapsed(p.steps[col]) && p.steps[col].up.length) k = chain.length;
    else if (w.y < -MAIN_H/2 - 6){ k = 0; while (k < chain.length && detailCenterY(k) > w.y) k++; }
    else if (Math.abs(dxc) < MAIN_W/2 - 40 && Math.abs(w.y) <= MAIN_H/2 + 40) k = chain.length;
    if (k !== null){
      if (ref.type === 'detail' && ref.i === col && ref.di === k) return null;   // same place
      return { kind: 'detail', col, k };
    }
  }
  const remaining = p.steps.map((s, i) => i).filter(i => i !== mainIdx);
  let t = Math.max(1, remaining.filter(i => i * SPACING < w.x).length);
  const goalLast = p.steps[n - 1].goal && mainIdx !== n - 1 && n > 1;
  if (goalLast) t = Math.min(t, remaining.length - 1);                        // nothing goes after the goal
  if (ref.type === 'main' && t === ref.i) return null;                         // same place
  return { kind: 'main', t, remaining };
}

function showDropIndicator(p, t){
  const ind = drag.ind;
  document.querySelectorAll('.card.drop-target').forEach(x => x.classList.remove('drop-target'));
  if (!t){ ind.classList.add('hidden'); return; }
  ind.classList.remove('hidden');
  if (t.kind === 'detail'){
    const pill = document.querySelector(`.stack-pill[data-pill="${t.col}"]`);
    const y = pill ? -(MAIN_H/2 + GAP + PILL_H + GAP/2) : t.k === 0 ? -MAIN_H/2 - GAP/2 : -(MAIN_H/2 + t.k * (GAP + DET_H) + GAP/2);
    ind.style.cssText = `left:${t.col * SPACING - DET_W/2}px;top:${y - 3}px;width:${DET_W}px;height:6px`;
    const tc = document.querySelector(`.card[data-main="${t.col}"]`); if (tc) tc.classList.add('drop-target');
  } else {
    const xa = t.remaining[t.t - 1] * SPACING;
    const xb = t.t < t.remaining.length ? t.remaining[t.t] * SPACING : xa + SPACING;
    ind.style.cssText = `left:${(xa + xb) / 2 - 3}px;top:${-MAIN_H/2 - 10}px;width:6px;height:${MAIN_H + 20}px`;
  }
}

function autoPan(sx){
  cancelAnimationFrame(autoPanRaf); autoPanRaf = null;
  const W = vp.clientWidth, edge = 70;
  const v = sx < edge ? (edge - sx) / 5 : sx > W - edge ? -(sx - (W - edge)) / 5 : 0;
  if (!v) return;
  const step = () => { if (!drag) return; view.x += v; apply(); updateDrag(true); autoPanRaf = requestAnimationFrame(step); };
  autoPanRaf = requestAnimationFrame(step);
}

function applyMove(ref, t){
  const p = current();
  const dest = t.kind === 'detail' ? p.steps[t.col] : null;   // grab before indexes shift
  let item, kids = [];
  if (ref.type === 'main'){ item = p.steps.splice(ref.i, 1)[0]; kids = item.up || []; }
  else item = p.steps[ref.i].up.splice(ref.di, 1)[0];
  if (t.kind === 'detail'){
    // details can't have details of their own, so a main step's stack comes along above it
    const flat = [item, ...kids].map(({ up, down, goal, milestone, ...rest }) => rest);
    dest.up.splice(t.k, 0, ...flat);
    if (kids.length) notify(`${kids.length} detaljsteg följde med och ligger nu ovanför.`);
  } else {
    item.up = ref.type === 'main' ? kids : []; item.down = [];
    p.steps.splice(t.t, 0, item);
  }
  p.steps.forEach((s, i) => { if (s.goal && i !== p.steps.length - 1) s.goal = false; });
  save(p); renderBoard();
}

/* ---------- pan & zoom ---------- */
const view = { x:0, y:0, s:1 };
const vp = $('#viewport'), world = $('#world');
function apply(){ world.style.transform = `translate(${view.x}px,${view.y}px) scale(${view.s})`; placeRuler(); }
function zoomAt(px, py, factor){
  const s = Math.min(2, Math.max(0.2, view.s * factor));
  const k = s / view.s;
  view.x = px - (px - view.x) * k; view.y = py - (py - view.y) * k; view.s = s; apply();
}
function fitView(){
  const W = vp.clientWidth, H = vp.clientHeight - RULER_H, pad = 80;
  const bw = bounds.x1 - bounds.x0, bh = bounds.y1 - bounds.y0;
  const want = Math.min((W - pad*2) / bw, (H - pad*2 - 40) / bh);
  view.s = Math.min(1, Math.max(0.55, want));
  // if everything doesn't fit at a readable size, show from the start instead of centering
  view.x = want < 0.55 ? Math.min(pad, W/2) - bounds.x0 * view.s : W/2 - (bounds.x0 + bw/2) * view.s;
  view.y = H/2 + 20 - (bounds.y0 + bh/2) * view.s;
  apply();
}
function ensureVisible(x, y){
  const W = vp.clientWidth, H = vp.clientHeight - RULER_H, m = 140;
  const sx = view.x + x * view.s, sy = view.y + y * view.s;
  if (sx > W - m) view.x -= sx - (W - m) + 120;
  if (sx < m) view.x += m - sx;
  if (sy < m) view.y += m - sy;
  if (sy > H - m) view.y -= sy - (H - m);
  apply();
}
$('#zoomIn').onclick = () => zoomAt(vp.clientWidth/2, vp.clientHeight/2, 1.2);
$('#zoomOut').onclick = () => zoomAt(vp.clientWidth/2, vp.clientHeight/2, 1/1.2);
$('#zoomFit').onclick = fitView;

vp.addEventListener('wheel', e => {
  e.preventDefault();
  if (e.ctrlKey || e.metaKey) zoomAt(e.clientX, e.clientY, Math.exp(-e.deltaY * 0.01));
  else { view.x -= e.deltaX; view.y -= e.deltaY; apply(); }
}, { passive:false });

const pointers = new Map(); let pinch = null;
vp.addEventListener('pointerdown', e => {
  if (e.target.closest('.card,.plus')) return;
  vp.setPointerCapture(e.pointerId);
  pointers.set(e.pointerId, { x:e.clientX, y:e.clientY });
  vp.classList.add('dragging');
  if (pointers.size === 2){
    const [a,b] = [...pointers.values()];
    pinch = { d: Math.hypot(a.x-b.x, a.y-b.y) };
  }
});
vp.addEventListener('pointermove', e => {
  const prev = pointers.get(e.pointerId); if (!prev) return;
  const cur = { x:e.clientX, y:e.clientY };
  if (pointers.size === 1){ view.x += cur.x - prev.x; view.y += cur.y - prev.y; apply(); }
  pointers.set(e.pointerId, cur);
  if (pointers.size === 2 && pinch){
    const [a,b] = [...pointers.values()];
    const d = Math.hypot(a.x-b.x, a.y-b.y);
    zoomAt((a.x+b.x)/2, (a.y+b.y)/2, d / pinch.d); pinch.d = d;
  }
});
const up = e => { pointers.delete(e.pointerId); if (pointers.size < 2) pinch = null; if (!pointers.size) vp.classList.remove('dragging'); };
vp.addEventListener('pointerup', up); vp.addEventListener('pointercancel', up);
window.addEventListener('resize', () => { if (current()) apply(); });

/* ---------- drawer ---------- */
let ctx = null;
const getItem = c => { const p = current(); const s = p.steps[c.i]; return c.type === 'detail' ? s[c.dir][c.di] : s; };
function openDrawer(c){
  ctx = c; const p = current();
  const isNew = c.mode === 'new';
  let item = isNew ? null : getItem(c);
  const now = new Date();
  // context text
  let ctxText = '';
  if (isNew && c.kind === 'start') ctxText = 'Startpunkten för projektet. Datumet sätts till idag.';
  else if (isNew && c.kind === 'main') ctxText = c.at >= p.steps.length ? `Nytt steg efter "${p.steps[p.steps.length-1].title}"` : `Nytt steg mellan "${p.steps[c.at-1].title}" och "${p.steps[c.at].title}"`;
  else if (isNew && c.kind === 'detail') ctxText = `Detaljsteg till "${p.steps[c.i].title}"`;
  else if (c.type === 'detail') ctxText = `Detaljsteg till "${p.steps[c.i].title}"`;
  else ctxText = c.i === 0 ? 'Projektets startpunkt' : `Steg ${c.i + 1} av ${p.steps.length}`;
  $('#dCtx').textContent = ctxText;

  $('#dTitle').textContent = isNew ? (c.kind === 'start' ? 'Starta projekt' : c.kind === 'detail' ? 'Nytt detaljsteg' : 'Nytt steg') : 'Redigera';
  $('#fTitle').value = item ? item.title : (c.kind === 'start' ? 'Projekt startat' : '');
  $('#fWho').value = item ? (item.who || '') : '';
  $('#fDate').value = item ? (item.date || '') : (c.kind === 'start' ? todayISO(now) : '');
  $('#fNote').value = item ? (item.note || '') : '';
  $('#fCrit').checked = item ? !!item.critical : false;
  const isStart = isNew ? c.kind === 'start' : (c.type === 'main' && c.i === 0);
  const isMainStep = isNew ? c.kind === 'main' : (c.type === 'main' && c.i > 0);
  // Goal is only offered for the last main step (nothing to its right)
  const goalOk = isNew ? (c.kind === 'main' && c.at >= p.steps.length) : (c.type === 'main' && c.i === p.steps.length - 1 && c.i > 0);
  $('#typeRow').classList.toggle('hidden', !isMainStep);
  $('#goalOpt').classList.toggle('hidden', !goalOk);
  setRadio('fType', item && item.goal && goalOk ? 'goal' : item && item.milestone ? 'milestone' : 'step');
  setRadio('fStatus', item ? (item.done ? 'done' : item.doing ? 'doing' : 'todo') : (isStart ? 'done' : 'todo'));
  const warn = item ? dateWarnings(p).get(item) : null;
  $('#dWarn').textContent = warn ? '⚠ ' + warn : ''; $('#dWarn').classList.toggle('hidden', !warn);
  $('#dDelete').classList.toggle('hidden', isNew);
  $('#dSave').textContent = isNew ? 'Lägg till' : 'Spara';

  // name suggestions
  const names = new Set();
  p.steps.forEach(s => [s, ...s.up].forEach(x => people(x).forEach(n => names.add(n))));
  const dl = $('#whoList'); dl.innerHTML = ''; names.forEach(n => { const o = el('option'); o.value = n; dl.appendChild(o); });

  $('#scrim').classList.remove('hidden');
  $('#drawer').classList.add('open'); $('#drawer').setAttribute('aria-hidden','false');
  setTimeout(() => { const t = $('#fTitle'); t.focus(); if (!isNew || c.kind === 'start') t.select(); }, 60);
}
function closeDrawer(silent){
  $('#drawer').classList.remove('open'); $('#drawer').setAttribute('aria-hidden','true');
  $('#scrim').classList.add('hidden'); ctx = null;
}
$('#dClose').onclick = () => closeDrawer();
$('#dCancel').onclick = () => closeDrawer();
$('#scrim').onclick = () => closeDrawer();
document.addEventListener('keydown', e => {
  if (e.key === 'Escape'){ if (!$('#modal').classList.contains('hidden')) closeModal(); else if (ctx) closeDrawer(); return; }
  if (!(e.metaKey || e.ctrlKey) || !current() || ctx || !$('#modal').classList.contains('hidden')) return;
  if (e.target.closest && e.target.closest('input,textarea,select')) return;
  const k = e.key.toLowerCase();
  if (k === 'z' && !e.shiftKey){ e.preventDefault(); undo(); }
  else if ((k === 'z' && e.shiftKey) || k === 'y'){ e.preventDefault(); redo(); }
});
$('#undoBtn').onclick = undo;
$('#redoBtn').onclick = redo;


/* ---------- print / PDF ---------- */
let prePrint = null;
const PRINT_W = 1000, PRINT_H = 500;
function beforePrint(){
  const p = current(); if (!p || prePrint) return;
  prePrint = { view: { ...view }, theme: document.documentElement.dataset.theme };
  document.documentElement.dataset.theme = 'light';
  $('#printHead h1').textContent = p.name;
  const items = allItems(p);
  $('#printHead p').textContent = [p.description, `${items.filter(x => x.done).length}/${items.length} klara`, 'Utskrivet ' + fmtDate({ date: todayISO(new Date()) }) + ' ' + new Date().getFullYear()].filter(Boolean).join('  ·  ');
  const bw = bounds.x1 - bounds.x0, bh = bounds.y1 - bounds.y0;
  view.s = Math.min(1, PRINT_W / bw, PRINT_H / bh);
  view.x = (PRINT_W - bw * view.s) / 2 - bounds.x0 * view.s;
  view.y = (PRINT_H - bh * view.s) / 2 - bounds.y0 * view.s;
  apply();
}
function afterPrint(){
  if (!prePrint) return;
  Object.assign(view, prePrint.view);
  if (prePrint.theme) document.documentElement.dataset.theme = prePrint.theme; else delete document.documentElement.dataset.theme;
  prePrint = null; apply();
}
window.addEventListener('beforeprint', beforePrint);
window.addEventListener('afterprint', afterPrint);
$('#printBtn').onclick = () => { beforePrint(); window.print(); };

$('#dForm').onsubmit = e => {
  e.preventDefault(); if (!ctx) return;
  const p = current();
  const status = getRadio('fStatus');
  const type = $('#typeRow').classList.contains('hidden') ? null : getRadio('fType');
  const data = {
    title: $('#fTitle').value.trim() || 'Namnlöst steg',
    who: people({ who: $('#fWho').value }).join(', '),
    date: $('#fDate').value, time: '',
    note: $('#fNote').value.trim(), critical: $('#fCrit').checked,
    done: status === 'done', doing: status === 'doing'
  };
  if (type !== null){ data.goal = type === 'goal'; data.milestone = type === 'milestone'; }
  let focus = null;
  if (ctx.mode === 'new'){
    const item = { id: uid(), created: Date.now(), ...data };
    if (ctx.kind === 'start' || ctx.kind === 'main'){
      item.up = []; item.down = [];
      const at = ctx.kind === 'start' ? 0 : ctx.at;
      p.steps.splice(at, 0, item); focus = { x: at * SPACING, y: 0 };
    } else {
      p.steps[ctx.i][ctx.dir].push(item);
      setCollapsed(p.steps[ctx.i], false);
      const n = p.steps[ctx.i][ctx.dir].length, sign = -1;
      focus = { x: ctx.i * SPACING, y: sign * (MAIN_H/2 + n * (GAP + DET_H) + 40) };
    }
  } else {
    Object.assign(getItem(ctx), data);
  }
  save(); closeDrawer(); renderBoard();
  if (focus) ensureVisible(focus.x + (focus.y ? 0 : SPACING/2), focus.y);
};

$('#dDelete').onclick = async () => {
  if (!ctx || ctx.mode !== 'edit') return;
  const p = current(), c = ctx;
  if (c.type === 'detail'){
    if (!await ask('Ta bort detaljsteget?', 'Ta bort', true)) return;
    p.steps[c.i][c.dir].splice(c.di, 1);
  } else {
    const s = p.steps[c.i]; const n = s.up.length;
    if (!await ask(n ? `Ta bort steget och dess ${n} detaljsteg?` : 'Ta bort steget?', 'Ta bort', true)) return;
    p.steps.splice(c.i, 1);
  }
  save(); closeDrawer(); renderBoard();
};

$('#backBtn').onclick = () => { location.hash = ''; };

async function boot(){
  $('#fatal').classList.add('hidden'); $('#loading').classList.remove('hidden');
  try {
    db.projects = await store.init();
    $('#loading').classList.add('hidden');
    route();
  } catch(e){
    console.error(e);
    $('#loading').classList.add('hidden');
    $('#fatalText').textContent = e && e.message ? e.message : String(e);
    $('#fatal').classList.remove('hidden');
  }
}
$('#fatalRetry').onclick = () => location.reload();
boot();
})();
