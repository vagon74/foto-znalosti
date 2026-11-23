/* Foto do znalostí – vyfotí snímek a odešle ho do OneDrive (Microsoft Graph).
 * Fotka se nikdy neukládá do Fotek iPhonu: pořizuje se přes <input capture>,
 * do odeslání čeká jen v úložišti aplikace (IndexedDB) a po potvrzení se smaže.
 * Převzato z MOPANI Backoffice: zmenšení přes canvas (lib/image.ts),
 * upload do 4 MB / po částech (storage/docStore.ts), offline fronta. */
'use strict';

const VERSION = '1.2.2';
const cfg = window.FOTO_CONFIG || {};
const DEMO = new URLSearchParams(location.search).has('demo'); // náhled bez přihlášení, upload se jen simuluje
const GRAPH = 'https://graph.microsoft.com/v1.0';
const AUTH = (cfg.authority || 'https://login.microsoftonline.com/consumers') + '/oauth2/v2.0';
const REDIRECT_URI = new URL('./', location.href).href;
const SIMPLE_LIMIT = 4 * 1024 * 1024;
const CHUNK = 320 * 1024 * 10; // násobek 320 KiB

const $ = (id) => document.getElementById(id);
const settings = {
  get size() { return Number(localStorage.getItem('fz.size') ?? 3072); },
  set size(v) { localStorage.setItem('fz.size', String(v)); },
  get folder() { return localStorage.getItem('fz.folder') || cfg.inboxFolder || 'Znalosti/Inbox'; },
  set folder(v) { localStorage.setItem('fz.folder', v.trim().replace(/^\/+|\/+$/g, '')); },
  get kind() { return localStorage.getItem('fz.kind') || 'auto'; },
  set kind(v) { localStorage.setItem('fz.kind', v); },
};

/* ---------------- přihlášení (OAuth 2.0 auth code + PKCE, bez serveru) ---------------- */

const b64url = (bytes) => btoa(String.fromCharCode(...new Uint8Array(bytes))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const randomStr = (n = 32) => b64url(crypto.getRandomValues(new Uint8Array(n)));

function loadTokens() { try { return JSON.parse(localStorage.getItem('fz.tokens')) || null; } catch { return null; } }
function saveTokens(t) { t ? localStorage.setItem('fz.tokens', JSON.stringify(t)) : localStorage.removeItem('fz.tokens'); }

async function login() {
  const verifier = randomStr(48);
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const state = randomStr(16);
  localStorage.setItem('fz.pkce', JSON.stringify({ verifier, state }));
  const q = new URLSearchParams({
    client_id: cfg.clientId, response_type: 'code', response_mode: 'query', redirect_uri: REDIRECT_URI,
    scope: cfg.scopes.join(' '), code_challenge: challenge, code_challenge_method: 'S256', state,
  });
  location.assign(`${AUTH}/authorize?${q}`);
}

async function tokenRequest(params) {
  const res = await fetch(`${AUTH}/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: cfg.clientId, scope: cfg.scopes.join(' '), ...params }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) { const e = new Error(data.error_description || data.error || `HTTP ${res.status}`); e.code = data.error; throw e; }
  const t = { access: data.access_token, refresh: data.refresh_token, exp: Date.now() + (data.expires_in - 60) * 1000 };
  saveTokens(t);
  return t;
}

async function handleRedirect() {
  const p = new URLSearchParams(location.search);
  if (!p.has('code') && !p.has('error')) return;
  window.history.replaceState(null, '', REDIRECT_URI);
  if (p.has('error')) { toast(`Přihlášení se nepovedlo: ${p.get('error_description') || p.get('error')}`); return; }
  const saved = JSON.parse(localStorage.getItem('fz.pkce') || '{}');
  localStorage.removeItem('fz.pkce');
  if (saved.state !== p.get('state')) { toast('Přihlášení se nepovedlo (neplatný stav), zkuste to znovu.'); return; }
  await tokenRequest({ grant_type: 'authorization_code', code: p.get('code'), redirect_uri: REDIRECT_URI, code_verifier: saved.verifier });
  await loadUser();
}

let refreshing = null;
/** Platný access token, nebo null (je potřeba se přihlásit). Síťová chyba se propaguje. */
async function getToken() {
  if (DEMO) return 'demo';
  const t = loadTokens();
  if (!t) return null;
  if (t.exp > Date.now()) return t.access;
  if (!t.refresh) { saveTokens(null); return null; }
  refreshing ??= tokenRequest({ grant_type: 'refresh_token', refresh_token: t.refresh })
    .catch((e) => { if (e.code) { saveTokens(null); return null; } throw e; }) // invalid_grant → nové přihlášení
    .finally(() => { refreshing = null; });
  const nt = await refreshing;
  return nt ? nt.access : null;
}

/** fetch s časovým limitem – když iOS uspí aplikaci uprostřed požadavku, jinak by visel navždy */
async function fetchT(url, init = {}, ms = 90000) {
  const ac = new AbortController();
  const t = setTimeout(() => ac.abort(), ms);
  try { return await fetch(url, { ...init, signal: ac.signal }); }
  catch (e) { throw Object.assign(new Error(e.name === 'AbortError' ? 'vypršel čas spojení' : (e.message || 'chyba sítě')), { network: true }); }
  finally { clearTimeout(t); }
}

async function graph(path, init = {}, retry = true) {
  const token = await getToken();
  if (!token) throw Object.assign(new Error('Je potřeba se přihlásit'), { needLogin: true });
  const res = await fetchT(path.startsWith('http') ? path : GRAPH + path, { ...init, headers: { ...init.headers, Authorization: `Bearer ${token}` } });
  if (res.status === 401 && retry) {
    const t = loadTokens(); if (t) { t.exp = 0; saveTokens(t); }
    return graph(path, init, false);
  }
  return res;
}

async function loadUser() {
  try {
    const res = await graph('/me/drive?$select=owner');
    if (res.ok) {
      const d = await res.json();
      const name = d.owner?.user?.displayName || d.owner?.user?.email || '';
      if (name) localStorage.setItem('fz.userName', name);
    }
  } catch { /* offline – nevadí */ }
}

/* ---------------- obrázky ---------------- */

async function toJpeg(file, max, quality = 0.85) {
  if (!max) return file; // originál
  let image, width, height, close;
  try {
    const bmp = await createImageBitmap(file, { imageOrientation: 'from-image' });
    image = bmp; width = bmp.width; height = bmp.height; close = () => bmp.close();
  } catch {
    const url = URL.createObjectURL(file);
    const img = new Image(); img.src = url;
    try { await img.decode(); } catch { URL.revokeObjectURL(url); throw new Error('Tento formát obrázku nejde otevřít.'); }
    image = img; width = img.naturalWidth; height = img.naturalHeight; close = () => URL.revokeObjectURL(url);
  }
  const scale = Math.min(1, max / Math.max(width, height));
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(width * scale); canvas.height = Math.round(height * scale);
  canvas.getContext('2d').drawImage(image, 0, 0, canvas.width, canvas.height);
  close();
  const blob = await new Promise((r) => canvas.toBlob(r, 'image/jpeg', quality));
  canvas.width = canvas.height = 0; // uvolnit paměť v Safari
  if (!blob) throw new Error('Obrázek nejde převést na JPEG.');
  return blob;
}

const extOf = (blob) => ({ 'image/jpeg': 'jpg', 'image/png': 'png', 'image/heic': 'heic', 'image/heif': 'heif', 'image/webp': 'webp' }[blob.type] || 'jpg');

/* ---------------- datum pořízení (EXIF) ----------------
 * Zmenšením přes canvas se EXIF ztratí, proto se datum čte z originálu předem.
 * Vrací { at: 'RRRR-MM-DDTHH:MM:SS+01:00', src: 'exif' | 'fotoaparat' | 'soubor' }. */

async function exifDate(file) {
  try {
    const v = new DataView(await file.slice(0, 256 * 1024).arrayBuffer());
    if (v.getUint16(0) !== 0xffd8) return null; // jen JPEG
    let p = 2;
    while (p + 4 < v.byteLength) {
      const marker = v.getUint16(p), len = v.getUint16(p + 2);
      if (marker === 0xffe1 && v.getUint32(p + 4) === 0x45786966) return parseTiff(v, p + 10); // "Exif"
      if ((marker & 0xff00) !== 0xff00 || marker === 0xffda) return null;
      p += 2 + len;
    }
  } catch { /* poškozený soubor – použije se náhradní datum */ }
  return null;
}

function parseTiff(v, t) {
  const le = v.getUint16(t) === 0x4949;
  const u16 = (o) => v.getUint16(t + o, le), u32 = (o) => v.getUint32(t + o, le);
  const ascii = (entry) => {
    const n = u32(entry + 4); const off = n > 4 ? u32(entry + 8) : entry + 8;
    let s = ''; for (let i = 0; i < n - 1; i++) s += String.fromCharCode(v.getUint8(t + off + i));
    return s.trim();
  };
  const ifd = (off) => {
    const tags = {}; const n = u16(off);
    for (let i = 0; i < n; i++) { const e = off + 2 + i * 12; tags[u16(e)] = e; }
    return tags;
  };
  const ifd0 = ifd(u32(4));
  const ex = ifd0[0x8769] ? ifd(u32(ifd0[0x8769] + 8)) : {};
  const raw = ex[0x9003] ? ascii(ex[0x9003]) : ifd0[0x0132] ? ascii(ifd0[0x0132]) : '';
  const m = raw.match(/^(\d{4}):(\d\d):(\d\d) (\d\d):(\d\d):(\d\d)/);
  if (!m || m[1] === '0000') return null;
  const offset = ex[0x9011] ? ascii(ex[0x9011]) : '';
  const local = `${m[1]}-${m[2]}-${m[3]}T${m[4]}:${m[5]}:${m[6]}`;
  if (/^[+-]\d\d:\d\d$/.test(offset)) return local + offset;
  // bez časové zóny = místní čas zařízení
  return isoLocal(new Date(+m[1], m[2] - 1, +m[3], +m[4], +m[5], +m[6]));
}

/** ISO s místní časovou zónou, např. 2026-11-11T14:16:52+01:00 */
function isoLocal(d) {
  const p = (n) => String(Math.abs(n)).padStart(2, '0');
  const tz = -d.getTimezoneOffset();
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}${tz >= 0 ? '+' : '-'}${p(Math.trunc(tz / 60))}:${p(tz % 60)}`;
}

async function takenAt(file, fromCamera) {
  const exif = await exifDate(file);
  if (exif) return { at: exif, src: 'exif' };
  if (fromCamera) return { at: isoLocal(new Date()), src: 'fotoaparat' };
  return { at: isoLocal(new Date(file.lastModified || Date.now())), src: 'soubor' };
}

/* ---------------- fronta v IndexedDB ---------------- */

let dbP;
function db() {
  dbP ??= new Promise((resolve, reject) => {
    const req = indexedDB.open('foto-znalosti', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('queue', { keyPath: 'id' });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbP;
}
async function tx(mode, fn) {
  const d = await db();
  return new Promise((resolve, reject) => {
    const t = d.transaction('queue', mode); const s = t.objectStore('queue');
    const r = fn(s); let out;
    if (r) r.onsuccess = () => { out = r.result; };
    t.oncomplete = () => resolve(out); t.onerror = () => reject(t.error); t.onabort = () => reject(t.error);
  });
}
const qAll = () => tx('readonly', (s) => s.getAll());
const qPut = (item) => tx('readwrite', (s) => s.put(item));
const qDel = (id) => tx('readwrite', (s) => s.delete(id));

function stamp(d = new Date()) {
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}

/** pages: [{ blob, taken: { at, src } }] – jméno souboru se odvozuje z data pořízení první strany */
async function enqueue(pages, kind, note) {
  const now = new Date();
  const first = [...pages].sort((a, b) => Date.parse(a.taken.at) - Date.parse(b.taken.at))[0];
  const base = `${stamp(new Date(first.taken.at))}_${randomStr(3).replace(/[-_]/g, 'x').slice(0, 4)}${kind !== 'auto' ? '_' + kind : ''}`;
  const item = {
    id: base, base, kind, note, createdAt: now.toISOString(), folder: settings.folder,
    pages: await Promise.all(pages.map(async (pg, i) => ({ bytes: await pg.blob.arrayBuffer(), type: pg.blob.type || 'image/jpeg', taken: pg.taken, name: `${base}${pages.length > 1 ? '-' + (i + 1) : ''}.${extOf(pg.blob)}`, done: false }))),
    metaDone: false, error: null,
  };
  await qPut(item);
  return item;
}

/* ---------------- upload do OneDrive ---------------- */

const encPath = (p) => p.split('/').filter(Boolean).map(encodeURIComponent).join('/');

async function putFile(folder, name, blob) {
  if (DEMO) { await new Promise((r) => setTimeout(r, 700)); return; }
  const item = `/me/drive/root:/${encPath(folder + '/' + name)}:`;
  if (blob.size <= SIMPLE_LIMIT) {
    const res = await graph(`${item}/content`, { method: 'PUT', headers: { 'Content-Type': blob.type || 'application/octet-stream' }, body: blob });
    if (!res.ok) throw await graphError(res);
    return;
  }
  const s = await graph(`${item}/createUploadSession`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }),
  });
  if (!s.ok) throw await graphError(s);
  const { uploadUrl } = await s.json();
  for (let start = 0; start < blob.size; start += CHUNK) {
    const end = Math.min(start + CHUNK, blob.size);
    const res = await fetchT(uploadUrl, { method: 'PUT', headers: { 'Content-Range': `bytes ${start}-${end - 1}/${blob.size}` }, body: blob.slice(start, end) });
    if (!res.ok && res.status !== 202) throw await graphError(res);
  }
}

async function graphError(res) {
  let msg = `HTTP ${res.status}`;
  try { const d = await res.json(); msg = d.error?.message || msg; } catch { /* */ }
  return Object.assign(new Error(msg), { status: res.status });
}

/** Data strany jako Blob. Nové položky mají `bytes` (ArrayBuffer) – Blob uložený v IndexedDB
 *  Safari na iPhonu po znovuotevření aplikace někdy neumí přečíst (WebKitBlobResource). */
async function pageBlob(page) {
  if (page.bytes) return new Blob([page.bytes], { type: page.type || 'image/jpeg' });
  if (!page.blob) throw Object.assign(new Error('fotka v telefonu chybí'), { broken: true });
  try { return new Blob([await page.blob.arrayBuffer()], { type: page.blob.type || 'image/jpeg' }); }
  catch { throw Object.assign(new Error('fotku v telefonu už nejde načíst – smažte ji z fronty a vyfoťte znovu'), { broken: true }); }
}

let pumping = false, pumpGen = 0, pumpAt = 0, retryTimer = null;
async function pump() {
  // pojistka: běh, který visí přes 3 minuty (aplikace byla uspaná), se nahradí novým
  if (pumping && Date.now() - pumpAt < 180000) return;
  const gen = ++pumpGen;
  pumping = true; pumpAt = Date.now(); clearTimeout(retryTimer);
  try {
    for (const item of (await qAll()).sort((a, b) => a.createdAt.localeCompare(b.createdAt))) {
      if (gen !== pumpGen) return;
      try {
        for (const page of item.pages) {
          if (page.done) continue;
          await putFile(item.folder, page.name, await pageBlob(page));
          if (gen !== pumpGen) return;
          page.done = true; page.blob = null; page.bytes = null; // hotovou stranu hned smazat z telefonu
          item.error = null; pumpAt = Date.now();
          await qPut(item); render();
        }
        // metadata až nakonec – na PC znamenají „snímek je kompletní, zpracuj ho“
        const taken = item.pages.map((p) => p.taken?.at).filter(Boolean).sort((a, b) => Date.parse(a) - Date.parse(b));
        const meta = {
          id: item.base, typ: item.kind, poznamka: item.note || '',
          porizeno: taken[0] || null,             // HLAVNÍ datum = pořízení první (nejstarší) fotky
          porizeno_do: taken[taken.length - 1] || null,
          nahrano: isoLocal(new Date()),          // jen doplňková informace
          strany: item.pages.map((p) => ({ soubor: p.name, porizeno: p.taken?.at || null, zdroj_data: p.taken?.src || null })),
          zdroj: `Foto do znalostí ${VERSION} (iPhone)`,
        };
        await putFile(item.folder, `${item.base}.json`, new Blob([JSON.stringify(meta, null, 2)], { type: 'application/json' }));
        await qDel(item.id);
        addHistory({ base: item.base, kind: item.kind, note: item.note, pages: item.pages.length, at: new Date().toISOString() });
      } catch (e) {
        if (gen !== pumpGen) return;
        item.error = e.needLogin ? 'čeká na přihlášení' : (navigator.onLine ? e.message : 'čeká na signál');
        await qPut(item).catch(() => {});
        if (e.needLogin) { showLoginIfNeeded(); break; }
        if (e.broken || (e.status && e.status !== 429 && e.status < 500)) continue; // vadný snímek neblokuje ostatní
        retryTimer = setTimeout(pump, 30000);
        break;
      }
    }
  } finally {
    if (gen === pumpGen) { pumping = false; render(); }
  }
}

/* ---------------- historie (jen metadata, žádné fotky) ---------------- */

function sentHistory() { try { return JSON.parse(localStorage.getItem('fz.history')) || []; } catch { return []; } }
function addHistory(h) { localStorage.setItem('fz.history', JSON.stringify([h, ...sentHistory()].slice(0, 30))); }

/* ---------------- UI ---------------- */

const KIND_LABEL = { auto: 'snímek', dokument: 'dokument', vizitka: 'vizitka', cedule: 'cedule', tabule: 'tabule' };
let draft = []; // stránky před odesláním (jen v paměti)
let draftUrls = [];

function toast(msg, ms = 3500) {
  const t = $('toast'); t.textContent = msg; t.hidden = false;
  clearTimeout(toast.t); toast.t = setTimeout(() => { t.hidden = true; }, ms);
}

function signedIn() { return DEMO || !!loadTokens(); }

function showLoginIfNeeded() { render(); }

async function render() {
  const logged = signedIn();
  const reviewing = draft.length > 0;
  // bez přihlášení jen úvodní obrazovka s tlačítkem (jako MOPANI Backoffice)
  $('loginView').hidden = logged;
  $('appHeader').hidden = !logged;
  $('appMain').hidden = !logged;
  $('configWarn').hidden = !!cfg.clientId;
  $('loginBtn').disabled = !cfg.clientId;
  if (!logged) {
    const n = (await qAll().catch(() => [])).length;
    $('loginPending').hidden = !n;
    $('loginPending').textContent = `V telefonu čeká ${n} ${n === 1 ? 'neodeslaný snímek' : n < 5 ? 'neodeslané snímky' : 'neodeslaných snímků'}. Po přihlášení se odešlou.`;
    return;
  }
  $('kinds').hidden = false;
  $('captureView').hidden = reviewing;
  $('reviewView').hidden = !reviewing;
  if (reviewing) renderNoteSuggest();
  $('statusView').hidden = reviewing;
  $('settingsView').hidden = reviewing;
  $('logoutBtn').hidden = !logged || DEMO;
  const name = DEMO ? 'demo režim' : localStorage.getItem('fz.userName');
  $('accountBtn').hidden = !logged || !name; $('accountBtn').textContent = name || '';

  const items = await qAll().catch(() => []);
  const q = $('queue'); q.textContent = '';
  if (!items.length) q.append(li('Nic nečeká na odeslání.', null, 'empty'));
  for (const it of items) {
    const done = it.pages.filter((p) => p.done).length;
    const row = li(`${KIND_LABEL[it.kind]} · ${fmtTime(it.createdAt)} · ${done}/${it.pages.length} str.${it.note ? ' · ' + it.note : ''}`,
      it.error ? `⚠ ${it.error}` : 'odesílám…', 'qitem', it.error ? 'err' : 'wait');
    row.onclick = () => queueAction(it);
    q.append(row);
  }
  if (items.length) q.append(li('Ťuknutím na snímek ho můžete odeslat znovu nebo smazat z fronty.', null, 'empty'));
  $('netState').textContent = navigator.onLine ? '' : 'Bez připojení – fotky počkají v aplikaci a odešlou se samy.';
  const h = $('history'); h.textContent = '';
  const hist = sentHistory();
  if (!hist.length) h.append(li('Zatím nic.', null, 'empty'));
  for (const x of hist.slice(0, 10)) {
    h.append(li(`${KIND_LABEL[x.kind]} · ${fmtTime(x.at)}${x.pages > 1 ? ` · ${x.pages} str.` : ''}${x.note ? ' · ' + x.note : ''}`, '✓ v OneDrive', null, 'ok'));
  }
}

async function queueAction(it) {
  if (confirm(`Odeslat „${KIND_LABEL[it.kind]}${it.note ? ' · ' + it.note : ''}“ znovu?\n\n(Zrušit = nabídne smazání z fronty)`)) {
    pumping = false; pumpGen++; pump(); return;
  }
  if (confirm('Smazat tento snímek z fronty v telefonu? Strany, které už jsou v OneDrive, tam zůstanou.')) {
    await qDel(it.id); render(); pump();
  }
}

function li(text, status, cls, stCls) {
  const el = document.createElement('li'); if (cls) el.className = cls;
  const a = document.createElement('span'); a.textContent = text; el.append(a);
  if (status) { const b = document.createElement('span'); b.className = `st ${stCls || ''}`; b.textContent = status; el.append(b); }
  return el;
}
const fmtTime = (iso) => new Date(iso).toLocaleString('cs-CZ', { day: 'numeric', month: 'numeric', hour: '2-digit', minute: '2-digit' });

function renderDraft() {
  draftUrls.forEach(URL.revokeObjectURL); draftUrls = [];
  const box = $('pages'); box.textContent = '';
  draft.forEach(({ blob, taken }, i) => {
    const url = URL.createObjectURL(blob); draftUrls.push(url);
    const d = document.createElement('div'); d.className = 'page';
    const img = document.createElement('img'); img.src = url; img.alt = `Strana ${i + 1}`;
    const del = document.createElement('button'); del.className = 'del'; del.textContent = '✕'; del.setAttribute('aria-label', 'Odebrat stranu');
    del.onclick = () => { draft.splice(i, 1); renderDraft(); render(); };
    const num = document.createElement('span'); num.className = 'num';
    num.textContent = `${i + 1} · ${fmtTime(taken.at)}`;
    num.title = { exif: 'datum pořízení z fotky', fotoaparat: 'vyfoceno teď', soubor: 'datum souboru (fotka nemá EXIF)' }[taken.src];
    d.append(img, del, num); box.append(d);
  });
}

async function addFiles(files, fromCamera = false) {
  const list = [...(files || [])];
  if (!list.length) return;
  try {
    for (const f of list) {
      const taken = await takenAt(f, fromCamera); // EXIF číst z originálu – po zmenšení už není
      draft.push({ blob: await toJpeg(f, settings.size), taken });
    }
  } catch (e) { toast(e.message); }
  renderDraft(); render();
}

function clearDraft() {
  draft = []; renderDraft(); $('note').value = ''; render();
}

/* poslední poznámky – nabídnou se při dalším focení (např. stejná akce) */
function recentNotes() { try { return JSON.parse(localStorage.getItem('fz.notes')) || []; } catch { return []; } }
function rememberNote(n) {
  if (!n) return;
  localStorage.setItem('fz.notes', JSON.stringify([n, ...recentNotes().filter((x) => x !== n)].slice(0, 5)));
}
function renderNoteSuggest() {
  const box = $('noteSuggest'); box.textContent = '';
  const cur = $('note').value.trim();
  for (const n of recentNotes()) {
    if (n === cur) continue;
    const b = document.createElement('button'); b.type = 'button'; b.className = 'chip suggest'; b.textContent = `↺ ${n}`;
    b.onclick = () => { $('note').value = n; renderNoteSuggest(); };
    box.append(b);
  }
  box.hidden = !box.childElementCount;
}

function bind() {
  $('loginBtn').onclick = () => login().catch((e) => toast(e.message));
  $('logoutBtn').onclick = () => {
    if (!confirm('Odhlásit? Neodeslané fotky zůstanou ve frontě.')) return;
    saveTokens(null); localStorage.removeItem('fz.userName'); render();
  };
  $('shootBtn').onclick = () => $('cameraInp').click();
  $('addPageBtn').onclick = () => $('cameraInp').click();
  $('galleryBtn').onclick = () => $('galleryInp').click();
  for (const id of ['cameraInp', 'galleryInp']) {
    $(id).onchange = (e) => { addFiles(e.target.files, id === 'cameraInp').finally(() => { e.target.value = ''; }); };
  }
  $('kinds').onclick = (e) => {
    const k = e.target.closest('[data-kind]')?.dataset.kind; if (!k) return;
    settings.kind = k; syncKinds();
  };
  $('discardBtn').onclick = () => { if (confirm('Zahodit vyfocené strany?')) clearDraft(); };
  $('sendBtn').onclick = async () => {
    $('sendBtn').disabled = true;
    try {
      const note = $('note').value.trim();
      await enqueue(draft, settings.kind, note);
      rememberNote(note);
      clearDraft();
      toast(navigator.onLine ? 'Odesílám do OneDrive…' : 'Uloženo ve frontě, odešle se po připojení.');
      pump();
    } catch (e) { toast(`Nepodařilo se uložit: ${e.message}`); }
    finally { $('sendBtn').disabled = false; }
  };
  $('sizeSel').value = String(settings.size);
  $('sizeSel').onchange = (e) => { settings.size = Number(e.target.value); };
  $('folderInp').value = settings.folder;
  $('folderInp').onchange = (e) => { settings.folder = e.target.value || cfg.inboxFolder; e.target.value = settings.folder; };
  $('ver').textContent = VERSION;
  addEventListener('online', () => { render(); pump(); });
  addEventListener('offline', render);
  document.addEventListener('visibilitychange', () => {
    if (document.hidden) return;
    pump();
    navigator.serviceWorker?.getRegistration().then((r) => r?.update()).catch(() => {}); // hledat novou verzi i po návratu do aplikace
  });
}

function syncKinds() {
  for (const b of $('kinds').querySelectorAll('[data-kind]')) {
    const on = b.dataset.kind === settings.kind; b.classList.toggle('on', on); b.setAttribute('aria-checked', on);
  }
}

(async function main() {
  bind(); syncKinds();
  if ('serviceWorker' in navigator) {
    const hadSw = !!navigator.serviceWorker.controller;
    navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).catch(() => {});
    // nová verze převzala řízení → jednou obnovit stránku (ne uprostřed focení, to by smazalo nafocené strany)
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (hadSw && !draft.length && !sessionStorage.getItem('fz.reloaded')) { sessionStorage.setItem('fz.reloaded', '1'); location.reload(); }
    });
  }
  navigator.storage?.persist?.().catch(() => {});
  try { await handleRedirect(); } catch (e) { toast(`Přihlášení se nepovedlo: ${e.message}`); }
  await render();
  pump();
})();
