/* =====================================================================
   Company data from one Excel workbook in SharePoint (Safety_App_Data.xlsx):

     "Job Information 2027", "... 2026", "... 2025"   Site Main ID lookup
         columns: Site Main ID | Site Secondary ID | VGi Job number | Site Address | Site City
         searched newest year first; a new year's sheet is picked up automatically
     "Staff List"      Employee | Title       -> crew dropdowns
     "Customer List"   one column of names    -> customer dropdown

   Inside Teams the person is signed in through Teams (Microsoft "nested app
   authentication"), the app downloads the workbook with read-only access and
   keeps a copy on the device so lookups still work with no signal. Outside
   Teams the saved copy is used if there is one; otherwise everything is typed
   by hand.

   Needs settings.js and app.js loaded first.
   ===================================================================== */
(function () {
  const MSAL_URL = 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@5.9.0/+esm';
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const CACHE_KEY = 'vgSafetyData.v1';
  const { kv } = window.SafetyApp;

  const cfg = window.SAFETY_SETTINGS || {};
  const ms = cfg.microsoft || {};
  const src = Object.assign({ library: 'Documents', jobSheets: 'Job Information', staffSheet: 'Staff List', customerSheet: 'Customer List' }, cfg.data || {});
  const configured = !!(ms.clientId && ms.tenantId && src.siteUrl && src.filePath);
  const SCOPES = [ms.scope || 'Sites.Read.All'];   // read-only; IT may use 'Sites.Selected' instead

  // lower case, single spaces, accents removed
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
  // Spreadsheet placeholders that mean "nothing here": 0, a lone space, "_Drop Down Selection"
  const clean = v => {
    const s = String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
    return (!s || s === '0' || s.startsWith('_')) ? '' : s;
  };

  /* ---------------- State ---------------- */
  const listeners = [];
  const D = {
    configured,
    jobs: [],          // [{ year, mainId, secondaryId, jobNo, address, city, row }]
    years: [],         // job sheet years found, newest first
    staff: [],         // [{ name, title }]
    customers: [],     // ['Bell Mobility', ...]
    updatedAt: null,
    state: configured ? 'idle' : 'off',   // off | idle | loading | ready | signin | outside | error
    detail: '',
    offline: false,
    ready() { return D.state === 'ready' || (D.offline && D.jobs.length + D.staff.length + D.customers.length > 0); },
    onChange(fn) { listeners.push(fn); fn(D); },
    /* Site Main ID lookup: newest year first. Returns { year, matches } or null. */
    findSite(id) {
      const k = norm(id);
      if (!k) return null;
      for (const y of D.years) {
        const m = D.jobs.filter(j => j.year === y && norm(j.mainId) === k);
        if (m.length) return { year: y, matches: m.slice().reverse() };   // later rows first (newest job numbers)
      }
      return null;
    },
    findStaff(name) { const k = norm(name); return k ? D.staff.find(p => norm(p.name) === k) || null : null; },
    start, refresh
  };
  const emit = () => listeners.forEach(fn => { try { fn(D); } catch (e) { if (window.console) console.error(e); } });
  const set = (state, detail) => { D.state = state; D.detail = detail || ''; emit(); };
  const sourceKey = () => [src.siteUrl, src.library, src.filePath].join('|');
  function apply(d) { D.jobs = d.jobs; D.years = d.years; D.staff = d.staff; D.customers = d.customers; }
  function loadCache() {
    try {
      const c = JSON.parse(kv.get(CACHE_KEY) || 'null');
      if (c && c.source === sourceKey() && c.data) { apply(c.data); D.updatedAt = c.at; return true; }
    } catch (e) {}
    return false;
  }

  /* ---------------- Start (once per page) ---------------- */
  let inTeams = false;
  async function start(teamsCtxPromise) {
    if (!configured) { emit(); return; }
    const hadCache = loadCache();
    if (hadCache) { D.offline = true; set('ready'); }
    const ctx = await teamsCtxPromise;
    inTeams = !!ctx;
    if (!inTeams) { if (!hadCache) set('outside'); return; }
    refresh(false);
  }

  /* ---------------- Sign-in through Teams ---------------- */
  let pcaPromise = null;
  function getPca() {
    if (!pcaPromise) {
      pcaPromise = import(MSAL_URL).then(m => m.createNestablePublicClientApplication({
        auth: { clientId: ms.clientId, authority: 'https://login.microsoftonline.com/' + ms.tenantId },
        cache: { cacheLocation: 'localStorage' }
      }));
      pcaPromise.catch(() => { pcaPromise = null; });
    }
    return pcaPromise;
  }
  async function getToken(interactive) {
    const pca = await getPca();
    const u = window.SafetyApp.teamsUser();
    const req = { scopes: SCOPES };
    if (u && u.upn) req.loginHint = u.upn;
    try { return (await pca.acquireTokenSilent(req)).accessToken; }
    catch (e) {
      if (!interactive) { const err = new Error('signin'); err.signin = true; err.cause = e; throw err; }
      return (await pca.acquireTokenPopup(req)).accessToken;
    }
  }

  /* ---------------- Download the workbook ---------------- */
  async function graph(path, token) {
    const r = await fetch(GRAPH + path, { headers: { Authorization: 'Bearer ' + token } });
    if (!r.ok) {
      const err = new Error(r.status === 404 ? 'not found' : r.status === 403 ? 'no access' : 'error ' + r.status);
      err.status = r.status; throw err;
    }
    return r.json();
  }
  const encPath = p => p.split('/').filter(Boolean).map(encodeURIComponent).join('/');
  async function download(token) {
    const u = new URL(src.siteUrl);
    const sitePath = u.pathname.replace(/\/+$/, '');
    const site = await graph(`/sites/${u.hostname}:${sitePath || '/'}?$select=id`, token);
    let driveBase = `/sites/${site.id}/drive`;
    const lib = (src.library || 'Documents').trim();
    if (lib && !/^(documents|shared documents)$/i.test(lib)) {
      const drives = await graph(`/sites/${site.id}/drives?$select=id,name,webUrl`, token);
      const d = (drives.value || []).find(x => norm(x.name) === norm(lib) || norm(decodeURIComponent((x.webUrl || '').split('/').pop())) === norm(lib));
      if (!d) { const err = new Error('library not found'); err.status = 404; throw err; }
      driveBase = `/drives/${d.id}`;
    }
    const item = await graph(`${driveBase}/root:/${encPath(src.filePath)}?$select=id,name,@microsoft.graph.downloadUrl`, token);
    const url = item['@microsoft.graph.downloadUrl'];
    if (!url) throw new Error('no download link');
    const r = await fetch(url);
    if (!r.ok) throw new Error('download failed ' + r.status);
    return r.arrayBuffer();
  }

  let running = null;
  function refresh(interactive) {
    if (!configured || !inTeams) return Promise.resolve();   // sign-in only works inside Teams
    if (running) return running;
    if (!D.updatedAt) set('loading');
    running = (async () => {
      try {
        const token = await getToken(interactive);
        const data = await readWorkbook(await download(token));
        apply(data); D.updatedAt = new Date().toISOString(); D.offline = false;
        kv.set(CACHE_KEY, JSON.stringify({ at: D.updatedAt, source: sourceKey(), data }));
        set('ready');
      } catch (e) {
        if (e.signin) set('signin');
        else if (navigator.onLine === false || e instanceof TypeError) { D.offline = true; D.updatedAt ? set('ready') : set('error', 'no connection'); }
        else set('error', explain(e));
        if (window.console) console.warn('[company data]', e, e && e.cause);
      } finally { running = null; }
    })();
    return running;
  }
  function explain(e) {
    if (e.status === 403) return 'you don’t have access to the company data file';
    if (e.status === 404) return 'the company data file wasn’t found. Check settings.js';
    if (e.errorCode) return e.errorCode.replace(/_/g, ' ');
    return e.message || String(e);
  }

  /* ================= Minimal .xlsx reader (no library) ================= */
  async function unzip(buf) {
    const dv = new DataView(buf), u8 = new Uint8Array(buf), dec = new TextDecoder();
    let eocd = -1;
    for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 66000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error('not an Excel (.xlsx) file');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const files = {};
    for (let n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      files[dec.decode(u8.subarray(p + 46, p + 46 + nlen))] = { method, csize, off };
      p += 46 + nlen + elen + clen;
    }
    return async function read(name) {
      const f = files[name] || files[Object.keys(files).find(k => k.toLowerCase() === name.toLowerCase())];
      if (!f) return null;
      const start = f.off + 30 + dv.getUint16(f.off + 26, true) + dv.getUint16(f.off + 28, true);
      const data = u8.subarray(start, start + f.csize);
      if (f.method === 0) return dec.decode(data);
      if (f.method !== 8) throw new Error('unsupported compression');
      if (typeof DecompressionStream === 'undefined') throw new Error('this device is too old to read the Excel file');
      return await new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).text();
    };
  }
  const xml = s => new DOMParser().parseFromString(s, 'application/xml');
  const kids = (el, tag) => Array.from(el.getElementsByTagNameNS('*', tag));
  const colIndex = ref => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + (ch.toUpperCase().charCodeAt(0) - 64); return n - 1; };

  async function readSheets(buf) {
    const read = await unzip(buf);
    const wb = xml(await read('xl/workbook.xml') || '');
    const rels = xml(await read('xl/_rels/workbook.xml.rels') || '');
    const relMap = {};
    kids(rels, 'Relationship').forEach(r => { relMap[r.getAttribute('Id')] = r.getAttribute('Target'); });
    const ssText = await read('xl/sharedStrings.xml');
    const shared = ssText ? kids(xml(ssText), 'si').map(si =>
      kids(si, 't').filter(t => !(t.parentNode && t.parentNode.localName === 'rPh') && !(t.parentNode && t.parentNode.parentNode && t.parentNode.parentNode.localName === 'rPh'))
        .map(t => t.textContent).join('')) : [];
    const sheets = [];
    for (const s of kids(wb, 'sheet')) {
      const rid = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      let t = relMap[rid] || '';
      t = t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, '');
      const text = await read(t);
      const rows = !text ? [] : kids(xml(text), 'row').map(r => {
        const out = [];
        kids(r, 'c').forEach((c, i) => {
          const ref = c.getAttribute('r'), ty = c.getAttribute('t');
          let v = '';
          if (ty === 'inlineStr') v = kids(c, 't').map(x => x.textContent).join('');
          else { const ve = kids(c, 'v')[0]; v = ve ? ve.textContent : ''; if (ty === 's') v = shared[+v] || ''; }
          out[ref ? colIndex(ref) : i] = v;
        });
        return out;
      });
      sheets.push({ name: s.getAttribute('name') || '', rows });
    }
    return sheets;
  }

  /* Find the header row (within the first 10) whose cells match the wanted columns */
  function header(rows, want) {
    for (let h = 0; h < Math.min(rows.length, 10); h++) {
      const head = (rows[h] || []).map(norm);
      const cols = {};
      for (const [key, test] of Object.entries(want)) cols[key] = head.findIndex(x => x && test(x));
      if (Object.values(cols).some(i => i >= 0)) return { h, cols };
    }
    return null;
  }
  const cell = (r, i) => (i >= 0 && r ? clean(r[i]) : '');

  async function readWorkbook(buf) {
    const sheets = await readSheets(buf);
    const byName = n => sheets.find(s => norm(s.name) === norm(n));

    /* Job Information YYYY sheets */
    const prefix = norm(src.jobSheets);
    const jobs = [], years = [];
    for (const s of sheets) {
      const n = norm(s.name);
      if (!n.startsWith(prefix)) continue;
      const m = n.match(/(\d{4})/);
      const year = m ? +m[1] : 0;
      const hd = header(s.rows, {
        mainId: x => x.includes('main') && x.includes('id'),
        secondaryId: x => x.includes('secondary'),
        jobNo: x => x.includes('job') && (x.includes('number') || x.includes('no') || x.includes('#')),
        address: x => x.includes('address'),
        city: x => x.includes('city') || x.includes('town')
      });
      if (!hd || hd.cols.mainId < 0) continue;
      years.push(year);
      s.rows.slice(hd.h + 1).forEach((r, i) => {
        const mainId = cell(r, hd.cols.mainId);
        if (!mainId) return;
        jobs.push({ year, mainId, secondaryId: cell(r, hd.cols.secondaryId), jobNo: cell(r, hd.cols.jobNo), address: cell(r, hd.cols.address), city: cell(r, hd.cols.city), row: hd.h + i + 2 });
      });
    }
    years.sort((a, b) => b - a);

    /* Staff List */
    const staff = [];
    const st = byName(src.staffSheet);
    if (st) {
      const hd = header(st.rows, {
        name: x => /^(empl|emplo|name|full name|worker|staff)/.test(x) || x === 'employee name',
        first: x => x.startsWith('first'),
        last: x => x.startsWith('last') || x === 'surname',
        title: x => /title|position|role|trade|classification/.test(x)
      });
      const name0 = hd ? hd.cols.name : 0, title0 = hd ? hd.cols.title : 1;
      const seen = new Set();
      (st.rows.slice(hd ? hd.h + 1 : 1)).forEach(r => {
        let name = name0 >= 0 ? cell(r, name0) : [cell(r, hd.cols.first), cell(r, hd.cols.last)].filter(Boolean).join(' ');
        if (!name || seen.has(norm(name))) return;
        seen.add(norm(name));
        staff.push({ name, title: cell(r, title0) });
      });
      staff.sort((a, b) => a.name.localeCompare(b.name));
    }

    /* Customer List: first column, after its heading */
    const customers = [];
    const cu = byName(src.customerSheet);
    if (cu) {
      const seen = new Set();
      cu.rows.forEach((r, i) => {
        const v = cell(r, 0);
        if (!v || (i === 0 && /customer/i.test(v)) || seen.has(norm(v))) return;
        seen.add(norm(v)); customers.push(v);
      });
    }

    if (!jobs.length && !staff.length && !customers.length) throw new Error('no job, staff or customer sheets found in the workbook');
    return { jobs, years, staff, customers };
  }

  D._readWorkbook = readWorkbook;   // used by tests
  window.SafetyApp.Data = D;
})();
