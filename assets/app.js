/* =====================================================================
   Shared code for every page: job details, form status, Teams setup.
   The job is entered once on the home page and read by each form.
   ===================================================================== */
(function () {
  const JOB_KEY = 'vgSafetyJob.v1';
  const FLOW_KEY = 'vgSafetyFlow.v1';

  /* Browser storage, with an in-memory fallback if it's blocked */
  const mem = {};
  const kv = {
    get(k) { try { const v = localStorage.getItem(k); if (v != null) return v; } catch (e) {} return mem[k] != null ? mem[k] : null; },
    set(k, v) { mem[k] = v; try { localStorage.setItem(k, v); } catch (e) {} },
    del(k) { delete mem[k]; try { localStorage.removeItem(k); } catch (e) {} }
  };

  const pad = n => String(n).padStart(2, '0');
  const localDate = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const localDateTime = d => `${localDate(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
  const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));

  /* ---------------- Job ---------------- */
  function blankJob() {
    return {
      id: uid(),
      createdAt: new Date().toISOString(),
      site: '', address: '', customer: 'Bell Mobility', contractor: 'The Violin Group',
      when: localDateTime(new Date()), jobNo: '', scope: '',
      leadName: '', leadTitle: '',
      crew: [],                 // [{ name, title }] — not including the crew lead
      musterPoint: '', emergencyContact: '',
      activities: [],           // activity ids from docs.js
      done: {}                  // docId -> { at, fileName }
    };
  }
  const Job = {
    load() {
      try { const j = JSON.parse(kv.get(JOB_KEY) || 'null'); if (j && j.id) return Object.assign(blankJob(), j); } catch (e) {}
      return null;
    },
    get() { return Job.load() || blankJob(); },
    save(job) { job.updatedAt = new Date().toISOString(); kv.set(JOB_KEY, JSON.stringify(job)); return job; },
    clear() { kv.del(JOB_KEY); },
    blank: blankJob,
    /* Everyone on site: crew lead first, then the crew list */
    workers(job) {
      const list = [];
      if (job.leadName) list.push({ name: job.leadName, title: job.leadTitle || '' });
      (job.crew || []).forEach(w => { if (w.name || w.title) list.push({ name: w.name || '', title: w.title || '' }); });
      return list;
    },
    isRequired(job, doc) {
      if (doc.optional) return false;
      if (doc.always) return true;
      return (doc.when || []).some(a => (job.activities || []).includes(a));
    },
    /* Forms call this after a successful submit */
    markDone(docId, info) {
      const job = Job.get();
      job.done[docId] = Object.assign({ at: new Date().toISOString() }, info || {});
      Job.save(job);
    },
    /* The minimum a form needs before it can be opened */
    missing(job) {
      const m = [];
      if (!job.site.trim()) m.push('site');
      if (!job.when) m.push('when');
      if (!job.leadName.trim()) m.push('leadName');
      return m;
    }
  };

  /* ---------------- Power Automate flow address ----------------
     Arrives in the Teams tab address (#flow=…) from config.html, so it never sits in this public code. */
  const validFlow = u => /^https:\/\/\S+$/i.test((u || '').trim());
  (function readFlowFromUrl() {
    const m = (location.hash + '&' + location.search).match(/[#?&]flow=([^&]+)/);
    if (m) { const u = decodeURIComponent(m[1]); if (validFlow(u)) kv.set(FLOW_KEY, u.trim()); }
  })();
  const Flow = { url: () => kv.get(FLOW_KEY) || '', set: u => kv.set(FLOW_KEY, u), clear: () => kv.del(FLOW_KEY), valid: validFlow };

  /* ---------------- Formatting ---------------- */
  const fmt = {
    when(v) {
      if (!v) return '';
      const d = new Date(v);
      return isNaN(d) ? v : d.toLocaleString('en-CA', { year: 'numeric', month: 'long', day: 'numeric', hour: 'numeric', minute: '2-digit' });
    },
    day(v) {
      if (!v) return '';
      const d = new Date(v.length <= 10 ? v + 'T12:00' : v);
      return isNaN(d) ? v : d.toLocaleDateString('en-CA', { weekday: 'short', month: 'short', day: 'numeric' });
    },
    time(iso) {
      const d = new Date(iso);
      return isNaN(d) ? '' : d.toLocaleTimeString('en-CA', { hour: 'numeric', minute: '2-digit' });
    }
  };
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  /* ---------------- Teams (optional; pages work in a normal browser too) ---------------- */
  let teamsUser = null;
  function initTeams(onUser) {
    const s = document.createElement('script');
    s.src = 'https://unpkg.com/@microsoft/teams-js@2.22.0/dist/MicrosoftTeams.min.js';
    s.async = true;
    s.onload = async () => {
      const mt = window.microsoftTeams;
      if (!mt || !mt.app) return;
      try {
        await mt.app.initialize();
        const theme = t => document.documentElement.setAttribute('data-theme', t === 'default' ? 'light' : 'dark');
        const ctx = await mt.app.getContext();
        if (ctx && ctx.app && ctx.app.theme) theme(ctx.app.theme);
        mt.app.registerOnThemeChangeHandler(theme);
        const u = ctx && ctx.user;
        if (u) { teamsUser = { name: u.displayName || '', upn: u.userPrincipalName || u.loginHint || '' }; if (onUser) onUser(teamsUser); }
        mt.app.notifySuccess();
      } catch (e) { /* not inside Teams */ }
    };
    document.head.appendChild(s);
  }

  window.SafetyApp = {
    Job, Flow, fmt, esc, kv, initTeams,
    teamsUser: () => teamsUser,
    localDate, localDateTime, uid,
    docs: () => window.SAFETY_DOCS || [],
    activities: () => window.SAFETY_ACTIVITIES || [],
    doc: id => (window.SAFETY_DOCS || []).find(d => d.id === id) || null,
    docHref: d => d.page ? d.page : 'form.html?doc=' + encodeURIComponent(d.id)
  };
})();
