/* ITR Generator — main application */
(function () {
  'use strict';

  const CFG = window.ITR_CONFIG || {};
  const LIBS = {
    xlsx: 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
    exceljs: 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js'
  };
  const LS = { token: 'itr_token', user: 'itr_user', api: 'itr_api_url', banner: 'itr_install_banner_seen', landing: 'itr_landing' };

  /* ======================= utilities ======================= */
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));
  const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const money = (n) => {
    const v = r2(n);
    const s = Math.abs(v).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    return v < 0 ? '(₱' + s + ')' : '₱' + s;
  };
  const plain = (n) => r2(n).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const parseNum = (s) => {
    const t = String(s || '').replace(/[₱,\s]/g, '').replace(/^\((.*)\)$/, '-$1');
    const n = parseFloat(t);
    return isFinite(n) ? r2(n) : 0;
  };
  const sizeFmt = (b) => b > 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB';
  const plural = (n, w) => n + ' ' + w + (n === 1 ? '' : 's');
  const uid = () => Math.random().toString(36).slice(2, 10);
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function toast(msg, type = '', ms = 4200) {
    const el = document.createElement('div');
    el.className = 'toast ' + type;
    el.textContent = msg;
    $('#toasts').appendChild(el);
    setTimeout(() => { el.style.transition = 'opacity .3s'; el.style.opacity = '0'; setTimeout(() => el.remove(), 320); }, ms);
  }

  const loaded = {};
  function loadScript(src) {
    if (!loaded[src]) {
      loaded[src] = new Promise((res, rej) => {
        const s = document.createElement('script');
        s.src = src; s.async = true;
        s.onload = res; s.onerror = () => { delete loaded[src]; rej(new Error('Could not load ' + src.split('/').pop() + '. Check your connection.')); };
        document.head.appendChild(s);
      });
    }
    return loaded[src];
  }

  function openModal(id) { $(id).hidden = false; }
  function closeModal(id) { $(id).hidden = true; }
  document.addEventListener('click', (e) => {
    const c = e.target.closest('[data-close]');
    if (c) { const m = c.closest('.modal'); if (m) { m.hidden = true; m.dispatchEvent(new Event('closed')); } }
  });

  /* ======================= state ======================= */
  const thisYear = new Date().getFullYear();
  const S = {
    token: localStorage.getItem(LS.token),
    user: JSON.parse(localStorage.getItem(LS.user) || 'null'),
    clients: [], client: null,
    year: thisYear, quarter: null, form: null,
    opts: { method: 'itemized', rcitRate: 0.25, mcitApplies: false, mcitRate: 0.02, mixedIncome: false, sellerOf: 'Goods' },
    salesType: 'VAT',
    preset: JSON.parse(localStorage.getItem('itr_landing') || 'null'),
    files: { sales: [], prior: [], expenses: [] },
    R: null, busy: false
  };

  /* ======================= API ======================= */
  function apiUrl() {
    const saved = localStorage.getItem(LS.api);
    if (saved) return saved;
    return /^https:\/\//.test(CFG.API_URL || '') ? CFG.API_URL : '';
  }

  async function api(action, payload = {}) {
    const url = apiUrl();
    if (!url) { openSettings(); throw new Error('Set the backend Web App URL first.'); }
    let res;
    try {
      // text/plain avoids a CORS preflight, which Apps Script does not answer
      res = await fetch(url, { method: 'POST', body: JSON.stringify(Object.assign({ action, token: S.token }, payload)), redirect: 'follow' });
    } catch (e) {
      throw new Error('Cannot reach the backend. Check your internet connection and the Web App URL.');
    }
    let data;
    try { data = await res.json(); } catch (e) { throw new Error('Unexpected response from the backend (HTTP ' + res.status + '). Is the Web App deployed for "Anyone"?'); }
    if (!data.ok) {
      if (data.code === 'AUTH' && action !== 'login') { logout(true); }
      throw new Error(data.error || 'Request failed.');
    }
    return data;
  }

  /* ======================= intro & install ======================= */
  let introDone = false, deferredPrompt = null, bannerPending = false;
  const isStandalone = () => matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
  const isIOS = () => /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);

  function runIntro() {
    const intro = $('#intro');
    const seen = sessionStorage.getItem('itr_intro');
    const finish = () => {
      if (introDone) return;
      introDone = true;
      intro.classList.add('done');
      setTimeout(() => intro.remove(), 700);
      sessionStorage.setItem('itr_intro', '1');
      if (bannerPending || !localStorage.getItem(LS.banner)) setTimeout(maybeShowBanner, 900);
    };
    intro.addEventListener('click', finish);
    setTimeout(finish, seen ? 900 : 2800);
  }

  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    $('#menuInstall').hidden = false;
    if (!localStorage.getItem(LS.banner)) maybeShowBanner();
  });
  window.addEventListener('appinstalled', () => {
    localStorage.setItem(LS.banner, '1');
    hideBanner();
    toast('ITR Generator installed.', 'ok');
  });

  function maybeShowBanner() {
    if (isStandalone() || localStorage.getItem(LS.banner)) return;
    if (!introDone) { bannerPending = true; return; }
    const b = $('#installBanner');
    if (!b.hidden) return;
    if (deferredPrompt) {
      $('#installHint').textContent = 'Add it to your home screen or desktop for a faster, full-screen experience.';
      $('#installNow').textContent = 'Install';
    } else if (isIOS()) {
      $('#installHint').innerHTML = 'Tap <b>Share</b> <span aria-hidden="true">⎋</span> then <b>Add to Home Screen</b>.';
      $('#installNow').textContent = 'Got it';
    } else {
      $('#installHint').textContent = 'Use your browser menu → "Install app" or "Add to Home screen".';
      $('#installNow').textContent = 'Got it';
    }
    b.hidden = false;
    requestAnimationFrame(() => requestAnimationFrame(() => b.classList.add('show')));
  }
  function hideBanner() {
    const b = $('#installBanner');
    b.classList.remove('show');
    setTimeout(() => { b.hidden = true; }, 500);
  }
  async function doInstall() {
    if (deferredPrompt) {
      deferredPrompt.prompt();
      const choice = await deferredPrompt.userChoice.catch(() => ({}));
      deferredPrompt = null;
      $('#menuInstall').hidden = true;
      if (choice && choice.outcome === 'accepted') localStorage.setItem(LS.banner, '1');
    }
    localStorage.setItem(LS.banner, '1');
    hideBanner();
  }
  $('#installNow').addEventListener('click', doInstall);
  $('#installLater').addEventListener('click', () => { localStorage.setItem(LS.banner, '1'); hideBanner(); });
  $('#menuInstall').addEventListener('click', () => { closeUserMenu(); doInstall(); });

  if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) {
    window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
  }

  /* ======================= settings ======================= */
  function openSettings() {
    $('#apiUrlInput').value = apiUrl();
    $('#settingsMsg').textContent = '';
    $('#settingsMsg').className = 'form-note';
    openModal('#settingsModal');
  }
  $('#openSettings').addEventListener('click', openSettings);
  $('#menuSettings').addEventListener('click', () => { closeUserMenu(); openSettings(); });
  async function testUrl(url) {
    const msg = $('#settingsMsg');
    msg.className = 'form-note'; msg.textContent = 'Testing…';
    try {
      const res = await fetch(url, { method: 'POST', body: JSON.stringify({ action: 'ping' }) });
      const d = await res.json();
      if (!d.ok) throw new Error(d.error || 'Failed');
      msg.className = 'form-note ok'; msg.textContent = 'Connected ✓ ' + d.app + ' v' + d.version + ' · model ' + d.model;
      return true;
    } catch (e) {
      msg.className = 'form-note bad'; msg.textContent = 'Could not connect: ' + e.message;
      return false;
    }
  }
  $('#testApi').addEventListener('click', () => testUrl($('#apiUrlInput').value.trim()));
  $('#saveApi').addEventListener('click', async () => {
    const url = $('#apiUrlInput').value.trim();
    if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(url)) {
      $('#settingsMsg').className = 'form-note bad';
      $('#settingsMsg').textContent = 'That does not look like an Apps Script Web App URL.';
      return;
    }
    localStorage.setItem(LS.api, url);
    if (await testUrl(url)) setTimeout(() => closeModal('#settingsModal'), 700);
  });

  /* ======================= auth ======================= */
  function showLanding() {
    $('#appView').hidden = true;
    $('#landingView').hidden = false;
    syncLanding();
    window.scrollTo(0, 0);
  }
  function showApp() {
    $('#landingView').hidden = true;
    $('#appView').hidden = false;
    window.scrollTo(0, 0);
    const name = (S.user && S.user.name) || 'User';
    $('#userName').textContent = name;
    $('#userAvatar').textContent = name.trim().charAt(0).toUpperCase();
  }
  function openSignin() {
    const p = landingValues();
    $('#siChips').innerHTML = '<span>Tax Year ' + p.year + '</span><span>BIR Form ' + p.form + '</span><span>' + p.vat + '</span>';
    $('#loginError').textContent = '';
    openModal('#signinModal');
    setTimeout(() => $('#loginUser').focus(), 80);
  }

  /* ---------- landing ---------- */
  (function initLanding() {
    const sel = $('#ldYear');
    const back = CFG.MAX_YEARS_BACK || 6;
    for (let y = thisYear; y >= thisYear - back; y--) sel.insertAdjacentHTML('beforeend', '<option value="' + y + '">' + y + '</option>');
    const p = S.preset || {};
    sel.value = String(p.year && p.year >= thisYear - back ? p.year : thisYear);
    $('#ldFormSel').value = p.form === '1702Q' ? '1702Q' : '1701Q';
    setLdVat(p.vat === 'Non-VAT' ? 'Non-VAT' : 'VAT');
    $('#ldVat').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) setLdVat(b.dataset.v); });
    $('#ldForm').addEventListener('submit', (e) => {
      e.preventDefault();
      saveLanding();
      if (S.token) enterApp(); else openSignin();
    });
    $('#ldSigninBtn').addEventListener('click', () => {
      if (S.token) logout(false); else { saveLanding(); openSignin(); }
    });
  })();
  function setLdVat(v) {
    $('#ldVat').dataset.v = v;
    $$('#ldVat button').forEach((b) => { const on = b.dataset.v === v; b.classList.toggle('on', on); b.setAttribute('aria-checked', String(on)); });
  }
  function landingValues() {
    return { year: Number($('#ldYear').value), form: $('#ldFormSel').value, vat: $('#ldVat').dataset.v || 'VAT' };
  }
  function saveLanding() {
    S.preset = landingValues();
    localStorage.setItem(LS.landing, JSON.stringify(S.preset));
  }
  function syncLanding() {
    const p = S.preset;
    if (p) { $('#ldYear').value = String(p.year); $('#ldFormSel').value = p.form; setLdVat(p.vat); }
    const signed = !!S.token;
    $('#ldHello').hidden = !signed;
    $('#ldHello').textContent = signed ? 'Signed in as ' + ((S.user && S.user.name) || '') : '';
    $('span', $('#ldSigninBtn')).textContent = signed ? 'Sign out' : 'Sign in';
  }
  async function enterApp() {
    const p = S.preset || landingValues();
    const changed = S.year !== p.year || S.form !== p.form || S.salesType !== p.vat;
    S.year = p.year; S.form = p.form; S.salesType = p.vat;
    $('#yearSelect').value = String(S.year);
    if (S.form === '1702Q' || S.salesType === 'VAT') { if (S.opts.method === 'eight') S.opts.method = 'itemized'; }
    S.opts.mcitRate = defaultMcitRate();
    if (changed && S.R) clearResults();
    showApp();
    refresh();
    if (!S.clients.length) await loadClients();
    if (!S.client && matchMedia('(pointer: fine)').matches) setTimeout(() => $('#clientInput').focus(), 250);
  }
  function setPreset(k, v) {
    S.preset = Object.assign({ year: S.year, form: S.form, vat: S.salesType }, S.preset || {}, { [k]: v });
    localStorage.setItem(LS.landing, JSON.stringify(S.preset));
  }

  $('#pwToggle').addEventListener('click', () => {
    const i = $('#loginPass');
    i.type = i.type === 'password' ? 'text' : 'password';
    $('#pwToggle').textContent = i.type === 'password' ? 'Show' : 'Hide';
  });

  $('#loginForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const btn = $('#loginBtn');
    $('#loginError').textContent = '';
    btn.classList.add('loading');
    try {
      const d = await api('login', { username: $('#loginUser').value.trim(), password: $('#loginPass').value });
      S.token = d.token; S.user = d.user;
      localStorage.setItem(LS.token, d.token);
      localStorage.setItem(LS.user, JSON.stringify(d.user));
      $('#loginPass').value = '';
      closeModal('#signinModal');
      await enterApp();
      toast('Welcome, ' + d.user.name + '!', 'ok');
      if (d.mustChangePassword) openPassword('You signed in with a temporary password. Please set your own now.');
    } catch (err) {
      $('#loginError').textContent = err.message;
    } finally {
      btn.classList.remove('loading');
    }
  });

  /* ---------- sign-up ---------- */
  let staffCache = null;
  async function openSignup() {
    closeModal('#signinModal');
    $('#suError').textContent = '';
    openModal('#signupModal');
    const sel = $('#suStaff');
    sel.innerHTML = '<option value="">Loading staff…</option>';
    sel.disabled = true;
    try {
      const d = await api('staffList');
      staffCache = d;
      const open = d.staff.filter((x) => !x.registered).length;
      sel.innerHTML = '<option value="">— Select your name —</option>' + d.staff.map((x) =>
        '<option value="' + esc(x.name) + '"' + (x.registered ? ' disabled' : '') + '>' + esc(x.name) +
        (x.registered ? ' (already registered)' : ' · ' + x.clients + ' client' + (x.clients === 1 ? '' : 's')) + '</option>').join('');
      sel.disabled = false;
      $('#suStaffHint').className = 'field-hint';
      $('#suStaffHint').textContent = open ? open + ' of ' + d.staff.length + ' staff can still enroll.' : 'All assigned staff already have accounts.';
    } catch (e) {
      sel.innerHTML = '<option value="">Could not load staff</option>';
      $('#suError').textContent = e.message;
    }
  }
  $('#suStaff').addEventListener('change', (e) => {
    const x = staffCache && staffCache.staff.find((s) => s.name === e.target.value);
    const h = $('#suStaffHint');
    if (x) { h.className = 'field-hint ok'; h.textContent = 'Username: ' + x.name + ' · ' + x.clients + ' assigned client' + (x.clients === 1 ? '' : 's'); }
  });
  $('#toSignup').addEventListener('click', openSignup);
  $('#toSignin').addEventListener('click', () => { closeModal('#signupModal'); openSignin(); });
  $('#signupForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const err = $('#suError');
    err.textContent = '';
    const staff = $('#suStaff').value, name = $('#suName').value.trim(), p1 = $('#suPass').value, p2 = $('#suPass2').value;
    if (!staff) { err.textContent = 'Select your name from the assigned staff list.'; return; }
    if (name.length < 2) { err.textContent = 'Enter your full name.'; return; }
    if (p1.length < 8) { err.textContent = 'Password must be at least 8 characters.'; return; }
    if (p1 !== p2) { err.textContent = 'Passwords do not match.'; return; }
    const btn = $('#suBtn');
    btn.classList.add('loading');
    try {
      const d = await api('signup', { staff, fullName: name, password: p1 });
      S.token = d.token; S.user = d.user;
      localStorage.setItem(LS.token, d.token);
      localStorage.setItem(LS.user, JSON.stringify(d.user));
      $('#signupForm').reset();
      closeModal('#signupModal');
      if (!S.preset) saveLanding();
      await enterApp();
      toast('Welcome, ' + d.user.name + '! Your account is ready.', 'ok', 5000);
    } catch (ex) {
      err.textContent = ex.message;
    } finally {
      btn.classList.remove('loading');
    }
  });

  function logout(expired) {
    if (!expired && S.token) api('logout').catch(() => {});
    S.token = null; S.user = null;
    localStorage.removeItem(LS.token); localStorage.removeItem(LS.user);
    resetAll(true);
    showLanding();
    if (expired) toast('Your session has expired. Please sign in again.', 'err');
  }

  // user menu
  function closeUserMenu() { $('#userMenu').hidden = true; $('#userMenuBtn').setAttribute('aria-expanded', 'false'); }
  $('#userMenuBtn').addEventListener('click', (e) => {
    e.stopPropagation();
    const m = $('#userMenu'); m.hidden = !m.hidden;
    $('#userMenuBtn').setAttribute('aria-expanded', String(!m.hidden));
  });
  document.addEventListener('click', (e) => { if (!e.target.closest('.tb-user')) closeUserMenu(); });
  $('#menuLogout').addEventListener('click', () => { closeUserMenu(); logout(false); });
  $('#menuHome').addEventListener('click', () => { closeUserMenu(); showLanding(); });
  $('#homeBtn').addEventListener('click', showLanding);

  function openPassword(note) {
    $('#pwNote').textContent = note || '';
    $('#pwMsg').textContent = '';
    $('#passwordForm').reset();
    openModal('#passwordModal');
  }
  $('#menuPassword').addEventListener('click', () => { closeUserMenu(); openPassword(''); });
  $('#passwordForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const btn = e.submitter || $('#passwordForm .btn-primary');
    btn.classList.add('loading');
    try {
      await api('changePassword', { oldPassword: $('#pwOld').value, newPassword: $('#pwNew').value });
      closeModal('#passwordModal');
      toast('Password updated.', 'ok');
    } catch (err) { $('#pwMsg').textContent = err.message; }
    finally { btn.classList.remove('loading'); }
  });

  /* ======================= step 1: clients combo ======================= */
  async function loadClients() {
    $('#clientInput').placeholder = 'Loading your clients…';
    try {
      const d = await api('clients');
      S.clients = d.clients || [];
      $('#clientInput').placeholder = S.clients.length ? 'Type a name or TIN (' + S.clients.length + ' assigned)…' : 'No clients assigned to you yet';
    } catch (e) {
      $('#clientInput').placeholder = 'Could not load clients';
      toast(e.message, 'err');
    }
  }

  let comboIdx = -1, comboItems = [];
  function highlight(text, q) {
    if (!q) return esc(text);
    const i = text.toLowerCase().indexOf(q.toLowerCase());
    if (i < 0) return esc(text);
    return esc(text.slice(0, i)) + '<mark>' + esc(text.slice(i, i + q.length)) + '</mark>' + esc(text.slice(i + q.length));
  }
  function renderCombo(q) {
    const list = $('#clientList');
    const qq = (q || '').trim().toLowerCase();
    const qDigits = qq.replace(/\D/g, '');
    comboItems = S.clients.filter((c) => !qq || c.name.toLowerCase().includes(qq) || (qDigits.length >= 3 && c.tin.replace(/\D/g, '').includes(qDigits)) || c.id.toLowerCase() === qq).slice(0, 60);
    comboIdx = comboItems.length ? 0 : -1;
    list.innerHTML = comboItems.length
      ? comboItems.map((c, i) => '<li role="option" data-i="' + i + '" class="' + (i === 0 ? 'active' : '') + '"><div><div class="cn">' + highlight(c.name, q) + '</div><div class="ct">' + esc(c.tin) + '</div></div><div><span class="tag ' + (c.type === 'Corporation' ? 'corp' : 'ind') + '">' + c.type + '</span>' + (c.vat ? ' <span class="tag ' + (c.vat === 'VAT' ? 'vat' : 'nonvat') + '">' + c.vat + '</span>' : '') + '</div></li>').join('')
      : '<li class="empty">' + (S.clients.length ? 'No matching client' : 'No clients assigned to you') + '</li>';
    list.hidden = false;
    $('#clientInput').setAttribute('aria-expanded', 'true');
  }
  function closeCombo() { $('#clientList').hidden = true; $('#clientInput').setAttribute('aria-expanded', 'false'); }
  function moveCombo(d) {
    if (!comboItems.length) return;
    comboIdx = (comboIdx + d + comboItems.length) % comboItems.length;
    $$('#clientList li').forEach((li, i) => li.classList.toggle('active', i === comboIdx));
    const a = $('#clientList li.active'); if (a) a.scrollIntoView({ block: 'nearest' });
  }
  $('#clientInput').addEventListener('input', (e) => {
    $('#clientClear').hidden = !e.target.value;
    if (S.client && e.target.value !== S.client.name) selectClient(null, true);
    renderCombo(e.target.value);
  });
  $('#clientInput').addEventListener('focus', (e) => { if (!S.client) renderCombo(e.target.value); else e.target.select(); });
  $('#clientInput').addEventListener('keydown', (e) => {
    const open = !$('#clientList').hidden;
    if (e.key === 'ArrowDown') { e.preventDefault(); if (!open) renderCombo(e.target.value); else moveCombo(1); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); moveCombo(-1); }
    else if (e.key === 'Enter') { e.preventDefault(); if (open && comboIdx >= 0) selectClient(comboItems[comboIdx]); }
    else if (e.key === 'Escape') closeCombo();
  });
  $('#clientList').addEventListener('mousedown', (e) => {
    const li = e.target.closest('li[data-i]');
    if (li) { e.preventDefault(); selectClient(comboItems[Number(li.dataset.i)]); }
  });
  $('#clientCaret').addEventListener('click', () => {
    if ($('#clientList').hidden) { renderCombo(''); $('#clientInput').focus(); } else closeCombo();
  });
  $('#clientClear').addEventListener('click', () => { $('#clientInput').value = ''; selectClient(null); $('#clientClear').hidden = true; $('#clientInput').focus(); renderCombo(''); });
  document.addEventListener('click', (e) => { if (!e.target.closest('#clientCombo')) closeCombo(); });

  function selectClient(c, silent) {
    const changed = (S.client && S.client.id) !== (c && c.id);
    S.client = c;
    if (c) {
      $('#clientInput').value = c.name;
      $('#clientClear').hidden = false;
      closeCombo();
      $('#clientCard').innerHTML =
        '<div><small>TIN</small><b>' + esc(c.tin || '—') + '</b></div>' +
        '<div><small>Taxpayer type</small><b>' + esc(c.type) + '</b></div>' +
        '<div><small>VAT status</small><b>' + esc(c.vat || 'Not set — choose in step 2') + '</b></div>' +
        '<div><small>RDO</small><b>' + esc(c.rdo || '—') + '</b></div>' +
        (c.lob ? '<div style="grid-column:1/-1"><small>Line of business</small><b>' + esc(c.lob) + '</b></div>' : '') +
        (c.forms ? '<div style="grid-column:1/-1"><small>Enrolled BIR forms</small><b class="forms">' + esc(c.forms) + '</b></div>' : '');
      $('#clientCard').hidden = false;
      S.opts.sellerOf = c.sellerOf || 'Goods';
      if (changed) {
        // BIR Form and VAT status come from the landing page selection; fall back to the client record
        S.form = (S.preset && S.preset.form) || (c.type === 'Corporation' ? '1702Q' : '1701Q');
        S.salesType = (S.preset && S.preset.vat) || c.vat || 'VAT';
        S.opts.method = 'itemized'; clearResults(); clearFiles();
      }
    } else {
      $('#clientCard').hidden = true;
      if (!silent) $('#clientInput').value = '';
      clearResults();
    }
    refresh();
  }

  /* ======================= step 2: period & form ======================= */
  (function initYears() {
    const sel = $('#yearSelect');
    const back = CFG.MAX_YEARS_BACK || 6;
    for (let y = thisYear; y >= thisYear - back; y--) sel.insertAdjacentHTML('beforeend', '<option value="' + y + '">' + y + '</option>');
    sel.value = String(S.year);
    sel.addEventListener('change', () => { S.year = Number(sel.value); setPreset('year', S.year); onPeriodChange(); });
    $('#yearPrev').addEventListener('click', () => { if (sel.selectedIndex < sel.options.length - 1) { sel.selectedIndex++; sel.dispatchEvent(new Event('change')); } });
    $('#yearNext').addEventListener('click', () => { if (sel.selectedIndex > 0) { sel.selectedIndex--; sel.dispatchEvent(new Event('change')); } });
  })();

  $('#quarterSeg').addEventListener('click', (e) => {
    const b = e.target.closest('button'); if (!b) return;
    S.quarter = Number(b.dataset.v);
    onPeriodChange();
  });
  $('#formPick').addEventListener('click', (e) => {
    const b = e.target.closest('button'); if (!b) return;
    if (S.form !== b.dataset.v) { S.form = b.dataset.v; setPreset('form', S.form); S.opts.method = 'itemized'; clearResults(); }
    refresh();
  });

  $('#vatSeg').addEventListener('click', (e) => {
    const b = e.target.closest('button'); if (!b) return;
    if (S.salesType !== b.dataset.v) {
      S.salesType = b.dataset.v; setPreset('vat', S.salesType);
      if (S.salesType === 'VAT' && S.opts.method === 'eight') S.opts.method = 'itemized';
      clearResults();
    }
    refresh();
  });

  function onPeriodChange() {
    if (S.R) clearResults();
    S.opts.mcitRate = defaultMcitRate();
    refresh();
  }
  function defaultMcitRate() {
    // CREATE Act: 1% from Jul 1, 2020 to Jun 30, 2023; 2% otherwise
    const y = S.year, q = S.quarter || 1;
    if (y < 2020 || (y === 2020 && q <= 2)) return 0.02;
    if (y < 2023 || (y === 2023 && q <= 2)) return 0.01;
    return 0.02;
  }

  function renderOptions() {
    const box = $('#taxOptions');
    if (!S.form || !S.client) { box.innerHTML = ''; return; }
    const o = S.opts;
    let html = '';
    if (S.form === '1701Q') {
      const vat = S.salesType === 'VAT';
      if (vat && o.method === 'eight') o.method = 'itemized';
      html += '<label class="field"><span>Tax method</span><select data-opt="method">' +
        '<option value="itemized"' + (o.method === 'itemized' ? ' selected' : '') + '>Graduated rates — Itemized deductions</option>' +
        '<option value="osd"' + (o.method === 'osd' ? ' selected' : '') + '>Graduated rates — OSD (40% of sales)</option>' +
        '<option value="eight"' + (o.method === 'eight' ? ' selected' : '') + (vat ? ' disabled' : '') + '>8% income tax option' + (vat ? ' (non-VAT only)' : '') + '</option>' +
        '</select></label>';
      if (o.method === 'eight') {
        html += '<label class="check"><input type="checkbox" data-opt="mixedIncome"' + (o.mixedIncome ? ' checked' : '') + ' /><span>Mixed income earner<small>Also earns compensation income: no ₱250,000 reduction</small></span></label>';
      }
    } else {
      html += '<label class="field"><span>Deduction method</span><select data-opt="method">' +
        '<option value="itemized"' + (o.method === 'itemized' ? ' selected' : '') + '>Itemized deductions</option>' +
        '<option value="osd"' + (o.method === 'osd' ? ' selected' : '') + '>OSD (40% of gross income)</option></select></label>';
      html += '<label class="field"><span>Income tax rate</span><select data-opt="rcitRate">' +
        [[0.25, '25% — Regular corporate rate'], [0.20, '20% — NTI ≤ ₱5M & assets ≤ ₱100M'], [0.10, '10% — Proprietary educational / hospital'], [0.30, '30% — Pre-CREATE (before Jul 2020)']]
          .map(([v, l]) => '<option value="' + v + '"' + (o.rcitRate === v ? ' selected' : '') + '>' + l + '</option>').join('') + '</select></label>';
      html += '<label class="check"><input type="checkbox" data-opt="mcitApplies"' + (o.mcitApplies ? ' checked' : '') + ' /><span>Subject to MCIT<small>From the 4th taxable year of operations</small></span></label>';
      if (o.mcitApplies) {
        html += '<label class="field"><span>MCIT rate</span><select data-opt="mcitRate">' +
          [[0.02, '2%'], [0.01, '1% (Jul 2020 – Jun 2023)']].map(([v, l]) => '<option value="' + v + '"' + (o.mcitRate === v ? ' selected' : '') + '>' + l + '</option>').join('') + '</select></label>';
      }
    }
    if (o.method === 'itemized') {
      html += '<label class="field"><span>Seller of (for entertainment limit)</span><select data-opt="sellerOf">' +
        ['Goods', 'Services'].map((v) => '<option' + (o.sellerOf === v ? ' selected' : '') + '>' + v + '</option>').join('') + '</select></label>';
    }
    box.innerHTML = html;
  }
  $('#taxOptions').addEventListener('change', (e) => {
    const k = e.target.dataset.opt; if (!k) return;
    const v = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
    S.opts[k] = ['rcitRate', 'mcitRate'].includes(k) ? Number(v) : v;
    refresh();
    if (S.R) renderResults();
  });

  /* ======================= step 3: uploads ======================= */
  const ACCEPT = /\.(pdf|png|jpe?g|webp|gif|heic|heif|xlsx|xls|xlsm|csv)$/i;
  function kindOf(f) {
    const n = f.name.toLowerCase();
    if (/\.(xlsx|xls|xlsm|csv)$/.test(n) || /sheet|excel|csv/.test(f.type)) return 'xls';
    if (/\.pdf$/.test(n) || f.type === 'application/pdf') return 'pdf';
    return 'img';
  }
  function addFiles(target, list) {
    const arr = Array.from(list || []);
    for (const f of arr) {
      if (!ACCEPT.test(f.name) && !/^image\//.test(f.type)) { toast('Skipped "' + f.name + '": use a photo, PDF or Excel file.', 'err'); continue; }
      if (f.size > 45 * 1048576) { toast('"' + f.name + '" is larger than 45 MB.', 'err'); continue; }
      const item = { id: uid(), file: f, kind: kindOf(f) };
      if (target === 'prior') S.files.prior = [item];
      else S.files[target].push(item);
    }
    if (S.R) clearResults();
    renderFiles();
    refresh();
  }
  function renderFiles() {
    ['sales', 'prior', 'expenses'].forEach((t) => {
      $('[data-list="' + t + '"]').innerHTML = S.files[t].map((x) =>
        '<li><span class="fi ' + x.kind + '">' + ({ pdf: 'PDF', img: 'IMG', xls: 'XLS' }[x.kind]) + '</span>' +
        '<span class="fn" title="' + esc(x.file.name) + '">' + esc(x.file.name) + '</span>' +
        '<span class="fs">' + sizeFmt(x.file.size) + '</span>' +
        '<button class="rm" data-rm="' + t + ':' + x.id + '" aria-label="Remove">&times;</button></li>').join('');
    });
    $('#expHint').hidden = S.files.expenses.length > 0;
  }
  function clearFiles() { S.files = { sales: [], prior: [], expenses: [] }; renderFiles(); }

  document.addEventListener('click', (e) => {
    const b = e.target.closest('[data-rm]'); if (!b) return;
    const [t, id] = b.dataset.rm.split(':');
    S.files[t] = S.files[t].filter((x) => x.id !== id);
    if (S.R) clearResults();
    renderFiles(); refresh();
  });
  $$('.drop').forEach((d) => {
    const input = $('input', d);
    input.addEventListener('change', () => { addFiles(d.dataset.target, input.files); input.value = ''; });
    ['dragenter', 'dragover'].forEach((ev) => d.addEventListener(ev, (e) => { e.preventDefault(); d.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => d.addEventListener(ev, (e) => { e.preventDefault(); d.classList.remove('over'); }));
    d.addEventListener('drop', (e) => addFiles(d.dataset.target, e.dataTransfer.files));
  });
  $$('.btn-upload input').forEach((input) => input.addEventListener('change', () => { addFiles(input.dataset.target, input.files); input.value = ''; }));
  const expBox = $('#uplExpenses');
  ['dragenter', 'dragover'].forEach((ev) => expBox.addEventListener(ev, (e) => { e.preventDefault(); expBox.style.borderColor = '#1d4a7e'; }));
  ['dragleave', 'drop'].forEach((ev) => expBox.addEventListener(ev, (e) => { e.preventDefault(); expBox.style.borderColor = ''; }));
  expBox.addEventListener('drop', (e) => addFiles('expenses', e.dataTransfer.files));

  function renderUploadLabels() {
    const vat = S.salesType === 'VAT';
    $('#salesTitle').textContent = 'Sales return — ' + (vat ? '2550Q (VAT)' : '2551Q (Non-VAT)');
    $('#salesSub').textContent = vat ? 'Quarterly VAT return for the quarter (photo, PDF or Excel)' : 'Quarterly percentage tax return for the quarter (photo, PDF or Excel)';
    const q = S.quarter;
    const prior = $('#uplPrior');
    if (q === 1) {
      prior.classList.add('disabled');
      $('#priorTitle').textContent = 'Previous quarter ITR';
      $('#priorSub').textContent = 'Not needed for Q1. Cumulative figures start fresh each year.';
      $('#priorBadge').textContent = 'N/A';
      if (S.files.prior.length) { S.files.prior = []; renderFiles(); }
    } else {
      prior.classList.remove('disabled');
      $('#priorTitle').textContent = 'Previous quarter ITR' + (q ? ' — Q' + (q - 1) + ' ' + S.year + ' ' + (S.form || '') : '');
      $('#priorSub').textContent = 'Used for prior-quarter taxable income, payments and credits';
      $('#priorBadge').textContent = 'Recommended';
    }
  }
  /* ======================= refresh UI state ======================= */
  function refresh() {
    const s1 = !!S.client, s2 = s1 && !!S.quarter && !!S.form, s3 = s2 && S.files.sales.length > 0;
    $('#step2').classList.toggle('locked', !s1);
    $('#step3').classList.toggle('locked', !s2);
    $$('#quarterSeg button').forEach((b) => { const on = Number(b.dataset.v) === S.quarter; b.classList.toggle('on', on); b.setAttribute('aria-checked', String(on)); });
    $$('#formPick button').forEach((b) => {
      b.classList.toggle('on', b.dataset.v === S.form);
      const rec = S.client && ((S.client.type === 'Corporation') === (b.dataset.v === '1702Q'));
      let tag = $('.rec', b);
      if (rec && !tag) b.insertAdjacentHTML('beforeend', '<span class="rec">Matches taxpayer type</span>');
      if (!rec && tag) tag.remove();
    });
    $$('#vatSeg button').forEach((b) => {
      const on = b.dataset.v === S.salesType;
      b.classList.toggle('on', on); b.setAttribute('aria-checked', String(on));
      let tag = $('.rec', b);
      const rec = S.client && S.client.vat === b.dataset.v;
      if (rec && !tag) b.insertAdjacentHTML('beforeend', '<span class="rec">Client record</span>');
      if (!rec && tag) tag.remove();
    });
    renderOptions();
    renderUploadLabels();
    $('#analyzeBtn').disabled = !s3 || S.busy;
    const st = [s1, s2, s3, !!S.R];
    const firstOpen = st.indexOf(false);
    $$('#stepper li').forEach((li, i) => {
      li.classList.toggle('done', st[i]);
      li.classList.toggle('active', i === firstOpen);
    });
  }

  /* ======================= file preparation ======================= */
  function abToB64(buf) {
    const bytes = new Uint8Array(buf);
    let bin = '';
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(bin);
  }
  async function compressImage(file) {
    const max = CFG.IMAGE_MAX_DIMENSION || 2000;
    let src;
    try { src = await createImageBitmap(file); }
    catch (e) {
      src = await new Promise((res, rej) => {
        const img = new Image();
        img.onload = () => res(img);
        img.onerror = () => rej(new Error('Cannot read image "' + file.name + '". Try JPG or PNG.'));
        img.src = URL.createObjectURL(file);
      });
    }
    const w = src.width, h = src.height;
    const scale = Math.min(1, max / Math.max(w, h));
    const c = document.createElement('canvas');
    c.width = Math.round(w * scale); c.height = Math.round(h * scale);
    const ctx = c.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
    ctx.drawImage(src, 0, 0, c.width, c.height);
    const url = c.toDataURL('image/jpeg', CFG.IMAGE_QUALITY || 0.85);
    return { mime: 'image/jpeg', data: url.split(',')[1] };
  }
  async function prepareFile(item) {
    const f = item.file;
    if (item.kind === 'xls') {
      await loadScript(LIBS.xlsx);
      const buf = await f.arrayBuffer();
      const wb = window.XLSX.read(buf, { type: 'array', cellDates: true });
      const text = wb.SheetNames.map((n) => '### Sheet: ' + n + '\n' + window.XLSX.utils.sheet_to_csv(wb.Sheets[n], { blankrows: false })).join('\n\n');
      if (!text.replace(/[#\s,]|Sheet:/g, '')) throw new Error('The spreadsheet appears to be empty.');
      return { name: f.name, mime: f.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', text, data: f.size < 8 * 1048576 ? abToB64(buf) : '' };
    }
    if (item.kind === 'pdf') return { name: f.name, mime: 'application/pdf', data: abToB64(await f.arrayBuffer()) };
    const img = await compressImage(f);
    return { name: f.name.replace(/\.(heic|heif|png|webp|gif)$/i, '.jpg'), mime: img.mime, data: img.data };
  }

  async function pool(items, limit, fn) {
    const out = new Array(items.length);
    let next = 0;
    const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
      while (next < items.length) { const i = next++; out[i] = await fn(items[i], i); }
    });
    await Promise.all(workers);
    return out;
  }

  /* ======================= expense % modal ======================= */
  function askPercent() {
    return new Promise((resolve) => {
      const m = $('#expenseModal');
      const num = $('#pctNumber'), rng = $('#pctRange');
      const sync = (v) => {
        v = Math.min(100, Math.max(1, Number(v) || 0));
        num.value = v; rng.value = v;
        $$('#pctChips button').forEach((b) => b.classList.toggle('on', Number(b.dataset.p) === v));
        $('#pctPreview').textContent = 'Expenses = ' + v + '% of the revenue found in the sales return.';
      };
      sync(num.value);
      num.oninput = () => sync(num.value);
      rng.oninput = () => sync(rng.value);
      $('#pctChips').onclick = (e) => { const b = e.target.closest('button'); if (b) sync(b.dataset.p); };
      let done = false;
      $('#pctConfirm').onclick = () => { done = true; closeModal('#expenseModal'); resolve(Number(num.value)); };
      m.addEventListener('closed', () => { if (!done) resolve(null); }, { once: true });
      openModal('#expenseModal');
      setTimeout(() => num.focus(), 60);
    });
  }

  /* ======================= analyze ======================= */
  $('#analyzeBtn').addEventListener('click', analyze);

  function progItem(label, sub) {
    const li = document.createElement('li');
    li.className = 'wait';
    li.innerHTML = '<span class="st"></span><div class="pl"><b>' + esc(label) + '</b><small>' + esc(sub) + '</small></div>';
    $('#progList').appendChild(li);
    return {
      run: (s) => { li.className = ''; if (s) $('small', li).textContent = s; },
      ok: (s) => { li.className = 'ok'; if (s) $('small', li).textContent = s; },
      err: (s) => { li.className = 'err'; $('small', li).textContent = s; }
    };
  }

  async function analyze() {
    if (S.busy) return;
    let autoPct = null;
    if (!S.files.expenses.length) {
      autoPct = await askPercent();
      if (autoPct == null) return;
    }
    S.busy = true; refresh();
    $('#progList').innerHTML = '';
    $('#pmClose').hidden = true;
    $('.spark-orb').classList.remove('done');
    $('#pmTitle').textContent = 'Muse Spark is reading your documents';
    $('#pmSub').textContent = 'Extracting figures, classifying expenses and computing tax…';
    openModal('#progressModal');

    const salesDoc = S.salesType === 'VAT' ? 'sales_vat' : 'sales_nonvat';
    const jobs = []
      .concat(S.files.sales.map((x) => ({ x, docType: salesDoc, label: 'Sales return (' + (S.salesType === 'VAT' ? '2550Q' : '2551Q') + ')' })))
      .concat(S.files.prior.map((x) => ({ x, docType: 'prior_itr', label: 'Previous quarter ' + S.form })))
      .concat(S.files.expenses.map((x) => ({ x, docType: 'expenses', label: 'Expenses' })));
    jobs.forEach((j) => { j.p = progItem(j.x.file.name, j.label + ' · waiting'); });
    const autoP = autoPct != null ? progItem('Auto-generate expenses', autoPct + '% of revenue · waiting for revenue') : null;
    const taxP = progItem('Compute ' + S.form + ' tax due', 'waiting');

    const common = { clientId: S.client.id, year: S.year, quarter: S.quarter, form: S.form, vat: S.salesType };
    const results = await pool(jobs, CFG.PARALLEL_UPLOADS || 3, async (j) => {
      try {
        j.p.run(j.label + ' · preparing file…');
        const file = await prepareFile(j.x);
        j.p.run(j.label + ' · Muse Spark is reading…');
        const d = await api('extract', Object.assign({ docType: j.docType, file }, common));
        const info = j.docType === 'expenses' ? d.data.items.length + ' items · ' + money(d.data.document_total)
          : j.docType === 'prior_itr' ? 'Taxable income to date ' + money(d.data.total_taxable_income_to_date)
            : 'Revenue ' + money(d.data.revenue);
        j.p.ok(j.label + ' · ' + info);
        return { ok: true, job: j, d };
      } catch (e) {
        j.p.err(e.message);
        return { ok: false, job: j, error: e.message };
      }
    });

    const sales = results.filter((r) => r.ok && r.job.docType !== 'prior_itr' && r.job.docType !== 'expenses');
    const prior = results.filter((r) => r.ok && r.job.docType === 'prior_itr')[0];
    const exps = results.filter((r) => r.ok && r.job.docType === 'expenses');
    const failed = results.filter((r) => !r.ok);

    if (!sales.length) {
      fail('No revenue could be read from the sales return. Check the file and try again.');
      if (autoP) autoP.err('Skipped: no revenue');
      taxP.err('Skipped');
      return;
    }

    const R = buildResult(sales, prior, exps, failed);
    if (autoPct != null) {
      try {
        autoP.run('Muse Spark is allocating ' + autoPct + '% of ' + money(R.revenue) + '…');
        const a = await api('autoExpenses', Object.assign({ revenue: R.revenue, percent: autoPct }, common));
        R.items = a.items.map((it) => Object.assign({ file: 'Auto-generated (' + autoPct + '%)' }, it));
        R.expenseSource = 'auto';
        R.auto = { percent: autoPct, target: a.target, rationale: a.rationale, source: a.source };
        autoP.ok(a.items.length + ' accounts · ' + money(a.target) + (a.source === 'template' ? ' (standard template)' : ''));
      } catch (e) {
        autoP.err(e.message);
        R.warnings.push('Auto-generated expenses failed: ' + e.message);
      }
    }
    taxP.ok('Done');
    S.R = R;
    $('.spark-orb').classList.add('done');
    await sleep(failed.length ? 1600 : 700);
    closeModal('#progressModal');
    S.busy = false;
    renderResults(true);
    refresh();
    $('#step4').scrollIntoView({ behavior: 'smooth', block: 'start' });
    if (failed.length) toast(failed.length + ' file(s) could not be read. See the notes in the results.', 'err', 6000);

    function fail(msg) {
      $('.spark-orb').classList.add('done');
      $('#pmTitle').textContent = 'Analysis stopped';
      $('#pmSub').textContent = msg;
      $('#pmClose').hidden = false;
      S.busy = false; refresh();
    }
  }
  $('#pmClose').addEventListener('click', () => closeModal('#progressModal'));

  const tinDigits = (t) => String(t || '').replace(/\D/g, '').slice(0, 9);

  function buildResult(sales, prior, exps, failed) {
    const W = [];
    const c = S.client;
    const revenueLines = [];
    let revenue = 0;
    const vatT = { vatable_sales: 0, sales_to_government: 0, zero_rated_sales: 0, exempt_sales: 0 };
    sales.forEach((s) => {
      const d = s.d.data;
      revenue += d.revenue;
      if (s.job.docType === 'sales_vat') {
        Object.keys(vatT).forEach((k) => { vatT[k] += d[k] || 0; });
        const parts = (d.vatable_sales || 0) + (d.sales_to_government || 0) + (d.zero_rated_sales || 0) + (d.exempt_sales || 0);
        if (Math.abs(parts - d.revenue) > 1 && parts > 0) W.push('In ' + s.d.fileName + ', the sales breakdown (' + money(parts) + ') does not match total sales (' + money(d.revenue) + '). Total sales was used.');
      } else {
        (d.lines || []).forEach((l) => revenueLines.push({ label: (l.atc ? l.atc + ' — ' : '') + (l.description || 'Gross receipts'), amount: l.taxable_amount }));
        if (!(d.lines || []).length) revenueLines.push({ label: 'Gross sales/receipts (' + s.d.fileName + ')', amount: d.revenue });
      }
      if (d.taxpayer_tin && c.tin && tinDigits(d.taxpayer_tin) !== tinDigits(c.tin)) W.push('TIN on ' + s.d.fileName + ' (' + d.taxpayer_tin + ') does not match the client\'s TIN (' + c.tin + ').');
      if (d.year && d.quarter && (Number(d.year) !== S.year || Number(d.quarter) !== S.quarter)) W.push(s.d.fileName + ' appears to be for Q' + d.quarter + ' ' + d.year + ', but you selected Q' + S.quarter + ' ' + S.year + '.');
      if (d.confidence === 'low') W.push('Muse Spark has low confidence reading ' + s.d.fileName + (d.notes ? ': ' + d.notes : '.'));
    });
    if (sales.some((s) => s.job.docType === 'sales_vat')) {
      [['vatable_sales', 'VATable sales (private)'], ['sales_to_government', 'Sales to government'], ['zero_rated_sales', 'Zero-rated sales'], ['exempt_sales', 'Exempt sales']]
        .forEach(([k, l]) => { if (vatT[k]) revenueLines.unshift({ label: l, amount: r2(vatT[k]) }); });
      // keep original order
      revenueLines.sort((a, b) => ['VATable sales (private)', 'Sales to government', 'Zero-rated sales', 'Exempt sales'].indexOf(a.label) - ['VATable sales (private)', 'Sales to government', 'Zero-rated sales', 'Exempt sales'].indexOf(b.label));
      if (!revenueLines.length) revenueLines.push({ label: 'Total sales', amount: r2(revenue) });
    }

    let P = { priorYearExcess: 0, taxPaymentsPrev: 0, cwtPrev: 0, taxableIncomePrev: 0, grossIncomePrev: 0, cumulativeIncomePrev: 0 };
    if (prior) {
      const d = prior.d.data;
      P = {
        priorYearExcess: d.prior_year_excess_credits,
        taxPaymentsPrev: r2(d.tax_payments_previous_quarters + Math.max(0, d.tax_payable)),
        cwtPrev: r2(d.cwt_previous_quarters + d.cwt_this_quarter),
        taxableIncomePrev: d.total_taxable_income_to_date,
        grossIncomePrev: d.cumulative_gross_income_to_date,
        cumulativeIncomePrev: d.cumulative_gross_income_to_date
      };
      if (S.form === '1702Q' && !P.grossIncomePrev && d.sales_this_quarter) {
        W.push('Gross income for previous quarters (MCIT base) was not found on the previous ITR. Check the amount in the tax computation.');
      }
      if (d.year && d.quarter && (Number(d.year) !== S.year || Number(d.quarter) !== S.quarter - 1)) W.push('The previous ITR uploaded is for Q' + d.quarter + ' ' + d.year + '. Expected Q' + (S.quarter - 1) + ' ' + S.year + '.');
      if (d.form_detected && !String(d.form_detected).includes(S.form.slice(0, 4))) W.push('The previous ITR looks like form ' + d.form_detected + ', but you are preparing ' + S.form + '.');
      if (d.confidence === 'low') W.push('Muse Spark has low confidence reading the previous ITR' + (d.notes ? ': ' + d.notes : '.'));
      if (d.tax_payable > 0) W.push('info:Tax payments for previous quarters assume the Q' + (S.quarter - 1) + ' tax payable of ' + money(Math.max(0, d.tax_payable)) + ' was paid. Edit the amount if it was not.');
    } else if (S.quarter > 1) {
      W.push('No previous quarter ITR was uploaded. Enter the taxable income, payments and credits from previous quarters in the tax computation.');
    }

    const items = [];
    exps.forEach((e) => {
      e.d.data.items.forEach((it) => items.push(Object.assign({ file: e.d.fileName }, it)));
      if (e.d.data.confidence === 'low') W.push('Muse Spark has low confidence reading ' + e.d.fileName + (e.d.data.notes ? ': ' + e.d.data.notes : '.'));
    });
    failed.forEach((f) => W.push('Could not read ' + f.job.x.file.name + ': ' + f.error));

    return {
      revenueLines, revenue: r2(revenue), revenueAuto: r2(revenue),
      otherIncome: 0, cwtThisQ: 0, otherCredits: 0,
      prior: P,
      items, overrides: {}, open: {},
      expenseSource: 'documents', auto: null,
      docs: sales.concat(prior ? [prior] : [], exps).map((r) => ({ type: r.job.docType, file: r.d.fileName, data: r.d.data, url: r.d.archivedUrl })),
      warnings: W
    };
  }

  /* ======================= consolidation & computation ======================= */
  const normAcct = (s) => String(s).toLowerCase().replace(/&/g, 'and').replace(/[^a-z0-9]+/g, ' ').trim();
  const CAT_ORDER = { cost_of_sales: 0, operating_expense: 1, non_deductible: 2 };

  function accounts(R) {
    const map = new Map();
    R.items.forEach((it) => {
      const key = it.category + '|' + normAcct(it.account);
      if (!map.has(key)) map.set(key, { key, account: it.account, category: it.category, amount: 0, items: [] });
      const a = map.get(key);
      a.amount = r2(a.amount + it.amount);
      a.items.push(it);
    });
    Object.keys(R.overrides).forEach((k) => {
      if (map.has(k)) { const a = map.get(k); a.extracted = a.amount; a.amount = R.overrides[k]; a.edited = true; }
    });
    return Array.from(map.values()).sort((a, b) => CAT_ORDER[a.category] - CAT_ORDER[b.category] || b.amount - a.amount);
  }

  function model() {
    const R = S.R, o = S.opts;
    const accts = accounts(R);
    const sum = (cat) => r2(accts.filter((a) => a.category === cat).reduce((s, a) => s + a.amount, 0));
    const cos = sum('cost_of_sales'), opex = sum('operating_expense'), nonded = sum('non_deductible');
    const ent = window.ITRTax.entertainmentCap(accts, R.revenue, o.sellerOf);
    const deductibleOpex = r2(opex - (o.method === 'itemized' ? ent.disallowed : 0));
    const tax = window.ITRTax.compute({
      form: S.form, year: S.year, quarter: S.quarter, method: o.method,
      vatStatus: S.salesType, mixedIncome: o.mixedIncome,
      rcitRate: o.rcitRate, mcitApplies: o.mcitApplies, mcitRate: o.mcitRate,
      revenue: R.revenue, otherIncome: R.otherIncome, costOfSales: cos, opex: deductibleOpex,
      prior: R.prior, cwtThisQ: R.cwtThisQ, otherCredits: R.otherCredits
    });
    const methodLabel = S.form === '1701Q'
      ? ({ itemized: 'Graduated rates — Itemized deductions', osd: 'Graduated rates — OSD (40%)', eight: '8% income tax option' + (o.mixedIncome ? ' (mixed income)' : '') }[o.method])
      : (o.method === 'osd' ? 'OSD (40% of gross income)' : 'Itemized deductions') + ' · RCIT ' + (o.rcitRate * 100) + '%' + (o.mcitApplies ? ' · MCIT ' + (o.mcitRate * 100) + '%' : '');
    return {
      client: S.client, year: S.year, quarter: S.quarter, form: S.form, method: o.method, methodLabel,
      revenueLines: R.revenueLines, revenue: R.revenue, revenueEdited: R.revenue !== R.revenueAuto,
      accounts: accts, cos, opex, nonded, ent, deductibleOpex,
      grossProfit: r2(R.revenue - cos), netIncome: r2(R.revenue - cos - opex),
      tax, R, user: S.user
    };
  }

  /* ======================= results rendering ======================= */
  $('#newBtn').addEventListener('click', () => resetAll(false));
  $('#downloadBtn').addEventListener('click', downloadExcel);

  function clearResults() {
    S.R = null;
    $('#step4').hidden = true;
    $('#results').innerHTML = '';
  }
  function resetAll(full) {
    clearResults(); clearFiles();
    if (full) { S.clients = []; }
    selectClient(null);
    S.quarter = null; S.form = null;
    refresh();
    window.scrollTo({ top: 0, behavior: 'smooth' });
    if (!full) setTimeout(() => $('#clientInput').focus(), 300);
  }

  function inp(key, val, edited) {
    return '<input class="num-input' + (edited ? ' edited' : '') + '" inputmode="decimal" data-edit="' + esc(key) + '" value="' + plain(val) + '" aria-label="Amount" />';
  }

  function renderResults(animate) {
    const M = model(), R = S.R, T = M.tax;
    const focusKey = document.activeElement && document.activeElement.dataset ? document.activeElement.dataset.edit : null;
    $('#step4').hidden = false;
    $('#resultTitle').textContent = M.form + ' · Q' + M.quarter + ' ' + M.year + ' · ' + M.client.name;

    const warn = R.warnings.concat(T.warnings);
    if (M.ent.disallowed > 0) warn.push('Representation & entertainment of ' + money(M.ent.reTotal) + ' is more than the ' + (M.ent.rate * 100) + '% limit (' + money(M.ent.cap) + '). ' + money(M.ent.disallowed) + ' was disallowed as a deduction.');
    if (M.method === 'osd' || M.method === 'eight') warn.push('info:' + (M.method === 'osd' ? 'OSD' : 'The 8% option') + ' is selected, so the expenses below are for reference only. They do not reduce the tax due.');
    if (R.auto) warn.push('info:Expenses were estimated at ' + R.auto.percent + '% of revenue (' + money(R.auto.target) + ')' + (R.auto.rationale ? '. ' + R.auto.rationale : '.') + ' Keep supporting documents for every deduction you claim.');

    const kpi = (label, val, sub, hl) => '<div class="kpi' + (hl ? ' hl' : '') + '"><small>' + label + '</small><b data-count="' + val + '">' + money(animate ? 0 : val) + '</b>' + (sub ? '<em>' + sub + '</em>' : '') + '</div>';

    let h = '<div class="kpis">' +
      kpi('Revenue', M.revenue, 'Q' + M.quarter + ' ' + M.year) +
      kpi('Cost of sales', M.cos, plural(M.accounts.filter((a) => a.category === 'cost_of_sales').length, 'account')) +
      kpi('Operating expenses', M.opex, plural(M.accounts.filter((a) => a.category === 'operating_expense').length, 'account')) +
      kpi(T.taxPayable < 0 ? 'Overpayment' : 'Tax payable', T.taxPayable, 'Tax due ' + money(T.taxDue), true) +
      '</div>';

    if (warn.length) h += '<div class="alerts">' + warn.map((w) => w.startsWith('info:') ? '<div class="alert info">' + esc(w.slice(5)) + '</div>' : '<div class="alert">' + esc(w) + '</div>').join('') + '</div>';

    // ---- Income statement
    let is = '<tr class="sec"><td colspan="2">Revenue</td></tr>';
    M.revenueLines.forEach((l) => { is += '<tr><td class="ind">' + esc(l.label) + '</td><td class="amt">' + money(l.amount) + '</td></tr>'; });
    is += '<tr class="sub"><td>Net sales / revenue' + (M.revenueEdited ? '<span class="note">Edited (extracted: ' + money(R.revenueAuto) + ')</span>' : '') + '</td><td class="amt">' + inp('r:revenue', M.revenue, M.revenueEdited) + '</td></tr>';

    const acctRows = (cat) => M.accounts.filter((a) => a.category === cat).map((a) => {
      const open = !!R.open[a.key];
      let row = '<tr><td class="ind acc' + (open ? ' open' : '') + '" data-toggle="' + esc(a.key) + '"><span class="car">▶</span>' + esc(a.account) + '<span class="cnt">' + a.items.length + ' item' + (a.items.length > 1 ? 's' : '') + '</span>' +
        (a.edited ? '<span class="note">Edited (extracted: ' + money(a.extracted) + ')</span>' : '') + '</td><td class="amt">' + inp('a:' + a.key, a.amount, a.edited) + '</td></tr>';
      if (open) {
        row += a.items.map((it) => '<tr class="item"><td>' + esc([it.date, it.vendor, it.description].filter(Boolean).join(' · ') || '—') + '<span class="note">' + esc(it.file) + '</span></td><td class="amt">' + money(it.amount) + '</td></tr>').join('');
      }
      return row;
    }).join('');

    is += '<tr class="sec"><td colspan="2">Cost of sales / services</td></tr>';
    is += acctRows('cost_of_sales') || '<tr class="muted"><td class="ind" colspan="2">None</td></tr>';
    is += '<tr class="tot"><td>Total cost of sales</td><td class="amt">' + money(M.cos) + '</td></tr>';
    is += '<tr class="sub"><td>Gross profit</td><td class="amt">' + money(M.grossProfit) + '</td></tr>';
    is += '<tr class="sec"><td colspan="2">Operating expenses</td></tr>';
    is += acctRows('operating_expense') || '<tr class="muted"><td class="ind" colspan="2">None</td></tr>';
    is += '<tr class="tot"><td>Total operating expenses</td><td class="amt">' + money(M.opex) + '</td></tr>';
    if (M.ent.disallowed > 0 && M.method === 'itemized') is += '<tr class="muted"><td class="ind">Less: Entertainment over limit (non-deductible)</td><td class="amt">(' + money(M.ent.disallowed) + ')</td></tr><tr class="sub"><td>Allowable operating expenses</td><td class="amt">' + money(M.deductibleOpex) + '</td></tr>';
    if (M.nonded) {
      is += '<tr class="sec"><td colspan="2">Non-deductible (excluded from tax)</td></tr>' + acctRows('non_deductible');
    }
    is += '<tr class="grand"><td>Net income / (loss) per books</td><td class="amt">' + money(M.netIncome) + '</td></tr>';

    // ---- Tax computation
    let tx = '<tr class="sec"><td colspan="2">' + esc(T.schedule) + '</td></tr>';
    T.lines.forEach((l) => {
      const cls = l.type === 'total' ? 'tot' : l.type === 'subtotal' ? 'sub' : '';
      let amt;
      if (l.type === 'rate') amt = l.amount + '%';
      else if (l.editable) amt = inp((l.editable === 'otherIncome' ? 'r:' : 'p:') + l.editable, l.amount, false);
      else amt = money(l.amount);
      tx += '<tr class="' + cls + '"><td>' + esc(l.label) + (l.note ? '<span class="note">' + esc(l.note) + '</span>' : '') + '</td><td class="amt">' + amt + '</td></tr>';
    });
    tx += '<tr class="sec"><td colspan="2">Less: Tax credits / payments</td></tr>';
    T.credits.forEach((c) => {
      const key = c.key === 'cwtThisQ' || c.key === 'otherCredits' ? 'r:' + c.key : 'p:' + c.key;
      tx += '<tr><td class="ind">' + esc(c.label) + '</td><td class="amt">' + inp(key, c.amount, false) + '</td></tr>';
    });
    tx += '<tr class="tot"><td>Total tax credits / payments</td><td class="amt">' + money(T.totalCredits) + '</td></tr>';
    tx += '<tr class="grand"><td>' + (T.taxPayable < 0 ? 'Overpayment' : 'Tax payable') + '</td><td class="amt">' + money(T.taxPayable) + '</td></tr>';

    // ---- Documents
    const docRows = R.docs.map((d) => {
      const conf = d.data.confidence ? '<span class="conf ' + esc(d.data.confidence) + '">' + esc(d.data.confidence) + '</span>' : '';
      const t = { sales_vat: '2550Q', sales_nonvat: '2551Q', prior_itr: 'Prior ' + S.form, expenses: 'Expenses' }[d.type];
      const val = d.type === 'expenses' ? d.data.items.length + ' items · ' + money(d.data.document_total)
        : d.type === 'prior_itr' ? 'Tax due ' + money(d.data.tax_due) : money(d.data.revenue);
      return '<tr><td><b>' + esc(t) + '</b> · ' + esc(d.file) + conf + (d.data.notes ? '<span class="note">' + esc(d.data.notes) + '</span>' : '') +
        (d.url ? '<span class="note"><a href="' + esc(d.url) + '" target="_blank" rel="noopener">Open archived copy</a></span>' : '') + '</td><td class="amt">' + val + '</td></tr>';
    }).join('') || '<tr class="muted"><td colspan="2">—</td></tr>';

    h += '<div class="res-grid">' +
      '<div class="panel"><div class="panel-h"><h4>Income statement</h4><small>' + (R.expenseSource === 'auto' ? '<span class="badge-src">Auto-generated expenses</span>' : 'Tap an account to see its items') + '</small></div><div class="scroll-x"><table class="tbl">' + is + '</table></div></div>' +
      '<div><div class="panel"><div class="panel-h"><h4>Tax computation · BIR Form ' + M.form + '</h4><small>' + esc(M.methodLabel) + '</small></div><div class="scroll-x"><table class="tbl">' + tx + '</table></div></div>' +
      '<div class="panel"><div class="panel-h"><h4>Source documents</h4><small>Read by Muse Spark</small></div><div class="scroll-x"><table class="tbl">' + docRows + '</table></div></div></div>' +
      '</div>';

    $('#results').innerHTML = h;
    if (animate) countUp();
    if (focusKey) { const el = $('[data-edit="' + CSS.escape(focusKey) + '"]'); if (el) { el.focus(); el.select(); } }
  }

  function countUp() {
    $$('.kpi b[data-count]').forEach((el) => {
      const target = Number(el.dataset.count), t0 = performance.now(), dur = 900;
      const step = (t) => {
        const p = Math.min(1, (t - t0) / dur), e = 1 - Math.pow(1 - p, 3);
        el.textContent = money(target * e);
        if (p < 1) requestAnimationFrame(step); else el.textContent = money(target);
      };
      requestAnimationFrame(step);
    });
  }

  $('#results').addEventListener('click', (e) => {
    const t = e.target.closest('[data-toggle]');
    if (t && !e.target.closest('input')) { const k = t.dataset.toggle; S.R.open[k] = !S.R.open[k]; renderResults(); }
  });
  $('#results').addEventListener('focusin', (e) => { if (e.target.matches('.num-input')) setTimeout(() => e.target.select(), 0); });
  $('#results').addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.matches('.num-input')) e.target.blur(); });
  $('#results').addEventListener('change', (e) => {
    const el = e.target; if (!el.matches('.num-input')) return;
    const v = parseNum(el.value);
    const [ns, key] = [el.dataset.edit.slice(0, 1), el.dataset.edit.slice(2)];
    if (ns === 'a') {
      const base = accounts(Object.assign({}, S.R, { overrides: {} })).find((a) => a.key === key);
      if (base && r2(base.amount) === v) delete S.R.overrides[key]; else S.R.overrides[key] = v;
    } else if (ns === 'p') S.R.prior[key] = v;
    else if (ns === 'r') S.R[key] = v;
    setTimeout(() => renderResults(), 0);
  });

  /* ======================= Excel ======================= */
  async function downloadExcel() {
    if (!S.R) return;
    const btn = $('#downloadBtn');
    btn.classList.add('loading');
    try {
      await loadScript(LIBS.exceljs);
      const M = model();
      const name = await window.ITRExcel.download(M);
      toast('Downloaded ' + name, 'ok');
      api('saveReport', {
        clientId: S.client.id,
        report: {
          form: M.form, year: M.year, quarter: M.quarter, method: M.methodLabel, revenue: M.revenue, costOfSales: M.cos,
          opex: M.deductibleOpex, taxableToDate: M.tax.taxableToDate, taxDue: M.tax.taxDue, credits: M.tax.totalCredits,
          taxPayable: M.tax.taxPayable, expenseSource: S.R.expenseSource === 'auto' ? 'Auto ' + S.R.auto.percent + '%' : 'Documents'
        }
      }).catch(() => {});
    } catch (e) {
      toast(e.message, 'err');
    } finally {
      btn.classList.remove('loading');
    }
  }

  /* ======================= boot ======================= */
  async function boot() {
    runIntro();
    renderFiles();
    refresh();
    showLanding();
    if (S.token && S.user) loadClients();
    if (!apiUrl()) setTimeout(openSettings, introDone ? 0 : 3000);
  }
  // exposed for testing / debugging
  window.ITRApp = { state: S, model: () => model(), renderResults };
  boot();
})();
