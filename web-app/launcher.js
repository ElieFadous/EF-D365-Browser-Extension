(function () {
  'use strict';

  // ── Idempotency guard ──────────────────────────────────────────────
  if (window.__EF_PPT_LAUNCHER) return;
  window.__EF_PPT_LAUNCHER = true;

  // ── Base URL (GitHub Pages) ─────────────────────────────────────────
  const BASE_URL =
    (document.currentScript && document.currentScript.src
      ? document.currentScript.src
      : ''
    ).replace(/\/launcher\.js.*$/, '') ||
    'https://eliefadous.github.io/EF-D365-Browser-Extension/web-app';

  const CONFIG_KEY = 'ef_ppt_config';
  const API_VERSION_DEFAULT = 'v9.2';

  // ── Helpers ─────────────────────────────────────────────────────────
  function escHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function loadConfig() {
    // The bookmarklet embeds the current config as window.__EF_PPT_INIT_CONFIG
    // (base64-encoded JSON) so it is available in the D365 page context without
    // any cross-origin iframe (which D365 CSP would block).  On first run we
    // cache it to D365 localStorage so subsequent tool opens within the same
    // session are instant even without re-clicking the bookmark.
    if (window.__EF_PPT_INIT_CONFIG) {
      const cfg = window.__EF_PPT_INIT_CONFIG;
      delete window.__EF_PPT_INIT_CONFIG;
      try { localStorage.setItem(CONFIG_KEY, JSON.stringify(cfg)); } catch (_) {}
      return cfg;
    }
    try {
      const raw = localStorage.getItem(CONFIG_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (err) {
      console.error('[EF PPT] Failed to parse config:', err);
      return null;
    }
  }

  function saveConfig(cfg) {
    try {
      localStorage.setItem(CONFIG_KEY, JSON.stringify(cfg));
      return true;
    } catch (err) {
      console.error('[EF PPT] Failed to save config:', err);
      return false;
    }
  }

  function currentEnv(cfg) {
    if (!cfg || !Array.isArray(cfg.environments)) return null;
    const origin = window.location.origin;
    return cfg.environments.find(function (e) { return e.url === origin; }) || null;
  }

  function apiVersion(cfg) {
    return (cfg && cfg.settings && cfg.settings.apiVersion) || API_VERSION_DEFAULT;
  }

  /** Clone-related settings, with the same defaults the extension used. */
  function cloneSettings(cfg) {
    const s = (cfg && cfg.settings) || {};
    return {
      apiVersion: s.apiVersion || API_VERSION_DEFAULT,
      clonePrefix: s.clonePrefix || '',
      cloneWhitelist: s.cloneWhitelist || null,
      cloneLookupMode: s.cloneLookupMode || 'skip',
    };
  }

  const STARTER_CONFIG = {
    environments: [
      {
        id: 'myorg_crm4',
        name: 'PROD',
        url: 'https://myorg.crm4.dynamics.com',
        color: '#b91c1c',
        warn: true,
        powerAppsId: ''
      }
    ],
    settings: {
      apiVersion: 'v9.2',
      defaultAppUniqueName: '',
      includedApps: []
    }
  };

  // ── Icon SVGs ───────────────────────────────────────────────────────
  // viewBox is cropped tightly around the polygon (instead of the full
  // 64x64 canvas) so the bolt fills the button instead of looking tiny.
  const BOLT_SVG =
    '<svg viewBox="21 15 28 38" width="22" height="22" aria-hidden="true">' +
    '<polygon points="36,21 28,34 33,34 27,47 43,31 37,31" fill="#FFB900"/>' +
    '</svg>';

  const TOOL_ICONS = {
    metadata:
      '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M3.5 4.5v4c0 1.1 2.5 2 5.5 2s5.5-.9 5.5-2v-4"/><path d="M3.5 8.5v4c0 1.1 2.5 2 5.5 2s5.5-.9 5.5-2v-4"/><ellipse cx="9" cy="4.5" rx="5.5" ry="2"/></svg>',
    ribbon:
      '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="12" height="12" rx="2"/><path d="M3 7h12M7 3v12"/></svg>',
    'plugin-trace':
      '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M9 2v3M9 13v3M2 9h3M13 9h3"/><circle cx="9" cy="9" r="3.5"/></svg>',
    flows:
      '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><circle cx="4.5" cy="4.5" r="2"/><circle cx="13.5" cy="9" r="2"/><circle cx="4.5" cy="13.5" r="2"/><path d="M6.2 5.5L12 8M6.2 12.5L12 10"/></svg>',
    'data-sync':
      '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7a6 6 0 0 1 10-2.5L15 6"/><path d="M15 11a6 6 0 0 1-10 2.5L3 12"/><path d="M15 3v3h-3M3 15v-3h3"/></svg>'
  };

  const TOOLS = [
    { name: 'metadata',     label: 'Metadata Browser' },
    { name: 'ribbon',       label: 'Ribbon Buttons'   },
    { name: 'plugin-trace', label: 'Plugin Trace'     },
    { name: 'flows',        label: 'Flows'            },
    { name: 'data-sync',    label: 'Data Sync'        }
  ];

  const GOTO_TYPES = [
    { id: 'open-in',   label: 'Open In'   },
    { id: 'api',       label: 'API'       },
    { id: 'solutions', label: 'Solutions' },
    { id: 'app',       label: 'App'       },
    { id: 'security',  label: 'Security'  }
  ];

  // Record Details pane icons
  const USER_ICON_SVG   = '<svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="7" cy="4.5" r="2.2"/><path d="M2 12.5c0-2.76 2.24-4.5 5-4.5s5 1.74 5 4.5"/></svg>';
  const TEAM_ICON_SVG   = '<svg viewBox="0 0 18 13" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="4" cy="4" r="2"/><path d="M1 12c0-1.66 1.34-3 3-3s3 1.34 3 3"/><circle cx="9" cy="3.5" r="2.2"/><path d="M5.5 12c0-1.93 1.57-3.5 3.5-3.5s3.5 1.57 3.5 3.5"/><circle cx="14" cy="4" r="2"/><path d="M11 12c0-1.66 1.34-3 3-3s3 1.34 3 3"/></svg>';
  const COPY_ICON_SVG   = '<svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3.5" y="2" width="8" height="10" rx="1"/><path d="M2 4.5H1.5a.5.5 0 0 0-.5.5v7a.5.5 0 0 0 .5.5h7a.5.5 0 0 0 .5-.5V12"/></svg>';
  const CLONE_ICON_SVG  = '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" width="12" height="12"><rect x="5" y="4" width="9" height="11" rx="1.5"/><path d="M2 11V2.5A1.5 1.5 0 0 1 3.5 1H11"/></svg>';
  const CHECK_ICON_SVG  = '<svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="11" height="11"><path d="M2 7l3.5 3.5L12 3"/></svg>';

  // ── State ───────────────────────────────────────────────────────────
  let CFG = loadConfig();
  let selectedTargetId = null;
  let selectedType = 'open-in';
  let selectedAppUniqueName = null;

  // Record Details pane state — populated whenever the current D365 page is
  // showing a record (see parseRecordContext / ensureRecordDetailsLoaded).
  let _rdState          = null;  // { etn, id, status: 'loading'|'loaded'|'error', data, entitySetName, error }
  let _rdCloneTargetId  = null;  // selected environment id in the clone dropdown
  let _rdConnecting     = false; // true while "Connect Target Environment" is in flight
  let _rdConnectError   = null;
  let _rdCloning        = false; // true while a clone operation is running
  let _rdCloneResult    = null;  // { ok:true, newId, isCross, targetEnvUrl } | { ok:false, error }

  // ════════════════════════════════════════════════════════════════════
  //  CROSS-ENVIRONMENT BRIDGE
  // ════════════════════════════════════════════════════════════════════
  // Plain page JS cannot fetch a DIFFERENT D365 org's API with credentials —
  // that's blocked by CORS (D365 grants it to nobody; browser extensions only
  // bypass this via declared host_permissions, which a bookmarklet has no
  // equivalent for). Workaround: open a second tab for the target org and let
  // ITS OWN launcher.js instance do its own same-origin fetch there, relaying
  // the request/response between the two tabs via postMessage + the window
  // handle from window.open() (postMessage works cross-origin even though
  // direct DOM/fetch access does not). Used both by tool iframes (Data Sync's
  // "Connect Target Environment") via postMessage below, and internally by
  // this file's own Record Details clone logic via _connectTarget/_selfFetch.
  const _targetWindows      = new Map(); // origin -> WindowProxy (opened by this tab)
  const _targetReady        = new Set(); // origins confirmed alive (launcher-ready received)
  const _connectWaiters     = new Map(); // origin -> [{ id, source } | { resolve, reject }]
  const _pendingRelays      = new Map(); // fetch id -> original requester's window (e.source)
  const _selfPendingFetches = new Map(); // fetch id -> { resolve, timer } (our own in-flight requests)

  function _originOf(url) {
    try { return new URL(url).origin; } catch (_) { return ''; }
  }

  // If THIS tab was itself opened by another EF PPT tab (via the
  // connect-target flow below), announce that this origin's bridge is live.
  if (window.opener) {
    try {
      window.opener.postMessage({ __efppt: 'launcher-ready', origin: window.location.origin }, '*');
    } catch (_) { /* opener gone or restricted — harmless to skip */ }
  }

  function _resolveConnectWaiters(origin, ok, error) {
    if (ok) _targetReady.add(origin);
    const waiters = _connectWaiters.get(origin) || [];
    _connectWaiters.delete(origin);
    waiters.forEach(function (w) {
      if (w.resolve) {
        if (ok) w.resolve(); else w.reject(new Error(error || 'Could not connect to the target environment.'));
      } else if (w.source && w.source.postMessage) {
        w.source.postMessage({ __efppt: 'connect-target-result', id: w.id, ok: ok, error: error }, '*');
      }
    });
  }

  /** Opens (or reuses) a tab for `origin` and resolves once its own launcher.js confirms it's live. */
  function _connectTarget(origin) {
    return new Promise(function (resolve, reject) {
      if (!origin) { reject(new Error('Invalid target environment URL.')); return; }
      const existing = _targetWindows.get(origin);
      if (_targetReady.has(origin) && existing && !existing.closed) { resolve(); return; }
      const winName = 'efppt_target_' + origin.replace(/[^a-z0-9]/gi, '_');
      let win;
      try { win = window.open(origin, winName); }
      catch (err) { reject(new Error('Could not open a tab for the target environment: ' + err.message)); return; }
      if (!win) { reject(new Error('The browser blocked the new tab. Allow pop-ups for this site and try again.')); return; }
      _targetWindows.set(origin, win);
      const waiters = _connectWaiters.get(origin) || [];
      waiters.push({ resolve: resolve, reject: reject });
      _connectWaiters.set(origin, waiters);
      // Give the user time to switch tabs and click the EF PPT bookmark there.
      setTimeout(function () {
        if (!_targetReady.has(origin)) {
          _resolveConnectWaiters(origin, false,
            'Timed out waiting for the target tab. Click the EF PPT bookmark in the new tab, then try again.');
        }
      }, 120000);
    });
  }

  /** Performs the actual fetch against `url` (this tab's own origin) and normalises the result. */
  async function _doFetch(url, method, extraHeaders, body) {
    try {
      const headers = Object.assign(
        { Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' },
        extraHeaders || {}
      );
      const init = { method: method || 'GET', credentials: 'include', headers: headers };
      if (body != null && init.method !== 'GET' && init.method !== 'HEAD') {
        init.body = typeof body === 'string' ? body : JSON.stringify(body);
        if (!headers['Content-Type'] && !headers['content-type']) headers['Content-Type'] = 'application/json';
      }

      const res = await fetch(url, init);
      const text = await res.text();
      let data = text;
      if (text) { try { data = JSON.parse(text); } catch (_) { /* leave as text */ } }

      if (!res.ok) {
        return { ok: false, status: res.status, error: (data && data.error && data.error.message) || res.statusText || ('HTTP ' + res.status), data: data };
      }
      return { ok: true, status: res.status, data: data };
    } catch (err) {
      console.error('[EF PPT] Fetch failed:', err);
      return { ok: false, status: 0, error: String(err && err.message ? err.message : err) };
    }
  }

  /**
   * Fetches `url` on behalf of THIS launcher (not a tool iframe) — same-origin
   * directly, or relayed through a connected target tab for another org.
   * Resolves with the full { ok, status, data, error } envelope, never rejects.
   */
  function _selfFetchRaw(url, extraHeaders, method, body) {
    const urlOrigin = _originOf(url);
    if (!urlOrigin || urlOrigin === window.location.origin) {
      return _doFetch(url, method, extraHeaders, body);
    }
    return new Promise(function (resolve) {
      const targetWin = _targetWindows.get(urlOrigin);
      if (!_targetReady.has(urlOrigin) || !targetWin || targetWin.closed) {
        resolve({ ok: false, status: 0, error: 'Target environment (' + urlOrigin + ') is not connected. Use "Connect Target Environment" first.' });
        return;
      }
      const id = Math.random().toString(36).slice(2) + Date.now();
      const timer = setTimeout(function () {
        if (_selfPendingFetches.has(id)) {
          _selfPendingFetches.delete(id);
          resolve({ ok: false, status: 0, error: 'Timed out waiting for the target tab to respond.' });
        }
      }, 30000);
      _selfPendingFetches.set(id, { resolve: resolve, timer: timer });
      try {
        targetWin.postMessage({ __efppt: 'fetch', id: id, url: url, method: method || 'GET', headers: extraHeaders || {}, body: body }, '*');
      } catch (err) {
        clearTimeout(timer);
        _selfPendingFetches.delete(id);
        resolve({ ok: false, status: 0, error: 'Could not reach the target tab: ' + err.message });
      }
    });
  }

  /** Like _selfFetchRaw but resolves with the parsed body, rejecting with an Error on failure. */
  function _selfFetch(url, extraHeaders, method, body) {
    return _selfFetchRaw(url, extraHeaders, method, body).then(function (r) {
      if (r.ok) return r.data;
      throw new Error(r.error || (r.data && r.data.error && r.data.error.message) || ('HTTP ' + r.status));
    });
  }

  // ════════════════════════════════════════════════════════════════════
  //  FETCH BRIDGE  (serves requests from tool iframes/tabs)
  // ════════════════════════════════════════════════════════════════════
  window.addEventListener('message', async function (e) {
    const msg = e.data;
    if (!msg || typeof msg !== 'object') return;

    if (msg.__efppt === 'close-overlay') {
      closeTool();
      return;
    }

    // A target tab we opened has confirmed its own bridge is live.
    if (msg.__efppt === 'launcher-ready' && msg.origin) {
      _resolveConnectWaiters(msg.origin, true, null);
      return;
    }

    // A reply to a fetch WE issued (via _selfFetchRaw) from a connected target tab.
    if (msg.__efppt === 'fetch-result' && _selfPendingFetches.has(msg.id)) {
      const pending = _selfPendingFetches.get(msg.id);
      _selfPendingFetches.delete(msg.id);
      clearTimeout(pending.timer);
      pending.resolve(msg);
      return;
    }

    // A reply relayed back FROM a target tab we forwarded a fetch to —
    // forward it on to the ORIGINAL requester (e.g. the Data Sync iframe).
    if (msg.__efppt === 'fetch-result' && _pendingRelays.has(msg.id)) {
      const originalSource = _pendingRelays.get(msg.id);
      _pendingRelays.delete(msg.id);
      if (originalSource && originalSource.postMessage) originalSource.postMessage(msg, '*');
      return;
    }

    // A tool (Data Sync) asking us to open + connect a target environment tab.
    if (msg.__efppt === 'connect-target') {
      const origin = _originOf(msg.targetOrigin);
      _connectTarget(origin).then(
        function () {
          if (e.source && e.source.postMessage) e.source.postMessage({ __efppt: 'connect-target-result', id: msg.id, ok: true, error: null }, '*');
        },
        function (err) {
          if (e.source && e.source.postMessage) e.source.postMessage({ __efppt: 'connect-target-result', id: msg.id, ok: false, error: err.message }, '*');
        }
      );
      return;
    }

    if (msg.__efppt !== 'fetch') return;

    const source = e.source;
    const id = msg.id;
    const reply = function (payload) {
      if (source && source.postMessage) {
        source.postMessage(Object.assign({ __efppt: 'fetch-result', id }, payload), '*');
      }
    };

    // Cross-environment: this tab's origin can't reach another org's API
    // directly, so relay through a previously connected target tab instead.
    const urlOrigin = _originOf(msg.url);
    if (urlOrigin && urlOrigin !== window.location.origin) {
      const targetWin = _targetWindows.get(urlOrigin);
      if (!_targetReady.has(urlOrigin) || !targetWin || targetWin.closed) {
        reply({ ok: false, status: 0, error: 'Target environment (' + urlOrigin + ') is not connected. Use "Connect Target Environment" first.' });
        return;
      }
      _pendingRelays.set(id, source);
      setTimeout(function () {
        if (_pendingRelays.has(id)) {
          _pendingRelays.delete(id);
          reply({ ok: false, status: 0, error: 'Timed out waiting for the target tab to respond.' });
        }
      }, 30000);
      try {
        targetWin.postMessage(msg, '*'); // forwarded as-is; target's own bridge replies with the same id
      } catch (err) {
        _pendingRelays.delete(id);
        reply({ ok: false, status: 0, error: 'Could not reach the target tab: ' + err.message });
      }
      return;
    }

    const result = await _doFetch(msg.url, msg.method, msg.headers, msg.body);
    reply(result);
  });

  // ════════════════════════════════════════════════════════════════════
  //  TOOL MODAL  (large centered dialog, with a pop-out-to-tab button)
  // ════════════════════════════════════════════════════════════════════
  function closeTool() {
    const ex = document.getElementById('__ef-ppt-tool-overlay');
    if (ex) ex.remove();
  }

  function buildToolUrl(toolName, extraParams) {
    const env = currentEnv(CFG);

    const params = new URLSearchParams();
    params.set('env', window.location.origin);
    if (env) {
      params.set('name', env.name || '');
      if (env.powerAppsId) params.set('paEnvId', env.powerAppsId);
    }
    params.set('_inModal', '1');
    // Embed the full config as base64 so the tool page can read it even when
    // browser storage partitioning blocks it from seeing github.io's
    // localStorage while running as a third-party iframe on this D365 page.
    if (CFG) {
      try {
        params.set('cfg', btoa(unescape(encodeURIComponent(JSON.stringify(CFG)))));
      } catch (_) { /* leave cfg out — tool falls back to its own localStorage */ }
    }
    if (extraParams && typeof extraParams === 'object') {
      Object.keys(extraParams).forEach(function (k) {
        if (extraParams[k] != null) params.set(k, extraParams[k]);
      });
    }

    return BASE_URL + '/' + toolName + '/' + toolName + '.html?' + params.toString();
  }

  function openTool(toolName, extraParams) {
    closeTool();
    const src = buildToolUrl(toolName, extraParams);
    const label = (TOOLS.find(function (t) { return t.name === toolName; }) || {}).label || toolName;

    const overlay = document.createElement('div');
    overlay.id = '__ef-ppt-tool-overlay';
    overlay.setAttribute(
      'style',
      [
        'position:fixed',
        'inset:0',
        'z-index:2147483646',
        'background:rgba(15,23,42,0.55)',
        'display:flex',
        'align-items:center',
        'justify-content:center',
        'margin:0',
        'padding:0',
        'font-family:Segoe UI,system-ui,sans-serif'
      ].join(';')
    );

    const dialog = document.createElement('div');
    dialog.setAttribute(
      'style',
      [
        'width:92vw',
        'height:88vh',
        'max-width:1400px',
        'background:#fff',
        'border-radius:12px',
        'box-shadow:0 20px 60px rgba(0,0,0,.45)',
        'display:flex',
        'flex-direction:column',
        'overflow:hidden'
      ].join(';')
    );

    const titleBar = document.createElement('div');
    titleBar.setAttribute(
      'style',
      [
        'flex:0 0 auto',
        'display:flex',
        'align-items:center',
        'justify-content:space-between',
        'padding:10px 14px',
        'background:#1B3A6B',
        'color:#fff'
      ].join(';')
    );

    const titleLabel = document.createElement('span');
    titleLabel.textContent = 'EF PPT — ' + label;
    titleLabel.setAttribute('style', 'font-size:13px;font-weight:600;');

    const btnGroup = document.createElement('div');
    btnGroup.setAttribute('style', 'display:flex;align-items:center;gap:6px;');

    function makeIconBtn(text, title) {
      const b = document.createElement('button');
      b.type = 'button';
      b.title = title;
      b.setAttribute('aria-label', title);
      b.textContent = text;
      b.setAttribute(
        'style',
        [
          'width:28px',
          'height:28px',
          'border:none',
          'border-radius:6px',
          'background:rgba(255,255,255,.12)',
          'color:#fff',
          'font-size:13px',
          'line-height:1',
          'cursor:pointer'
        ].join(';')
      );
      b.addEventListener('mouseenter', function () { b.style.background = 'rgba(255,255,255,.25)'; });
      b.addEventListener('mouseleave', function () { b.style.background = 'rgba(255,255,255,.12)'; });
      return b;
    }

    const popoutBtn = makeIconBtn('⧉', 'Open in new tab');
    popoutBtn.addEventListener('click', function () {
      // No 'noopener' — the popped-out tab keeps window.opener so its D365
      // fetch bridge can still relay calls back through this launcher.
      window.open(src, '_blank');
      closeTool();
    });

    const closeBtn = makeIconBtn('✕', 'Close');
    closeBtn.addEventListener('click', closeTool);

    btnGroup.appendChild(popoutBtn);
    btnGroup.appendChild(closeBtn);
    titleBar.appendChild(titleLabel);
    titleBar.appendChild(btnGroup);

    const iframe = document.createElement('iframe');
    iframe.src = src;
    iframe.setAttribute(
      'style',
      [
        'flex:1 1 auto',
        'width:100%',
        'border:none',
        'background:#fff'
      ].join(';')
    );
    iframe.setAttribute('allow', 'clipboard-write');

    dialog.appendChild(titleBar);
    dialog.appendChild(iframe);
    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    overlay.addEventListener('click', function (ev) {
      if (ev.target === overlay) closeTool();
    });
  }

  // ════════════════════════════════════════════════════════════════════
  //  GO TO  logic
  // ════════════════════════════════════════════════════════════════════
  function parseRecordContext() {
    // Read entity + record id from the current D365 URL.
    const p = new URLSearchParams(window.location.search);
    return {
      etn: p.get('etn') || '',
      id: (p.get('id') || '').replace(/[{}]/g, '')
    };
  }

  // ── Configured apps (for the "App" Go-To type) ─────────────────────────
  // settings.includedApps is a flat list of app uniquenames (global, not
  // per-environment — matches the original extension's model). Friendly
  // display names are best-effort: fetched once from the CURRENT environment
  // (same-origin, no bridge needed) and used purely for labelling the
  // dropdown. Navigation itself only ever needs the uniquename, via D365's
  // `/apps/uniquename/<name>` launch URL — the same reliable pattern the
  // extension uses, which needs no cross-environment app-id resolution at all.
  let _appNameCache = null; // Map<uniquename, displayName> | null while loading

  function configuredAppUniqueNames() {
    const settings = (CFG && CFG.settings) || {};
    const list = Array.isArray(settings.includedApps) ? settings.includedApps.slice() : [];
    if (settings.defaultAppUniqueName && list.indexOf(settings.defaultAppUniqueName) === -1) {
      list.unshift(settings.defaultAppUniqueName);
    }
    return list;
  }

  async function ensureAppNames() {
    if (_appNameCache) return _appNameCache;
    _appNameCache = new Map();
    try {
      const ver = apiVersion(CFG);
      const res = await fetch(
        window.location.origin + '/api/data/' + ver + '/appmodules?$select=uniquename,name',
        {
          credentials: 'include',
          headers: { Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }
        }
      );
      if (res.ok) {
        const data = await res.json();
        (data.value || []).forEach(function (a) {
          if (a.uniquename) _appNameCache.set(a.uniquename, a.name || a.uniquename);
        });
      }
    } catch (err) {
      console.error('[EF PPT] Failed to fetch app names:', err);
    }
    return _appNameCache;
  }

  function appLabel(uniquename) {
    return (_appNameCache && _appNameCache.get(uniquename)) || uniquename;
  }

  async function doGoTo() {
    const target =
      (CFG && CFG.environments || []).find(function (e) { return e.id === selectedTargetId; });
    if (!target) return;

    const ver = apiVersion(CFG);
    let targetUrl = '';

    switch (selectedType) {
      case 'open-in': {
        // Deliberately does NOT pass appuniquename: the app configured in the
        // CURRENT environment has no guaranteed match in the TARGET one (different
        // org, different solution deployment), and forcing a mismatched app causes
        // D365 to throw an unhandled-exception error page. Omitting it lets the
        // target org fall back to its own default/last-used app for the record.
        const ctx = parseRecordContext();
        const qp = new URLSearchParams();
        qp.set('pagetype', 'entityrecord');
        if (ctx.etn) qp.set('etn', ctx.etn);
        if (ctx.id) qp.set('id', ctx.id);
        targetUrl = target.url + '/main.aspx?' + qp.toString();
        break;
      }
      case 'api':
        targetUrl = target.url + '/api/data/' + ver + '/';
        break;
      case 'solutions':
        targetUrl = target.powerAppsId
          ? 'https://make.powerapps.com/environments/' + target.powerAppsId + '/solutions'
          : target.url + '/tools/solution';
        break;
      case 'app': {
        const apps = configuredAppUniqueNames();
        const uname = selectedAppUniqueName || apps[0];
        if (!uname) return;
        targetUrl = target.url + '/apps/uniquename/' + encodeURIComponent(uname);
        break;
      }
      case 'security':
        targetUrl = target.url + '/main.aspx?pagetype=entitylist&etn=role';
        break;
      default:
        targetUrl = target.url;
    }

    window.open(targetUrl, '_blank', 'noopener');
  }

  // ════════════════════════════════════════════════════════════════════
  //  RECORD DETAILS PANE  (left column of the flyout when on a record page)
  // ════════════════════════════════════════════════════════════════════
  // Unlike the other tools, this is rendered directly inside the flyout —
  // no modal, no iframe. Reading the current record is always same-origin
  // (it's the D365 page the launcher is running on), so it needs no bridge
  // at all. Only Clone-to-a-different-environment needs the cross-tab relay,
  // which reuses the exact same _connectTarget/_selfFetch machinery Data
  // Sync drives externally via postMessage.

  const _entitySetNameCache = new Map();

  async function fetchEntitySetName(baseUrl, etn) {
    const cacheKey = baseUrl + ':' + etn;
    if (_entitySetNameCache.has(cacheKey)) return _entitySetNameCache.get(cacheKey);
    try {
      const ver = apiVersion(CFG);
      const d = await _selfFetch(baseUrl + '/api/data/' + ver + "/EntityDefinitions(LogicalName='" + etn + "')?$select=EntitySetName");
      const name = d && d.EntitySetName ? d.EntitySetName : null;
      if (name) _entitySetNameCache.set(cacheKey, name);
      return name;
    } catch (err) {
      console.error('[EF PPT] Failed to resolve entity set name:', err);
      return null;
    }
  }

  function ensureRecordDetailsLoaded(ctx) {
    if (!ctx || !ctx.etn || !ctx.id) { _rdState = null; return; }
    if (_rdState && _rdState.etn === ctx.etn && _rdState.id === ctx.id) return; // already loading/loaded
    _rdState = { etn: ctx.etn, id: ctx.id, status: 'loading', data: null, entitySetName: null, error: null };
    _rdCloneTargetId = null;
    _rdConnecting = false;
    _rdConnectError = null;
    _rdCloning = false;
    _rdCloneResult = null;
    loadRecordDetails(ctx.etn, ctx.id);
  }

  async function loadRecordDetails(etn, id) {
    try {
      const entitySetName = await fetchEntitySetName(window.location.origin, etn);
      if (!entitySetName) throw new Error('Could not resolve entity type.');
      const ver = apiVersion(CFG);
      const url =
        window.location.origin + '/api/data/' + ver + '/' + entitySetName + '(' + id + ')' +
        '?$select=createdon,modifiedon,_ownerid_value' +
        '&$expand=createdby($select=fullname,systemuserid),modifiedby($select=fullname,systemuserid)';
      const data = await _selfFetch(url, {
        'Prefer': 'odata.include-annotations="OData.Community.Display.V1.FormattedValue,Microsoft.Dynamics.CRM.lookuplogicalname"',
      });
      if (_rdState && _rdState.etn === etn && _rdState.id === id) {
        _rdState.status = 'loaded';
        _rdState.data = data;
        _rdState.entitySetName = entitySetName;
        renderFlyout();
      }
    } catch (err) {
      if (_rdState && _rdState.etn === etn && _rdState.id === id) {
        _rdState.status = 'error';
        _rdState.error = err.message;
        renderFlyout();
      }
    }
  }

  function _rdFormatDateTime(iso) {
    if (!iso) return '—';
    try {
      const d = new Date(iso);
      const dd = String(d.getDate()).padStart(2, '0');
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      const hh = String(d.getHours()).padStart(2, '0');
      const min = String(d.getMinutes()).padStart(2, '0');
      return dd + '/' + mm + '/' + d.getFullYear() + ' ' + hh + ':' + min;
    } catch (e) { return iso; }
  }

  function _rdRecordUrl(envUrl, etn, id) {
    if (!envUrl || !id) return null;
    return envUrl + '/main.aspx?pagetype=entityrecord&etn=' + etn + '&id=' + id;
  }

  /** Builds one card's HTML. rows: [{ text, mono, link, icon }] */
  function _rdCard(label, rows) {
    const rowsHtml = rows.filter(function (r) { return r.text != null; }).map(function (r) {
      const cls = r.mono ? 'rd-mono' : 'rd-value';
      const inner = r.link
        ? '<a class="' + cls + ' rd-link" href="' + escHtml(r.link) + '" target="_blank" rel="noopener noreferrer">' + escHtml(r.text) + '</a>'
        : '<span class="' + cls + '">' + escHtml(r.text) + '</span>';
      return (
        '<div class="rd-value-row">' +
          (r.icon ? '<span class="rd-icon">' + r.icon + '</span>' : '') +
          inner +
          '<button type="button" class="rd-copy-btn" data-copy="' + escHtml(r.text) + '" title="Copy">' + COPY_ICON_SVG + '</button>' +
        '</div>'
      );
    }).join('');
    return '<div class="rd-card"><span class="rd-label">' + escHtml(label) + '</span>' + rowsHtml + '</div>';
  }

  function buildRecordPaneMarkup() {
    if (!_rdState) return '';

    if (_rdState.status === 'loading') {
      return '<div class="rd-state"><span class="rd-mini-spinner"></span> Loading record details…</div>';
    }
    if (_rdState.status === 'error') {
      return '<div class="rd-state rd-state--error">Failed to load: ' + escHtml(_rdState.error) + '</div>';
    }

    const env = currentEnv(CFG);
    const cSettings = cloneSettings(CFG);
    const data = _rdState.data || {};
    const apiLink = env ? (env.url + '/api/data/' + cSettings.apiVersion + '/' + _rdState.entitySetName + '(' + _rdState.id + ')') : null;

    let html = '<div class="rd-cards">';
    html += _rdCard('Record ID', [{ text: _rdState.id, mono: true, link: apiLink }]);
    html += _rdCard('Entity Type', [{ text: _rdState.etn }]);
    html += _rdCard('Created On', [{ text: _rdFormatDateTime(data.createdon) }]);
    html += _rdCard('Created By', [
      { text: (data.createdby && data.createdby.fullname) || '—', link: env && data.createdby ? _rdRecordUrl(env.url, 'systemuser', data.createdby.systemuserid) : null },
      { text: data.createdby ? data.createdby.systemuserid : null, mono: true },
    ]);
    html += _rdCard('Modified On', [{ text: _rdFormatDateTime(data.modifiedon) }]);
    html += _rdCard('Modified By', [
      { text: (data.modifiedby && data.modifiedby.fullname) || '—', link: env && data.modifiedby ? _rdRecordUrl(env.url, 'systemuser', data.modifiedby.systemuserid) : null },
      { text: data.modifiedby ? data.modifiedby.systemuserid : null, mono: true },
    ]);

    const ownerName = data['_ownerid_value@OData.Community.Display.V1.FormattedValue'] || null;
    const ownerId   = data['_ownerid_value'] || null;
    const ownerType = data['_ownerid_value@Microsoft.Dynamics.CRM.lookuplogicalname'] || 'systemuser';
    if (ownerName || ownerId) {
      const isTeam = ownerType === 'team';
      html += _rdCard('Owner', [
        { text: ownerName || '—', icon: isTeam ? TEAM_ICON_SVG : USER_ICON_SVG, link: env ? _rdRecordUrl(env.url, isTeam ? 'team' : 'systemuser', ownerId) : null },
        { text: ownerId, mono: true },
      ]);
    }
    html += '</div>';

    // ── Clone ──────────────────────────────────────────────────────────────
    const cloneAllowed = env && (!cSettings.cloneWhitelist || cSettings.cloneWhitelist.indexOf((_rdState.etn || '').toLowerCase()) !== -1);
    if (cloneAllowed) {
      const envs = (CFG && CFG.environments) || [];
      const cloneEnvs = [env].concat(envs.filter(function (e) { return e.url !== env.url; }));
      const targetId = _rdCloneTargetId || env.id;
      const target = cloneEnvs.filter(function (e) { return e.id === targetId; })[0] || env;
      const targetOrigin = _originOf(target.url);
      const isCross = targetOrigin !== _originOf(env.url);
      const needsConnect = isCross && !_targetReady.has(targetOrigin);

      const options = cloneEnvs.map(function (e) {
        return '<option value="' + escHtml(e.id) + '"' + (e.id === target.id ? ' selected' : '') + '>' + escHtml(e.name) + '</option>';
      }).join('');

      html += '<div class="rd-clone-bar">';
      html += '<div class="rd-clone-row">';
      html += (
        '<button type="button" class="rd-clone-btn" id="ef-rd-clone-btn"' + ((_rdCloning || needsConnect) ? ' disabled' : '') + '>' +
          (_rdCloning
            ? '<span class="rd-mini-spinner"></span> ' + (isCross ? 'Copying…' : 'Cloning…')
            : CLONE_ICON_SVG + ' Clone') +
        '</button>'
      );
      html += '<select class="rd-clone-select" id="ef-rd-clone-target">' + options + '</select>';
      html += '</div>';

      if (needsConnect) {
        html += (
          '<div class="rd-alert ' + (_rdConnectError ? 'rd-alert--error' : 'rd-alert--warning') + '">' +
            '<div>' + (_rdConnectError
              ? escHtml(_rdConnectError)
              : ('Cloning to ' + escHtml(target.name) + ' needs a live connection. Click Connect, then click the EF PPT bookmark in the new tab that opens.')) +
            '</div>' +
            '<button type="button" class="rd-btn-secondary" id="ef-rd-connect-btn"' + (_rdConnecting ? ' disabled' : '') + '>' +
              (_rdConnecting ? 'Connecting…' : 'Connect Target Environment') +
            '</button>' +
          '</div>'
        );
      }

      if (_rdCloneResult) {
        if (_rdCloneResult.ok) {
          const newUrl = _rdCloneResult.targetEnvUrl + '/main.aspx?pagetype=entityrecord&etn=' + _rdState.etn + '&id=' + _rdCloneResult.newId;
          html += (
            '<div class="rd-clone-success">' + CHECK_ICON_SVG + (_rdCloneResult.isCross ? ' Copied!' : ' Cloned!') +
            ' <a href="' + escHtml(newUrl) + '" target="_blank" rel="noopener noreferrer">Open record ↗</a></div>'
          );
        } else {
          html += (
            '<div class="rd-clone-error">' +
              '<span style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;">Error occurred.</span>' +
              '<button type="button" class="rd-link-btn" id="ef-rd-clone-err-details">View details</button>' +
            '</div>'
          );
        }
      }

      html += '</div>';
    }

    return html;
  }

  function onRdCloneTargetChange(selectEl) {
    _rdCloneTargetId = selectEl.value;
    _rdConnectError = null;
    _rdCloneResult = null;
    renderFlyout();
  }

  function onRdConnectClick() {
    const envs = (CFG && CFG.environments) || [];
    const target = envs.filter(function (e) { return e.id === _rdCloneTargetId; })[0];
    if (!target) return;
    const origin = _originOf(target.url);
    _rdConnecting = true;
    _rdConnectError = null;
    renderFlyout();
    _connectTarget(origin).then(
      function () {
        _rdConnecting = false;
        renderFlyout();
      },
      function (err) {
        _rdConnecting = false;
        _rdConnectError = (err && err.message) || 'Could not connect to the target environment.';
        renderFlyout();
      }
    );
  }

  function onRdCloneClick() {
    if (!_rdState || _rdState.status !== 'loaded') return;
    const env = currentEnv(CFG);
    if (!env) return;
    const envs = (CFG && CFG.environments) || [];
    const targetId = _rdCloneTargetId || env.id;
    const target = envs.filter(function (e) { return e.id === targetId; })[0] || env;
    const isCross = _originOf(target.url) !== _originOf(env.url);

    _rdCloning = true;
    _rdCloneResult = null;
    renderFlyout();

    cloneRecord(env.url, target.url, _rdState.etn, _rdState.id, _rdState.entitySetName).then(
      function (result) {
        _rdCloning = false;
        _rdCloneResult = { ok: true, newId: result.newId, isCross: isCross, targetEnvUrl: target.url };
        renderFlyout();
      },
      function (err) {
        _rdCloning = false;
        _rdCloneResult = { ok: false, error: err.message };
        renderFlyout();
      }
    );
  }

  function openRdErrorModal(text) {
    closeRdErrorModal();
    const overlay = document.createElement('div');
    overlay.id = '__ef-ppt-rd-err-overlay';
    overlay.setAttribute('style', 'position:fixed;inset:0;z-index:2147483647;background:rgba(15,23,42,.5);display:flex;align-items:center;justify-content:center;font-family:Segoe UI,system-ui,sans-serif;margin:0;padding:16px;');
    overlay.innerHTML =
      '<div style="background:#fff;border-radius:10px;max-width:480px;width:100%;max-height:70vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 10px 40px rgba(0,0,0,.4);">' +
        '<div style="font-size:13px;font-weight:700;color:#dc2626;padding:12px 14px 8px;border-bottom:1px solid #fecaca;">Clone Error Details</div>' +
        '<pre style="flex:1;overflow-y:auto;padding:12px 14px;font-size:11px;font-family:Consolas,monospace;color:#1e293b;white-space:pre-wrap;word-break:break-word;background:#fef2f2;margin:0;">' + escHtml(text || '') + '</pre>' +
        '<button id="__ef-ppt-rd-err-close" type="button" style="padding:9px 14px;background:#1B3A6B;color:#fff;border:none;font-size:12px;font-weight:600;cursor:pointer;">Close</button>' +
      '</div>';
    document.body.appendChild(overlay);
    overlay.addEventListener('click', function (ev) { if (ev.target === overlay) closeRdErrorModal(); });
    overlay.querySelector('#__ef-ppt-rd-err-close').addEventListener('click', closeRdErrorModal);
  }

  function closeRdErrorModal() {
    const ex = document.getElementById('__ef-ppt-rd-err-overlay');
    if (ex) ex.remove();
  }

  // ── Clone engine ─────────────────────────────────────────────────────────
  // Same-env: POST a new record, then re-create every M2M association.
  // Cross-env: PATCH (upsert) the SAME record ID on the target — self-healing
  // against fields the target entity schema doesn't have — then reconcile M2M
  // associations exactly. Every request goes through _selfFetch/_selfFetchRaw,
  // which transparently picks direct-fetch vs. connected-tab-relay per URL.

  const CLONE_EXCLUDE = new Set([
    'createdon', 'modifiedon', 'overriddencreatedon',
    'createdby', 'modifiedby', 'createdonbehalfby', 'modifiedonbehalfby',
    'versionnumber', 'exchangerate', 'importsequencenumber',
    'timezoneruleversionnumber', 'utcconversiontimezonecode',
    'owningbusinessunit', 'owninguser', 'owningteam',
  ]);
  const CLONE_SKIP_TYPES = new Set(['Virtual', 'EntityName', 'ManagedProperty', 'Uniqueidentifier']);

  async function cloneRecord(sourceUrl, targetUrl, etn, id, sourceEntitySetName) {
    const isCross      = _originOf(targetUrl) !== _originOf(sourceUrl);
    const apiBase       = sourceUrl + '/api/data/' + cloneSettings(CFG).apiVersion;
    const targetApiBase = targetUrl + '/api/data/' + cloneSettings(CFG).apiVersion;

    // ── 1. Entity definition — primary name/id attrs, writable attrs, M2M ──
    const results = await Promise.all([
      _selfFetch(
        apiBase + "/EntityDefinitions(LogicalName='" + etn + "')?$select=PrimaryNameAttribute,PrimaryIdAttribute" +
        "&$expand=ManyToManyRelationships($select=SchemaName,Entity1LogicalName,Entity2LogicalName,Entity1NavigationPropertyName,Entity2NavigationPropertyName)"
      ),
      _selfFetch(apiBase + "/EntityDefinitions(LogicalName='" + etn + "')/Attributes?$select=LogicalName,AttributeType,IsValidForCreate"),
      _selfFetch(apiBase + "/EntityDefinitions(LogicalName='" + etn + "')/ManyToOneRelationships?$select=ReferencingAttribute,ReferencingEntityNavigationPropertyName"),
    ]);
    const def = results[0], attrData = results[1], m2oData = results[2];

    const primaryName   = def.PrimaryNameAttribute;
    const primaryIdAttr = def.PrimaryIdAttribute || (etn + 'id');
    const m2m           = def.ManyToManyRelationships || [];
    const attrs         = attrData.value || [];
    const navPropMap    = {};
    (m2oData.value || []).forEach(function (r) { navPropMap[r.ReferencingAttribute] = r.ReferencingEntityNavigationPropertyName; });

    let primaryMaxLen = 100;
    if (primaryName) {
      try {
        const lenD = await _selfFetch(
          apiBase + "/EntityDefinitions(LogicalName='" + etn + "')/Attributes/Microsoft.Dynamics.CRM.StringAttributeMetadata" +
          "?$select=LogicalName,MaxLength&$filter=LogicalName eq '" + primaryName + "'"
        );
        primaryMaxLen = (lenD.value && lenD.value[0] && lenD.value[0].MaxLength) || 100;
      } catch (_) { /* keep default */ }
    }

    // ── 2. Full record (all fields + lookup annotations) ───────────────────
    const record = await _selfFetch(apiBase + '/' + sourceEntitySetName + '(' + id + ')', {
      'Prefer': 'odata.include-annotations="OData.Community.Display.V1.FormattedValue,Microsoft.Dynamics.CRM.lookuplogicalname"',
    });

    // ── 3. Build payload ─────────────────────────────────────────────────────
    const EXCLUDE = new Set([primaryIdAttr]);
    CLONE_EXCLUDE.forEach(function (x) { EXCLUDE.add(x); });
    const scalarPayload = {};
    const lookupFields  = []; // { logicalName, guid, targetEtn }

    attrs.forEach(function (attr) {
      if (!attr.IsValidForCreate) return;
      const name = attr.LogicalName;
      if (EXCLUDE.has(name)) return;
      if (CLONE_SKIP_TYPES.has(attr.AttributeType)) return;

      if (attr.AttributeType === 'Lookup' || attr.AttributeType === 'Customer' || attr.AttributeType === 'Owner') {
        const guidKey = '_' + name + '_value';
        const guid    = record[guidKey];
        if (!guid) return;
        const targetEtn = record[guidKey + '@Microsoft.Dynamics.CRM.lookuplogicalname'];
        if (targetEtn) lookupFields.push({ logicalName: name, guid: guid, targetEtn: targetEtn });
      } else {
        const val = record[name];
        if (val === null || val === undefined) return;
        if (name === primaryName && !isCross) {
          const rawPrefix = (cloneSettings(CFG).clonePrefix || '').trim();
          const prefix    = rawPrefix ? ('[' + rawPrefix + '] ') : '';
          const full      = prefix + String(val);
          scalarPayload[name] = full.length > primaryMaxLen ? full.slice(0, primaryMaxLen) : full;
        } else {
          scalarPayload[name] = val;
        }
      }
    });

    return isCross
      ? _cloneCrossEnv({ etn: etn, id: id, apiBase: apiBase, targetApiBase: targetApiBase, scalarPayload: scalarPayload, lookupFields: lookupFields, navPropMap: navPropMap, m2m: m2m, sourceUrl: sourceUrl, targetUrl: targetUrl })
      : _cloneSameEnv({ apiBase: apiBase, sourceEntitySetName: sourceEntitySetName, primaryIdAttr: primaryIdAttr, scalarPayload: scalarPayload, lookupFields: lookupFields, navPropMap: navPropMap, m2m: m2m, sourceUrl: sourceUrl, id: id });
  }

  async function _cloneSameEnv(ctx) {
    const apiBase = ctx.apiBase, sourceEntitySetName = ctx.sourceEntitySetName, primaryIdAttr = ctx.primaryIdAttr;
    const scalarPayload = ctx.scalarPayload, lookupFields = ctx.lookupFields, navPropMap = ctx.navPropMap;
    const m2m = ctx.m2m, sourceUrl = ctx.sourceUrl, id = ctx.id;

    // Resolve entity set names for all lookup targets (parallel, cached)
    const uniqueTargetEtns = Array.from(new Set(lookupFields.map(function (f) { return f.targetEtn; })));
    const etnToSet = {};
    await Promise.all(uniqueTargetEtns.map(async function (t) { etnToSet[t] = await fetchEntitySetName(sourceUrl, t); }));

    const fullPayload = Object.assign({}, scalarPayload);
    lookupFields.forEach(function (f) {
      const targetSet = etnToSet[f.targetEtn];
      if (!targetSet) return;
      const navProp = navPropMap[f.logicalName] || f.logicalName;
      fullPayload[navProp + '@odata.bind'] = '/' + targetSet + '(' + f.guid + ')';
    });

    // Prefer: return=representation — the bridge doesn't forward response headers
    // (e.g. OData-EntityId), so the new ID must come back in the JSON body instead.
    const createHeaders = { 'Content-Type': 'application/json', 'MSCRM.SuppressDuplicateDetection': 'true', 'Prefer': 'return=representation' };
    const createResult  = await _selfFetchRaw(apiBase + '/' + sourceEntitySetName, createHeaders, 'POST', fullPayload);
    if (!createResult.ok) {
      const errText = createResult.error || (createResult.data && createResult.data.error && createResult.data.error.message) || ('HTTP ' + createResult.status);
      throw new Error('Create failed (' + createResult.status + '): ' + errText);
    }
    const newId = createResult.data && createResult.data[primaryIdAttr];
    if (!newId) throw new Error('Clone created but new record ID could not be determined.');

    // M2M: add all source associations to the new clone
    for (let i = 0; i < m2m.length; i++) {
      await _syncM2mAdd({ rel: m2m[i], apiBase: apiBase, entitySetName: sourceEntitySetName, sourceUrl: sourceUrl, id: id, newId: newId });
    }

    return { newId: newId, entitySetName: sourceEntitySetName };
  }

  /** Adds every source-side association of `rel` to the newly cloned record (same env only). */
  async function _syncM2mAdd(ctx) {
    const rel = ctx.rel, apiBase = ctx.apiBase, entitySetName = ctx.entitySetName;
    const sourceUrl = ctx.sourceUrl, id = ctx.id, newId = ctx.newId;

    const candidates = [
      { navProp: rel.Entity1NavigationPropertyName, relatedEtn: rel.Entity2LogicalName },
      { navProp: rel.Entity2NavigationPropertyName, relatedEtn: rel.Entity1LogicalName },
    ];

    for (let i = 0; i < candidates.length; i++) {
      const navProp = candidates[i].navProp, relatedEtn = candidates[i].relatedEtn;
      if (!navProp || !relatedEtn) continue;
      const relatedEntitySet = await fetchEntitySetName(sourceUrl, relatedEtn);
      if (!relatedEntitySet) continue;
      const pkField = relatedEtn + 'id';

      let relatedIds;
      try {
        const d = await _selfFetch(apiBase + '/' + entitySetName + '(' + id + ')/' + navProp + '?$top=500&$select=' + pkField);
        relatedIds = (d.value || []).map(function (r) { return r[pkField]; }).filter(Boolean);
      } catch (e) {
        continue; // this nav prop doesn't apply to our side of the relationship — try the other one
      }

      for (let j = 0; j < relatedIds.length; j++) {
        try {
          await _selfFetch(apiBase + '/' + entitySetName + '(' + newId + ')/' + navProp + '/$ref', { 'Content-Type': 'application/json' }, 'POST', {
            '@odata.id': apiBase + '/' + relatedEntitySet + '(' + relatedIds[j] + ')',
          });
        } catch (e) {
          console.error('[EF PPT] M2M associate failed:', e);
        }
      }
      return; // matched side handled — don't also try the other nav prop
    }
  }

  async function _cloneCrossEnv(ctx) {
    const etn = ctx.etn, id = ctx.id, apiBase = ctx.apiBase, targetApiBase = ctx.targetApiBase;
    const scalarPayload = ctx.scalarPayload, lookupFields = ctx.lookupFields, navPropMap = ctx.navPropMap;
    const m2m = ctx.m2m, sourceUrl = ctx.sourceUrl, targetUrl = ctx.targetUrl;

    const sourceEntitySetName = await fetchEntitySetName(sourceUrl, etn);
    const targetEntitySetName = await fetchEntitySetName(targetUrl, etn);
    if (!targetEntitySetName) throw new Error("Cannot resolve entity set for '" + etn + "' in target environment.");

    // Target schema: writable attrs + nav prop map
    let tAttrSet = null, targetNavMap = {};
    try {
      const results = await Promise.all([
        _selfFetch(targetApiBase + "/EntityDefinitions(LogicalName='" + etn + "')/Attributes?$select=LogicalName,IsValidForCreate"),
        _selfFetch(targetApiBase + "/EntityDefinitions(LogicalName='" + etn + "')/ManyToOneRelationships?$select=ReferencingAttribute,ReferencingEntityNavigationPropertyName"),
      ]);
      tAttrSet = new Set((results[0].value || []).filter(function (a) { return a.IsValidForCreate; }).map(function (a) { return a.LogicalName; }));
      (results[1].value || []).forEach(function (r) { targetNavMap[r.ReferencingAttribute] = r.ReferencingEntityNavigationPropertyName; });
    } catch (e) {
      console.error('[EF PPT] Failed to fetch target schema:', e);
    }

    // Validate scalar fields against target schema
    const validatedScalar = {};
    Object.keys(scalarPayload).forEach(function (k) {
      if (!tAttrSet || tAttrSet.has(k)) validatedScalar[k] = scalarPayload[k];
    });

    // Rebuild lookup bindings using the target's own entity-set names + nav props
    const validatedLookups = {};
    for (let i = 0; i < lookupFields.length; i++) {
      const f = lookupFields[i];
      if (tAttrSet && !tAttrSet.has(f.logicalName)) continue;
      const targetLookupSet = await fetchEntitySetName(targetUrl, f.targetEtn);
      if (!targetLookupSet) continue;
      const navProp = targetNavMap[f.logicalName] || navPropMap[f.logicalName] || f.logicalName;
      validatedLookups[navProp + '@odata.bind'] = '/' + targetLookupSet + '(' + f.guid + ')';
    }

    // ── Upsert scalar fields — self-healing retry against fields the target
    // entity schema doesn't have (source-only custom fields, etc). ─────────
    const writeHeaders = { 'Content-Type': 'application/json', 'MSCRM.SuppressDuplicateDetection': 'true' };
    let remaining = Object.assign({}, validatedScalar);
    let lastError = null;
    for (let attempt = 0; attempt < 30 && Object.keys(remaining).length > 0; attempt++) {
      const r = await _selfFetchRaw(targetApiBase + '/' + targetEntitySetName + '(' + id + ')', writeHeaders, 'PATCH', remaining);
      if (r.ok) { remaining = {}; break; }
      const errText = r.error || (r.data && r.data.error && r.data.error.message) || ('HTTP ' + r.status);
      lastError = 'Upsert failed (' + r.status + '): ' + errText;

      const m = errText.match(/[`'"]?([a-z_][a-z0-9_]*)[`'"]?\s+(?:field\s+)?missing\s+from\s+target\s+entity/i)
             || errText.match(/missing\s+from\s+target\s+entity[^:]*:\s*([a-z_][a-z0-9_]*)/i);
      const badField = m && m[1];
      if (badField && badField in remaining) { delete remaining[badField]; continue; }
      throw new Error(lastError);
    }
    if (Object.keys(remaining).length > 0 && lastError) throw new Error(lastError);

    // ── PATCH each lookup individually (skip or fail per cloneLookupMode) ──
    const failOnLookup = cloneSettings(CFG).cloneLookupMode === 'fail';
    const lookupEntries = Object.keys(validatedLookups).map(function (k) { return [k, validatedLookups[k]]; });
    for (let i = 0; i < lookupEntries.length; i++) {
      const bindKey = lookupEntries[i][0], bindVal = lookupEntries[i][1];
      const body = {}; body[bindKey] = bindVal;
      const r = await _selfFetchRaw(targetApiBase + '/' + targetEntitySetName + '(' + id + ')', writeHeaders, 'PATCH', body);
      if (!r.ok && failOnLookup) {
        throw new Error('Lookup PATCH failed (' + r.status + '): ' + (r.error || (r.data && r.data.error && r.data.error.message) || ('HTTP ' + r.status)));
      }
    }

    // ── Exact M2M sync in target (add missing, remove extra) ───────────────
    for (let i = 0; i < m2m.length; i++) {
      await _syncM2mExact({ rel: m2m[i], apiBase: apiBase, targetApiBase: targetApiBase, sourceEntitySetName: sourceEntitySetName, targetEntitySetName: targetEntitySetName, sourceUrl: sourceUrl, targetUrl: targetUrl, id: id });
    }

    return { newId: id, entitySetName: targetEntitySetName };
  }

  async function _syncM2mExact(ctx) {
    const rel = ctx.rel, apiBase = ctx.apiBase, targetApiBase = ctx.targetApiBase;
    const sourceEntitySetName = ctx.sourceEntitySetName, targetEntitySetName = ctx.targetEntitySetName;
    const sourceUrl = ctx.sourceUrl, targetUrl = ctx.targetUrl, id = ctx.id;

    const candidates = [
      { navProp: rel.Entity1NavigationPropertyName, relatedEtn: rel.Entity2LogicalName },
      { navProp: rel.Entity2NavigationPropertyName, relatedEtn: rel.Entity1LogicalName },
    ];

    for (let i = 0; i < candidates.length; i++) {
      const navProp = candidates[i].navProp, relatedEtn = candidates[i].relatedEtn;
      if (!navProp || !relatedEtn) continue;
      const pkField = relatedEtn + 'id';

      let sourceIds;
      try {
        const d = await _selfFetch(apiBase + '/' + sourceEntitySetName + '(' + id + ')/' + navProp + '?$top=500&$select=' + pkField);
        sourceIds = (d.value || []).map(function (r) { return r[pkField]; }).filter(Boolean);
      } catch (e) {
        continue; // wrong side of the relationship for our entity — try the other nav prop
      }

      const relatedSetSource = await fetchEntitySetName(sourceUrl, relatedEtn);
      const relatedSetTarget = await fetchEntitySetName(targetUrl, relatedEtn);
      if (!relatedSetSource || !relatedSetTarget) return;

      let targetIds = [];
      try {
        const d = await _selfFetch(targetApiBase + '/' + targetEntitySetName + '(' + id + ')/' + navProp + '?$top=500&$select=' + pkField);
        targetIds = (d.value || []).map(function (r) { return r[pkField]; }).filter(Boolean);
      } catch (e) {
        console.error('[EF PPT] Failed to fetch target M2M IDs:', e);
      }

      const srcSet = new Set(sourceIds);
      const tgtSet = new Set(targetIds);

      for (let s = 0; s < sourceIds.length; s++) {
        if (tgtSet.has(sourceIds[s])) continue;
        try {
          await _selfFetch(targetApiBase + '/' + targetEntitySetName + '(' + id + ')/' + navProp + '/$ref', { 'Content-Type': 'application/json' }, 'POST', {
            '@odata.id': targetApiBase + '/' + relatedSetTarget + '(' + sourceIds[s] + ')',
          });
        } catch (e) { console.error('[EF PPT] M2M associate failed:', e); }
      }
      for (let t = 0; t < targetIds.length; t++) {
        if (srcSet.has(targetIds[t])) continue;
        try {
          await _selfFetch(targetApiBase + '/' + targetEntitySetName + '(' + id + ')/' + navProp + '(' + targetIds[t] + ')/$ref', {}, 'DELETE');
        } catch (e) { console.error('[EF PPT] M2M disassociate failed:', e); }
      }
      return; // matched side handled
    }
  }

  // ════════════════════════════════════════════════════════════════════
  //  CONFIG VIEWER OVERLAY  (appended to body, not shadow DOM)
  // ════════════════════════════════════════════════════════════════════
  // Read-only by design. Config lives in ONE place — the github.io web app's
  // localStorage — and reaches every D365 environment by being embedded in
  // the bookmarklet URL when you drag it. Editing here would write only to
  // this D365 origin's localStorage, silently diverging from every other
  // environment and from github.io itself. Always edit on github.io, then
  // re-drag the bookmark so the new config reaches every environment.
  function closeConfigEditor() {
    const ex = document.getElementById('__ef-ppt-config-overlay');
    if (ex) ex.remove();
  }

  function openConfigEditor() {
    closeConfigEditor();
    const current = CFG || STARTER_CONFIG;
    const json = JSON.stringify(current, null, 2);
    const isPlaceholder = !CFG;
    const configPageUrl = BASE_URL + '/index.html';

    const overlay = document.createElement('div');
    overlay.id = '__ef-ppt-config-overlay';
    overlay.setAttribute(
      'style',
      [
        'position:fixed',
        'inset:0',
        'z-index:2147483647',
        'background:rgba(15,23,42,0.6)',
        'display:flex',
        'align-items:center',
        'justify-content:center',
        'font-family:Segoe UI,system-ui,sans-serif',
        'margin:0',
        'padding:16px'
      ].join(';')
    );

    overlay.innerHTML =
      '<div style="background:#fff;border-radius:12px;max-width:640px;width:100%;max-height:90vh;overflow:auto;box-shadow:0 10px 40px rgba(0,0,0,.4);padding:20px 22px;box-sizing:border-box;">' +
        '<h2 style="margin:0 0 4px;font-size:18px;color:#1B3A6B;">EF PPT — Configuration</h2>' +
        (isPlaceholder
          ? '<p style="margin:0 0 12px;font-size:12px;color:#b91c1c;">No config found. Set up your environments on the web app, then drag the bookmark into your bookmark bar.</p>'
          : '<p style="margin:0 0 12px;font-size:12px;color:#475569;">Read-only — this is the config currently active on this environment. ' +
            'To make changes, edit it on the web app and drag a fresh bookmark so every environment picks up the update.</p>') +
        '<textarea id="__ef-ppt-cfg-text" readonly spellcheck="false" style="width:100%;height:280px;box-sizing:border-box;font-family:Consolas,monospace;font-size:12px;border:1px solid #cbd5e1;border-radius:8px;padding:10px;resize:vertical;background:#f8fafc;color:#334155;">' +
          escHtml(json) +
        '</textarea>' +
        '<div style="display:flex;gap:8px;justify-content:flex-end;margin-top:14px;">' +
          '<button id="__ef-ppt-cfg-close" type="button" style="padding:8px 16px;border:1px solid #cbd5e1;background:#fff;border-radius:8px;cursor:pointer;font-size:13px;">Close</button>' +
          '<a id="__ef-ppt-cfg-open" href="' + escHtml(configPageUrl) + '" target="_blank" rel="noopener" ' +
            'style="padding:8px 16px;border:none;background:#1B3A6B;color:#fff;border-radius:8px;cursor:pointer;font-size:13px;text-decoration:none;display:inline-flex;align-items:center;">' +
            'Edit on Web App ↗' +
          '</a>' +
        '</div>' +
      '</div>';

    document.body.appendChild(overlay);

    overlay.addEventListener('click', function (ev) {
      if (ev.target === overlay) closeConfigEditor();
    });

    overlay.querySelector('#__ef-ppt-cfg-close').addEventListener('click', closeConfigEditor);
  }

  // ════════════════════════════════════════════════════════════════════
  //  SHADOW DOM HOSTS + FLYOUT
  // ════════════════════════════════════════════════════════════════════
  // Two separate shadow hosts are used:
  //  - `host` / `boltShadow`  — just the bolt button, injected into the D365
  //    command bar (or a fixed fallback) so it visually sits with the other
  //    toolbar icons.
  //  - `panelHost` / `panelShadow` — the flyout panel, appended directly to
  //    <body>. Keeping it OUTSIDE the command bar's DOM subtree matters:
  //    D365's header often has an ancestor with a CSS transform/filter,
  //    which would silently hijack position:fixed to be relative to that
  //    ancestor instead of the viewport, rendering the panel off-screen.
  //
  // `host` is the element removed on Close (may be an <li> wrapper).
  // The inner shadow-hosting <div> is always a <div>, since attachShadow()
  // only supports a specific allowlist of tags and <li> is NOT one of them
  // (throws NotSupportedError, silently aborting the whole script).
  let host = document.getElementById('__ef-ppt-host');
  let boltShadow, panelShadow, panelHost;
  let _inToolbar = false;

  // True on any Dynamics 365 page — via the client API global once the page
  // has loaded, or the org hostname pattern otherwise (covers the moment
  // right after injection, before Xrm exists yet). Deliberately independent
  // of whether the current org is one of the user's *configured*
  // environments — the launcher should still work on an unconfigured org.
  function _isDynamicsPage() {
    try { if (window.Xrm && window.Xrm.Utility) return true; } catch (_) { /* ignore */ }
    return /(^|\.)crm\d*\.dynamics\.com$/i.test(window.location.hostname);
  }

  function _findCommandBar() {
    return document.querySelector('ul[data-id="CommandBar"]');
  }

  // Re-parents the floating bolt button into the D365 command bar once it
  // appears. attachShadow() is intrinsic to the element it was called on, so
  // moving that same <div> into a new wrapper keeps its shadow root (and
  // everything rendered inside it) intact — no need to rebuild it.
  function _upgradeToToolbar(cmdBar) {
    if (_inToolbar) return;
    const inner = host; // the floating mode's shadow-hosting <div>
    const li = document.createElement('li');
    li.id = '__ef-ppt-host';
    li.setAttribute('role', 'presentation');
    li.setAttribute('style', 'display:flex;align-items:center;list-style:none;position:relative;');
    inner.removeAttribute('id');
    inner.setAttribute('style', '');
    if (inner.parentNode) inner.parentNode.removeChild(inner);
    li.appendChild(inner);
    cmdBar.appendChild(li);
    host = li;
    _inToolbar = true;
    renderFlyout();
  }

  if (!host) {
    // Prefer injecting into the D365 global command bar so the button sits
    // alongside New / Notifications / Settings / Help instead of overlapping
    // the user-profile icon in the top-right corner.
    const cmdBar = _findCommandBar();
    const boltShadowHost = document.createElement('div');
    if (cmdBar) {
      host = document.createElement('li');
      host.id = '__ef-ppt-host';
      host.setAttribute('role', 'presentation');
      host.setAttribute('style', 'display:flex;align-items:center;list-style:none;position:relative;');
      host.appendChild(boltShadowHost);
      cmdBar.appendChild(host);
      _inToolbar = true;
    } else {
      // No command bar yet — show the floating button immediately (so
      // something is ALWAYS visible right away, on any page) rather than
      // waiting. If this genuinely is a D365 page that just hasn't finished
      // rendering its command bar (SPA hydration), watch for it and move the
      // button in once it shows up instead of leaving it stuck floating for
      // the rest of the session; a plain non-D365 page skips the watch
      // entirely and keeps the floating button, as intended.
      host = boltShadowHost;
      host.id = '__ef-ppt-host';
      host.setAttribute('style', 'all:initial;position:fixed;top:8px;right:8px;z-index:2147483647;');
      document.body.appendChild(host);

      if (_isDynamicsPage()) {
        const observer = new MutationObserver(function () {
          const bar = _findCommandBar();
          if (bar) {
            observer.disconnect();
            _upgradeToToolbar(bar);
          }
        });
        observer.observe(document.body, { childList: true, subtree: true });
        setTimeout(function () { observer.disconnect(); }, 8000);
      }
    }
    boltShadow = boltShadowHost.attachShadow({ mode: 'open' });

    panelHost = document.createElement('div');
    panelHost.id = '__ef-ppt-panel-host';
    panelHost.setAttribute('style', 'all:initial;position:absolute;top:0;left:0;width:0;height:0;overflow:visible;z-index:2147483647;');
    document.body.appendChild(panelHost);
    panelShadow = panelHost.attachShadow({ mode: 'open' });
  } else {
    _inToolbar = host.tagName.toLowerCase() === 'li';
    boltShadow = _inToolbar ? host.firstElementChild.shadowRoot : host.shadowRoot;
    panelHost = document.getElementById('__ef-ppt-panel-host');
    panelShadow = panelHost.shadowRoot;
  }

  // Close the flyout on any click outside it (or outside the bolt button).
  // Capture phase so this still runs even if D365's own handlers later call
  // stopPropagation() during bubbling. composedPath() is used (rather than
  // e.target) because it correctly crosses the shadow DOM boundary — a click
  // on an element inside panelShadow/boltShadow reports panelHost/host as
  // ancestors in the path even though they're separate shadow trees.
  document.addEventListener('click', function (ev) {
    const panel = panelShadow.querySelector('#ef-panel');
    if (!panel || !panel.classList.contains('open')) return;
    const path = ev.composedPath ? ev.composedPath() : [];
    if (path.indexOf(panelHost) !== -1 || path.indexOf(host) !== -1) return;
    panel.classList.remove('open');
  }, true);

  // A function (not a precomputed constant) because _inToolbar can flip from
  // false to true at runtime via _upgradeToToolbar(), after this file's top
  // level has already run once — the styles must reflect *current* mode.
  function buildBoltStyles() {
    return (
      ':host,*{box-sizing:border-box;}' +
      '.wrap{font-family:Segoe UI,system-ui,sans-serif;}' +
      (_inToolbar
        // Toolbar mode: transparent button that blends with the D365 header
        ? '.bolt-btn{width:36px;height:36px;background:transparent;color:#fff;border:none;' +
            'border-radius:6px;cursor:pointer;display:flex;align-items:center;' +
            'justify-content:center;padding:0;position:relative;}' +
          '.bolt-btn:hover{background:rgba(255,255,255,.15);}' +
          '.env-dot{position:absolute;bottom:3px;right:3px;width:7px;height:7px;' +
            'border-radius:50%;pointer-events:none;border:1.5px solid rgba(0,0,0,.25);}'
        // Fallback: fixed pill in top-right corner
        : '.bolt-btn{width:36px;height:36px;' +
            'background:#1B3A6B;color:#fff;border:none;border-radius:8px;cursor:pointer;' +
            'display:flex;align-items:center;justify-content:center;' +
            'border-left:4px solid #888;box-shadow:0 2px 8px rgba(0,0,0,.3);padding:0;position:relative;}' +
          '.bolt-btn:hover{filter:brightness(1.1);}' +
          '.env-dot{position:absolute;bottom:3px;right:3px;width:7px;height:7px;' +
            'border-radius:50%;pointer-events:none;border:1.5px solid rgba(0,0,0,.25);}')
    );
  }

  const PANEL_STYLES =
    ':host,*{box-sizing:border-box;}' +
    '.wrap{font-family:Segoe UI,system-ui,sans-serif;}' +
    '.panel{position:fixed;width:300px;background:#fff;' +
      'border-radius:10px;box-shadow:0 8px 30px rgba(0,0,0,.25);' +
      'overflow:hidden;display:none;color:#1e293b;font-size:13px;}' +
    '.panel.open{display:block;}' +
    // Wide mode: Record Details pane on the left, divider, original content on the right.
    '.panel.panel--wide{width:601px;display:none;flex-direction:row;align-items:stretch;' +
      'max-height:min(620px,calc(100vh - 32px));}' +
    '.panel.panel--wide.open{display:flex;}' +
    '.rd-pane{flex:0 0 300px;max-width:300px;overflow-y:auto;padding:12px 14px;}' +
    '.pane-divider{flex:0 0 1px;align-self:stretch;background:#e2e8f0;}' +
    '.main-pane{flex:1 1 300px;min-width:0;display:flex;flex-direction:column;overflow-y:auto;}' +
    '.hdr{padding:12px 14px;border-bottom:1px solid #e2e8f0;}' +
    '.hdr-row{display:flex;align-items:center;gap:8px;}' +
    '.hdr-name{font-weight:700;font-size:14px;}' +
    '.badge{display:inline-block;width:12px;height:12px;border-radius:3px;flex:0 0 auto;}' +
    '.hdr-url{font-size:11px;color:#64748b;margin-top:3px;word-break:break-all;}' +
    '.sect{padding:12px 14px;border-bottom:1px solid #e2e8f0;}' +
    '.sect:last-child{border-bottom:none;}' +
    '.lbl{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;' +
      'color:#64748b;margin-bottom:8px;}' +
    '.envs{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px;}' +
    '.env-chip{display:flex;align-items:center;gap:5px;padding:4px 8px;border-radius:6px;' +
      'border:1px solid #cbd5e1;background:#fff;cursor:pointer;font-size:12px;color:#334155;}' +
    '.env-chip.active{border-color:#1B3A6B;background:#eef2ff;font-weight:600;}' +
    '.types{display:flex;flex-wrap:wrap;gap:4px;margin-bottom:10px;}' +
    '.type-btn{padding:4px 9px;border-radius:14px;border:1px solid #cbd5e1;background:#fff;' +
      'cursor:pointer;font-size:11px;color:#475569;}' +
    '.type-btn.active{background:#1B3A6B;color:#fff;border-color:#1B3A6B;}' +
    '.app-select{width:100%;padding:6px 8px;margin-bottom:10px;border-radius:6px;' +
      'border:1px solid #cbd5e1;background:#fff;font-size:12px;color:#334155;}' +
    '.go-btn{width:100%;padding:8px;border:none;border-radius:8px;background:#1B3A6B;' +
      'color:#fff;cursor:pointer;font-size:13px;font-weight:600;}' +
    '.go-btn:hover{filter:brightness(1.1);}' +
    '.tools-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;}' +
    '.tool-btn{display:flex;align-items:center;gap:6px;padding:7px 8px;border:1px solid #e2e8f0;' +
      'border-radius:8px;background:#f8fafc;cursor:pointer;font-size:12px;color:#1e293b;' +
      'text-align:left;overflow:hidden;}' +
    '.tool-btn:hover{background:#eef2ff;border-color:#c7d2fe;}' +
    '.tool-btn svg{flex:0 0 auto;color:#1B3A6B;width:15px;height:15px;}' +
    '.tool-btn span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0;}' +
    '.ftr{padding:8px 14px;background:#f8fafc;display:flex;align-items:center;justify-content:space-between;}' +
    '.cfg-link{font-size:12px;color:#1B3A6B;cursor:pointer;text-decoration:none;background:none;border:none;padding:0;}' +
    '.cfg-link:hover{text-decoration:underline;}' +
    '.rm-link{font-size:12px;color:#94a3b8;cursor:pointer;background:none;border:none;padding:0;}' +
    '.rm-link:hover{color:#b91c1c;}' +
    '.empty{font-size:12px;color:#64748b;}' +
    // ── Record Details pane ──────────────────────────────────────────────
    '.rd-state{font-size:12px;color:#64748b;padding:6px 2px;}' +
    '.rd-state--error{color:#dc2626;}' +
    '.rd-cards{display:flex;flex-direction:column;gap:6px;margin-bottom:10px;}' +
    '.rd-card{background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:6px 9px;' +
      'display:flex;flex-direction:column;gap:1px;}' +
    '.rd-label{font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;' +
      'color:#64748b;margin-bottom:2px;}' +
    '.rd-value-row{display:flex;align-items:center;gap:5px;min-width:0;}' +
    '.rd-value{font-size:12.5px;font-weight:600;color:#1e293b;overflow:hidden;' +
      'text-overflow:ellipsis;white-space:nowrap;min-width:0;}' +
    '.rd-mono{font-size:10px;color:#64748b;font-family:Consolas,monospace;overflow:hidden;' +
      'text-overflow:ellipsis;white-space:nowrap;}' +
    '.rd-link{color:#2855a0;text-decoration:underline;text-underline-offset:2px;cursor:pointer;}' +
    '.rd-link:hover{opacity:.75;}' +
    '.rd-icon{display:inline-flex;align-items:center;flex-shrink:0;color:#64748b;}' +
    '.rd-icon svg{width:12px;height:12px;}' +
    '.rd-copy-btn{background:none;border:none;cursor:pointer;color:#94a3b8;padding:2px;' +
      'border-radius:3px;display:inline-flex;align-items:center;flex-shrink:0;opacity:.5;margin-left:auto;}' +
    '.rd-card:hover .rd-copy-btn{opacity:.85;}' +
    '.rd-copy-btn:hover{opacity:1;color:#2855a0;}' +
    '.rd-copy-btn.copied{color:#16a34a;opacity:1;}' +
    '.rd-copy-btn svg{width:11px;height:11px;}' +
    '.rd-clone-bar{border-top:1px solid #e2e8f0;padding-top:10px;}' +
    '.rd-clone-row{display:flex;gap:6px;align-items:center;}' +
    '.rd-clone-btn{display:inline-flex;align-items:center;gap:5px;flex-shrink:0;padding:6px 10px;' +
      'background:#1B3A6B;color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:600;' +
      'cursor:pointer;white-space:nowrap;}' +
    '.rd-clone-btn:hover:not(:disabled){filter:brightness(1.1);}' +
    '.rd-clone-btn:disabled{opacity:.55;cursor:not-allowed;}' +
    '.rd-clone-select{flex:1;min-width:0;padding:6px 7px;border:1px solid #cbd5e1;border-radius:6px;' +
      'font-size:11.5px;color:#334155;background:#fff;}' +
    '.rd-alert{margin-top:8px;padding:7px 9px;border-radius:6px;font-size:11px;line-height:1.45;}' +
    '.rd-alert--warning{background:#fffbeb;color:#92400e;border:1px solid #fde68a;}' +
    '.rd-alert--error{background:#fef2f2;color:#991b1b;border:1px solid #fecaca;}' +
    '.rd-btn-secondary{margin-top:6px;padding:5px 10px;border-radius:6px;border:1px solid #cbd5e1;' +
      'background:#fff;color:#334155;font-size:11px;font-weight:600;cursor:pointer;}' +
    '.rd-btn-secondary:hover:not(:disabled){background:#f8fafc;}' +
    '.rd-btn-secondary:disabled{opacity:.6;cursor:not-allowed;}' +
    '.rd-clone-success{margin-top:8px;display:flex;align-items:center;gap:5px;font-size:11.5px;' +
      'font-weight:600;color:#16a34a;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:6px;padding:6px 9px;}' +
    '.rd-clone-success a{color:#2855a0;text-decoration:underline;text-underline-offset:2px;font-weight:600;}' +
    '.rd-clone-error{margin-top:8px;display:flex;align-items:center;gap:6px;font-size:11.5px;color:#dc2626;' +
      'background:#fef2f2;border:1px solid #fecaca;border-radius:6px;padding:6px 9px;}' +
    '.rd-link-btn{background:none;border:none;padding:0;font-size:11px;font-weight:600;color:#b91c1c;' +
      'cursor:pointer;text-decoration:underline;text-underline-offset:2px;flex-shrink:0;}' +
    '.rd-link-btn:hover{color:#7f1d1d;}' +
    '.rd-mini-spinner{display:inline-block;width:9px;height:9px;border:1.5px solid rgba(27,58,107,.25);' +
      'border-top-color:#1B3A6B;border-radius:50%;vertical-align:-1px;animation:rdspin .7s linear infinite;}' +
    '.rd-clone-btn .rd-mini-spinner{border-color:rgba(255,255,255,.5);border-top-color:#fff;}' +
    '@keyframes rdspin{to{transform:rotate(360deg);}}';

  function buildBoltMarkup() {
    const env = currentEnv(CFG);
    const envColor = env && env.color ? env.color : '#888';
    const boltInner = _inToolbar
      ? BOLT_SVG + '<span class="env-dot" style="background:' + escHtml(envColor) + '"></span>'
      : BOLT_SVG;
    const boltStyle = _inToolbar
      ? ''
      : 'style="border-left-color:' + escHtml(envColor) + '"';

    return (
      '<div class="wrap">' +
        '<button type="button" class="bolt-btn" id="ef-bolt" title="EF Power Platform Tools" ' +
          boltStyle + '>' + boltInner + '</button>' +
      '</div>'
    );
  }

  function buildPanelMarkup() {
    const env = currentEnv(CFG);
    const envColor = env && env.color ? env.color : '#888';
    const envs = (CFG && CFG.environments) || [];

    // Header
    let hdr;
    if (env) {
      hdr =
        '<div class="hdr-row">' +
          '<span class="badge" style="background:' + escHtml(envColor) + '"></span>' +
          '<span class="hdr-name">' + escHtml(env.name || '') + '</span>' +
        '</div>' +
        '<div class="hdr-url">' + escHtml(env.url || '') + '</div>';
    } else {
      hdr =
        '<div class="hdr-row"><span class="hdr-name">Unknown environment</span></div>' +
        '<div class="hdr-url">' + escHtml(window.location.origin) + '</div>';
    }

    // Go-To env chips
    let envChips;
    if (envs.length) {
      envChips = envs.map(function (e) {
        const active = e.id === selectedTargetId ? ' active' : '';
        return (
          '<button type="button" class="env-chip' + active + '" data-env-id="' + escHtml(e.id) + '">' +
            '<span class="badge" style="background:' + escHtml(e.color || '#888') + '"></span>' +
            escHtml(e.name || e.id) +
          '</button>'
        );
      }).join('');
    } else {
      envChips = '<div class="empty">No environments configured.</div>';
    }

    // Type toggles
    const typeBtns = GOTO_TYPES.map(function (t) {
      const active = t.id === selectedType ? ' active' : '';
      return '<button type="button" class="type-btn' + active + '" data-type="' + t.id + '">' + escHtml(t.label) + '</button>';
    }).join('');

    // App dropdown — only shown when "App" is the selected Go-To type.
    let appSelect = '';
    if (selectedType === 'app') {
      const apps = configuredAppUniqueNames();
      if (apps.length) {
        if (!selectedAppUniqueName || apps.indexOf(selectedAppUniqueName) === -1) {
          selectedAppUniqueName = apps[0];
        }
        appSelect =
          '<select class="app-select" id="ef-app-select">' +
            apps.map(function (u) {
              const sel = u === selectedAppUniqueName ? ' selected' : '';
              return '<option value="' + escHtml(u) + '"' + sel + '>' + escHtml(appLabel(u)) + '</option>';
            }).join('') +
          '</select>';
      } else {
        appSelect = '<div class="empty" style="margin-bottom:10px;">No apps configured (settings.includedApps).</div>';
      }
    }

    // Tools grid
    const toolBtns = TOOLS.map(function (t) {
      return (
        '<button type="button" class="tool-btn" data-tool="' + escHtml(t.name) + '">' +
          (TOOL_ICONS[t.name] || '') +
          '<span>' + escHtml(t.label) + '</span>' +
        '</button>'
      );
    }).join('');

    const mainInner =
      '<div class="hdr">' + hdr + '</div>' +
      '<div class="sect">' +
        '<div class="lbl">Go To</div>' +
        '<div class="envs" id="ef-envs">' + envChips + '</div>' +
        '<div class="types" id="ef-types">' + typeBtns + '</div>' +
        appSelect +
        '<button type="button" class="go-btn" id="ef-go">Go →</button>' +
      '</div>' +
      '<div class="sect">' +
        '<div class="lbl">Tools</div>' +
        '<div class="tools-grid" id="ef-tools">' + toolBtns + '</div>' +
      '</div>' +
      '<div class="ftr">' +
        '<button type="button" class="cfg-link" id="ef-cfg">⚙ Config</button>' +
        '<button type="button" class="rm-link" id="ef-remove" title="Remove launcher from this page">✕ Close</button>' +
      '</div>';

    const hasRecordPane = !!_rdState;
    const panelBody = hasRecordPane
      ? '<div class="rd-pane" id="ef-rd-pane">' + buildRecordPaneMarkup() + '</div>' +
        '<div class="pane-divider"></div>' +
        '<div class="main-pane">' + mainInner + '</div>'
      : mainInner;

    return (
      '<div class="wrap">' +
        '<div class="panel' + (hasRecordPane ? ' panel--wide' : '') + '" id="ef-panel">' +
          panelBody +
        '</div>' +
      '</div>'
    );
  }

  // Positions the panel (fixed, body-level) just under/right of the bolt
  // button (which may live inside the D365 toolbar).
  function positionPanel() {
    const boltBtn = boltShadow.querySelector('#ef-bolt');
    const panel = panelShadow.querySelector('#ef-panel');
    if (!boltBtn || !panel) return;
    const rect = boltBtn.getBoundingClientRect();
    panel.style.top = (rect.bottom + 6) + 'px';
    panel.style.right = Math.max(8, window.innerWidth - rect.right) + 'px';
  }

  function renderFlyout() {
    const panelWasOpen = (function () {
      const p = panelShadow.querySelector('#ef-panel');
      return p ? p.classList.contains('open') : false;
    })();

    boltShadow.innerHTML = '<style>' + buildBoltStyles() + '</style>' + buildBoltMarkup();
    panelShadow.innerHTML = '<style>' + PANEL_STYLES + '</style>' + buildPanelMarkup();

    const boltBtn = boltShadow.querySelector('#ef-bolt');
    const panel = panelShadow.querySelector('#ef-panel');
    if (panelWasOpen) {
      panel.classList.add('open');
      // innerHTML was replaced, so any inline top/right positioning from a
      // previous positionPanel() call was lost along with the old element.
      positionPanel();
    }

    boltBtn.addEventListener('click', function () {
      const willOpen = !panel.classList.contains('open');
      if (willOpen) {
        // Re-check the record context on every open — D365's SPA navigation
        // can switch records without a full page reload, so what was cached
        // in _rdState from the last open may now be stale.
        ensureRecordDetailsLoaded(parseRecordContext());
        renderFlyout();
        const freshPanel = panelShadow.querySelector('#ef-panel');
        freshPanel.classList.add('open');
        positionPanel();
        return;
      }
      panel.classList.toggle('open');
    });

    // Env chips
    panelShadow.querySelectorAll('.env-chip').forEach(function (chip) {
      chip.addEventListener('click', function () {
        selectedTargetId = chip.getAttribute('data-env-id');
        panelShadow.querySelectorAll('.env-chip').forEach(function (c) { c.classList.remove('active'); });
        chip.classList.add('active');
      });
    });

    // Type buttons
    panelShadow.querySelectorAll('.type-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        selectedType = btn.getAttribute('data-type');
        // Re-render: the "App" type shows an extra dropdown that only exists
        // in the markup when selectedType === 'app', so a full rebuild is
        // needed either way (adding or removing that element).
        renderFlyout();
        if (selectedType === 'app') {
          ensureAppNames().then(function () {
            if (selectedType === 'app') renderFlyout(); // refresh labels once friendly names arrive
          });
        }
      });
    });

    // App dropdown (only present when selectedType === 'app')
    const appSelectEl = panelShadow.querySelector('#ef-app-select');
    if (appSelectEl) {
      appSelectEl.addEventListener('change', function () {
        selectedAppUniqueName = appSelectEl.value;
      });
    }

    // Go button
    panelShadow.querySelector('#ef-go').addEventListener('click', function () {
      if (!selectedTargetId) {
        const first = (CFG && CFG.environments && CFG.environments[0]);
        if (first) selectedTargetId = first.id;
        else return;
      }
      doGoTo();
    });

    // Tool buttons
    panelShadow.querySelectorAll('.tool-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        panel.classList.remove('open');
        openTool(btn.getAttribute('data-tool'));
      });
    });

    // Config link
    panelShadow.querySelector('#ef-cfg').addEventListener('click', openConfigEditor);

    // Remove / close launcher
    panelShadow.querySelector('#ef-remove').addEventListener('click', destroyLauncher);

    // Record Details pane wiring (present only when on a record page)
    panelShadow.querySelectorAll('.rd-copy-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        const text = btn.getAttribute('data-copy');
        navigator.clipboard.writeText(text).then(function () {
          btn.classList.add('copied');
          setTimeout(function () { btn.classList.remove('copied'); }, 1500);
        });
      });
    });
    const rdCloneTargetEl = panelShadow.querySelector('#ef-rd-clone-target');
    if (rdCloneTargetEl) {
      rdCloneTargetEl.addEventListener('change', function () { onRdCloneTargetChange(rdCloneTargetEl); });
    }
    const rdConnectBtn = panelShadow.querySelector('#ef-rd-connect-btn');
    if (rdConnectBtn) rdConnectBtn.addEventListener('click', onRdConnectClick);
    const rdCloneBtn = panelShadow.querySelector('#ef-rd-clone-btn');
    if (rdCloneBtn) rdCloneBtn.addEventListener('click', onRdCloneClick);
    const rdErrBtn = panelShadow.querySelector('#ef-rd-clone-err-details');
    if (rdErrBtn) {
      rdErrBtn.addEventListener('click', function () { openRdErrorModal(_rdCloneResult && _rdCloneResult.error); });
    }
  }

  // ── Destroy ─────────────────────────────────────────────────────────
  function destroyLauncher() {
    const cfgOv = document.getElementById('__ef-ppt-config-overlay');
    if (cfgOv) cfgOv.remove();
    const toolOv = document.getElementById('__ef-ppt-tool-overlay');
    if (toolOv) toolOv.remove();
    const rdErrOv = document.getElementById('__ef-ppt-rd-err-overlay');
    if (rdErrOv) rdErrOv.remove();
    if (panelHost && panelHost.parentNode) panelHost.parentNode.removeChild(panelHost);
    if (host && host.parentNode) host.parentNode.removeChild(host);
    delete window.__EF_PPT_LAUNCHER;
  }

  // ── Bootstrap ───────────────────────────────────────────────────────
  // No automatic "no config found" prompt — the flyout renders regardless
  // (with an empty-state message under Go To if nothing's configured yet)
  // and the user can open Config from the footer link whenever they want.
  ensureRecordDetailsLoaded(parseRecordContext());
  renderFlyout();

})();
