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

  // ── State ───────────────────────────────────────────────────────────
  let CFG = loadConfig();
  let selectedTargetId = null;
  let selectedType = 'open-in';
  let selectedAppUniqueName = null;

  // ════════════════════════════════════════════════════════════════════
  //  FETCH BRIDGE
  // ════════════════════════════════════════════════════════════════════
  window.addEventListener('message', async function (e) {
    const msg = e.data;
    if (!msg || typeof msg !== 'object') return;

    if (msg.__efppt === 'close-overlay') {
      closeTool();
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

    try {
      const headers = Object.assign(
        {
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0'
        },
        msg.headers || {}
      );
      const method = msg.method || 'GET';
      const init = {
        method: method,
        credentials: 'include',
        headers: headers
      };
      if (msg.body != null && method !== 'GET' && method !== 'HEAD') {
        init.body =
          typeof msg.body === 'string' ? msg.body : JSON.stringify(msg.body);
        if (!headers['Content-Type'] && !headers['content-type']) {
          headers['Content-Type'] = 'application/json';
        }
      }

      const res = await fetch(msg.url, init);
      const text = await res.text();
      let data = text;
      if (text) {
        try { data = JSON.parse(text); } catch (_) { /* leave as text */ }
      }

      if (!res.ok) {
        reply({ ok: false, status: res.status, error: (data && data.error && data.error.message) || res.statusText || ('HTTP ' + res.status), data: data });
      } else {
        reply({ ok: true, status: res.status, data: data });
      }
    } catch (err) {
      console.error('[EF PPT] Bridge fetch failed:', err);
      reply({ ok: false, status: 0, error: String(err && err.message ? err.message : err) });
    }
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

  if (!host) {
    // Prefer injecting into the D365 global command bar so the button sits
    // alongside New / Notifications / Settings / Help instead of overlapping
    // the user-profile icon in the top-right corner.
    const cmdBar = document.querySelector('ul[data-id="CommandBar"]');
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
      host = boltShadowHost;
      host.id = '__ef-ppt-host';
      host.setAttribute('style', 'all:initial;position:fixed;top:8px;right:8px;z-index:2147483647;');
      document.body.appendChild(host);
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

  const BOLT_STYLES =
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
          'border-radius:50%;pointer-events:none;border:1.5px solid rgba(0,0,0,.25);}');

  const PANEL_STYLES =
    ':host,*{box-sizing:border-box;}' +
    '.wrap{font-family:Segoe UI,system-ui,sans-serif;}' +
    '.panel{position:fixed;width:300px;background:#fff;' +
      'border-radius:10px;box-shadow:0 8px 30px rgba(0,0,0,.25);' +
      'overflow:hidden;display:none;color:#1e293b;font-size:13px;}' +
    '.panel.open{display:block;}' +
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
    '.empty{font-size:12px;color:#64748b;}';

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

    return (
      '<div class="wrap">' +
        '<div class="panel" id="ef-panel">' +
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
          '</div>' +
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

    boltShadow.innerHTML = '<style>' + BOLT_STYLES + '</style>' + buildBoltMarkup();
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
      if (willOpen) positionPanel();
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
  }

  // ── Destroy ─────────────────────────────────────────────────────────
  function destroyLauncher() {
    const cfgOv = document.getElementById('__ef-ppt-config-overlay');
    if (cfgOv) cfgOv.remove();
    const toolOv = document.getElementById('__ef-ppt-tool-overlay');
    if (toolOv) toolOv.remove();
    if (panelHost && panelHost.parentNode) panelHost.parentNode.removeChild(panelHost);
    if (host && host.parentNode) host.parentNode.removeChild(host);
    delete window.__EF_PPT_LAUNCHER;
  }

  // ── Bootstrap ───────────────────────────────────────────────────────
  renderFlyout();
  if (!CFG) openConfigEditor();

})();
