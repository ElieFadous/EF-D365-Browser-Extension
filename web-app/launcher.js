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
      localStorage.setItem(CONFIG_KEY, JSON.stringify(cfg, null, 2));
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
  const BOLT_SVG =
    '<svg viewBox="0 0 64 64" width="20" height="20" aria-hidden="true">' +
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
  //  TOOL OVERLAY  (full-screen iframe, appended to body)
  // ════════════════════════════════════════════════════════════════════
  function closeTool() {
    const ex = document.getElementById('__ef-ppt-tool-overlay');
    if (ex) ex.remove();
  }

  function openTool(toolName, extraParams) {
    closeTool();
    const env = currentEnv(CFG);

    const params = new URLSearchParams();
    params.set('env', window.location.origin);
    if (env) {
      params.set('name', env.name || '');
      if (env.powerAppsId) params.set('paEnvId', env.powerAppsId);
    }
    params.set('_inModal', '1');
    if (extraParams && typeof extraParams === 'object') {
      Object.keys(extraParams).forEach(function (k) {
        if (extraParams[k] != null) params.set(k, extraParams[k]);
      });
    }

    const src =
      BASE_URL + '/' + toolName + '/' + toolName + '.html?' + params.toString();

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
        'flex-direction:column',
        'margin:0',
        'padding:0'
      ].join(';')
    );

    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.title = 'Close';
    closeBtn.setAttribute('aria-label', 'Close tool');
    closeBtn.textContent = '✕';
    closeBtn.setAttribute(
      'style',
      [
        'position:absolute',
        'top:10px',
        'right:14px',
        'width:34px',
        'height:34px',
        'border:none',
        'border-radius:8px',
        'background:#1B3A6B',
        'color:#fff',
        'font-size:16px',
        'line-height:1',
        'cursor:pointer',
        'z-index:2',
        'box-shadow:0 2px 8px rgba(0,0,0,.35)'
      ].join(';')
    );
    closeBtn.addEventListener('click', closeTool);

    const iframe = document.createElement('iframe');
    iframe.src = src;
    iframe.setAttribute(
      'style',
      [
        'flex:1 1 auto',
        'width:100%',
        'height:100%',
        'border:none',
        'background:#fff'
      ].join(';')
    );
    iframe.setAttribute('allow', 'clipboard-write');

    overlay.appendChild(closeBtn);
    overlay.appendChild(iframe);
    document.body.appendChild(overlay);
  }

  // ════════════════════════════════════════════════════════════════════
  //  GO TO  logic
  // ════════════════════════════════════════════════════════════════════
  async function resolveCurrentAppUniqueName() {
    try {
      const ver = apiVersion(CFG);
      const cur = new URLSearchParams(window.location.search);
      const appId = cur.get('appid');
      if (!appId) return '';
      const url =
        window.location.origin +
        '/api/data/' + ver +
        '/appmodules?$select=uniquename&$filter=appmoduleid eq ' + appId;
      const res = await fetch(url, {
        credentials: 'include',
        headers: {
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0'
        }
      });
      if (!res.ok) return '';
      const data = await res.json();
      if (data && data.value && data.value[0]) return data.value[0].uniquename || '';
      return '';
    } catch (err) {
      console.error('[EF PPT] Failed to resolve app unique name:', err);
      return '';
    }
  }

  function parseRecordContext() {
    // Read entity + record id from the current D365 URL.
    const p = new URLSearchParams(window.location.search);
    return {
      etn: p.get('etn') || '',
      id: (p.get('id') || '').replace(/[{}]/g, '')
    };
  }

  async function doGoTo() {
    const target =
      (CFG && CFG.environments || []).find(function (e) { return e.id === selectedTargetId; });
    if (!target) return;

    const ver = apiVersion(CFG);
    let targetUrl = '';

    switch (selectedType) {
      case 'open-in': {
        const ctx = parseRecordContext();
        const uniquename = await resolveCurrentAppUniqueName();
        const qp = new URLSearchParams();
        if (uniquename) qp.set('appuniquename', uniquename);
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
        const dflt = CFG && CFG.settings && CFG.settings.defaultAppUniqueName;
        targetUrl = dflt
          ? target.url + '/main.aspx?appuniquename=' + encodeURIComponent(dflt)
          : target.url + '/main.aspx';
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
  //  CONFIG EDITOR OVERLAY  (appended to body, not shadow DOM)
  // ════════════════════════════════════════════════════════════════════
  function bookmarkletCode() {
    const loaderSrc = BASE_URL + '/launcher.js';
    return (
      "javascript:(function(){var s=document.createElement('script');" +
      "s.src='" + loaderSrc + "?v='+Date.now();document.body.appendChild(s);})();"
    );
  }

  function closeConfigEditor() {
    const ex = document.getElementById('__ef-ppt-config-overlay');
    if (ex) ex.remove();
  }

  function openConfigEditor() {
    closeConfigEditor();
    const current = CFG || STARTER_CONFIG;
    const json = JSON.stringify(current, null, 2);
    const bm = bookmarkletCode();

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
        '<h2 style="margin:0 0 4px;font-size:18px;color:#1B3A6B;">EF PPT — Configure</h2>' +
        '<p style="margin:0 0 12px;font-size:12px;color:#475569;">Edit the JSON config below. It is stored in this browser’s localStorage.</p>' +
        '<textarea id="__ef-ppt-cfg-text" spellcheck="false" style="width:100%;height:280px;box-sizing:border-box;font-family:Consolas,monospace;font-size:12px;border:1px solid #cbd5e1;border-radius:8px;padding:10px;resize:vertical;">' +
          escHtml(json) +
        '</textarea>' +
        '<div id="__ef-ppt-cfg-err" style="color:#b91c1c;font-size:12px;min-height:16px;margin:6px 2px;"></div>' +
        '<div style="display:flex;gap:8px;justify-content:flex-end;margin-bottom:16px;">' +
          '<button id="__ef-ppt-cfg-cancel" type="button" style="padding:8px 16px;border:1px solid #cbd5e1;background:#fff;border-radius:8px;cursor:pointer;font-size:13px;">Cancel</button>' +
          '<button id="__ef-ppt-cfg-save" type="button" style="padding:8px 16px;border:none;background:#1B3A6B;color:#fff;border-radius:8px;cursor:pointer;font-size:13px;">Save</button>' +
        '</div>' +
        '<div style="border-top:1px solid #e2e8f0;padding-top:12px;">' +
          '<div style="font-size:12px;font-weight:600;color:#334155;margin-bottom:4px;">Bookmarklet</div>' +
          '<p style="margin:0 0 6px;font-size:11px;color:#64748b;">Create a new bookmark and paste this as the URL:</p>' +
          '<textarea readonly spellcheck="false" onclick="this.select()" style="width:100%;height:64px;box-sizing:border-box;font-family:Consolas,monospace;font-size:11px;border:1px solid #cbd5e1;border-radius:8px;padding:8px;background:#f8fafc;color:#334155;">' +
            escHtml(bm) +
          '</textarea>' +
        '</div>' +
      '</div>';

    document.body.appendChild(overlay);

    overlay.addEventListener('click', function (ev) {
      if (ev.target === overlay) closeConfigEditor();
    });

    const errEl = overlay.querySelector('#__ef-ppt-cfg-err');
    overlay.querySelector('#__ef-ppt-cfg-cancel').addEventListener('click', closeConfigEditor);
    overlay.querySelector('#__ef-ppt-cfg-save').addEventListener('click', function () {
      const txt = overlay.querySelector('#__ef-ppt-cfg-text').value;
      let parsed;
      try {
        parsed = JSON.parse(txt);
      } catch (err) {
        errEl.textContent = 'Invalid JSON: ' + (err && err.message ? err.message : err);
        return;
      }
      if (!parsed || !Array.isArray(parsed.environments)) {
        errEl.textContent = 'Config must contain an "environments" array.';
        return;
      }
      if (!saveConfig(parsed)) {
        errEl.textContent = 'Could not write to localStorage.';
        return;
      }
      CFG = parsed;
      closeConfigEditor();
      renderFlyout();
    });
  }

  // ════════════════════════════════════════════════════════════════════
  //  SHADOW DOM HOST + FLYOUT
  // ════════════════════════════════════════════════════════════════════
  let host = document.getElementById('__ef-ppt-host');
  let shadow;

  if (!host) {
    host = document.createElement('div');
    host.id = '__ef-ppt-host';
    host.setAttribute('style', 'all:initial;position:fixed;top:0;right:0;z-index:2147483647;');
    document.body.appendChild(host);
    shadow = host.attachShadow({ mode: 'open' });
  } else {
    shadow = host.shadowRoot;
  }

  const STYLES =
    ':host,*{box-sizing:border-box;}' +
    '.wrap{font-family:Segoe UI,system-ui,sans-serif;}' +
    '.bolt-btn{position:fixed;top:8px;right:8px;width:36px;height:36px;' +
      'background:#1B3A6B;color:#fff;border:none;border-radius:8px;cursor:pointer;' +
      'display:flex;align-items:center;justify-content:center;' +
      'border-left:4px solid #888;box-shadow:0 2px 8px rgba(0,0,0,.3);padding:0;}' +
    '.bolt-btn:hover{filter:brightness(1.1);}' +
    '.panel{position:fixed;top:52px;right:8px;width:300px;background:#fff;' +
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
    '.go-btn{width:100%;padding:8px;border:none;border-radius:8px;background:#1B3A6B;' +
      'color:#fff;cursor:pointer;font-size:13px;font-weight:600;}' +
    '.go-btn:hover{filter:brightness(1.1);}' +
    '.tools-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;}' +
    '.tool-btn{display:flex;align-items:center;gap:6px;padding:8px;border:1px solid #e2e8f0;' +
      'border-radius:8px;background:#f8fafc;cursor:pointer;font-size:12px;color:#1e293b;' +
      'text-align:left;}' +
    '.tool-btn:hover{background:#eef2ff;border-color:#c7d2fe;}' +
    '.tool-btn svg{flex:0 0 auto;color:#1B3A6B;}' +
    '.ftr{padding:8px 14px;background:#f8fafc;text-align:right;}' +
    '.cfg-link{font-size:12px;color:#1B3A6B;cursor:pointer;text-decoration:none;background:none;border:none;padding:0;}' +
    '.cfg-link:hover{text-decoration:underline;}' +
    '.empty{font-size:12px;color:#64748b;}';

  function buildFlyoutMarkup() {
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
        '<button type="button" class="bolt-btn" id="ef-bolt" title="EF Power Platform Tools" ' +
          'style="border-left-color:' + escHtml(envColor) + '">' + BOLT_SVG + '</button>' +
        '<div class="panel" id="ef-panel">' +
          '<div class="hdr">' + hdr + '</div>' +
          '<div class="sect">' +
            '<div class="lbl">Go To</div>' +
            '<div class="envs" id="ef-envs">' + envChips + '</div>' +
            '<div class="types" id="ef-types">' + typeBtns + '</div>' +
            '<button type="button" class="go-btn" id="ef-go">Go →</button>' +
          '</div>' +
          '<div class="sect">' +
            '<div class="lbl">Tools</div>' +
            '<div class="tools-grid" id="ef-tools">' + toolBtns + '</div>' +
          '</div>' +
          '<div class="ftr">' +
            '<button type="button" class="cfg-link" id="ef-cfg">⚙ Config</button>' +
          '</div>' +
        '</div>' +
      '</div>'
    );
  }

  function renderFlyout() {
    const panelWasOpen = (function () {
      const p = shadow.querySelector('#ef-panel');
      return p ? p.classList.contains('open') : false;
    })();

    shadow.innerHTML = '<style>' + STYLES + '</style>' + buildFlyoutMarkup();

    const boltBtn = shadow.querySelector('#ef-bolt');
    const panel = shadow.querySelector('#ef-panel');
    if (panelWasOpen) panel.classList.add('open');

    boltBtn.addEventListener('click', function () {
      panel.classList.toggle('open');
    });

    // Env chips
    shadow.querySelectorAll('.env-chip').forEach(function (chip) {
      chip.addEventListener('click', function () {
        selectedTargetId = chip.getAttribute('data-env-id');
        shadow.querySelectorAll('.env-chip').forEach(function (c) { c.classList.remove('active'); });
        chip.classList.add('active');
      });
    });

    // Type buttons
    shadow.querySelectorAll('.type-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        selectedType = btn.getAttribute('data-type');
        shadow.querySelectorAll('.type-btn').forEach(function (b) { b.classList.remove('active'); });
        btn.classList.add('active');
      });
    });

    // Go button
    shadow.querySelector('#ef-go').addEventListener('click', function () {
      if (!selectedTargetId) {
        const first = (CFG && CFG.environments && CFG.environments[0]);
        if (first) selectedTargetId = first.id;
        else return;
      }
      doGoTo();
    });

    // Tool buttons
    shadow.querySelectorAll('.tool-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        panel.classList.remove('open');
        openTool(btn.getAttribute('data-tool'));
      });
    });

    // Config link
    shadow.querySelector('#ef-cfg').addEventListener('click', openConfigEditor);
  }

  // ── Bootstrap ───────────────────────────────────────────────────────
  renderFlyout();

  // No config? Open the editor automatically so the user can set up.
  if (!CFG) {
    openConfigEditor();
  }

})();
