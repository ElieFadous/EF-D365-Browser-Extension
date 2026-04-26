/**
 * EF Power Platform Tools — Popup Script
 *
 * Architecture note: this file is intentionally structured so adding a new
 * feature only requires:
 *   1. A new <button data-section="myfeature"> in the nav.
 *   2. A new <section id="section-myfeature"> in main.
 *   3. Registering an init function in FEATURE_REGISTRY below.
 */

// ─── Feature Registry ─────────────────────────────────────────────────────────

/**
 * Maps nav tab data-section values to their initialiser functions.
 * Add future features here.
 */
const FEATURE_REGISTRY = {
  goto:             initGoTo,
  metadata:         initMetadata,
  ribbon:           initRibbon,
  'plugin-trace':   initPluginTrace,
  'flows':          initFlows,
  'data-sync':      initDataSync,
  'record-details': initRecordDetails,
};

// ─── State ────────────────────────────────────────────────────────────────────

let environments    = [];
let settings        = { apiVersion: 'v9.2' };
let currentTab      = null;
let currentEnv      = null;
let currentWindowId = null; // used to open new tabs in the same window (preserves incognito)

// Cache for EntitySetName lookups (LogicalName → set name), keyed by "baseUrl:etn".
const _entitySetNameCache = new Map();

// ─── Bootstrap ────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {
  try {
    [environments, settings] = await Promise.all([loadEnvironments(), loadSettings()]);
  } catch (e) {
    showStatus('Could not load environments: ' + e.message, 'error');
    return;
  }

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  currentTab      = tab ?? null;
  currentEnv      = detectEnvironment(tab?.url);
  currentWindowId = tab?.windowId ?? null; // used as fallback in _openTab / GOTO_ENVIRONMENT

  // Stamp the footer version from manifest so it never goes stale.
  const { version } = chrome.runtime.getManifest();
  const vEl = document.getElementById('footer-version');
  if (vEl) vEl.textContent = `v${version}`;

  renderEnvBadge();
  updateHeaderTheme();
  setupNavTabs();
  setupSidebar();

  // Show the Record Details tab and make it the default when on a D365 record form with an ID.
  const _initCtx = parseD365PageContext(currentTab?.url);
  if (currentEnv && _initCtx?.pagetype === 'entityrecord' && _initCtx?.id) {
    const rdTab = document.querySelector('[data-section="record-details"]');
    rdTab?.classList.remove('hidden');
    // Switch default active tab to Record Details.
    document.querySelectorAll('.nav-tab').forEach(t => {
      t.classList.remove('active');
      t.setAttribute('aria-selected', 'false');
    });
    document.querySelectorAll('.feature-section').forEach(s => s.classList.remove('active'));
    rdTab?.classList.add('active');
    rdTab?.setAttribute('aria-selected', 'true');
    document.getElementById('section-record-details')?.classList.add('active');
    document.body.classList.add('rd-active');
  }

  // Initialise the default active section.
  const activeSection = document.querySelector('.nav-tab.active')?.dataset.section;
  if (activeSection && FEATURE_REGISTRY[activeSection]) {
    FEATURE_REGISTRY[activeSection]();
  }

  // Footer settings link.
  document.getElementById('btn-settings').addEventListener('click', () => {
    chrome.runtime.openOptionsPage();
  });

});

// ─── Tab / modal helpers ──────────────────────────────────────────────────────

/**
 * Opens a new tab in the correct browser window.
 *
 * For https:// URLs (Go To, API Go To, etc.) Edge correctly honours windowId,
 * so they open in the InPrivate window when launched from InPrivate.
 *
 * For chrome-extension:// URLs (tool pages) Edge categorically blocks them in
 * InPrivate tabs — even typing the URL manually in InPrivate is blocked with
 * ERR_BLOCKED_BY_CLIENT. There is no programmatic workaround; they always open
 * in the regular window. The tool pages still function normally from there.
 */
async function _openTab(url, active = true) {
  let targetWindowId = currentWindowId;

  try {
    const fw = await chrome.runtime.sendMessage({ type: 'GET_FOCUSED_WINDOW' });
    if (fw?.windowId != null) targetWindowId = fw.windowId;
  } catch (e) {}

  try {
    const opts = { url, active };
    if (targetWindowId !== null) opts.windowId = targetWindowId;
    await chrome.tabs.create(opts);
  } catch (e) {
    await chrome.tabs.create({ url, active });
  }
}

/**
 * Injects the tool as a full-screen modal overlay into the current D365 tab.
 * Falls back to opening a new tab if injection is not possible (no D365 tab,
 * or host-permission mismatch).
 *
 * @param {string} toolName   - Identifier used as the modal's DOM key, e.g. 'metadata'
 * @param {string} toolTitle  - Human-readable title shown in the modal title bar
 * @param {string} toolUrl    - Full chrome-extension:// URL (with any query params)
 */
async function _injectTool(toolName, toolTitle, toolUrl) {
  if (currentTab?.id != null) {
    try {
      const result = await chrome.runtime.sendMessage({
        type:      'INJECT_TOOL',
        tabId:     currentTab.id,
        toolName,
        toolTitle,
        toolUrl,
      });
      if (!result?.error) return; // Success — modal injected.
    } catch (e) {}
  }
  // Fallback: open as a regular tab (e.g. not on a D365 tab, or scripting blocked).
  await _openTab(toolUrl);
}

// ─── Environment helpers ──────────────────────────────────────────────────────

async function loadEnvironments() {
  // Prefer user-saved environments (from options page).
  const stored = await chrome.storage.local.get('environments');
  if (Array.isArray(stored.environments) && stored.environments.length > 0) {
    return stored.environments;
  }
  // Fall back to bundled environments.json.
  const res  = await fetch(chrome.runtime.getURL('environments.json'));
  const data = await res.json();
  return data.environments;
}

async function loadSettings() {
  const stored = await chrome.storage.local.get('settings');
  const raw = {
    apiVersion:           'v9.2',
    clonePrefix:          '',
    cloneWhitelist:       null,
    cloneLookupMode:      'skip',
    defaultAppUniqueName: '',
    includedApps:         [],
    syncBatchSize:        250,
    ...stored.settings,
  };
  // Validate API version format — must be vX.Y (e.g. v9.2). Default to v9.2 if invalid.
  if (!raw.apiVersion || !/^v\d+\.\d+$/.test(String(raw.apiVersion).trim())) {
    raw.apiVersion = 'v9.2';
  }
  return raw;
}

/** Persists a partial settings update to chrome.storage.local. */
async function _saveSettings(partial) {
  settings = { ...settings, ...partial };
  await chrome.storage.local.set({ settings });
}

function detectEnvironment(url) {
  if (!url) return null;
  try {
    const parsed = new URL(url);

    // 1. Exact D365 origin match (primary path).
    const byOrigin = environments.find(env => env.url === parsed.origin);
    if (byOrigin) return byOrigin;

    // 2. make.powerapps.com/environments/{paEnvId}/… — match by powerAppsId if configured.
    if (parsed.hostname === 'make.powerapps.com') {
      const m = parsed.pathname.match(/^\/environments\/([^/]+)/i);
      if (m) {
        const paId = m[1].toLowerCase();
        return environments.find(
          env => env.powerAppsId && env.powerAppsId.toLowerCase() === paId
        ) ?? null;
      }
    }

    return null;
  } catch {
    return null;
  }
}

/**
 * Returns true only when the active tab is on a real D365 origin
 * (not make.powerapps.com or any other site).
 * Used to gate the "Open In…" feature which requires a D365 page context.
 */
function _isOnD365Tab() {
  if (!currentTab?.url || !currentEnv) return false;
  try {
    return new URL(currentTab.url).origin === new URL(currentEnv.url).origin;
  } catch {
    return false;
  }
}

function hasAppId(url) {
  if (!url) return false;
  try {
    return new URLSearchParams(new URL(url).search).has('appid');
  } catch {
    return false;
  }
}

// ─── Header theme ─────────────────────────────────────────────────────────────

/**
 * Darkens a hex colour by `amount` (0–1 fraction).
 * e.g. _darkenColor('#1B3A6B', 0.15) → slightly darker navy.
 */
function _darkenColor(hex, amount = 0.15) {
  const h = hex.replace('#', '');
  const r = Math.max(0, Math.round(parseInt(h.slice(0, 2), 16) * (1 - amount)));
  const g = Math.max(0, Math.round(parseInt(h.slice(2, 4), 16) * (1 - amount)));
  const b = Math.max(0, Math.round(parseInt(h.slice(4, 6), 16) * (1 - amount)));
  return '#' + [r, g, b].map(x => x.toString(16).padStart(2, '0')).join('');
}

/**
 * Returns '#000' or '#fff' — whichever has better contrast against hexColor.
 * Uses the WCAG relative-luminance formula.
 */
function _contrastFg(hexColor) {
  const h = hexColor.replace('#', '');
  const toLinear = x => {
    const c = parseInt(x, 16) / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  };
  const L = 0.2126 * toLinear(h.slice(0,2))
          + 0.7152 * toLinear(h.slice(2,4))
          + 0.0722 * toLinear(h.slice(4,6));
  return L > 0.179 ? '#000' : '#fff';
}

function updateHeaderTheme() {
  const root  = document.documentElement;
  const badge = document.getElementById('current-env-badge');

  let bg, fg;

  if (currentEnv?.color) {
    bg = currentEnv.color;
    fg = _contrastFg(bg);
  } else {
    bg = '#6b7280';   // neutral gray when no env detected
    fg = '#ffffff';
  }

  const isDark    = fg === '#fff';
  const fgMuted   = isDark ? 'rgba(255,255,255,0.60)' : 'rgba(0,0,0,0.50)';
  const logoBg    = isDark ? 'rgba(255,255,255,0.18)' : 'rgba(0,0,0,0.12)';
  const badgeBg   = isDark ? 'rgba(255,255,255,0.16)' : 'rgba(0,0,0,0.10)';
  const badgeBdr  = isDark ? 'rgba(255,255,255,0.28)' : 'rgba(0,0,0,0.18)';

  root.style.setProperty('--header-bg',       bg);
  root.style.setProperty('--header-fg',       fg);
  root.style.setProperty('--header-fg-muted', fgMuted);
  root.style.setProperty('--header-logo-bg',  logoBg);
  root.style.setProperty('--tab-accent',      _darkenColor(bg, 0.15));

  // Badge sits inside the header so it needs its own contrast treatment.
  badge.style.background  = badgeBg;
  badge.style.borderColor = badgeBdr;
  badge.style.color       = fg;
  badge.style.border      = `1px solid ${badgeBdr}`;
}

function renderEnvBadge() {
  const badge = document.getElementById('current-env-badge');
  if (currentEnv) {
    badge.textContent = currentEnv.name;
    badge.className   = `env-badge env-badge--${currentEnv.id}`;
  } else {
    badge.textContent = 'Unknown';
    badge.className   = 'env-badge env-badge--unknown';
  }
}

// ─── Nav tab switching ────────────────────────────────────────────────────────

function setupNavTabs() {
  document.querySelectorAll('.nav-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      // Deactivate all.
      document.querySelectorAll('.nav-tab').forEach(b => {
        b.classList.remove('active');
        b.setAttribute('aria-selected', 'false');
      });
      document.querySelectorAll('.feature-section').forEach(s => s.classList.remove('active'));

      // Activate clicked.
      btn.classList.add('active');
      btn.setAttribute('aria-selected', 'true');
      const section = document.getElementById(`section-${btn.dataset.section}`);
      if (section) section.classList.add('active');

      // Run initialiser (idempotent—features should guard against double-init).
      const key = btn.dataset.section;
      document.body.classList.toggle('rd-active', key === 'record-details');
      // Collapse pane width when leaving record-details tab.
      if (key !== 'record-details') {
        document.documentElement.classList.remove('rd-pane-open');
        document.body.classList.remove('rd-pane-open');
      }
      if (FEATURE_REGISTRY[key]) FEATURE_REGISTRY[key]();
      // Force a layout reflow on the record-details pane so its width is
      // calculated correctly the first time the section becomes visible.
      if (key === 'record-details') {
        requestAnimationFrame(() => {
          const pane = document.querySelector('.rd-open-in-pane');
          if (pane) void pane.offsetWidth; // trigger reflow
        });
      }
    });
  });
}

// ─── Sidebar ──────────────────────────────────────────────────────────────────

function setupSidebar() {
  const toggle = document.getElementById('sidebar-toggle');

  function _applyExpanded(expanded) {
    // Both html AND body must carry the class so Chrome resizes the popup window.
    document.documentElement.classList.toggle('nav-expanded', expanded);
    document.body.classList.toggle('nav-expanded', expanded);
  }

  // Restore saved state; default is expanded unless explicitly saved as false.
  chrome.storage.local.get('navExpanded', ({ navExpanded }) => {
    _applyExpanded(navExpanded !== false);
  });

  toggle.addEventListener('click', () => {
    const expanded = !document.body.classList.contains('nav-expanded');
    _applyExpanded(expanded);
    chrome.storage.local.set({ navExpanded: expanded });
  });
}

// ─── Go To Feature ────────────────────────────────────────────────────────────

let gotoInitialised = false;

function initGoTo() {
  if (gotoInitialised) return;
  gotoInitialised = true;

  // Hide the "Open In…" sub-tab and its panel when the user is not on a D365 tab.
  if (!_isOnD365Tab()) {
    const openInTab   = document.querySelector('.goto-tab[data-goto-tab="open-in"]');
    const openInPanel = document.getElementById('goto-panel-open-in');
    if (openInTab)   openInTab.classList.add('hidden');
    if (openInPanel) openInPanel.classList.add('hidden');
    // Activate the first remaining visible sub-tab automatically.
    const firstVisible = document.querySelector('.goto-tab:not(.hidden)');
    if (firstVisible) firstVisible.click();
  }

  _wireGotoSubTabs();
  _initOpenInPanel();
  _initApiPanel();
  _initSolutionsPanel();   // async, fire-and-forget
  _initAppPanel();         // async, fire-and-forget
  _initSecurityPanel();
}

// ── Sub-tab switching ──────────────────────────────────────────────────────────

function _wireGotoSubTabs() {
  document.querySelectorAll('.goto-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.goto-tab').forEach(b => {
        b.classList.remove('active');
        b.setAttribute('aria-selected', 'false');
      });
      document.querySelectorAll('.goto-panel').forEach(p => p.classList.add('hidden'));
      btn.classList.add('active');
      btn.setAttribute('aria-selected', 'true');
      document.getElementById(`goto-panel-${btn.dataset.gotoTab}`).classList.remove('hidden');
    });
  });
}

// ── Panel 1: Open This Page In… ───────────────────────────────────────────────

function _initOpenInPanel() {
  const elLoading = document.getElementById('goto-loading');
  const elNotD365 = document.getElementById('goto-not-d365');
  const elNoAppId = document.getElementById('goto-no-appid');
  const elEnvList = document.getElementById('goto-env-list');
  const elButtons = document.getElementById('goto-buttons');

  elLoading.classList.add('hidden');

  if (!currentEnv) {
    elNotD365.classList.remove('hidden');
    return;
  }

  if (!hasAppId(currentTab?.url)) {
    elNoAppId.classList.remove('hidden');
    return;
  }

  // Build buttons for every environment that is NOT the current one and enabled for this feature.
  const targets = environments.filter(env => env.id !== currentEnv.id && isEnabledFor(env, 'goto-open-in'));

  if (targets.length === 0) {
    elNotD365.textContent = 'No other environments configured.';
    elNotD365.classList.remove('hidden');
    return;
  }

  targets.forEach(env => {
    elButtons.appendChild(buildGoToButton(env));
  });

  elEnvList.classList.remove('hidden');
}

function buildGoToButton(env) {
  const btn = document.createElement('button');
  btn.className = `goto-btn${env.warn ? ' goto-btn--warn' : ''}`;
  btn.style.setProperty('--env-color', env.color ?? '#1B3A6B');

  const shortUrl = env.url.replace('https://', '');
  btn.innerHTML = `
    <div class="goto-btn__info">
      <span class="goto-btn__name">${escHtml(env.name)}</span>
      <span class="goto-btn__url">${escHtml(shortUrl)}</span>
      ${env.warn ? '<span class="goto-btn__warn-tag">Production</span>' : ''}
    </div>
    <span class="goto-btn__arrow">&#8250;</span>
  `;

  btn.addEventListener('click', () => handleGoTo(env));
  return btn;
}

async function handleGoTo(targetEnv) {
  setAllGoToBtnsDisabled(true);
  showStatus(`Resolving app in ${targetEnv.name}\u2026`, 'loading');

  try {
    const response = await chrome.runtime.sendMessage({
      type:         'GOTO_ENVIRONMENT',
      currentTabId: currentTab.id,
      currentUrl:   currentTab.url,
      windowId:     currentWindowId,
      targetEnv
    });

    if (response?.error) {
      showStatus(response.error, 'error');
      setAllGoToBtnsDisabled(false);
    } else {
      showStatus(`Opened in ${targetEnv.name}.`, 'success');
      setTimeout(() => window.close(), 900);
    }
  } catch (e) {
    showStatus(e.message, 'error');
    setAllGoToBtnsDisabled(false);
  }
}

function setAllGoToBtnsDisabled(disabled) {
  document.querySelectorAll('#goto-buttons .goto-btn, .rd-env-btn').forEach(btn => {
    btn.disabled = disabled;
  });
}

// ── Panel 2: API ──────────────────────────────────────────────────────────────

async function _initApiPanel() {
  const apiVersion = settings.apiVersion || 'v9.2';
  const container  = document.getElementById('api-buttons');

  let pageContext   = null;
  let entitySetName = null;

  if (currentTab?.url && currentEnv) {
    pageContext = parseD365PageContext(currentTab.url);
  }

  // If we detected an entity type, fetch its OData set name for smart URLs.
  if (pageContext?.etn && currentTab?.id) {
    const hint = document.createElement('div');
    hint.className = 'api-context-hint';
    hint.innerHTML = '<span class="spinner" style="width:10px;height:10px;border-width:1.5px"></span> Detecting page context…';
    container.appendChild(hint);

    try {
      entitySetName = await _fetchEntitySetName(currentEnv.url, pageContext.etn);
    } catch (_) { /* fall through — use base URL */ }

    hint.remove();
  }

  const apiEnvs = environments.filter(env => isEnabledFor(env, 'goto-api'));

  apiEnvs.forEach(env => {
    let targetUrl, sublabel;

    if (entitySetName && pageContext?.pagetype === 'entityrecord' && pageContext?.id) {
      targetUrl = `${env.url}/api/data/${apiVersion}/${entitySetName}(${pageContext.id})`;
      sublabel  = `${entitySetName}(${pageContext.id.slice(0, 8)}…)`;
    } else if (entitySetName && pageContext?.pagetype === 'entitylist') {
      targetUrl = `${env.url}/api/data/${apiVersion}/${entitySetName}`;
      sublabel  = `api/data/${apiVersion}/${entitySetName}`;
    } else {
      targetUrl = `${env.url}/api/data/${apiVersion}/`;
      sublabel  = `api/data/${apiVersion}`;
    }

    container.appendChild(_buildDirectLinkButton(env, targetUrl, sublabel));
  });
}

// ── D365 URL context parser ────────────────────────────────────────────────────

/**
 * Parses a Dynamics 365 main.aspx URL and returns { pagetype, etn, id } or null.
 * Handles both query-string params and hash-encoded JSON (UCI format).
 */
function parseD365PageContext(url) {
  if (!url) return null;
  try {
    const u = new URL(url);

    // Primary: standard query string params (classic / unified)
    let pagetype = u.searchParams.get('pagetype');
    let etn      = u.searchParams.get('etn');
    let id       = u.searchParams.get('id');

    // Secondary: hash-based JSON routing used by newer UCI pages
    if (!pagetype && u.hash) {
      try {
        const hashStr = u.hash.startsWith('#') ? u.hash.slice(1) : u.hash;
        // Some hashes are raw JSON, some are key=value pairs
        let parsed = null;
        try { parsed = JSON.parse(decodeURIComponent(hashStr)); } catch (_) {}
        if (!parsed) {
          const hp = new URLSearchParams(hashStr);
          const state = hp.get('state');
          if (state) try { parsed = JSON.parse(decodeURIComponent(state)); } catch (_) {}
        }
        if (parsed) {
          pagetype = pagetype ?? parsed.pagetype ?? null;
          etn      = etn      ?? parsed.etn      ?? null;
          id       = id       ?? parsed.id        ?? null;
        }
      } catch (_) { /* ignore malformed hash */ }
    }

    if (!pagetype) return null;

    // Strip curly braces from GUIDs — D365 sometimes wraps them: {xxxxxxxx-…}
    const cleanId = id ? id.replace(/[{}]/g, '') : null;

    return { pagetype, etn: etn ?? null, id: cleanId };
  } catch {
    return null;
  }
}

// ── Entity set name resolver ───────────────────────────────────────────────────

/**
 * Fetches EntitySetName for a D365 entity logical name via a content script
 * injected into the current D365 tab (uses the page's auth cookies).
 * Results are cached in _entitySetNameCache for the lifetime of the popup.
 */
async function _fetchEntitySetName(baseUrl, etn) {
  const cacheKey = `${baseUrl}:${etn}`;
  if (_entitySetNameCache.has(cacheKey)) return _entitySetNameCache.get(cacheKey);

  const apiVersion = settings.apiVersion || 'v9.2';
  const apiUrl = `${baseUrl}/api/data/${apiVersion}/EntityDefinitions(LogicalName='${etn}')?$select=EntitySetName`;

  const [injection] = await chrome.scripting.executeScript({
    target: { tabId: currentTab.id },
    func: async (url) => {
      try {
        const res = await fetch(url, {
          headers: {
            'Accept':           'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version':    '4.0',
          },
        });
        if (!res.ok) return null;
        const data = await res.json();
        return data.EntitySetName ?? null;
      } catch { return null; }
    },
    args: [apiUrl],
  });

  const entitySetName = injection?.result ?? null;
  if (entitySetName) _entitySetNameCache.set(cacheKey, entitySetName);
  return entitySetName;
}

// ── Generic tab script executor ────────────────────────────────────────────────

/**
 * Executes an async function inside the current D365 tab (so it runs with the
 * page's authentication cookies) and returns its result.
 */
async function _executeInTab(func, args = []) {
  const [injection] = await chrome.scripting.executeScript({
    target: { tabId: currentTab.id },
    func,
    args,
  });
  return injection?.result ?? null;
}

/** Like _executeInTab but targets an explicit tabId (used for cross-env proxy tabs). */
async function _executeInProxyTab(tabId, func, args = []) {
  const [injection] = await chrome.scripting.executeScript({
    target: { tabId },
    func,
    args,
  });
  return injection?.result ?? null;
}

/** Resolves the OData entity set name for an entity in a proxy tab (cross-env). */
async function _fetchEntitySetNameInTab(tabId, apiBase, etn) {
  const cacheKey = `proxy:${tabId}:${etn}`;
  if (_entitySetNameCache.has(cacheKey)) return _entitySetNameCache.get(cacheKey);
  const result = await _executeInProxyTab(tabId, async (url) => {
    try {
      const res = await fetch(url, { headers: { Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' } });
      if (!res.ok) return null;
      const d = await res.json();
      return d.EntitySetName ?? null;
    } catch { return null; }
  }, [`${apiBase}/EntityDefinitions(LogicalName='${etn}')?$select=EntitySetName`]);
  if (result) _entitySetNameCache.set(cacheKey, result);
  return result;
}

// ── Clone Record ───────────────────────────────────────────────────────────────

/**
 * Clones a Dynamics 365 record:
 *  1. Fetches entity metadata (writable attributes + M2M relationships).
 *  2. Fetches the full record (all fields + lookup annotations).
 *  3. Builds a POST payload — transforms lookups to @odata.bind, prepends
 *     configured prefix (e.g. "[COPY] ") to the primary name field (cropped to MaxLength).
 *  4. POSTs to create the new record.
 *  5. Re-creates every many-to-many association.
 *
 * Returns { newId, entitySetName } on success.
 */
async function cloneRecord(envUrl, etn, id, entitySetName, targetEnvUrl = null) {
  const isCrossEnv    = !!targetEnvUrl && targetEnvUrl !== envUrl;
  const apiBase       = `${envUrl}/api/data/${settings.apiVersion || 'v9.2'}`;
  const targetApiBase = isCrossEnv ? `${targetEnvUrl}/api/data/${settings.apiVersion || 'v9.2'}` : apiBase;

  // ── 1. Entity definition — primary name attr + writable attrs + M2M ────────
  const meta = await _executeInTab(async (apiBase, etn) => {
    const h = { 'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' };
    try {
      const [defR, attrR, m2oR] = await Promise.all([
        fetch(
          `${apiBase}/EntityDefinitions(LogicalName='${etn}')?$select=PrimaryNameAttribute` +
          `&$expand=ManyToManyRelationships($select=SchemaName,Entity1LogicalName,Entity2LogicalName,Entity1NavigationPropertyName,Entity2NavigationPropertyName)`,
          { headers: h }
        ),
        // MaxLength is NOT on the base AttributeMetadata type — omit it here
        fetch(
          `${apiBase}/EntityDefinitions(LogicalName='${etn}')/Attributes` +
          `?$select=LogicalName,AttributeType,IsValidForCreate`,
          { headers: h }
        ),
        // ManyToOneRelationships gives us the correct single-valued navigation property
        // names to use for @odata.bind (the attr LogicalName is NOT always the nav prop name)
        fetch(
          `${apiBase}/EntityDefinitions(LogicalName='${etn}')/ManyToOneRelationships` +
          `?$select=ReferencingAttribute,ReferencingEntityNavigationPropertyName`,
          { headers: h }
        ),
      ]);
      if (!defR.ok)  return { error: `EntityDefinitions ${defR.status}` };
      if (!attrR.ok) return { error: `Attributes ${attrR.status}` };
      const def      = await defR.json();
      const attrData = await attrR.json();

      // Build a map: attribute logical name → correct OData navigation property name
      let navPropMap = {};
      if (m2oR.ok) {
        const m2oData = await m2oR.json();
        navPropMap = Object.fromEntries(
          (m2oData.value ?? []).map(r => [r.ReferencingAttribute, r.ReferencingEntityNavigationPropertyName])
        );
      }

      // Fetch MaxLength for the primary name field via the typed StringAttributeMetadata endpoint
      let primaryMaxLen = 100;
      const primaryName = def.PrimaryNameAttribute;
      if (primaryName) {
        const lenR = await fetch(
          `${apiBase}/EntityDefinitions(LogicalName='${etn}')/Attributes/Microsoft.Dynamics.CRM.StringAttributeMetadata` +
          `?$select=LogicalName,MaxLength&$filter=LogicalName eq '${primaryName}'`,
          { headers: h }
        );
        if (lenR.ok) {
          const lenData = await lenR.json();
          primaryMaxLen = lenData.value?.[0]?.MaxLength ?? 100;
        }
      }

      return {
        primaryName,
        primaryMaxLen,
        navPropMap,
        m2m:   def.ManyToManyRelationships ?? [],
        attrs: attrData.value ?? [],
      };
    } catch (e) { console.error('[EF PPT proxy]', e); return { error: e.message }; }
  }, [apiBase, etn]);

  if (meta?.error) throw new Error(meta.error);
  if (!meta)       throw new Error('Failed to fetch entity metadata.');

  // ── 2. Full record (all fields + annotation headers for lookups) ────────────
  const record = await _executeInTab(async (url) => {
    try {
      const res = await fetch(url, {
        headers: {
          'Accept':           'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version':    '4.0',
          'Prefer':           'odata.include-annotations="OData.Community.Display.V1.FormattedValue,Microsoft.Dynamics.CRM.lookuplogicalname"',
        },
      });
      if (!res.ok) return { error: `Record fetch ${res.status}` };
      return await res.json();
    } catch (e) { console.error('[EF PPT proxy]', e); return { error: e.message }; }
  }, [`${apiBase}/${entitySetName}(${id})`]);

  if (record?.error) throw new Error(record.error);
  if (!record)       throw new Error('Failed to fetch record data.');

  // ── 3. Build POST payload ───────────────────────────────────────────────────
  const { primaryName, primaryMaxLen, navPropMap, m2m, attrs } = meta;
  const EXCLUDE = new Set([
    `${etn}id`,
    'createdon', 'modifiedon', 'overriddencreatedon',
    'createdby', 'modifiedby', 'createdonbehalfby', 'modifiedonbehalfby',
    'versionnumber', 'exchangerate', 'importsequencenumber',
    'timezoneruleversionnumber', 'utcconversiontimezonecode',
    'owningbusinessunit', 'owninguser', 'owningteam',
  ]);
  const SKIP_TYPES = new Set(['Virtual', 'EntityName', 'ManagedProperty', 'Uniqueidentifier']);

  const payload      = {};
  const lookupFields = []; // { logicalName, guid, targetEtn }

  for (const attr of attrs) {
    if (!attr.IsValidForCreate) continue;
    const name = attr.LogicalName;
    if (EXCLUDE.has(name)) continue;
    if (SKIP_TYPES.has(attr.AttributeType)) continue;

    if (attr.AttributeType === 'Lookup' || attr.AttributeType === 'Customer' || attr.AttributeType === 'Owner') {
      const guidKey   = `_${name}_value`;
      const guid      = record[guidKey];
      if (!guid) continue;
      const targetEtn = record[`${guidKey}@Microsoft.Dynamics.CRM.lookuplogicalname`];
      if (targetEtn)  lookupFields.push({ logicalName: name, guid, targetEtn });
    } else {
      const val = record[name];
      if (val === null || val === undefined) continue;

      if (name === primaryName && !isCrossEnv) {
        const rawPrefix = settings.clonePrefix?.trim() ?? '';
        const prefix    = rawPrefix ? `[${rawPrefix}] ` : '';
        const full      = prefix + String(val);
        payload[name]   = full.length > primaryMaxLen ? full.slice(0, primaryMaxLen) : full;
      } else {
        payload[name] = val;
      }
    }
  }

  // Resolve entity set names for all lookup targets (parallel, cached)
  const uniqueTargetEtns = [...new Set(lookupFields.map(f => f.targetEtn))];
  const etnToSet = Object.fromEntries(
    await Promise.all(uniqueTargetEtns.map(async t => [t, await _fetchEntitySetName(envUrl, t)]))
  );
  for (const { logicalName, guid, targetEtn } of lookupFields) {
    const targetSet = etnToSet[targetEtn];
    if (!targetSet) continue;
    // Use the declared OData navigation property name, not the raw attribute logical name
    const navProp = navPropMap[logicalName] ?? logicalName;
    payload[`${navProp}@odata.bind`] = `/${targetSet}(${guid})`;
  }

  // ── 4 & 5. Create / upsert + M2M ────────────────────────────────────────────

  // Separate lookup bindings from scalar payload (needed for cross-env skip mode)
  const scalarPayload = {};
  const lookupBindings = {}; // { key: value } where key ends in @odata.bind
  for (const [k, v] of Object.entries(payload)) {
    if (k.endsWith('@odata.bind')) lookupBindings[k] = v;
    else scalarPayload[k] = v;
  }

  if (isCrossEnv) {
    // ── Cross-env: open proxy tab on target, PATCH (upsert) with same record ID ──
    let proxyTabId = null, proxyReused = false;
    const proxyResult = await chrome.runtime.sendMessage({
      type: 'OPEN_PROXY_TAB', env: targetEnvUrl, windowId: currentWindowId,
    });
    if (proxyResult?.error) throw new Error(`Cannot connect to target environment: ${proxyResult.error}`);
    proxyTabId  = proxyResult.tabId;
    proxyReused = proxyResult.reused;

    try {
      const targetEntitySetName = await _fetchEntitySetNameInTab(proxyTabId, targetApiBase, etn);
      if (!targetEntitySetName) throw new Error(`Cannot resolve entity set for '${etn}' in target environment.`);

      // Step 3b — Validate + remap fields against target environment schema
      const targetSchema = await _executeInProxyTab(proxyTabId, async (apiBase, etn) => {
        const h = { Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' };
        try {
          const [attrR, m2oR] = await Promise.all([
            fetch(`${apiBase}/EntityDefinitions(LogicalName='${etn}')/Attributes?$select=LogicalName,IsValidForCreate`, { headers: h }),
            fetch(`${apiBase}/EntityDefinitions(LogicalName='${etn}')/ManyToOneRelationships?$select=ReferencingAttribute,ReferencingEntityNavigationPropertyName`, { headers: h }),
          ]);
          const attrNames = attrR.ok ? ((await attrR.json()).value ?? []).filter(a => a.IsValidForCreate).map(a => a.LogicalName) : null;
          const navMap    = m2oR.ok ? Object.fromEntries(((await m2oR.json()).value ?? []).map(r => [r.ReferencingAttribute, r.ReferencingEntityNavigationPropertyName])) : {};
          return { attrNames, navMap };
        } catch (e) { console.error('[EF PPT]', 'Failed to fetch target schema:', e); return null; }
      }, [targetApiBase, etn]);

      const tAttrSet = targetSchema?.attrNames ? new Set(targetSchema.attrNames) : null;

      // Filter scalar payload: only fields that exist in target
      const validatedScalar = {};
      for (const [k, v] of Object.entries(scalarPayload)) {
        if (!tAttrSet || tAttrSet.has(k)) validatedScalar[k] = v;
      }

      // Rebuild lookup bindings: use target nav prop map + only fields existing in target
      const validatedLookups = {};
      for (const { logicalName, guid, targetEtn } of lookupFields) {
        if (tAttrSet && !tAttrSet.has(logicalName)) continue; // field missing in target — skip
        const targetLookupSet = await _fetchEntitySetNameInTab(proxyTabId, targetApiBase, targetEtn)
                             ?? etnToSet[targetEtn];
        if (!targetLookupSet) continue;
        const navProp = targetSchema?.navMap?.[logicalName] ?? navPropMap[logicalName] ?? logicalName;
        validatedLookups[`${navProp}@odata.bind`] = `/${targetLookupSet}(${guid})`;
      }

      // Step 4a — PATCH scalar fields with self-healing retry.
      // D365 may reject fields that are present in the source but absent from the target
      // entity schema (e.g. custom fields only in one environment). When the error message
      // identifies the offending field by name, that field is removed and the PATCH retried
      // automatically. Up to 30 such removals are attempted before giving up.
      const scalarResult = await _executeInProxyTab(proxyTabId, async (url, bodyEntries) => {
        const headers = { 'Content-Type': 'application/json', 'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0', 'MSCRM.SuppressDuplicateDetection': 'true' };
        let remaining = Object.fromEntries(bodyEntries);
        let lastError = null;

        for (let attempt = 0; attempt < 30; attempt++) {
          if (Object.keys(remaining).length === 0) return { ok: true };
          try {
            const res = await fetch(url, { method: 'PATCH', headers, body: JSON.stringify(remaining) });
            if (res.ok) return { ok: true };
            const errText = await res.text().catch(() => res.statusText);
            lastError = `Upsert failed (${res.status}): ${errText}`;

            // Try to extract field name from "XYZ field missing from target entity"
            // or "Attribute logical name missing from target entity: xyz"
            const m = errText.match(/[`'"]?([a-z_][a-z0-9_]*)[`'"]?\s+(?:field\s+)?missing\s+from\s+target\s+entity/i)
                   ?? errText.match(/missing\s+from\s+target\s+entity[^:]*:\s*([a-z_][a-z0-9_]*)/i);
            const badField = m?.[1];
            if (badField && badField in remaining) {
              delete remaining[badField];
              continue; // retry without that field
            }
            return { error: lastError }; // non-field error — fail immediately
          } catch (e) { console.error('[EF PPT proxy]', e); return { error: e.message }; }
        }
        return lastError ? { error: lastError } : { ok: true };
      }, [`${targetApiBase}/${targetEntitySetName}(${id})`, Object.entries(validatedScalar)]);

      if (scalarResult?.error) throw new Error(scalarResult.error);

      // Step 4b — PATCH each lookup individually (skip or fail on error)
      const failOnLookup = settings.cloneLookupMode === 'fail';
      for (const [bindKey, bindVal] of Object.entries(validatedLookups)) {
        const lResult = await _executeInProxyTab(proxyTabId, async (url, body) => {
          try {
            const res = await fetch(url, {
              method: 'PATCH',
              headers: { 'Content-Type': 'application/json', 'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0', 'MSCRM.SuppressDuplicateDetection': 'true' },
              body: JSON.stringify(body),
            });
            if (!res.ok) { const e = await res.text().catch(() => res.statusText); return { error: `Lookup PATCH failed (${res.status}): ${e}` }; }
            return { ok: true };
          } catch (e) { console.error('[EF PPT proxy]', e); return { error: e.message }; }
        }, [`${targetApiBase}/${targetEntitySetName}(${id})`, { [bindKey]: bindVal }]);

        if (lResult?.error && failOnLookup) throw new Error(lResult.error);
      }

      // Step 5 — exact M2M sync in target
      const h = { 'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' };
      for (const rel of m2m) {
        const isEntity1      = rel.Entity1LogicalName === etn;
        const myNavProp      = isEntity1 ? rel.Entity1NavigationPropertyName : rel.Entity2NavigationPropertyName;
        const relatedEtn     = isEntity1 ? rel.Entity2LogicalName            : rel.Entity1LogicalName;
        if (!myNavProp || !relatedEtn) continue;

        const relatedSetSource = await _fetchEntitySetName(envUrl, relatedEtn);
        const relatedSetTarget = await _fetchEntitySetNameInTab(proxyTabId, targetApiBase, relatedEtn);
        if (!relatedSetSource || !relatedSetTarget) continue;

        const pkField = `${relatedEtn}id`;

        const sourceIds = await _executeInTab(async (url, pk) => {
          try {
            const res = await fetch(`${url}&$select=${pk}`, { headers: { Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' } });
            if (!res.ok) return [];
            const d = await res.json();
            return (d.value ?? []).map(r => r[pk]).filter(Boolean);
          } catch (e) { console.error('[EF PPT]', 'Failed to fetch source M2M IDs:', e); return []; }
        }, [`${apiBase}/${entitySetName}(${id})/${myNavProp}?$top=500`, pkField]);

        const targetIds = await _executeInProxyTab(proxyTabId, async (url, pk) => {
          try {
            const res = await fetch(`${url}&$select=${pk}`, { headers: { Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' } });
            if (!res.ok) return [];
            const d = await res.json();
            return (d.value ?? []).map(r => r[pk]).filter(Boolean);
          } catch (e) { console.error('[EF PPT]', 'Failed to fetch target M2M IDs:', e); return []; }
        }, [`${targetApiBase}/${targetEntitySetName}(${id})/${myNavProp}?$top=500`, pkField]);

        const srcSet = new Set(sourceIds);
        const tgtSet = new Set(targetIds);

        // Add missing associations
        for (const rid of sourceIds) {
          if (tgtSet.has(rid)) continue;
          await _executeInProxyTab(proxyTabId, async (assocUrl, body) => {
            try { await fetch(assocUrl, { method: 'POST', headers: { 'Content-Type': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }, body: JSON.stringify(body) }); } catch (_) { console.error('[EF PPT]', 'M2M associate failed:', _); }
          }, [`${targetApiBase}/${targetEntitySetName}(${id})/${myNavProp}/$ref`, { '@odata.id': `${targetApiBase}/${relatedSetTarget}(${rid})` }]);
        }

        // Remove extra associations
        for (const rid of targetIds) {
          if (srcSet.has(rid)) continue;
          await _executeInProxyTab(proxyTabId, async (url) => {
            try { await fetch(url, { method: 'DELETE', headers: { 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' } }); } catch (_) { console.error('[EF PPT]', 'M2M disassociate failed:', _); }
          }, [`${targetApiBase}/${targetEntitySetName}(${id})/${myNavProp}(${rid})/$ref`]);
        }
      }

      return { newId: id, entitySetName: targetEntitySetName, crossEnv: true, targetEnvUrl };

    } finally {
      if (!proxyReused) chrome.runtime.sendMessage({ type: 'CLOSE_PROXY_TAB', tabId: proxyTabId, reused: false }).catch(() => {});
    }

  } else {
    // ── Same-env: POST (create new record) ─────────────────────────────────────
    const fullPayload = { ...scalarPayload, ...lookupBindings };
    const createResult = await _executeInTab(async (url, body) => {
      try {
        const res = await fetch(url, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json', 'Accept': 'application/json',
            'OData-MaxVersion': '4.0', 'OData-Version': '4.0', 'Prefer': 'return=minimal',
            'MSCRM.SuppressDuplicateDetection': 'true',
          },
          body: JSON.stringify(body),
        });
        if (!res.ok) { const e = await res.text().catch(() => res.statusText); return { error: `Create failed (${res.status}): ${e}` }; }
        const entityId = res.headers.get('OData-EntityId') ?? '';
        const match    = entityId.match(/\(([^)]+)\)$/);
        return { newId: match?.[1] ?? null };
      } catch (e) { return { error: e.message }; }
    }, [`${apiBase}/${entitySetName}`, fullPayload]);

    if (createResult?.error) throw new Error(createResult.error);
    const newId = createResult?.newId;
    if (!newId) throw new Error('Clone created but new record ID could not be determined.');

    // M2M: add all source associations to the new clone
    for (const rel of m2m) {
      const isEntity1      = rel.Entity1LogicalName === etn;
      const myNavProp      = isEntity1 ? rel.Entity1NavigationPropertyName : rel.Entity2NavigationPropertyName;
      const relatedEtn     = isEntity1 ? rel.Entity2LogicalName            : rel.Entity1LogicalName;
      if (!myNavProp || !relatedEtn) continue;

      const relatedEntitySet = await _fetchEntitySetName(envUrl, relatedEtn);
      if (!relatedEntitySet) continue;

      const relatedIds = await _executeInTab(async (url, pkField) => {
        try {
          const res = await fetch(`${url}&$select=${pkField}`, {
            headers: { 'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' },
          });
          if (!res.ok) return [];
          const data = await res.json();
          return (data.value ?? []).map(r => r[pkField]).filter(Boolean);
        } catch (e) { console.error('[EF PPT]', 'Failed to fetch related M2M IDs:', e); return []; }
      }, [`${apiBase}/${entitySetName}(${id})/${myNavProp}?$top=500`, `${relatedEtn}id`]);

      for (const relatedId of relatedIds) {
        await _executeInTab(async (assocUrl, body) => {
          try {
            await fetch(assocUrl, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' },
              body: JSON.stringify(body),
            });
          } catch (_) { console.error('[EF PPT]', 'M2M associate failed:', _); /* best-effort */ }
        }, [
          `${apiBase}/${entitySetName}(${newId})/${myNavProp}/$ref`,
          { '@odata.id': `${apiBase}/${relatedEntitySet}(${relatedId})` },
        ]);
      }
    }

    return { newId, entitySetName };
  }
}

// ── Panel 3: Solutions ────────────────────────────────────────────────────────

function _initSolutionsPanel() {
  const container = document.getElementById('solutions-buttons');

  for (const env of environments.filter(env => isEnabledFor(env, 'goto-solutions'))) {
    if (env.powerAppsId) {
      const url = `https://make.powerapps.com/environments/${env.powerAppsId}/solutions/`;
      container.appendChild(_buildDirectLinkButton(env, url, 'Solutions'));
    } else {
      container.appendChild(_buildSolutionsMissingIdButton(env));
    }
  }
}

function _buildSolutionsMissingIdButton(env) {
  const el = document.createElement('div');
  el.className = `goto-btn goto-btn--no-id${env.warn ? ' goto-btn--warn' : ''}`;
  el.style.setProperty('--env-color', env.color ?? '#1B3A6B');
  el.innerHTML = `
    <div class="goto-btn__info">
      <span class="goto-btn__name">${escHtml(env.name)}</span>
      <span class="goto-btn__url">${escHtml(env.url.replace('https://', ''))}</span>
    </div>
    <span class="goto-btn__no-id-tag">No PA Env ID</span>
  `;
  return el;
}

// ── Panel 4: App ──────────────────────────────────────────────────────────────

async function _initAppPanel() {
  const statusEl   = document.getElementById('app-status');
  const selectorEl = document.getElementById('app-selector');
  const appSelect  = document.getElementById('app-select');
  const envSection = document.getElementById('app-env-section');
  const envButtons = document.getElementById('app-buttons');

  // ── Rebuild "Open in environment" buttons for the chosen app ──
  function _rebuildEnvButtons(uniquename) {
    envButtons.innerHTML = '';
    if (!uniquename) {
      envButtons.innerHTML = '<div class="state-msg state-info">Select an app above first.</div>';
      envSection.classList.remove('hidden');
      return;
    }
    if (environments.length === 0) {
      envButtons.innerHTML = '<div class="state-msg state-info">No environments configured.</div>';
      envSection.classList.remove('hidden');
      return;
    }
    environments.forEach(env => {
      const url = `${env.url}/apps/uniquename/${encodeURIComponent(uniquename)}`;
      envButtons.appendChild(_buildDirectLinkButton(env, url, uniquename));
    });
    envSection.classList.remove('hidden');
  }

  // ── Read stored app cache + settings ──
  const stored        = await chrome.storage.local.get('appCache');
  const appCache      = stored.appCache ?? {};
  const defaultApp    = settings.defaultAppUniqueName ?? '';
  const includedApps  = Array.isArray(settings.includedApps) ? settings.includedApps : [];

  // Merge all apps from cache
  const allApps = new Map(); // uniquename → display name
  for (const apps of Object.values(appCache)) {
    for (const a of apps) {
      if (!allApps.has(a.uniquename)) allApps.set(a.uniquename, a.name || a.uniquename);
    }
  }

  // Keep only: default app (always) + explicitly included apps
  const visibleApps = [...allApps.entries()]
    .filter(([uname]) => uname === defaultApp || includedApps.includes(uname))
    .sort((a, b) => a[1].localeCompare(b[1]));

  statusEl.style.display = 'none';

  if (visibleApps.length === 0) {
    // Nothing configured yet — guide user to Settings
    statusEl.textContent = allApps.size === 0
      ? 'No apps loaded yet. Open Settings → Model Driven Apps to load and configure apps.'
      : 'No apps selected. Open Settings → Model Driven Apps to choose which apps appear here.';
    statusEl.className    = 'state-msg state-info';
    statusEl.style.display = '';
    return;
  }

  // ── Build <select> from visible apps ──
  appSelect.innerHTML = '';

  for (const [uname, label] of visibleApps) {
    const opt = document.createElement('option');
    opt.value       = uname;
    opt.textContent = label;
    if (uname === defaultApp) opt.selected = true;
    appSelect.appendChild(opt);
  }

  selectorEl.classList.remove('hidden');

  if (!appSelect._changeListenerAdded) {
    appSelect._changeListenerAdded = true;
    appSelect.addEventListener('change', () => _rebuildEnvButtons(appSelect.value));
  }

  _rebuildEnvButtons(appSelect.value || defaultApp);
}

// ── Panel 5: Security ─────────────────────────────────────────────────────────

function _initSecurityPanel() {
  const container = document.getElementById('security-buttons');
  environments.filter(env => isEnabledFor(env, 'goto-security')).forEach(env => {
    const url = `${env.url}/tools/AdminSecurity/adminsecurity_area.aspx`;
    container.appendChild(_buildDirectLinkButton(env, url, 'Admin Security'));
  });
}

// ── Shared direct-link button builder ─────────────────────────────────────────

function _buildDirectLinkButton(env, targetUrl, sublabel) {
  const btn = document.createElement('button');
  btn.className = `goto-btn${env.warn ? ' goto-btn--warn' : ''}`;
  btn.style.setProperty('--env-color', env.color ?? '#1B3A6B');

  const shortUrl = targetUrl.replace('https://', '');
  btn.innerHTML = `
    <div class="goto-btn__info">
      <span class="goto-btn__name">${escHtml(env.name)}</span>
      <span class="goto-btn__url">${escHtml(shortUrl)}</span>
      ${env.warn ? '<span class="goto-btn__warn-tag">Production</span>' : ''}
    </div>
    <span class="goto-btn__arrow">&#8250;</span>
  `;

  btn.addEventListener('click', async () => {
    await _openTab(targetUrl);
    window.close();
  });
  return btn;
}

// ─── Metadata Browser Feature ─────────────────────────────────────────────────

let metadataInitialised = false;

function initMetadata() {
  if (metadataInitialised) return;
  metadataInitialised = true;

  const section = document.getElementById('section-metadata');

  // If not on a recognised D365 environment, default to the first configured env.
  const metaEnvs   = environments.filter(env => isEnabledFor(env, 'metadata'));
  const defaultEnv = (metaEnvs.find(e => e.id === currentEnv?.id)) ?? metaEnvs[0] ?? null;

  if (metaEnvs.length === 0) {
    section.innerHTML = `<div class="state-msg state-info">${
      environments.length === 0
        ? 'No environments configured.<br/>Add one in Settings first.'
        : 'No environments are enabled for Metadata Browser.<br/>Configure in Settings → Edit environment.'
    }</div>`;
    return;
  }

  // Build a small launcher: environment selector + Open button.
  section.innerHTML = `
    <p class="env-list-label" style="margin-bottom:8px">Browse metadata in:</p>
    <div id="meta-env-dropdown-slot"></div>
    <button id="btn-open-metadata" class="goto-btn" style="--env-color:var(--accent);width:100%;justify-content:center;gap:8px;">
      <span style="font-weight:700">Open Metadata Browser</span>
      <svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" style="vertical-align:middle;flex-shrink:0"><rect x="0.5" y="1" width="15" height="13" rx="2" stroke-opacity="0.3" stroke-width="1"/><rect x="3" y="3.5" width="10" height="8" rx="1.3"/><line x1="3" y1="6" x2="13" y2="6"/></svg>
    </button>
    <p style="font-size:11px;color:var(--text-muted);margin-top:8px;text-align:center;line-height:1.4;">
      Opens as an overlay in this tab — or in a new tab if not on a Dynamics 365 page.
    </p>
  `;

  const metaDropdown = buildEnvDropdown(metaEnvs, defaultEnv);
  document.getElementById('meta-env-dropdown-slot').replaceWith(metaDropdown);

  document.getElementById('btn-open-metadata').addEventListener('click', async () => {
    const { url, name } = metaDropdown.getValue();
    const pageUrl = chrome.runtime.getURL('metadata/metadata.html')
                  + `?env=${encodeURIComponent(url)}&name=${encodeURIComponent(name)}`;
    await _injectTool('metadata', 'Metadata Browser', pageUrl);
    window.close();
  });
}

// ─── Ribbon Buttons Feature ───────────────────────────────────────────────────

let ribbonInitialised = false;

function initRibbon() {
  if (ribbonInitialised) return;
  ribbonInitialised = true;

  const section = document.getElementById('section-ribbon');

  const ribbonEnvs   = environments.filter(env => isEnabledFor(env, 'ribbon'));
  const defaultEnv   = (ribbonEnvs.find(e => e.id === currentEnv?.id)) ?? ribbonEnvs[0] ?? null;

  if (ribbonEnvs.length === 0) {
    section.innerHTML = `<div class="state-msg state-info">${
      environments.length === 0
        ? 'No environments configured.<br/>Add one in Settings first.'
        : 'No environments are enabled for Ribbon Buttons.<br/>Configure in Settings → Edit environment.'
    }</div>`;
    return;
  }

  section.innerHTML = `
    <p class="env-list-label" style="margin-bottom:8px">Browse ribbon buttons in:</p>
    <div id="ribbon-env-dropdown-slot"></div>
    <button id="btn-open-ribbon" class="goto-btn" style="--env-color:var(--accent);width:100%;justify-content:center;gap:8px;">
      <span style="font-weight:700">Open Ribbon Buttons</span>
      <svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" style="vertical-align:middle;flex-shrink:0"><rect x="0.5" y="1" width="15" height="13" rx="2" stroke-opacity="0.3" stroke-width="1"/><rect x="3" y="3.5" width="10" height="8" rx="1.3"/><line x1="3" y1="6" x2="13" y2="6"/></svg>
    </button>
    <p style="font-size:11px;color:var(--text-muted);margin-top:8px;text-align:center;line-height:1.4;">
      Opens as an overlay in this tab — or in a new tab if not on a Dynamics 365 page.
    </p>
  `;

  const ribbonDropdown = buildEnvDropdown(ribbonEnvs, defaultEnv);
  document.getElementById('ribbon-env-dropdown-slot').replaceWith(ribbonDropdown);

  document.getElementById('btn-open-ribbon').addEventListener('click', async () => {
    const { url, name } = ribbonDropdown.getValue();
    const pageUrl = chrome.runtime.getURL('ribbon/ribbon.html')
                  + `?env=${encodeURIComponent(url)}&name=${encodeURIComponent(name)}`;
    await _injectTool('ribbon', 'Ribbon Buttons', pageUrl);
    window.close();
  });
}

// ─── Custom environment dropdown ─────────────────────────────────────────────

/**
 * Build a fully custom styled environment picker.
 * Returns the wrapper element, which exposes a `.getValue()` method returning
 * { url, name } of the currently selected environment.
 *
 * @param {Array}  envs        Filtered list of environments to show.
 * @param {Object} defaultEnv  Pre-selected environment (or null).
 */
function buildEnvDropdown(envs, defaultEnv) {
  let selected = defaultEnv ?? envs[0];

  const wrapper  = document.createElement('div');
  wrapper.className = 'env-dropdown';

  const trigger  = document.createElement('button');
  trigger.type   = 'button';
  trigger.className = 'env-dropdown__trigger';

  const list = document.createElement('div');
  list.className = 'env-dropdown__list hidden';

  function _updateTrigger() {
    const color = selected?.color ?? '#1B3A6B';
    trigger.style.setProperty('--env-color', color);
    trigger.innerHTML = `
      <span class="env-dropdown__swatch" style="background:${escHtml(color)}"></span>
      <span class="env-dropdown__selected-name">${escHtml(selected?.name ?? '')}</span>
      <span class="env-dropdown__chevron">&#9660;</span>
    `;
  }

  function _buildList() {
    list.innerHTML = '';
    envs.forEach(env => {
      const opt = document.createElement('div');
      const isActive = env.id === selected?.id;
      opt.className = `env-dropdown__option${isActive ? ' env-dropdown__option--active' : ''}`;
      if (isActive) opt.style.setProperty('--env-color', env.color ?? '#1B3A6B');
      opt.innerHTML = `
        <span class="env-dropdown__swatch" style="background:${escHtml(env.color ?? '#1B3A6B')}"></span>
        <div class="env-dropdown__opt-info">
          <div class="env-dropdown__opt-name">${escHtml(env.name)}</div>
          <div class="env-dropdown__opt-url">${escHtml(env.url.replace('https://', ''))}</div>
        </div>
        ${isActive ? '<span class="env-dropdown__check">&#10003;</span>' : ''}
      `;
      opt.addEventListener('click', () => {
        selected = env;
        _updateTrigger();
        _buildList();
        _close();
      });
      list.appendChild(opt);
    });
  }

  // Invisible spacer: appended to <body> on open to force the popup window
  // to grow tall enough for the flyout to fit inside its viewport.
  const spacer = document.createElement('div');
  spacer.setAttribute('aria-hidden', 'true');
  spacer.style.cssText = 'pointer-events:none;flex-shrink:0;';

  function _open() {
    const rect        = trigger.getBoundingClientRect();
    // Each option is ~60px tall; cap at the max-height of the list (220px).
    const listH       = Math.min(220, envs.length * 62) + 10;
    spacer.style.height = `${listH}px`;

    // Append spacer first so Chrome expands the popup window before we
    // measure / position the fixed list.
    document.body.appendChild(spacer);

    list.style.top   = `${rect.bottom + 5}px`;
    list.style.left  = `${rect.left}px`;
    list.style.width = `${rect.width}px`;
    document.body.appendChild(list);
    list.classList.remove('hidden');
    trigger.classList.add('open');
  }

  function _close() {
    trigger.classList.remove('open');
    list.classList.add('hidden');
    if (list.parentNode)   list.parentNode.removeChild(list);
    if (spacer.parentNode) spacer.parentNode.removeChild(spacer);
  }

  trigger.addEventListener('click', e => {
    e.stopPropagation();
    list.parentNode ? _close() : _open();
  });

  // Close on any outside click.
  document.addEventListener('click', e => {
    if (list.parentNode && !wrapper.contains(e.target) && !list.contains(e.target)) {
      _close();
    }
  });

  _updateTrigger();
  _buildList();
  wrapper.appendChild(trigger);
  // list is NOT appended here — it's portalled to <body> in _open() / removed in _close()

  /** Returns the currently selected { url, name }. */
  wrapper.getValue = () => ({ url: selected?.url ?? '', name: selected?.name ?? '' });

  return wrapper;
}

// ─── Plugin Trace Logs Feature ───────────────────────────────────────────────

let pluginTraceInitialised = false;

function initPluginTrace() {
  if (pluginTraceInitialised) return;
  pluginTraceInitialised = true;

  const section = document.getElementById('section-plugin-trace');

  const traceEnvs  = environments.filter(env => isEnabledFor(env, 'plugin-trace'));
  const defaultEnv = (traceEnvs.find(e => e.id === currentEnv?.id)) ?? traceEnvs[0] ?? null;

  if (traceEnvs.length === 0) {
    section.innerHTML = `<div class="state-msg state-info">${
      environments.length === 0
        ? 'No environments configured.<br/>Add one in Settings first.'
        : 'No environments are enabled for Plugin Trace Logs.<br/>Configure in Settings → Edit environment.'
    }</div>`;
    return;
  }

  section.innerHTML = `
    <p class="env-list-label" style="margin-bottom:8px">Browse trace logs in:</p>
    <div id="trace-env-dropdown-slot"></div>
    <button id="btn-open-trace" class="goto-btn" style="--env-color:var(--accent);width:100%;justify-content:center;gap:8px;">
      <span style="font-weight:700">Open Plugin Trace Logs</span>
      <svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" style="vertical-align:middle;flex-shrink:0"><rect x="0.5" y="1" width="15" height="13" rx="2" stroke-opacity="0.3" stroke-width="1"/><rect x="3" y="3.5" width="10" height="8" rx="1.3"/><line x1="3" y1="6" x2="13" y2="6"/></svg>
    </button>
    <p style="font-size:11px;color:var(--text-muted);margin-top:8px;text-align:center;line-height:1.4;">
      Opens as an overlay in this tab — or in a new tab if not on a Dynamics 365 page.
    </p>
  `;

  const traceDropdown = buildEnvDropdown(traceEnvs, defaultEnv);
  document.getElementById('trace-env-dropdown-slot').replaceWith(traceDropdown);

  document.getElementById('btn-open-trace').addEventListener('click', async () => {
    const { url, name } = traceDropdown.getValue();
    const pageUrl = chrome.runtime.getURL('plugin-trace/plugin-trace.html')
                  + `?env=${encodeURIComponent(url)}&name=${encodeURIComponent(name)}`;
    await _injectTool('plugin-trace', 'Plugin Trace Logs', pageUrl);
    window.close();
  });
}

// ─── Flows Feature ────────────────────────────────────────────────────────────

let flowsInitialised = false;

function initFlows() {
  if (flowsInitialised) return;
  flowsInitialised = true;

  const section    = document.getElementById('section-flows');
  const flowsEnvs  = environments.filter(env => isEnabledFor(env, 'flows'));
  const defaultEnv = flowsEnvs.find(e => e.id === currentEnv?.id) ?? flowsEnvs[0] ?? null;

  if (flowsEnvs.length === 0) {
    section.innerHTML = `<div class="state-msg state-info">${
      environments.length === 0
        ? 'No environments configured.<br/>Add one in Settings first.'
        : 'No environments are enabled for Flows.<br/>Configure in Settings → Edit environment.'
    }</div>`;
    return;
  }

  section.innerHTML = `
    <p class="env-list-label" style="margin-bottom:8px">Browse flows in:</p>
    <div id="flows-env-dropdown-slot"></div>
    <button id="btn-open-flows" class="goto-btn" style="--env-color:var(--accent);width:100%;justify-content:center;gap:8px;">
      <span style="font-weight:700">Open Flows Viewer</span>
      <svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" style="vertical-align:middle;flex-shrink:0"><rect x="0.5" y="1" width="15" height="13" rx="2" stroke-opacity="0.3" stroke-width="1"/><rect x="3" y="3.5" width="10" height="8" rx="1.3"/><line x1="3" y1="6" x2="13" y2="6"/></svg>
    </button>
    <p style="font-size:11px;color:var(--text-muted);margin-top:8px;text-align:center;line-height:1.4;">
      Opens as an overlay in this tab — or in a new tab if not on a Dynamics 365 page.
    </p>
  `;

  const flowsDropdown = buildEnvDropdown(flowsEnvs, defaultEnv);
  document.getElementById('flows-env-dropdown-slot').replaceWith(flowsDropdown);

  document.getElementById('btn-open-flows').addEventListener('click', async () => {
    const { url, name } = flowsDropdown.getValue();
    const env         = flowsEnvs.find(e => e.url === url);
    const paEnvId     = env?.powerAppsId ?? '';
    const pageUrl = chrome.runtime.getURL('flows/flows.html')
                  + `?env=${encodeURIComponent(url)}&name=${encodeURIComponent(name)}`
                  + (paEnvId ? `&paEnvId=${encodeURIComponent(paEnvId)}` : '');
    await _injectTool('flows', 'Flows Viewer', pageUrl);
    window.close();
  });
}

// ─── Record Details Feature ───────────────────────────────────────────────────

// ─── Data Sync ────────────────────────────────────────────────────────────────

let dataSyncInitialised = false;

function initDataSync() {
  if (dataSyncInitialised) return;
  dataSyncInitialised = true;

  const section = document.getElementById('section-data-sync');
  section.innerHTML = `
    <button id="btn-open-data-sync" class="goto-btn" style="--env-color:var(--accent);width:100%;justify-content:center;gap:8px;">
      <span style="font-weight:700">Open Data Sync</span>
      <svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" style="vertical-align:middle;flex-shrink:0"><rect x="0.5" y="1" width="15" height="13" rx="2" stroke-opacity="0.3" stroke-width="1"/><rect x="3" y="3.5" width="10" height="8" rx="1.3"/><line x1="3" y1="6" x2="13" y2="6"/></svg>
    </button>
    <p style="font-size:11px;color:var(--text-muted);margin-top:10px;text-align:center;line-height:1.5;">
      Synchronise records between two environments.<br/>
      Opens as an overlay in this tab — or in a new tab if not on a Dynamics 365 page.
    </p>`;

  document.getElementById('btn-open-data-sync').addEventListener('click', async () => {
    let pageUrl = chrome.runtime.getURL('data-sync/data-sync.html');
    if (currentEnv?.url) pageUrl += `?sourceEnv=${encodeURIComponent(currentEnv.url)}`;
    await _injectTool('data-sync', 'Data Sync', pageUrl);
    window.close();
  });
}

let recordDetailsInitialised = false;

function initRecordDetails() {
  if (recordDetailsInitialised) return;
  recordDetailsInitialised = true;

  const section     = document.getElementById('section-record-details');
  const pageContext = parseD365PageContext(currentTab?.url);

  if (!currentEnv || !pageContext?.etn || !pageContext?.id) {
    section.innerHTML = `<div class="state-msg state-info">Not on a Dynamics 365 record form.</div>`;
    return;
  }

  section.innerHTML = `
    <div class="state-msg state-loading">
      <span class="spinner"></span> Loading record details&hellip;
    </div>`;

  fetchRecordDetails(currentEnv.url, pageContext.etn, pageContext.id)
    .then(data  => _renderRecordDetails(section, data, pageContext))
    .catch(err  => {
      section.innerHTML = `<div class="state-msg state-error">Failed to load: ${escHtml(err.message)}</div>`;
    });
}

async function fetchRecordDetails(baseUrl, etn, id) {
  const entitySetName = await _fetchEntitySetName(baseUrl, etn);
  if (!entitySetName) throw new Error('Could not resolve entity type.');

  const apiVersion = settings.apiVersion || 'v9.2';
  const apiUrl = (
    `${baseUrl}/api/data/${apiVersion}/${entitySetName}(${id})` +
    `?$select=createdon,modifiedon,_ownerid_value` +
    `&$expand=createdby($select=fullname,systemuserid),modifiedby($select=fullname,systemuserid)`
  );

  const [injection] = await chrome.scripting.executeScript({
    target: { tabId: currentTab.id },
    func: async (url) => {
      try {
        const res = await fetch(url, {
          headers: {
            'Accept':           'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version':    '4.0',
            'Prefer':           'odata.include-annotations="OData.Community.Display.V1.FormattedValue,Microsoft.Dynamics.CRM.lookuplogicalname"',
          },
        });
        if (!res.ok) return { error: `API ${res.status}: ${res.statusText}` };
        return await res.json();
      } catch (e) {
        return { error: e.message };
      }
    },
    args: [apiUrl],
  });

  const result = injection?.result;
  if (result?.error) throw new Error(result.error);
  return { ...result, _entitySetName: entitySetName };
}

/** Shows a full-screen modal overlay with the complete clone error text. */
function _showCloneErrorModal(errorText) {
  document.getElementById('clone-err-overlay')?.remove();
  const overlay = document.createElement('div');
  overlay.id    = 'clone-err-overlay';
  overlay.className = 'clone-err-overlay';
  overlay.innerHTML = `
    <div class="clone-err-modal">
      <div class="clone-err-modal__title">Clone Error Details</div>
      <pre class="clone-err-modal__body">${escHtml(errorText)}</pre>
      <button class="clone-err-modal__close">Close</button>
    </div>`;
  overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
  overlay.querySelector('.clone-err-modal__close').addEventListener('click', () => overlay.remove());
  document.body.appendChild(overlay);
}

function _renderRecordDetails(section, data, pageContext) {
  if (!data) {
    section.innerHTML = `<div class="state-msg state-error">No data returned from API.</div>`;
    return;
  }

  const createdBy  = data.createdby;
  const modifiedBy = data.modifiedby;

  // ── Layout ────────────────────────────────────────────────────────────────
  const layout = document.createElement('div');
  layout.className = 'rd-layout';

  const main = document.createElement('div');
  main.className = 'rd-main';

  // ── Clone Record bar (single row: Clone button + environment dropdown) ─────
  const cloneAllowed = !settings.cloneWhitelist ||
    settings.cloneWhitelist.includes((pageContext.etn ?? '').toLowerCase());

  if (cloneAllowed) {
    const cloneBar = document.createElement('div');
    cloneBar.className = 'rd-clone-bar';

    const cloneSVG = `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6"
        stroke-linecap="round" stroke-linejoin="round" width="13" height="13">
      <rect x="5" y="4" width="9" height="11" rx="1.5"/>
      <path d="M2 11V2.5A1.5 1.5 0 0 1 3.5 1H11"/>
    </svg>`;

    /** Reusable error UI: compact error bar + "View details" link. */
    function _showCloneError(container, err, onRetry) {
      container.innerHTML = '';
      const errEl = document.createElement('div');
      errEl.className = 'rd-clone-error';
      const errMsg = document.createElement('span');
      errMsg.style.cssText = 'flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;';
      errMsg.textContent = 'Error occurred.';
      const detailsBtn = document.createElement('button');
      detailsBtn.className = 'rd-clone-error-details';
      detailsBtn.textContent = 'View details';
      detailsBtn.addEventListener('click', () => _showCloneErrorModal(err.message));
      errEl.appendChild(errMsg);
      errEl.appendChild(detailsBtn);
      const retryBtn = document.createElement('button');
      retryBtn.className = 'rd-clone-btn';
      retryBtn.style.marginTop = '3px';
      retryBtn.innerHTML = `${cloneSVG} Retry`;
      retryBtn.addEventListener('click', onRetry);
      container.appendChild(errEl);
      container.appendChild(retryBtn);
    }

    // ── Environment dropdown (all envs — current env first) ───────────────────
    const allCloneEnvs = [currentEnv, ...environments.filter(e => e.url !== currentEnv.url)];
    const cloneDropdown = buildEnvDropdown(allCloneEnvs, currentEnv);

    // ── Clone button ──────────────────────────────────────────────────────────
    const cloneBtn = document.createElement('button');
    cloneBtn.className = 'rd-clone-btn';
    cloneBtn.innerHTML = `${cloneSVG} Clone`;

    const cloneRow = document.createElement('div');
    cloneRow.className = 'rd-clone-row';
    cloneRow.appendChild(cloneBtn);
    cloneRow.appendChild(cloneDropdown);
    cloneBar.appendChild(cloneRow);

    const doClone = async () => {
      if (cloneBtn.disabled) return;
      const { url: selectedUrl } = cloneDropdown.getValue();
      const isCross = selectedUrl !== currentEnv.url;

      cloneBtn.disabled = true;
      cloneBtn.innerHTML = `<span class="spinner" style="width:10px;height:10px;border-width:1.5px;flex-shrink:0"></span> ${isCross ? 'Copying\u2026' : 'Cloning\u2026'}`;

      try {
        const { newId } = await cloneRecord(
          currentEnv.url, pageContext.etn, pageContext.id, data._entitySetName,
          isCross ? selectedUrl : null
        );
        const baseUrl = isCross ? selectedUrl : currentEnv.url;
        const newRecordUrl = `${baseUrl}/main.aspx?pagetype=entityrecord&etn=${pageContext.etn}&id=${newId}`;
        cloneBar.innerHTML = `
          <div class="rd-clone-success">
            <svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="2"
                stroke-linecap="round" stroke-linejoin="round" width="12" height="12">
              <path d="M2 7l3.5 3.5L12 3"/>
            </svg>
            ${isCross ? 'Copied!' : 'Cloned!'}&nbsp;
            <a href="${newRecordUrl}" target="_blank" rel="noopener noreferrer" class="rd-clone-open-link">Open record ↗</a>
          </div>`;
      } catch (err) {
        _showCloneError(cloneBar, err, () => {
          cloneBar.innerHTML = '';
          cloneBtn.disabled = false;
          cloneBtn.innerHTML = `${cloneSVG} Clone`;
          cloneBar.appendChild(cloneRow);
        });
      }
    };

    cloneBtn.addEventListener('click', doClone);
    main.appendChild(cloneBar);
  }

  // ── Block builder ─────────────────────────────────────────────────────────
  /** Builds an rd-block with one or more copyable value rows. */
  function _block(label, ...rows) {
    const block = document.createElement('div');
    block.className = 'rd-block';

    const lbl = document.createElement('span');
    lbl.className = 'rd-label';
    lbl.textContent = label;
    block.appendChild(lbl);

    rows.forEach(({ text, mono, link, icon }) => {
      if (!text) return;
      const row = document.createElement('div');
      row.className = 'rd-value-row';

      if (icon) {
        const iconEl = document.createElement('span');
        iconEl.className = 'rd-owner-icon';
        iconEl.innerHTML = icon;
        row.appendChild(iconEl);
      }

      let valEl;
      if (link) {
        valEl = document.createElement('a');
        valEl.className = mono ? 'rd-user-id rd-link' : 'rd-value rd-link';
        valEl.href = link;
        valEl.target = '_blank';
        valEl.rel = 'noopener noreferrer';
        valEl.textContent = text;
      } else {
        valEl = document.createElement('span');
        valEl.className = mono ? 'rd-user-id' : 'rd-value';
        valEl.textContent = text;
      }

      const btn = document.createElement('button');
      btn.className = 'rd-copy-btn';
      btn.title = 'Copy';
      btn.innerHTML = `<svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3.5" y="2" width="8" height="10" rx="1"/><path d="M2 4.5H1.5a.5.5 0 0 0-.5.5v7a.5.5 0 0 0 .5.5h7a.5.5 0 0 0 .5-.5V12"/></svg>`;
      btn.addEventListener('click', () => {
        navigator.clipboard.writeText(text).then(() => {
          btn.classList.add('copied');
          setTimeout(() => btn.classList.remove('copied'), 1500);
        });
      });

      row.appendChild(valEl);
      row.appendChild(btn);
      block.appendChild(row);
    });

    return block;
  }

  // ── Blocks ────────────────────────────────────────────────────────────────

  // Record ID — linked to Web API record endpoint
  if (pageContext?.id) {
    const apiLink = data._entitySetName && currentEnv
      ? `${currentEnv.url}/api/data/${settings.apiVersion || 'v9.2'}/${data._entitySetName}(${pageContext.id})`
      : null;
    main.appendChild(_block('Record ID', { text: pageContext.id, mono: true, link: apiLink }));
  }

  // Entity type name
  if (pageContext?.etn) {
    main.appendChild(_block('Entity Type', { text: pageContext.etn, mono: false }));
  }

  /** Builds a D365 record URL for a given entity type + GUID. */
  function _d365RecordUrl(etn, id) {
    if (!currentEnv?.url || !id) return null;
    return `${currentEnv.url}/main.aspx?pagetype=entityrecord&etn=${etn}&id=${id}`;
  }

  // Created On / By
  main.appendChild(_block('Created On', { text: _formatDateTime(data.createdon), mono: false }));
  main.appendChild(_block('Created By',
    { text: createdBy?.fullname     ?? '—',  mono: false,
      link: _d365RecordUrl('systemuser', createdBy?.systemuserid) },
    { text: createdBy?.systemuserid ?? null,  mono: true  },
  ));

  // Modified On / By
  main.appendChild(_block('Modified On', { text: _formatDateTime(data.modifiedon), mono: false }));
  main.appendChild(_block('Modified By',
    { text: modifiedBy?.fullname     ?? '—', mono: false,
      link: _d365RecordUrl('systemuser', modifiedBy?.systemuserid) },
    { text: modifiedBy?.systemuserid ?? null, mono: true  },
  ));

  // Owner (user or team)
  const ownerName = data['_ownerid_value@OData.Community.Display.V1.FormattedValue'] ?? null;
  const ownerId   = data['_ownerid_value'] ?? null;
  const ownerType = data['_ownerid_value@Microsoft.Dynamics.CRM.lookuplogicalname'] ?? 'systemuser';
  if (ownerName || ownerId) {
    const isTeam   = ownerType === 'team';
    const userIcon = `<svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="7" cy="4.5" r="2.2"/><path d="M2 12.5c0-2.76 2.24-4.5 5-4.5s5 1.74 5 4.5"/></svg>`;
    const teamIcon = `<svg viewBox="0 0 18 13" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="4" cy="4" r="2"/><path d="M1 12c0-1.66 1.34-3 3-3s3 1.34 3 3"/><circle cx="9" cy="3.5" r="2.2"/><path d="M5.5 12c0-1.93 1.57-3.5 3.5-3.5s3.5 1.57 3.5 3.5"/><circle cx="14" cy="4" r="2"/><path d="M11 12c0-1.66 1.34-3 3-3s3 1.34 3 3"/></svg>`;
    main.appendChild(_block('Owner',
      { text: ownerName ?? '—', mono: false, icon: isTeam ? teamIcon : userIcon,
        link: _d365RecordUrl(isTeam ? 'team' : 'systemuser', ownerId) },
      { text: ownerId ?? null, mono: true },
    ));
  }

  // Open in Metadata Browser
  if (currentEnv && pageContext?.etn) {
    const actionBlock = document.createElement('div');
    actionBlock.className = 'rd-block rd-block--action';

    const metaBtn = document.createElement('button');
    metaBtn.className = 'rd-meta-btn';
    metaBtn.innerHTML = `
      <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">
        <ellipse cx="8" cy="4" rx="5" ry="1.8"/>
        <path d="M3 4v3.5c0 1 2.2 1.8 5 1.8s5-.8 5-1.8V4"/>
        <path d="M3 7.5V11c0 1 2.2 1.8 5 1.8s5-.8 5-1.8V7.5"/>
      </svg>
      Open in Metadata Browser`;
    metaBtn.addEventListener('click', async () => {
      const url = chrome.runtime.getURL('metadata/metadata.html')
        + `?env=${encodeURIComponent(currentEnv.url)}`
        + `&name=${encodeURIComponent(currentEnv.name)}`
        + `&etn=${encodeURIComponent(pageContext.etn)}`;
      await _openTab(url);
    });

    actionBlock.appendChild(metaBtn);
    main.appendChild(actionBlock);
  }

  // ── Open In pane ──────────────────────────────────────────────────────────
  const pane = document.createElement('div');
  pane.className = 'rd-open-in-pane';

  // Restore saved open/closed state.
  chrome.storage.local.get('rdPaneOpen', ({ rdPaneOpen }) => {
    if (rdPaneOpen) {
      pane.classList.add('open');
      document.documentElement.classList.add('rd-pane-open');
      document.body.classList.add('rd-pane-open');
    }
  });

  const paneToggle = document.createElement('button');
  paneToggle.className = 'rd-pane-toggle';
  paneToggle.title = 'Open In\u2026';
  paneToggle.innerHTML = `<svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 2L4 7l5 5"/></svg>`;
  paneToggle.addEventListener('click', () => {
    const open = pane.classList.toggle('open');
    document.documentElement.classList.toggle('rd-pane-open', open);
    document.body.classList.toggle('rd-pane-open', open);
    chrome.storage.local.set({ rdPaneOpen: open });
  });

  const paneContent = document.createElement('div');
  paneContent.className = 'rd-pane-content';

  const paneTitle = document.createElement('div');
  paneTitle.className = 'rd-pane-title';
  paneTitle.textContent = 'Open In\u2026';
  paneContent.appendChild(paneTitle);

  const paneTargets = environments.filter(env => env.id !== currentEnv?.id && isEnabledFor(env, 'goto-open-in'));
  if (paneTargets.length === 0) {
    const msg = document.createElement('div');
    msg.className = 'rd-pane-empty';
    msg.textContent = 'No other environments configured.';
    paneContent.appendChild(msg);
  } else {
    paneTargets.forEach(env => {
      const btn = document.createElement('button');
      btn.className = `rd-env-btn${env.warn ? ' rd-env-btn--warn' : ''}`;
      btn.title = env.name;
      btn.style.setProperty('--env-color', env.color ?? '#1B3A6B');
      btn.innerHTML = `
        <span class="rd-env-btn__name">${escHtml(env.name)}</span>
        ${env.warn ? '<span class="rd-env-btn__warn">Prod</span>' : ''}`;
      btn.addEventListener('click', () => handleGoTo(env));
      paneContent.appendChild(btn);
    });
  }

  pane.appendChild(paneToggle);
  pane.appendChild(paneContent);

  layout.appendChild(main);
  layout.appendChild(pane);

  section.innerHTML = '';
  section.appendChild(layout);
}

/** Formats an ISO 8601 timestamp as DD/MM/YYYY HH:mm (local time). */
function _formatDateTime(iso) {
  if (!iso) return '—';
  try {
    const d   = new Date(iso);
    const dd  = String(d.getDate()).padStart(2, '0');
    const mm  = String(d.getMonth() + 1).padStart(2, '0');
    const hh  = String(d.getHours()).padStart(2, '0');
    const min = String(d.getMinutes()).padStart(2, '0');
    return `${dd}/${mm}/${d.getFullYear()} ${hh}:${min}`;
  } catch {
    return iso;
  }
}

// ─── Feature filtering ────────────────────────────────────────────────────────

/** Returns true if `env` is enabled for the given feature.
 *  Absent / empty enabledFor means all features are on (backward-compatible). */
function isEnabledFor(env, featureId) {
  if (!Array.isArray(env.enabledFor) || env.enabledFor.length === 0) return true;
  return env.enabledFor.includes(featureId);
}

// ─── Status bar ───────────────────────────────────────────────────────────────

function showStatus(message, type = 'loading') {
  const bar = document.getElementById('status-bar');
  const icon = type === 'loading' ? '<span class="spinner"></span>'
             : type === 'success' ? '&#10003;'
             : '&#9888;';
  bar.innerHTML = `${icon} ${escHtml(message)}`;
  bar.className = `status-bar status-bar--${type}`;
  bar.classList.remove('hidden');
}

// ─── Utility ──────────────────────────────────────────────────────────────────

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
