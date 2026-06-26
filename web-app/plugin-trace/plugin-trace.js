/**
 * EF Power Platform Tools — Plugin Trace Logs (web-app)
 *
 * Full-page viewer for Dynamics 365 Plug-in Trace Log records.
 * Runs as an iframe opened by launcher.js (which runs in the D365 page context).
 * All D365 API calls are relayed via postMessage bridge — no direct fetch.
 */

// ─── URL params ────────────────────────────────────────────────────────────────

const _params    = new URLSearchParams(location.search);
const ENV_URL    = (_params.get('env')   ?? '').replace(/\/$/, '');
const ENV_NAME   = _params.get('name')  ?? ENV_URL;

// ─── State ────────────────────────────────────────────────────────────────────

let _allLogs       = [];          // full page of results from the API
let _filteredLogs  = [];          // after client-side text search
let _openLogId     = null;        // currently expanded row
let _entitySetName = null;        // resolved via EntityDefinitions on first fetch
let _showFilter    = 'all'; // 'all' | 'exceptions'
let _ignoredTypes  = new Set();   // short plugin type names excluded from results
let _fetchingMore  = false;       // true while paginating (more pages still in flight)

// ─── Ignored types ────────────────────────────────────────────────────────────

function _loadIgnoredTypes() {
  const raw = localStorage.getItem('ef_ppt_pluginTraceIgnored');
  _ignoredTypes = new Set(raw ? JSON.parse(raw) : []);
}

function _saveIgnoredTypes() {
  localStorage.setItem('ef_ppt_pluginTraceIgnored', JSON.stringify([..._ignoredTypes]));
  return Promise.resolve();
}

async function _ignoreType(shortName) {
  _ignoredTypes.add(shortName);
  await _saveIgnoredTypes();
  _applyClientFilter();
}

// ─── Bootstrap ────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {
  document.getElementById('env-name').textContent = ENV_NAME;
  document.title = `Plugin Trace Logs — ${ENV_NAME}`;

  _loadIgnoredTypes();

  // Wire controls.
  document.getElementById('btn-refresh').addEventListener('click', () => _fetchLogs());
  document.getElementById('btn-apply').addEventListener('click', () => _fetchLogs());
  document.getElementById('btn-clear').addEventListener('click', () => {
    document.getElementById('filter-from').value = '';
    document.getElementById('filter-to').value   = '';
    document.getElementById('search').value       = '';
    _fetchLogs();
  });

  // Live client-side search as the user types.
  document.getElementById('search').addEventListener('input', _applyClientFilter);

  // All / Exceptions segmented control.
  document.querySelectorAll('.pt-seg__btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.pt-seg__btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      _showFilter = btn.dataset.filter;
      _applyClientFilter();
    });
  });

  // Enter key in date fields triggers apply.
  ['filter-from', 'filter-to'].forEach(id => {
    document.getElementById(id).addEventListener('keydown', e => {
      if (e.key === 'Enter') _fetchLogs();
    });
  });

  if (!ENV_URL) {
    _showError('No environment URL provided. Please open this page from the EF Power Platform Tools launcher.');
    return;
  }

  _showLoading('Fetching plug-in trace logs…');
  try {
    await _fetchLogs();
  } catch (e) {
    console.error('[EF PPT]', e);
    _showError(e.message);
  }
});

// ─── Fetch bridge ─────────────────────────────────────────────────────────────

const _inModal = true; // always iframe in web-app

function _bridgeFetch(url, extraHeaders = {}, method = 'GET', body = null) {
  return new Promise((resolve, reject) => {
    const id    = Math.random().toString(36).slice(2) + Date.now();
    const timer = setTimeout(() => {
      window.removeEventListener('message', onMsg);
      reject(new Error('Bridge fetch timeout'));
    }, 30_000);
    function onMsg(e) {
      if (!e.data || e.data.__efppt !== 'fetch-result' || e.data.id !== id) return;
      clearTimeout(timer);
      window.removeEventListener('message', onMsg);
      if (e.data.ok) resolve(e.data.data);
      else reject(new Error(e.data.error || 'HTTP ' + e.data.status));
    }
    window.addEventListener('message', onMsg);
    window.parent.postMessage({ __efppt: 'fetch', id, url, method, headers: extraHeaders, body }, '*');
  });
}

async function _d365Fetch(url, extraHeaders = {}) {
  return _bridgeFetch(url, extraHeaders);
}

// ─── Entity set name resolver ─────────────────────────────────────────────────

/**
 * Resolves the OData entity set name for plugintracelog.
 * Cached for the lifetime of the page.
 */
async function _resolveEntitySetName() {
  if (_entitySetName) return _entitySetName;

  const url = (
    `${ENV_URL}/api/data/v9.2/EntityDefinitions(LogicalName='plugintracelog')` +
    `?$select=EntitySetName`
  );
  const data = await _d365Fetch(url);

  if (!data.EntitySetName) throw new Error('plugintracelog EntitySetName not returned by metadata API.');
  _entitySetName = data.EntitySetName;
  return _entitySetName;
}

// ─── Fetch (with full pagination) ────────────────────────────────────────────

async function _fetchLogs() {
  _showLoading(_entitySetName ? 'Fetching plug-in trace logs…' : 'Resolving entity metadata…');

  try {
    const entitySet = await _resolveEntitySetName();

    _showLoading('Fetching plug-in trace logs…');

    const rawFrom = document.getElementById('filter-from').value;
    const rawTo   = document.getElementById('filter-to').value;
    const filters = [];
    if (rawFrom) filters.push(`createdon ge ${new Date(rawFrom).toISOString()}`);
    if (rawTo)   filters.push(`createdon le ${new Date(rawTo).toISOString()}`);

    const filterStr = filters.length
      ? `&$filter=${encodeURIComponent(filters.join(' and '))}`
      : '';

    // No $top — follow @odata.nextLink until all pages are retrieved.
    const firstUrl = (
      `${ENV_URL}/api/data/v9.2/${entitySet}` +
      `?$select=plugintracelogid,typename,messagename,createdon,mode,depth` +
      `,exceptiondetails,messageblock,performanceexecutionduration,correlationid,requestid` +
      `&$orderby=createdon%20desc` +
      filterStr
    );

    // Reset state; pages are streamed progressively into _allLogs.
    _allLogs      = [];
    _openLogId    = null;
    _fetchingMore = true;

    await _fetchAllPages(firstUrl, (pageRecords) => {
      _allLogs.push(...pageRecords);
      _applyClientFilter();  // render after each page so user sees results immediately
    });
  } catch (e) {
    _fetchingMore = false;
    _showError(e.message);
    return;
  }
  _fetchingMore = false;
  _applyClientFilter();  // final render to clear "loading more…" indicator
}

/**
 * Follows @odata.nextLink pagination, calling onPage(records) after each page
 * so the caller can render results progressively.
 * Uses odata.maxpagesize=500 for a responsive first-paint.
 *
 * @param {string}   firstUrl - Initial OData request URL.
 * @param {Function} onPage   - Called with each page's record array as it arrives.
 */
async function _fetchAllPages(firstUrl, onPage) {
  let nextUrl = firstUrl;

  while (nextUrl) {
    const data = await _d365Fetch(nextUrl, { 'Prefer': 'odata.maxpagesize=500' });
    onPage(data.value ?? []);
    nextUrl = data['@odata.nextLink'] ?? null;
  }
}

// ─── Client-side search filter ────────────────────────────────────────────────

function _applyClientFilter() {
  const q = document.getElementById('search').value.trim().toLowerCase();

  _filteredLogs = _allLogs.filter(log => {
    if (_showFilter === 'exceptions' && !log.exceptiondetails) return false;
    if (_ignoredTypes.has(_shortName(log.typename ?? ''))) return false;
    if (!q) return true;
    return (
      (log.typename         ?? '').toLowerCase().includes(q) ||
      (log.messagename      ?? '').toLowerCase().includes(q) ||
      (log.exceptiondetails ?? '').toLowerCase().includes(q)
    );
  });

  _renderLogs();
}

// ─── Rendering ────────────────────────────────────────────────────────────────

function _renderLogs() {
  const container = document.getElementById('log-list');
  const info      = document.getElementById('results-info');

  // While the very first page hasn't arrived yet, leave the spinner visible.
  if (_allLogs.length === 0 && _fetchingMore) return;

  if (_filteredLogs.length === 0) {
    info.classList.add('hidden');
    container.innerHTML = `<div class="pt-empty">${
      _allLogs.length === 0
        ? 'No plug-in trace logs found. Adjust the date range or check that trace logging is enabled in this environment.'
        : 'No logs match the current search.'
    }</div>`;
    return;
  }

  const total     = _allLogs.length;
  const shown     = _filteredLogs.length;
  const fetchedAt = _fmtDateTime(new Date().toISOString());
  const suffix    = _fetchingMore
    ? ' — loading more…'
    : ` · fetched ${fetchedAt}`;
  info.textContent = shown === total
    ? `${total} log${total !== 1 ? 's' : ''}${suffix}`
    : `${shown} of ${total} logs (filtered)${suffix}`;
  info.classList.remove('hidden');

  container.innerHTML = '';
  _filteredLogs.forEach(log => container.appendChild(_buildRow(log)));

  // Restore previously expanded row if still in the filtered set.
  if (_openLogId) {
    const row = container.querySelector(`[data-id="${CSS.escape(_openLogId)}"]`);
    if (row) row.classList.add('expanded');
  }
}

function _buildRow(log) {
  const hasException = !!log.exceptiondetails;
  const hasTrace     = !!log.messageblock;
  const isSync       = log.mode === 0;
  const modeLabel    = isSync ? 'Sync' : log.mode === 1 ? 'Async' : `Mode ${log.mode}`;
  const modeClass    = isSync ? 'pt-badge--sync' : 'pt-badge--async';
  const duration     = log.performanceexecutionduration != null
    ? `${log.performanceexecutionduration} ms`
    : '—';

  const row = document.createElement('div');
  row.className = `pt-log-row${hasException ? ' pt-log-row--error' : ''}`;
  row.dataset.id = log.plugintracelogid;

  // ── Summary row ──
  const summary = document.createElement('div');
  summary.className = 'pt-log-summary';
  summary.innerHTML = `
    <span class="pt-log-indicator ${hasException ? 'pt-log-indicator--error' : 'pt-log-indicator--ok'}"></span>
    <div class="pt-log-main">
      <span class="pt-log-type" title="${_esc(log.typename ?? '')}">${_esc(_shortName(log.typename ?? '—'))}</span>
      <span class="pt-log-msg">${_esc(log.messagename ?? '—')}</span>
    </div>
    <div class="pt-log-meta">
      <span class="pt-badge ${modeClass}">${_esc(modeLabel)}</span>
      <span class="pt-log-time">${_esc(_fmtDateTime(log.createdon))}</span>
      <span class="pt-log-dur">${_esc(duration)}</span>
    </div>
    <span class="pt-log-chevron">&#8250;</span>
  `;

  // ── Detail panel ──
  const detail = document.createElement('div');
  detail.className = 'pt-log-detail';

  // Exception block
  if (hasException) {
    detail.appendChild(_detailSection('Exception', log.exceptiondetails, true));
  }

  // Trace block
  if (hasTrace) {
    detail.appendChild(_detailSection('Trace', log.messageblock, false));
  }

  if (!hasException && !hasTrace) {
    const empty = document.createElement('div');
    empty.className = 'pt-detail-empty';
    empty.textContent = 'No exception or trace data for this log entry.';
    detail.appendChild(empty);
  }

  // Metadata grid
  const metaSection = document.createElement('div');
  metaSection.className = 'pt-detail-section';
  const metaTitle = document.createElement('div');
  metaTitle.className = 'pt-detail-section-title';
  metaTitle.textContent = 'Details';
  metaSection.appendChild(metaTitle);

  const grid = document.createElement('div');
  grid.className = 'pt-detail-grid';
  const rows = [
    ['Plugin Type',  log.typename    ?? '—', false],
    ['Message',      log.messagename ?? '—', false],
    ['Mode',         modeLabel,               false],
    ['Depth',        String(log.depth ?? '—'), false],
    ['Duration',     duration,                false],
    ['Created On',   _fmtDateTime(log.createdon), false],
  ];
  if (log.correlationid) rows.push(['Correlation ID', log.correlationid, true]);
  if (log.requestid)     rows.push(['Request ID',     log.requestid,     true]);

  rows.forEach(([key, val, mono]) => {
    const k = document.createElement('span');
    k.className = 'pt-detail-key';
    k.textContent = key;
    const v = document.createElement('span');
    v.className = `pt-detail-val${mono ? ' pt-detail-val--mono' : ''}`;
    v.textContent = val;
    v.title = val;
    grid.appendChild(k);
    grid.appendChild(v);
  });

  metaSection.appendChild(grid);
  detail.appendChild(metaSection);

  // Toggle expand on summary click.
  summary.addEventListener('click', () => {
    const isOpen = row.classList.contains('expanded');
    document.querySelectorAll('.pt-log-row.expanded').forEach(r => r.classList.remove('expanded'));
    if (!isOpen) {
      row.classList.add('expanded');
      _openLogId = log.plugintracelogid;
    } else {
      _openLogId = null;
    }
  });

  // Right-click context menu.
  summary.addEventListener('contextmenu', e => {
    e.preventDefault();
    _showContextMenu(e.clientX, e.clientY, _shortName(log.typename ?? ''));
  });

  row.appendChild(summary);
  row.appendChild(detail);
  return row;
}

// ─── Context menu ─────────────────────────────────────────────────────────────

function _showContextMenu(x, y, shortName) {
  _hideContextMenu();

  const menu = document.createElement('div');
  menu.className = 'pt-ctx-menu';
  menu.id = '_pt-ctx-menu';

  const ignoreSvg = `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
    <circle cx="8" cy="8" r="6.5"/>
    <line x1="4.5" y1="8" x2="11.5" y2="8"/>
  </svg>`;

  const ignoreBtn = document.createElement('button');
  ignoreBtn.className = 'pt-ctx-item pt-ctx-item--danger';
  ignoreBtn.innerHTML = `${ignoreSvg} Ignore "${shortName}"`;
  ignoreBtn.addEventListener('click', e => {
    e.stopPropagation();
    _ignoreType(shortName);
    _hideContextMenu();
  });

  menu.appendChild(ignoreBtn);
  document.body.appendChild(menu);

  // Smart position — flip if menu would overflow viewport
  const mw = menu.offsetWidth  || 240;
  const mh = menu.offsetHeight || 48;
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  menu.style.left = `${x + mw > vw ? Math.max(0, x - mw) : x}px`;
  menu.style.top  = `${y + mh > vh ? Math.max(0, y - mh) : y}px`;

  // Close on any outside click, context menu, or Escape
  const close = () => _hideContextMenu();
  setTimeout(() => {
    document.addEventListener('click',       close, { once: true });
    document.addEventListener('contextmenu', close, { once: true });
    document.addEventListener('keydown', function esc(e) {
      if (e.key === 'Escape') { close(); document.removeEventListener('keydown', esc); }
    });
  }, 0);
}

function _hideContextMenu() {
  document.getElementById('_pt-ctx-menu')?.remove();
}

function _detailSection(title, content, isError) {
  const section = document.createElement('div');
  section.className = 'pt-detail-section';

  // Header row: title + copy button
  const header = document.createElement('div');
  header.className = 'pt-detail-section-header';

  const titleEl = document.createElement('div');
  titleEl.className = 'pt-detail-section-title';
  titleEl.textContent = title;

  const copyBtn = document.createElement('button');
  copyBtn.className = 'pt-copy-btn';
  copyBtn.title = 'Copy to clipboard';
  copyBtn.innerHTML = `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
    <rect x="5" y="4" width="8" height="10" rx="1.2"/>
    <path d="M3 12V3a1 1 0 0 1 1-1h7"/>
  </svg>`;
  copyBtn.addEventListener('click', async (e) => {
    e.stopPropagation();
    try {
      await navigator.clipboard.writeText(content);
      copyBtn.classList.add('copied');
      setTimeout(() => copyBtn.classList.remove('copied'), 1800);
    } catch { /* clipboard not available */ }
  });

  header.appendChild(titleEl);
  header.appendChild(copyBtn);

  const pre = document.createElement('pre');
  pre.className = `pt-detail-pre${isError ? ' pt-detail-pre--error' : ''}`;
  pre.textContent = content;

  section.appendChild(header);
  section.appendChild(pre);
  return section;
}

// ─── State helpers ────────────────────────────────────────────────────────────

function _showLoading(msg) {
  document.getElementById('results-info').classList.add('hidden');
  document.getElementById('log-list').innerHTML = `
    <div class="pt-state">
      <span class="pt-spinner"></span>
      <span>${_esc(msg)}</span>
    </div>`;
}

function _updateLoadingMsg(msg) {
  const el = document.querySelector('#log-list .pt-state span:last-child');
  if (el) el.textContent = msg;
}

function _showError(msg) {
  document.getElementById('results-info').classList.add('hidden');
  document.getElementById('log-list').innerHTML = `
    <div class="pt-state pt-state--error">
      <strong>Could not load trace logs</strong>
      ${_esc(msg)}
    </div>`;
}

// ─── Formatting helpers ───────────────────────────────────────────────────────

/** Extract the plugin class name from a fully-qualified assembly-qualified type name.
 *  e.g. "MyOrg.Plugins.MyPlugin, MyOrg.Plugins, Version=1.0.0.0, …"
 *       → "MyOrg.Plugins.MyPlugin"
 */
function _shortName(fullName) {
  return (fullName.split(',')[0] ?? fullName).trim();
}

/** Format an ISO timestamp as DD/MM/YYYY HH:mm (local time). */
function _fmtDateTime(iso) {
  if (!iso) return '—';
  try {
    const d  = new Date(iso);
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const hh = String(d.getHours()).padStart(2, '0');
    const mi = String(d.getMinutes()).padStart(2, '0');
    return `${dd}/${mm}/${d.getFullYear()} ${hh}:${mi}`;
  } catch { return iso; }
}

/** HTML-escape a string. */
function _esc(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
