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

const PAGE_SIZE = 100;

let _openLogId     = null;        // currently expanded row
let _entitySetName = null;        // resolved via EntityDefinitions on first fetch
let _showFilter    = 'all'; // 'all' | 'exceptions'
let _ignoredTypes  = new Set();   // short plugin type names excluded from results
let _pages         = [];
let _moreAfter     = [];
let _pageIndex     = 0;
let _totalCount    = null;
let _countState    = null;
let _filterXml     = '';
let _queryToken    = 0;
let _searchTimer   = null;
let _loadingPage   = false;
let _fetchedAt     = null;

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
  _fetchLogs();
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

  document.getElementById('search').addEventListener('input', () => {
    clearTimeout(_searchTimer);
    _searchTimer = setTimeout(() => _fetchLogs(), 400);
  });
  document.getElementById('search').addEventListener('keydown', e => {
    if (e.key === 'Enter') { clearTimeout(_searchTimer); _fetchLogs(); }
  });

  // All / Exceptions segmented control.
  document.querySelectorAll('.pt-seg__btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.pt-seg__btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      _showFilter = btn.dataset.filter;
      _fetchLogs();
    });
  });

  // Enter key in date fields triggers apply.
  ['filter-from', 'filter-to'].forEach(id => {
    document.getElementById(id).addEventListener('keydown', e => {
      if (e.key === 'Enter') _fetchLogs();
    });
  });

  document.getElementById('results-text').addEventListener('click', async e => {
    if (!e.target.closest('[data-unignore]')) return;
    _ignoredTypes.clear();
    await _saveIgnoredTypes();
    _fetchLogs();
  });

  document.getElementById('pager').addEventListener('click', e => {
    const btn = e.target.closest('[data-page]');
    if (!btn || btn.disabled) return;
    const to = btn.dataset.page === 'first' ? 0
      : btn.dataset.page === 'prev' ? _pageIndex - 1
      : btn.dataset.page === 'last' ? _lastPageIndex()
      : _pageIndex + 1;
    if (to != null) _loadPage(to);
  });
  document.getElementById('pager').addEventListener('keydown', e => {
    if (e.key === 'Enter' && e.target.matches('.pt-pager__input')) _goToTypedPage(e.target);
  });
  document.getElementById('pager').addEventListener('change', e => {
    if (e.target.matches('.pt-pager__input')) _goToTypedPage(e.target);
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
      else reject(new Error(e.data.error || (e.data.data && e.data.data.error && e.data.data.error.message) || 'HTTP ' + e.data.status));
    }
    window.addEventListener('message', onMsg);
    // window.opener when popped out to its own tab; window.parent when embedded.
    (window.opener || window.parent).postMessage({ __efppt: 'fetch', id, url, method, headers: extraHeaders, body }, '*');
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

// ─── Fetch (server-side paging) ──────────────────────────────────────────────

const _ATTRS = [
  'plugintracelogid', 'typename', 'messagename', 'createdon', 'mode', 'depth', 'exceptiondetails',
  'messageblock', 'performanceexecutionduration', 'correlationid', 'requestid',
];

function _xml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function _likeEscape(s) {
  return String(s).replace(/[[%_]/g, '[$&]');
}

function _cond(attr, op, value) {
  return `<condition attribute="${attr}" operator="${op}" value="${_xml(value)}"/>`;
}

function _buildFilterXml() {
  const parts = [];
  const rawFrom = document.getElementById('filter-from').value;
  const rawTo   = document.getElementById('filter-to').value;
  if (rawFrom) parts.push(_cond('createdon', 'ge', new Date(rawFrom).toISOString()));
  if (rawTo)   parts.push(_cond('createdon', 'le', new Date(rawTo).toISOString()));
  if (_showFilter === 'exceptions') parts.push(_cond('exceptiondetails', 'like', '_%'));

  const q = document.getElementById('search').value.trim();
  if (q) {
    const v = `%${_likeEscape(q)}%`;
    parts.push(`<filter type="or">${['typename', 'messagename', 'exceptiondetails'].map(attr => _cond(attr, 'like', v)).join('')}</filter>`);
  }
  _ignoredTypes.forEach(name => {
    parts.push(_cond('typename', 'ne', name));
    parts.push(_cond('typename', 'not-like', `${_likeEscape(name)},%`));
  });
  return parts.length ? `<filter type="and">${parts.join('')}</filter>` : '';
}

function _fetchXmlUrl(xml) {
  return `${ENV_URL}/api/data/v9.2/${_entitySetName}?fetchXml=${encodeURIComponent(xml)}`;
}

function _pageUrl(index) {
  return _fetchXmlUrl(
    `<fetch count="${PAGE_SIZE}" page="${index + 1}" returntotalrecordcount="true">` +
      '<entity name="plugintracelog">' +
        _ATTRS.map(attr => `<attribute name="${attr}"/>`).join('') +
        '<order attribute="createdon" descending="true"/>' +
        '<order attribute="plugintracelogid" descending="true"/>' +
        _filterXml +
      '</entity>' +
    '</fetch>'
  );
}

function _countUrl() {
  return _fetchXmlUrl(
    '<fetch aggregate="true">' +
      '<entity name="plugintracelog">' +
        '<attribute name="plugintracelogid" alias="n" aggregate="count"/>' +
        _filterXml +
      '</entity>' +
    '</fetch>'
  );
}

async function _fetchLogs() {
  const token = ++_queryToken;
  clearTimeout(_searchTimer);
  _showLoading(_entitySetName ? 'Fetching plug-in trace logs…' : 'Resolving entity metadata…');

  try {
    await _resolveEntitySetName();
  } catch (e) {
    if (token === _queryToken) _showError(e.message);
    return;
  }
  if (token !== _queryToken) return;

  _filterXml   = _buildFilterXml();
  _pages       = [];
  _moreAfter   = [];
  _pageIndex   = 0;
  _totalCount  = null;
  _countState  = null;
  _openLogId   = null;
  _loadingPage = false;
  await _loadPage(0, token);
}

async function _countAll(token) {
  let n = null;
  try {
    const data = await _d365Fetch(_countUrl());
    n = data.value && data.value[0] ? Number(data.value[0].n) : null;
  } catch (e) {
    console.warn('[EF PPT] Could not count all trace logs:', e);
  }
  if (token !== _queryToken) return;
  if (n != null && !isNaN(n)) {
    _totalCount = n;
    _countState = 'exact';
  } else {
    _countState = 'capped';
  }
  _renderInfo();
}

function _lastPageIndex() {
  return _countState === 'exact' ? Math.max(0, Math.ceil(_totalCount / PAGE_SIZE) - 1) : null;
}

function _goToTypedPage(input) {
  const n = parseInt(input.value, 10);
  const last = _lastPageIndex();
  if (isNaN(n) || n < 1) { input.value = _pageIndex + 1; return; }
  const to = last == null ? n - 1 : Math.min(n - 1, last);
  if (to === _pageIndex) { input.value = _pageIndex + 1; return; }
  _loadPage(to);
}

async function _loadPage(index, token = _queryToken) {
  if (index < 0 || _loadingPage) return;
  if (_pages[index]) {
    _pageIndex = index;
    _renderLogs();
    return;
  }

  _loadingPage = true;
  if (index === 0 && !_pages.length) _showLoading('Fetching plug-in trace logs…');
  else _renderPager(true);

  let data;
  try {
    data = await _d365Fetch(_pageUrl(index), { 'Prefer': 'odata.include-annotations="*"' });
  } catch (e) {
    if (token !== _queryToken) return;
    _loadingPage = false;
    if (!_pages.length) _showError(e.message);
    else { _renderPager(false); _showPagerError(e.message); }
    return;
  }
  if (token !== _queryToken) return;
  _loadingPage = false;

  const rows = (data.value ?? []).slice(0, PAGE_SIZE);
  if (!rows.length && index > 0) {
    _renderPager(false);
    _showPagerError(`page ${index + 1} is empty`);
    return;
  }

  if (_countState == null) {
    const total = data['@Microsoft.Dynamics.CRM.totalrecordcount'];
    if (data['@Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded'] === true) {
      _totalCount = typeof total === 'number' && total > 0 ? total : 5000;
      _countState = 'counting';
      _countAll(token);
    } else if (typeof total === 'number' && total >= 0) {
      _totalCount = total;
      _countState = 'exact';
    }
  }
  _pages[index]     = rows;
  _moreAfter[index] = data['@Microsoft.Dynamics.CRM.morerecords'] === true;
  _pageIndex        = index;
  _fetchedAt        = new Date();
  _renderLogs();
}

// ─── Rendering ────────────────────────────────────────────────────────────────

function _pageCountText() {
  if (_totalCount == null) return null;
  const pages = Math.max(1, Math.ceil(_totalCount / PAGE_SIZE));
  return _countState === 'exact' ? pages.toLocaleString() : `${pages.toLocaleString()}+`;
}

function _hasNext() {
  const last = _lastPageIndex();
  return last != null ? _pageIndex < last : !!_moreAfter[_pageIndex];
}

function _renderPager(busy) {
  const pager = document.getElementById('pager');
  const of    = _pageCountText();
  const last  = _lastPageIndex();
  const atStart = _pageIndex === 0;
  const lastTitle = last != null ? `Last page (${last + 1})`
    : _countState === 'counting' ? 'Counting all logs — available in a moment'
    : 'Too many logs to count — narrow the date range or filters';
  pager.innerHTML = `
    <button class="pt-pager__btn" data-page="first" title="First page" ${busy || atStart ? 'disabled' : ''}>&laquo;</button>
    <button class="pt-pager__btn" data-page="prev" title="Previous page" ${busy || atStart ? 'disabled' : ''}>&lsaquo; Prev</button>
    <span class="pt-pager__label">${busy ? 'Loading…' : `Page <input type="number" class="pt-pager__input" min="1"${last != null ? ` max="${last + 1}"` : ''} value="${_pageIndex + 1}" aria-label="Page number" title="Type a page number and press Enter"> ${of ? `of ${of}` : ''}`}</span>
    <button class="pt-pager__btn" data-page="next" title="Next page" ${busy || !_hasNext() ? 'disabled' : ''}>Next &rsaquo;</button>
    <button class="pt-pager__btn" data-page="last" title="${_xml(lastTitle)}" ${busy || last == null || _pageIndex >= last ? 'disabled' : ''}>&raquo;</button>
  `;
  pager.classList.toggle('hidden', !busy && atStart && !_hasNext());
}

function _showPagerError(msg) {
  document.getElementById('results-text').textContent = `Could not load that page: ${msg}`;
}

function _renderLogs() {
  const container = document.getElementById('log-list');
  const info      = document.getElementById('results-info');
  const logs      = _pages[_pageIndex] ?? [];
  const filtered  = _showFilter === 'exceptions' || document.getElementById('search').value.trim() || _ignoredTypes.size;

  if (logs.length === 0) {
    info.classList.add('hidden');
    container.innerHTML = `<div class="pt-empty">${
      filtered
        ? 'No logs match the current search and filters.'
        : 'No plug-in trace logs found. Adjust the date range or check that trace logging is enabled in this environment.'
    }</div>`;
    return;
  }

  _renderInfo();

  const frag = document.createDocumentFragment();
  logs.forEach(log => frag.appendChild(_buildRow(log)));
  container.replaceChildren(frag);
  container.scrollTop = 0;

  if (_openLogId) {
    const row = container.querySelector(`[data-id="${CSS.escape(_openLogId)}"]`);
    if (row) _expandRow(row, logs.find(l => l.plugintracelogid === _openLogId));
  }
}

function _renderInfo() {
  const info     = document.getElementById('results-info');
  const text     = document.getElementById('results-text');
  const logs     = _pages[_pageIndex] ?? [];
  if (!logs.length) return;
  const filtered = _showFilter === 'exceptions' || document.getElementById('search').value.trim() || _ignoredTypes.size;
  const from  = _pageIndex * PAGE_SIZE + 1;
  const to    = from + logs.length - 1;
  const total = _totalCount == null ? null
    : `${_totalCount.toLocaleString()}${_countState === 'exact' ? '' : '+'}${_countState === 'counting' ? ' (counting…)' : ''}`;
  const ignored = [..._ignoredTypes].sort().join('\n');
  text.innerHTML = _esc(
    `${from.toLocaleString()}–${to.toLocaleString()}${total ? ` of ${total}` : ''} log${total === '1' ? '' : 's'}` +
    `${filtered ? ' (filtered)' : ''}`
  ) +
    (_ignoredTypes.size
      ? ` · <span title="${_esc(ignored)}">${_ignoredTypes.size} plugin type${_ignoredTypes.size === 1 ? '' : 's'} ignored</span> ` +
        `<button type="button" class="pt-link-btn" data-unignore>Show all</button>`
      : '') +
    _esc(` · fetched ${_fmtDateTime(_fetchedAt.toISOString())}`);
  info.classList.remove('hidden');
  _renderPager(_loadingPage);
}

function _expandRow(row, log) {
  if (!row.querySelector('.pt-log-detail')) row.appendChild(_buildDetail(log));
  row.classList.add('expanded');
}

function _buildRow(log) {
  const hasException = !!log.exceptiondetails;
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

  // Toggle expand on summary click.
  summary.addEventListener('click', () => {
    const isOpen = row.classList.contains('expanded');
    document.querySelectorAll('.pt-log-row.expanded').forEach(r => r.classList.remove('expanded'));
    if (!isOpen) {
      _expandRow(row, log);
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
  return row;
}

function _buildDetail(log) {
  const hasException = !!log.exceptiondetails;
  const hasTrace     = !!log.messageblock;
  const isSync       = log.mode === 0;
  const modeLabel    = isSync ? 'Sync' : log.mode === 1 ? 'Async' : `Mode ${log.mode}`;
  const duration     = log.performanceexecutionduration != null
    ? `${log.performanceexecutionduration} ms`
    : '—';

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

  return detail;
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
