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
let _view          = 'logs';
let _sumPeriod     = '7';
let _sumKey        = null;
let _sumToken      = 0;
let _sumData       = null;
let _pane          = null;

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
  _ignoredChanged();
}

function _ignoredChanged() {
  _sumKey = null;
  _fetchLogs();
  if (_view === 'summary') _loadSummary();
}

// ─── Bootstrap ────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {
  document.getElementById('env-name').textContent = ENV_NAME;
  document.title = `Plugin Trace Logs — ${ENV_NAME}`;

  _loadIgnoredTypes();

  // Wire controls.
  document.getElementById('btn-refresh').addEventListener('click', () => {
    if (_view === 'summary') _loadSummary(true);
    else _fetchLogs();
  });
  _bindSummary();
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
    _ignoredChanged();
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

  _setView(_readView());
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

function _bridgeFetch(url, extraHeaders = {}, method = 'GET', body = null, timeoutMs = 30_000) {
  return new Promise((resolve, reject) => {
    const id    = Math.random().toString(36).slice(2) + Date.now();
    const timer = setTimeout(() => {
      window.removeEventListener('message', onMsg);
      reject(new Error(`D365 didn't respond within ${Math.round(timeoutMs / 1000)} seconds — try a shorter period.`));
    }, timeoutMs);
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

async function _d365Fetch(url, extraHeaders = {}, timeoutMs = 30_000) {
  return _bridgeFetch(url, extraHeaders, 'GET', null, timeoutMs);
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
    const inMainList = row.parentElement && row.parentElement.id === 'log-list';
    (row.parentElement || document).querySelectorAll('.pt-log-row.expanded').forEach(r => r.classList.remove('expanded'));
    if (!isOpen) {
      _expandRow(row, log);
      if (inMainList) _openLogId = log.plugintracelogid;
    } else if (inMainList) {
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

function _readView() {
  try { return localStorage.getItem('ef_ppt_pluginTraceView') === 'summary' ? 'summary' : 'logs'; }
  catch { return 'logs'; }
}

function _setView(view, remember = false) {
  _view = view === 'summary' ? 'summary' : 'logs';
  document.querySelectorAll('.pt-tab').forEach(btn => {
    const on = btn.dataset.view === _view;
    btn.classList.toggle('active', on);
    btn.setAttribute('aria-selected', on ? 'true' : 'false');
  });
  document.getElementById('logs-view').classList.toggle('hidden', _view !== 'logs');
  document.getElementById('summary-view').classList.toggle('hidden', _view !== 'summary');
  if (remember) { try { localStorage.setItem('ef_ppt_pluginTraceView', _view); } catch { } }
  if (_view === 'summary' && ENV_URL) _loadSummary();
}

function _bindSummary() {
  document.querySelectorAll('.pt-tab').forEach(btn =>
    btn.addEventListener('click', () => _setView(btn.dataset.view, true))
  );
  document.querySelectorAll('.sum-period__btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.sum-period__btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      _sumPeriod = btn.dataset.period;
      _loadSummary();
    });
  });
  document.getElementById('summary-view').addEventListener('click', async e => {
    if (e.target.closest('[data-unignore]')) {
      _ignoredTypes.clear();
      await _saveIgnoredTypes();
      _ignoredChanged();
      return;
    }
    const target = e.target.closest('[data-open]');
    if (target) _openPane(target.dataset.open, target.dataset.plugin || null);
  });
  document.getElementById('sum-pane-close').addEventListener('click', () => _closePane());
  document.getElementById('sum-pane-pager').addEventListener('click', e => {
    const btn = e.target.closest('[data-pp]');
    if (!btn || btn.disabled || !_pane) return;
    _paneLoad(btn.dataset.pp === 'first' ? 0 : btn.dataset.pp === 'prev' ? _pane.page - 1 : _pane.page + 1);
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && _pane && !document.getElementById('_pt-ctx-menu')) _closePane();
  });
}

function _periodText() {
  return _sumPeriod === '1' ? 'last 24 hours' : `last ${_sumPeriod} days`;
}

let _sumWindow = null;

function _sumFrom() {
  return _sumWindow ? _sumWindow.start : new Date(Date.now() - Number(_sumPeriod) * 86_400_000).toISOString();
}

function _makeSumWindow() {
  const end = Date.now();
  const span = Number(_sumPeriod) * 86_400_000;
  const n = _sumPeriod === '1' ? 4 : 7;
  const step = span / n;
  return {
    start: new Date(end - span).toISOString(),
    slices: Array.from({ length: n }, (_, i) => ({
      from: new Date(end - step * (i + 1)).toISOString(),
      to: i === 0 ? null : new Date(end - step * i).toISOString(),
    })),
  };
}

function _exceptionCond() {
  return _cond('exceptiondetails', 'like', '_%');
}

function _pluginFilter(short) {
  return `<filter type="or">${_cond('typename', 'eq', short)}${_cond('typename', 'like', `${_likeEscape(short)},%`)}</filter>`;
}

function _sumFilter(extra = [], range = null) {
  const ignored = [..._ignoredTypes].flatMap(name => [
    _cond('typename', 'ne', name),
    _cond('typename', 'not-like', `${_likeEscape(name)},%`),
  ]);
  const time = range
    ? [_cond('createdon', 'ge', range.from), ...(range.to ? [_cond('createdon', 'lt', range.to)] : [])]
    : [_cond('createdon', 'ge', _sumFrom())];
  return `<filter type="and">${[...time, ...ignored, ...extra].join('')}</filter>`;
}

function _fmtMs(ms) {
  if (ms == null || isNaN(ms)) return '—';
  ms = Math.round(Number(ms));
  if (ms < 1000) return `${ms.toLocaleString()} ms`;
  const s = ms / 1000;
  if (s < 60) return `${s.toFixed(s < 10 ? 2 : 1)} s`;
  const m = Math.floor(s / 60);
  return `${m}m ${Math.round(s % 60)}s`;
}

const LONG_RUN_MS = 10_000;
const SUMMARY_TIMEOUT_MS = 90_000;

function _longRunCond() {
  return _cond('performanceexecutionduration', 'ge', String(LONG_RUN_MS));
}

function _isRecordLimitError(e) {
  return /AggregateQueryRecordLimit|record limit|50000|50,000/i.test(String(e && e.message));
}

async function _groupSlice(extra, kind, range) {
  const filter = _sumFilter(extra, range);
  const aggXml =
    '<fetch aggregate="true"><entity name="plugintracelog">' +
      '<attribute name="typename" alias="t" groupby="true"/>' +
      '<attribute name="plugintracelogid" alias="n" aggregate="count"/>' +
      '<attribute name="createdon" alias="last" aggregate="max"/>' +
      (kind === 'long' ? '<attribute name="performanceexecutionduration" alias="mx" aggregate="max"/>' : '') +
      filter +
    '</entity></fetch>';
  try {
    const d = await _d365Fetch(_fetchXmlUrl(aggXml), {}, SUMMARY_TIMEOUT_MS);
    return {
      rows: (d.value ?? []).map(r => ({ t: r.t, n: Number(r.n) || 0, mx: r.mx != null ? Number(r.mx) : null, last: r.last ?? null })),
      exact: true,
    };
  } catch (e) {
    if (!_isRecordLimitError(e)) throw e;
    console.warn('[EF PPT] Too many logs to aggregate in one slice; using its newest 5,000.', e);
  }
  const sampleXml =
    '<fetch count="5000" page="1"><entity name="plugintracelog">' +
      '<attribute name="typename"/><attribute name="createdon"/><attribute name="performanceexecutionduration"/>' +
      '<order attribute="createdon" descending="true"/>' +
      filter +
    '</entity></fetch>';
  const d = await _d365Fetch(_fetchXmlUrl(sampleXml), { 'Prefer': 'odata.include-annotations="*"' }, SUMMARY_TIMEOUT_MS);
  return {
    rows: (d.value ?? []).map(r => ({
      t: r.typename, n: 1,
      mx: r.performanceexecutionduration != null ? Number(r.performanceexecutionduration) : null,
      last: r.createdon,
    })),
    exact: d['@Microsoft.Dynamics.CRM.morerecords'] !== true,
    sampled: true,
  };
}

async function _groupStats(extra, kind, slices, onProgress) {
  const groups = new Map();
  const state = { groups: [], exact: true, sampled: false, done: 0, slices: slices.length, failed: 0 };
  let lastError = null;
  const merge = rows => rows.forEach(r => {
    const name = _shortName(r.t ?? '') || '(unknown plugin)';
    const g = groups.get(name) || { name, n: 0, mx: null, last: null };
    g.n += r.n;
    if (r.mx != null && (g.mx == null || r.mx > g.mx)) g.mx = r.mx;
    if (r.last && (!g.last || r.last > g.last)) g.last = r.last;
    groups.set(name, g);
  });
  await Promise.all(slices.map(async range => {
    try {
      const r = await _groupSlice(extra, kind, range);
      merge(r.rows);
      if (!r.exact) state.exact = false;
      if (r.sampled) state.sampled = true;
    } catch (e) {
      state.failed++;
      state.exact = false;
      lastError = e;
    }
    state.done++;
    state.groups = [...groups.values()];
    onProgress({ ...state });
  }));
  if (state.failed === slices.length) throw lastError;
  return { ...state };
}

async function _recentExceptions(slices, limit) {
  const found = [];
  let lastError = null;
  for (const range of slices) {
    const xml =
      `<fetch count="${limit - found.length}" page="1"><entity name="plugintracelog">` +
        ['plugintracelogid', 'typename', 'messagename', 'createdon', 'exceptiondetails', 'mode', 'performanceexecutionduration']
          .map(a => `<attribute name="${a}"/>`).join('') +
        '<order attribute="createdon" descending="true"/>' +
        _sumFilter([_exceptionCond()], range) +
      '</entity></fetch>';
    try {
      const d = await _d365Fetch(_fetchXmlUrl(xml), {}, SUMMARY_TIMEOUT_MS);
      found.push(...(d.value ?? []));
    } catch (e) {
      lastError = e;
    }
    if (found.length >= limit) break;
  }
  if (!found.length && lastError) throw lastError;
  return found.slice(0, limit);
}

async function _loadSummary(force = false) {
  const key = `${_sumPeriod}|${[..._ignoredTypes].sort().join(',')}`;
  if (!force && key === _sumKey) return;
  _sumKey = key;
  const token = ++_sumToken;
  _closePane(false);
  _sumWindow = _makeSumWindow();

  document.getElementById('sum-scope').innerHTML = _esc(`${_periodText()} · exceptions and runs of ${LONG_RUN_MS / 1000} s or more`) +
    (_ignoredTypes.size
      ? ` · <span title="${_esc([..._ignoredTypes].sort().join('\n'))}">${_ignoredTypes.size} plugin type${_ignoredTypes.size === 1 ? '' : 's'} ignored</span> <button type="button" class="pt-link-btn" data-unignore>Show all</button>`
      : '');
  document.getElementById('sum-updated').textContent = '';

  const data = _sumData = { ex: null, exErr: null, long: null, longErr: null, recent: null, recentErr: null };
  _renderSummary();

  try {
    await _resolveEntitySetName();
  } catch (e) {
    if (token !== _sumToken) return;
    _sumData = null;
    document.getElementById('sum-content').innerHTML = `<div class="sum-msg sum-msg--error">${_esc(e.message)}</div>`;
    return;
  }
  if (token !== _sumToken) return;
  const update = () => { if (token === _sumToken) _renderSummary(); };
  const slices = _sumWindow.slices;

  await Promise.all([
    _recentExceptions(slices, 12).then(r => { data.recent = r; }, e => { data.recentErr = e.message; }).then(update),
    _groupStats([_exceptionCond()], 'exceptions', slices, p => { data.ex = p; update(); })
      .then(r => { data.ex = r; }, e => { data.exErr = e.message; }).then(update),
    _groupStats([_longRunCond()], 'long', slices, p => { data.long = p; update(); })
      .then(r => { data.long = r; }, e => { data.longErr = e.message; }).then(update),
  ]);
  if (token !== _sumToken) return;
  const now = new Date();
  document.getElementById('sum-updated').textContent =
    `Updated ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
}

function _pluginButton(kind, name) {
  const active = _pane && _pane.kind === kind && _pane.plugin === name;
  return `<button type="button" class="sum-flow${active ? ' is-active' : ''}" data-open="${kind}" data-plugin="${_esc(name)}" title="${_esc(name)}">${_esc(name)}</button>`;
}

function _sumKpi(kind, label, value, sub, tone = '') {
  const active = _pane && _pane.kind === kind && !_pane.plugin;
  return (
    `<button type="button" class="sum-kpi${tone ? ` sum-kpi--${tone}` : ''}${active ? ' is-active' : ''}" data-open="${kind}">` +
      `<span class="sum-kpi-label">${_esc(label)}</span>` +
      `<span class="sum-kpi-value">${value}</span>` +
      `<span class="sum-kpi-sub">${sub}</span>` +
    `</button>`
  );
}

function _sumBars(kind, rows, valueOf, labelOf, titleOf, tone = '') {
  if (!rows.length) return `<div class="sum-msg">Nothing in the ${_esc(_periodText())}.</div>`;
  const peak = Math.max(...rows.map(valueOf)) || 1;
  return `<ul class="sum-bars${tone ? ` sum-bars--${tone}` : ''}">${rows.map(r => `
    <li>
      ${_pluginButton(kind, r.name)}
      <span class="sum-bar"><span style="width:${Math.max(4, (valueOf(r) / peak) * 100)}%"></span></span>
      <span class="sum-val" title="${_esc(titleOf(r))}">${_esc(labelOf(r))}</span>
    </li>`).join('')}</ul>`;
}

function _renderSummary() {
  const d = _sumData;
  if (!d) return;
  const content = document.getElementById('sum-content');
  if (!content.querySelector('.sum-kpis')) {
    content.innerHTML = `
      <div class="sum-kpis" id="sum-kpis"></div>
      <div class="sum-grid">
        <section class="sum-card">
          <header class="sum-card-head"><span>Most exceptions</span><small>failed executions per plugin</small></header>
          <div class="sum-card-body" id="sum-exceptions"></div>
        </section>
        <section class="sum-card">
          <header class="sum-card-head"><span>Long-running plugins</span><small>runs of ${LONG_RUN_MS / 1000} s or more</small></header>
          <div class="sum-card-body" id="sum-long"></div>
        </section>
        <section class="sum-card sum-card--wide">
          <header class="sum-card-head"><span>Most recent exceptions</span><small>click a plugin to see all its exceptions</small></header>
          <div class="sum-card-body" id="sum-recent"></div>
        </section>
      </div>`;
  }
  const loading = '<div class="sum-msg sum-msg--loading">Loading…</div>';
  const err = msg => `<div class="sum-msg sum-msg--error">${_esc(msg)}</div>`;
  const counting = r => r && r.done < r.slices;
  const sampledNote = r => !r ? '' :
    (counting(r) ? `<div class="sum-note">Counting… ${r.done} of ${r.slices} time slices done.</div>` : '') +
    (r.failed ? `<div class="sum-note sum-note--warn">${r.failed} of ${r.slices} time slices didn't load in time — these numbers are incomplete. Try Refresh or a shorter period.</div>` : '') +
    (r.sampled && !r.exact && !r.failed ? '<div class="sum-note">Too many logs to total exactly — part of the period is based on its newest 5,000 logs.</div>' : '');
  const plus = r => r && (!r.exact || counting(r)) ? '+' : '';
  const progress = (r, text) => counting(r) ? `counting… ${r.done}/${r.slices}` : text;

  const ex = d.ex ? d.ex.groups : null;
  const long = d.long ? d.long.groups : null;
  const exTotal = ex ? ex.reduce((a, g) => a + g.n, 0) : null;
  const longTotal = long ? long.reduce((a, g) => a + g.n, 0) : null;
  const slow = long ? long.filter(g => g.mx != null).sort((a, b) => b.mx - a.mx) : [];
  const longLabel = `${LONG_RUN_MS / 1000} s or more`;

  document.getElementById('sum-kpis').innerHTML =
    _sumKpi('exceptions', 'Exceptions', d.exErr ? '—' : ex ? `${exTotal.toLocaleString()}${plus(d.ex)}` : '…', progress(d.ex, _esc(_periodText())), exTotal ? 'bad' : '') +
    _sumKpi('exceptions', 'Plugins with exceptions', d.exErr ? '—' : ex ? ex.length.toLocaleString() : '…', progress(d.ex, 'with at least one exception')) +
    _sumKpi('long', 'Long runs', d.longErr ? '—' : long ? `${longTotal.toLocaleString()}${plus(d.long)}` : '…', progress(d.long, _esc(longLabel)), longTotal ? 'warn' : '') +
    _sumKpi('long', 'Plugins with long runs', d.longErr ? '—' : long ? long.length.toLocaleString() : '…', progress(d.long, _esc(longLabel))) +
    _sumKpi('long', 'Slowest run', d.longErr ? '—' : long ? (slow[0] ? _esc(_fmtMs(slow[0].mx)) : '—') : '…', slow[0] ? _esc(slow[0].name) : long ? 'no long runs' : '');

  document.getElementById('sum-exceptions').innerHTML = d.exErr ? err(d.exErr) : !ex ? loading :
    !ex.length && counting(d.ex) ? loading + sampledNote(d.ex) :
    !ex.length ? `<div class="sum-msg sum-msg--ok">No exceptions in the ${_esc(_periodText())}.</div>` :
    _sumBars('exceptions', ex.slice().sort((a, b) => b.n - a.n).slice(0, 10), g => g.n, g => g.n.toLocaleString(), g => g.last ? `Last exception ${_fmtDateTime(g.last)}` : '', 'bad') + sampledNote(d.ex);

  document.getElementById('sum-long').innerHTML = d.longErr ? err(d.longErr) : !long ? loading :
    !long.length && counting(d.long) ? loading + sampledNote(d.long) :
    !long.length ? `<div class="sum-msg sum-msg--ok">No runs of ${_esc(longLabel)} in the ${_esc(_periodText())}.</div>` :
    _sumBars('long', slow.slice(0, 10), g => g.mx, g => _fmtMs(g.mx), g => `${g.n.toLocaleString()} run${g.n === 1 ? '' : 's'} of ${longLabel}${g.last ? ` · last ${_fmtDateTime(g.last)}` : ''}`, 'warn') + sampledNote(d.long);

  const recentEl = document.getElementById('sum-recent');
  if (d.recentErr) recentEl.innerHTML = err(d.recentErr);
  else if (!d.recent) recentEl.innerHTML = loading;
  else if (!d.recent.length) recentEl.innerHTML = `<div class="sum-msg sum-msg--ok">No exceptions in the ${_esc(_periodText())}.</div>`;
  else {
    recentEl.innerHTML = `<table class="sum-recent">
      <colgroup><col style="width:130px"><col style="width:30%"><col style="width:110px"><col></colgroup>
      <tbody>${d.recent.map(log => {
        const name = _shortName(log.typename ?? '') || '(unknown plugin)';
        const first = String(log.exceptiondetails ?? '').split(/\r?\n/).find(l => l.trim()) ?? '';
        return `<tr>
          <td class="sum-time">${_esc(_fmtDateTime(log.createdon))}</td>
          <td>${_pluginButton('exceptions', name)}</td>
          <td class="sum-msgname">${_esc(log.messagename ?? '—')}</td>
          <td class="sum-err" title="${_esc(String(log.exceptiondetails ?? '').slice(0, 1500))}">${_esc(first.length > 160 ? first.slice(0, 160) + '…' : first)}</td>
        </tr>`;
      }).join('')}</tbody>
    </table>`;
  }
}

const PANE_KINDS = {
  exceptions: { title: 'Exceptions', sub: 'newest first',  order: 'createdon', exceptions: true },
  long:       { title: 'Long runs',  sub: `${LONG_RUN_MS / 1000} s or more · slowest first`, order: 'duration', long: true },
};
let _paneToken = 0;

function _openPane(kind, plugin) {
  const cfg = PANE_KINDS[kind];
  if (!cfg) return;
  const stats = _sumData && (kind === 'exceptions' ? _sumData.ex : _sumData.long);
  let total = null;
  if (stats && stats.done === stats.slices && stats.exact) {
    total = plugin
      ? (stats.groups.find(g => g.name === plugin) || { n: 0 }).n
      : stats.groups.reduce((a, g) => a + g.n, 0);
  }
  _pane = { kind, plugin: plugin || null, pages: [], more: [], total, capped: false, page: 0, loading: false, token: ++_paneToken };
  document.getElementById('sum-pane-title').textContent = `${cfg.title} — ${plugin || 'all plugins'}`;
  document.getElementById('sum-pane-title').title = plugin || '';
  document.getElementById('sum-pane-sub').textContent = `${_periodText()} · ${cfg.sub}`;
  document.getElementById('sum-pane').classList.remove('hidden');
  document.getElementById('summary-view').classList.add('pt-summary--pane');
  _renderSummary();
  _paneLoad(0);
}

function _closePane(rerender = true) {
  if (!_pane) return;
  _pane = null;
  _paneToken++;
  document.getElementById('sum-pane').classList.add('hidden');
  document.getElementById('summary-view').classList.remove('pt-summary--pane');
  if (rerender) _renderSummary();
}

function _paneXml(p, index) {
  const cfg = PANE_KINDS[p.kind];
  const extra = [];
  if (cfg.exceptions) extra.push(_exceptionCond());
  if (cfg.long) extra.push(_longRunCond());
  if (p.plugin) extra.push(_pluginFilter(p.plugin));
  const order = cfg.order === 'duration'
    ? '<order attribute="performanceexecutionduration" descending="true"/><order attribute="createdon" descending="true"/>'
    : '<order attribute="createdon" descending="true"/>';
  return `<fetch count="${PAGE_SIZE}" page="${index + 1}">` +
    '<entity name="plugintracelog">' +
      _ATTRS.map(attr => `<attribute name="${attr}"/>`).join('') +
      order +
      '<order attribute="plugintracelogid" descending="true"/>' +
      _sumFilter(extra) +
    '</entity></fetch>';
}

async function _paneLoad(index) {
  const p = _pane;
  if (!p || index < 0 || p.loading) return;
  if (p.pages[index]) {
    p.page = index;
    _renderPane();
    return;
  }
  p.loading = true;
  if (index === 0) {
    document.getElementById('sum-pane-list').innerHTML = '<div class="sum-msg sum-msg--loading">Loading logs…</div>';
    document.getElementById('sum-pane-count').textContent = '';
    document.getElementById('sum-pane-pager').classList.add('hidden');
  } else {
    _renderPanePager(true);
  }
  let result;
  try {
    if (PANE_KINDS[p.kind].exceptions) {
      const prev = index > 0 ? p.pages[index - 1] : null;
      result = await _paneFetchSliced(p, prev && prev.length ? prev[prev.length - 1].createdon : null);
    } else {
      const d = await _d365Fetch(_fetchXmlUrl(_paneXml(p, index)), { 'Prefer': 'odata.include-annotations="*"' }, SUMMARY_TIMEOUT_MS);
      result = { rows: (d.value ?? []).slice(0, PAGE_SIZE), more: d['@Microsoft.Dynamics.CRM.morerecords'] === true };
    }
  } catch (e) {
    if (p !== _pane) return;
    p.loading = false;
    if (index === 0) document.getElementById('sum-pane-list').innerHTML = `<div class="sum-msg sum-msg--error">${_esc(e.message)}</div>`;
    else _renderPanePager(false);
    return;
  }
  if (p !== _pane) return;
  p.loading = false;
  p.pages[index] = result.rows;
  p.more[index] = result.more;
  p.page = index;
  _renderPane();
}

async function _paneFetchSliced(p, before) {
  const extra = [_exceptionCond()];
  if (p.plugin) extra.push(_pluginFilter(p.plugin));
  const slices = _sumWindow.slices.filter(slice => !before || slice.from < before);
  const results = await Promise.all(slices.map(async slice => {
    const range = { from: slice.from, to: before && (!slice.to || before < slice.to) ? before : slice.to };
    const xml = `<fetch count="${PAGE_SIZE}" page="1"><entity name="plugintracelog">` +
      _ATTRS.map(attr => `<attribute name="${attr}"/>`).join('') +
      '<order attribute="createdon" descending="true"/>' +
      '<order attribute="plugintracelogid" descending="true"/>' +
      _sumFilter(extra, range) +
      '</entity></fetch>';
    const d = await _d365Fetch(_fetchXmlUrl(xml), { 'Prefer': 'odata.include-annotations="*"' }, SUMMARY_TIMEOUT_MS);
    return { rows: d.value ?? [], more: d['@Microsoft.Dynamics.CRM.morerecords'] === true };
  }));
  const merged = results.flatMap(r => r.rows)
    .sort((a, b) => String(b.createdon).localeCompare(String(a.createdon)) || String(b.plugintracelogid).localeCompare(String(a.plugintracelogid)));
  return { rows: merged.slice(0, PAGE_SIZE), more: merged.length > PAGE_SIZE || results.some(r => r.more) };
}

function _renderPanePager(busy) {
  const p = _pane;
  const pager = document.getElementById('sum-pane-pager');
  const show = busy || p.page > 0 || !!p.more[p.page];
  pager.classList.toggle('hidden', !show);
  if (!show) return;
  const pages = p.total != null ? Math.max(1, Math.ceil(p.total / PAGE_SIZE)) : null;
  pager.innerHTML = `
    <button class="pt-pager__btn" data-pp="first" title="First page" ${busy || p.page === 0 ? 'disabled' : ''}>&laquo;</button>
    <button class="pt-pager__btn" data-pp="prev" ${busy || p.page === 0 ? 'disabled' : ''}>&lsaquo; Prev</button>
    <span class="pt-pager__label">${busy ? 'Loading…' : `Page ${p.page + 1}${pages ? ` of ${pages.toLocaleString()}${p.capped ? '+' : ''}` : ''}`}</span>
    <button class="pt-pager__btn" data-pp="next" ${busy || !p.more[p.page] ? 'disabled' : ''}>Next &rsaquo;</button>
  `;
}

function _renderPane() {
  const p = _pane;
  const rows = p.pages[p.page] ?? [];
  const list = document.getElementById('sum-pane-list');
  const from = p.page * PAGE_SIZE + 1;
  document.getElementById('sum-pane-count').textContent = rows.length
    ? `${from.toLocaleString()}–${(from + rows.length - 1).toLocaleString()}${p.total != null ? ` of ${p.total.toLocaleString()}${p.capped ? '+' : ''}` : ''}`
    : '0';
  _renderPanePager(false);
  if (!rows.length) {
    list.innerHTML = '<div class="sum-msg">No logs match.</div>';
    return;
  }
  const frag = document.createDocumentFragment();
  rows.forEach(log => frag.appendChild(_buildRow(log)));
  list.replaceChildren(frag);
  list.scrollTop = 0;
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
