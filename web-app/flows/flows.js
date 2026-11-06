/**
 * EF Power Platform Tools — Power Automate Flows (web-app edition)
 *
 * Full-page viewer for Power Automate (Modern Flow) records and executions.
 * Runs as an iframe opened by launcher.js; all D365 fetches are relayed
 * through the parent via the __efppt postMessage bridge.
 */

// ─── URL params ────────────────────────────────────────────────────────────────

const _params  = new URLSearchParams(location.search);
const ENV_URL  = (_params.get('env')     ?? '').replace(/\/$/, '');
const ENV_NAME = _params.get('name')    ?? ENV_URL;
const PA_ENV_ID = _params.get('paEnvId') ?? '';

// ─── State ─────────────────────────────────────────────────────────────────────

let _solutions           = [];
let _selectedSolutionIds = [];
let _favSolutionIds      = [];
let _pendingSolutionIds  = null;
let _allFlows            = [];
let _flowIdSet           = new Set();
let _filteredFlows       = [];
let _selectedFlowId      = null;
const _orgId             = PA_ENV_ID; // Power Platform environment ID from extension settings
const RUN_PAGE_SIZE      = 100;
const RUN_MAX_FLOWS      = 300;
const RUN_SELECT         = 'name,workflowid,status,errorcode,starttime,createdon,duration,modifiedon';
let _runMetaCache        = null;
let _runErrorSearchOff   = false;
let _execReloadTimer     = null;
let _flowsToken          = 0;
let _execList            = null;
let _paneList            = null;
let _paneQuery           = null;
let _view                = 'runs';
let _summaryPeriod       = '7';
let _summaryKey          = null;
let _summaryToken        = 0;
let _summaryData         = null;

const DEFAULT_SOL_KEY = `flows_defaultSol_${ENV_URL}`;
const FAV_SOL_KEY     = `flows_favSols_${ENV_URL}`;
const VIEW_KEY        = 'flows_view';

// ─── Bootstrap ─────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {
  document.title = `Flows — ${ENV_NAME}`;
  document.getElementById('env-name').textContent = ENV_NAME;
  document.getElementById('env-url').textContent  = ENV_URL.replace('https://', '');

  _setDefaultDates();
  _execList = _createRunList({
    listId: 'exec-list',
    pagerId: 'exec-pager',
    sort: { key: 'date', dir: -1 },
    buildQuery: _execQuery,
    onCount: _setExecCount,
    emptyText: () => `No executions match the selected date range and filters.${
      _runErrorSearchOff && document.getElementById('exec-search').value.trim() ? ' (Search covers flow names only in this environment.)' : ''
    }`,
    onFirstPageError: () => {
      if (!document.getElementById('exec-search').value.trim() || _runErrorSearchOff) return false;
      _runErrorSearchOff = true;
      return true;
    },
  });
  _paneList = _createRunList({
    listId: 'sum-pane-list',
    pagerId: 'sum-pane-pager',
    sort: { key: 'date', dir: -1 },
    buildQuery: _paneQueryUrl,
    onCount: (text, title) => {
      const el = document.getElementById('sum-pane-count');
      el.textContent = text;
      el.title = title || '';
    },
    emptyText: () => 'No runs match.',
  });

  _loadFavourites();
  _bindControls();
  _setView(_readView(), false);

  if (!ENV_URL) {
    _showPanelMsg('flow-list', 'No environment URL provided. Open from the EF Power Platform Tools popup.', 'error');
    return;
  }

  _showPanelMsg('flow-list', 'Loading solutions…', 'loading');
  try {
    await _loadSolutions();
  } catch (e) {
    console.error('[EF PPT]', e);
    _showPanelMsg('flow-list', e.message, 'error');
  }
});

// ─── Controls ──────────────────────────────────────────────────────────────────

function _bindControls() {
  _bindSolutionPicker();

  document.querySelectorAll('.view-tab').forEach(btn =>
    btn.addEventListener('click', () => _setView(btn.dataset.view, true))
  );

  // Refresh
  document.getElementById('btn-refresh').addEventListener('click', async () => {
    if (_selectedSolutionIds.length) await _loadFlows(true);
  });

  // Flow search + status filter
  document.getElementById('flow-search').addEventListener('input', _renderFlows);
  document.querySelectorAll('#flow-status-filter .seg-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#flow-status-filter .seg-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      _renderFlows();
    });
  });

  // Execution controls
  document.getElementById('btn-apply-exec').addEventListener('click', _loadExecutions);
  document.getElementById('exec-search').addEventListener('input', _queueLoadExecutions);
  document.getElementById('exec-search').addEventListener('keydown', e => { if (e.key === 'Enter') _loadExecutions(); });
  document.getElementById('exec-status-filter').addEventListener('change', _loadExecutions);
  ['exec-from', 'exec-to'].forEach(id =>
    document.getElementById(id).addEventListener('keydown', e => { if (e.key === 'Enter') _loadExecutions(); })
  );

  // Duration filter
  ['dur-op', 'dur-unit'].forEach(id =>
    document.getElementById(id).addEventListener('change', _queueLoadExecutions)
  );
  document.getElementById('dur-value').addEventListener('input', _queueLoadExecutions);
  document.getElementById('btn-clear-dur').addEventListener('click', () => {
    document.getElementById('dur-op').value    = '';
    document.getElementById('dur-value').value = '';
    document.getElementById('dur-unit').value  = 'ms';
    _loadExecutions();
  });

  // Clear flow filter
  document.getElementById('btn-clear-flow-filter').addEventListener('click', () => {
    _selectedFlowId = null;
    document.getElementById('exec-scope').textContent = '';
    document.getElementById('btn-clear-flow-filter').classList.add('hidden');
    document.querySelectorAll('.flow-item').forEach(el => el.classList.remove('selected'));
    _loadExecutions();
  });

  document.querySelectorAll('#sum-period .seg-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#sum-period .seg-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      _summaryPeriod = btn.dataset.period;
      _loadSummary();
    });
  });
  document.getElementById('summary-view').addEventListener('click', e => {
    const target = e.target.closest('[data-open]');
    if (!target) return;
    _openPane(target.dataset.open, target.dataset.flow || null);
  });
  document.getElementById('sum-pane-close').addEventListener('click', _closePane);
  document.addEventListener('keydown', e => {
    if (e.key !== 'Escape') return;
    if (!document.getElementById('sol-ms-panel').classList.contains('hidden')) _closeSolutionPicker(true);
    else if (!document.getElementById('sum-pane').classList.contains('hidden')) _closePane();
  });
}

function _setDefaultDates() {
  const from = new Date();
  from.setDate(from.getDate() - 7);
  from.setHours(0, 0, 0, 0);   // start of day, 7 days ago

  const to = new Date();
  to.setHours(23, 59, 0, 0);   // end of today

  document.getElementById('exec-from').value = _toInputDateTime(from);
  document.getElementById('exec-to').value   = _toInputDateTime(to);
}

function _readView() {
  try { return localStorage.getItem(VIEW_KEY) === 'summary' ? 'summary' : 'runs'; }
  catch { return 'runs'; }
}

function _setView(view, remember) {
  _view = view === 'summary' ? 'summary' : 'runs';
  document.querySelectorAll('.view-tab').forEach(btn => {
    const on = btn.dataset.view === _view;
    btn.classList.toggle('active', on);
    btn.setAttribute('aria-selected', on ? 'true' : 'false');
  });
  document.getElementById('runs-view').classList.toggle('hidden', _view !== 'runs');
  document.getElementById('summary-view').classList.toggle('hidden', _view !== 'summary');
  if (remember) { try { localStorage.setItem(VIEW_KEY, _view); } catch { } }
  if (_view === 'summary' && _allFlows.length) _loadSummary();
}

// ─── Fetch bridge ──────────────────────────────────────────────────────────────

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
    // window.opener when popped out to its own tab; window.parent when embedded.
    (window.opener || window.parent).postMessage({ __efppt: 'fetch', id, url, method, headers: extraHeaders, body }, '*');
  });
}

async function _d365Fetch(url, extraHeaders = {}) {
  return _bridgeFetch(url, extraHeaders);
}

async function _fetchAllPages(firstUrl, maxPages = 20) {
  const records = [];
  let nextUrl   = firstUrl;
  let page      = 0;
  while (nextUrl && page < maxPages) {
    page++;
    const data = await _d365Fetch(nextUrl, {
      'Prefer': 'odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=500',
    });
    records.push(...(data.value ?? []));
    nextUrl = data['@odata.nextLink'] ?? null;
  }
  return records;
}

function _runPrefer(pageSize) {
  return { 'Prefer': `odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=${pageSize}` };
}

async function _runMeta() {
  if (_runMetaCache) return _runMetaCache;
  const meta = { idText: false, durationNumeric: true };
  try {
    const filter = encodeURIComponent("LogicalName eq 'workflowid' or LogicalName eq 'duration'");
    const data = await _d365Fetch(
      `${ENV_URL}/api/data/v9.2/EntityDefinitions(LogicalName='flowrun')/Attributes?$select=LogicalName,AttributeType&$filter=${filter}`
    );
    (data.value ?? []).forEach(attr => {
      if (attr.LogicalName === 'workflowid') meta.idText = attr.AttributeType === 'String' || attr.AttributeType === 'Memo';
      if (attr.LogicalName === 'duration') meta.durationNumeric = ['Integer', 'BigInt', 'Decimal', 'Double', 'Money'].includes(attr.AttributeType);
    });
  } catch (e) {
    console.warn('[EF PPT] Could not read the flowrun column types:', e);
  }
  _runMetaCache = meta;
  return meta;
}

// ─── Solutions ─────────────────────────────────────────────────────────────────

function _loadFavourites() {
  try {
    const raw = localStorage.getItem(FAV_SOL_KEY);
    if (raw) {
      _favSolutionIds = JSON.parse(raw).filter(id => typeof id === 'string');
      return;
    }
    const legacy = localStorage.getItem(DEFAULT_SOL_KEY);
    _favSolutionIds = legacy ? [legacy] : [];
    if (legacy) localStorage.setItem(FAV_SOL_KEY, JSON.stringify(_favSolutionIds));
  } catch {
    _favSolutionIds = [];
  }
}

function _saveFavourites() {
  try { localStorage.setItem(FAV_SOL_KEY, JSON.stringify(_favSolutionIds)); } catch { }
}

function _solName(sol) {
  return sol.friendlyname || sol.uniquename || sol.solutionid;
}

async function _loadSolutions() {
  const btn = document.getElementById('sol-ms-btn');
  btn.disabled = true;
  _setSolutionLabel('Loading solutions…');

  try {
    const url = `${ENV_URL}/api/data/v9.2/solutions` +
      `?$select=solutionid,uniquename,friendlyname,ismanaged` +
      `&$orderby=ismanaged asc,friendlyname asc`;

    _solutions = await _fetchAllPages(url);
    btn.disabled = false;

    const known = new Set(_solutions.map(s => s.solutionid));
    _selectedSolutionIds = _favSolutionIds.filter(id => known.has(id));
    _renderSolutionLabel();

    if (_selectedSolutionIds.length) {
      await _loadFlows();
    } else {
      _showPanelMsg('flow-list', 'Choose one or more solutions above. Star (☆) a solution to load it automatically next time.', 'info');
      _showPanelMsg('exec-list', 'Load flows to see executions.', 'info');
      _renderSummaryEmpty('Choose one or more solutions above to see a summary.');
    }
  } catch (e) {
    _setSolutionLabel('Failed to load solutions');
    console.error('[EF PPT]', e);
    _showPanelMsg('flow-list', `Error loading solutions: ${e.message}`, 'error');
  }
}

function _setSolutionLabel(text, placeholder = true) {
  const label = document.getElementById('sol-ms-label');
  label.textContent = text;
  label.classList.toggle('sol-ms-label--placeholder', placeholder);
}

function _renderSolutionLabel() {
  const names = _selectedSolutionIds
    .map(id => _solutions.find(s => s.solutionid === id))
    .filter(Boolean)
    .map(_solName);
  if (!names.length) _setSolutionLabel('Choose solutions…');
  else _setSolutionLabel(names.length > 1 ? `${names.length} solutions: ${names.join(', ')}` : names[0], false);
  document.getElementById('sol-ms-btn').title = names.join('\n');
}

function _bindSolutionPicker() {
  const btn   = document.getElementById('sol-ms-btn');
  const panel = document.getElementById('sol-ms-panel');
  const list  = document.getElementById('sol-ms-list');

  btn.addEventListener('click', () => {
    if (panel.classList.contains('hidden')) _openSolutionPicker();
    else _closeSolutionPicker(true);
  });
  document.getElementById('sol-ms-search').addEventListener('input', _renderSolutionList);
  document.getElementById('sol-ms-apply').addEventListener('click', () => _closeSolutionPicker(true));
  document.getElementById('sol-ms-clear').addEventListener('click', () => {
    _pendingSolutionIds = new Set();
    _renderSolutionList();
  });
  list.addEventListener('change', e => {
    const cb = e.target.closest('input[data-sol]');
    if (!cb) return;
    if (cb.checked) _pendingSolutionIds.add(cb.dataset.sol);
    else _pendingSolutionIds.delete(cb.dataset.sol);
  });
  list.addEventListener('click', e => {
    const star = e.target.closest('[data-fav]');
    if (!star) return;
    e.preventDefault();
    const id = star.dataset.fav;
    const i = _favSolutionIds.indexOf(id);
    if (i === -1) {
      _favSolutionIds.push(id);
      _pendingSolutionIds.add(id);
    } else {
      _favSolutionIds.splice(i, 1);
    }
    _saveFavourites();
    _renderSolutionList();
  });
  document.addEventListener('mousedown', e => {
    if (!panel.classList.contains('hidden') && !document.getElementById('sol-ms').contains(e.target)) _closeSolutionPicker(true);
  });
}

function _openSolutionPicker() {
  _pendingSolutionIds = new Set(_selectedSolutionIds);
  document.getElementById('sol-ms-search').value = '';
  document.getElementById('sol-ms-panel').classList.remove('hidden');
  document.getElementById('sol-ms').classList.add('open');
  document.getElementById('sol-ms-btn').setAttribute('aria-expanded', 'true');
  _renderSolutionList();
  document.getElementById('sol-ms-search').focus();
}

function _closeSolutionPicker(apply) {
  const panel = document.getElementById('sol-ms-panel');
  if (panel.classList.contains('hidden')) return;
  panel.classList.add('hidden');
  document.getElementById('sol-ms').classList.remove('open');
  document.getElementById('sol-ms-btn').setAttribute('aria-expanded', 'false');
  if (!apply || !_pendingSolutionIds) return;
  const next = _solutions.map(s => s.solutionid).filter(id => _pendingSolutionIds.has(id));
  _pendingSolutionIds = null;
  if (next.join(',') === _selectedSolutionIds.join(',')) return;
  _selectedSolutionIds = next;
  _renderSolutionLabel();
  if (next.length) {
    _loadFlows();
  } else {
    _flowsToken++;
    _setFlows([]);
    _renderFlows();
    document.getElementById('flow-count').textContent = '';
    _showPanelMsg('flow-list', 'Choose one or more solutions above.', 'info');
    _showPanelMsg('exec-list', 'Load flows to see executions.', 'info');
    _setExecCount('');
    _renderSummaryEmpty('Choose one or more solutions above to see a summary.');
  }
}

function _renderSolutionList() {
  const list = document.getElementById('sol-ms-list');
  const q = document.getElementById('sol-ms-search').value.trim().toLowerCase();
  const fav = new Set(_favSolutionIds);
  const matches = _solutions.filter(s => !q || _solName(s).toLowerCase().includes(q) || (s.uniquename ?? '').toLowerCase().includes(q));
  const ordered = matches.filter(s => fav.has(s.solutionid)).concat(matches.filter(s => !fav.has(s.solutionid)));
  if (!ordered.length) {
    list.innerHTML = '<li class="sol-ms-empty">No matching solutions.</li>';
    return;
  }
  list.innerHTML = ordered.map(s => {
    const isFav = fav.has(s.solutionid);
    return (
      `<li class="sol-ms-opt${isFav ? ' sol-ms-opt--fav' : ''}">` +
        `<label>` +
          `<input type="checkbox" data-sol="${_esc(s.solutionid)}"${_pendingSolutionIds.has(s.solutionid) ? ' checked' : ''}>` +
          `<span class="sol-ms-name">${_esc(_solName(s))}</span>` +
          (s.ismanaged ? '<span class="sol-ms-tag">Managed</span>' : '') +
        `</label>` +
        `<button type="button" class="sol-star${isFav ? ' sol-star--on' : ''}" data-fav="${_esc(s.solutionid)}" ` +
          `title="${isFav ? 'Remove from favourites' : 'Add to favourites — loads automatically next time'}" aria-pressed="${isFav}">${isFav ? '★' : '☆'}</button>` +
      `</li>`
    );
  }).join('');
}

// ─── Flows ─────────────────────────────────────────────────────────────────────

async function _loadFlows(force = false) {
  if (!_selectedSolutionIds.length) return;
  const token = ++_flowsToken;

  _setFlows([]);
  _selectedFlowId = null;
  _summaryKey    = null;
  document.getElementById('exec-scope').textContent = '';
  document.getElementById('btn-clear-flow-filter').classList.add('hidden');

  _showPanelMsg('flow-list', 'Loading flows…', 'loading');
  _showPanelMsg('exec-list', 'Loading executions…', 'loading');
  if (_view === 'summary') _renderSummaryEmpty('Loading flows…', 'loading');

  try {
    // Step 1: get workflow component IDs for this solution (componenttype 29 = workflow/flow)
    const solFilter = _selectedSolutionIds.map(id => `_solutionid_value eq ${id}`).join(' or ');
    const compUrl = `${ENV_URL}/api/data/v9.2/solutioncomponents` +
      `?$select=objectid` +
      `&$filter=${encodeURIComponent(`(${solFilter}) and componenttype eq 29`)}`;
    const components = await _fetchAllPages(compUrl);
    if (token !== _flowsToken) return;
    const flowIdSet  = new Set(components.map(c => (c.objectid ?? '').toLowerCase()));

    if (flowIdSet.size === 0) {
      document.getElementById('flow-count').textContent = '0';
      _showPanelMsg('flow-list', 'No flows found in the selected solutions.', 'info');
      _showPanelMsg('exec-list', 'No flows in the selected solutions.', 'info');
      _renderSummaryEmpty('No flows in the selected solutions.');
      return;
    }

    // Step 2: fetch all modern flows (category = 5) and filter to this solution's set
    const flowsUrl = `${ENV_URL}/api/data/v9.2/workflows` +
      `?$select=workflowid,name,statecode,statuscode,category,createdon,modifiedon` +
      `&$filter=category eq 5` +
      `&$orderby=name asc`;
    const allFlows = await _fetchAllPages(flowsUrl);
    if (token !== _flowsToken) return;
    _setFlows(allFlows.filter(f => flowIdSet.has((f.workflowid ?? '').toLowerCase())));

    _renderFlows();
    _loadExecutions();
    if (_view === 'summary') _loadSummary(force);
  } catch (e) {
    if (token !== _flowsToken) return;
    console.error('[EF PPT]', e);
    _showPanelMsg('flow-list', `Error: ${e.message}`, 'error');
    _showPanelMsg('exec-list', `Error: ${e.message}`, 'error');
    _renderSummaryEmpty(`Error: ${e.message}`, 'error');
  }
}

function _flowIds() {
  return _allFlows.map(f => (f.workflowid ?? '').toLowerCase()).filter(Boolean);
}

function _flowNameMap() {
  return new Map(_allFlows.map(f => [(f.workflowid ?? '').toLowerCase(), f.name ?? '(Unknown)']));
}

function _setFlows(flows) {
  _allFlows = flows;
  _flowIdSet = new Set(_flowIds());
}

function _inScope(run) {
  return _flowIdSet.has((run.workflowid ?? '').toLowerCase());
}

const STATUS_VALUES = {
  succeeded: ['Succeeded'],
  failed:    ['Failed'],
  running:   ['Running'],
  waiting:   ['Waiting', 'Suspended'],
  canceled:  ['Cancelled', 'Canceled'],
};

const SORT_COLUMNS = { date: 'starttime', duration: 'duration', status: 'status', error: 'errorcode' };

function _str(v) {
  return `'${String(v).replace(/'/g, "''")}'`;
}

function _idClause(meta, ids) {
  return `(${ids.map(id => `workflowid eq ${meta.idText ? _str(id) : id}`).join(' or ')})`;
}

function _statusClause(status) {
  const values = STATUS_VALUES[status];
  return values ? `(${values.map(v => `status eq ${_str(v)}`).join(' or ')})` : null;
}

function _orderBy(sortKey, sortDir) {
  const col = SORT_COLUMNS[sortKey] ?? 'starttime';
  const dir = sortKey in SORT_COLUMNS ? (sortDir === 1 ? 'asc' : 'desc') : 'desc';
  return col === 'starttime' ? `starttime ${dir}` : `${col} ${dir},starttime desc`;
}

function _runsUrl(clauses, order, select = RUN_SELECT) {
  return `${ENV_URL}/api/data/v9.2/flowruns` +
    `?$select=${select}` +
    `&$filter=${encodeURIComponent(clauses.filter(Boolean).join(' and '))}` +
    `&$orderby=${encodeURIComponent(order)}`;
}

function _runUrl(run) {
  return _orgId ? `https://make.powerautomate.com/environments/${_orgId}/flows/${run.workflowid}/runs/${run.name}` : '';
}

function _flowUrl(flowId) {
  return _orgId
    ? `https://make.powerautomate.com/environments/${_orgId}/flows/${flowId}`
    : `${ENV_URL}/main.aspx?pagetype=entityrecord&etn=workflow&id=${flowId}`;
}

function _createRunList(opts) {
  const st = {
    pages: [], next: [], page: 0, firstUrl: null, keep: null, loading: false, token: 0,
    sortKey: opts.sort.key, sortDir: opts.sort.dir, selectedId: null,
  };
  const listEl  = () => document.getElementById(opts.listId);
  const pagerEl = () => document.getElementById(opts.pagerId);

  pagerEl().addEventListener('click', e => {
    const btn = e.target.closest('[data-page]');
    if (!btn || btn.disabled) return;
    loadPage(btn.dataset.page === 'first' ? 0 : btn.dataset.page === 'prev' ? st.page - 1 : st.page + 1);
  });

  async function reload() {
    const token = ++st.token;
    st.pages = [];
    st.next = [];
    st.page = 0;
    st.loading = false;
    renderPager(false);
    opts.onCount('');
    _showPanelMsg(opts.listId, 'Loading runs…', 'loading');

    const meta = await _runMeta();
    if (token !== st.token) return;
    const q = opts.buildQuery(meta, st.sortKey, st.sortDir);
    if (q.message) {
      _showPanelMsg(opts.listId, q.message, q.type || 'info');
      return;
    }
    st.firstUrl = q.url;
    st.keep = q.keep || null;
    await loadPage(0, token);
  }

  async function loadPage(index, token = st.token) {
    if (index < 0 || st.loading) return;
    if (st.pages[index]) {
      st.page = index;
      render();
      return;
    }
    const url = index === 0 ? st.firstUrl : st.next[index - 1];
    if (!url) return;

    st.loading = true;
    if (index > 0) renderPager(true);
    let data;
    try {
      data = await _d365Fetch(url, _runPrefer(RUN_PAGE_SIZE));
    } catch (e) {
      if (token !== st.token) return;
      st.loading = false;
      if (index === 0 && opts.onFirstPageError && opts.onFirstPageError(e)) {
        console.warn('[EF PPT] Retrying the run query without error-text search:', e);
        reload();
        return;
      }
      console.error('[EF PPT]', e);
      if (index === 0) _showPanelMsg(opts.listId, `Error loading runs: ${e.message}`, 'error');
      else { renderPager(false); opts.onCount(`Couldn't load page ${index + 1}`); }
      return;
    }
    if (token !== st.token) return;
    st.loading = false;

    st.pages[index] = (data.value ?? []).slice(0, RUN_PAGE_SIZE).filter(run => _inScope(run) && (!st.keep || st.keep(run)));
    st.next[index] = data['@odata.nextLink'] ?? null;
    st.page = index;
    render();
  }

  function renderPager(busy) {
    const pager = pagerEl();
    const hasNext = !!st.next[st.page];
    const show = busy || st.page > 0 || hasNext;
    pager.classList.toggle('hidden', !show);
    if (!show) return;
    pager.innerHTML = `
      <button class="exec-pager__btn" data-page="first" title="First page" ${busy || st.page === 0 ? 'disabled' : ''}>&laquo;</button>
      <button class="exec-pager__btn" data-page="prev" ${busy || st.page === 0 ? 'disabled' : ''}>&lsaquo; Prev</button>
      <span class="exec-pager__label">${busy ? 'Loading…' : `Page ${st.page + 1}`}</span>
      <button class="exec-pager__btn" data-page="next" ${busy || !hasNext ? 'disabled' : ''}>Next &rsaquo;</button>
    `;
  }

  function sortIcon(key) {
    if (st.sortKey !== key) return '<span class="exec-sort-icon">⇅</span>';
    return `<span class="exec-sort-icon exec-sort-icon--active">${st.sortDir === 1 ? '▲' : '▼'}</span>`;
  }

  function render() {
    const flowMap = _flowNameMap();
    const rows = (st.pages[st.page] ?? []).slice();
    const from = st.page * RUN_PAGE_SIZE + 1;
    const more = !!st.next[st.page];

    opts.onCount(
      rows.length ? `${from.toLocaleString()}–${(from + rows.length - 1).toLocaleString()}${more ? '+' : ''}` : '0',
      more ? 'More runs on the next page' : ''
    );
    renderPager(false);

    const list = listEl();
    if (rows.length === 0) {
      list.innerHTML = `<div class="state-msg state-info">${_esc(opts.emptyText())}</div>`;
      return;
    }

    if (st.sortKey === 'flow') {
      rows.sort((a, b) => {
        const av = (flowMap.get((a.workflowid ?? '').toLowerCase()) ?? '').toLowerCase();
        const bv = (flowMap.get((b.workflowid ?? '').toLowerCase()) ?? '').toLowerCase();
        return av < bv ? -st.sortDir : av > bv ? st.sortDir : 0;
      });
    }

    list.innerHTML = '';
    const table = document.createElement('table');
    table.className = 'exec-table';
    table.innerHTML = `
      <colgroup>
        <col><col><col><col><col><col>
      </colgroup>
      <thead><tr>
        <th class="exec-th" data-sort="status">Status ${sortIcon('status')}</th>
        <th class="exec-th" data-sort="flow" title="Sorts the runs on this page">Flow ${sortIcon('flow')}</th>
        <th class="exec-th" data-sort="date">Date ${sortIcon('date')}</th>
        <th class="exec-th" data-sort="duration">Duration ${sortIcon('duration')}</th>
        <th class="exec-th" data-sort="error">Error ${sortIcon('error')}</th>
        <th class="exec-th exec-th--static">Run</th>
      </tr></thead>
    `;

    table.addEventListener('click', e => {
      const th = e.target.closest('[data-sort]');
      if (!th) return;
      const key = th.dataset.sort;
      if (st.sortKey === key) {
        st.sortDir *= -1;
      } else {
        st.sortKey = key;
        st.sortDir = key === 'date' ? -1 : 1;
      }
      if (key === 'flow') render();
      else reload();
    });

    const tbody = document.createElement('tbody');
    rows.forEach(exec => {
      const status     = _execStatus(exec);
      const flowName   = flowMap.get((exec.workflowid ?? '').toLowerCase()) ?? '—';
      const errFull    = exec.errorcode ?? '';
      const errPreview = errFull.length > 90 ? errFull.slice(0, 90) + '…' : errFull;
      const runUrl     = _runUrl(exec);

      const tr = document.createElement('tr');
      if (exec.name === st.selectedId) tr.classList.add('selected');
      tr.innerHTML = `
        <td><span class="exec-badge exec-badge--${status.cls}">${_esc(status.label)}</span></td>
        <td class="exec-flow-name" title="${_esc(flowName)}">${_esc(flowName)}</td>
        <td class="exec-date">${_fmtDate(exec.starttime)}</td>
        <td class="exec-duration">${_fmtDurationMs(_execDurationMs(exec))}</td>
        <td class="exec-error" title="${_esc(errFull)}">${_esc(errPreview)}</td>
        <td class="exec-open">${runUrl
          ? `<a class="exec-run-link" href="${runUrl}" target="_blank" rel="noopener noreferrer" title="Open this run in Power Automate">Open &#8599;</a>`
          : ''}</td>
      `;
      tr.addEventListener('click', e => {
        if (e.target.closest('a')) return;
        st.selectedId = exec.name === st.selectedId ? null : exec.name;
        tbody.querySelectorAll('tr').forEach(r => r.classList.remove('selected'));
        if (st.selectedId) tr.classList.add('selected');
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    list.appendChild(table);
    list.scrollTop = 0;
  }

  return {
    reload,
    cancel() { st.token++; st.loading = false; },
    setSort(key, dir) { st.sortKey = key; st.sortDir = dir; },
  };
}

// ─── Executions ────────────────────────────────────────────────────────────────

function _queueLoadExecutions() {
  clearTimeout(_execReloadTimer);
  _execReloadTimer = setTimeout(_loadExecutions, 400);
}

function _loadExecutions() {
  clearTimeout(_execReloadTimer);
  _execList.reload();
}

function _durationFilterMs() {
  const op   = document.getElementById('dur-op').value;
  const raw  = parseFloat(document.getElementById('dur-value').value);
  const unit = document.getElementById('dur-unit').value;
  if (!op || isNaN(raw)) return null;
  const ms = unit === 'ms' ? raw : unit === 's' ? raw * 1_000 : unit === 'min' ? raw * 60_000 : raw * 3_600_000;
  return { op, ms: Math.round(ms) };
}

function _execQuery(meta, sortKey, sortDir) {
  if (_allFlows.length === 0) return { message: 'No flows loaded.' };
  if (!_orgId) {
    return {
      message: 'Power Apps Environment ID is not configured for this environment. ' +
        'Open Settings → Edit environment and fill in the "Power Apps Environment ID" field.',
      type: 'error',
    };
  }

  const flowMap = new Map(_allFlows.map(f => [(f.workflowid ?? '').toLowerCase(), (f.name ?? '').toLowerCase()]));
  let ids = _selectedFlowId ? [_selectedFlowId.toLowerCase()] : _flowIds();
  const clauses = [];

  const query = document.getElementById('exec-search').value.trim();
  let searchClause = null;
  if (query) {
    const q = query.toLowerCase();
    const nameIds = ids.filter(id => (flowMap.get(id) ?? '').includes(q));
    const parts = [];
    if (nameIds.length) parts.push(_idClause(meta, nameIds));
    if (!_runErrorSearchOff) parts.push(`contains(errorcode,${_str(query)})`);
    if (!parts.length) return { message: 'No flows or errors match the search.' };
    if (_runErrorSearchOff) ids = nameIds;
    else searchClause = parts.length > 1 ? `(${parts.join(' or ')})` : parts[0];
  }

  if (ids.length > RUN_MAX_FLOWS) {
    return { message: `The selected solutions have ${ids.length} flows — too many to query together. Select a flow on the left, or search by flow name.` };
  }
  clauses.push(_idClause(meta, ids));
  if (searchClause) clauses.push(searchClause);

  const fromVal = document.getElementById('exec-from').value;
  const toVal   = document.getElementById('exec-to').value;
  if (fromVal) clauses.push(`starttime ge ${new Date(fromVal).toISOString()}`);
  if (toVal)   clauses.push(`starttime le ${new Date(toVal).toISOString()}`);

  clauses.push(_statusClause(document.getElementById('exec-status-filter').value));

  const dur = _durationFilterMs();
  if (dur && meta.durationNumeric) clauses.push(`duration ${dur.op} ${dur.ms}`);
  const keep = dur && !meta.durationNumeric
    ? run => { const ms = _execDurationMs(run); return ms !== null && (dur.op === 'gt' ? ms > dur.ms : ms < dur.ms); }
    : null;

  return { url: _runsUrl(clauses, _orderBy(sortKey, sortDir)), keep };
}

function _setExecCount(text, title = '') {
  const el = document.getElementById('exec-count');
  el.textContent = text;
  el.title = title;
}

// ─── Render flows ──────────────────────────────────────────────────────────────

function _renderFlows() {
  const query        = document.getElementById('flow-search').value.trim().toLowerCase();
  const statusFilter = document.querySelector('#flow-status-filter .seg-btn.active')?.dataset.status ?? 'all';

  _filteredFlows = _allFlows.filter(flow => {
    const st = _flowStatus(flow);
    if (statusFilter === 'active'   && st.cls !== 'active')   return false;
    if (statusFilter === 'inactive' && st.cls === 'active')   return false;
    if (query && !(flow.name ?? '').toLowerCase().includes(query)) return false;
    return true;
  });

  document.getElementById('flow-count').textContent = `${_filteredFlows.length}`;

  const list = document.getElementById('flow-list');
  if (_filteredFlows.length === 0) {
    list.innerHTML = `<div class="state-msg state-info">${
      query || statusFilter !== 'all' ? 'No matching flows.' : 'No flows found.'
    }</div>`;
    return;
  }

  list.innerHTML = '';
  _filteredFlows.forEach(flow => {
    const st      = _flowStatus(flow);
    const openUrl = _flowUrl(flow.workflowid);

    const item = document.createElement('div');
    item.className  = `flow-item${flow.workflowid === _selectedFlowId ? ' selected' : ''}`;
    item.dataset.id = flow.workflowid;
    item.title      = flow.name ?? '';
    item.innerHTML  = `
      <span class="flow-status-dot flow-status-dot--${st.cls}" title="${_esc(st.label)}"></span>
      <div class="flow-item__info">
        <div class="flow-item__name">${_esc(flow.name ?? '(Unnamed)')}</div>
        <div class="flow-item__meta">Modified ${_fmtDate(flow.modifiedon)}</div>
      </div>
      <a class="flow-item__link" href="${openUrl}" target="_blank" rel="noopener noreferrer" title="Open in Power Automate">&#8599;</a>
    `;

    item.addEventListener('click', e => {
      if (e.target.closest('.flow-item__link')) return;
      const wasSelected = _selectedFlowId === flow.workflowid;
      _selectedFlowId   = wasSelected ? null : flow.workflowid;

      document.querySelectorAll('.flow-item').forEach(el => el.classList.remove('selected'));
      if (!wasSelected) item.classList.add('selected');

      const scopeEl  = document.getElementById('exec-scope');
      const clearBtn = document.getElementById('btn-clear-flow-filter');
      if (_selectedFlowId) {
        scopeEl.textContent = flow.name ?? '';
        clearBtn.classList.remove('hidden');
      } else {
        scopeEl.textContent = '';
        clearBtn.classList.add('hidden');
      }
      _loadExecutions();
    });

    list.appendChild(item);
  });
}

function _periodLabel() {
  return _summaryPeriod === '1' ? 'last 24 hours' : `last ${_summaryPeriod} days`;
}

function _summaryFrom() {
  return new Date(Date.now() - Number(_summaryPeriod) * 86_400_000).toISOString();
}

function _renderSummaryEmpty(msg, type = 'info') {
  _summaryKey = null;
  _summaryData = null;
  _summaryToken++;
  _closePane();
  document.getElementById('sum-scope').textContent = '';
  document.getElementById('sum-updated').textContent = '';
  document.getElementById('sum-content').innerHTML = `<div class="state-msg state-${type}">${_esc(msg)}</div>`;
}

async function _loadSummary(force = false) {
  if (!_selectedSolutionIds.length || !_allFlows.length) return;
  const ids = _flowIds();
  const key = `${_summaryPeriod}|${ids.join(',')}`;
  if (!force && key === _summaryKey) return;
  _summaryKey = key;
  const token = ++_summaryToken;
  _closePane();

  const solCount = _selectedSolutionIds.length;
  document.getElementById('sum-scope').textContent =
    `${ids.length} flow${ids.length === 1 ? '' : 's'} in ${solCount} solution${solCount === 1 ? '' : 's'} · ${_periodLabel()}`;
  document.getElementById('sum-updated').textContent = '';

  if (ids.length > RUN_MAX_FLOWS) {
    _summaryData = null;
    document.getElementById('sum-content').innerHTML =
      `<div class="state-msg state-info">The selected solutions have ${ids.length} flows — too many to summarise together. Choose fewer solutions.</div>`;
    return;
  }

  const data = _summaryData = {
    total: null, totalCapped: false, totalErr: false,
    failures: [], failuresStarted: false, failuresDone: false, failuresCapped: false, failuresErr: null,
    longest: null, longestErr: null,
    running: null, runningMore: false, runningErr: false,
  };
  _renderSummary();

  const meta = await _runMeta();
  if (token !== _summaryToken) return;
  const base = [_idClause(meta, ids), `starttime ge ${_summaryFrom()}`];
  const update = () => { if (token === _summaryToken) _renderSummary(); };

  const failures = (async () => {
    let url = _runsUrl([...base, _statusClause('failed')], 'starttime desc');
    let pages = 0;
    try {
      while (url && pages < 10) {
        const d = await _d365Fetch(url, _runPrefer(500));
        if (token !== _summaryToken) return;
        data.failures.push(...(d.value ?? []).filter(_inScope));
        data.failuresStarted = true;
        url = d['@odata.nextLink'] ?? null;
        pages++;
        data.failuresCapped = !!url && pages >= 10;
        update();
      }
    } catch (e) {
      data.failuresErr = e.message;
    }
    data.failuresStarted = true;
    data.failuresDone = true;
    update();
  })();

  const longest = (async () => {
    if (!meta.durationNumeric) { data.longestErr = 'Run durations aren’t available in this environment.'; update(); return; }
    let remaining = ids.slice();
    const seen = new Set();
    try {
      for (let round = 0; round < 5 && remaining.length && seen.size < 8; round++) {
        const d = await _d365Fetch(
          _runsUrl([_idClause(meta, remaining), base[1], 'duration ne null'], 'duration desc'),
          _runPrefer(200)
        );
        if (token !== _summaryToken) return;
        const rows = (d.value ?? []).filter(_inScope);
        data.longest = (data.longest ?? []).concat(rows);
        rows.forEach(run => seen.add((run.workflowid ?? '').toLowerCase()));
        remaining = remaining.filter(id => !seen.has(id));
        update();
        if (!rows.length || !d['@odata.nextLink']) break;
      }
      if (!data.longest) data.longest = [];
    } catch (e) {
      if (!data.longest) data.longestErr = e.message;
    }
    update();
  })();

  const total = (async () => {
    try {
      const d = await _d365Fetch(_runsUrl(base, 'starttime desc', 'name') + '&$count=true', {
        'Prefer': 'odata.include-annotations="*",odata.maxpagesize=1',
      });
      if (token !== _summaryToken) return;
      data.total = typeof d['@odata.count'] === 'number' ? d['@odata.count'] : null;
      data.totalCapped = d['@Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded'] === true;
      data.totalErr = data.total == null;
    } catch {
      data.totalErr = true;
    }
    update();
  })();

  const running = (async () => {
    try {
      const d = await _d365Fetch(
        _runsUrl([_idClause(meta, ids), _statusClause('running')], 'starttime desc', 'name,workflowid'),
        _runPrefer(500)
      );
      if (token !== _summaryToken) return;
      data.running = (d.value ?? []).filter(_inScope).length;
      data.runningMore = !!d['@odata.nextLink'];
    } catch {
      data.runningErr = true;
    }
    update();
  })();

  await Promise.all([failures, longest, total, running]);
  if (token !== _summaryToken) return;
  const now = new Date();
  document.getElementById('sum-updated').textContent =
    `Updated ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
}

function _groupByFlow(runs, flowMap) {
  const groups = new Map();
  runs.forEach(run => {
    const id = (run.workflowid ?? '').toLowerCase();
    if (!groups.has(id)) groups.set(id, { id, name: flowMap.get(id) ?? '(Unknown flow)', runs: [] });
    groups.get(id).runs.push(run);
  });
  return [...groups.values()];
}

function _flowButton(kind, id, name) {
  const active = _paneQuery && _paneQuery.kind === kind && _paneQuery.flowId === id;
  return `<button type="button" class="sum-flow${active ? ' is-active' : ''}" data-open="${kind}" data-flow="${_esc(id)}" title="${_esc(name)}">${_esc(name)}</button>`;
}

function _kpi(kind, label, value, sub, tone = '') {
  const active = _paneQuery && _paneQuery.kind === kind && !_paneQuery.flowId;
  return (
    `<button type="button" class="sum-kpi${tone ? ` sum-kpi--${tone}` : ''}${active ? ' is-active' : ''}" data-open="${kind}">` +
      `<span class="sum-kpi-label">${_esc(label)}</span>` +
      `<span class="sum-kpi-value">${value}</span>` +
      `<span class="sum-kpi-sub">${sub}</span>` +
    `</button>`
  );
}

function _loadingLine(text = 'Loading…') {
  return `<div class="state-msg state-loading">${_esc(text)}</div>`;
}

function _renderSummary() {
  const data = _summaryData;
  if (!data) return;
  const flowMap = _flowNameMap();
  const content = document.getElementById('sum-content');
  if (!content.querySelector('.sum-kpis')) {
    content.innerHTML = `
      <div class="sum-kpis" id="sum-kpis"></div>
      <div class="sum-grid">
        <section class="sum-card">
          <header class="sum-card-head"><span>Longest-running flows</span><small>slowest run per flow</small></header>
          <div class="sum-card-body" id="sum-longest"></div>
        </section>
        <section class="sum-card">
          <header class="sum-card-head"><span>Most failures</span><small>failed runs per flow</small></header>
          <div class="sum-card-body" id="sum-failures"></div>
        </section>
        <section class="sum-card sum-card--wide">
          <header class="sum-card-head"><span>Most recent failures</span><small>click a flow to see all its failures</small></header>
          <div class="sum-card-body" id="sum-recent"></div>
        </section>
      </div>`;
  }

  const failCount = data.failures.length;
  const failText = !data.failuresStarted ? '…' : data.failuresErr && !failCount ? '—' : `${failCount.toLocaleString()}${data.failuresCapped ? '+' : ''}`;
  const totalText = data.totalErr ? '—' : data.total == null ? '…' : `${data.total.toLocaleString()}${data.totalCapped ? '+' : ''}`;
  const rate = data.failuresDone && !data.failuresCapped && !data.totalErr && data.total && !data.totalCapped
    ? `${((failCount / data.total) * 100).toFixed(failCount / data.total < 0.1 ? 1 : 0)}% of runs`
    : data.failuresDone ? _esc(_periodLabel()) : 'counting…';
  const failingFlows = new Set(data.failures.map(r => (r.workflowid ?? '').toLowerCase())).size;
  const runningText = data.runningErr ? '—' : data.running == null ? '…' : `${data.running.toLocaleString()}${data.runningMore ? '+' : ''}`;
  const top = data.longest && data.longest[0];
  const longestText = data.longestErr ? '—' : data.longest == null ? '…' : top ? _esc(_fmtDurationMs(_execDurationMs(top))) : '—';
  const longestSub = top ? _esc(flowMap.get((top.workflowid ?? '').toLowerCase()) ?? '') : data.longest ? 'no completed runs' : '';

  document.getElementById('sum-kpis').innerHTML =
    _kpi('all', 'Runs', totalText, _esc(_periodLabel())) +
    _kpi('failed', 'Failed runs', failText, rate, failCount ? 'bad' : '') +
    _kpi('failed', 'Flows with failures', !data.failuresStarted ? '…' : failingFlows.toLocaleString(), `of ${_allFlows.length} flows`) +
    _kpi('running', 'Running now', runningText, 'in progress', data.running ? 'info' : '') +
    _kpi('longest', 'Longest run', longestText, longestSub);

  const longestEl = document.getElementById('sum-longest');
  if (data.longestErr) longestEl.innerHTML = `<div class="state-msg state-info">${_esc(data.longestErr)}</div>`;
  else if (!data.longest) longestEl.innerHTML = _loadingLine();
  else {
    const groups = _groupByFlow(data.longest, flowMap)
      .map(g => ({ ...g, max: Math.max(...g.runs.map(r => _execDurationMs(r) ?? 0)) }))
      .sort((a, b) => b.max - a.max)
      .slice(0, 8);
    const peak = groups[0] ? groups[0].max : 0;
    longestEl.innerHTML = groups.length
      ? `<ul class="sum-bars">${groups.map(g => `
          <li>
            ${_flowButton('longest', g.id, g.name)}
            <span class="sum-bar"><span style="width:${peak ? Math.max(4, (g.max / peak) * 100) : 0}%"></span></span>
            <span class="sum-val">${_esc(_fmtDurationMs(g.max))}</span>
          </li>`).join('')}</ul>`
      : `<div class="state-msg state-info">No completed runs in the ${_esc(_periodLabel())}.</div>`;
  }

  const failuresEl = document.getElementById('sum-failures');
  const recentEl = document.getElementById('sum-recent');
  if (!data.failuresStarted) {
    failuresEl.innerHTML = _loadingLine();
    recentEl.innerHTML = _loadingLine();
  } else if (data.failuresErr && !failCount) {
    failuresEl.innerHTML = `<div class="state-msg state-error">${_esc(data.failuresErr)}</div>`;
    recentEl.innerHTML = '';
  } else if (!failCount) {
    const none = `<div class="state-msg state-ok">No failures in the ${_esc(_periodLabel())}.</div>`;
    failuresEl.innerHTML = none;
    recentEl.innerHTML = none;
  } else {
    const groups = _groupByFlow(data.failures, flowMap).sort((a, b) => b.runs.length - a.runs.length).slice(0, 8);
    const peak = groups[0].runs.length;
    failuresEl.innerHTML = `<ul class="sum-bars sum-bars--bad">${groups.map(g => `
        <li>
          ${_flowButton('failed', g.id, g.name)}
          <span class="sum-bar"><span style="width:${Math.max(4, (g.runs.length / peak) * 100)}%"></span></span>
          <span class="sum-val" title="Last failed ${_esc(_fmtDate(g.runs[0].starttime))}">${g.runs.length.toLocaleString()}</span>
        </li>`).join('')}</ul>` +
      (data.failuresDone ? '' : '<div class="sum-note">Still counting…</div>') +
      (data.failuresCapped ? '<div class="sum-note">Counted the newest 5,000 failures.</div>' : '');

    recentEl.innerHTML = `<table class="sum-recent">
        <colgroup><col style="width:130px"><col style="width:32%"><col><col style="width:64px"></colgroup>
        <tbody>${data.failures.slice(0, 12).map(run => {
          const id = (run.workflowid ?? '').toLowerCase();
          const err = run.errorcode ?? '';
          const url = _runUrl(run);
          return `<tr>
            <td class="exec-date">${_fmtDate(run.starttime)}</td>
            <td>${_flowButton('failed', id, flowMap.get(id) ?? '(Unknown flow)')}</td>
            <td class="exec-error" title="${_esc(err)}">${_esc(err.length > 110 ? err.slice(0, 110) + '…' : err)}</td>
            <td class="exec-open">${url ? `<a class="exec-run-link" href="${url}" target="_blank" rel="noopener noreferrer" title="Open this run in Power Automate">Open &#8599;</a>` : ''}</td>
          </tr>`;
        }).join('')}</tbody>
      </table>`;
  }
}

const PANE_KINDS = {
  all:     { title: 'Runs',          sort: { key: 'date', dir: -1 },     sub: 'newest first' },
  failed:  { title: 'Failed runs',   sort: { key: 'date', dir: -1 },     sub: 'newest first', status: 'failed' },
  running: { title: 'Running now',   sort: { key: 'date', dir: -1 },     sub: 'in progress', status: 'running', anyTime: true },
  longest: { title: 'Longest runs',  sort: { key: 'duration', dir: -1 }, sub: 'slowest first' },
};

function _openPane(kind, flowId) {
  const cfg = PANE_KINDS[kind];
  if (!cfg) return;
  _paneQuery = { kind, flowId: flowId ? flowId.toLowerCase() : null };
  const flowName = flowId ? (_flowNameMap().get(_paneQuery.flowId) ?? '(Unknown flow)') : null;
  const sort = kind === 'longest' && _runMetaCache && !_runMetaCache.durationNumeric ? { key: 'date', dir: -1 } : cfg.sort;

  document.getElementById('sum-pane-title').textContent = flowName ? `${cfg.title} — ${flowName}` : `${cfg.title} — all flows`;
  document.getElementById('sum-pane-title').title = flowName ?? '';
  document.getElementById('sum-pane-sub').textContent = cfg.anyTime ? cfg.sub : `${_periodLabel()} · ${cfg.sub}`;
  const link = document.getElementById('sum-pane-flow');
  link.classList.toggle('hidden', !flowId);
  if (flowId) link.href = _flowUrl(_paneQuery.flowId);

  document.getElementById('sum-pane').classList.remove('hidden');
  document.getElementById('summary-view').classList.add('summary--pane');
  _paneList.setSort(sort.key, sort.dir);
  _paneList.reload();
  _renderSummary();
}

function _closePane() {
  if (!_paneList) return;
  _paneList.cancel();
  _paneQuery = null;
  document.getElementById('sum-pane').classList.add('hidden');
  document.getElementById('summary-view').classList.remove('summary--pane');
  _renderSummary();
}

function _paneQueryUrl(meta, sortKey, sortDir) {
  if (!_paneQuery) return { message: 'Choose something on the dashboard.' };
  const cfg = PANE_KINDS[_paneQuery.kind];
  const ids = _paneQuery.flowId ? [_paneQuery.flowId] : _flowIds();
  if (ids.length > RUN_MAX_FLOWS) return { message: 'Too many flows to query together — click a single flow instead.' };
  const clauses = [_idClause(meta, ids)];
  if (!cfg.anyTime) clauses.push(`starttime ge ${_summaryFrom()}`);
  if (cfg.status) clauses.push(_statusClause(cfg.status));
  if (_paneQuery.kind === 'longest' && meta.durationNumeric && sortKey === 'duration') clauses.push('duration ne null');
  return { url: _runsUrl(clauses, _orderBy(sortKey, sortDir)) };
}

// ─── Status helpers ────────────────────────────────────────────────────────────

function _flowStatus(flow) {
  if (flow.statecode === 1) return { label: 'On',        cls: 'active'    };
  if (flow.statecode === 2) return { label: 'Suspended', cls: 'suspended' };
  return                           { label: 'Off',       cls: 'inactive'  };
}

function _execStatus(exec) {
  const s = (exec.status ?? '').toLowerCase();
  if (s === 'succeeded')                      return { label: 'Succeeded', cls: 'succeeded' };
  if (s === 'failed')                         return { label: 'Failed',    cls: 'failed'    };
  if (s === 'running')                        return { label: 'Running',   cls: 'running'   };
  if (s === 'waiting' || s === 'suspended')   return { label: 'Waiting',   cls: 'waiting'   };
  if (s === 'cancelled' || s === 'canceled')  return { label: 'Canceled',  cls: 'canceled'  };
  return { label: exec.status ?? '?', cls: 'unknown' };
}

// ─── Utility ───────────────────────────────────────────────────────────────────

function _showPanelMsg(panelId, msg, type = 'info') {
  document.getElementById(panelId).innerHTML =
    `<div class="state-msg state-${type}">${_esc(msg)}</div>`;
}

function _fmtDate(iso) {
  if (!iso) return '—';
  try {
    const d   = new Date(iso);
    const pad = n => String(n).padStart(2, '0');
    return `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  } catch { return iso; }
}

// Returns duration in ms for an execution record.
// Prefers the explicit `duration` field (available on some environments);
// falls back to modifiedon − createdon for completed runs.
function _execDurationMs(exec) {
  if (exec.duration != null) {
    const n = Number(exec.duration);
    if (!isNaN(n) && n >= 0) return n;
  }
  const status = (exec.status ?? '').toLowerCase();
  const start  = exec.starttime ?? exec.createdon;
  if (start && exec.modifiedon && status !== 'running') {
    const ms = new Date(exec.modifiedon) - new Date(start);
    if (ms > 0) return ms;
  }
  return null;
}

// Formats a raw millisecond value into a human-readable string.
function _fmtDurationMs(ms) {
  if (ms === null || ms === undefined) return '—';
  ms = Number(ms);
  if (isNaN(ms) || ms < 0) return '—';
  if (ms < 1000)            return `${ms} ms`;
  const s = Math.floor(ms / 1000);
  if (s < 60)               return `${s}s`;
  const m = Math.floor(s / 60), rs = s % 60;
  if (m < 60)               return rs > 0 ? `${m}m ${rs}s` : `${m}m`;
  const h = Math.floor(m / 60), rm = m % 60;
  return                           rm > 0 ? `${h}h ${rm}m` : `${h}h`;
}

function _toInputDateTime(d) {
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function _esc(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
