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
let _selectedSolution    = null;
let _allFlows            = [];
let _filteredFlows       = [];
let _selectedFlowId      = null;
let _allExecutions       = [];
let _filteredExecutions  = [];
const _orgId             = PA_ENV_ID; // Power Platform environment ID from extension settings
const RUN_PAGE_SIZE      = 100;
const RUN_MAX_FLOWS      = 300;
let _execToken           = 0;
let _runMetaCache        = null;
let _runPages            = [];
let _runNext             = [];
let _runPage             = 0;
let _runFirstUrl         = null;
let _runLoading          = false;
let _runErrorSearchOff   = false;
let _execReloadTimer     = null;

// Execution table sort state
let _execSortKey = 'date';
let _execSortDir = -1;   // -1 = descending, 1 = ascending

// Selected execution row
let _selectedExecId = null;

const DEFAULT_SOL_KEY = `flows_defaultSol_${ENV_URL}`;

// ─── Bootstrap ─────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {
  document.title = `Flows — ${ENV_NAME}`;
  document.getElementById('env-name').textContent = ENV_NAME;
  document.getElementById('env-url').textContent  = ENV_URL.replace('https://', '');

  _setDefaultDates();
  _bindControls();

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
  // Solution select
  document.getElementById('solution-select').addEventListener('change', async (e) => {
    _selectedSolution = _solutions.find(s => s.solutionid === e.target.value) ?? null;
    _updateDefaultStar();
    if (_selectedSolution) await _loadFlows();
  });

  // Set default solution
  document.getElementById('btn-set-default').addEventListener('click', () => {
    if (!_selectedSolution) return;
    localStorage.setItem(DEFAULT_SOL_KEY, _selectedSolution.solutionid);
    const btn = document.getElementById('btn-set-default');
    btn.textContent = '★';
    btn.classList.add('sol-icon-btn--active');
    btn.title = 'Default solution saved';
  });

  // Refresh
  document.getElementById('btn-refresh').addEventListener('click', async () => {
    if (_selectedSolution) await _loadFlows();
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
  document.getElementById('exec-pager').addEventListener('click', e => {
    const btn = e.target.closest('[data-page]');
    if (!btn || btn.disabled) return;
    _loadRunPage(btn.dataset.page === 'first' ? 0 : btn.dataset.page === 'prev' ? _runPage - 1 : _runPage + 1);
  });
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

function _updateDefaultStar() {
  const defaultId = localStorage.getItem(DEFAULT_SOL_KEY);
  const btn = document.getElementById('btn-set-default');
  const isDefault = _selectedSolution && defaultId === _selectedSolution.solutionid;
  btn.textContent = isDefault ? '★' : '☆';
  btn.title = isDefault ? 'This is the default solution' : 'Set as default solution';
  btn.classList.toggle('sol-icon-btn--active', isDefault);
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

async function _fetchAllPages(firstUrl, maxPages = 20, onPage = null) {
  const records = [];
  let nextUrl   = firstUrl;
  let page      = 0;
  while (nextUrl && page < maxPages) {
    page++;
    const data = await _d365Fetch(nextUrl, {
      'Prefer': 'odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=500',
    });
    const value = data.value ?? [];
    records.push(...value);
    if (onPage) onPage(value);
    nextUrl = data['@odata.nextLink'] ?? null;
  }
  records.truncated = !!nextUrl;
  return records;
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

async function _loadSolutions() {
  const select = document.getElementById('solution-select');
  select.innerHTML = '<option value="">Loading…</option>';
  select.disabled  = true;

  try {
    const url = `${ENV_URL}/api/data/v9.2/solutions` +
      `?$select=solutionid,uniquename,friendlyname,ismanaged` +
      `&$orderby=ismanaged asc,friendlyname asc`;

    _solutions = await _fetchAllPages(url);

    select.innerHTML = '<option value="">— Select a solution —</option>';
    _solutions.forEach(sol => {
      const opt     = document.createElement('option');
      opt.value     = sol.solutionid;
      opt.textContent = `${sol.friendlyname || sol.uniquename}${sol.ismanaged ? '  (Managed)' : ''}`;
      select.appendChild(opt);
    });
    select.disabled = false;

    // Restore default solution
    const defaultId = localStorage.getItem(DEFAULT_SOL_KEY);
    if (defaultId && _solutions.some(s => s.solutionid === defaultId)) {
      select.value      = defaultId;
      _selectedSolution = _solutions.find(s => s.solutionid === defaultId);
      _updateDefaultStar();
      await _loadFlows();
    } else {
      _showPanelMsg('flow-list', 'Select a solution to load flows.', 'info');
      _showPanelMsg('exec-list', 'Load flows to see executions.', 'info');
    }
  } catch (e) {
    select.innerHTML = `<option value="">Failed to load solutions</option>`;
    console.error('[EF PPT]', e);
    _showPanelMsg('flow-list', `Error loading solutions: ${e.message}`, 'error');
  }
}

// ─── Flows ─────────────────────────────────────────────────────────────────────

async function _loadFlows() {
  if (!_selectedSolution) return;

  _allFlows      = [];
  _selectedFlowId = null;
  document.getElementById('exec-scope').textContent = '';
  document.getElementById('btn-clear-flow-filter').classList.add('hidden');

  _showPanelMsg('flow-list', 'Loading flows…', 'loading');
  _showPanelMsg('exec-list', 'Loading executions…', 'loading');

  try {
    // Step 1: get workflow component IDs for this solution (componenttype 29 = workflow/flow)
    const compUrl = `${ENV_URL}/api/data/v9.2/solutioncomponents` +
      `?$select=objectid` +
      `&$filter=_solutionid_value eq ${_selectedSolution.solutionid} and componenttype eq 29`;
    const components = await _fetchAllPages(compUrl);
    const flowIdSet  = new Set(components.map(c => (c.objectid ?? '').toLowerCase()));

    if (flowIdSet.size === 0) {
      _showPanelMsg('flow-list', 'No flows found in this solution.', 'info');
      _showPanelMsg('exec-list', 'No flows in this solution.', 'info');
      return;
    }

    // Step 2: fetch all modern flows (category = 5) and filter to this solution's set
    const flowsUrl = `${ENV_URL}/api/data/v9.2/workflows` +
      `?$select=workflowid,name,statecode,statuscode,category,createdon,modifiedon` +
      `&$filter=category eq 5` +
      `&$orderby=name asc`;
    const allFlows = await _fetchAllPages(flowsUrl);
    _allFlows = allFlows.filter(f => flowIdSet.has((f.workflowid ?? '').toLowerCase()));

    _renderFlows();
    await _loadExecutions();
  } catch (e) {
    console.error('[EF PPT]', e);
    _showPanelMsg('flow-list', `Error: ${e.message}`, 'error');
    _showPanelMsg('exec-list', `Error: ${e.message}`, 'error');
  }
}


// ─── Executions ────────────────────────────────────────────────────────────────

function _queueLoadExecutions() {
  clearTimeout(_execReloadTimer);
  _execReloadTimer = setTimeout(_loadExecutions, 400);
}

function _durationFilterMs() {
  const op   = document.getElementById('dur-op').value;
  const raw  = parseFloat(document.getElementById('dur-value').value);
  const unit = document.getElementById('dur-unit').value;
  if (!op || isNaN(raw)) return null;
  const ms = unit === 'ms' ? raw : unit === 's' ? raw * 1_000 : unit === 'min' ? raw * 60_000 : raw * 3_600_000;
  return { op, ms: Math.round(ms) };
}

const STATUS_VALUES = {
  succeeded: ['Succeeded'],
  failed:    ['Failed'],
  running:   ['Running'],
  waiting:   ['Waiting', 'Suspended'],
  canceled:  ['Cancelled', 'Canceled'],
};

const SORT_COLUMNS = { date: 'starttime', duration: 'duration', status: 'status', error: 'errorcode' };

function _runQueryUrl(meta) {
  const quoteIds = meta.idText;
  const flowMap = new Map(_allFlows.map(f => [(f.workflowid ?? '').toLowerCase(), (f.name ?? '').toLowerCase()]));
  let ids = _selectedFlowId ? [_selectedFlowId.toLowerCase()] : [...flowMap.keys()].filter(Boolean);
  const str = v => `'${String(v).replace(/'/g, "''")}'`;
  const idEq = list => list.map(id => `workflowid eq ${quoteIds ? str(id) : id}`).join(' or ');
  const clauses = [];

  const query = document.getElementById('exec-search').value.trim();
  let searchClause = null;
  if (query) {
    const q = query.toLowerCase();
    const nameIds = ids.filter(id => (flowMap.get(id) ?? '').includes(q));
    const parts = [];
    if (nameIds.length) parts.push(`(${idEq(nameIds)})`);
    if (!_runErrorSearchOff) parts.push(`contains(errorcode,${str(query)})`);
    if (!parts.length) return { empty: true };
    if (_runErrorSearchOff) ids = nameIds;
    else searchClause = parts.length > 1 ? `(${parts.join(' or ')})` : parts[0];
  }

  if (ids.length > RUN_MAX_FLOWS) return { tooMany: ids.length };
  clauses.push(`(${idEq(ids)})`);
  if (searchClause) clauses.push(searchClause);

  const fromVal = document.getElementById('exec-from').value;
  const toVal   = document.getElementById('exec-to').value;
  if (fromVal) clauses.push(`starttime ge ${new Date(fromVal).toISOString()}`);
  if (toVal)   clauses.push(`starttime le ${new Date(toVal).toISOString()}`);

  const status = document.getElementById('exec-status-filter').value;
  if (status && STATUS_VALUES[status]) {
    clauses.push(`(${STATUS_VALUES[status].map(v => `status eq ${str(v)}`).join(' or ')})`);
  }

  const dur = _durationFilterMs();
  if (dur && meta.durationNumeric) clauses.push(`duration ${dur.op} ${dur.ms}`);

  const sortCol = SORT_COLUMNS[_execSortKey] ?? 'starttime';
  const dir = _execSortKey in SORT_COLUMNS ? (_execSortDir === 1 ? 'asc' : 'desc') : 'desc';
  const order = sortCol === 'starttime' ? `starttime ${dir}` : `${sortCol} ${dir},starttime desc`;

  return {
    url: `${ENV_URL}/api/data/v9.2/flowruns` +
      `?$select=name,workflowid,status,errorcode,starttime,createdon,duration,modifiedon` +
      `&$filter=${encodeURIComponent(clauses.join(' and '))}` +
      `&$orderby=${encodeURIComponent(order)}`,
    hasErrorSearch: !!query && !_runErrorSearchOff,
  };
}

async function _loadExecutions() {
  clearTimeout(_execReloadTimer);
  const token = ++_execToken;
  _runPages   = [];
  _runNext    = [];
  _runPage    = 0;
  _runLoading = false;
  _allExecutions = [];
  _renderPager();

  if (_allFlows.length === 0) {
    _showPanelMsg('exec-list', 'No flows loaded.', 'info');
    return;
  }
  if (!_orgId) {
    _showPanelMsg('exec-list',
      'Power Apps Environment ID is not configured for this environment. ' +
      'Open Settings → Edit environment and fill in the "Power Apps Environment ID" field.',
      'error');
    return;
  }

  _showPanelMsg('exec-list', 'Loading executions…', 'loading');
  const meta = await _runMeta();
  if (token !== _execToken) return;

  const q = _runQueryUrl(meta);
  if (q.empty) {
    _setExecCount('');
    _showPanelMsg('exec-list', 'No flows or errors match the search.', 'info');
    return;
  }
  if (q.tooMany) {
    _setExecCount('');
    _showPanelMsg('exec-list',
      `This solution has ${q.tooMany} flows — too many to query together. Select a flow on the left, or search by flow name.`,
      'info');
    return;
  }
  _runFirstUrl = q.url;
  await _loadRunPage(0, token, q.hasErrorSearch);
}

async function _loadRunPage(index, token = _execToken, retryable = false) {
  if (index < 0 || _runLoading) return;
  if (_runPages[index]) {
    _runPage = index;
    _allExecutions = _runPages[index];
    _renderExecutions();
    return;
  }
  const url = index === 0 ? _runFirstUrl : _runNext[index - 1];
  if (!url) return;

  _runLoading = true;
  if (index > 0) _renderPager(true);
  let data;
  try {
    data = await _d365Fetch(url, {
      'Prefer': `odata.include-annotations="OData.Community.Display.V1.FormattedValue",odata.maxpagesize=${RUN_PAGE_SIZE}`,
    });
  } catch (e) {
    if (token !== _execToken) return;
    _runLoading = false;
    if (index === 0 && retryable) {
      console.warn('[EF PPT] Searching flow run errors is not supported here; searching flow names only.', e);
      _runErrorSearchOff = true;
      _loadExecutions();
      return;
    }
    console.error('[EF PPT]', e);
    if (index === 0) _showPanelMsg('exec-list', `Error loading executions: ${e.message}`, 'error');
    else { _renderPager(false); _setExecCount(`Couldn't load page ${index + 1}`); }
    return;
  }
  if (token !== _execToken) return;
  _runLoading = false;

  const flowIdSet = new Set(_allFlows.map(f => (f.workflowid ?? '').toLowerCase()));
  const dur = _runMetaCache && !_runMetaCache.durationNumeric ? _durationFilterMs() : null;
  _runPages[index] = (data.value ?? []).slice(0, RUN_PAGE_SIZE).filter(e => {
    if (!flowIdSet.has((e.workflowid ?? '').toLowerCase())) return false;
    if (!dur) return true;
    const ms = _execDurationMs(e);
    return ms !== null && (dur.op === 'gt' ? ms > dur.ms : ms < dur.ms);
  });
  _runNext[index]  = data['@odata.nextLink'] ?? null;
  _runPage         = index;
  _allExecutions   = _runPages[index];
  _renderExecutions();
}

function _setExecCount(text, title = '') {
  const el = document.getElementById('exec-count');
  el.textContent = text;
  el.title = title;
}

function _renderPager(busy = false) {
  const pager = document.getElementById('exec-pager');
  const hasNext = !!_runNext[_runPage];
  const show = busy || _runPage > 0 || hasNext;
  pager.classList.toggle('hidden', !show);
  if (!show) return;
  pager.innerHTML = `
    <button class="exec-pager__btn" data-page="first" title="First page" ${busy || _runPage === 0 ? 'disabled' : ''}>&laquo;</button>
    <button class="exec-pager__btn" data-page="prev" ${busy || _runPage === 0 ? 'disabled' : ''}>&lsaquo; Prev</button>
    <span class="exec-pager__label">${busy ? 'Loading…' : `Page ${_runPage + 1}`}</span>
    <button class="exec-pager__btn" data-page="next" ${busy || !hasNext ? 'disabled' : ''}>Next &rsaquo;</button>
  `;
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
    const openUrl = _orgId
      ? `https://make.powerautomate.com/environments/${_orgId}/flows/${flow.workflowid}`
      : `${ENV_URL}/main.aspx?pagetype=entityrecord&etn=workflow&id=${flow.workflowid}`;

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

// ─── Render executions ─────────────────────────────────────────────────────────

function _renderExecutions() {
  const flowMap = new Map(_allFlows.map(f => [(f.workflowid ?? '').toLowerCase(), f.name ?? '(Unknown)']));
  const rows = _allExecutions.slice();
  const from = _runPage * RUN_PAGE_SIZE + 1;
  const more = !!_runNext[_runPage];

  _setExecCount(
    rows.length ? `${from.toLocaleString()}–${(from + rows.length - 1).toLocaleString()}${more ? '+' : ''}` : '0',
    more ? 'More runs on the next page' : ''
  );
  _renderPager(false);

  const list = document.getElementById('exec-list');
  if (rows.length === 0) {
    list.innerHTML = `<div class="state-msg state-info">No executions match the selected date range and filters.${
      _runErrorSearchOff && document.getElementById('exec-search').value.trim() ? ' (Search covers flow names only in this environment.)' : ''
    }</div>`;
    return;
  }

  if (_execSortKey === 'flow') {
    rows.sort((a, b) => {
      const av = (flowMap.get((a.workflowid ?? '').toLowerCase()) ?? '').toLowerCase();
      const bv = (flowMap.get((b.workflowid ?? '').toLowerCase()) ?? '').toLowerCase();
      return av < bv ? -_execSortDir : av > bv ? _execSortDir : 0;
    });
  }

  list.innerHTML = '';

  const table = document.createElement('table');
  table.className = 'exec-table';
  table.innerHTML = `
    <colgroup>
      <col><col><col><col><col>
    </colgroup>
    <thead><tr>
      <th class="exec-th" data-sort="status">Status ${_execSortIcon('status')}</th>
      <th class="exec-th" data-sort="flow" title="Sorts the runs on this page">Flow ${_execSortIcon('flow')}</th>
      <th class="exec-th" data-sort="date">Date ${_execSortIcon('date')}</th>
      <th class="exec-th" data-sort="duration">Duration ${_execSortIcon('duration')}</th>
      <th class="exec-th" data-sort="error">Error ${_execSortIcon('error')}</th>
    </tr></thead>
  `;

  table.addEventListener('click', e => {
    const th = e.target.closest('[data-sort]');
    if (!th) return;
    const key = th.dataset.sort;
    if (_execSortKey === key) {
      _execSortDir *= -1;
    } else {
      _execSortKey = key;
      _execSortDir = key === 'date' ? -1 : 1;
    }
    if (key === 'flow') _renderExecutions();
    else _loadExecutions();
  });

  const tbody = document.createElement('tbody');

  rows.forEach(exec => {
    const st         = _execStatus(exec);
    const flowId     = (exec.workflowid ?? '').toLowerCase();
    const flowName   = flowMap.get(flowId) ?? '—';
    const started    = _fmtDate(exec.starttime);
    const duration   = _fmtDurationMs(_execDurationMs(exec));
    const errFull    = exec.errorcode ?? '';
    const errPreview = errFull.length > 90 ? errFull.slice(0, 90) + '…' : errFull;
    const openUrl    = _orgId
      ? `https://make.powerautomate.com/environments/${_orgId}/flows/${exec.workflowid}/runs/${exec.name}`
      : '';

    const tr = document.createElement('tr');
    if (exec.name === _selectedExecId) tr.classList.add('selected');
    tr.innerHTML = `
      <td><span class="exec-badge exec-badge--${st.cls}">${_esc(st.label)}</span></td>
      <td class="exec-flow-name" title="${_esc(flowName)}">
        ${openUrl
          ? `<a class="exec-flow-link" href="${openUrl}" target="_blank" rel="noopener noreferrer">${_esc(flowName)}</a>`
          : _esc(flowName)}
      </td>
      <td class="exec-date">${started}</td>
      <td class="exec-duration">${duration}</td>
      <td class="exec-error" title="${_esc(errFull)}">${_esc(errPreview)}</td>
    `;

    tr.addEventListener('click', () => {
      _selectedExecId = exec.name === _selectedExecId ? null : exec.name;
      tbody.querySelectorAll('tr').forEach(r => r.classList.remove('selected'));
      if (_selectedExecId) tr.classList.add('selected');
    });

    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  list.appendChild(table);
  list.scrollTop = 0;
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

// Returns the sort-icon HTML for a given column key.
function _execSortIcon(key) {
  if (_execSortKey !== key) return '<span class="exec-sort-icon">⇅</span>';
  return `<span class="exec-sort-icon exec-sort-icon--active">${_execSortDir === 1 ? '▲' : '▼'}</span>`;
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
