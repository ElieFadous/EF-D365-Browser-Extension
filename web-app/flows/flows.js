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
  document.getElementById('exec-search').addEventListener('input', _renderExecutions);
  document.getElementById('exec-status-filter').addEventListener('change', _renderExecutions);
  ['exec-from', 'exec-to'].forEach(id =>
    document.getElementById(id).addEventListener('keydown', e => { if (e.key === 'Enter') _loadExecutions(); })
  );

  // Duration filter (client-side, no API reload needed)
  ['dur-op', 'dur-unit'].forEach(id =>
    document.getElementById(id).addEventListener('change', _renderExecutions)
  );
  document.getElementById('dur-value').addEventListener('input', _renderExecutions);
  document.getElementById('btn-clear-dur').addEventListener('click', () => {
    document.getElementById('dur-op').value    = '';
    document.getElementById('dur-value').value = '';
    document.getElementById('dur-unit').value  = 'ms';
    _renderExecutions();
  });

  // Clear flow filter
  document.getElementById('btn-clear-flow-filter').addEventListener('click', () => {
    _selectedFlowId = null;
    document.getElementById('exec-scope').textContent = '';
    document.getElementById('btn-clear-flow-filter').classList.add('hidden');
    document.querySelectorAll('.flow-item').forEach(el => el.classList.remove('selected'));
    _renderExecutions();
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
    window.parent.postMessage({ __efppt: 'fetch', id, url, method, headers: extraHeaders, body }, '*');
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

async function _loadExecutions() {
  if (_allFlows.length === 0) {
    _showPanelMsg('exec-list', 'No flows loaded.', 'info');
    return;
  }

  _showPanelMsg('exec-list', 'Loading executions…', 'loading');

  if (!_orgId) {
    _showPanelMsg('exec-list',
      'Power Apps Environment ID is not configured for this environment. ' +
      'Open Settings → Edit environment and fill in the "Power Apps Environment ID" field.',
      'error');
    return;
  }

  try {
    const fromVal = document.getElementById('exec-from').value;
    const toVal   = document.getElementById('exec-to').value;

    // Server-side date filter (best-effort — some environments don't honour OData
    // filtering on flowruns.starttime; _renderExecutions also filters client-side).
    const conditions = [];
    if (fromVal) conditions.push(`starttime ge ${new Date(fromVal).toISOString()}`);
    if (toVal)   conditions.push(`starttime le ${new Date(toVal).toISOString()}`);

    const filter  = conditions.length > 0 ? `&$filter=${conditions.join(' and ')}` : '';
    const execUrl = `${ENV_URL}/api/data/v9.2/flowruns` +
      `?$select=name,workflowid,status,errorcode,starttime,createdon,duration,modifiedon` +
      `&$orderby=starttime desc` +
      filter;

    const flowIdSet = new Set(_allFlows.map(f => (f.workflowid ?? '').toLowerCase()));
    const raw = await _fetchAllPages(execUrl, 10); // up to 5 000 records

    // Filter client-side to only runs belonging to flows in this solution.
    _allExecutions = raw.filter(e =>
      flowIdSet.has((e.workflowid ?? '').toLowerCase())
    );

    _renderExecutions();
  } catch (e) {
    console.error('[EF PPT]', e);
    _showPanelMsg('exec-list', `Error loading executions: ${e.message}`, 'error');
  }
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
      _renderExecutions();
    });

    list.appendChild(item);
  });
}

// ─── Render executions ─────────────────────────────────────────────────────────

function _renderExecutions() {
  const query        = document.getElementById('exec-search').value.trim().toLowerCase();
  const statusFilter = document.getElementById('exec-status-filter').value;

  // Client-side date bounds (reliable backstop — server OData filter on flowruns
  // createdon is not consistently honoured by all environments).
  const fromVal  = document.getElementById('exec-from').value;
  const toVal    = document.getElementById('exec-to').value;
  const fromDate = fromVal ? new Date(fromVal) : null;
  const toDate   = toVal   ? new Date(toVal)   : null;

  // Duration filter
  const durOp    = document.getElementById('dur-op').value;
  const durRaw   = parseFloat(document.getElementById('dur-value').value);
  const durUnit  = document.getElementById('dur-unit').value;
  let   durMs    = NaN;
  if (durOp && !isNaN(durRaw)) {
    durMs = durUnit === 'ms'  ? durRaw
          : durUnit === 's'   ? durRaw * 1_000
          : durUnit === 'min' ? durRaw * 60_000
          :                     durRaw * 3_600_000;   // 'h'
  }

  // Flow name lookup map
  const flowMap = new Map(_allFlows.map(f => [(f.workflowid ?? '').toLowerCase(), f.name ?? '(Unknown)']));

  _filteredExecutions = _allExecutions.filter(exec => {
    const flowId = (exec.workflowid ?? '').toLowerCase();

    // Filter to selected flow
    if (_selectedFlowId && flowId !== _selectedFlowId.toLowerCase()) return false;

    // Client-side date filter
    if (fromDate || toDate) {
      const d = exec.starttime ? new Date(exec.starttime) : null;
      if (d) {
        if (fromDate && d < fromDate) return false;
        if (toDate   && d > toDate)   return false;
      }
    }

    // Status filter
    if (statusFilter) {
      const st = _execStatus(exec);
      if (st.cls !== statusFilter) return false;
    }

    // Duration filter
    if (!isNaN(durMs)) {
      const dur = _execDurationMs(exec);
      if (dur === null) return false;
      if (durOp === 'gt' && dur <= durMs) return false;
      if (durOp === 'lt' && dur >= durMs) return false;
    }

    // Text search: flow name or error code
    if (query) {
      const name = (flowMap.get(flowId) ?? '').toLowerCase();
      const err  = (exec.errorcode ?? '').toLowerCase();
      if (!name.includes(query) && !err.includes(query)) return false;
    }

    return true;
  });

  document.getElementById('exec-count').textContent =
    `${_filteredExecutions.length} of ${_allExecutions.length}`;

  const list = document.getElementById('exec-list');
  if (_filteredExecutions.length === 0) {
    list.innerHTML = `<div class="state-msg state-info">${
      _allExecutions.length === 0
        ? 'No executions found for the selected date range.'
        : 'No executions match the current filter.'
    }</div>`;
    return;
  }

  // ── Sort ───────────────────────────────────────────────────────────────────────
  _filteredExecutions.sort((a, b) => {
    let av, bv;
    switch (_execSortKey) {
      case 'status':
        av = _execStatus(a).label;
        bv = _execStatus(b).label;
        break;
      case 'flow':
        av = (flowMap.get((a.workflowid ?? '').toLowerCase()) ?? '').toLowerCase();
        bv = (flowMap.get((b.workflowid ?? '').toLowerCase()) ?? '').toLowerCase();
        break;
      case 'date':
        av = a.starttime ?? '';
        bv = b.starttime ?? '';
        break;
      case 'duration':
        av = _execDurationMs(a) ?? -1;   // nulls sort to bottom
        bv = _execDurationMs(b) ?? -1;
        break;
      case 'error':
        av = (a.errorcode ?? '').toLowerCase();
        bv = (b.errorcode ?? '').toLowerCase();
        break;
      default: return 0;
    }
    if (av < bv) return -_execSortDir;
    if (av > bv) return  _execSortDir;
    return 0;
  });

  list.innerHTML = '';

  const table = document.createElement('table');
  table.className = 'exec-table';
  table.innerHTML = `
    <colgroup>
      <col><col><col><col><col>
    </colgroup>
    <thead><tr>
      <th class="exec-th" data-sort="status">Status ${_execSortIcon('status')}</th>
      <th class="exec-th" data-sort="flow">Flow ${_execSortIcon('flow')}</th>
      <th class="exec-th" data-sort="date">Date ${_execSortIcon('date')}</th>
      <th class="exec-th" data-sort="duration">Duration ${_execSortIcon('duration')}</th>
      <th class="exec-th" data-sort="error">Error ${_execSortIcon('error')}</th>
    </tr></thead>
  `;

  // Delegated click on the header — no need to re-query after every render.
  table.addEventListener('click', e => {
    const th = e.target.closest('[data-sort]');
    if (!th) return;
    const key = th.dataset.sort;
    if (_execSortKey === key) {
      _execSortDir *= -1;
    } else {
      _execSortKey = key;
      _execSortDir = key === 'date' ? -1 : 1;  // date defaults desc, others asc
    }
    _renderExecutions();
  });

  const tbody = document.createElement('tbody');

  _filteredExecutions.forEach(exec => {
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
