/**
 * EF Power Platform Tools — Data Sync (web-app)
 *
 * Synchronises records between two D365 environments.
 * Steps: Environments → Table → Columns & N:N → Filter → Results
 *
 * Runs as an iframe opened by launcher.js (which runs in the D365 page context).
 * All D365 API calls are relayed via the postMessage bridge — no direct fetch and
 * no chrome.* APIs.
 *
 * IMPORTANT LIMITATION:
 * The bridge (launcher.js on the source D365 page) can only relay same-origin calls.
 * Cross-environment sync (source ≠ target) is NOT supported in the web app — only
 * single-environment operations where source = target = the current D365 page.
 */

// ─── Bridge fetch ───────────────────────────────────────────────────────────────
// launcher.js (running in D365 context) relays this same-origin call with
// credentials included, then posts back { __efppt: 'fetch-result', id, ok, data }.
// Resolves with the parsed response body (caller reads .value etc. itself).

const _bridgeFetch = (url, extraHeaders = {}, method = 'GET', body = null) =>
  new Promise((resolve, reject) => {
    const id    = Math.random().toString(36).slice(2) + Date.now();
    const timer = setTimeout(() => {
      window.removeEventListener('message', onMsg);
      reject(new Error('Request timed out after 30s'));
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

// HEAD-style existence check — needs the status code, not the body. The bridge
// rejects on non-2xx, so a resolved promise means the record exists (200).
const _bridgeExists = (url) =>
  _bridgeFetch(url).then(() => true).catch(() => false);

// ─── Config (localStorage) ───────────────────────────────────────────────────────

function _loadConfig() {
  try { return JSON.parse(localStorage.getItem('ef_ppt_config')) ?? {}; }
  catch { return {}; }
}

// ─── State ────────────────────────────────────────────────────────────────────

let environments   = [];
let settings       = {};
let sourceEnv      = null;   // { url, name, id }
let targetEnv      = null;
let entityMeta     = null;   // { logicalName, displayName, entitySetName, primaryIdAttribute, primaryNameAttribute }
let allAttributes  = [];     // full attr list from source
let allM2m         = [];     // full N:N list from source
let selectedAttrs  = null;   // null = all; otherwise Set of logicalName
let selectedM2m    = null;   // null = all; otherwise Set of schemaName
let filterConditions = [];   // [{ attribute, operator, value }]
let syncResults    = [];     // accumulates per-record results
let currentStep    = 1;

// Sort state for summary table
let sortCol = 'name', sortDir = 'asc';

// Option set value cache for filter builder — keyed by "{logicalName}|{attrName}"
const _optionSetCache = new Map();

// ─── Bootstrap ────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  const config = _loadConfig();
  environments = Array.isArray(config.environments) ? config.environments : [];
  settings = {
    apiVersion: 'v9.2',
    syncBatchSize: 250,
    ...(config.settings ?? {}),
  };

  // The launcher passes ?env= (the current D365 origin) and ?name= for the page.
  const _params         = new URLSearchParams(location.search);
  const _currentEnvUrl  = (_params.get('env') ?? '').replace(/\/$/, '');

  // Pre-select the source env matching the current D365 page if we recognise it.
  const _defaultSourceUrl = environments.find(e =>
    (e.url ?? '').replace(/\/$/, '') === _currentEnvUrl)?.url ?? '';

  populateEnvDropdowns(_defaultSourceUrl);
  bindStep1();
  bindStep4();
  bindStep5();
  updateFooterButtons(1);

  // Stepper click — allow navigating back to any completed step
  document.querySelectorAll('.step[data-step]').forEach(el => {
    el.addEventListener('click', async () => {
      const targetStep = parseInt(el.dataset.step, 10);
      if (el.classList.contains('done') && targetStep < currentStep) {
        await navigateBack(targetStep);
      }
    });
  });

  // Reset button — return to step 1 and clear all state
  document.getElementById('btn-reset').addEventListener('click', () => navigateBack(1));
});

// ─── Environment dropdowns ─────────────────────────────────────────────────────

function populateEnvDropdowns(defaultSourceUrl = '') {
  const srcSel = document.getElementById('source-env');
  const tgtSel = document.getElementById('target-env');

  environments.forEach(env => {
    srcSel.appendChild(new Option(env.name, env.url));
    tgtSel.appendChild(new Option(env.name, env.url));
  });

  if (defaultSourceUrl) {
    srcSel.value = defaultSourceUrl;
    // In single-env mode the only supported target is the source — default it too.
    tgtSel.value = defaultSourceUrl;
  }

  srcSel.addEventListener('change', validateStep1);
  tgtSel.addEventListener('change', validateStep1);

  // Run validation immediately so Next button state reflects the pre-selection
  validateStep1();
}

/** Current D365 origin the bridge can reach (the page hosting the launcher). */
function _currentOrigin() {
  const fromParam = (new URLSearchParams(location.search).get('env') ?? '').replace(/\/$/, '');
  if (fromParam) { try { return new URL(fromParam).origin; } catch { /* fall through */ } }
  // Fallback: the parent (D365) origin via the document referrer.
  try { return new URL(document.referrer).origin; } catch { return ''; }
}

function _sameOrigin(envUrl) {
  if (!envUrl) return false;
  const cur = _currentOrigin();
  try { return new URL(envUrl).origin === cur; } catch { return false; }
}

function validateStep1() {
  const src     = document.getElementById('source-env').value;
  const tgt     = document.getElementById('target-env').value;
  const errEl   = document.getElementById('env-error');
  const nextBtn = document.getElementById('btn-step1-next');

  // Nothing selected yet → disable Next, clear error.
  if (!src || !tgt) {
    errEl.classList.add('hidden');
    nextBtn.disabled = true;
    return;
  }

  // Cross-environment sync is impossible via the same-origin bridge.
  if (src !== tgt) {
    errEl.textContent =
      'Cross-environment sync is not supported in the web app version. ' +
      'To sync between environments, use the browser extension. ' +
      'You can sync within the same environment (source = target = current page).';
    errEl.classList.remove('hidden');
    nextBtn.disabled = true;
    return;
  }

  // Single-env mode: the selected env must be the D365 page the launcher runs on,
  // otherwise the bridge cannot reach it (cross-origin).
  if (!_sameOrigin(src)) {
    errEl.textContent =
      'The selected environment is not the current D365 page. ' +
      'The web app can only sync within the environment you launched it from. ' +
      'Open this tool from the environment you want to sync, or use the browser extension.';
    errEl.classList.remove('hidden');
    nextBtn.disabled = true;
    return;
  }

  errEl.classList.add('hidden');
  nextBtn.disabled = false;
}

// ─── Step bindings ────────────────────────────────────────────────────────────

function bindStep1() {
  document.getElementById('btn-step1-next').addEventListener('click', () => {
    const srcUrl = document.getElementById('source-env').value;
    const tgtUrl = document.getElementById('target-env').value;
    sourceEnv = environments.find(e => e.url === srcUrl);
    targetEnv = environments.find(e => e.url === tgtUrl);

    // No proxy tab to open — the bridge talks to the current (source) origin.
    goToStep(2);
    loadEntityList();
  });
}

function setStep1Loading(loading) {
  document.getElementById('btn-step1-next').disabled = loading;
  document.getElementById('source-env').disabled = loading;
  document.getElementById('target-env').disabled = loading;
}

function showEnvError(msg) {
  const el = document.getElementById('env-error');
  el.textContent = msg;
  el.classList.remove('hidden');
}

// ─── Step 2: Entity list ───────────────────────────────────────────────────────

async function loadEntityList() {
  const stateEl = document.getElementById('table-list-state');
  const listEl  = document.getElementById('table-list');

  stateEl.innerHTML = '<span class="spinner"></span> Fetching entity list…';
  stateEl.className = 'state-msg state-loading';
  stateEl.style.display = '';
  listEl.classList.add('hidden');

  const apiBase = `${sourceEnv.url}/api/data/${settings.apiVersion}`;

  // Fetch without server-side $filter — D365 metadata API has unreliable filter support
  // for complex types like IsCustomizable/Value. Filter and sort client-side instead.
  const url = `${apiBase}/EntityDefinitions` +
    `?$select=LogicalName,DisplayName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,IsCustomizable`;

  let entities;
  try {
    const d = await _bridgeFetch(url);
    entities = d.value ?? [];
  } catch (e) {
    console.error('[Data Sync] Failed to load entity list:', e);
    stateEl.textContent = `Failed to load tables: ${e.message}`;
    stateEl.className = 'state-msg state-error';
    return;
  }

  // Filter client-side: only customisable entities
  const filtered = entities.filter(e => e.IsCustomizable?.Value === true);

  // Sort alphabetically by display name
  filtered.sort((a, b) => {
    const la = (a.DisplayName?.UserLocalizedLabel?.Label ?? a.LogicalName).toLowerCase();
    const lb = (b.DisplayName?.UserLocalizedLabel?.Label ?? b.LogicalName).toLowerCase();
    return la < lb ? -1 : la > lb ? 1 : 0;
  });

  if (filtered.length === 0) {
    stateEl.textContent = 'No customisable tables found.';
    stateEl.className = 'state-msg state-info';
    return;
  }

  renderEntityList(filtered);

  // Wire up search
  document.getElementById('table-search').addEventListener('input', e => {
    filterEntityList(e.target.value.trim().toLowerCase());
  });

  // Back button
  document.getElementById('btn-step2-back').addEventListener('click', () => {
    goToStep(1);
    setStep1Loading(false);
    document.getElementById('source-env').disabled = false;
    document.getElementById('target-env').disabled = false;
  });

  // Next button — reset downstream state since the table may have changed
  document.getElementById('btn-step2-next').addEventListener('click', () => {
    filterConditions = [];
    syncResults = [];
    goToStep(3);
    loadColumnsAndM2m();
  });
}

function renderEntityList(entities) {
  const listEl  = document.getElementById('table-list');
  const stateEl = document.getElementById('table-list-state');

  listEl.innerHTML = '';
  entities.forEach(ent => {
    const label = ent.DisplayName?.UserLocalizedLabel?.Label ?? ent.LogicalName;
    const li = document.createElement('li');
    li.dataset.logicalName = ent.LogicalName;
    li.dataset.displayName = label.toLowerCase();
    li.dataset.entitySet   = ent.EntitySetName;
    li.dataset.primaryId   = ent.PrimaryIdAttribute;
    li.dataset.primaryName = ent.PrimaryNameAttribute;
    li.innerHTML = `<span class="item-display">${escHtml(label)}</span>
                    <span class="item-logical">${escHtml(ent.LogicalName)}</span>`;
    li.addEventListener('click', () => selectEntity(li, ent, label));
    listEl.appendChild(li);
  });

  stateEl.style.display = 'none';
  listEl.classList.remove('hidden');
}

function filterEntityList(query) {
  document.querySelectorAll('#table-list li').forEach(li => {
    const show = !query
      || li.dataset.displayName.includes(query)
      || li.dataset.logicalName.includes(query);
    li.style.display = show ? '' : 'none';
  });
}

function selectEntity(li, ent, label) {
  document.querySelectorAll('#table-list li').forEach(el => el.classList.remove('selected'));
  li.classList.add('selected');
  entityMeta = {
    logicalName:          ent.LogicalName,
    displayName:          label,
    entitySetName:        ent.EntitySetName,
    primaryIdAttribute:   ent.PrimaryIdAttribute,
    primaryNameAttribute: ent.PrimaryNameAttribute,
  };
  document.getElementById('btn-step2-next').disabled = false;
}

// ─── Step 3: Columns & N:N ────────────────────────────────────────────────────

async function loadColumnsAndM2m() {
  const colsStateEl = document.getElementById('cols-state');
  const colsListEl  = document.getElementById('cols-list');
  const m2mStateEl  = document.getElementById('m2m-state');
  const m2mListEl   = document.getElementById('m2m-list');

  colsListEl.classList.add('hidden');
  m2mListEl.classList.add('hidden');
  colsStateEl.innerHTML = '<span class="spinner"></span> Loading…';
  colsStateEl.style.display = '';
  colsStateEl.className = 'state-msg state-loading';
  m2mStateEl.classList.add('hidden');

  const apiBase = `${sourceEnv.url}/api/data/${settings.apiVersion}`;
  const etn     = entityMeta.logicalName;

  let schema;
  try {
    const [attrD, m2mD] = await Promise.all([
      _bridgeFetch(
        `${apiBase}/EntityDefinitions(LogicalName='${etn}')/Attributes` +
        `?$select=LogicalName,DisplayName,AttributeType,AttributeTypeName,IsValidForCreate`
      ),
      _bridgeFetch(
        `${apiBase}/EntityDefinitions(LogicalName='${etn}')/ManyToManyRelationships` +
        `?$select=SchemaName,Entity1LogicalName,Entity2LogicalName,Entity1NavigationPropertyName,Entity2NavigationPropertyName`
      ),
    ]);
    schema = { attrs: attrD.value ?? [], m2m: m2mD.value ?? [] };
  } catch (e) {
    console.error('[Data Sync] Failed to load columns/M2M schema:', e);
    colsStateEl.textContent = `Failed to load schema: ${e.message}`;
    colsStateEl.className = 'state-msg state-error';
    return;
  }

  allAttributes = schema.attrs;
  allM2m        = schema.m2m;
  selectedAttrs = null; // all by default
  selectedM2m   = null;

  renderColsList(allAttributes, '');
  renderM2mList(allM2m);

  // Column search
  document.getElementById('cols-search').addEventListener('input', e => {
    renderColsList(allAttributes, e.target.value.trim().toLowerCase());
  });

  // Select All / Clear All columns
  document.getElementById('btn-cols-all').addEventListener('click', () => {
    document.querySelectorAll('#cols-list input[type="checkbox"]').forEach(cb => { cb.checked = true; });
    updateSelectedAttrs();
  });
  document.getElementById('btn-cols-clear').addEventListener('click', () => {
    document.querySelectorAll('#cols-list input[type="checkbox"]').forEach(cb => { cb.checked = false; });
    updateSelectedAttrs();
  });

  // Select All / Clear All N:N
  document.getElementById('btn-m2m-all').addEventListener('click', () => {
    document.querySelectorAll('#m2m-list input[type="checkbox"]').forEach(cb => { cb.checked = true; });
    updateSelectedM2m();
  });
  document.getElementById('btn-m2m-clear').addEventListener('click', () => {
    document.querySelectorAll('#m2m-list input[type="checkbox"]').forEach(cb => { cb.checked = false; });
    updateSelectedM2m();
  });

  // Back / Next
  document.getElementById('btn-step3-back').addEventListener('click', () => {
    goToStep(2);
  });
  document.getElementById('btn-step3-next').addEventListener('click', () => {
    updateSelectedAttrs();
    updateSelectedM2m();
    goToStep(4);
    initFilterStep();
  });
}

/**
 * D365 returns AttributeType='Virtual' for MultiSelectPicklist columns.
 * Use AttributeTypeName.Value to distinguish them from true virtual columns.
 */
function getEffectiveAttrType(attr) {
  if (attr.AttributeType === 'Virtual' && attr.AttributeTypeName?.Value === 'MultiSelectPicklistType') {
    return 'MultiSelectPicklist';
  }
  return attr.AttributeType;
}

function renderColsList(attrs, query) {
  const listEl  = document.getElementById('cols-list');
  const stateEl = document.getElementById('cols-state');

  const SKIP_TYPES = new Set(['Virtual', 'EntityName', 'ManagedProperty', 'Uniqueidentifier']);
  const filtered = attrs.filter(a => {
    // Use effective type so MultiSelectPicklist (reported as Virtual by D365) is not excluded
    if (SKIP_TYPES.has(getEffectiveAttrType(a))) return false;
    if (!query) return true;
    const label = (a.DisplayName?.UserLocalizedLabel?.Label ?? a.LogicalName).toLowerCase();
    return label.includes(query) || a.LogicalName.includes(query);
  });

  // Sort alphabetically by display label
  filtered.sort((a, b) => {
    const la = (a.DisplayName?.UserLocalizedLabel?.Label ?? a.LogicalName).toLowerCase();
    const lb = (b.DisplayName?.UserLocalizedLabel?.Label ?? b.LogicalName).toLowerCase();
    return la < lb ? -1 : la > lb ? 1 : 0;
  });

  listEl.innerHTML = '';
  filtered.forEach(attr => {
    const label = attr.DisplayName?.UserLocalizedLabel?.Label ?? attr.LogicalName;
    const li = document.createElement('li');
    const isChecked = selectedAttrs === null || selectedAttrs.has(attr.LogicalName);
    li.innerHTML = `
      <input type="checkbox" value="${escHtml(attr.LogicalName)}" ${isChecked ? 'checked' : ''} />
      <div>
        <div class="item-display">${escHtml(label)}</div>
        <div class="item-logical">${escHtml(attr.LogicalName)}</div>
      </div>`;
    li.querySelector('input').addEventListener('change', updateSelectedAttrs);
    listEl.appendChild(li);
  });

  stateEl.style.display = 'none';
  listEl.classList.remove('hidden');
}

function renderM2mList(m2m) {
  const listEl  = document.getElementById('m2m-list');
  const stateEl = document.getElementById('m2m-state');

  if (m2m.length === 0) {
    stateEl.classList.remove('hidden');
    stateEl.style.display = '';
    listEl.classList.add('hidden');
    return;
  }

  // Sort alphabetically by schema name
  const sortedM2m = [...m2m].sort((a, b) => a.SchemaName.toLowerCase() < b.SchemaName.toLowerCase() ? -1 : a.SchemaName.toLowerCase() > b.SchemaName.toLowerCase() ? 1 : 0);

  listEl.innerHTML = '';
  sortedM2m.forEach(rel => {
    const li = document.createElement('li');
    const isChecked = selectedM2m === null || selectedM2m.has(rel.SchemaName);
    li.innerHTML = `
      <input type="checkbox" value="${escHtml(rel.SchemaName)}" ${isChecked ? 'checked' : ''} />
      <div>
        <div class="item-display">${escHtml(rel.SchemaName)}</div>
        <div class="item-logical">${escHtml(rel.Entity1LogicalName)} ↔ ${escHtml(rel.Entity2LogicalName)}</div>
      </div>`;
    li.querySelector('input').addEventListener('change', updateSelectedM2m);
    listEl.appendChild(li);
  });

  listEl.classList.remove('hidden');
}

function updateSelectedAttrs() {
  const checked = Array.from(
    document.querySelectorAll('#cols-list input[type="checkbox"]:checked')
  ).map(cb => cb.value);
  const all = Array.from(
    document.querySelectorAll('#cols-list input[type="checkbox"]')
  ).map(cb => cb.value);
  selectedAttrs = checked.length === all.length ? null : new Set(checked);
}

function updateSelectedM2m() {
  const checked = Array.from(
    document.querySelectorAll('#m2m-list input[type="checkbox"]:checked')
  ).map(cb => cb.value);
  const all = Array.from(
    document.querySelectorAll('#m2m-list input[type="checkbox"]')
  ).map(cb => cb.value);
  selectedM2m = checked.length === all.length ? null : new Set(checked);
}

// ─── Step 4: Filter builder ────────────────────────────────────────────────────

function initFilterStep() {
  filterConditions = [];
  renderFilterConditions();
  updateFetchXmlPreview();

  document.getElementById('filter-logic').addEventListener('change', updateFetchXmlPreview);
}

/** Custom checkbox-based multi-select dropdown for filter conditions. */
function buildMultiSelectDropdown(capturedIdx, capturedAttr, attrType) {
  const wrapper = document.createElement('div');
  wrapper.className = 'ms-dropdown';

  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'ms-dropdown-btn';
  btn.textContent = 'Loading…';
  btn.disabled = true;

  const panel = document.createElement('div');
  panel.className = 'ms-dropdown-panel hidden';

  wrapper.appendChild(btn);
  wrapper.appendChild(panel);

  btn.addEventListener('click', e => {
    e.stopPropagation();
    const wasHidden = panel.classList.contains('hidden');
    // Close any other open panels first
    document.querySelectorAll('.ms-dropdown-panel').forEach(p => p.classList.add('hidden'));
    if (wasHidden) panel.classList.remove('hidden');
  });

  // Prevent clicks inside the panel from bubbling to the document close handler
  panel.addEventListener('click', e => e.stopPropagation());

  function updateLabel() {
    const checked = Array.from(panel.querySelectorAll('input[type="checkbox"]:checked'));
    if (checked.length === 0) { btn.textContent = 'Select values…'; return; }
    // Show comma-separated labels; the button has overflow:hidden + text-overflow:ellipsis
    btn.textContent = checked.map(cb => cb.nextElementSibling?.textContent ?? cb.value).join(', ');
  }

  fetchOptionSetValues(entityMeta.logicalName, capturedAttr, attrType)
    .then(options => {
      panel.innerHTML = '';
      const currentVals = new Set((filterConditions[capturedIdx]?.value ?? '').split(',').filter(v => v));
      options.forEach(o => {
        const lbl = document.createElement('label');
        lbl.className = 'ms-option';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.value = String(o.value);
        cb.checked = currentVals.has(String(o.value));
        const span = document.createElement('span');
        span.textContent = o.label;
        lbl.appendChild(cb);
        lbl.appendChild(span);
        cb.addEventListener('change', () => {
          filterConditions[capturedIdx].value = Array.from(
            panel.querySelectorAll('input[type="checkbox"]:checked')
          ).map(c => c.value).join(',');
          updateLabel();
          updateFetchXmlPreview();
        });
        panel.appendChild(lbl);
      });
      btn.disabled = false;
      updateLabel();
    })
    .catch(e => {
      console.error('[Data Sync] Failed to load multi-select option values:', e);
      btn.textContent = 'Failed to load options';
    });

  return wrapper;
}

function bindStep4() {
  // Close all multi-select panels when clicking outside
  document.addEventListener('click', () => {
    document.querySelectorAll('.ms-dropdown-panel').forEach(p => p.classList.add('hidden'));
  });

  document.getElementById('btn-add-condition').addEventListener('click', () => {
    filterConditions.push({ attribute: '', operator: 'eq', value: '' });
    renderFilterConditions();
    updateFetchXmlPreview();
  });

  document.getElementById('btn-step4-back').addEventListener('click', () => {
    goToStep(3);
  });

  document.getElementById('btn-view-updates').addEventListener('click', () => {
    goToStep(5);
    runViewUpdates();
  });

  document.getElementById('btn-sync').addEventListener('click', () => {
    goToStep(5);
    runSync();
  });
}

// D365 date operators that take no value (self-closing condition in FetchXML)
const DATE_NO_VALUE_OPS = new Set([
  'today','yesterday','tomorrow',
  'this-week','last-week','next-week',
  'this-month','last-month','next-month',
  'this-year','last-year','next-year',
  'last-7-days','next-7-days','last-30-days','next-30-days',
]);

// D365 date operators that take a single integer value (e.g. "last x days = 7")
const DATE_NUMBER_VALUE_OPS = new Set([
  'last-x-days','last-x-weeks','last-x-months','last-x-years',
  'next-x-days','next-x-weeks','next-x-months','next-x-years',
  'older-than-x-days','older-than-x-weeks','older-than-x-months','older-than-x-years',
]);

function renderFilterConditions() {
  const container = document.getElementById('filter-conditions');
  container.innerHTML = '';

  filterConditions.forEach((cond, idx) => {
    const row = document.createElement('div');
    row.className = 'filter-row';

    // Attribute dropdown — sorted alphabetically by display label
    const attrSel = document.createElement('select');
    attrSel.innerHTML = `<option value="">— Attribute —</option>`;
    const SKIP_TYPES = new Set(['Virtual', 'EntityName', 'ManagedProperty', 'Uniqueidentifier']);
    const filterableAttrs = allAttributes
      .filter(a => !SKIP_TYPES.has(getEffectiveAttrType(a)))
      .slice()
      .sort((a, b) => {
        const la = (a.DisplayName?.UserLocalizedLabel?.Label ?? a.LogicalName).toLowerCase();
        const lb = (b.DisplayName?.UserLocalizedLabel?.Label ?? b.LogicalName).toLowerCase();
        return la < lb ? -1 : la > lb ? 1 : 0;
      });
    filterableAttrs.forEach(attr => {
      const label = attr.DisplayName?.UserLocalizedLabel?.Label ?? attr.LogicalName;
      const opt = new Option(`${label} (${attr.LogicalName})`, attr.LogicalName);
      if (attr.LogicalName === cond.attribute) opt.selected = true;
      attrSel.appendChild(opt);
    });
    attrSel.addEventListener('change', () => {
      filterConditions[idx].attribute = attrSel.value;
      filterConditions[idx].value     = '';
      // Default to the first valid operator for the new attribute type
      const newAttr = allAttributes.find(a => a.LogicalName === attrSel.value);
      const newType = newAttr ? getEffectiveAttrType(newAttr) : '';
      filterConditions[idx].operator  = getOperators(newType)[0]?.[0] ?? 'eq';
      renderFilterConditions();
      updateFetchXmlPreview();
    });

    // Operator dropdown
    const opSel = document.createElement('select');
    const _foundAttr = allAttributes.find(a => a.LogicalName === cond.attribute);
    const attrType = _foundAttr ? getEffectiveAttrType(_foundAttr) : '';
    getOperators(attrType).forEach(([val, label]) => {
      const opt = new Option(label, val);
      if (val === cond.operator) opt.selected = true;
      opSel.appendChild(opt);
    });
    opSel.addEventListener('change', () => {
      filterConditions[idx].operator = opSel.value;
      filterConditions[idx].value    = '';
      renderFilterConditions();
      updateFetchXmlPreview();
    });

    // Value input — type-aware
    const valWrap = document.createElement('div');
    valWrap.style.width = '100%';
    const noVal = ['null', 'not-null'].includes(cond.operator) || DATE_NO_VALUE_OPS.has(cond.operator);
    if (!noVal) {
      let valEl;
      if (attrType === 'DateTime') {
        if (DATE_NUMBER_VALUE_OPS.has(cond.operator)) {
          // e.g. "last-x-days" — needs a plain integer
          valEl = document.createElement('input');
          valEl.type        = 'number';
          valEl.min         = '1';
          valEl.value       = cond.value || '7';
          valEl.placeholder = 'e.g. 7';
        } else {
          valEl = document.createElement('input');
          valEl.type  = 'date';
          valEl.value = cond.value;
        }
      } else if (attrType === 'Boolean') {
        valEl = document.createElement('select');
        valEl.innerHTML = `<option value="1" ${cond.value === '1' ? 'selected' : ''}>Yes</option>
                           <option value="0" ${cond.value === '0' ? 'selected' : ''}>No</option>`;
      } else if (['Integer', 'Decimal', 'Double', 'Money', 'BigInt'].includes(attrType)) {
        valEl = document.createElement('input');
        valEl.type  = 'number';
        valEl.value = cond.value;
        valEl.placeholder = '0';
      } else if (['Lookup', 'Customer', 'Owner'].includes(attrType)) {
        valEl = document.createElement('input');
        valEl.type  = 'text';
        valEl.value = cond.value;
        valEl.placeholder = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx';
      } else if (['Picklist', 'State', 'Status'].includes(attrType) && cond.attribute) {
        // Async option set select — render a loading placeholder, then populate
        valEl = document.createElement('select');
        valEl.innerHTML = `<option value="">Loading options…</option>`;
        valEl.disabled  = true;
        const capturedIdx1  = idx;
        const capturedAttr1 = cond.attribute;
        const capturedType1 = attrType;
        fetchOptionSetValues(entityMeta.logicalName, capturedAttr1, capturedType1)
          .then(options => {
            valEl.innerHTML = `<option value="">— Select —</option>`;
            options.forEach(o => {
              const opt = new Option(o.label, String(o.value));
              if (String(o.value) === filterConditions[capturedIdx1]?.value) opt.selected = true;
              valEl.appendChild(opt);
            });
            valEl.disabled = false;
            if (filterConditions[capturedIdx1]?.value) valEl.value = filterConditions[capturedIdx1].value;
          })
          .catch(e => {
            console.error('[Data Sync] Failed to load option set values:', e);
            valEl.innerHTML = `<option value="">Failed to load options</option>`;
            valEl.disabled  = true;
          });
      } else if (attrType === 'MultiSelectPicklist' && cond.attribute) {
        // Custom checkbox dropdown — handled entirely inside buildMultiSelectDropdown
        valWrap.appendChild(buildMultiSelectDropdown(idx, cond.attribute, attrType));
        // Skip generic valEl listener attachment below
        valEl = null;
      } else {
        valEl = document.createElement('input');
        valEl.type  = 'text';
        valEl.value = cond.value;
        valEl.placeholder = 'Value';
      }
      if (valEl) {
        valEl.style.width = '100%';
        const readValue = () => valEl.value;
        valEl.addEventListener('change', () => { filterConditions[idx].value = readValue(); updateFetchXmlPreview(); });
        valEl.addEventListener('input',  () => { filterConditions[idx].value = readValue(); updateFetchXmlPreview(); });
        valWrap.appendChild(valEl);
      }
    }

    // Remove button
    const removeBtn = document.createElement('button');
    removeBtn.className = 'btn-remove';
    removeBtn.title = 'Remove condition';
    removeBtn.textContent = '×';
    removeBtn.addEventListener('click', () => {
      filterConditions.splice(idx, 1);
      renderFilterConditions();
      updateFetchXmlPreview();
    });

    row.appendChild(attrSel);
    row.appendChild(opSel);
    row.appendChild(valWrap);
    row.appendChild(removeBtn);
    container.appendChild(row);
  });
}

function getOperators(attrType) {
  const common = [['null', 'Is Null'], ['not-null', 'Is Not Null']];
  switch (attrType) {
    case 'String':
    case 'Memo':
      return [['eq','Equals'],['ne','Not Equals'],['like','Contains'],['not-like','Does Not Contain'],['begins-with','Begins With'],['ends-with','Ends With'], ...common];
    case 'Integer':
    case 'Decimal':
    case 'Double':
    case 'Money':
    case 'BigInt':
      return [['eq','Equals'],['ne','Not Equals'],['gt','Greater Than'],['ge','Greater or Equal'],['lt','Less Than'],['le','Less or Equal'], ...common];
    case 'DateTime':
      return [
        ['eq','Equals (exact date)'], ['ne','Not Equals'], ['gt','After'], ['lt','Before'],
        ['on-or-before','On or Before'], ['on-or-after','On or After'],
        // Relative — no value needed
        ['today','Today'], ['yesterday','Yesterday'], ['tomorrow','Tomorrow'],
        ['this-week','This Week'], ['last-week','Last Week'], ['next-week','Next Week'],
        ['this-month','This Month'], ['last-month','Last Month'], ['next-month','Next Month'],
        ['this-year','This Year'], ['last-year','Last Year'], ['next-year','Next Year'],
        ['last-7-days','Last 7 Days'], ['next-7-days','Next 7 Days'],
        ['last-30-days','Last 30 Days'], ['next-30-days','Next 30 Days'],
        // Relative — number value required
        ['last-x-days','Last X Days'], ['next-x-days','Next X Days'],
        ['last-x-weeks','Last X Weeks'], ['next-x-weeks','Next X Weeks'],
        ['last-x-months','Last X Months'], ['next-x-months','Next X Months'],
        ['last-x-years','Last X Years'], ['next-x-years','Next X Years'],
        ['older-than-x-days','Older Than X Days'], ['older-than-x-weeks','Older Than X Weeks'],
        ['older-than-x-months','Older Than X Months'], ['older-than-x-years','Older Than X Years'],
        ...common,
      ];
    case 'Boolean':
      return [['eq','Equals'], ...common];
    case 'Picklist':
    case 'State':
    case 'Status':
      return [['eq','Equals'],['ne','Not Equals'], ...common];
    case 'MultiSelectPicklist':
      return [['contain-values','Contains Values'],['not-contain-values','Does Not Contain Values'], ...common];
    case 'Lookup':
    case 'Customer':
    case 'Owner':
      return [['eq','Equals'],['ne','Not Equals'], ...common];
    default:
      return [['eq','Equals'],['ne','Not Equals'], ...common];
  }
}

/**
 * Fetches option set values (Picklist / State / Status) from D365 via the bridge.
 * Results are cached so repeat renders don't re-fetch.
 * Returns an array of { value: number, label: string }.
 */
async function fetchOptionSetValues(etn, attrName, attrType) {
  const cacheKey = `${etn}|${attrName}`;
  if (_optionSetCache.has(cacheKey)) return _optionSetCache.get(cacheKey);

  const typeSegment = attrType === 'Picklist'             ? 'PicklistAttributeMetadata'
                    : attrType === 'MultiSelectPicklist' ? 'MultiSelectPicklistAttributeMetadata'
                    : attrType === 'State'               ? 'StateAttributeMetadata'
                    :                                      'StatusAttributeMetadata';

  const apiBase = `${sourceEnv.url}/api/data/${settings.apiVersion}`;
  // Use $expand=OptionSet — $select alone does not inline the Options collection
  const url = `${apiBase}/EntityDefinitions(LogicalName='${etn}')/Attributes(LogicalName='${attrName}')/Microsoft.Dynamics.CRM.${typeSegment}?$expand=OptionSet`;

  let options;
  try {
    const d = await _bridgeFetch(url);
    const opts = d.OptionSet?.Options ?? d.GlobalOptionSet?.Options ?? [];
    options = opts.map(o => ({
      value: o.Value,
      label: o.Label?.UserLocalizedLabel?.Label ?? String(o.Value),
    }));
  } catch (e) {
    console.error('[Data Sync] fetchOptionSetValues failed:', e);
    throw e;
  }

  _optionSetCache.set(cacheKey, options);
  return options;
}

function buildFetchXml(forCount = false) {
  const etn    = entityMeta.logicalName;
  const idAttr = entityMeta.primaryIdAttribute;

  const fetchAttrs = forCount
    ? `    <attribute name="${idAttr}" alias="cnt" aggregate="count" />`
    : `    <all-attributes />`;

  const aggregateAttr = forCount ? ` aggregate="true"` : '';

  const filterLogic = document.getElementById('filter-logic')?.value ?? 'and';
  const conditions  = filterConditions
    .filter(c => c.attribute)
    .map(c => {
      if (['null', 'not-null'].includes(c.operator) || DATE_NO_VALUE_OPS.has(c.operator)) {
        return `      <condition attribute="${c.attribute}" operator="${c.operator}" />`;
      }
      if (['contain-values', 'not-contain-values'].includes(c.operator)) {
        // Compact form — no whitespace inside the condition element to avoid D365 parsing issues
        const vals = (c.value ?? '').split(',').map(v => v.trim()).filter(Boolean)
          .map(v => `<value>${escXml(v)}</value>`).join('');
        return vals
          ? `      <condition attribute="${c.attribute}" operator="${c.operator}">${vals}</condition>`
          : `      <condition attribute="${c.attribute}" operator="${c.operator}" />`;
      }
      return `      <condition attribute="${c.attribute}" operator="${c.operator}" value="${escXml(c.value)}" />`;
    }).join('\n');

  const filterXml = conditions
    ? `\n  <filter type="${filterLogic}">\n${conditions}\n  </filter>`
    : '';

  return `<fetch${aggregateAttr}>\n  <entity name="${etn}">\n${fetchAttrs}${filterXml}\n  </entity>\n</fetch>`;
}

function updateFetchXmlPreview() {
  const el = document.getElementById('fetchxml-code');
  if (el) el.textContent = buildFetchXml(false);
}

// ─── Step 5: View Updates ──────────────────────────────────────────────────────

async function runViewUpdates() {
  document.getElementById('results-title').textContent = 'View Updates';
  showResultsView('progress');
  setProgress(0, 'Counting records…');

  const apiBase       = `${sourceEnv.url}/api/data/${settings.apiVersion}`;
  // Single-env mode: target === source. The bridge reaches both the same way.
  const targetApiBase = `${targetEnv.url}/api/data/${settings.apiVersion}`;

  try {
    // 1. Count
    const count = await fetchRecordCount(apiBase);
    const threshold = settings.syncBatchSize ?? 250;

    if (count > threshold) {
      const proceed = await showCountWarning(count, threshold);
      if (!proceed) { goToStep(4); return; }
    }

    // 2. Fetch all source records
    setProgress(10, 'Fetching source records…');
    const records = await fetchAllRecords(apiBase);

    // 3. For each record — check create vs update against target
    setProgress(40, 'Checking target environment…');
    const entitySetName = entityMeta.entitySetName;
    const idAttr        = entityMeta.primaryIdAttribute;
    const nameAttr      = entityMeta.primaryNameAttribute;
    const previewRows   = [];
    const total         = records.length;

    for (let i = 0; i < records.length; i++) {
      const rec = records[i];
      const id  = rec[idAttr];
      const name= rec[nameAttr] ?? id;
      setProgress(40 + Math.round((i / total) * 55), `Checking: ${name}…`);

      const exists = await _bridgeExists(`${targetApiBase}/${entitySetName}(${id})?$select=${idAttr}`);
      previewRows.push({ id, name, action: exists ? 'update' : 'create' });
    }

    setProgress(100, 'Done.');
    renderPreviewTable(previewRows);
    showResultsView('preview');

  } catch (e) {
    console.error('[Data Sync] View Updates error:', e);
    setProgress(0, '');
    showResultsView('progress');
    document.getElementById('progress-status').textContent = `Error: ${e.message}`;
  }
}

function recordUrl(envUrl, logicalName, id) {
  return `${envUrl}/main.aspx?etn=${encodeURIComponent(logicalName)}&id=${encodeURIComponent(id)}&pagetype=entityrecord`;
}

function recordLinkCell(envUrl, logicalName, id, show) {
  if (!show) return '<td></td>';
  const url = escHtml(recordUrl(envUrl, logicalName, id));
  return `<td><a class="record-link" href="${url}" target="_blank" rel="noopener">Open ↗</a></td>`;
}

/**
 * Renders rows into tbody grouped by action, with a collapsible header per group.
 * renderRowFn(r) must return a <tr> element; the group member attribute is set here.
 */
function renderActionGroups(tbody, rows, colspan, renderRowFn) {
  ['create', 'update'].forEach(action => {
    const group = rows.filter(r => r.action === action);
    if (group.length === 0) return;

    const label      = action === 'create' ? 'Create' : 'Update';
    const badgeClass = action === 'create' ? 'badge--create' : 'badge--update';

    const headerTr = document.createElement('tr');
    headerTr.className = 'group-row';
    headerTr.dataset.group = action;
    headerTr.innerHTML = `
      <td colspan="${colspan}">
        <button type="button" class="group-toggle">
          <span class="group-arrow">▾</span>
          <span class="badge ${badgeClass}">${label}</span>
          <span class="group-count">${group.length} record${group.length !== 1 ? 's' : ''}</span>
        </button>
      </td>`;
    headerTr.querySelector('.group-toggle').addEventListener('click', () => {
      const collapsed = headerTr.dataset.collapsed === 'true';
      headerTr.dataset.collapsed = String(!collapsed);
      headerTr.querySelector('.group-arrow').textContent = collapsed ? '▾' : '▸';
      tbody.querySelectorAll(`tr[data-group-member="${action}"]`).forEach(r => {
        r.style.display = collapsed ? '' : 'none';
      });
    });
    tbody.appendChild(headerTr);

    group.forEach(r => {
      const tr = renderRowFn(r);
      tr.dataset.groupMember = action;
      tbody.appendChild(tr);
    });
  });
}

function renderPreviewTable(rows) {
  const tbody = document.getElementById('preview-tbody');
  tbody.innerHTML = '';

  renderActionGroups(tbody, rows, 4, r => {
    const tr = document.createElement('tr');
    const srcCell = recordLinkCell(sourceEnv.url, entityMeta.logicalName, r.id, true);
    const tgtCell = recordLinkCell(targetEnv.url, entityMeta.logicalName, r.id, r.action === 'update');
    tr.innerHTML = `<td>${escHtml(r.name)}</td><td class="cell-id">${r.id}</td>${srcCell}${tgtCell}`;
    return tr;
  });

  const creates = rows.filter(r => r.action === 'create').length;
  const updates = rows.filter(r => r.action === 'update').length;
  document.getElementById('preview-stats').textContent =
    `${rows.length} record${rows.length !== 1 ? 's' : ''} — ${creates} to create, ${updates} to update`;
}

// ─── Step 5: Sync ─────────────────────────────────────────────────────────────

async function runSync(records = null) {
  document.getElementById('results-title').textContent = 'Sync';
  showResultsView('progress');
  setProgress(0, 'Preparing…');
  syncResults = [];

  const apiBase       = `${sourceEnv.url}/api/data/${settings.apiVersion}`;
  const targetApiBase = `${targetEnv.url}/api/data/${settings.apiVersion}`;

  try {
    // Count check (only if not coming from preview which already checked)
    if (!records) {
      const count     = await fetchRecordCount(apiBase);
      const threshold = settings.syncBatchSize ?? 250;
      if (count > threshold) {
        const proceed = await showCountWarning(count, threshold);
        if (!proceed) { goToStep(4); return; }
      }
    }

    // Fetch all records if not provided
    if (!records) {
      setProgress(5, 'Fetching source records…');
      records = await fetchAllRecords(apiBase);
    }

    // Fetch target schema (attrs + navMap)
    setProgress(15, 'Fetching target schema…');
    const etn = entityMeta.logicalName;
    let targetSchema;
    try {
      const [attrD, m2oD] = await Promise.all([
        _bridgeFetch(`${targetApiBase}/EntityDefinitions(LogicalName='${etn}')/Attributes?$select=LogicalName,IsValidForCreate`),
        _bridgeFetch(`${targetApiBase}/EntityDefinitions(LogicalName='${etn}')/ManyToOneRelationships?$select=ReferencingAttribute,ReferencingEntityNavigationPropertyName`),
      ]);
      const attrNames = (attrD.value ?? []).filter(a => a.IsValidForCreate).map(a => a.LogicalName);
      const navMap    = Object.fromEntries((m2oD.value ?? []).map(r => [r.ReferencingAttribute, r.ReferencingEntityNavigationPropertyName]));
      targetSchema = { attrNames, navMap };
    } catch (e) {
      console.error('[Data Sync] Failed to fetch target schema:', e);
      targetSchema = { attrNames: null, navMap: {} };
    }

    const tAttrSet = targetSchema?.attrNames ? new Set(targetSchema.attrNames) : null;

    // Fetch source navMap (for lookups)
    let sourceNavMap = {};
    try {
      const d = await _bridgeFetch(
        `${apiBase}/EntityDefinitions(LogicalName='${etn}')/ManyToOneRelationships?$select=ReferencingAttribute,ReferencingEntityNavigationPropertyName`
      );
      sourceNavMap = Object.fromEntries((d.value ?? []).map(r => [r.ReferencingAttribute, r.ReferencingEntityNavigationPropertyName]));
    } catch (e) {
      console.error('[Data Sync] Failed to fetch source navMap:', e);
    }

    const targetEntitySetName = await fetchEntitySetName(targetApiBase, etn);
    if (!targetEntitySetName) throw new Error(`Cannot resolve entity set name for '${etn}' in target environment.`);

    // Process records
    const total = records.length;
    const m2mToSync = selectedM2m === null
      ? allM2m
      : allM2m.filter(r => selectedM2m.has(r.SchemaName));

    for (let i = 0; i < records.length; i++) {
      const rec  = records[i];
      const idAttr   = entityMeta.primaryIdAttribute;
      const nameAttr = entityMeta.primaryNameAttribute;
      const id       = rec[idAttr];
      const name     = rec[nameAttr] ?? id;

      setProgress(20 + Math.round((i / total) * 75), `Processing: ${name}…`);

      const result = await syncRecord(
        rec, id, name, apiBase, targetApiBase,
        targetEntitySetName, tAttrSet, targetSchema?.navMap ?? {},
        sourceNavMap, m2mToSync
      );
      syncResults.push(result);
    }

    setProgress(100, 'Sync complete.');
    renderSummaryTable();
    showResultsView('summary');

  } catch (e) {
    console.error('[Data Sync] Sync error:', e);
    setProgress(0, '');
    document.getElementById('progress-status').textContent = `Error: ${e.message}`;
  }
}

// ─── Sync engine ──────────────────────────────────────────────────────────────

const EXCLUDE = new Set([
  'createdon', 'modifiedon', 'overriddencreatedon',
  'createdby', 'modifiedby', 'createdonbehalfby', 'modifiedonbehalfby',
  'versionnumber', 'exchangerate', 'importsequencenumber',
  'timezoneruleversionnumber', 'utcconversiontimezonecode',
  'owningbusinessunit', 'owninguser', 'owningteam',
  // State/status must be set as a separate final PATCH after all field updates,
  // because sending statuscode alongside other fields while the state differs
  // causes a 400 "not a valid status code for state" error.
  'statecode', 'statuscode',
]);
const SKIP_TYPES = new Set(['Virtual', 'EntityName', 'ManagedProperty', 'Uniqueidentifier']);

// Standard write headers for D365 OData PATCH/POST.
const WRITE_HEADERS = {
  'Content-Type': 'application/json',
  'Accept': 'application/json',
  'OData-MaxVersion': '4.0',
  'OData-Version': '4.0',
  'MSCRM.SuppressDuplicateDetection': 'true',
};

async function syncRecord(
  record, id, name,
  sourceApiBase, targetApiBase,
  targetEntitySetName, tAttrSet, targetNavMap, sourceNavMap,
  m2mRels
) {
  const etn      = entityMeta.logicalName;
  const idAttr   = entityMeta.primaryIdAttribute;
  let action = 'update'; // will be overwritten if PATCH returns 404-style create

  try {
    // Build scalar payload + lookup fields
    const scalarPayload = {};
    const lookupFields  = []; // { logicalName, guid, targetEtn }

    for (const attr of allAttributes) {
      if (!attr.IsValidForCreate) continue;
      const attrName = attr.LogicalName;
      if (EXCLUDE.has(attrName)) continue;
      if (attrName === idAttr) continue;
      // Column filter
      if (selectedAttrs !== null && !selectedAttrs.has(attrName)) continue;
      // Target schema filter
      if (tAttrSet && !tAttrSet.has(attrName)) continue;

      const effectiveType = getEffectiveAttrType(attr);
      if (SKIP_TYPES.has(effectiveType)) continue;
      if (effectiveType === 'Lookup' || effectiveType === 'Customer' || effectiveType === 'Owner') {
        const guidKey   = `_${attrName}_value`;
        const guid      = record[guidKey];
        if (!guid) continue;
        const targetEtn = record[`${guidKey}@Microsoft.Dynamics.CRM.lookuplogicalname`];
        if (targetEtn) lookupFields.push({ logicalName: attrName, guid, targetEtn });
      } else {
        const val = record[attrName];
        if (val === null || val === undefined) continue;
        scalarPayload[attrName] = val;
      }
    }

    // Determine create vs update by checking if record exists in target
    const existsInTarget = await _bridgeExists(`${targetApiBase}/${targetEntitySetName}(${id})?$select=${idAttr}`);
    action = existsInTarget ? 'update' : 'create';

    // Resolve lookup entity set names in target
    const lookupPayload = {};
    for (const { logicalName, guid, targetEtn } of lookupFields) {
      const targetLookupSet = await fetchEntitySetName(targetApiBase, targetEtn);
      if (!targetLookupSet) continue;
      const navProp = targetNavMap[logicalName] ?? sourceNavMap[logicalName] ?? logicalName;
      lookupPayload[`${navProp}@odata.bind`] = `/${targetLookupSet}(${guid})`;
    }

    // PATCH scalar (self-healing retry — drops fields the target rejects as missing)
    if (Object.keys(scalarPayload).length > 0) {
      let remaining = { ...scalarPayload };
      let lastError = null;
      const patchUrl = `${targetApiBase}/${targetEntitySetName}(${id})`;

      for (let attempt = 0; attempt < 30; attempt++) {
        if (Object.keys(remaining).length === 0) break;
        try {
          await _bridgeFetch(patchUrl, WRITE_HEADERS, 'PATCH', remaining);
          lastError = null;
          break;
        } catch (e) {
          const errText = e.message || '';
          lastError = `Upsert failed: ${errText}`;
          const m = errText.match(/[`'"]?([a-z_][a-z0-9_]*)[`'"]?\s+(?:field\s+)?missing\s+from\s+target\s+entity/i)
                 ?? errText.match(/missing\s+from\s+target\s+entity[^:]*:\s*([a-z_][a-z0-9_]*)/i);
          const badField = m?.[1];
          if (badField && badField in remaining) { delete remaining[badField]; continue; }
          throw new Error(lastError);
        }
      }
      if (lastError) throw new Error(lastError);
    }

    // PATCH each lookup individually (skip on failure)
    for (const [bindKey, bindVal] of Object.entries(lookupPayload)) {
      try {
        await _bridgeFetch(
          `${targetApiBase}/${targetEntitySetName}(${id})`,
          WRITE_HEADERS, 'PATCH', { [bindKey]: bindVal }
        );
      } catch (_) { /* skip lookups the target can't resolve */ }
    }

    // N:N delta sync
    for (const rel of m2mRels) {
      const isEntity1  = rel.Entity1LogicalName === etn;
      const myNavProp  = isEntity1 ? rel.Entity1NavigationPropertyName : rel.Entity2NavigationPropertyName;
      const relatedEtn = isEntity1 ? rel.Entity2LogicalName            : rel.Entity1LogicalName;
      if (!myNavProp || !relatedEtn) continue;

      const relatedSetTarget = await fetchEntitySetName(targetApiBase, relatedEtn);
      if (!relatedSetTarget) continue;

      const pkField = `${relatedEtn}id`;

      let sourceIds = [];
      try {
        const d = await _bridgeFetch(`${sourceApiBase}/${entityMeta.entitySetName}(${id})/${myNavProp}?$top=500&$select=${pkField}`);
        sourceIds = (d.value ?? []).map(r => r[pkField]).filter(Boolean);
      } catch (_) { sourceIds = []; }

      let targetIds = [];
      try {
        const d = await _bridgeFetch(`${targetApiBase}/${targetEntitySetName}(${id})/${myNavProp}?$top=500&$select=${pkField}`);
        targetIds = (d.value ?? []).map(r => r[pkField]).filter(Boolean);
      } catch (_) { targetIds = []; }

      const tgtSet = new Set(targetIds);
      const srcSet = new Set(sourceIds);

      for (const rid of sourceIds) {
        if (tgtSet.has(rid)) continue;
        try {
          await _bridgeFetch(
            `${targetApiBase}/${targetEntitySetName}(${id})/${myNavProp}/$ref`,
            { 'Content-Type': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' },
            'POST',
            { '@odata.id': `${targetApiBase}/${relatedSetTarget}(${rid})` }
          );
        } catch (_) { /* skip */ }
      }

      for (const rid of targetIds) {
        if (srcSet.has(rid)) continue;
        try {
          await _bridgeFetch(
            `${targetApiBase}/${targetEntitySetName}(${id})/${myNavProp}(${rid})/$ref`,
            { 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' },
            'DELETE'
          );
        } catch (_) { /* skip */ }
      }
    }

    // Apply state/status as a final separate PATCH.
    // Sending statecode+statuscode alongside other fields causes a 400 if the target
    // is in a different state, because D365 validates statuscode against the *current*
    // statecode before applying the new one.
    const srcStateCode  = record['statecode'];
    const srcStatusCode = record['statuscode'];
    const shouldSyncState = (srcStateCode !== null && srcStateCode !== undefined)
      && (selectedAttrs === null || selectedAttrs.has('statecode'));

    if (shouldSyncState) {
      try {
        await _bridgeFetch(
          `${targetApiBase}/${targetEntitySetName}(${id})`,
          WRITE_HEADERS, 'PATCH',
          { statecode: srcStateCode, statuscode: srcStatusCode }
        );
      } catch (e) {
        throw new Error(`State sync failed: ${e.message}`);
      }
    }

    return { id, name, action, status: 'success', error: null };

  } catch (e) {
    console.error('[Data Sync] syncRecord failed for', id, ':', e);
    return { id, name, action, status: 'failed', error: e.message };
  }
}

// ─── Summary table ────────────────────────────────────────────────────────────

function renderSummaryTable() {
  const succeeded = syncResults.filter(r => r.status === 'success').length;
  const failed    = syncResults.filter(r => r.status === 'failed').length;
  document.getElementById('summary-stats').textContent =
    `${syncResults.length} records — ${succeeded} succeeded, ${failed} failed`;

  bindSummaryControls();
  renderFilteredSummary();
}

function bindSummaryControls() {
  document.getElementById('summary-search').addEventListener('input', renderFilteredSummary);
  document.getElementById('summary-status-filter').addEventListener('change', renderFilteredSummary);
  document.getElementById('btn-export-csv').addEventListener('click', exportCsv);
  document.getElementById('btn-new-sync').addEventListener('click', () => location.reload());

  document.querySelectorAll('#summary-table .sortable').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.col;
      if (sortCol === col) sortDir = sortDir === 'asc' ? 'desc' : 'asc';
      else { sortCol = col; sortDir = 'asc'; }
      document.querySelectorAll('#summary-table .sortable').forEach(h => {
        h.querySelector('.sort-arrow').textContent = h.dataset.col === sortCol
          ? (sortDir === 'asc' ? '▲' : '▼')
          : '↕';
      });
      renderFilteredSummary();
    });
  });
}

function renderFilteredSummary() {
  const query      = document.getElementById('summary-search').value.trim().toLowerCase();
  const statusFilt = document.getElementById('summary-status-filter').value;

  let rows = syncResults.filter(r => {
    if (statusFilt && r.status !== statusFilt) return false;
    if (query) {
      return r.name.toLowerCase().includes(query)
          || r.id.toLowerCase().includes(query)
          || (r.error ?? '').toLowerCase().includes(query);
    }
    return true;
  });

  // Sort
  rows = [...rows].sort((a, b) => {
    let av = a[sortCol] ?? '', bv = b[sortCol] ?? '';
    if (typeof av === 'string') av = av.toLowerCase();
    if (typeof bv === 'string') bv = bv.toLowerCase();
    if (av < bv) return sortDir === 'asc' ? -1 : 1;
    if (av > bv) return sortDir === 'asc' ? 1 : -1;
    return 0;
  });

  const tbody = document.getElementById('summary-tbody');
  tbody.innerHTML = '';

  renderActionGroups(tbody, rows, 6, r => {
    const tr = document.createElement('tr');
    const statusBadge = r.status === 'success'
      ? '<span class="badge badge--success">✓ Success</span>'
      : '<span class="badge badge--failed">✗ Failed</span>';
    const srcCell = recordLinkCell(sourceEnv.url, entityMeta.logicalName, r.id, true);
    const showTgt = r.action === 'update' || (r.action === 'create' && r.status === 'success');
    const tgtCell = recordLinkCell(targetEnv.url, entityMeta.logicalName, r.id, showTgt);
    tr.innerHTML = `
      <td>${escHtml(r.name)}</td>
      <td class="cell-id">${r.id}</td>
      <td>${statusBadge}</td>
      ${srcCell}
      ${tgtCell}
      <td><span class="cell-error" title="${escHtml(r.error ?? '')}">${escHtml(r.error ?? '')}</span></td>`;
    return tr;
  });
}

function exportCsv() {
  const cols = ['Name', 'Record ID', 'Action', 'Status', 'Error'];
  const rows = syncResults.map(r => [
    csvEscape(r.name),
    csvEscape(r.id),
    csvEscape(r.action),
    csvEscape(r.status),
    csvEscape(r.error ?? ''),
  ]);
  const csv  = [cols, ...rows].map(r => r.join(',')).join('\r\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `sync-results-${new Date().toISOString().replace(/[:.]/g, '-')}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function csvEscape(val) {
  const s = String(val ?? '');
  return s.includes(',') || s.includes('"') || s.includes('\n')
    ? `"${s.replace(/"/g, '""')}"`
    : s;
}

// ─── Shared helpers ────────────────────────────────────────────────────────────

// EntitySetName cache
const _entitySetCache = new Map();

async function fetchEntitySetName(apiBase, etn) {
  const key = `${apiBase}:${etn}`;
  if (_entitySetCache.has(key)) return _entitySetCache.get(key);
  let result = null;
  try {
    const d = await _bridgeFetch(`${apiBase}/EntityDefinitions(LogicalName='${etn}')?$select=EntitySetName`);
    result = d.EntitySetName ?? null;
  } catch (e) {
    console.error('[Data Sync] fetchEntitySetName failed for', etn, ':', e);
    result = null;
  }
  if (result) _entitySetCache.set(key, result);
  return result;
}

async function fetchRecordCount(apiBase) {
  const fetchXml  = buildFetchXml(true);
  const entitySet = entityMeta.entitySetName;
  const encoded   = encodeURIComponent(fetchXml);
  const url       = `${apiBase}/${entitySet}?fetchXml=${encoded}`;

  try {
    const d = await _bridgeFetch(url);
    return parseInt(d.value?.[0]?.cnt ?? '0', 10) || 0;
  } catch (e) {
    console.error('[Data Sync] fetchRecordCount failed:', e);
    throw e;
  }
}

async function fetchAllRecords(apiBase) {
  const fetchXml  = buildFetchXml(false);
  const entitySet = entityMeta.entitySetName;
  const encoded   = encodeURIComponent(fetchXml);
  let url = `${apiBase}/${entitySet}?fetchXml=${encoded}`;

  const headers = {
    Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue,Microsoft.Dynamics.CRM.lookuplogicalname",odata.maxpagesize=250',
  };

  const allRecords = [];
  try {
    while (url) {
      const d = await _bridgeFetch(url, headers);
      allRecords.push(...(d.value ?? []));
      url = d['@odata.nextLink'] ?? null;
    }
    return allRecords;
  } catch (e) {
    console.error('[Data Sync] fetchAllRecords failed:', e);
    throw e;
  }
}

// ─── Count warning ─────────────────────────────────────────────────────────────

function showCountWarning(count, threshold) {
  return new Promise(resolve => {
    const warningEl = document.getElementById('count-warning');
    const msgEl     = document.getElementById('count-warning-msg');
    msgEl.textContent =
      `Warning: ${count} records found. This exceeds the configured threshold of ${threshold}. ` +
      `Syncing a large number of records may take a long time.`;
    warningEl.classList.remove('hidden');

    const cont = document.getElementById('btn-warning-continue');
    const cancel = document.getElementById('btn-warning-cancel');

    const cleanup = () => { warningEl.classList.add('hidden'); cont.onclick = null; cancel.onclick = null; };
    cont.onclick   = () => { cleanup(); resolve(true); };
    cancel.onclick = () => { cleanup(); resolve(false); };
  });
}

// ─── UI helpers ───────────────────────────────────────────────────────────────

const _FOOTER_BTNS = [
  'btn-step1-next','btn-step2-back','btn-step2-next',
  'btn-step3-back','btn-step3-next',
  'btn-step4-back','btn-view-updates','btn-sync',
  'btn-back-from-preview','btn-sync-from-preview','btn-new-sync',
];

function updateFooterButtons(step, subView = null) {
  // Hide all footer buttons first
  _FOOTER_BTNS.forEach(id => document.getElementById(id)?.classList.add('hidden'));

  const show = id => document.getElementById(id)?.classList.remove('hidden');

  switch (step) {
    case 1:
      show('btn-step1-next');
      break;
    case 2:
      show('btn-step2-back');
      show('btn-step2-next');
      break;
    case 3:
      show('btn-step3-back');
      show('btn-step3-next');
      break;
    case 4:
      show('btn-step4-back');
      show('btn-view-updates');
      show('btn-sync');
      break;
    case 5:
      // Back is always visible on step 5 so the user can correct configuration
      show('btn-back-from-preview');
      if (subView === 'preview') show('btn-sync-from-preview');
      if (subView === 'summary') show('btn-new-sync');
      break;
  }

  // Show separator between nav and action groups only when action buttons are visible
  const actionButtonsVisible = ['btn-view-updates','btn-sync','btn-sync-from-preview','btn-new-sync']
    .some(id => !document.getElementById(id)?.classList.contains('hidden'));
  document.getElementById('footer-sep-actions')?.classList.toggle('hidden', !actionButtonsVisible);
}

function goToStep(step) {
  currentStep = step;

  // Update stepper
  document.querySelectorAll('.step').forEach(el => {
    const n = parseInt(el.dataset.step, 10);
    el.classList.toggle('active', n === step);
    el.classList.toggle('done',   n < step);
  });

  // Show/hide panels
  document.querySelectorAll('.step-panel').forEach(el => {
    el.classList.toggle('active', el.id === `panel-${step}`);
  });

  updateFooterButtons(step);
  updateStepSummary(step);
}

/**
 * Navigate back to a previous step, cascading state resets as needed.
 * - Going back past step 2: re-enable env dropdowns
 * - Going back past step 3: clear table/columns/M2M/filter/results state
 * - Going back past step 4: clear filter conditions
 * - Going back past step 5: clear sync results
 */
async function navigateBack(targetStep) {
  if (targetStep >= currentStep) return;

  // Re-enable env dropdowns when leaving the table/sync flow entirely
  if (targetStep < 2) {
    document.getElementById('source-env').disabled = false;
    document.getElementById('target-env').disabled = false;
    setStep1Loading(false);
  }

  // Cascade resets
  if (targetStep < 5) syncResults = [];
  if (targetStep < 4) filterConditions = [];
  if (targetStep < 3) {
    entityMeta    = null;
    allAttributes = [];
    allM2m        = [];
    selectedAttrs = null;
    selectedM2m   = null;
    document.getElementById('table-search').value = '';
    document.querySelectorAll('#table-list li').forEach(li => li.classList.remove('selected'));
    document.getElementById('btn-step2-next').disabled = true;
  }

  goToStep(targetStep);
}

function updateStepSummary(step) {
  const el = document.getElementById('step-summary');
  if (step <= 1) { el.classList.add('hidden'); return; }

  const parts = [];

  // Environments (always shown from step 2)
  if (sourceEnv) parts.push(`<span class="crumb"><span class="crumb-label">Source</span> <span class="crumb-value">${escHtml(sourceEnv.name)}</span></span>`);
  if (targetEnv) parts.push(`<span class="crumb"><span class="crumb-label">Target</span> <span class="crumb-value">${escHtml(targetEnv.name)}</span></span>`);

  // Table (step 3+)
  if (step >= 3 && entityMeta) {
    parts.push(`<span class="crumb"><span class="crumb-label">Table</span> <span class="crumb-value">${escHtml(entityMeta.displayName)}</span></span>`);
  }

  // Columns & N:N (step 4+)
  if (step >= 4) {
    const SKIP = new Set(['Virtual', 'EntityName', 'ManagedProperty', 'Uniqueidentifier']);
    const totalCols = allAttributes.filter(a => !SKIP.has(a.AttributeType)).length;
    const selCols   = selectedAttrs === null ? totalCols : selectedAttrs.size;
    parts.push(`<span class="crumb"><span class="crumb-label">Columns</span> <span class="crumb-value">${selCols} / ${totalCols}</span></span>`);
    if (allM2m.length > 0) {
      const selM2m = selectedM2m === null ? allM2m.length : selectedM2m.size;
      parts.push(`<span class="crumb"><span class="crumb-label">N:N</span> <span class="crumb-value">${selM2m} / ${allM2m.length}</span></span>`);
    }
  }

  // Filter (step 5)
  if (step >= 5) {
    const n = filterConditions.filter(c => c.attribute).length;
    parts.push(`<span class="crumb"><span class="crumb-label">Filter</span> <span class="crumb-value">${n === 0 ? 'None' : `${n} condition${n !== 1 ? 's' : ''}`}</span></span>`);
  }

  el.innerHTML = parts.join('');
  el.classList.remove('hidden');
}

function showResultsView(view) {
  document.getElementById('progress-view').classList.toggle('hidden', view !== 'progress');
  document.getElementById('count-warning').classList.add('hidden');
  document.getElementById('preview-view').classList.toggle('hidden', view !== 'preview');
  document.getElementById('summary-view').classList.toggle('hidden', view !== 'summary');
  // Keep Back always visible; show view-specific action buttons
  updateFooterButtons(5, view);
}

function setProgress(pct, msg) {
  document.getElementById('progress-bar').style.width = `${pct}%`;
  document.getElementById('progress-status').textContent = msg;
}

function bindStep5() {
  document.getElementById('btn-back-from-preview').addEventListener('click', () => {
    goToStep(4);
  });
  document.getElementById('btn-sync-from-preview').addEventListener('click', async () => {
    showResultsView('progress');
    const apiBase  = `${sourceEnv.url}/api/data/${settings.apiVersion}`;
    setProgress(5, 'Fetching source records…');
    const records = await fetchAllRecords(apiBase);
    await runSync(records);
  });
}

function escHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escXml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
