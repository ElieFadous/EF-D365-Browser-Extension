/**
 * EF Power Platform Tools — Options Page Script
 *
 * Manages the list of Dynamics 365 environments stored in chrome.storage.local.
 * The stored list is the source of truth; environments.json is only the default.
 */

// ─── Feature definitions ──────────────────────────────────────────────────────

const FEATURES = [
  { id: 'goto-open-in',   label: 'Go To — Open In…' },
  { id: 'goto-api',       label: 'Go To — API' },
  { id: 'goto-solutions', label: 'Go To — Solutions' },
  { id: 'goto-config',    label: 'Go To — Config' },
  { id: 'goto-security',  label: 'Go To — Security' },
  { id: 'metadata',       label: 'Metadata Browser' },
  { id: 'ribbon',         label: 'Ribbon Buttons' },
  { id: 'plugin-trace',   label: 'Plugin Trace Logs' },
  { id: 'flows',          label: 'Power Automate Flows' },
];

/** Returns true if the environment is enabled for the given feature.
 *  Absent or empty enabledFor means ALL features are on (backward-compatible). */
function isEnabledFor(env, featureId) {
  if (!Array.isArray(env.enabledFor) || env.enabledFor.length === 0) return true;
  return env.enabledFor.includes(featureId);
}

// ─── State ────────────────────────────────────────────────────────────────────

let environments = [];
let settings     = { apiVersion: 'v9.2' };

// ─── Bootstrap ────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {
  [environments, settings] = await Promise.all([loadEnvironments(), loadSettings()]);
  renderTable();
  renderSettingsForm();
  bindControls();
});

// ─── Load / Save ─────────────────────────────────────────────────────────────

async function loadEnvironments() {
  const stored = await chrome.storage.local.get('environments');
  if (Array.isArray(stored.environments) && stored.environments.length > 0) {
    return stored.environments;
  }
  // First run — seed from bundled defaults.
  const res  = await fetch(chrome.runtime.getURL('../environments.json'));
  const data = await res.json();
  await chrome.storage.local.set({ environments: data.environments });
  return data.environments;
}

async function saveEnvironments() {
  await chrome.storage.local.set({ environments });
}

// ─── Settings Load / Save ─────────────────────────────────────────────────────

async function loadSettings() {
  const stored = await chrome.storage.local.get('settings');
  return {
    apiVersion:           'v9.2',
    clonePrefix:          '',
    cloneWhitelist:       null,
    cloneLookupMode:      'skip',
    defaultAppUniqueName: '',
    includedApps:         [],
    syncBatchSize:        250,
    ...stored.settings,
  };
}

async function saveSettings() {
  await chrome.storage.local.set({ settings });
}

function renderSettingsForm() {
  document.getElementById('api-version').value = settings.apiVersion ?? 'v9.2';

  // Clone settings
  document.getElementById('clone-prefix').value = settings.clonePrefix ?? '';
  const scope = settings.cloneWhitelist === null ? 'all' : 'whitelist';
  const scopeRadio = document.querySelector(`input[name="clone-scope"][value="${scope}"]`);
  if (scopeRadio) scopeRadio.checked = true;
  document.getElementById('clone-whitelist').value = (settings.cloneWhitelist ?? []).join('\n');
  document.getElementById('clone-whitelist-row').style.display = scope === 'whitelist' ? '' : 'none';
  const lookupMode = settings.cloneLookupMode ?? 'skip';
  const lookupRadio = document.querySelector(`input[name="clone-lookup-mode"][value="${lookupMode}"]`);
  if (lookupRadio) lookupRadio.checked = true;

  // Toggle whitelist visibility on radio change
  document.querySelectorAll('input[name="clone-scope"]').forEach(r =>
    r.addEventListener('change', () => {
      document.getElementById('clone-whitelist-row').style.display =
        r.value === 'whitelist' ? '' : 'none';
    })
  );

  // Data Sync settings
  document.getElementById('sync-batch-size').value = settings.syncBatchSize ?? 250;
}

// ─── Feature chip helpers ─────────────────────────────────────────────────────

/** Render feature checkboxes into #field-enabled-for.
 *  enabledFor === null/undefined → all checked. */
function _renderFeatureChips(enabledFor) {
  const grid  = document.getElementById('field-enabled-for');
  grid.innerHTML = '';
  const isAll = !Array.isArray(enabledFor) || enabledFor.length === 0;

  FEATURES.forEach(f => {
    const lbl = document.createElement('label');
    lbl.className        = 'feature-chip';
    lbl.dataset.tooltip  = f.label;   // used by the CSS ::after tooltip

    const cb = document.createElement('input');
    cb.type    = 'checkbox';
    cb.value   = f.id;
    cb.checked = isAll || enabledFor.includes(f.id);

    lbl.appendChild(cb);
    lbl.appendChild(document.createTextNode(f.label));
    grid.appendChild(lbl);
  });
}

/** Table cell content: "All" badge or "N / 7" badge. */
function _featuresSummary(env) {
  if (!Array.isArray(env.enabledFor) || env.enabledFor.length === 0) {
    return '<span class="features-badge features-badge--all">All</span>';
  }
  const n = env.enabledFor.length;
  const m = FEATURES.length;
  const tip = escHtml(env.enabledFor.map(id => FEATURES.find(f => f.id === id)?.label ?? id).join(', '));
  return `<span class="features-badge features-badge--partial" title="${tip}">${n}&thinsp;/&thinsp;${m}</span>`;
}

// ─── Table Rendering ─────────────────────────────────────────────────────────

function renderTable() {
  const tbody = document.getElementById('env-tbody');
  tbody.innerHTML = '';

  if (environments.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="7" style="text-align:center;color:var(--text-muted);padding:24px">
          No environments configured. Click <strong>+ Add Environment</strong> to get started.
        </td>
      </tr>`;
    return;
  }

  environments.forEach((env, index) => {
    const tr = document.createElement('tr');
    tr.dataset.index = index;
    tr.innerHTML = `
      <td>
        <span class="drag-handle" title="Drag to reorder">&#8597;</span>
        <span class="cell-name">${escHtml(env.name)}</span>
      </td>
      <td><span class="cell-url">${escHtml(env.url)}</span></td>
      <td>
        <span class="color-swatch" style="background:${escHtml(env.color ?? '#888')}"></span>
      </td>
      <td>
        ${env.warn ? '<span class="warn-badge">Yes</span>' : '<span style="color:var(--text-muted)">No</span>'}
      </td>
      <td>
        ${env.powerAppsId
          ? `<span class="pa-id-cell" title="${escHtml(env.powerAppsId)}">${escHtml(env.powerAppsId.slice(0, 8))}&hellip;</span>`
          : '<span style="color:var(--text-muted)">—</span>'}
      </td>
      <td>${_featuresSummary(env)}</td>
      <td>
        <div class="actions-cell">
          <button class="btn btn--secondary btn--icon-sm" data-action="edit" data-index="${index}">Edit</button>
          <button class="btn btn--danger    btn--icon-sm" data-action="delete" data-index="${index}">Delete</button>
          <button class="btn btn--secondary btn--icon-sm" data-action="up"   data-index="${index}" ${index === 0 ? 'disabled' : ''} title="Move up">&#8593;</button>
          <button class="btn btn--secondary btn--icon-sm" data-action="down" data-index="${index}" ${index === environments.length - 1 ? 'disabled' : ''} title="Move down">&#8595;</button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });

  // Delegate click events.
  tbody.querySelectorAll('[data-action]').forEach(btn => {
    btn.addEventListener('click', e => {
      const action = e.currentTarget.dataset.action;
      const index  = parseInt(e.currentTarget.dataset.index, 10);
      switch (action) {
        case 'edit':   openModal(index); break;
        case 'delete': deleteEnv(index); break;
        case 'up':     moveEnv(index, -1); break;
        case 'down':   moveEnv(index, +1); break;
      }
    });
  });
}

// ─── CRUD operations ─────────────────────────────────────────────────────────

function deleteEnv(index) {
  const env = environments[index];
  if (!confirm(`Remove "${env.name}" from the list?`)) return;
  environments.splice(index, 1);
  saveAndRender();
  showToast(`"${env.name}" removed.`, 'success');
}

function moveEnv(index, direction) {
  const target = index + direction;
  if (target < 0 || target >= environments.length) return;
  [environments[index], environments[target]] = [environments[target], environments[index]];
  saveAndRender();
}

async function saveAndRender() {
  await saveEnvironments();
  renderTable();
}

// ─── Modal ───────────────────────────────────────────────────────────────────

function openModal(index = -1) {
  const overlay = document.getElementById('modal-overlay');
  const title   = document.getElementById('modal-title');
  const form    = document.getElementById('env-form');
  const err     = document.getElementById('form-error');

  document.getElementById('field-index').value = index;
  err.classList.add('hidden');

  if (index === -1) {
    // Add mode — clear form.
    title.textContent = 'Add Environment';
    form.reset();
    document.getElementById('field-id').value            = '';
    document.getElementById('field-color').value          = '#2563eb';
    document.getElementById('field-power-apps-id').value  = '';
    document.getElementById('field-highlight-tab').checked = false;
    _renderFeatureChips(null);  // all checked by default
  } else {
    // Edit mode — populate form.
    const env = environments[index];
    title.textContent = 'Edit Environment';
    document.getElementById('field-url').value          = env.url;
    document.getElementById('field-id').value           = deriveEnvId(env.url) || env.id;
    document.getElementById('field-name').value         = env.name;
    document.getElementById('field-power-apps-id').value = env.powerAppsId ?? '';
    document.getElementById('field-color').value        = env.color ?? '#2563eb';
    document.getElementById('field-warn').checked          = !!env.warn;
    document.getElementById('field-highlight-tab').checked = !!env.highlightTab;
    _renderFeatureChips(env.enabledFor ?? null);
  }

  overlay.classList.remove('hidden');
  document.getElementById('field-url').focus();
}

function closeModal() {
  document.getElementById('modal-overlay').classList.add('hidden');
}

function bindControls() {
  // Settings left-tab switching.
  document.querySelectorAll('.settings-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.settings-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.settings-panel').forEach(p => p.classList.add('hidden'));
      tab.classList.add('active');
      const panelId = `panel-${tab.dataset.tab}`;
      document.getElementById(panelId)?.classList.remove('hidden');
      // Lazy-init the Model Driven Apps panel on first open
      if (tab.dataset.tab === 'model-driven-apps' && !tab._mdaInit) {
        tab._mdaInit = true;
        initMdaPanel();
      }
      // Lazy-init the Plugin Trace panel on first open
      if (tab.dataset.tab === 'plugin-trace' && !tab._ptInit) {
        tab._ptInit = true;
        initPluginTracePanel();
      }
    });
  });

  // Feature chip All / None shortcuts.
  document.getElementById('feat-all').addEventListener('click', () => {
    document.querySelectorAll('#field-enabled-for input[type="checkbox"]')
      .forEach(cb => { cb.checked = true; });
  });
  document.getElementById('feat-none').addEventListener('click', () => {
    document.querySelectorAll('#field-enabled-for input[type="checkbox"]')
      .forEach(cb => { cb.checked = false; });
  });

  // Save API Settings button — with format validation.
  document.getElementById('btn-save-settings').addEventListener('click', async () => {
    const errEl = document.getElementById('api-version-error');
    const val   = document.getElementById('api-version').value.trim();
    if (!val || !/^v\d+\.\d+$/.test(val)) {
      errEl.textContent = 'Invalid format. Use e.g. v9.2 (v + digits + . + digits).';
      errEl.classList.remove('hidden');
      return;
    }
    errEl.classList.add('hidden');
    settings.apiVersion = val;
    await saveSettings();
    showToast('API settings saved.', 'success');
  });

  document.getElementById('btn-save-clone-settings').addEventListener('click', async () => {
    settings.clonePrefix = document.getElementById('clone-prefix').value.trim();
    const isScopeAll = document.querySelector('input[name="clone-scope"]:checked')?.value !== 'whitelist';
    settings.cloneWhitelist = isScopeAll
      ? null
      : document.getElementById('clone-whitelist').value
          .split(/[\n,]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
    settings.cloneLookupMode = document.querySelector('input[name="clone-lookup-mode"]:checked')?.value ?? 'skip';
    await saveSettings();
    showToast('Clone settings saved.', 'success');
  });

  document.getElementById('btn-save-sync-settings').addEventListener('click', async () => {
    settings.syncBatchSize = parseInt(document.getElementById('sync-batch-size').value, 10) || 250;
    await saveSettings();
    showToast('Data Sync settings saved.', 'success');
  });

  // Auto-derive ID from URL as user types
  document.getElementById('field-url').addEventListener('input', () => {
    document.getElementById('field-id').value = deriveEnvId(
      document.getElementById('field-url').value.trim()
    );
  });

  // Add button.
  document.getElementById('btn-add').addEventListener('click', () => openModal());

  // Cancel button.
  document.getElementById('btn-cancel').addEventListener('click', closeModal);

  // Close on overlay click.
  document.getElementById('modal-overlay').addEventListener('click', e => {
    if (e.target === e.currentTarget) closeModal();
  });

  // Form submit.
  document.getElementById('env-form').addEventListener('submit', async e => {
    e.preventDefault();
    const err = document.getElementById('form-error');
    err.classList.add('hidden');

    const index       = parseInt(document.getElementById('field-index').value, 10);
    const id          = document.getElementById('field-id').value.trim();
    const name        = document.getElementById('field-name').value.trim();
    const url         = document.getElementById('field-url').value.trim().replace(/\/$/, '');
    const powerAppsId  = document.getElementById('field-power-apps-id').value.trim();
    const color       = document.getElementById('field-color').value;
    const warn         = document.getElementById('field-warn').checked;
    const highlightTab = document.getElementById('field-highlight-tab').checked;

    // Collect enabled features; omit the field entirely when all are checked
    // so that future features are automatically enabled for this environment.
    const checkedIds  = Array.from(
      document.querySelectorAll('#field-enabled-for input[type="checkbox"]:checked')
    ).map(cb => cb.value);
    const enabledFor  = checkedIds.length === FEATURES.length ? null : checkedIds;

    // Validate.
    if (!id || !name || !url) {
      showFormError('All fields are required.'); return;
    }
    if (!/^[a-z0-9_-]+$/.test(id)) {
      showFormError('ID may only contain lowercase letters, numbers, hyphens and underscores.'); return;
    }
    try { new URL(url); } catch {
      showFormError('Please enter a valid URL (e.g. https://myenv.crm4.dynamics.com).'); return;
    }

    // Check ID uniqueness (allow same index in edit mode).
    const duplicate = environments.findIndex((env, i) => env.id === id && i !== index);
    if (duplicate !== -1) {
      showFormError(`An environment with ID "${id}" already exists.`); return;
    }

    const entry = {
      id, name, url, color,
      ...(powerAppsId          ? { powerAppsId }           : {}),
      ...(enabledFor           ? { enabledFor }            : {}),
      ...(warn                 ? { warn: true }             : {}),
      ...(highlightTab         ? { highlightTab: true }    : {}),
    };

    if (index === -1) {
      environments.push(entry);
      showToast(`"${name}" added.`, 'success');
    } else {
      environments[index] = entry;
      showToast(`"${name}" updated.`, 'success');
    }

    await saveAndRender();
    closeModal();
  });

  // Export — full backup of all extension storage (environments + settings + ignored types + app cache).
  document.getElementById('btn-export').addEventListener('click', async () => {
    const all  = await chrome.storage.local.get(null); // get every key
    const json = JSON.stringify(all, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `ef-ppt-backup-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('Full backup exported.', 'success');
  });

  // Import — restores whichever keys are present in the file.
  // Supports both old single-key exports ({ environments: [...] })
  // and new full-backup exports (all keys).
  document.getElementById('file-import').addEventListener('change', async e => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const text = await file.text();
      const data = JSON.parse(text);
      if (typeof data !== 'object' || data === null) throw new Error('Invalid backup file.');

      const toRestore = {};
      const restored  = [];

      if (Array.isArray(data.environments)) {
        toRestore.environments = data.environments;
        environments = data.environments;
        restored.push('environments');
      }
      if (data.settings && typeof data.settings === 'object') {
        toRestore.settings = data.settings;
        settings = { ...settings, ...data.settings };
        restored.push('settings');
      }
      if (Array.isArray(data.pluginTraceIgnored)) {
        toRestore.pluginTraceIgnored = data.pluginTraceIgnored;
        restored.push('ignored plugin types');
      }
      if (data.appCache && typeof data.appCache === 'object') {
        toRestore.appCache = data.appCache;
        restored.push('app cache');
      }

      if (Object.keys(toRestore).length === 0) {
        throw new Error('No recognised keys found in backup file.');
      }

      await chrome.storage.local.set(toRestore);
      renderTable();
      renderSettingsForm();
      showToast(`Imported: ${restored.join(', ')}.`, 'success');
    } catch (err) {
      showToast('Import failed: ' + err.message, 'error');
    }
    // Reset so the same file can be re-imported.
    e.target.value = '';
  });

  // Reset to defaults.
  document.getElementById('btn-reset').addEventListener('click', async () => {
    if (!confirm('Reset all environments to the built-in defaults? This cannot be undone.')) return;
    const res  = await fetch(chrome.runtime.getURL('../environments.json'));
    const data = await res.json();
    environments = data.environments;
    await saveAndRender();
    showToast('Reset to defaults.', 'success');
  });

  // Escape key closes modal.
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeModal();
  });
}

// ─── UI helpers ───────────────────────────────────────────────────────────────

function showFormError(message) {
  const err = document.getElementById('form-error');
  err.textContent = message;
  err.classList.remove('hidden');
}

let toastTimer = null;

function showToast(message, type = 'success') {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.className   = `toast toast--${type}`;
  toast.classList.remove('hidden');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.add('hidden'), 3000);
}

function deriveEnvId(url) {
  try {
    const parts = new URL(url).hostname.split('.');
    // e.g. myenv.crm4.dynamics.com → myenv_crm4
    if (parts.length >= 2) return `${parts[0]}_${parts[1]}`;
    return parts[0];
  } catch { return ''; }
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ─── Model Driven Apps Panel ──────────────────────────────────────────────────

/** Fetch apps from a single environment via the service worker (uses host_permissions + credentials). */
async function _fetchMdaFromEnv(env) {
  try {
    const result = await chrome.runtime.sendMessage({
      type: 'FETCH_APPMODULES',
      envUrl: env.url,
      apiVersion: settings.apiVersion || 'v9.2',
    });
    if (result?.error || !result?.apps) return null;
    return result.apps;
  } catch (e) { console.error('[EF PPT]', 'FETCH_APPMODULES failed:', e); return null; }
}

/** Render the app list table from an array of apps. */
function _renderMdaList(apps) {
  const listEl   = document.getElementById('mda-list');
  const saveRow  = document.getElementById('mda-save-row');
  const statusEl = document.getElementById('mda-status');
  listEl.innerHTML = '';

  if (!apps || apps.length === 0) {
    statusEl.textContent = 'No model-driven apps found. Make sure you are signed in to the selected environment.';
    statusEl.className   = 'state-msg state-warn';
    statusEl.classList.remove('hidden');
    listEl.classList.add('hidden');
    saveRow.classList.add('hidden');
    return;
  }

  const sorted       = [...apps].sort((a, b) => (a.name || a.uniquename).localeCompare(b.name || b.uniquename));
  const defaultApp   = settings.defaultAppUniqueName ?? '';
  const includedApps = Array.isArray(settings.includedApps) ? settings.includedApps : [];

  const table = document.createElement('table');
  table.className = 'mda-table';
  table.innerHTML = `
    <thead>
      <tr>
        <th>App Name</th>
        <th>Unique Name</th>
        <th class="col-center" title="This app is always shown in Go To → App">Default</th>
        <th class="col-center" title="Show in Go To → App dropdown">Show in Go To</th>
      </tr>
    </thead>
    <tbody id="mda-tbody"></tbody>
  `;

  const tbody = table.querySelector('#mda-tbody');
  for (const { uniquename, name } of sorted) {
    const isDefault  = uniquename === defaultApp;
    const isIncluded = includedApps.includes(uniquename);

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="font-weight:600">${escHtml(name || uniquename)}</td>
      <td class="mda-uniquename">${escHtml(uniquename)}</td>
      <td class="col-center">
        <input type="radio" name="mda-default" value="${escHtml(uniquename)}" ${isDefault ? 'checked' : ''} />
      </td>
      <td class="col-center">
        <input type="checkbox" class="mda-include-cb" value="${escHtml(uniquename)}" ${isIncluded || isDefault ? 'checked' : ''} />
      </td>
    `;
    tbody.appendChild(tr);

    // Auto-check "Show in Go To" when this app is selected as default
    const radio = tr.querySelector('input[type="radio"]');
    const cb    = tr.querySelector('.mda-include-cb');
    radio.addEventListener('change', () => { if (radio.checked) cb.checked = true; });
  }

  listEl.appendChild(table);
  listEl.classList.remove('hidden');
  saveRow.classList.remove('hidden');
  statusEl.classList.add('hidden');
}

/** Wire up the MDA panel once when the tab is first opened. */
function initMdaPanel() {
  const statusEl  = document.getElementById('mda-status');
  const loadBtn   = document.getElementById('btn-load-mda');
  const saveBtn   = document.getElementById('btn-save-mda');
  const envSelect = document.getElementById('mda-env-select');

  // Populate environment dropdown
  envSelect.innerHTML = '';
  if (environments.length === 0) {
    envSelect.innerHTML = '<option value="">No environments configured</option>';
    loadBtn.disabled = true;
  } else {
    for (const env of environments) {
      const opt = document.createElement('option');
      opt.value       = env.url;
      opt.textContent = env.name;
      envSelect.appendChild(opt);
    }
  }

  // If we already have a cached result, render it and pre-select its environment.
  chrome.storage.local.get('appCache').then(({ appCache }) => {
    if (appCache && appCache.envUrl && Array.isArray(appCache.apps) && appCache.apps.length > 0) {
      envSelect.value = appCache.envUrl;
      _renderMdaList(appCache.apps);
    }
  });

  // Load button — fetches from the selected environment only.
  loadBtn.addEventListener('click', async () => {
    const selectedUrl = envSelect.value;
    const env = environments.find(e => e.url === selectedUrl);
    if (!env) return;

    loadBtn.disabled    = true;
    loadBtn.textContent = 'Loading…';
    statusEl.textContent = `Loading apps from ${env.name}…`;
    statusEl.className  = 'state-msg state-loading';
    statusEl.classList.remove('hidden');
    document.getElementById('mda-list').classList.add('hidden');
    document.getElementById('mda-save-row').classList.add('hidden');

    const apps = await _fetchMdaFromEnv(env);

    loadBtn.disabled    = false;
    loadBtn.textContent = 'Load Apps';

    if (!apps || apps.length === 0) {
      showToast(`Could not load apps from ${env.name}. Make sure you are signed in.`, 'error');
      statusEl.textContent = `Could not load apps from ${env.name}. Make sure you are signed in.`;
      statusEl.className   = 'state-msg state-warn';
      return;
    }

    await chrome.storage.local.set({ appCache: { envUrl: selectedUrl, apps } });
    statusEl.classList.add('hidden');
    _renderMdaList(apps);
    showToast(`Apps loaded from ${env.name}.`, 'success');
  });

  // Save button
  saveBtn.addEventListener('click', async () => {
    const defaultRadio = document.querySelector('input[name="mda-default"]:checked');
    const checkedBoxes = [...document.querySelectorAll('.mda-include-cb:checked')];
    settings.defaultAppUniqueName = defaultRadio?.value ?? '';
    settings.includedApps         = checkedBoxes.map(cb => cb.value);
    await saveSettings();
    showToast('App settings saved.', 'success');
  });
}

// ─── Plugin Trace Panel ───────────────────────────────────────────────────────

async function initPluginTracePanel() {
  const stored = await chrome.storage.local.get('pluginTraceIgnored');
  let ignoredTypes = Array.isArray(stored.pluginTraceIgnored) ? [...stored.pluginTraceIgnored] : [];

  async function _saveIgnored() {
    await chrome.storage.local.set({ pluginTraceIgnored: ignoredTypes });
  }

  function renderIgnoredList() {
    const container = document.getElementById('ignored-types-list');
    if (ignoredTypes.length === 0) {
      container.innerHTML = `
        <p style="font-size:13px;color:var(--text-muted);padding:10px 0 4px">
          No ignored plugin types. Right-click any log row in the Plugin Trace Logs viewer to ignore a type.
        </p>`;
      return;
    }

    container.innerHTML = '';
    const wrap = document.createElement('div');
    wrap.className = 'table-wrap';
    wrap.style.maxWidth = '580px';

    const table = document.createElement('table');
    table.innerHTML = `
      <thead>
        <tr>
          <th>Plugin Type</th>
          <th style="width:90px">Actions</th>
        </tr>
      </thead>`;
    const tbody = document.createElement('tbody');

    ignoredTypes.forEach((type, i) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><span class="cell-url">${escHtml(type)}</span></td>
        <td>
          <div class="actions-cell">
            <button class="btn btn--danger btn--icon-sm" data-remove="${i}">Remove</button>
          </div>
        </td>`;
      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    wrap.appendChild(table);
    container.appendChild(wrap);

    tbody.querySelectorAll('[data-remove]').forEach(btn => {
      btn.addEventListener('click', async () => {
        const idx = parseInt(btn.dataset.remove, 10);
        const removed = ignoredTypes.splice(idx, 1)[0];
        await _saveIgnored();
        renderIgnoredList();
        showToast(`"${removed}" removed from ignore list.`, 'success');
      });
    });
  }

  renderIgnoredList();

  const addInput = document.getElementById('new-ignored-type');
  const addBtn   = document.getElementById('btn-add-ignored-type');

  async function addType() {
    const val = addInput.value.trim();
    if (!val) return;
    if (ignoredTypes.includes(val)) {
      showToast(`"${val}" is already in the ignore list.`, 'error');
      return;
    }
    ignoredTypes.push(val);
    await _saveIgnored();
    renderIgnoredList();
    showToast(`"${val}" added to ignore list.`, 'success');
    addInput.value = '';
  }

  addBtn.addEventListener('click', addType);
  addInput.addEventListener('keydown', e => { if (e.key === 'Enter') addType(); });
}
