/**
 * EF Power Platform Tools — Metadata Browser (web-app)
 *
 * Runs as an iframe loaded by launcher.js from GitHub Pages.
 * All D365 Web API calls are relayed through the parent D365 tab via postMessage.
 */

'use strict';

// ─── State ────────────────────────────────────────────────────────────────────

let envUrl          = '';
let envName         = '';
let allEntities     = [];
let allOptionSets   = [];
// Multi-tab state — each open table is a tab with its own fetched data.
let openTabs     = [];   // [{ entity, allAttributes, allRels, loading, error }]
let activeTabIdx = -1;

/** Returns the currently active tab object, or null if no tabs are open. */
function currentTab() { return activeTabIdx >= 0 ? openTabs[activeTabIdx] : null; }
let currentOptionSet= null;
let sidebarMode     = 'entities'; // 'entities' | 'optionsets'
let colSortKey      = 'display';
let colSortDir      = 1; // 1 = asc, -1 = desc

// ─── Attribute type registry ──────────────────────────────────────────────────

const ATTR_TYPES = {
  StringAttributeMetadata:              { label: 'Text',             color: 'var(--type-text)',   bg: 'var(--type-text-bg)' },
  MemoAttributeMetadata:                { label: 'Multiline Text',   color: 'var(--type-text)',   bg: 'var(--type-text-bg)' },
  IntegerAttributeMetadata:             { label: 'Integer',          color: 'var(--type-num)',    bg: 'var(--type-num-bg)' },
  BigIntAttributeMetadata:              { label: 'Big Integer',      color: 'var(--type-num)',    bg: 'var(--type-num-bg)' },
  DecimalAttributeMetadata:             { label: 'Decimal',          color: 'var(--type-num)',    bg: 'var(--type-num-bg)' },
  DoubleAttributeMetadata:              { label: 'Float',            color: 'var(--type-num)',    bg: 'var(--type-num-bg)' },
  MoneyAttributeMetadata:               { label: 'Currency',         color: 'var(--type-money)',  bg: 'var(--type-money-bg)' },
  BooleanAttributeMetadata:             { label: 'Yes / No',         color: 'var(--type-bool)',   bg: 'var(--type-bool-bg)' },
  DateTimeAttributeMetadata:            { label: 'Date & Time',      color: 'var(--type-dt)',     bg: 'var(--type-dt-bg)' },
  LookupAttributeMetadata:              { label: 'Lookup',           color: 'var(--type-lookup)', bg: 'var(--type-lookup-bg)' },
  OwnerAttributeMetadata:               { label: 'Owner',            color: 'var(--type-lookup)', bg: 'var(--type-lookup-bg)' },
  CustomerAttributeMetadata:            { label: 'Customer',         color: 'var(--type-lookup)', bg: 'var(--type-lookup-bg)' },
  PicklistAttributeMetadata:            { label: 'Choice',           color: 'var(--type-choice)', bg: 'var(--type-choice-bg)' },
  MultiSelectPicklistAttributeMetadata: { label: 'Multi-Choice',     color: 'var(--type-choice)', bg: 'var(--type-choice-bg)' },
  StateAttributeMetadata:               { label: 'State',            color: 'var(--type-choice)', bg: 'var(--type-choice-bg)' },
  StatusAttributeMetadata:              { label: 'Status',           color: 'var(--type-choice)', bg: 'var(--type-choice-bg)' },
  UniqueIdentifierAttributeMetadata:    { label: 'Unique ID',        color: 'var(--type-other)',  bg: 'var(--type-other-bg)' },
  ImageAttributeMetadata:               { label: 'Image',            color: 'var(--type-other)',  bg: 'var(--type-other-bg)' },
  FileAttributeMetadata:                { label: 'File',             color: 'var(--type-other)',  bg: 'var(--type-other-bg)' },
  EntityNameAttributeMetadata:          { label: 'Entity Name',      color: 'var(--type-other)',  bg: 'var(--type-other-bg)' },
  VirtualAttributeMetadata:             { label: 'Virtual',          color: 'var(--type-other)',  bg: 'var(--type-other-bg)' },
};

// ─── Bootstrap ────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {
  const params = new URLSearchParams(location.search);
  envUrl  = params.get('env')  ?? '';
  envName = params.get('name') ?? 'Unknown';
  const autoEtn = params.get('etn') ?? '';

  document.title = `Metadata — ${envName}`;
  document.getElementById('env-name').textContent = envName;
  document.getElementById('env-url').textContent  = envUrl.replace('https://', '');

  if (!envUrl) {
    showGlobalError('No environment URL provided.\nOpen this page from the EF Power Platform Tools popup.');
    return;
  }

  wireEventListeners();

  // The D365 Metadata API does NOT support $orderby — sort client-side.
  setLoadingText('Loading metadata…');
  try {
    [allEntities, allOptionSets] = await Promise.all([
      fetchAllPages(
        `/EntityDefinitions?$select=` +
        `LogicalName,SchemaName,DisplayName,DisplayCollectionName,` +
        `OwnershipType,IsCustomEntity,IsManaged,IsActivity,IsIntersect,` +
        `PrimaryIdAttribute,PrimaryNameAttribute,EntitySetName,ObjectTypeCode`
      ).then(rows => rows.sort((a, b) =>
        getLabel(a.DisplayName).localeCompare(getLabel(b.DisplayName))
      )),
      fetchAllPages(`/GlobalOptionSetDefinitions`)
        .then(rows => rows.sort((a, b) =>
          getLabel(a.DisplayName).localeCompare(getLabel(b.DisplayName))
        ))
    ]);

    document.getElementById('sidebar-count-entities').textContent  = allEntities.length;
    document.getElementById('sidebar-count-optionsets').textContent = allOptionSets.length;

    renderEntityList();
    renderOptionSetList();

    document.getElementById('global-loading').classList.add('hidden');
    document.getElementById('app-body').style.display = 'flex';

    // Auto-select entity when opened via "Open in Metadata Browser" from Record Details.
    if (autoEtn) {
      const match = allEntities.find(e => e.LogicalName === autoEtn);
      if (match) {
        const li = document.querySelector(`#entity-list [data-ln="${CSS.escape(autoEtn)}"]`);
        if (li) li.scrollIntoView({ block: 'center' });
        openEntityTab(match);
      }
    }
  } catch (e) {
    console.error('[EF PPT]', e);
    const envLink = document.getElementById('env-link');
    if (envLink) envLink.href = envUrl;
    showGlobalError(`Failed to load metadata:\n${e.message}`);
  }
});

// ─── Event wiring ─────────────────────────────────────────────────────────────

function wireEventListeners() {
  // Sidebar type tabs.
  document.querySelectorAll('.sidebar-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      sidebarMode = tab.dataset.list;
      document.querySelectorAll('.sidebar-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');

      const isEntities = sidebarMode === 'entities';
      document.getElementById('entity-list').classList.toggle('hidden', !isEntities);
      document.getElementById('optionset-list').classList.toggle('hidden', isEntities);
      document.getElementById('sidebar-filters-entities').classList.toggle('hidden', !isEntities);
      document.getElementById('sidebar-filters-optionsets').classList.toggle('hidden', isEntities);
      document.getElementById('search-input').placeholder = isEntities ? 'Search tables…' : 'Search option sets…';

      // Update count display.
      updateCountDisplay();

      // Show/hide tab bar and detail panel based on sidebar mode.
      const tabsBar = document.getElementById('entity-tabs-bar');
      if (isEntities) {
        document.getElementById('optionset-detail').classList.add('hidden');
        if (openTabs.length > 0) {
          renderTabBar();
          switchToTab(activeTabIdx);
        } else {
          tabsBar.classList.add('hidden');
          document.getElementById('empty-state').classList.remove('hidden');
          document.getElementById('entity-detail').classList.add('hidden');
        }
      } else {
        tabsBar.classList.add('hidden');
        document.getElementById('empty-state').classList.remove('hidden');
        document.getElementById('entity-detail').classList.add('hidden');
        document.getElementById('optionset-detail').classList.add('hidden');
      }
    });
  });

  // Shared search.
  document.getElementById('search-input').addEventListener('input', () => {
    if (sidebarMode === 'entities') renderEntityList();
    else renderOptionSetList();
  });

  // Entity filters.
  document.getElementById('filter-custom').addEventListener('change',     renderEntityList);
  document.getElementById('filter-managed').addEventListener('change',    renderEntityList);
  document.getElementById('filter-intersect').addEventListener('change',  renderEntityList);
  document.getElementById('filter-adx').addEventListener('change',        renderEntityList);

  // Option set filters.
  document.getElementById('filter-os-custom').addEventListener('change',  renderOptionSetList);
  document.getElementById('filter-os-managed').addEventListener('change', renderOptionSetList);

  // Option search inside option set detail.
  document.getElementById('os-option-search').addEventListener('input', renderCurrentOptionSetOptions);

  // Entity detail tabs.
  document.querySelectorAll('.detail-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.detail-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      const panel = tab.dataset.panel;
      document.getElementById('panel-columns').classList.toggle('hidden',       panel !== 'columns');
      document.getElementById('panel-relationships').classList.toggle('hidden', panel !== 'relationships');
    });
  });

  // Relationship sub-tabs.
  document.querySelectorAll('.rel-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.rel-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      document.querySelectorAll('.rel-panel').forEach(p => p.classList.add('hidden'));
      document.getElementById(`panel-${tab.dataset.rel}`).classList.remove('hidden');
      // Refresh count display for the newly active panel.
      const _r = currentTab()?.allRels;
      if (_r && (_r.oneToMany.length || _r.manyToOne.length || _r.manyToMany.length)) {
        renderRelationships();
      }
    });
  });

  // Column panel filters.
  document.getElementById('col-search').addEventListener('input',        renderColumnsTable);
  document.getElementById('col-hide-system').addEventListener('change',  renderColumnsTable);
  document.getElementById('col-hide-managed').addEventListener('change', renderColumnsTable);

  // Relationship search.
  document.getElementById('rel-search').addEventListener('input', renderRelationships);

  // Options expand row toggle (delegated).
  document.addEventListener('click', e => {
    const link = e.target.closest('.view-options-link');
    if (!link) return;
    e.preventDefault();
    toggleOptionsExpandRow(link);
  });

  // Column header sort.
  document.getElementById('col-table').querySelectorAll('th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
      const key = th.dataset.sort;
      if (colSortKey === key) colSortDir *= -1;
      else { colSortKey = key; colSortDir = 1; }
      renderColumnsTable();
    });
  });

  // Schema Name info-button tooltip.
  // Uses a fixed-position overlay div to avoid being clipped by the scrollable table wrapper.
  const _schemaTip = document.createElement('div');
  _schemaTip.className = 'col-info-tip';
  _schemaTip.textContent =
    'For lookup attributes: if the OData navigation property name ' +
    '(ReferencingEntityNavigationPropertyName) differs from the Schema Name, ' +
    'it is shown in a second row in the same cell.';
  document.body.appendChild(_schemaTip);

  document.addEventListener('mouseover', e => {
    if (!e.target.closest('.col-th-info')) return;
    const rect = e.target.closest('.col-th-info').getBoundingClientRect();
    _schemaTip.style.left = `${rect.left + rect.width / 2}px`;
    _schemaTip.style.top  = `${rect.bottom + 6}px`;
    _schemaTip.classList.add('col-info-tip--visible');
  });
  document.addEventListener('mouseout', e => {
    if (e.target.closest('.col-th-info')) _schemaTip.classList.remove('col-info-tip--visible');
  });
  // Prevent the info button click from bubbling up to the <th> sort handler.
  document.addEventListener('click', e => {
    if (e.target.closest('.col-th-info')) e.stopPropagation();
  }, true /* capture — fires before the sort listener */);
}

function updateCountDisplay() {
  if (sidebarMode === 'entities') {
    const filtered = getFilteredEntities();
    document.getElementById('entity-count').textContent = `${filtered.length} / ${allEntities.length}`;
  } else {
    const filtered = getFilteredOptionSets();
    document.getElementById('entity-count').textContent = `${filtered.length} / ${allOptionSets.length}`;
  }
}

// ─── Entity list ─────────────────────────────────────────────────────────────

function getFilteredEntities() {
  const query          = document.getElementById('search-input').value.trim().toLowerCase();
  const customOnly     = document.getElementById('filter-custom').checked;
  const hideManaged    = document.getElementById('filter-managed').checked;
  const hideIntersect  = document.getElementById('filter-intersect').checked;
  const hideAdx        = document.getElementById('filter-adx').checked;

  return allEntities.filter(e => {
    if (customOnly    && !e.IsCustomEntity)                              return false;
    if (hideManaged   &&  e.IsManaged)                                   return false;
    if (hideIntersect &&  e.IsIntersect)                                 return false;
    if (hideAdx       && (e.LogicalName ?? '').startsWith('adx_'))       return false;
    if (!query) return true;
    return getLabel(e.DisplayName).toLowerCase().includes(query)
        || e.LogicalName.toLowerCase().includes(query);
  });
}

function renderEntityList() {
  const filtered = getFilteredEntities();
  document.getElementById('entity-count').textContent = `${filtered.length} / ${allEntities.length}`;

  const list = document.getElementById('entity-list');
  list.innerHTML = '';

  filtered.forEach(entity => {
    const displayName = getLabel(entity.DisplayName) || entity.LogicalName;
    const li = document.createElement('li');
    li.className  = 'entity-item';
    li.role       = 'option';
    li.dataset.ln = entity.LogicalName;
    li.innerHTML  = `
      <div class="entity-item__display">${esc(displayName)}</div>
      <div class="entity-item__logical">${esc(entity.LogicalName)}</div>
      <div class="entity-item__badges">
        ${entity.IsCustomEntity ? '<span class="badge badge--custom">Custom</span>'    : ''}
        ${entity.IsActivity     ? '<span class="badge badge--activity">Activity</span>': ''}
        ${entity.IsManaged      ? '<span class="badge badge--managed">Managed</span>'  : ''}
      </div>`;
    li.addEventListener('click', () => openEntityTab(entity));
    list.appendChild(li);
  });
}

// ─── Tab management ───────────────────────────────────────────────────────────

/** Renders the horizontal tab strip at the top of the detail panel. */
function renderTabBar() {
  const bar = document.getElementById('entity-tabs-bar');
  if (openTabs.length === 0) { bar.classList.add('hidden'); return; }

  bar.classList.remove('hidden');
  bar.innerHTML = '';

  openTabs.forEach((tab, i) => {
    const displayName = getLabel(tab.entity.DisplayName) || tab.entity.LogicalName;
    const el = document.createElement('div');
    el.className = `entity-tab${i === activeTabIdx ? ' active' : ''}`;
    el.setAttribute('role', 'tab');
    el.setAttribute('aria-selected', String(i === activeTabIdx));
    el.title = `${displayName}\n${tab.entity.LogicalName}`;

    const nameEl = document.createElement('span');
    nameEl.className = 'entity-tab__name';
    nameEl.textContent = displayName;

    const closeBtn = document.createElement('button');
    closeBtn.className = 'entity-tab__close';
    closeBtn.title = 'Close tab';
    closeBtn.innerHTML = '&times;';
    closeBtn.addEventListener('click', e => {
      e.stopPropagation();
      closeEntityTab(i);
    });

    el.appendChild(nameEl);
    el.appendChild(closeBtn);
    el.addEventListener('click', () => { if (i !== activeTabIdx) switchToTab(i); });
    bar.appendChild(el);
  });

  // Scroll active tab into view.
  bar.querySelector('.entity-tab.active')?.scrollIntoView({ inline: 'nearest', block: 'nearest' });
}

/** Updates sidebar item highlighting: active tab = primary, background tabs = subtle. */
function updateSidebarHighlights() {
  const openNames   = new Set(openTabs.map(t => t.entity.LogicalName));
  const activeName  = currentTab()?.entity.LogicalName;
  document.querySelectorAll('#entity-list .entity-item').forEach(li => {
    const ln = li.dataset.ln;
    li.classList.toggle('entity-item--active', ln === activeName);
    li.classList.toggle('entity-item--open',   openNames.has(ln) && ln !== activeName);
  });
}

/** Renders the active tab's entity data into the detail panel. */
function renderActiveTab() {
  const tab = currentTab();
  if (!tab) return;

  // Reset sort + filter state when switching context.
  colSortKey = 'display';
  colSortDir = 1;
  document.getElementById('col-search').value = '';
  document.getElementById('rel-search').value = '';

  renderEntityHeader(tab.entity);
  renderColumnsTable();
  renderRelationships();

  // Reset to Columns sub-tab.
  document.querySelectorAll('.detail-tab').forEach(t =>
    t.classList.toggle('active', t.dataset.panel === 'columns')
  );
  document.getElementById('panel-columns').classList.remove('hidden');
  document.getElementById('panel-relationships').classList.add('hidden');

  // Reset relationship sub-tab to 1:N.
  document.querySelectorAll('.rel-tab').forEach(t => t.classList.toggle('active', t.dataset.rel === '1n'));
  document.querySelectorAll('.rel-panel').forEach(p => p.classList.add('hidden'));
  document.getElementById('panel-1n').classList.remove('hidden');

  document.getElementById('empty-state').classList.add('hidden');
  document.getElementById('detail-loading').classList.add('hidden');
  document.getElementById('entity-detail').classList.remove('hidden');
  document.getElementById('optionset-detail').classList.add('hidden');
}

/** Switches focus to the tab at index `idx`, re-rendering from cached data. */
function switchToTab(idx) {
  if (idx < 0 || idx >= openTabs.length) return;
  activeTabIdx = idx;
  renderTabBar();
  updateSidebarHighlights();

  const tab = openTabs[idx];
  if (tab.loading) {
    showDetailLoading();
  } else if (tab.error) {
    hideDetailLoading();
    showEmptyState('Failed to load table metadata', tab.error);
  } else {
    renderActiveTab();
  }
}

/** Closes the tab at index `idx`, switching to the nearest remaining tab. */
function closeEntityTab(idx) {
  openTabs.splice(idx, 1);

  if (openTabs.length === 0) {
    activeTabIdx = -1;
    document.getElementById('entity-tabs-bar').classList.add('hidden');
    document.getElementById('empty-state').classList.remove('hidden');
    document.getElementById('entity-detail').classList.add('hidden');
    document.getElementById('detail-loading').classList.add('hidden');
    updateSidebarHighlights();
    return;
  }

  // Stay on the tab to the left of the closed one, clamped to a valid index.
  activeTabIdx = Math.min(idx, openTabs.length - 1);
  switchToTab(activeTabIdx);
}

// ─── Entity selection ─────────────────────────────────────────────────────────

async function openEntityTab(entity) {
  // If already open, just switch to that tab.
  const existingIdx = openTabs.findIndex(t => t.entity.LogicalName === entity.LogicalName);
  if (existingIdx !== -1) { switchToTab(existingIdx); return; }

  // Create a placeholder tab marked as loading.
  const tab = { entity, allAttributes: [], allRels: null, loading: true, error: null };
  openTabs.push(tab);
  activeTabIdx = openTabs.length - 1;

  renderTabBar();
  updateSidebarHighlights();
  showDetailLoading();

  try {
    const base = `/EntityDefinitions(LogicalName='${entity.LogicalName}')`;

    // Fetch attributes, relationships, and choice-type option sets all in parallel.
    //
    // The base /Attributes endpoint does NOT return OptionSet.Options inline —
    // D365 requires casting to the specific type and using $expand=OptionSet to
    // get the options array. We fetch all four choice types explicitly so that
    // both local and global option sets are always resolved.
    const [
      attrs, relOtM, relMtO, relMtM,
      pickOpts, multiPickOpts, stateOpts, statusOpts
    ] = await Promise.all([
      fetchAllPages(`${base}/Attributes`),
      fetchAllPages(`${base}/OneToManyRelationships`),
      fetchAllPages(`${base}/ManyToOneRelationships`),
      fetchAllPages(`${base}/ManyToManyRelationships`),
      fetchAllPages(`${base}/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet`)
        .catch(() => []),
      fetchAllPages(`${base}/Attributes/Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet`)
        .catch(() => []),
      fetchAllPages(`${base}/Attributes/Microsoft.Dynamics.CRM.StateAttributeMetadata?$select=LogicalName&$expand=OptionSet`)
        .catch(() => []),
      fetchAllPages(`${base}/Attributes/Microsoft.Dynamics.CRM.StatusAttributeMetadata?$select=LogicalName&$expand=OptionSet`)
        .catch(() => []),
    ]);

    // Build a LogicalName → { OptionSet, GlobalOptionSet } map from the
    // explicitly-expanded choice attribute fetches.
    const optionMap = new Map();
    [...pickOpts, ...multiPickOpts, ...stateOpts, ...statusOpts].forEach(a => {
      if (a.LogicalName) {
        optionMap.set(a.LogicalName, {
          OptionSet:       a.OptionSet       ?? null,
          GlobalOptionSet: a.GlobalOptionSet ?? null,
        });
      }
    });

    // Build a lookup-attribute LogicalName → ReferencingEntityNavigationPropertyName map.
    // The ManyToOne relationships response already carries this property on each relationship
    // object (via rel.ReferencingAttribute = the lookup field's logical name), so we harvest
    // it from there rather than making an extra API call.
    const lookupNavMap = new Map();
    relMtO.forEach(rel => {
      if (rel.ReferencingAttribute && rel.ReferencingEntityNavigationPropertyName) {
        lookupNavMap.set(rel.ReferencingAttribute, rel.ReferencingEntityNavigationPropertyName);
      }
    });

    // Merge option data + lookup nav-prop name into the main attribute list, then sort.
    tab.allAttributes = attrs
      .map(a => {
        const extra   = optionMap.get(a.LogicalName);
        const navProp = lookupNavMap.get(a.LogicalName);
        return {
          ...a,
          ...(extra   ? extra   : {}),
          ...(navProp ? { ReferencingEntityNavigationPropertyName: navProp } : {}),
        };
      })
      .sort((a, b) => getLabel(a.DisplayName).localeCompare(getLabel(b.DisplayName)));

    tab.allRels  = { oneToMany: relOtM, manyToOne: relMtO, manyToMany: relMtM };
    tab.loading  = false;

    // Only render if this tab is still the active one (user may have switched away).
    if (currentTab() === tab) renderActiveTab();

  } catch (e) {
    console.error('[EF PPT]', e);
    tab.loading = false;
    tab.error   = e.message;
    if (currentTab() === tab) {
      hideDetailLoading();
      showEmptyState('Failed to load table metadata', e.message);
    }
  }
}

// ─── Entity header ────────────────────────────────────────────────────────────

function renderEntityHeader(entity) {
  const displayName    = getLabel(entity.DisplayName)           || entity.LogicalName;
  const collectionName = getLabel(entity.DisplayCollectionName) || '';

  document.getElementById('entity-header').innerHTML = `
    <div class="entity-header__title">
      ${esc(displayName)}
      ${entity.IsCustomEntity ? '<span class="badge badge--custom">Custom</span>'    : ''}
      ${entity.IsActivity     ? '<span class="badge badge--activity">Activity</span>': ''}
      ${entity.IsManaged      ? '<span class="badge badge--managed">Managed</span>'  : ''}
    </div>
    <div class="entity-header__meta">
      ${metaItem('Logical Name',    entity.LogicalName,          true)}
      ${metaItem('Schema Name',     entity.SchemaName,           true)}
      ${metaItem('Collection Name', collectionName)}
      ${metaItem('Entity Set',      entity.EntitySetName,        true)}
      ${metaItem('Ownership',       entity.OwnershipType)}
      ${metaItem('Primary Key',     entity.PrimaryIdAttribute,   true)}
      ${metaItem('Primary Name',    entity.PrimaryNameAttribute, true)}
      ${metaItem('Type Code',       entity.ObjectTypeCode)}
    </div>`;
}

function metaItem(label, value, copyable = false) {
  if (!value && value !== 0) return '';
  const v = esc(String(value));
  const valHtml = copyable
    ? `${v} <button class="copy-btn" data-copy="${v}" title="Copy">&#128203;</button>`
    : v;
  return `
    <div class="entity-header__meta-item">
      <span class="entity-header__meta-label">${esc(label)}:</span>
      <span>${valHtml}</span>
    </div>`;
}

// Delegate copy button clicks (header copy buttons — these are inline in text).
document.addEventListener('click', e => {
  const btn = e.target.closest('.copy-btn[data-copy]');
  if (!btn) return;
  copyToClipboard(btn.dataset.copy, btn);
});

// ─── Columns table ────────────────────────────────────────────────────────────

function renderColumnsTable() {
  const query      = document.getElementById('col-search').value.trim().toLowerCase();
  const hideSystem = document.getElementById('col-hide-system').checked;
  const hideManaged= document.getElementById('col-hide-managed').checked;

  const _attrs = currentTab()?.allAttributes ?? [];
  let rows = _attrs.filter(attr => {
    if (hideManaged && attr.IsManaged)    return false;
    if (hideSystem  && isSystemField(attr)) return false;
    if (!query) return true;
    return getLabel(attr.DisplayName).toLowerCase().includes(query)
        || (attr.LogicalName ?? '').toLowerCase().includes(query)
        || (attr.SchemaName  ?? '').toLowerCase().includes(query);
  });

  rows = sortAttributes(rows);

  document.getElementById('col-count').textContent = `${rows.length} columns`;
  document.getElementById('tab-count-columns').textContent = _attrs.length;

  const tbody = document.getElementById('col-tbody');
  tbody.innerHTML = '';

  rows.forEach(attr => {
    const typeMeta    = getAttrTypeMeta(attr);
    const reqInfo     = getRequiredInfo(attr);
    const displayName = getLabel(attr.DisplayName) || attr.LogicalName;
    const details     = buildDetailsHtml(attr);

    // Schema Name cell — for lookup types, show ReferencingEntityNavigationPropertyName
    // underneath SchemaName when the two values differ, each with its own copy button.
    const schemaName = attr.SchemaName ?? '';
    const odataType  = (attr['@odata.type'] ?? '').split('.').pop();
    const isLookup   = ['LookupAttributeMetadata', 'OwnerAttributeMetadata', 'CustomerAttributeMetadata'].includes(odataType);
    const navProp    = isLookup ? (attr.ReferencingEntityNavigationPropertyName ?? '') : '';
    const showNavProp = navProp && navProp !== schemaName && (attr.Targets ?? []).length <= 1;

    const schemaCellHtml = showNavProp
      ? `<td data-no-autocopy>
          <div class="cell-schema-row">
            <span class="cell-logical">${esc(schemaName)}</span>
            <button class="copy-btn" data-copy="${esc(schemaName)}" title="Copy schema name">&#128203;</button>
          </div>
          <div class="cell-schema-row cell-schema-row--nav">
            <span class="cell-logical cell-logical--muted">${esc(navProp)}</span>
            <button class="copy-btn" data-copy="${esc(navProp)}" title="Copy navigation property name">&#128203;</button>
          </div>
        </td>`
      : `<td data-copytext="${esc(schemaName)}"><div class="cell-logical">${esc(schemaName)}</div></td>`;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td data-copytext="${esc(displayName)}">
        <div class="cell-name">${esc(displayName)}</div>
      </td>
      <td data-copytext="${esc(attr.LogicalName ?? '')}">
        <div class="cell-logical">${esc(attr.LogicalName ?? '')}</div>
      </td>
      ${schemaCellHtml}
      <td data-copytext="${esc(typeMeta.label)}">
        <span class="type-badge" style="color:${typeMeta.color};background:${typeMeta.bg}">
          ${esc(typeMeta.label)}
        </span>
      </td>
      <td data-copytext="${buildCopyTextForDetails(attr)}">${details}</td>
      <td data-copytext="${esc(reqInfo.label)}">
        ${reqInfo.label
          ? `<span class="req-badge req-badge--${reqInfo.cls}">${esc(reqInfo.label)}</span>`
          : '<span style="color:var(--border)">—</span>'}
      </td>`;

    tbody.appendChild(tr);
  });

  addCopyButtons(document.getElementById('col-table'));
  makeTableResizable(document.getElementById('col-table'));
}

function sortAttributes(rows) {
  return [...rows].sort((a, b) => {
    let va, vb;
    switch (colSortKey) {
      case 'display': va = getLabel(a.DisplayName);     vb = getLabel(b.DisplayName);     break;
      case 'logical': va = a.LogicalName ?? '';         vb = b.LogicalName ?? '';         break;
      case 'schema':  va = a.SchemaName  ?? '';         vb = b.SchemaName  ?? '';         break;
      case 'type':    va = getAttrTypeMeta(a).label;   vb = getAttrTypeMeta(b).label;   break;
      case 'req':     va = a.RequiredLevel?.Value ?? ''; vb = b.RequiredLevel?.Value ?? ''; break;
      default:        return 0;
    }
    return colSortDir * va.localeCompare(vb);
  });
}

function isSystemField(attr) {
  const systemNames = new Set([
    'createdon','modifiedon','createdby','modifiedby','createdonbehalfby',
    'modifiedonbehalfby','ownerid','owningbusinessunit','owningteam','owninguser',
    'statecode','statuscode','versionnumber','timezoneruleversionnumber',
    'utcconversiontimezonecode','importsequencenumber','overriddencreatedon',
    'overwritetime','organizationid'
  ]);
  return !attr.IsCustomAttribute && systemNames.has(attr.LogicalName ?? '');
}

/**
 * Resolves the options array for a choice/picklist attribute.
 *
 * D365 returns OptionSet.Options inline for local option sets.
 * For global option sets the Options array may be absent — we cross-reference
 * the already-fetched `allOptionSets` by MetadataId and then by Name as
 * further fallbacks so options always resolve without extra API calls.
 *
 * @returns {{ options: object[], globalName: string }}
 */
function getOptionsForAttr(attr) {
  // 1. Inline local (or already-populated global) option set.
  const direct = attr.OptionSet?.Options;
  if (Array.isArray(direct) && direct.length > 0) {
    return { options: direct, globalName: '' };
  }

  // 2. Navigation property GlobalOptionSet (present when D365 expands it).
  const fromNav = attr.GlobalOptionSet?.Options;
  if (Array.isArray(fromNav) && fromNav.length > 0) {
    return { options: fromNav, globalName: attr.GlobalOptionSet?.Name ?? '' };
  }

  // 3. Cross-reference allOptionSets by MetadataId.
  const metadataId = attr.OptionSet?.MetadataId;
  if (metadataId) {
    const found = allOptionSets.find(os => os.MetadataId === metadataId);
    if (Array.isArray(found?.Options) && found.Options.length > 0) {
      return { options: found.Options, globalName: found.Name ?? '' };
    }
  }

  // 4. Cross-reference allOptionSets by Name (OptionSet.Name or GlobalOptionSetName).
  const osName = attr.OptionSet?.Name ?? attr.GlobalOptionSetName ?? null;
  if (osName) {
    const found = allOptionSets.find(os => os.Name === osName);
    if (Array.isArray(found?.Options) && found.Options.length > 0) {
      return { options: found.Options, globalName: found.Name ?? '' };
    }
  }

  return { options: [], globalName: '' };
}

function getAttrTypeMeta(attr) {
  const key = (attr['@odata.type'] ?? '').split('.').pop();
  return ATTR_TYPES[key] ?? {
    label: attr.AttributeType ?? key,
    color: 'var(--type-other)',
    bg:    'var(--type-other-bg)'
  };
}

function getRequiredInfo(attr) {
  switch (attr.RequiredLevel?.Value) {
    case 'SystemRequired':      return { label: 'System',   cls: 'system' };
    case 'ApplicationRequired': return { label: 'Required', cls: 'app'    };
    case 'Recommended':         return { label: 'Rec.',     cls: 'rec'    };
    default:                    return { label: '',         cls: ''        };
  }
}

/** Returns a plain-text string suitable for clipboard copy from a details cell. */
function buildCopyTextForDetails(attr) {
  const key = (attr['@odata.type'] ?? '').split('.').pop();

  switch (key) {
    case 'StringAttributeMetadata':
    case 'MemoAttributeMetadata': {
      const parts = [];
      if (attr.MaxLength != null) parts.push(`Max length: ${attr.MaxLength}`);
      if (attr.Format && attr.Format !== 'Text') parts.push(`Format: ${attr.Format}`);
      return parts.join(' | ');
    }
    case 'IntegerAttributeMetadata': {
      const parts = [];
      if (attr.MinValue != null) parts.push(`Min: ${attr.MinValue}`);
      if (attr.MaxValue != null) parts.push(`Max: ${attr.MaxValue}`);
      if (attr.Format && attr.Format !== 'None') parts.push(`Format: ${attr.Format}`);
      return parts.join(' | ');
    }
    case 'DecimalAttributeMetadata':
    case 'DoubleAttributeMetadata':
    case 'MoneyAttributeMetadata': {
      const parts = [];
      if (attr.MinValue  != null) parts.push(`Min: ${attr.MinValue}`);
      if (attr.MaxValue  != null) parts.push(`Max: ${attr.MaxValue}`);
      if (attr.Precision != null) parts.push(`Precision: ${attr.Precision}`);
      return parts.join(' | ');
    }
    case 'BooleanAttributeMetadata': {
      const t = attr.OptionSet?.TrueOption?.Label?.UserLocalizedLabel?.Label  ?? 'Yes';
      const f = attr.OptionSet?.FalseOption?.Label?.UserLocalizedLabel?.Label ?? 'No';
      return `True: ${t} | False: ${f}`;
    }
    case 'DateTimeAttributeMetadata': {
      const parts = [];
      if (attr.Format) parts.push(`Format: ${attr.Format}`);
      const beh = attr.DateTimeBehavior?.Value ?? attr.DateTimeBehavior;
      if (beh) parts.push(`Behavior: ${beh}`);
      return parts.join(' | ');
    }
    case 'LookupAttributeMetadata':
    case 'OwnerAttributeMetadata':
    case 'CustomerAttributeMetadata':
      return (attr.Targets ?? []).join(', ');
    case 'PicklistAttributeMetadata':
    case 'MultiSelectPicklistAttributeMetadata':
    case 'StateAttributeMetadata':
    case 'StatusAttributeMetadata': {
      const { options } = getOptionsForAttr(attr);
      return options
        .slice()
        .sort((a, b) => (a.Label?.UserLocalizedLabel?.Label ?? '').localeCompare(b.Label?.UserLocalizedLabel?.Label ?? ''))
        .map(o => `${o.Value}: ${o.Label?.UserLocalizedLabel?.Label ?? ''}`)
        .join(' | ');
    }
    default:
      return '';
  }
}

/** Builds the HTML for the Details column of the columns table. */
function buildDetailsHtml(attr) {
  const key = (attr['@odata.type'] ?? '').split('.').pop();
  const pills = [];

  switch (key) {
    case 'StringAttributeMetadata':
      if (attr.MaxLength != null) pills.push(`Max length: ${attr.MaxLength.toLocaleString()}`);
      if (attr.Format && attr.Format !== 'Text') pills.push(`Format: ${attr.Format}`);
      break;

    case 'MemoAttributeMetadata':
      if (attr.MaxLength != null) pills.push(`Max length: ${attr.MaxLength.toLocaleString()}`);
      break;

    case 'IntegerAttributeMetadata':
      if (attr.MinValue != null) pills.push(`Min: ${attr.MinValue.toLocaleString()}`);
      if (attr.MaxValue != null) pills.push(`Max: ${attr.MaxValue.toLocaleString()}`);
      if (attr.Format && attr.Format !== 'None') pills.push(`Format: ${attr.Format}`);
      break;

    case 'DecimalAttributeMetadata':
    case 'DoubleAttributeMetadata':
      if (attr.MinValue   != null) pills.push(`Min: ${attr.MinValue}`);
      if (attr.MaxValue   != null) pills.push(`Max: ${attr.MaxValue}`);
      if (attr.Precision  != null) pills.push(`Precision: ${attr.Precision}`);
      break;

    case 'MoneyAttributeMetadata':
      if (attr.MinValue        != null) pills.push(`Min: ${attr.MinValue}`);
      if (attr.MaxValue        != null) pills.push(`Max: ${attr.MaxValue}`);
      if (attr.Precision       != null) pills.push(`Precision: ${attr.Precision}`);
      if (attr.PrecisionSource != null) pills.push(`Precision source: ${attr.PrecisionSource}`);
      break;

    case 'BooleanAttributeMetadata': {
      const t = attr.OptionSet?.TrueOption?.Label?.UserLocalizedLabel?.Label  ?? 'Yes';
      const f = attr.OptionSet?.FalseOption?.Label?.UserLocalizedLabel?.Label ?? 'No';
      pills.push(`True: ${t}`, `False: ${f}`);
      break;
    }

    case 'DateTimeAttributeMetadata': {
      if (attr.Format) pills.push(`Format: ${attr.Format}`);
      const beh = attr.DateTimeBehavior?.Value ?? attr.DateTimeBehavior;
      if (beh)  pills.push(`Behavior: ${beh}`);
      break;
    }

    case 'LookupAttributeMetadata':
    case 'OwnerAttributeMetadata':
    case 'CustomerAttributeMetadata':
      return buildLookupHtml(attr.Targets ?? []);

    case 'PicklistAttributeMetadata':
    case 'MultiSelectPicklistAttributeMetadata':
    case 'StateAttributeMetadata':
    case 'StatusAttributeMetadata': {
      const { options, globalName } = getOptionsForAttr(attr);
      return buildOptionsLinkHtml(options, globalName);
    }

    case 'ImageAttributeMetadata':
    case 'FileAttributeMetadata':
      if (attr.MaxSizeInKB != null) pills.push(`Max: ${attr.MaxSizeInKB.toLocaleString()} KB`);
      if (attr.IsPrimaryImage)      pills.push('Primary image');
      break;

    default: break;
  }

  return pills.length
    ? pills.map(p => `<span class="detail-pill">${esc(p)}</span>`).join('')
    : '<span style="color:var(--border)">—</span>';
}

function buildLookupHtml(targets) {
  if (!targets.length) return '<span style="color:var(--border)">—</span>';
  return targets
    .map(t => `<span class="detail-pill detail-pill--lookup">&#128279; ${esc(t)}</span>`)
    .join('');
}

/**
 * Renders a "▶ N options" link for choice/picklist cells.
 * Clicking it inserts an expand row below the current table row.
 * The options data is stored as JSON in a data attribute on the link.
 */
function buildOptionsLinkHtml(options, globalName) {
  if (!options.length) return '<span style="color:var(--border)">—</span>';

  // Encode safely for use in an HTML attribute.
  const encoded  = escAttr(JSON.stringify(options));
  const count    = options.length;
  const gBadge   = globalName
    ? ` <span class="detail-pill" style="font-size:10px;opacity:0.8">global: ${esc(globalName)}</span>`
    : '';

  return `<a class="view-options-link" data-options="${encoded}" href="#" title="Show / hide options">&#9654; ${count} option${count !== 1 ? 's' : ''}</a>${gBadge}`;
}

/**
 * Toggles an expand row below the column's <tr> that shows all options
 * with Value, Label, and Color columns — each cell copy-able on hover.
 */
function toggleOptionsExpandRow(link) {
  const parentTr = link.closest('tr');
  const nextTr   = parentTr.nextElementSibling;
  const isOpen   = nextTr?.classList.contains('options-expand-row');

  // Always parse from data — never infer count from link text (text starts
  // with a ▶/▼ arrow so parseInt would return NaN and fall back to 0).
  let options;
  try {
    options = JSON.parse(link.dataset.options);
  } catch {
    return;
  }

  const count = options.length;

  if (isOpen) {
    nextTr.remove();
    link.innerHTML = `&#9654; ${count} option${count !== 1 ? 's' : ''}`;
    return;
  }

  // Sort options alphabetically by label.
  options = options.slice().sort((a, b) =>
    (a.Label?.UserLocalizedLabel?.Label ?? '').localeCompare(b.Label?.UserLocalizedLabel?.Label ?? '')
  );

  const colCount  = parentTr.cells.length;
  const tbody     = document.createElement('tbody');

  options.forEach(o => {
    const label      = o.Label?.UserLocalizedLabel?.Label ?? '—';
    const colorValid = o.Color && o.Color !== '#00000000';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td data-copytext="${esc(String(o.Value))}">
        <span class="os-value-pill">${esc(String(o.Value))}</span>
      </td>
      <td data-copytext="${esc(label)}">
        <strong>${esc(label)}</strong>
      </td>
      <td data-copytext="${esc(o.Color ?? '')}">
        ${colorValid
          ? `<span class="os-color-swatch" style="background:${esc(o.Color)};margin-right:5px" title="${esc(o.Color)}"></span><span style="font-size:11px;color:var(--text-muted);font-family:monospace">${esc(o.Color)}</span>`
          : '<span style="color:var(--border)">—</span>'}
      </td>`;
    tbody.appendChild(tr);
  });

  const innerTable = document.createElement('table');
  innerTable.className = 'data-table options-expand-table';
  innerTable.innerHTML = `<thead><tr><th>Value</th><th>Label</th><th>Color</th></tr></thead>`;
  innerTable.appendChild(tbody);

  addCopyButtons(innerTable);

  const expandTr = document.createElement('tr');
  expandTr.className = 'options-expand-row';

  // Empty spacer cell covering Display Name, Logical Name, Type columns.
  const spacerTd = document.createElement('td');
  spacerTd.colSpan = colCount - 2;
  spacerTd.className = 'options-expand-spacer';

  // Content cell spanning Details + Required columns.
  const expandTd = document.createElement('td');
  expandTd.colSpan = 2;
  expandTd.className = 'options-expand-cell';

  const wrap = document.createElement('div');
  wrap.className = 'options-expand-wrap';
  wrap.appendChild(innerTable);
  expandTd.appendChild(wrap);
  expandTr.appendChild(spacerTd);
  expandTr.appendChild(expandTd);

  parentTr.after(expandTr);
  link.innerHTML = `&#9660; ${options.length} option${options.length !== 1 ? 's' : ''}`;
}

// ─── Relationships ────────────────────────────────────────────────────────────

function renderRelationships() {
  const { oneToMany, manyToOne, manyToMany } = currentTab()?.allRels ?? { oneToMany: [], manyToOne: [], manyToMany: [] };
  const total = oneToMany.length + manyToOne.length + manyToMany.length;

  document.getElementById('tab-count-rels').textContent = total;
  document.getElementById('tab-count-1n').textContent   = oneToMany.length;
  document.getElementById('tab-count-n1').textContent   = manyToOne.length;
  document.getElementById('tab-count-nn').textContent   = manyToMany.length;

  const shown1n = renderRelTable('panel-1n', oneToMany,  buildOneToManyRow,  ['Schema Name', 'Related Table', 'Lookup Field', 'Referenced Field', 'Delete', 'Assign', 'Share']);
  const shownN1 = renderRelTable('panel-n1', manyToOne,  buildManyToOneRow,  ['Schema Name', 'Related Table (Parent)', 'Lookup Field', 'Referenced Field', 'Delete']);
  const shownNN = renderRelTable('panel-nn', manyToMany, buildManyToManyRow, ['Schema Name', 'Entity 1', 'Entity 2', 'Intersect Table']);

  // Update the count badge to reflect the active panel's filtered count.
  const activeRel = document.querySelector('.rel-tab.active')?.dataset.rel;
  const shownMap  = { '1n': shown1n, 'n1': shownN1, 'nn': shownNN };
  const totalMap  = { '1n': oneToMany.length, 'n1': manyToOne.length, 'nn': manyToMany.length };
  const s = shownMap[activeRel] ?? 0;
  const t = totalMap[activeRel] ?? 0;
  const q = document.getElementById('rel-search').value.trim();
  document.getElementById('rel-count').textContent = q ? `${s} / ${t}` : `${t}`;
}

/** Returns the number of visible rows after filtering. */
function renderRelTable(panelId, rels, rowBuilder, headers) {
  const panel = document.getElementById(panelId);
  const query = (document.getElementById('rel-search')?.value ?? '').trim().toLowerCase();

  const sorted = [...rels].sort((a, b) =>
    (a.SchemaName ?? '').localeCompare(b.SchemaName ?? '')
  );

  const visible = query
    ? sorted.filter(rel => JSON.stringify(rel).toLowerCase().includes(query))
    : sorted;

  if (!visible.length) {
    panel.innerHTML = rels.length
      ? `<p style="padding:20px;color:var(--text-muted);font-style:italic">No relationships match the filter.</p>`
      : `<p style="padding:20px;color:var(--text-muted);font-style:italic">No relationships of this type.</p>`;
    return 0;
  }

  const table = document.createElement('table');
  table.className = 'data-table rel-table';
  table.innerHTML = `
    <thead><tr>${headers.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead>
    <tbody>${visible.map(rel => rowBuilder(rel)).join('')}</tbody>`;

  panel.innerHTML = '';
  panel.appendChild(table);
  addCopyButtons(table);
  makeTableResizable(table);
  return visible.length;
}

function buildOneToManyRow(rel) {
  return `<tr>
    <td data-copytext="${esc(rel.SchemaName ?? '')}">
      <span class="cell-logical">${esc(rel.SchemaName ?? '—')}</span>
    </td>
    <td data-copytext="${esc(rel.ReferencingEntity ?? '')}">
      <span class="detail-pill detail-pill--lookup">&#128279; ${esc(rel.ReferencingEntity ?? '—')}</span>
    </td>
    <td data-copytext="${esc(rel.ReferencingAttribute ?? '')}">
      <span class="cell-logical">${esc(rel.ReferencingAttribute ?? '—')}</span>
    </td>
    <td data-copytext="${esc(rel.ReferencedAttribute ?? '')}">
      <span class="cell-logical">${esc(rel.ReferencedAttribute ?? '—')}</span>
    </td>
    <td data-copytext="${esc(rel.CascadeConfiguration?.Delete ?? '')}">${cascadeBadge(rel.CascadeConfiguration?.Delete)}</td>
    <td data-copytext="${esc(rel.CascadeConfiguration?.Assign ?? '')}">${cascadeBadge(rel.CascadeConfiguration?.Assign)}</td>
    <td data-copytext="${esc(rel.CascadeConfiguration?.Share  ?? '')}">${cascadeBadge(rel.CascadeConfiguration?.Share)}</td>
  </tr>`;
}

function buildManyToOneRow(rel) {
  return `<tr>
    <td data-copytext="${esc(rel.SchemaName ?? '')}">
      <span class="cell-logical">${esc(rel.SchemaName ?? '—')}</span>
    </td>
    <td data-copytext="${esc(rel.ReferencedEntity ?? '')}">
      <span class="detail-pill detail-pill--lookup">&#128279; ${esc(rel.ReferencedEntity ?? '—')}</span>
    </td>
    <td data-copytext="${esc(rel.ReferencingAttribute ?? '')}">
      <span class="cell-logical">${esc(rel.ReferencingAttribute ?? '—')}</span>
    </td>
    <td data-copytext="${esc(rel.ReferencedAttribute ?? '')}">
      <span class="cell-logical">${esc(rel.ReferencedAttribute ?? '—')}</span>
    </td>
    <td data-copytext="${esc(rel.CascadeConfiguration?.Delete ?? '')}">${cascadeBadge(rel.CascadeConfiguration?.Delete)}</td>
  </tr>`;
}

function buildManyToManyRow(rel) {
  return `<tr>
    <td data-copytext="${esc(rel.SchemaName ?? '')}">
      <span class="cell-logical">${esc(rel.SchemaName ?? '—')}</span>
    </td>
    <td data-copytext="${esc(rel.Entity1LogicalName ?? '')}">
      <span class="detail-pill detail-pill--lookup">&#128279; ${esc(rel.Entity1LogicalName ?? '—')}</span>
    </td>
    <td data-copytext="${esc(rel.Entity2LogicalName ?? '')}">
      <span class="detail-pill detail-pill--lookup">&#128279; ${esc(rel.Entity2LogicalName ?? '—')}</span>
    </td>
    <td data-copytext="${esc(rel.IntersectEntityName ?? '')}">
      <span class="cell-logical">${esc(rel.IntersectEntityName ?? '—')}</span>
    </td>
  </tr>`;
}

function cascadeBadge(value) {
  if (!value) return '<span style="color:var(--border)">—</span>';
  const styles = {
    Cascade:    'background:#fee2e2;color:#b91c1c',
    Active:     'background:#fef3c7;color:#92400e',
    UserOwned:  'background:#ede9fe;color:#6d28d9',
    RemoveLink: 'background:#e0f2fe;color:#0369a1',
    Restrict:   'background:#dcfce7;color:#15803d',
    NoCascade:  'background:#f3f4f6;color:#6b7280',
  };
  const s = styles[value] ?? 'background:#f3f4f6;color:#6b7280';
  return `<span class="req-badge" style="${s}">${esc(value)}</span>`;
}

// ─── Global Option Sets ───────────────────────────────────────────────────────

function getFilteredOptionSets() {
  const query      = document.getElementById('search-input').value.trim().toLowerCase();
  const customOnly = document.getElementById('filter-os-custom').checked;
  const hideManaged= document.getElementById('filter-os-managed').checked;

  return allOptionSets.filter(os => {
    if (customOnly  && !os.IsCustomOptionSet) return false;
    if (hideManaged &&  os.IsManaged)         return false;
    if (!query) return true;
    return getLabel(os.DisplayName).toLowerCase().includes(query)
        || (os.Name ?? '').toLowerCase().includes(query);
  });
}

function renderOptionSetList() {
  const filtered = getFilteredOptionSets();
  if (sidebarMode === 'optionsets') {
    document.getElementById('entity-count').textContent = `${filtered.length} / ${allOptionSets.length}`;
  }

  const list = document.getElementById('optionset-list');
  list.innerHTML = '';

  filtered.forEach(os => {
    const displayName = getLabel(os.DisplayName) || os.Name;
    const optCount    = os.Options?.length ?? 0;

    const li = document.createElement('li');
    li.className   = 'entity-item';
    li.role        = 'option';
    li.dataset.name = os.Name;
    li.innerHTML   = `
      <div class="entity-item__display">${esc(displayName)}</div>
      <div class="entity-item__logical">${esc(os.Name ?? '')}</div>
      <div class="optset-item__count">${optCount} option${optCount !== 1 ? 's' : ''}</div>
      <div class="entity-item__badges">
        ${os.IsCustomOptionSet ? '<span class="badge badge--custom">Custom</span>'  : ''}
        ${os.IsManaged         ? '<span class="badge badge--managed">Managed</span>': ''}
      </div>`;
    li.addEventListener('click', () => selectOptionSet(os));
    list.appendChild(li);
  });
}

function selectOptionSet(os) {
  currentOptionSet = os;
  markListItemActive('optionset-list', os.Name, 'name');

  showDetailLoading();

  // Render immediately — no API call needed, options came with the list.
  renderOptionSetHeader(os);
  renderCurrentOptionSetOptions();

  hideDetailLoading();
  document.getElementById('entity-detail').classList.add('hidden');
  document.getElementById('optionset-detail').classList.remove('hidden');
}

function renderOptionSetHeader(os) {
  const displayName = getLabel(os.DisplayName) || os.Name;
  const description = getLabel(os.Description);

  document.getElementById('optionset-header').innerHTML = `
    <div class="entity-header__title">
      ${esc(displayName)}
      ${os.IsCustomOptionSet ? '<span class="badge badge--custom">Custom</span>'  : ''}
      ${os.IsManaged         ? '<span class="badge badge--managed">Managed</span>': ''}
    </div>
    <div class="entity-header__meta">
      ${metaItem('Name',           os.Name,        true)}
      ${metaItem('Type',           os['@odata.type']?.split('.').pop() ?? '')}
      ${description ? metaItem('Description', description) : ''}
    </div>`;
}

function renderCurrentOptionSetOptions() {
  if (!currentOptionSet) return;

  const query = document.getElementById('os-option-search').value.trim().toLowerCase();
  const allOpts = currentOptionSet.Options ?? [];

  const filtered = (query
    ? allOpts.filter(o => {
        const label = o.Label?.UserLocalizedLabel?.Label ?? '';
        return label.toLowerCase().includes(query) || String(o.Value).includes(query);
      })
    : allOpts
  ).slice().sort((a, b) =>
    (a.Label?.UserLocalizedLabel?.Label ?? '').localeCompare(b.Label?.UserLocalizedLabel?.Label ?? '')
  );

  document.getElementById('os-option-count').textContent = `${filtered.length} / ${allOpts.length} options`;

  const tbody = document.getElementById('os-tbody');
  tbody.innerHTML = '';

  filtered.forEach(o => {
    const label       = o.Label?.UserLocalizedLabel?.Label ?? '—';
    const description = o.Description?.UserLocalizedLabel?.Label ?? '';
    const colorValid  = o.Color && o.Color !== '#00000000';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td data-copytext="${esc(String(o.Value))}">
        <span class="os-value-pill">${esc(String(o.Value))}</span>
      </td>
      <td data-copytext="${esc(label)}">
        <strong>${esc(label)}</strong>
      </td>
      <td data-copytext="${esc(o.Color ?? '')}">
        ${colorValid
          ? `<span class="os-color-swatch" style="background:${esc(o.Color)}" title="${esc(o.Color)}"></span>
             <span style="font-size:11px;color:var(--text-muted);font-family:monospace;margin-left:5px">${esc(o.Color)}</span>`
          : '<span style="color:var(--border)">—</span>'}
      </td>
      <td data-copytext="${esc(description)}">${description ? esc(description) : '<span style="color:var(--border)">—</span>'}</td>`;

    tbody.appendChild(tr);
  });

  addCopyButtons(document.getElementById('os-table'));
  makeTableResizable(document.getElementById('os-table'));
}

// ─── Copy-to-clipboard on hover ───────────────────────────────────────────────

/**
 * Adds a hover-revealed copy button to every <tbody> <td> in the given table.
 *
 * Uses a `data-copytext` attribute on each <td> (set during row construction)
 * for the value to copy, falling back to the cell's plain textContent.
 *
 * Buttons are appended once and reused if the table is re-rendered; calling
 * this on an already-processed table is safe (old buttons are replaced with
 * the new DOM nodes from innerHTML assignment).
 */
function addCopyButtons(table) {
  if (!table) return;
  table.querySelectorAll('tbody td').forEach(td => {
    // Skip cells that already have a copy button from a previous render pass.
    if (td.querySelector('.cell-copy')) return;
    // Skip cells that manage their own inline copy buttons.
    if (td.hasAttribute('data-no-autocopy')) return;

    // Determine what text to copy: data-copytext > plain textContent.
    const getCopyText = () =>
      td.dataset.copytext?.trim() ||
      td.textContent.replace(/\s+/g, ' ').trim();

    if (!getCopyText()) return; // nothing to copy

    const btn = document.createElement('button');
    btn.className = 'cell-copy';
    btn.title     = 'Copy to clipboard';
    btn.innerHTML = '&#128203;';

    btn.addEventListener('click', e => {
      e.stopPropagation();
      copyToClipboard(getCopyText(), btn);
    });

    td.appendChild(btn);
  });
}

// ─── Resizable columns ────────────────────────────────────────────────────────

/**
 * Adds drag handles to every <th> in a table so columns can be resized.
 * Safe to call multiple times on the same table (guarded by _resizeDone).
 */
function makeTableResizable(table) {
  if (!table || table._resizeDone) return;
  table._resizeDone = true;

  const ths  = Array.from(table.querySelectorAll('thead th'));
  const cols = Array.from(table.querySelectorAll('col'));

  ths.forEach((th, i) => {
    const handle = document.createElement('div');
    handle.className = 'col-resizer';
    th.appendChild(handle);

    handle.addEventListener('mousedown', e => {
      e.preventDefault();
      const startX = e.clientX;
      const startW = th.offsetWidth;

      function onMove(ev) {
        const newW = Math.max(40, startW + ev.clientX - startX);
        if (cols[i]) cols[i].style.width = newW + 'px';
      }
      function onUp() {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup',   onUp);
      }
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup',   onUp);
    });
  });
}

// ─── Fetch bridge (web-app) ───────────────────────────────────────────────────

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

// ─── D365 Web API helpers ─────────────────────────────────────────────────────

/** Fetches all OData pages, following @odata.nextLink automatically. */
async function fetchAllPages(path) {
  const results = [];
  // Accept both relative paths (/EntityDefinitions...) and full URLs (nextLink).
  let url = path.startsWith('http') ? path : `${envUrl}/api/data/v9.2${path}`;

  while (url) {
    const data = await _d365Fetch(url);
    results.push(...(data.value ?? []));
    url = data['@odata.nextLink'] ?? null;
  }

  return results;
}

// ─── UI helpers ───────────────────────────────────────────────────────────────

function showDetailLoading() {
  document.getElementById('empty-state').classList.add('hidden');
  document.getElementById('entity-detail').classList.add('hidden');
  document.getElementById('optionset-detail').classList.add('hidden');
  document.getElementById('detail-loading').classList.remove('hidden');
}

function hideDetailLoading() {
  document.getElementById('detail-loading').classList.add('hidden');
}

function showEmptyState(title, sub) {
  document.getElementById('empty-state').classList.remove('hidden');
  document.querySelector('.empty-state__title').textContent = title;
  document.querySelector('.empty-state__sub').textContent   = sub;
}

function showGlobalError(message) {
  document.getElementById('global-loading').classList.add('hidden');
  document.getElementById('error-text').textContent = message;
  document.getElementById('global-error').classList.remove('hidden');
}

function setLoadingText(text) {
  document.getElementById('loading-text').textContent = text;
}

function markListItemActive(listId, value, dataAttr = 'ln') {
  document.querySelectorAll(`#${listId} .entity-item`).forEach(li => {
    li.classList.toggle('entity-item--active', li.dataset[dataAttr] === value);
  });
}

let copyFeedbackTimer = null;

function copyToClipboard(text, triggerEl) {
  navigator.clipboard.writeText(text).then(() => {
    if (!triggerEl) return;
    const orig = triggerEl.innerHTML;
    triggerEl.innerHTML = '✓';
    triggerEl.classList.add('cell-copy--copied');
    clearTimeout(copyFeedbackTimer);
    copyFeedbackTimer = setTimeout(() => {
      triggerEl.innerHTML = orig;
      triggerEl.classList.remove('cell-copy--copied');
    }, 1300);
  }).catch(() => {});
}

// ─── Utility ──────────────────────────────────────────────────────────────────

function getLabel(labelObj) {
  if (!labelObj) return '';
  return (
    labelObj.UserLocalizedLabel?.Label ??
    labelObj.LocalizedLabels?.[0]?.Label ??
    ''
  );
}

function esc(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Like esc() but also escapes single quotes — safe for JSON stored in data-* attributes. */
function escAttr(str) {
  return esc(str).replace(/'/g, '&#39;');
}
