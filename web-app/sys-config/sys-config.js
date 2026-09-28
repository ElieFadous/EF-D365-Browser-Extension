/**
 * EF Power Platform Tools — System Configurator (web-app)
 *
 * For tables that store system configuration as data. You choose which
 * tables and columns to show (with order, default sort and which columns are
 * editable); the tool then lists those tables in a side pane and shows each
 * one's records in a grid with search, per-column filters, sorting, inline
 * edits and row actions (open, clone, deactivate/activate).
 *
 * Settings are stored per environment in the D365 org's own localStorage
 * via the launcher (tool pages' own storage is partitioned when embedded).
 * `settings.systemConfigurator` in the web-app config seeds any environment
 * that has nothing saved yet; Import / Export copies settings between envs.
 */
(function () {
  'use strict';

  const D = window.EFD365;
  const esc = D.escHtml;
  const STORE_KEY = 'ef_ppt_tool_sysconfig';

  const SYSTEM_ATTRS = new Set([
    'createdon', 'createdby', 'createdonbehalfby', 'modifiedon', 'modifiedby', 'modifiedonbehalfby',
    'overriddencreatedon', 'versionnumber', 'importsequencenumber', 'timezoneruleversionnumber',
    'utcconversiontimezonecode', 'owningbusinessunit', 'owninguser', 'owningteam', 'ownerid',
    'statecode', 'statuscode', 'exchangerate', 'transactioncurrencyid', 'solutionid',
    'overwritetime', 'componentstate', 'ismanaged', 'iscustomizable',
  ]);
  const EDITABLE_TYPES = new Set(['String', 'Memo', 'Integer', 'BigInt', 'Decimal', 'Double', 'Money', 'Boolean', 'Picklist', 'DateTime']);
  const NUMERIC_TYPES  = new Set(['Integer', 'BigInt', 'Decimal', 'Double', 'Money']);
  const LOOKUP_TYPES   = new Set(['Lookup', 'Customer', 'Owner']);
  const SKIP_TYPES     = new Set(['EntityName', 'ManagedProperty', 'CalendarRules', 'PartyList']);
  const CHOICE_TYPES   = new Set(['Picklist', 'Boolean', 'State', 'Status']);

  let settings = { version: 1, tables: [] };
  let settingsSource = 'none'; // 'org' | 'config' | 'none'

  const metaCache = new Map();  // logicalName -> Promise<meta>
  let entityListPromise = null;

  // Viewer state
  let activeTable = null;       // logicalName
  let activeView = 'active';
  let grid = null;
  let currentSelect = '';
  let loadToken = 0;
  let tableCtx = null;          // { meta, table, cols } for the grid on screen
  let query = { search: '', filters: {}, sort: null };
  let queryTimer = null;
  const PAGE_SIZE = 250;

  // Settings-editor state
  let draft = null;
  let draftSel = null;
  let availableFilter = '';

  // ════════════════════════════════════════════════════════════════════════
  //  Bootstrap
  // ════════════════════════════════════════════════════════════════════════

  document.addEventListener('DOMContentLoaded', async function () {
    document.getElementById('env-url').textContent = D.envUrl;
    document.getElementById('env-name').textContent = D.envName;
    document.title = 'System Configurator — ' + D.envName;

    document.getElementById('btn-configure').addEventListener('click', function () {
      if (document.getElementById('settings-view').classList.contains('hidden')) openSettings();
      else closeSettings(false);
    });
    document.getElementById('btn-empty-configure').addEventListener('click', openSettings);
    document.getElementById('view-select').addEventListener('change', function (e) {
      activeView = e.target.value;
      loadTable();
    });
    document.getElementById('search').addEventListener('input', function (e) { if (grid) grid.setSearch(e.target.value); });
    document.getElementById('btn-clear').addEventListener('click', function () {
      document.getElementById('search').value = '';
      if (grid) grid.clearFilters();
    });
    document.getElementById('btn-refresh').addEventListener('click', loadTable);
    document.getElementById('table-list').addEventListener('click', function (e) {
      const li = e.target.closest('li[data-table]');
      if (!li || li.dataset.table === activeTable) return;
      selectTable(li.dataset.table);
    });
    document.getElementById('grid-wrap').addEventListener('click', onRowAction);
    document.getElementById('actionbar').addEventListener('click', onBulkAction);
    wireSettings();

    if (!D.envUrl) {
      bootState('No environment URL provided. Open this tool from the EF Power Platform Tools launcher.', true);
      return;
    }

    await loadSettings();
    document.getElementById('boot-state').classList.add('hidden');
    showViewer();
  });

  function bootState(msg, isError) {
    const el = document.getElementById('boot-state');
    el.textContent = msg;
    el.className = 'state-msg' + (isError ? ' state-error' : '');
  }

  // ════════════════════════════════════════════════════════════════════════
  //  Settings persistence
  // ════════════════════════════════════════════════════════════════════════

  function normalizeSettings(s) {
    const tables = (s && Array.isArray(s.tables) ? s.tables : [])
      .filter(function (t) { return t && typeof t.logicalName === 'string' && t.logicalName; })
      .map(function (t) {
        return {
          logicalName: t.logicalName,
          columns: (Array.isArray(t.columns) ? t.columns : [])
            .filter(function (c) { return c && typeof c.name === 'string' && c.name; })
            .map(function (c) { return { name: c.name, editable: !!c.editable, json: !!c.json }; }),
          sort: t.sort && t.sort.column ? { column: t.sort.column, dir: t.sort.dir === 'desc' ? 'desc' : 'asc' } : null,
        };
      });
    return { version: 1, tables: tables };
  }

  async function loadSettings() {
    try {
      const raw = await D.storeGet(STORE_KEY);
      if (raw) {
        settings = normalizeSettings(JSON.parse(raw));
        settingsSource = 'org';
        return;
      }
    } catch (err) {
      console.warn('[EF PPT] Could not read System Configurator settings:', err);
    }
    const seed = D.cfg.settings && D.cfg.settings.systemConfigurator;
    if (seed) {
      settings = normalizeSettings(seed);
      settingsSource = 'config';
    }
  }

  async function saveSettings(next) {
    await D.storeSet(STORE_KEY, JSON.stringify(next));
    settings = next;
    settingsSource = 'org';
  }

  // ════════════════════════════════════════════════════════════════════════
  //  Metadata
  // ════════════════════════════════════════════════════════════════════════

  function label(l) {
    return (l && l.UserLocalizedLabel && l.UserLocalizedLabel.Label) || '';
  }

  function loadMeta(logicalName) {
    if (!metaCache.has(logicalName)) {
      metaCache.set(logicalName, fetchMeta(logicalName).catch(function (err) {
        metaCache.delete(logicalName);
        throw err;
      }));
    }
    return metaCache.get(logicalName);
  }

  async function fetchMeta(ln) {
    const base = "/EntityDefinitions(LogicalName='" + ln + "')";
    const none = function () { return { value: [] }; };
    const results = await Promise.all([
      D.request(base + '?$select=LogicalName,DisplayName,DisplayCollectionName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute'),
      D.request(base + '/Attributes?$select=LogicalName,DisplayName,AttributeType,AttributeTypeName,IsValidForRead,IsValidForUpdate,AttributeOf'),
      D.request(base + '/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet').catch(none),
      D.request(base + '/Attributes/Microsoft.Dynamics.CRM.BooleanAttributeMetadata?$select=LogicalName&$expand=OptionSet').catch(none),
      D.request(base + '/Attributes/Microsoft.Dynamics.CRM.StateAttributeMetadata?$select=LogicalName&$expand=OptionSet').catch(none),
      D.request(base + '/Attributes/Microsoft.Dynamics.CRM.StatusAttributeMetadata?$select=LogicalName&$expand=OptionSet').catch(none),
    ]);
    const def = results[0];

    // Choice labels for Picklist, State and Status columns — used by the edit
    // dropdowns and to translate filters/searches into server-side values.
    const options = {};
    [results[2], results[4], results[5]].forEach(function (res) {
      (res.value || []).forEach(function (p) {
        const os = p.OptionSet || p.GlobalOptionSet;
        options[p.LogicalName] = ((os && os.Options) || []).map(function (o) {
          return { value: String(o.Value), label: label(o.Label) || String(o.Value) };
        });
      });
    });
    const booleans = {};
    (results[3].value || []).forEach(function (b) {
      const os = b.OptionSet || {};
      booleans[b.LogicalName] = {
        t: label(os.TrueOption && os.TrueOption.Label) || 'Yes',
        f: label(os.FalseOption && os.FalseOption.Label) || 'No',
      };
    });
    let hasState = false, activeStatus = null, inactiveStatus = null;
    (results[4].value || []).forEach(function (s) {
      hasState = true;
      ((s.OptionSet && s.OptionSet.Options) || []).forEach(function (o) {
        if (o.Value === 0) activeStatus = o.DefaultStatus;
        if (o.Value === 1) inactiveStatus = o.DefaultStatus;
      });
    });

    const attrs = (results[1].value || [])
      .filter(function (a) {
        if (a.IsValidForRead === false || a.AttributeOf || SKIP_TYPES.has(a.AttributeType)) return false;
        if (a.AttributeType === 'Virtual') return a.AttributeTypeName && a.AttributeTypeName.Value === 'MultiSelectPicklistType';
        return true;
      })
      .map(function (a) {
        const type = a.AttributeType === 'Virtual' ? 'MultiSelectPicklist' : a.AttributeType;
        let reason = '';
        if (SYSTEM_ATTRS.has(a.LogicalName) || a.LogicalName === def.PrimaryIdAttribute) reason = 'System field — read-only';
        else if (!a.IsValidForUpdate) reason = 'This column can’t be updated';
        else if (LOOKUP_TYPES.has(type)) reason = 'Lookups can’t be edited inline';
        else if (!EDITABLE_TYPES.has(type)) reason = 'This column type can’t be edited inline';
        return {
          name: a.LogicalName,
          label: label(a.DisplayName) || a.LogicalName,
          type: type,
          canEdit: !reason,
          editReason: reason,
          isSystem: SYSTEM_ATTRS.has(a.LogicalName) || a.LogicalName === def.PrimaryIdAttribute,
        };
      })
      .sort(function (x, y) { return x.label.localeCompare(y.label, undefined, { sensitivity: 'base' }); });

    const attrByName = new Map();
    attrs.forEach(function (a) { attrByName.set(a.name, a); });

    return {
      logicalName:    ln,
      label:          label(def.DisplayName) || ln,
      plural:         label(def.DisplayCollectionName) || label(def.DisplayName) || ln,
      entitySet:      def.EntitySetName,
      primaryId:      def.PrimaryIdAttribute,
      primaryName:    def.PrimaryNameAttribute,
      hasState:       hasState,
      activeStatus:   activeStatus,
      inactiveStatus: inactiveStatus,
      attrs:          attrs,
      attrByName:     attrByName,
      options:        options,
      booleans:       booleans,
    };
  }

  function loadEntityList() {
    if (!entityListPromise) {
      entityListPromise = D.request('/EntityDefinitions?$select=LogicalName,DisplayName,DisplayCollectionName,IsIntersect,IsPrivate')
        .then(function (d) {
          return (d.value || [])
            .filter(function (e) { return !e.IsIntersect && !e.IsPrivate; })
            .map(function (e) {
              return { name: e.LogicalName, label: label(e.DisplayCollectionName) || label(e.DisplayName) || e.LogicalName };
            })
            .sort(function (a, b) { return a.label.localeCompare(b.label, undefined, { sensitivity: 'base' }); });
        })
        .catch(function (err) { entityListPromise = null; throw err; });
    }
    return entityListPromise;
  }

  // ════════════════════════════════════════════════════════════════════════
  //  Viewer
  // ════════════════════════════════════════════════════════════════════════

  function showViewer() {
    const hasTables = settings.tables.length > 0;
    document.getElementById('settings-view').classList.add('hidden');
    document.getElementById('viewer').classList.toggle('hidden', !hasTables);
    document.getElementById('empty-view').classList.toggle('hidden', hasTables);
    document.getElementById('btn-configure').innerHTML = '&#9881; Configure';
    if (!hasTables) return;

    if (!activeTable || !settings.tables.some(function (t) { return t.logicalName === activeTable; })) {
      activeTable = settings.tables[0].logicalName;
    }
    renderTableList();
    selectTable(activeTable);
  }

  function renderTableList() {
    const ul = document.getElementById('table-list');
    ul.innerHTML = settings.tables.map(function (t) {
      return (
        '<li data-table="' + esc(t.logicalName) + '"' + (t.logicalName === activeTable ? ' class="active"' : '') + '>' +
          '<div class="sc-item-text">' +
            '<div class="sc-item-name" data-label="' + esc(t.logicalName) + '">' + esc(t.logicalName) + '</div>' +
            '<div class="sc-item-sub">' + esc(t.logicalName) + '</div>' +
          '</div>' +
        '</li>'
      );
    }).join('');
    // Fill in friendly names as metadata arrives.
    settings.tables.forEach(function (t) {
      loadMeta(t.logicalName).then(function (m) {
        const el = ul.querySelector('[data-label="' + CSS.escape(t.logicalName) + '"]');
        if (el) el.textContent = m.plural;
      }).catch(function () { /* shown when selected */ });
    });
  }

  async function selectTable(logicalName) {
    activeTable = logicalName;
    activeView = 'active';
    tableCtx = null;
    document.getElementById('search').value = '';
    document.querySelectorAll('#table-list li').forEach(function (li) {
      li.classList.toggle('active', li.dataset.table === logicalName);
    });

    const token = ++loadToken;
    const table = tableConfig(logicalName);
    if (!table) return;
    viewerState('Loading…', 'loading');
    setCount('');

    let meta;
    try {
      meta = await loadMeta(table.logicalName);
    } catch (err) {
      if (token === loadToken) viewerState('Could not load table “' + table.logicalName + '”: ' + err.message, 'error');
      return;
    }
    if (token !== loadToken) return;

    const viewSel = document.getElementById('view-select');
    viewSel.innerHTML =
      '<option value="active">Active ' + esc(meta.plural) + '</option>' +
      (meta.hasState ? '<option value="inactive">Inactive ' + esc(meta.plural) + '</option>' : '');
    viewSel.value = 'active';
    viewSel.disabled = !meta.hasState;

    const cols = table.columns
      .map(function (c) { return { cfg: c, attr: meta.attrByName.get(c.name) }; })
      .filter(function (x) { return x.attr; });
    if (!cols.length) {
      viewerState('No columns are configured for this table. Use Configure to choose some.', 'info');
      return;
    }

    const select = new Set([meta.primaryId]);
    if (meta.primaryName) select.add(meta.primaryName);
    cols.forEach(function (x) { select.add(selectName(x.attr)); });
    currentSelect = Array.from(select).join(',');

    tableCtx = { meta: meta, table: table, cols: cols };
    buildGrid(meta, table, cols);
    updateActionBar([]);
    await fetchPage(true);
  }

  function loadTable() {
    return fetchPage(false);
  }

  function viewerState(msg, kind) {
    const el = document.getElementById('viewer-state');
    const wrap = document.getElementById('grid-wrap');
    if (!msg) { el.classList.add('hidden'); wrap.classList.remove('hidden'); return; }
    el.textContent = msg;
    el.className = 'state-msg' + (kind ? ' state-' + kind : '');
    wrap.classList.add('hidden');
  }

  function setCount(text, isError) {
    const el = document.getElementById('count');
    el.textContent = text;
    el.className = 'count' + (isError ? ' count--error' : '');
  }

  function selectName(a) {
    return LOOKUP_TYPES.has(a.type) ? '_' + a.name + '_value' : a.name;
  }

  function cellDisplay(a, rec) {
    const k = selectName(a);
    const f = rec[k + '@' + 'OData.Community.Display.V1.FormattedValue'];
    if (f != null) return f;
    const raw = rec[k];
    return raw == null ? '' : String(raw);
  }

  function tableConfig(ln) {
    return settings.tables.filter(function (t) { return t.logicalName === ln; })[0];
  }

  // ── Server-side query ───────────────────────────────────────────────────
  // Only PAGE_SIZE rows are ever loaded, so search, column filters and sorting
  // all run in D365 ($filter / $orderby) — filtering the loaded rows would
  // silently miss everything past the first page.

  function odataString(s) {
    return "'" + String(s).replace(/'/g, "''") + "'";
  }

  function isJsonCol(x) {
    return x.cfg.json && (x.attr.type === 'String' || x.attr.type === 'Memo');
  }

  function choiceOptions(meta, a) {
    if (a.type === 'Boolean') {
      const b = meta.booleans[a.name] || { t: 'Yes', f: 'No' };
      return [{ value: 'true', label: b.t }, { value: 'false', label: b.f }];
    }
    return meta.options[a.name] || [];
  }

  /** OData condition matching `q` in one column, or null if this column can't match that text. */
  function termFor(meta, a, q) {
    const k = selectName(a);
    const ql = q.toLowerCase();
    switch (a.type) {
      case 'String':
      case 'Memo':
        return 'contains(' + a.name + ',' + odataString(q) + ')';
      case 'Integer':
      case 'BigInt':
        return /^-?\d+$/.test(q) ? k + ' eq ' + q : null;
      case 'Decimal':
      case 'Double':
      case 'Money':
        return q !== '' && isFinite(Number(q)) ? k + ' eq ' + Number(q) : null;
      case 'Boolean':
      case 'Picklist':
      case 'State':
      case 'Status': {
        const hits = choiceOptions(meta, a).filter(function (o) { return o.label.toLowerCase().indexOf(ql) !== -1; });
        if (!hits.length) return null;
        return hits.map(function (o) { return k + ' eq ' + o.value; }).join(' or ');
      }
      case 'Uniqueidentifier':
        return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(q) ? k + ' eq ' + q : null;
      case 'DateTime':
        return /^\d{4}-\d{2}-\d{2}$/.test(q)
          ? 'Microsoft.Dynamics.CRM.On(PropertyName=' + odataString(a.name) + ',PropertyValue=' + odataString(q) + ')'
          : null;
      default:
        return null;
    }
  }

  function filterClause(meta, a, f, noMatch) {
    const k = selectName(a);
    if (CHOICE_TYPES.has(a.type)) {
      if (!f.value) return null;
      return f.value === '__blank__' ? k + ' eq null' : k + ' eq ' + f.value;
    }
    const parts = [];
    if (f.mode === 'blank') parts.push(k + ' eq null');
    else if (f.mode === 'notblank') parts.push(k + ' ne null');
    const text = (f.text || '').trim();
    if (text) parts.push(termFor(meta, a, text) || noMatch);
    return parts.length ? parts.join(' and ') : null;
  }

  function buildQueryUrl() {
    const meta = tableCtx.meta, cols = tableCtx.cols;
    const noMatch = meta.primaryId + ' eq null'; // matches nothing — for text no column can contain
    const clauses = [];
    if (meta.hasState) clauses.push('statecode eq ' + (activeView === 'inactive' ? 1 : 0));

    const search = (query.search || '').trim();
    if (search) {
      const terms = cols.map(function (x) { return termFor(meta, x.attr, search); }).filter(Boolean);
      clauses.push(terms.length ? terms.map(function (t) { return '(' + t + ')'; }).join(' or ') : noMatch);
    }
    cols.forEach(function (x) {
      const f = query.filters['c_' + x.attr.name];
      const c = f && filterClause(meta, x.attr, f, noMatch);
      if (c) clauses.push(c);
    });

    let url = '/' + meta.entitySet + '?$select=' + currentSelect + '&$count=true';
    if (clauses.length) url += '&$filter=' + encodeURIComponent(clauses.map(function (c) { return '(' + c + ')'; }).join(' and '));

    const sortX = query.sort && cols.filter(function (x) { return 'c_' + x.attr.name === query.sort.key; })[0];
    const order = [];
    if (sortX) order.push(selectName(sortX.attr) + ' ' + (query.sort.dir === 'desc' ? 'desc' : 'asc'));
    if (!sortX || sortX.attr.name !== meta.primaryId) order.push(meta.primaryId + ' asc'); // stable order
    url += '&$orderby=' + encodeURIComponent(order.join(','));
    return url;
  }

  async function fetchPage(initial) {
    if (!tableCtx) return;
    const token = ++loadToken;
    const meta = tableCtx.meta;
    if (initial) viewerState('Loading…', 'loading');
    else setCount('Loading…');

    let d;
    try {
      d = await D.request(buildQueryUrl(), {
        headers: { Prefer: 'odata.include-annotations="*",odata.maxpagesize=' + PAGE_SIZE },
      });
    } catch (err) {
      if (token !== loadToken) return;
      if (initial) viewerState('Failed to load records: ' + err.message, 'error');
      else setCount('Search failed: ' + err.message, true);
      return;
    }
    if (token !== loadToken) return;

    const rows = (d.value || []).slice(0, PAGE_SIZE).map(function (rec) {
      return { id: rec[meta.primaryId], name: (meta.primaryName && rec[meta.primaryName]) || rec[meta.primaryId], rec: rec };
    });
    viewerState(null);
    grid.setRows(rows);
    updateCount(rows.length, d['@odata.count'], d['@Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded'] === true);
  }

  function updateCount(shown, total, exceeded) {
    const kind = activeView === 'inactive' ? 'inactive' : 'active';
    const hasTotal = typeof total === 'number';
    const n = hasTotal ? total : shown;
    const noun = kind + ' ' + (n === 1 && !exceeded ? 'record' : 'records');
    if (!exceeded && (!hasTotal || shown >= total)) {
      setCount(n.toLocaleString() + ' ' + noun);
      return;
    }
    const totalText = exceeded ? '5,000+' : total.toLocaleString();
    setCount('Showing ' + shown.toLocaleString() + ' of ' + totalText + ' ' + noun + ' — search or filter to find others');
  }

  function filterPlaceholder(a) {
    if (NUMERIC_TYPES.has(a.type)) return '= number';
    if (a.type === 'DateTime') return 'YYYY-MM-DD';
    if (a.type === 'Uniqueidentifier') return 'GUID';
    return 'Contains…';
  }

  function buildGrid(meta, table, cols) {
    const wrap = document.getElementById('grid-wrap');
    wrap.innerHTML = '';
    const host = document.createElement('div');
    host.className = 'grid-host';
    wrap.appendChild(host);

    const columns = cols.map(function (x) {
      const a = x.attr;
      const k = selectName(a);
      const col = {
        key: 'c_' + a.name,
        label: a.label,
        value: function (r) { return cellDisplay(a, r.rec); },
        sortable: a.type !== 'Memo' && a.type !== 'MultiSelectPicklist',
        filter: CHOICE_TYPES.has(a.type) ? 'select' : 'value',
        filterOptions: CHOICE_TYPES.has(a.type) ? choiceOptions(meta, a) : null,
        filterNoText: LOOKUP_TYPES.has(a.type) || a.type === 'MultiSelectPicklist',
        filterPlaceholder: filterPlaceholder(a),
      };
      const editable = !!(x.cfg.editable && a.canEdit);
      if (isJsonCol(x)) {
        // JSON columns are viewed/edited in the JSON viewer rather than inline.
        col.render = function (r) { return jsonCellHtml(a, r, editable); };
      } else if (editable) {
        col.editable = true;
        col.editValue = function (r) { const raw = r.rec[k]; return raw == null ? '' : String(raw); };
        if (a.type === 'Boolean' || a.type === 'Picklist') {
          col.editOptions = function () {
            const opts = choiceOptions(meta, a);
            return a.type === 'Picklist' ? [{ value: '', label: '(none)' }].concat(opts) : opts;
          };
        }
        col.editHint = editHint(a);
        col.edit = function (r, text) { return saveCell(meta, a, r, text); };
      }
      return col;
    });

    const sortCfg = table.sort && cols.some(function (x) { return x.attr.name === table.sort.column; })
      ? { key: 'c_' + table.sort.column, dir: table.sort.dir }
      : { key: 'c_' + cols[0].attr.name, dir: 'asc' };

    grid = new window.DataGrid(host, {
      columns: columns,
      rowKey: function (r) { return r.id; },
      defaultSort: sortCfg,
      tableMinWidth: (cols.length * 130 + 40) + 'px',
      selectable: true,
      onSelectionChange: updateActionBar,
      serverMode: true,
      onQueryChange: function (q) {
        query = q;
        clearTimeout(queryTimer);
        queryTimer = setTimeout(function () { fetchPage(false); }, 350);
      },
    });
    query = grid.query();
  }

  // ── JSON columns ───────────────────────────────────────────────────────

  function jsonCellHtml(a, r, editable) {
    const v = r.rec[a.name];
    const btn = '<button type="button" class="sc-json-btn" data-sc-json="' + esc(a.name) + '" data-id="' + esc(r.id) + '" title="' +
      (editable ? 'View / edit JSON' : 'View JSON') + '">{ }</button>';
    if (v == null || v === '') return editable ? '<div class="sc-json-cell">' + btn + '<span class="dg-blank">—</span></div>' : '<span class="dg-blank">—</span>';
    const preview = String(v).replace(/\s+/g, ' ').slice(0, 300);
    return '<div class="sc-json-cell">' + btn + '<span class="dg-text mono" title="Open the JSON viewer to see it formatted">' + esc(preview) + '</span></div>';
  }

  /** Pretty JSON with syntax colouring. Input is escaped first; quotes are left alone so strings still tokenise. */
  function highlightJson(text) {
    const safe = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    return safe.replace(/("(?:\\u[\da-fA-F]{4}|\\[^u]|[^\\"])*")(\s*:)?|\b(true|false)\b|\bnull\b|-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/g,
      function (m, str, colon, bool) {
        if (str) return colon ? '<span class="j-key">' + str + '</span>' + colon : '<span class="j-str">' + str + '</span>';
        if (bool) return '<span class="j-bool">' + m + '</span>';
        if (m === 'null') return '<span class="j-null">null</span>';
        return '<span class="j-num">' + m + '</span>';
      });
  }

  function jsonLinesHtml(highlighted) {
    return highlighted.split('\n').map(function (line, i) {
      return '<div class="j-line"><span class="j-ln">' + (i + 1) + '</span><span class="j-code">' + (line || ' ') + '</span></div>';
    }).join('');
  }

  function openJsonViewer(meta, a, row, editable) {
    const dlg = showDialog(
      '<div class="sc-dialog-title">' + esc(a.label) + ' <span class="sc-col-sub">' + esc(row.name) + '</span></div>' +
      '<div class="sc-dialog-body sc-json-body" id="sc-json-body"></div>' +
      '<div class="sc-dialog-actions" id="sc-json-actions"></div>',
      'xl'
    );
    const body = dlg.el.querySelector('#sc-json-body');
    const actions = dlg.el.querySelector('#sc-json-actions');

    function current() { const v = row.rec[a.name]; return v == null ? '' : String(v); }

    function view() {
      const raw = current();
      let parsed, ok = true;
      if (raw.trim()) { try { parsed = JSON.parse(raw); } catch (_) { ok = false; } }
      if (!raw.trim()) {
        body.innerHTML = '<div class="sc-json-empty">This column is empty.</div>';
      } else if (ok) {
        body.innerHTML = '<pre class="sc-json">' + jsonLinesHtml(highlightJson(JSON.stringify(parsed, null, 2))) + '</pre>';
      } else {
        body.innerHTML = '<div class="sc-status sc-status--warn">This value isn’t valid JSON — showing it as plain text.</div>' +
          '<pre class="sc-json sc-json--raw">' + esc(raw) + '</pre>';
      }
      actions.innerHTML =
        (raw.trim() ? '<button type="button" class="btn" data-j="copy">Copy</button>' : '') +
        (editable ? '<button type="button" class="btn" data-j="edit">Edit</button>' : '') +
        '<button type="button" class="btn btn--primary" data-j="close">Close</button>';
    }

    function edit() {
      const raw = current();
      let text = raw;
      try { if (raw.trim()) text = JSON.stringify(JSON.parse(raw), null, 2); } catch (_) { /* keep raw */ }
      body.innerHTML =
        '<textarea class="sc-json-editor" id="sc-json-text" spellcheck="false"></textarea>' +
        '<div id="sc-json-valid" class="sc-json-valid"></div>';
      const ta = body.querySelector('#sc-json-text');
      ta.value = text;
      actions.innerHTML =
        '<button type="button" class="btn" data-j="format">Format</button>' +
        '<span class="toolbar-spacer"></span>' +
        '<button type="button" class="btn" data-j="cancel">Cancel</button>' +
        '<button type="button" class="btn btn--primary" data-j="save">Save</button>';
      ta.addEventListener('input', validate);
      ta.addEventListener('keydown', function (e) {
        if (e.key === 'Tab') { // indent instead of leaving the editor
          e.preventDefault();
          const s = ta.selectionStart, en = ta.selectionEnd;
          ta.value = ta.value.slice(0, s) + '  ' + ta.value.slice(en);
          ta.selectionStart = ta.selectionEnd = s + 2;
          validate();
        }
      });
      validate();
      ta.focus();
    }

    function validate() {
      const ta = body.querySelector('#sc-json-text');
      const out = body.querySelector('#sc-json-valid');
      if (!ta || !out) return true;
      if (!ta.value.trim()) { out.className = 'sc-json-valid'; out.textContent = 'Empty — saving will clear this column.'; return true; }
      try {
        JSON.parse(ta.value);
        out.className = 'sc-json-valid ok';
        out.textContent = '✓ Valid JSON';
        return true;
      } catch (e) {
        out.className = 'sc-json-valid err';
        out.textContent = '✕ ' + e.message;
        return false;
      }
    }

    actions.addEventListener('click', async function (e) {
      const b = e.target.closest('[data-j]');
      if (!b) return;
      const act = b.dataset.j;
      if (act === 'close') dlg.close();
      else if (act === 'edit') edit();
      else if (act === 'cancel') view();
      else if (act === 'copy') {
        Promise.resolve(navigator.clipboard && navigator.clipboard.writeText(current()))
          .then(function () { b.textContent = 'Copied'; setTimeout(function () { b.textContent = 'Copy'; }, 1200); })
          .catch(function () { b.textContent = 'Copy failed'; });
      } else if (act === 'format') {
        const ta = body.querySelector('#sc-json-text');
        try { ta.value = JSON.stringify(JSON.parse(ta.value), null, 2); } catch (_) { /* validate() shows why */ }
        validate();
      } else if (act === 'save') {
        const ta = body.querySelector('#sc-json-text');
        if (!validate()) return;
        const out = body.querySelector('#sc-json-valid');
        b.disabled = true;
        b.textContent = 'Saving…';
        try {
          await saveCell(meta, a, row, ta.value.trim() ? ta.value : '');
          grid.refresh();
          view();
        } catch (err) {
          out.className = 'sc-json-valid err';
          out.textContent = 'Save failed: ' + err.message;
          b.disabled = false;
          b.textContent = 'Save';
        }
      }
    });

    view();
  }

  function editHint(a) {
    if (a.type === 'Boolean' || a.type === 'Picklist') return 'Pick a value, then Enter or Save. Esc to cancel.';
    if (a.type === 'DateTime') return 'Use YYYY-MM-DD or a full ISO date/time. Leave empty to clear. Enter to save, Esc to cancel.';
    if (NUMERIC_TYPES.has(a.type)) return 'Numbers only. Leave empty to clear. Enter to save, Esc to cancel.';
    return 'Enter to save · Shift+Enter for a new line · Esc to cancel. Leave empty to clear.';
  }

  function parseValue(a, text) {
    const t = text.trim();
    switch (a.type) {
      case 'String':
      case 'Memo':
        return text === '' ? null : text;
      case 'Integer':
      case 'BigInt':
        if (t === '') return null;
        if (!/^-?\d+$/.test(t)) throw new Error('"' + t + '" is not a whole number.');
        return parseInt(t, 10);
      case 'Decimal':
      case 'Double':
      case 'Money':
        if (t === '') return null;
        if (!isFinite(Number(t))) throw new Error('"' + t + '" is not a valid number.');
        return Number(t);
      case 'Boolean':
        return t === 'true';
      case 'Picklist':
        return t === '' ? null : parseInt(t, 10);
      case 'DateTime':
        if (t === '') return null;
        if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return t;
        if (isNaN(Date.parse(t))) throw new Error('"' + t + '" is not a valid date.');
        return new Date(t).toISOString();
      default:
        throw new Error('This column type can’t be edited here.');
    }
  }

  async function saveCell(meta, a, row, text) {
    const body = {};
    body[a.name] = parseValue(a, text);
    await D.request('/' + meta.entitySet + '(' + row.id + ')', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: body,
    });
    await refreshRow(meta, row);
  }

  /** Re-reads a record so formatted values (choices, dates, numbers) are exactly what D365 shows. */
  async function refreshRow(meta, row) {
    const rec = await D.request('/' + meta.entitySet + '(' + row.id + ')?$select=' + currentSelect, {
      headers: { Prefer: D.FORMATTED },
    });
    row.rec = rec;
    row.name = (meta.primaryName && rec[meta.primaryName]) || row.name;
  }

  // ── Row actions ────────────────────────────────────────────────────────

  function cloneAllowed(meta) {
    const wl = D.cfg.settings && D.cfg.settings.cloneWhitelist;
    return !Array.isArray(wl) || wl.indexOf(meta.logicalName.toLowerCase()) !== -1;
  }

  // ── Action bar — every action works on the selected rows ────────────────

  function updateActionBar(rows) {
    rows = rows || (grid ? grid.getSelectedRows() : []);
    const bar = document.getElementById('actionbar');
    const meta = tableCtx && tableCtx.meta;
    const n = rows.length;
    const btn = function (name) { return bar.querySelector('[data-bulk="' + name + '"]'); };

    btn('open').disabled = n === 0;

    const clone = btn('clone');
    const canClone = !!meta && cloneAllowed(meta);
    clone.disabled = n === 0 || !canClone;
    clone.title = meta && !canClone
      ? 'Cloning this table is turned off by the clone whitelist in your config'
      : 'Clone the selected records';

    const state = btn('state');
    const deactivate = activeView !== 'inactive';
    state.classList.toggle('hidden', !meta || !meta.hasState);
    state.classList.toggle('sc-cmd--danger', deactivate);
    state.querySelector('[data-bulk-label]').textContent = deactivate ? 'Deactivate' : 'Activate';
    state.title = (deactivate ? 'Deactivate' : 'Activate') + ' the selected records';
    state.disabled = n === 0;

    document.getElementById('sel-count').textContent = n
      ? n + (n === 1 ? ' record' : ' records') + ' selected'
      : 'Select rows to act on them';
    btn('clear').classList.toggle('hidden', n === 0);
  }

  function onBulkAction(e) {
    const btn = e.target.closest('[data-bulk]');
    if (!btn || btn.disabled || !grid || !tableCtx) return;
    const act = btn.dataset.bulk;
    if (act === 'clear') { grid.clearSelection(); return; }
    const rows = grid.getSelectedRows();
    if (!rows.length) return;
    const meta = tableCtx.meta;
    if (act === 'open') openRecords(meta, rows);
    else if (act === 'clone') openCloneDialog(meta, rows);
    else if (act === 'state') confirmStateChange(meta, rows, activeView !== 'inactive');
  }

  // JSON viewer buttons live inside cells.
  function onRowAction(e) {
    const jsonBtn = e.target.closest('[data-sc-json]');
    if (!jsonBtn || !grid || !tableCtx) return;
    const row = grid.byId.get(jsonBtn.dataset.id);
    const x = tableCtx.cols.filter(function (c) { return c.attr.name === jsonBtn.dataset.scJson; })[0];
    if (row && x) openJsonViewer(tableCtx.meta, x.attr, row, !!(x.cfg.editable && x.attr.canEdit));
  }

  function recordUrl(envUrl, meta, id) {
    return envUrl + '/main.aspx?pagetype=entityrecord&etn=' + encodeURIComponent(meta.logicalName) +
      '&id=' + encodeURIComponent(id);
  }

  function nameList(rows, max) {
    const shown = rows.slice(0, max).map(function (r) { return '<li>' + esc(r.name) + '</li>'; }).join('');
    const more = rows.length > max ? '<li class="sc-more">…and ' + (rows.length - max) + ' more</li>' : '';
    return '<ul class="sc-name-list">' + shown + more + '</ul>';
  }

  function openRecords(meta, rows) {
    const doOpen = function () {
      // Browsers usually allow one pop-up per click; any tab they block is
      // listed afterwards so it can still be opened by hand.
      const blocked = [];
      rows.forEach(function (r) {
        const w = window.open(recordUrl(D.envUrl, meta, r.id), '_blank');
        if (w) { try { w.opener = null; } catch (_) { /* cross-origin already */ } }
        else blocked.push(r);
      });
      if (!blocked.length) return;
      const dlg = showDialog(
        '<div class="sc-dialog-title">Some tabs were blocked</div>' +
        '<div class="sc-dialog-body">' +
          '<p>The browser blocked ' + blocked.length + ' of ' + rows.length + ' tabs. Allow pop-ups for this site to open them all at once next time, or open them here:</p>' +
          '<ul class="sc-name-list">' + blocked.map(function (r) {
            return '<li><a class="link" href="' + esc(recordUrl(D.envUrl, meta, r.id)) + '" target="_blank" rel="noopener">' + esc(r.name) + ' ↗</a></li>';
          }).join('') + '</ul>' +
        '</div>' +
        '<div class="sc-dialog-actions"><button type="button" class="btn btn--primary" data-dlg="ok">Close</button></div>'
      );
      dlg.el.querySelector('[data-dlg="ok"]').addEventListener('click', dlg.close);
    };

    if (rows.length <= 10) { doOpen(); return; }
    const dlg = showDialog(
      '<div class="sc-dialog-title">Open ' + rows.length + ' tabs?</div>' +
      '<div class="sc-dialog-body"><p>This opens one browser tab per selected record.</p></div>' +
      '<div class="sc-dialog-actions">' +
        '<button type="button" class="btn" data-dlg="no">Cancel</button>' +
        '<button type="button" class="btn btn--primary" data-dlg="yes">Open ' + rows.length + ' tabs</button>' +
      '</div>'
    );
    dlg.el.querySelector('[data-dlg="no"]').addEventListener('click', dlg.close);
    dlg.el.querySelector('[data-dlg="yes"]').addEventListener('click', function () { dlg.close(); doOpen(); });
  }

  function confirmStateChange(meta, rows, deactivate) {
    const verb = deactivate ? 'Deactivate' : 'Activate';
    const n = rows.length;
    const dlg = showDialog(
      '<div class="sc-dialog-title">' + verb + ' ' + (n === 1 ? 'record' : n + ' records') + '?</div>' +
      '<div class="sc-dialog-body">' +
        '<p>' + (n === 1 ? 'This record' : 'These records') + ' will move to <em>' + (deactivate ? 'Inactive' : 'Active') + ' ' + esc(meta.plural) + '</em>:</p>' +
        nameList(rows, 8) +
        '<div id="sc-state-status" class="sc-status hidden"></div>' +
      '</div>' +
      '<div class="sc-dialog-actions">' +
        '<button type="button" class="btn" data-dlg="no">No</button>' +
        '<button type="button" class="btn ' + (deactivate ? 'btn--danger' : 'btn--primary') + '" data-dlg="yes">Yes, ' + verb.toLowerCase() + '</button>' +
      '</div>'
    );
    const yes = dlg.el.querySelector('[data-dlg="yes"]');
    const no = dlg.el.querySelector('[data-dlg="no"]');
    const statusEl = dlg.el.querySelector('#sc-state-status');
    no.addEventListener('click', function () { if (!no.disabled) dlg.close(); });

    yes.addEventListener('click', async function () {
      yes.disabled = true; no.disabled = true;
      const body = { statecode: deactivate ? 1 : 0 };
      const status = deactivate ? meta.inactiveStatus : meta.activeStatus;
      if (status != null) body.statuscode = status;

      const failures = [];
      for (let i = 0; i < rows.length; i++) {
        yes.textContent = (deactivate ? 'Deactivating ' : 'Activating ') + (i + 1) + ' of ' + n + '…';
        try {
          await D.request('/' + meta.entitySet + '(' + rows[i].id + ')', {
            method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: body,
          });
        } catch (err) {
          failures.push({ row: rows[i], error: err.message });
        }
      }

      if (!failures.length) {
        dlg.close();
        grid.clearSelection();
        loadTable();
        return;
      }
      const done = n - failures.length;
      setStatus(statusEl, 'err',
        (done ? done + ' ' + (deactivate ? 'deactivated' : 'activated') + '. ' : '') + failures.length + ' failed:' +
        '<ul class="sc-name-list">' + failures.map(function (f) {
          return '<li><strong>' + esc(f.row.name) + '</strong> — ' + esc(f.error) + '</li>';
        }).join('') + '</ul>');
      yes.classList.add('hidden');
      no.disabled = false;
      no.textContent = 'Close';
      no.addEventListener('click', function () { loadTable(); }, { once: true });
    });
  }

  // ── Clone ──────────────────────────────────────────────────────────────

  function originOf(url) { try { return new URL(url).origin; } catch (_) { return ''; } }

  function cloneTargets() {
    const envs = Array.isArray(D.cfg.environments) ? D.cfg.environments : [];
    const here = originOf(D.envUrl);
    const current = envs.filter(function (e) { return originOf(e.url) === here; })[0] || { name: D.envName, url: D.envUrl };
    return [current].concat(envs.filter(function (e) { return originOf(e.url) !== here; }));
  }

  function openCloneDialog(meta, rows) {
    const targets = cloneTargets();
    const n = rows.length;
    const dlg = showDialog(
      '<div class="sc-dialog-title">Clone ' + (n === 1 ? 'record' : n + ' records') + '</div>' +
      '<div class="sc-dialog-body">' +
        '<ul class="sc-results" id="sc-clone-results">' + rows.map(function (r) {
          return '<li data-id="' + esc(r.id) + '"><span class="sc-res-icon">•</span><span class="sc-res-name">' + esc(r.name) + '</span><span class="sc-res-msg"></span></li>';
        }).join('') + '</ul>' +
        '<p><label for="sc-clone-target">Target environment</label>' +
          '<select id="sc-clone-target">' + targets.map(function (t, i) {
            return '<option value="' + esc(t.url) + '">' + esc(t.name) + (i === 0 ? ' (this environment)' : '') + '</option>';
          }).join('') + '</select></p>' +
        '<div id="sc-clone-status" class="sc-status"></div>' +
      '</div>' +
      '<div class="sc-dialog-actions">' +
        '<button type="button" class="btn" data-dlg="close">Close</button>' +
        '<button type="button" class="btn hidden" data-dlg="connect">Connect Target Environment</button>' +
        '<button type="button" class="btn btn--primary" data-dlg="clone">Clone' + (n === 1 ? '' : ' ' + n) + '</button>' +
      '</div>',
      n > 1
    );
    const sel = dlg.el.querySelector('#sc-clone-target');
    const statusEl = dlg.el.querySelector('#sc-clone-status');
    const results = dlg.el.querySelector('#sc-clone-results');
    const btnClose = dlg.el.querySelector('[data-dlg="close"]');
    const btnConnect = dlg.el.querySelector('[data-dlg="connect"]');
    const btnClone = dlg.el.querySelector('[data-dlg="clone"]');
    let busy = false;
    let pending = rows.slice(); // rows not yet cloned successfully

    function targetName() { return sel.options[sel.selectedIndex].textContent.replace(' (this environment)', ''); }
    function isCross() { return originOf(sel.value) !== originOf(D.envUrl); }
    function setResult(row, kind, html) {
      const li = results.querySelector('li[data-id="' + CSS.escape(row.id) + '"]');
      if (!li) return;
      li.className = kind ? 'sc-res--' + kind : '';
      li.querySelector('.sc-res-icon').textContent = kind === 'ok' ? '✓' : kind === 'err' ? '✕' : kind === 'run' ? '…' : '•';
      li.querySelector('.sc-res-msg').innerHTML = html || '';
    }

    async function refresh() {
      btnConnect.classList.add('hidden');
      if (!isCross()) {
        setStatus(statusEl, '', 'Creates a new copy of ' + (n === 1 ? 'this record' : 'each record') + ' in this environment.');
        btnClone.disabled = !pending.length;
        return;
      }
      btnClone.disabled = true;
      setStatus(statusEl, '', 'Checking connection…');
      let ready = false;
      try { ready = (await D.launcherCall('target-status', { targetOrigin: sel.value })).ready; } catch (_) { ready = false; }
      if (ready) {
        setStatus(statusEl, 'ok', 'Connected to ' + esc(targetName()) + '. Copies ' + (n === 1 ? 'this record' : 'each record') + ' (same ID) into that environment.');
        btnClone.disabled = !pending.length;
      } else {
        setStatus(statusEl, 'warn', 'Cloning to ' + esc(targetName()) + ' needs a live connection. Click Connect, then click the EF PPT bookmark in the new tab that opens.');
        btnConnect.classList.remove('hidden');
      }
    }

    sel.addEventListener('change', function () { if (!busy) refresh(); });
    btnClose.addEventListener('click', function () { if (!busy) dlg.close(); });

    btnConnect.addEventListener('click', async function () {
      busy = true;
      btnConnect.disabled = true; sel.disabled = true;
      btnConnect.textContent = 'Waiting for the new tab…';
      try {
        await D.launcherCall('connect-target', { targetOrigin: originOf(sel.value) }, 130000);
        busy = false;
        await refresh();
      } catch (err) {
        busy = false;
        setStatus(statusEl, 'err', esc(err.message));
      } finally {
        btnConnect.disabled = false; sel.disabled = false;
        btnConnect.textContent = 'Connect Target Environment';
      }
    });

    btnClone.addEventListener('click', async function () {
      const cross = isCross();
      const targetUrl = sel.value.replace(/\/$/, '');
      const batch = pending.slice();
      busy = true;
      btnClone.disabled = true; btnClose.disabled = true; sel.disabled = true;

      let ok = 0;
      for (let i = 0; i < batch.length; i++) {
        const row = batch[i];
        btnClone.textContent = (cross ? 'Copying ' : 'Cloning ') + (i + 1) + ' of ' + batch.length + '…';
        setResult(row, 'run', '');
        try {
          const r = await D.launcherCall('clone-record', { etn: meta.logicalName, recordId: row.id, targetUrl: targetUrl }, 300000);
          ok++;
          pending = pending.filter(function (p) { return p !== row; });
          setResult(row, 'ok', '<a class="link" href="' + esc(recordUrl(targetUrl, meta, r.newId)) + '" target="_blank" rel="noopener">Open ↗</a>');
        } catch (err) {
          setResult(row, 'err', esc(err.message));
        }
      }

      busy = false;
      btnClose.disabled = false; sel.disabled = false;
      const failed = batch.length - ok;
      if (!failed) {
        setStatus(statusEl, 'ok', (cross ? 'Copied ' : 'Cloned ') + ok + (ok === 1 ? ' record' : ' records') + (cross ? ' to ' + esc(targetName()) : '') + '.');
        btnClone.classList.add('hidden');
      } else {
        setStatus(statusEl, 'err', ok + ' succeeded, ' + failed + ' failed — see the errors above. Retry clones only the failed ones.');
        btnClone.disabled = false;
        btnClone.textContent = 'Retry ' + failed + ' failed';
      }
      if (!cross && ok) loadTable();
    });

    refresh();
  }

  // ── Dialog helpers ─────────────────────────────────────────────────────

  function showDialog(html, wide) {
    const overlay = document.createElement('div');
    overlay.className = 'sc-overlay';
    const size = wide === 'xl' ? ' sc-dialog--xl' : wide ? ' sc-dialog--wide' : '';
    overlay.innerHTML = '<div class="sc-dialog' + size + '" role="dialog" aria-modal="true">' + html + '</div>';
    document.body.appendChild(overlay);
    function onKey(e) { if (e.key === 'Escape') close(); }
    function close() { document.removeEventListener('keydown', onKey); overlay.remove(); }
    document.addEventListener('keydown', onKey);
    return { el: overlay, close: close };
  }

  function setStatus(el, kind, html) {
    el.className = 'sc-status' + (kind ? ' sc-status--' + kind : '');
    el.innerHTML = html;
  }

  // ════════════════════════════════════════════════════════════════════════
  //  Settings editor
  // ════════════════════════════════════════════════════════════════════════

  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  function draftTable() {
    return draft.tables.filter(function (t) { return t.logicalName === draftSel; })[0] || null;
  }

  function openSettings() {
    draft = clone(settings);
    draftSel = (activeTable && draft.tables.some(function (t) { return t.logicalName === activeTable; }))
      ? activeTable
      : (draft.tables[0] && draft.tables[0].logicalName) || null;
    availableFilter = '';
    document.getElementById('viewer').classList.add('hidden');
    document.getElementById('empty-view').classList.add('hidden');
    document.getElementById('settings-view').classList.remove('hidden');
    document.getElementById('btn-configure').innerHTML = '&#8592; Back to tables';
    document.getElementById('cfg-storage-note').textContent =
      (settingsSource === 'config' ? 'Currently using the defaults from your web-app config. ' : '') +
      'Saved in this browser for ' + D.envName + '.';
    renderCfgList();
    renderCfgEditor();
  }

  function isDirty() {
    return JSON.stringify(draft) !== JSON.stringify(settings);
  }

  function updateDirty() {
    document.getElementById('cfg-dirty').textContent = isDirty() ? 'Unsaved changes' : '';
  }

  function closeSettings(force) {
    if (!force && isDirty()) {
      const dlg = showDialog(
        '<div class="sc-dialog-title">Discard changes?</div>' +
        '<div class="sc-dialog-body"><p>You have unsaved configuration changes.</p></div>' +
        '<div class="sc-dialog-actions">' +
          '<button type="button" class="btn" data-dlg="stay">Keep editing</button>' +
          '<button type="button" class="btn btn--danger" data-dlg="discard">Discard</button>' +
        '</div>'
      );
      dlg.el.querySelector('[data-dlg="stay"]').addEventListener('click', dlg.close);
      dlg.el.querySelector('[data-dlg="discard"]').addEventListener('click', function () { dlg.close(); closeSettings(true); });
      return;
    }
    draft = null;
    showViewer();
  }

  function renderCfgList() {
    const ul = document.getElementById('cfg-table-list');
    if (!draft.tables.length) {
      ul.innerHTML = '<li class="sc-empty">No tables yet — add one below.</li>';
    } else {
      ul.innerHTML = draft.tables.map(function (t) {
        return (
          '<li data-table="' + esc(t.logicalName) + '"' + (t.logicalName === draftSel ? ' class="active"' : '') + '>' +
            '<div class="sc-item-text">' +
              '<div class="sc-item-name" data-label="' + esc(t.logicalName) + '">' + esc(t.logicalName) + '</div>' +
              '<div class="sc-item-sub">' + t.columns.length + ' column' + (t.columns.length === 1 ? '' : 's') + '</div>' +
            '</div>' +
            '<button type="button" class="sc-remove" data-remove="' + esc(t.logicalName) + '" title="Remove table">✕</button>' +
          '</li>'
        );
      }).join('');
      draft.tables.forEach(function (t) {
        loadMeta(t.logicalName).then(function (m) {
          const el = ul.querySelector('[data-label="' + CSS.escape(t.logicalName) + '"]');
          if (el) el.textContent = m.plural;
        }).catch(function () {});
      });
    }
    updateDirty();
  }

  async function renderCfgEditor() {
    const host = document.getElementById('cfg-editor');
    const t = draftTable();
    if (!t) {
      host.innerHTML = '<div class="state-msg">Add a table on the left to start configuring.</div>';
      updateDirty();
      return;
    }
    const sel = draftSel;
    host.innerHTML = '<div class="state-msg state-loading">Loading columns…</div>';
    let meta;
    try {
      meta = await loadMeta(t.logicalName);
    } catch (err) {
      if (draftSel === sel) host.innerHTML = '<div class="state-msg state-error">Could not load “' + esc(t.logicalName) + '”: ' + esc(err.message) + '</div>';
      return;
    }
    if (draftSel !== sel || !draft) return;

    const selectedCols = t.columns.map(function (c) { return { cfg: c, attr: meta.attrByName.get(c.name) }; });
    const sortCol = t.sort ? t.sort.column : '';
    host.innerHTML =
      '<h2>' + esc(meta.plural) + '</h2>' +
      '<div class="sc-cfg-sub">' + esc(meta.logicalName) + (meta.hasState ? '' : ' · no Active/Inactive state') + '</div>' +
      '<div class="sc-cfg-row">' +
        '<label for="cfg-sort-col">Default sort</label>' +
        '<select id="cfg-sort-col"><option value="">First column</option>' +
          selectedCols.filter(function (x) { return x.attr; }).map(function (x) {
            return '<option value="' + esc(x.attr.name) + '"' + (x.attr.name === sortCol ? ' selected' : '') + '>' + esc(x.attr.label) + '</option>';
          }).join('') +
        '</select>' +
        '<select id="cfg-sort-dir">' +
          '<option value="asc"' + (!t.sort || t.sort.dir === 'asc' ? ' selected' : '') + '>Ascending</option>' +
          '<option value="desc"' + (t.sort && t.sort.dir === 'desc' ? ' selected' : '') + '>Descending</option>' +
        '</select>' +
      '</div>' +
      '<div class="sc-cols">' +
        '<div class="sc-panel">' +
          '<div class="sc-panel-head">Selected columns (' + selectedCols.length + ') <small>top to bottom = left to right in the grid</small></div>' +
          (selectedCols.length ? '<ol class="sc-selected">' + selectedCols.map(function (x, i) {
            const a = x.attr;
            const canEdit = a && a.canEdit;
            return (
              '<li data-col="' + esc(x.cfg.name) + '">' +
                '<div class="sc-move">' +
                  '<button type="button" data-move="-1" title="Move up"' + (i === 0 ? ' disabled' : '') + '>▲</button>' +
                  '<button type="button" data-move="1" title="Move down"' + (i === selectedCols.length - 1 ? ' disabled' : '') + '>▼</button>' +
                '</div>' +
                '<div class="sc-col-text">' +
                  '<div class="sc-col-name">' + esc(a ? a.label : x.cfg.name) + '</div>' +
                  '<div class="sc-col-sub">' + esc(x.cfg.name) + ' · ' + esc(a ? a.type : 'not found in this environment') + '</div>' +
                '</div>' +
                '<label class="sc-editable' + (canEdit ? '' : ' disabled') + '" title="' + esc(canEdit ? 'Allow editing this column in the grid' : (a ? a.editReason : '')) + '">' +
                  '<input type="checkbox" data-editable' + (canEdit && x.cfg.editable ? ' checked' : '') + (canEdit ? '' : ' disabled') + '> Editable' +
                '</label>' +
                (a && (a.type === 'String' || a.type === 'Memo')
                  ? '<label class="sc-editable" title="Values may contain JSON — show a formatted JSON viewer' +
                      (canEdit ? ' (and editor, when Editable is ticked)' : '') + '">' +
                      '<input type="checkbox" data-json' + (x.cfg.json ? ' checked' : '') + '> JSON' +
                    '</label>'
                  : '<span class="sc-editable disabled" title="Only text columns can hold JSON"><input type="checkbox" disabled> JSON</span>') +
                '<button type="button" class="sc-remove" data-remove-col title="Remove column">✕</button>' +
              '</li>'
            );
          }).join('') + '</ol>' : '<div class="sc-panel-empty">No columns selected yet — add some from the list on the right.</div>') +
        '</div>' +
        '<div class="sc-panel">' +
          '<div class="sc-panel-head">Add columns</div>' +
          '<div class="sc-panel-search"><input type="search" id="cfg-col-search" placeholder="Search columns…" autocomplete="off" value="' + esc(availableFilter) + '"></div>' +
          '<ul class="sc-available" id="cfg-available"></ul>' +
        '</div>' +
      '</div>';

    renderAvailable(meta, t);
    updateDirty();
  }

  function renderAvailable(meta, t) {
    const ul = document.getElementById('cfg-available');
    if (!ul) return;
    const chosen = new Set(t.columns.map(function (c) { return c.name; }));
    const q = availableFilter.trim().toLowerCase();
    const list = meta.attrs.filter(function (a) {
      if (chosen.has(a.name)) return false;
      return !q || a.label.toLowerCase().indexOf(q) !== -1 || a.name.indexOf(q) !== -1;
    });
    ul.innerHTML = list.length ? list.map(function (a) {
      return (
        '<li>' +
          '<div class="sc-col-text">' +
            '<div class="sc-col-name">' + esc(a.label) + '</div>' +
            '<div class="sc-col-sub">' + esc(a.name) + ' · ' + esc(a.type) + (a.isSystem ? ' · system' : '') + '</div>' +
          '</div>' +
          '<button type="button" class="sc-add" data-add-col="' + esc(a.name) + '">+ Add</button>' +
        '</li>'
      );
    }).join('') : '<li class="sc-panel-empty">' + (q ? 'No matching columns.' : 'All columns are selected.') + '</li>';
  }

  function wireSettings() {
    document.getElementById('cfg-table-list').addEventListener('click', function (e) {
      const rm = e.target.closest('[data-remove]');
      if (rm) {
        const ln = rm.dataset.remove;
        draft.tables = draft.tables.filter(function (t) { return t.logicalName !== ln; });
        if (draftSel === ln) draftSel = draft.tables[0] ? draft.tables[0].logicalName : null;
        renderCfgList();
        renderCfgEditor();
        return;
      }
      const li = e.target.closest('li[data-table]');
      if (!li || li.dataset.table === draftSel) return;
      draftSel = li.dataset.table;
      availableFilter = '';
      renderCfgList();
      renderCfgEditor();
    });

    const editor = document.getElementById('cfg-editor');
    editor.addEventListener('click', async function (e) {
      const t = draftTable();
      if (!t) return;
      const add = e.target.closest('[data-add-col]');
      if (add) {
        t.columns.push({ name: add.dataset.addCol, editable: false, json: false });
        await renderCfgEditor();
        renderCfgList();
        return;
      }
      const li = e.target.closest('li[data-col]');
      if (!li) return;
      const idx = t.columns.findIndex(function (c) { return c.name === li.dataset.col; });
      if (idx === -1) return;
      const move = e.target.closest('[data-move]');
      if (move && !move.disabled) {
        const to = idx + parseInt(move.dataset.move, 10);
        const item = t.columns.splice(idx, 1)[0];
        t.columns.splice(to, 0, item);
        await renderCfgEditor();
        return;
      }
      if (e.target.closest('[data-remove-col]')) {
        const removed = t.columns.splice(idx, 1)[0];
        if (t.sort && t.sort.column === removed.name) t.sort = null;
        await renderCfgEditor();
        renderCfgList();
      }
    });
    editor.addEventListener('change', function (e) {
      const t = draftTable();
      if (!t) return;
      if (e.target.matches('[data-editable]')) {
        const li = e.target.closest('li[data-col]');
        const c = t.columns.filter(function (x) { return x.name === li.dataset.col; })[0];
        if (c) c.editable = e.target.checked;
      } else if (e.target.matches('[data-json]')) {
        const li = e.target.closest('li[data-col]');
        const c = t.columns.filter(function (x) { return x.name === li.dataset.col; })[0];
        if (c) c.json = e.target.checked;
      } else if (e.target.id === 'cfg-sort-col' || e.target.id === 'cfg-sort-dir') {
        const col = document.getElementById('cfg-sort-col').value;
        const dir = document.getElementById('cfg-sort-dir').value;
        t.sort = col ? { column: col, dir: dir } : null;
      }
      updateDirty();
    });
    editor.addEventListener('input', function (e) {
      if (e.target.id !== 'cfg-col-search') return;
      availableFilter = e.target.value;
      const t = draftTable();
      if (t) loadMeta(t.logicalName).then(function (m) { renderAvailable(m, t); });
    });

    document.getElementById('btn-add-table').addEventListener('click', openTablePicker);
    document.getElementById('btn-cfg-cancel').addEventListener('click', function () { closeSettings(false); });
    document.getElementById('btn-cfg-save').addEventListener('click', async function () {
      const btn = this;
      btn.disabled = true;
      btn.textContent = 'Saving…';
      try {
        await saveSettings(normalizeSettings(draft));
        metaCache.forEach(function (_, k) { if (!settings.tables.some(function (t) { return t.logicalName === k; })) metaCache.delete(k); });
        closeSettings(true);
      } catch (err) {
        const dlg = showDialog(
          '<div class="sc-dialog-title">Couldn’t save</div>' +
          '<div class="sc-dialog-body"><p>' + esc(err.message) + '</p></div>' +
          '<div class="sc-dialog-actions"><button type="button" class="btn btn--primary" data-dlg="ok">OK</button></div>'
        );
        dlg.el.querySelector('[data-dlg="ok"]').addEventListener('click', dlg.close);
      } finally {
        btn.disabled = false;
        btn.textContent = 'Save';
      }
    });
    document.getElementById('btn-io').addEventListener('click', openImportExport);
  }

  function openTablePicker() {
    const dlg = showDialog(
      '<div class="sc-dialog-title">Add table</div>' +
      '<div class="sc-dialog-body">' +
        '<input type="search" id="sc-pick-search" placeholder="Search tables by name…" autocomplete="off">' +
        '<ul class="sc-picker-list" id="sc-pick-list"><li class="sc-empty">Loading tables…</li></ul>' +
      '</div>' +
      '<div class="sc-dialog-actions"><button type="button" class="btn" data-dlg="close">Close</button></div>',
      true
    );
    const input = dlg.el.querySelector('#sc-pick-search');
    const list = dlg.el.querySelector('#sc-pick-list');
    dlg.el.querySelector('[data-dlg="close"]').addEventListener('click', dlg.close);
    input.focus();

    let entities = [];
    function render() {
      const q = input.value.trim().toLowerCase();
      const configured = new Set(draft.tables.map(function (t) { return t.logicalName; }));
      const matches = entities.filter(function (e) {
        return !q || e.label.toLowerCase().indexOf(q) !== -1 || e.name.indexOf(q) !== -1;
      }).slice(0, 200);
      list.innerHTML = matches.length ? matches.map(function (e) {
        return (
          '<li data-pick="' + esc(e.name) + '"' + (configured.has(e.name) ? ' class="configured"' : '') + '>' +
            '<div class="sc-col-name">' + esc(e.label) + (configured.has(e.name) ? ' · already added' : '') + '</div>' +
            '<div class="sc-col-sub">' + esc(e.name) + '</div>' +
          '</li>'
        );
      }).join('') : '<li class="sc-empty">No matching tables.</li>';
    }
    input.addEventListener('input', render);
    list.addEventListener('click', async function (e) {
      const li = e.target.closest('li[data-pick]');
      if (!li) return;
      const ln = li.dataset.pick;
      if (!draft.tables.some(function (t) { return t.logicalName === ln; })) {
        const t = { logicalName: ln, columns: [], sort: null };
        try {
          const meta = await loadMeta(ln);
          if (meta.primaryName && meta.attrByName.has(meta.primaryName)) t.columns.push({ name: meta.primaryName, editable: false, json: false });
        } catch (_) { /* editor shows the error */ }
        draft.tables.push(t);
      }
      draftSel = ln;
      availableFilter = '';
      dlg.close();
      renderCfgList();
      renderCfgEditor();
    });

    loadEntityList().then(function (list_) { entities = list_; render(); }, function (err) {
      list.innerHTML = '<li class="sc-empty">Could not load tables: ' + esc(err.message) + '</li>';
    });
  }

  function openImportExport() {
    const dlg = showDialog(
      '<div class="sc-dialog-title">Import / Export settings</div>' +
      '<div class="sc-dialog-body">' +
        '<p>Copy this JSON to reuse the same configuration in another environment, or paste settings here and apply them. ' +
          'To make it the default everywhere, add it to your web-app config as <code>settings.systemConfigurator</code>.</p>' +
        '<textarea id="sc-io-text" spellcheck="false"></textarea>' +
        '<div id="sc-io-status" class="sc-status hidden"></div>' +
      '</div>' +
      '<div class="sc-dialog-actions">' +
        '<button type="button" class="btn" data-dlg="close">Close</button>' +
        '<button type="button" class="btn" data-dlg="copy">Copy</button>' +
        '<button type="button" class="btn btn--primary" data-dlg="apply">Apply to editor</button>' +
      '</div>',
      true
    );
    const ta = dlg.el.querySelector('#sc-io-text');
    const statusEl = dlg.el.querySelector('#sc-io-status');
    ta.value = JSON.stringify(normalizeSettings(draft), null, 2);
    dlg.el.querySelector('[data-dlg="close"]').addEventListener('click', dlg.close);
    dlg.el.querySelector('[data-dlg="copy"]').addEventListener('click', function () {
      ta.select();
      Promise.resolve(navigator.clipboard && navigator.clipboard.writeText(ta.value))
        .then(function () { setStatus(statusEl, 'ok', 'Copied to the clipboard.'); })
        .catch(function () { setStatus(statusEl, '', 'Select the text and press Ctrl+C to copy.'); });
    });
    dlg.el.querySelector('[data-dlg="apply"]').addEventListener('click', function () {
      let parsed;
      try { parsed = JSON.parse(ta.value); } catch (e) { setStatus(statusEl, 'err', 'Invalid JSON: ' + esc(e.message)); return; }
      if (!parsed || !Array.isArray(parsed.tables)) { setStatus(statusEl, 'err', 'Expected an object with a "tables" array.'); return; }
      draft = normalizeSettings(parsed);
      draftSel = draft.tables[0] ? draft.tables[0].logicalName : null;
      dlg.close();
      renderCfgList();
      renderCfgEditor();
    });
  }
})();
