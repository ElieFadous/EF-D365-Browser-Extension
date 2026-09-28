/**
 * EF Power Platform Tools — DataGrid: sortable columns, per-column filters,
 * global search, inline cell editing and resizable columns. Exposes
 * window.DataGrid.
 *
 * Column spec:
 *   key, label, width        — identity / header text / optional CSS width
 *   value(row)               — raw value for sort, filter and search (default row[key])
 *   filterValues(row)        — values a select filter matches against (default [value])
 *   sortValue(row)           — override sort key
 *   render(row)              — cell HTML (default: escaped value, "—" when blank)
 *   filter                   — 'text' | 'select' | 'value' (Any/Blank/Not blank + contains) | null
 *   filterOptions            — [{ value, label }] for a select filter (required in server mode)
 *   filterNoText             — 'value' filter without the contains box (blank/not blank only)
 *   filterPlaceholder        — placeholder for the filter's text box
 *   sortable                 — default true
 *   editable, edit(row, v)   — inline edit; edit() returns a Promise and mutates row
 *   editValue(row), editHint — initial editor text / hint shown under the editor
 *   editOptions(row)         — [{ value, label }] to edit with a dropdown instead of text
 *
 * Options: columns, rowKey, defaultSort, onCountChange, tableMinWidth (CSS length),
 *   selectable + onSelectionChange(rows) — checkbox column; click selects a row,
 *   Ctrl/Cmd-click toggles one, Shift-click selects a range, header box selects all;
 *   serverMode + onQueryChange({ search, filters, sort }) — rows are shown exactly
 *   as given and every search/filter/sort change is handed to the caller to
 *   re-query the server instead of being applied to the loaded rows.
 */
(function () {
  'use strict';

  const MIN_COL_WIDTH = 60;

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function isBlank(v) { return v == null || String(v).trim() === ''; }
  function norm(v) { return v == null ? '' : String(v); }
  function compare(a, b) {
    if (typeof a === 'number' && typeof b === 'number') return a - b;
    return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
  }

  function DataGrid(host, opts) {
    this.host = host;
    this.cols = opts.columns;
    this.rowKey = opts.rowKey;
    this.onCount = opts.onCountChange || function () {};
    this.serverMode = !!opts.serverMode;
    this.onQueryChange = opts.onQueryChange || function () {};
    this.rows = [];
    this.byId = new Map();
    this.sort = opts.defaultSort || null;
    this.filters = {};
    this.search = '';
    this.editing = null; // { id, key, draft, error, saving }
    this.tableMinWidth = opts.tableMinWidth || null;
    this.widthsFrozen = false;
    this.selectable = !!opts.selectable;
    this.onSelectionChange = opts.onSelectionChange || function () {};
    this.selected = new Set(); // row ids
    this.anchorId = null;      // for shift-click ranges
    this._build();
  }

  DataGrid.prototype._col = function (key) {
    return this.cols.filter(function (c) { return c.key === key; })[0];
  };
  DataGrid.prototype._val = function (col, row) {
    return col.value ? col.value(row) : row[col.key];
  };
  DataGrid.prototype._filterVals = function (col, row) {
    return col.filterValues ? col.filterValues(row) : [this._val(col, row)];
  };

  DataGrid.prototype.query = function () {
    return { search: this.search, filters: JSON.parse(JSON.stringify(this.filters)), sort: this.sort };
  };

  DataGrid.prototype._emitQuery = function () {
    this.onQueryChange(this.query());
  };

  DataGrid.prototype._filterHtml = function (c) {
    const k = esc(c.key);
    const ph = esc(c.filterPlaceholder || (c.filter === 'value' ? 'Contains…' : 'Filter…'));
    if (c.filter === 'text') {
      return '<input type="search" class="dg-f-text" data-key="' + k + '" placeholder="' + ph + '" autocomplete="off">';
    }
    if (c.filter === 'select') {
      return '<select class="dg-f-select" data-key="' + k + '"><option value="">All</option></select>';
    }
    if (c.filter === 'value') {
      return (
        '<div class="dg-f-value' + (c.filterNoText ? ' dg-f-value--mode-only' : '') + '">' +
          '<select class="dg-f-mode" data-key="' + k + '">' +
            '<option value="">Any</option><option value="blank">Blank</option><option value="notblank">Not blank</option>' +
          '</select>' +
          (c.filterNoText ? '' : '<input type="search" class="dg-f-text" data-key="' + k + '" placeholder="' + ph + '" autocomplete="off">') +
        '</div>'
      );
    }
    return '';
  };

  DataGrid.prototype._build = function () {
    const self = this;
    const offset = this.selectable ? 1 : 0; // the checkbox column sits before the data columns
    this.host.classList.add('dg');
    this.host.innerHTML =
      '<div class="dg-scroll"><table class="dg-table"' +
        (this.tableMinWidth ? ' style="min-width:' + esc(this.tableMinWidth) + '"' : '') + '>' +
        '<colgroup>' + (this.selectable ? '<col style="width:40px">' : '') + this.cols.map(function (c) {
          return '<col' + (c.width ? ' style="width:' + c.width + '"' : '') + '>';
        }).join('') + '</colgroup>' +
        '<thead>' +
          '<tr class="dg-head">' +
            (this.selectable ? '<th class="dg-sel"><input type="checkbox" class="dg-all-cb" title="Select all loaded rows" aria-label="Select all"></th>' : '') +
            this.cols.map(function (c, i) {
            return '<th>' + (c.sortable === false
              ? '<span class="dg-th-label">' + esc(c.label) + '</span>'
              : '<button type="button" class="dg-sort" data-key="' + esc(c.key) + '">' +
                  '<span>' + esc(c.label) + '</span><span class="dg-arrow"></span></button>') +
              '<span class="dg-resize" data-index="' + (i + offset) + '" title="Drag to resize"></span></th>';
          }).join('') + '</tr>' +
          '<tr class="dg-filters">' + (this.selectable ? '<th class="dg-sel"></th>' : '') + this.cols.map(function (c) {
            return '<th>' + self._filterHtml(c) + '</th>';
          }).join('') + '</tr>' +
        '</thead>' +
        '<tbody></tbody>' +
      '</table></div>';
    this.table = this.host.querySelector('table');
    this.colEls = Array.from(this.host.querySelectorAll('col'));
    this.tbody = this.host.querySelector('tbody');

    const thead = this.host.querySelector('thead');
    thead.addEventListener('click', function (e) {
      const b = e.target.closest('.dg-sort');
      if (!b) return;
      const key = b.dataset.key;
      self.sort = (self.sort && self.sort.key === key)
        ? { key: key, dir: self.sort.dir === 'asc' ? 'desc' : 'asc' }
        : { key: key, dir: 'asc' };
      self._renderHead();
      if (self.serverMode) self._emitQuery();
      else self._renderBody();
    });
    thead.addEventListener('mousedown', function (e) {
      const handle = e.target.closest('.dg-resize');
      if (handle) self._startResize(e, handle);
    });

    this.host.addEventListener('input', function (e) {
      const t = e.target;
      if (t.classList.contains('dg-f-text')) self._setFilter(t.dataset.key, 'text', t.value);
    });
    this.host.addEventListener('change', function (e) {
      const t = e.target;
      if (t.classList.contains('dg-f-select')) self._setFilter(t.dataset.key, 'value', t.value);
      else if (t.classList.contains('dg-f-mode')) self._setFilter(t.dataset.key, 'mode', t.value);
    });

    this.tbody.addEventListener('click', function (e) {
      const btn = e.target.closest('[data-dg-act]');
      if (!btn) return;
      const act = btn.dataset.dgAct;
      if (act === 'edit') self._startEdit(btn.closest('tr').dataset.id, btn.dataset.key);
      else if (act === 'save') self._saveEdit();
      else if (act === 'cancel') self._cancelEdit();
    });
    if (this.selectable) {
      this.host.querySelector('.dg-all-cb').addEventListener('change', function (e) {
        self.selected = e.target.checked
          ? new Set(self._visibleRows().map(function (r) { return String(self.rowKey(r)); }))
          : new Set();
        self.anchorId = null;
        self._syncSelection();
      });
      this.tbody.addEventListener('click', function (e) {
        const tr = e.target.closest('tr[data-id]');
        if (!tr) return;
        const id = tr.dataset.id;
        if (e.target.classList.contains('dg-row-cb')) {
          if (e.shiftKey && self.anchorId) self._selectRange(self.anchorId, id, true);
          else { self._toggle(id); self.anchorId = id; }
          self._syncSelection();
          return;
        }
        // Leave clicks on controls and open editors alone.
        if (e.target.closest('a, button, input, select, textarea, label, .dg-editing')) return;
        if (e.shiftKey && self.anchorId) {
          self._selectRange(self.anchorId, id, !(e.ctrlKey || e.metaKey));
        } else if (e.ctrlKey || e.metaKey) {
          self._toggle(id);
          self.anchorId = id;
        } else {
          self.selected = new Set([id]);
          self.anchorId = id;
        }
        self._syncSelection();
      });
    }
    this.tbody.addEventListener('dblclick', function (e) {
      const td = e.target.closest('td.dg-editable');
      if (!td || td.classList.contains('dg-editing')) return;
      self._startEdit(td.parentNode.dataset.id, td.dataset.key);
    });
    this.tbody.addEventListener('keydown', function (e) {
      if (!e.target.classList.contains('dg-editor')) return;
      if (e.key === 'Escape') { e.preventDefault(); self._cancelEdit(); }
      else if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); self._saveEdit(); }
    });

    if (this.serverMode) this._refreshOptions();
  };

  // ── Row selection ─────────────────────────────────────────────────────

  DataGrid.prototype._toggle = function (id) {
    if (this.selected.has(id)) this.selected.delete(id);
    else this.selected.add(id);
  };

  /** Selects every visible row between two ids (inclusive); `replace` drops the rest of the selection. */
  DataGrid.prototype._selectRange = function (fromId, toId, replace) {
    const self = this;
    const ids = this._visibleRows().map(function (r) { return String(self.rowKey(r)); });
    const a = ids.indexOf(fromId), b = ids.indexOf(toId);
    if (a === -1 || b === -1) { this.selected = new Set([toId]); this.anchorId = toId; return; }
    if (replace) this.selected = new Set();
    ids.slice(Math.min(a, b), Math.max(a, b) + 1).forEach(function (id) { self.selected.add(id); });
  };

  /** Reflects this.selected in the DOM without re-rendering rows, then notifies the caller. */
  DataGrid.prototype._syncSelection = function () {
    const self = this;
    this.tbody.querySelectorAll('tr[data-id]').forEach(function (tr) {
      const on = self.selected.has(tr.dataset.id);
      tr.classList.toggle('dg-selected', on);
      const cb = tr.querySelector('.dg-row-cb');
      if (cb) cb.checked = on;
    });
    const all = this.host.querySelector('.dg-all-cb');
    if (all) {
      const visible = this._visibleRows().length;
      const n = this.selected.size;
      all.checked = visible > 0 && n >= visible;
      all.indeterminate = n > 0 && n < visible;
    }
    this.onSelectionChange(this.getSelectedRows());
  };

  DataGrid.prototype.getSelectedRows = function () {
    const self = this;
    return this.rows.filter(function (r) { return self.selected.has(String(self.rowKey(r))); });
  };

  DataGrid.prototype.clearSelection = function () {
    this.selected = new Set();
    this.anchorId = null;
    this._syncSelection();
  };

  // ── Column resizing ───────────────────────────────────────────────────

  /**
   * Switches every column from its initial (often %) width to its current
   * pixel width, and the table to the sum of them — so dragging one column
   * edge changes only that column instead of redistributing the others.
   */
  DataGrid.prototype._freezeWidths = function () {
    if (this.widthsFrozen) return;
    const ths = this.host.querySelectorAll('.dg-head th');
    const self = this;
    ths.forEach(function (th, i) {
      self.colEls[i].style.width = Math.max(MIN_COL_WIDTH, Math.round(th.getBoundingClientRect().width)) + 'px';
    });
    this.table.style.minWidth = '0';
    this._syncTableWidth();
    this.widthsFrozen = true;
  };

  DataGrid.prototype._syncTableWidth = function () {
    const total = this.colEls.reduce(function (sum, c) { return sum + (parseFloat(c.style.width) || 0); }, 0);
    this.table.style.width = total + 'px';
  };

  DataGrid.prototype._startResize = function (e, handle) {
    e.preventDefault();
    e.stopPropagation();
    const self = this;
    this._freezeWidths();
    const i = parseInt(handle.dataset.index, 10);
    const colEl = this.colEls[i];
    const startX = e.clientX;
    const startW = parseFloat(colEl.style.width) || MIN_COL_WIDTH;
    handle.classList.add('active');
    document.body.classList.add('dg-resizing');

    function onMove(ev) {
      colEl.style.width = Math.max(MIN_COL_WIDTH, Math.round(startW + ev.clientX - startX)) + 'px';
      self._syncTableWidth();
    }
    function onUp() {
      handle.classList.remove('active');
      document.body.classList.remove('dg-resizing');
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    }
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  };

  // ── Filters / search ──────────────────────────────────────────────────

  DataGrid.prototype._setFilter = function (key, prop, value) {
    const f = this.filters[key] || (this.filters[key] = {});
    f[prop] = value;
    if (this.serverMode) this._emitQuery();
    else this._renderBody();
  };

  DataGrid.prototype.setSearch = function (q) {
    this.search = (q || '').trim();
    if (this.serverMode) this._emitQuery();
    else this._renderBody();
  };

  DataGrid.prototype.clearFilters = function () {
    this.filters = {};
    this.search = '';
    this.host.querySelectorAll('.dg-f-text').forEach(function (el) { el.value = ''; });
    this.host.querySelectorAll('.dg-f-select, .dg-f-mode').forEach(function (el) { el.value = ''; });
    if (this.serverMode) this._emitQuery();
    else this._renderBody();
  };

  DataGrid.prototype.setRows = function (rows) {
    const self = this;
    this.rows = rows;
    this.byId = new Map();
    rows.forEach(function (r) { self.byId.set(String(self.rowKey(r)), r); });
    this.editing = null;
    // Keep the selection for rows that are still loaded; drop the rest.
    this.selected = new Set(Array.from(this.selected).filter(function (id) { return self.byId.has(id); }));
    if (this.anchorId && !this.byId.has(this.anchorId)) this.anchorId = null;
    if (!this.serverMode) this._refreshOptions();
    this._renderHead();
    this._renderBody();
    if (this.selectable) this._syncSelection();
  };

  DataGrid.prototype._refreshOptions = function () {
    const self = this;
    this.cols.filter(function (c) { return c.filter === 'select'; }).forEach(function (c) {
      const sel = self.host.querySelector('.dg-f-select[data-key="' + c.key + '"]');
      if (!sel) return;
      const current = sel.value;
      let opts;
      let hasBlank;
      if (self.serverMode || c.filterOptions) {
        opts = (c.filterOptions || []).map(function (o) { return { value: String(o.value), label: o.label }; });
        hasBlank = true;
      } else {
        const values = new Set();
        hasBlank = false;
        self.rows.forEach(function (r) {
          const vals = self._filterVals(c, r);
          if (!vals.length || vals.every(isBlank)) hasBlank = true;
          vals.forEach(function (v) { if (!isBlank(v)) values.add(String(v)); });
        });
        opts = Array.from(values).sort(compare).map(function (v) { return { value: v, label: v }; });
      }
      sel.innerHTML =
        '<option value="">All</option>' +
        (hasBlank ? '<option value="__blank__">(blank)</option>' : '') +
        opts.map(function (o) { return '<option value="' + esc(o.value) + '">' + esc(o.label) + '</option>'; }).join('');
      const keep = (current === '__blank__' && hasBlank) || opts.some(function (o) { return o.value === current; });
      sel.value = keep ? current : '';
      if (sel.value !== current) {
        const f = self.filters[c.key] || (self.filters[c.key] = {});
        f.value = sel.value;
      }
    });
  };

  DataGrid.prototype._passes = function (row) {
    const self = this;
    if (this.search) {
      const q = this.search.toLowerCase();
      const hit = this.cols.some(function (c) {
        if (c.searchable === false) return false;
        return [self._val(c, row)].concat(self._filterVals(c, row)).some(function (v) {
          return norm(v).toLowerCase().indexOf(q) !== -1;
        });
      });
      if (!hit) return false;
    }
    return this.cols.every(function (c) {
      const f = self.filters[c.key];
      if (!f || !c.filter) return true;
      if (c.filter === 'text') {
        if (!f.text) return true;
        const q = f.text.trim().toLowerCase();
        return self._filterVals(c, row).some(function (v) { return norm(v).toLowerCase().indexOf(q) !== -1; });
      }
      if (c.filter === 'select') {
        if (!f.value) return true;
        const vals = self._filterVals(c, row);
        if (f.value === '__blank__') return !vals.length || vals.every(isBlank);
        return vals.some(function (v) { return String(v) === f.value; });
      }
      if (c.filter === 'value') {
        const v = self._val(c, row);
        if (f.mode === 'blank' && !isBlank(v)) return false;
        if (f.mode === 'notblank' && isBlank(v)) return false;
        if (f.text && norm(v).toLowerCase().indexOf(f.text.trim().toLowerCase()) === -1) return false;
      }
      return true;
    });
  };

  DataGrid.prototype._visibleRows = function () {
    const self = this;
    if (this.serverMode) return this.rows;
    const rows = this.rows.filter(function (r) { return self._passes(r); });
    if (!this.sort) return rows;
    const col = this._col(this.sort.key);
    if (!col) return rows;
    const dir = this.sort.dir === 'desc' ? -1 : 1;
    const key = function (r) { return col.sortValue ? col.sortValue(r) : self._val(col, r); };
    return rows.sort(function (r1, r2) {
      const a = key(r1), b = key(r2);
      const ab = isBlank(a), bb = isBlank(b);
      if (ab !== bb) return ab ? 1 : -1; // blanks always last, either direction
      if (ab) return 0;
      return compare(a, b) * dir;
    });
  };

  DataGrid.prototype._renderHead = function () {
    const sort = this.sort;
    this.host.querySelectorAll('.dg-sort').forEach(function (b) {
      const active = sort && sort.key === b.dataset.key;
      b.classList.toggle('active', !!active);
      b.querySelector('.dg-arrow').textContent = active ? (sort.dir === 'asc' ? '▲' : '▼') : '';
      b.closest('th').setAttribute('aria-sort', active ? (sort.dir === 'asc' ? 'ascending' : 'descending') : 'none');
    });
  };

  // ── Rendering ─────────────────────────────────────────────────────────

  DataGrid.prototype._defaultRender = function (col, row) {
    const v = this._val(col, row);
    if (isBlank(v)) return '<span class="dg-blank">—</span>';
    return '<span class="dg-text" title="' + esc(v) + '">' + esc(v) + '</span>';
  };

  DataGrid.prototype._cellHtml = function (row, col) {
    const id = String(this.rowKey(row));
    const cls = [col.className || '', col.editable ? 'dg-editable' : ''].join(' ').trim();
    const ed = this.editing;

    if (ed && ed.id === id && ed.key === col.key) {
      const lines = ed.draft.split('\n').length;
      const options = col.editOptions ? col.editOptions(row) : null;
      const editor = options
        ? '<select class="dg-editor"' + (ed.saving ? ' disabled' : '') + '>' +
            options.map(function (o) {
              return '<option value="' + esc(o.value) + '"' + (String(o.value) === ed.draft ? ' selected' : '') + '>' +
                esc(o.label) + '</option>';
            }).join('') + '</select>'
        : '<textarea class="dg-editor" rows="' + Math.min(8, Math.max(2, lines)) + '"' +
            (ed.saving ? ' disabled' : '') + ' spellcheck="false">' + esc(ed.draft) + '</textarea>';
      return (
        '<td class="' + cls + ' dg-editing" data-key="' + esc(col.key) + '">' +
          editor +
          (ed.error ? '<div class="dg-error">' + esc(ed.error) + '</div>' : '') +
          (col.editHint ? '<div class="dg-hint">' + esc(col.editHint) + '</div>' : '') +
          '<div class="dg-edit-actions">' +
            '<button type="button" class="dg-btn dg-btn--primary" data-dg-act="save"' + (ed.saving ? ' disabled' : '') + '>' +
              (ed.saving ? 'Saving…' : 'Save') + '</button>' +
            '<button type="button" class="dg-btn" data-dg-act="cancel"' + (ed.saving ? ' disabled' : '') + '>Cancel</button>' +
          '</div>' +
        '</td>'
      );
    }

    const content = col.render ? col.render(row) : this._defaultRender(col, row);
    const editBtn = col.editable
      ? '<button type="button" class="dg-edit-btn" data-dg-act="edit" data-key="' + esc(col.key) + '" title="Edit">✎</button>'
      : '';
    return (
      '<td class="' + cls + '" data-key="' + esc(col.key) + '">' +
        '<div class="dg-cell"><div class="dg-cell-content">' + content + '</div>' + editBtn + '</div>' +
      '</td>'
    );
  };

  DataGrid.prototype._renderBody = function () {
    const self = this;
    // Keep whatever's been typed into an open editor across re-renders.
    const open = this.tbody.querySelector('.dg-editor');
    if (open && this.editing && !this.editing.saving) this.editing.draft = open.value;

    const rows = this._visibleRows();
    this.onCount(rows.length, this.rows.length);
    if (!rows.length) {
      const filtered = this.serverMode ? !!(this.search || Object.keys(this.filters).length) : this.rows.length;
      this.tbody.innerHTML = '<tr><td class="dg-empty" colspan="' + (this.cols.length + (this.selectable ? 1 : 0)) + '">' +
        (filtered ? 'No rows match the current search or filters.' : 'No records found.') + '</td></tr>';
      return;
    }
    this.tbody.innerHTML = rows.map(function (r) {
      const id = String(self.rowKey(r));
      const on = self.selectable && self.selected.has(id);
      return '<tr data-id="' + esc(id) + '"' + (on ? ' class="dg-selected"' : '') + '>' +
        (self.selectable
          ? '<td class="dg-sel"><input type="checkbox" class="dg-row-cb" aria-label="Select row"' + (on ? ' checked' : '') + '></td>'
          : '') +
        self.cols.map(function (c) { return self._cellHtml(r, c); }).join('') + '</tr>';
    }).join('');
  };

  /** Re-renders rows without re-filtering or re-sorting (e.g. after a row was updated elsewhere). */
  DataGrid.prototype.refresh = function () {
    this._renderBody();
  };

  // ── Inline editing ────────────────────────────────────────────────────

  DataGrid.prototype._focusEditor = function () {
    const ta = this.tbody.querySelector('.dg-editor');
    if (!ta || ta.disabled) return;
    ta.focus();
    if (ta.tagName === 'TEXTAREA') ta.setSelectionRange(ta.value.length, ta.value.length);
  };

  DataGrid.prototype._startEdit = function (id, key) {
    const row = this.byId.get(String(id));
    const col = this._col(key);
    if (!row || !col || !col.editable) return;
    if (this.editing && this.editing.saving) return;
    this.editing = {
      id: String(id), key: key, error: null, saving: false,
      draft: col.editValue ? norm(col.editValue(row)) : norm(this._val(col, row)),
    };
    this._renderBody();
    this._focusEditor();
  };

  DataGrid.prototype._cancelEdit = function () {
    if (!this.editing || this.editing.saving) return;
    this.editing = null;
    this._renderBody();
  };

  DataGrid.prototype._saveEdit = function () {
    const self = this;
    const ed = this.editing;
    if (!ed || ed.saving) return;
    const ta = this.tbody.querySelector('.dg-editor');
    if (ta) ed.draft = ta.value;
    const row = this.byId.get(ed.id);
    const col = this._col(ed.key);
    ed.saving = true;
    ed.error = null;
    this._renderBody();

    Promise.resolve()
      .then(function () { return col.edit(row, ed.draft); })
      .then(function () {
        if (self.editing !== ed) return;
        self.editing = null;
        if (!self.serverMode) self._refreshOptions();
        self._renderBody();
      }, function (err) {
        if (self.editing !== ed) return;
        ed.saving = false;
        ed.error = (err && err.message) || String(err);
        self._renderBody();
        self._focusEditor();
      });
  };

  window.DataGrid = DataGrid;
})();
