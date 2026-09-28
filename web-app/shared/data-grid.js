/**
 * EF Power Platform Tools — DataGrid: sortable columns, per-column filters,
 * global search and inline cell editing. Exposes window.DataGrid.
 *
 * Column spec:
 *   key, label, width        — identity / header text / optional CSS width
 *   value(row)               — raw value for sort, filter and search (default row[key])
 *   filterValues(row)        — values a select filter matches against (default [value])
 *   sortValue(row)           — override sort key
 *   render(row)              — cell HTML (default: escaped value, "—" when blank)
 *   filter                   — 'text' | 'select' | 'value' (Any/Blank/Not blank + contains) | null
 *   sortable                 — default true
 *   editable, edit(row, v)   — inline edit; edit() returns a Promise and mutates row
 *   editValue(row), editHint — initial editor text / hint shown under the editor
 *   editOptions(row)         — [{ value, label }] to edit with a dropdown instead of text
 *
 * Options: columns, rowKey, defaultSort, onCountChange, tableMinWidth (CSS length)
 */
(function () {
  'use strict';

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
    this.rows = [];
    this.byId = new Map();
    this.sort = opts.defaultSort || null;
    this.filters = {};
    this.search = '';
    this.editing = null; // { id, key, draft, error, saving }
    this.tableMinWidth = opts.tableMinWidth || null;
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

  DataGrid.prototype._filterHtml = function (c) {
    const k = esc(c.key);
    if (c.filter === 'text') {
      return '<input type="search" class="dg-f-text" data-key="' + k + '" placeholder="Filter…" autocomplete="off">';
    }
    if (c.filter === 'select') {
      return '<select class="dg-f-select" data-key="' + k + '"><option value="">All</option></select>';
    }
    if (c.filter === 'value') {
      return (
        '<div class="dg-f-value">' +
          '<select class="dg-f-mode" data-key="' + k + '">' +
            '<option value="">Any</option><option value="blank">Blank</option><option value="notblank">Not blank</option>' +
          '</select>' +
          '<input type="search" class="dg-f-text" data-key="' + k + '" placeholder="Contains…" autocomplete="off">' +
        '</div>'
      );
    }
    return '';
  };

  DataGrid.prototype._build = function () {
    const self = this;
    this.host.classList.add('dg');
    this.host.innerHTML =
      '<div class="dg-scroll"><table class="dg-table"' +
        (this.tableMinWidth ? ' style="min-width:' + esc(this.tableMinWidth) + '"' : '') + '>' +
        '<colgroup>' + this.cols.map(function (c) {
          return '<col' + (c.width ? ' style="width:' + c.width + '"' : '') + '>';
        }).join('') + '</colgroup>' +
        '<thead>' +
          '<tr class="dg-head">' + this.cols.map(function (c) {
            return '<th>' + (c.sortable === false
              ? '<span class="dg-th-label">' + esc(c.label) + '</span>'
              : '<button type="button" class="dg-sort" data-key="' + esc(c.key) + '">' +
                  '<span>' + esc(c.label) + '</span><span class="dg-arrow"></span></button>') + '</th>';
          }).join('') + '</tr>' +
          '<tr class="dg-filters">' + this.cols.map(function (c) {
            return '<th>' + self._filterHtml(c) + '</th>';
          }).join('') + '</tr>' +
        '</thead>' +
        '<tbody></tbody>' +
      '</table></div>';
    this.tbody = this.host.querySelector('tbody');

    this.host.querySelector('thead').addEventListener('click', function (e) {
      const b = e.target.closest('.dg-sort');
      if (!b) return;
      const key = b.dataset.key;
      self.sort = (self.sort && self.sort.key === key)
        ? { key: key, dir: self.sort.dir === 'asc' ? 'desc' : 'asc' }
        : { key: key, dir: 'asc' };
      self._renderHead();
      self._renderBody();
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
  };

  DataGrid.prototype._setFilter = function (key, prop, value) {
    const f = this.filters[key] || (this.filters[key] = {});
    f[prop] = value;
    this._renderBody();
  };

  DataGrid.prototype.setSearch = function (q) {
    this.search = (q || '').trim().toLowerCase();
    this._renderBody();
  };

  DataGrid.prototype.clearFilters = function () {
    this.filters = {};
    this.search = '';
    this.host.querySelectorAll('.dg-f-text').forEach(function (el) { el.value = ''; });
    this.host.querySelectorAll('.dg-f-select, .dg-f-mode').forEach(function (el) { el.value = ''; });
    this._renderBody();
  };

  DataGrid.prototype.setRows = function (rows) {
    const self = this;
    this.rows = rows;
    this.byId = new Map();
    rows.forEach(function (r) { self.byId.set(String(self.rowKey(r)), r); });
    this.editing = null;
    this._refreshOptions();
    this._renderHead();
    this._renderBody();
  };

  DataGrid.prototype._refreshOptions = function () {
    const self = this;
    this.cols.filter(function (c) { return c.filter === 'select'; }).forEach(function (c) {
      const sel = self.host.querySelector('.dg-f-select[data-key="' + c.key + '"]');
      if (!sel) return;
      const values = new Set();
      let hasBlank = false;
      self.rows.forEach(function (r) {
        const vals = self._filterVals(c, r);
        if (!vals.length || vals.every(isBlank)) hasBlank = true;
        vals.forEach(function (v) { if (!isBlank(v)) values.add(String(v)); });
      });
      const current = sel.value;
      const opts = Array.from(values).sort(compare);
      sel.innerHTML =
        '<option value="">All</option>' +
        (hasBlank ? '<option value="__blank__">(blank)</option>' : '') +
        opts.map(function (v) { return '<option value="' + esc(v) + '">' + esc(v) + '</option>'; }).join('');
      sel.value = (current === '__blank__' && hasBlank) || opts.indexOf(current) !== -1 ? current : '';
      if (sel.value !== current) self._setFilterSilently(c.key, 'value', sel.value);
    });
  };

  DataGrid.prototype._setFilterSilently = function (key, prop, value) {
    const f = this.filters[key] || (this.filters[key] = {});
    f[prop] = value;
  };

  DataGrid.prototype._passes = function (row) {
    const self = this;
    if (this.search) {
      const q = this.search;
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
      this.tbody.innerHTML = '<tr><td class="dg-empty" colspan="' + this.cols.length + '">' +
        (this.rows.length ? 'No rows match the current filters.' : 'No records found.') + '</td></tr>';
      return;
    }
    this.tbody.innerHTML = rows.map(function (r) {
      return '<tr data-id="' + esc(self.rowKey(r)) + '">' +
        self.cols.map(function (c) { return self._cellHtml(r, c); }).join('') + '</tr>';
    }).join('');
  };

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
        self._refreshOptions();
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
