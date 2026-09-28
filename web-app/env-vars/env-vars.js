/**
 * EF Power Platform Tools — Environment Variables (web-app)
 *
 * Grid of every environment variable definition with its default value and
 * current value (the environmentvariablevalue record, if any). The current
 * value is editable in place: PATCH when a value record exists, POST one when
 * it doesn't, DELETE it when cleared (falls back to the default value).
 */
(function () {
  'use strict';

  const D = window.EFD365;
  const esc = D.escHtml;

  const TYPE_NAMES = {
    100000000: 'String',
    100000001: 'Number',
    100000002: 'Boolean',
    100000003: 'JSON',
    100000004: 'Data source',
    100000005: 'Secret',
  };
  const TYPE = { NUMBER: 100000001, BOOLEAN: 100000002, JSON: 100000003 };

  let grid;

  document.addEventListener('DOMContentLoaded', function () {
    document.getElementById('env-url').textContent = D.envUrl;
    document.getElementById('env-name').textContent = D.envName;
    document.title = 'Environment Variables — ' + D.envName;

    grid = new window.DataGrid(document.getElementById('grid'), {
      rowKey: function (r) { return r.id; },
      defaultSort: { key: 'displayName', dir: 'asc' },
      onCountChange: function (shown, total) {
        document.getElementById('count').textContent =
          shown === total ? total + ' variables' : shown + ' of ' + total + ' variables';
      },
      columns: [
        { key: 'displayName', label: 'Display Name', width: '19%', filter: 'text',
          render: function (r) {
            return '<span class="dg-text" title="' + esc(r.description || r.displayName) + '">' + esc(r.displayName || r.schemaName) + '</span>';
          } },
        { key: 'schemaName', label: 'Unique Name', width: '19%', filter: 'text', className: 'mono' },
        { key: 'type', label: 'Type', width: '9%', filter: 'select' },
        { key: 'introducedIn', label: 'Introduced In', width: '15%', filter: 'select',
          filterValues: function (r) { return r.solutions.map(function (s) { return s.friendlyname; }); },
          render: renderSolutions },
        { key: 'defaultValue', label: 'Default Value', width: '19%', filter: 'value', className: 'mono' },
        { key: 'currentValue', label: 'Current Value', width: '19%', filter: 'value', className: 'mono',
          editable: true,
          editValue: function (r) { return r.currentValue || ''; },
          editHint: 'Enter to save · Shift+Enter for a new line · Esc to cancel. Leave empty to remove the current value (falls back to the default).',
          edit: saveCurrentValue },
      ],
    });

    document.getElementById('search').addEventListener('input', function (e) { grid.setSearch(e.target.value); });
    document.getElementById('btn-clear').addEventListener('click', function () {
      document.getElementById('search').value = '';
      grid.clearFilters();
    });
    document.getElementById('btn-refresh').addEventListener('click', load);

    if (!D.envUrl) {
      showState('No environment URL provided. Open this tool from the EF Power Platform Tools launcher.', 'error');
      return;
    }
    load();
  });

  function showState(msg, kind) {
    const el = document.getElementById('state');
    el.textContent = msg;
    el.className = 'state-msg' + (kind ? ' state-' + kind : '');
    el.classList.remove('hidden');
    document.getElementById('grid').classList.add('hidden');
  }

  function renderSolutions(r) {
    if (!r.solutions.length) return '<span class="dg-blank">—</span>';
    const first = r.solutions[0].solution;
    const all = r.solutions.map(function (s) { return s.friendlyname; }).join('\n');
    return (
      '<a class="link" href="' + esc(D.solutionUrl(first.solutionid)) + '" target="_blank" rel="noopener" title="' + esc(all) + '">' +
        esc(first.friendlyname) + '</a>' +
      (first.ismanaged ? ' <span class="pill">Managed</span>' : '') +
      (r.solutions.length > 1 ? '<span class="more" title="Also in:\n' + esc(all) + '">+' + (r.solutions.length - 1) + '</span>' : '')
    );
  }

  async function load() {
    showState('Loading environment variables…', 'loading');
    try {
      const results = await Promise.all([
        D.fetchAll('/environmentvariabledefinitions?$select=environmentvariabledefinitionid,displayname,schemaname,defaultvalue,type,description,ismanaged'),
        D.fetchAll('/environmentvariablevalues?$select=environmentvariablevalueid,value,_environmentvariabledefinitionid_value'),
        D.solutionMap(),
      ]);
      const defs = results[0], values = results[1], solMap = results[2];

      const valueByDef = new Map();
      values.forEach(function (v) { valueByDef.set(v._environmentvariabledefinitionid_value, v); });

      const compSols = await D.componentSolutions(defs.map(function (d) { return d.environmentvariabledefinitionid; }), solMap);

      const rows = defs.map(function (d) {
        const v = valueByDef.get(d.environmentvariabledefinitionid);
        const sols = (compSols.get(d.environmentvariabledefinitionid) || []).map(function (x) {
          return { solution: x.solution, friendlyname: x.solution.friendlyname };
        });
        return {
          id:           d.environmentvariabledefinitionid,
          displayName:  d.displayname || '',
          schemaName:   d.schemaname || '',
          description:  d.description || '',
          typeCode:     d.type,
          type:         D.formatted(d, 'type') || TYPE_NAMES[d.type] || '',
          defaultValue: d.defaultvalue,
          currentValue: v ? v.value : null,
          valueId:      v ? v.environmentvariablevalueid : null,
          solutions:    sols,
          introducedIn: sols.length ? sols[0].friendlyname : '',
        };
      });

      document.getElementById('state').classList.add('hidden');
      document.getElementById('grid').classList.remove('hidden');
      grid.setRows(rows);
    } catch (err) {
      console.error('[EF PPT] Environment variables load failed:', err);
      showState('Failed to load environment variables: ' + err.message, 'error');
    }
  }

  function validate(row, text) {
    const t = text.trim();
    if (row.typeCode === TYPE.NUMBER && !isFinite(Number(t))) {
      throw new Error('"' + t + '" is not a valid number.');
    }
    if (row.typeCode === TYPE.BOOLEAN && !/^(yes|no|true|false)$/i.test(t)) {
      throw new Error('Boolean values must be yes or no.');
    }
    if (row.typeCode === TYPE.JSON) {
      try { JSON.parse(t); } catch (e) { throw new Error('Invalid JSON: ' + e.message); }
    }
  }

  async function saveCurrentValue(row, text) {
    if (text.trim() === '') {
      if (!row.valueId) return;
      await D.request('/environmentvariablevalues(' + row.valueId + ')', { method: 'DELETE' });
      row.valueId = null;
      row.currentValue = null;
      return;
    }

    validate(row, text);

    if (row.valueId) {
      await D.request('/environmentvariablevalues(' + row.valueId + ')', {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: { value: text },
      });
    } else {
      const nav = await D.navProperty('environmentvariablevalue', 'environmentvariabledefinitionid');
      const body = { value: text, schemaname: row.schemaName };
      body[nav + '@odata.bind'] = '/environmentvariabledefinitions(' + row.id + ')';
      const created = await D.request('/environmentvariablevalues', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Prefer: 'return=representation' },
        body: body,
      });
      row.valueId = created && created.environmentvariablevalueid;
    }
    row.currentValue = text;
  }
})();
