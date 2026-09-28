/**
 * EF Power Platform Tools — Connection References (web-app)
 *
 * Grid of every connection reference: connector, the connection it's bound
 * to, solution, status and owner. The bound connection is editable in place
 * (PATCH connectionid); the connection's signed-in account and health live
 * in the Power Platform connections service, so they're linked to rather
 * than shown.
 */
(function () {
  'use strict';

  const D = window.EFD365;
  const esc = D.escHtml;

  const CONNECTOR_NAMES = {
    shared_commondataserviceforapps: 'Microsoft Dataverse',
    shared_commondataservice:        'Dataverse (legacy)',
    shared_office365:                'Office 365 Outlook',
    shared_office365users:           'Office 365 Users',
    shared_office365groups:          'Office 365 Groups',
    shared_sharepointonline:         'SharePoint',
    shared_teams:                    'Microsoft Teams',
    shared_onedriveforbusiness:      'OneDrive for Business',
    shared_excelonlinebusiness:      'Excel Online (Business)',
    shared_wordonlinebusiness:       'Word Online (Business)',
    shared_planner:                  'Planner',
    shared_approvals:                'Approvals',
    shared_outlook:                  'Outlook.com',
    shared_sendmail:                 'Mail',
    shared_sql:                      'SQL Server',
    shared_azureblob:                'Azure Blob Storage',
    shared_azurequeues:              'Azure Queues',
    shared_keyvault:                 'Azure Key Vault',
    shared_servicebus:               'Service Bus',
    shared_documentdb:               'Azure Cosmos DB',
    shared_azuread:                  'Microsoft Entra ID',
    shared_webcontents:              'HTTP with Microsoft Entra ID',
    shared_flowmanagement:           'Power Automate Management',
    shared_powerappsforappmakers:    'Power Apps for Makers',
    shared_dynamicscrmonline:        'Dynamics 365 (deprecated)',
  };

  let grid;

  function connectorApiName(connectorId) {
    return String(connectorId || '').split('/').pop();
  }

  function titleCase(s) {
    return s.split(/[\s_]+/).filter(Boolean)
      .map(function (w) { return w.charAt(0).toUpperCase() + w.slice(1); }).join(' ');
  }

  function connectorLabel(connectorId) {
    const api = connectorApiName(connectorId);
    if (!api) return '';
    if (CONNECTOR_NAMES[api]) return CONNECTOR_NAMES[api];
    const raw = api.replace(/^shared_/, '');
    // Custom connector API names hex-escape their characters ("-5f" = "_",
    // "-20" = space) as <publisher prefix>_<name>_<hash>.
    if (/-[0-9a-f]{2}/i.test(raw)) {
      const parts = raw.replace(/-([0-9a-f]{2})/gi, function (_, h) { return String.fromCharCode(parseInt(h, 16)); }).split('_');
      if (parts.length > 2 && /^[0-9a-f]+$/i.test(parts[parts.length - 1])) parts.pop();
      if (parts.length > 1) parts.shift();
      return titleCase(parts.join(' ')) + ' (custom)';
    }
    return titleCase(raw);
  }

  function connectionUrl(row) {
    if (!D.paEnvId || !row.connectionId) return null;
    return 'https://make.powerapps.com/environments/' + D.paEnvId + '/connections/' +
      encodeURIComponent(row.connectorApi) + '/' + encodeURIComponent(row.connectionId) + '/details';
  }

  document.addEventListener('DOMContentLoaded', function () {
    document.getElementById('env-url').textContent = D.envUrl;
    document.getElementById('env-name').textContent = D.envName;
    document.title = 'Connection References — ' + D.envName;
    if (!D.paEnvId) document.getElementById('pa-notice').classList.remove('hidden');

    grid = new window.DataGrid(document.getElementById('grid'), {
      rowKey: function (r) { return r.id; },
      defaultSort: { key: 'displayName', dir: 'asc' },
      onCountChange: function (shown, total) {
        document.getElementById('count').textContent =
          shown === total ? total + ' connection references' : shown + ' of ' + total + ' connection references';
      },
      columns: [
        { key: 'displayName', label: 'Display Name', width: '16%', filter: 'text',
          filterValues: function (r) { return [r.displayName, r.description]; },
          render: function (r) {
            return '<span class="dg-text" title="' + esc(r.description || r.displayName) + '">' + esc(r.displayName || r.logicalName) + '</span>';
          } },
        { key: 'logicalName', label: 'Logical Name', width: '14%', filter: 'text', className: 'mono' },
        { key: 'connector', label: 'Type (Connector)', width: '13%', filter: 'select',
          render: function (r) {
            if (!r.connector) return '<span class="dg-blank">—</span>';
            return esc(r.connector) + '<span class="sub mono">' + esc(r.connectorApi) + '</span>';
          } },
        { key: 'connectionId', label: 'Connection', width: '16%', filter: 'value', className: 'mono',
          editable: true,
          editValue: function (r) { return r.connectionId || ''; },
          editHint: 'Paste a connection ID, or the connection\'s URL from Power Apps. Leave empty to unbind. Enter to save, Esc to cancel.',
          edit: saveConnection,
          render: function (r) {
            if (!r.connectionId) return '<span class="pill pill--warn">Not connected</span>';
            return '<span class="dg-text" title="' + esc(r.connectionId) + '">' + esc(r.connectionId) + '</span>';
          } },
        { key: 'introducedIn', label: 'Introduced In', width: '13%', filter: 'select',
          filterValues: function (r) { return r.solutions.map(function (s) { return s.friendlyname; }); },
          render: renderSolutions },
        { key: 'status', label: 'Status', width: '8%', filter: 'select',
          render: function (r) {
            const cls = r.stateCode === 0 ? 'pill--ok' : 'pill--bad';
            return '<span class="pill ' + cls + '">' + esc(r.status) + '</span>';
          } },
        { key: 'owner', label: 'Owner', width: '9%', filter: 'text' },
        { key: 'actions', label: 'Open', width: '11%', sortable: false, searchable: false,
          render: renderActions },
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

  function renderActions(r) {
    const conn = connectionUrl(r);
    const connLink = conn
      ? '<a class="link" href="' + esc(conn) + '" target="_blank" rel="noopener">View connection ↗</a>'
      : '<span class="disabled" title="' + (r.connectionId ? 'Needs this environment\'s powerAppsId in your config' : 'No connection bound') + '">View connection</span>';
    const connectionsList = D.paEnvId
      ? '<a class="link" href="https://make.powerapps.com/environments/' + esc(D.paEnvId) + '/connections" target="_blank" rel="noopener">All connections ↗</a>'
      : '';
    return '<div class="actions">' + connLink + connectionsList + '</div>';
  }

  async function load() {
    showState('Loading connection references…', 'loading');
    try {
      const results = await Promise.all([
        D.fetchAll('/connectionreferences?$select=connectionreferenceid,connectionreferencedisplayname,connectionreferencelogicalname,' +
          'connectorid,connectionid,description,statecode,statuscode,ismanaged,_ownerid_value'),
        D.solutionMap(),
      ]);
      const refs = results[0], solMap = results[1];
      const compSols = await D.componentSolutions(refs.map(function (r) { return r.connectionreferenceid; }), solMap);

      const rows = refs.map(function (c) {
        const sols = (compSols.get(c.connectionreferenceid) || []).map(function (x) {
          return { solution: x.solution, friendlyname: x.solution.friendlyname };
        });
        return {
          id:           c.connectionreferenceid,
          displayName:  c.connectionreferencedisplayname || '',
          logicalName:  c.connectionreferencelogicalname || '',
          description:  c.description || '',
          connectorApi: connectorApiName(c.connectorid),
          connector:    connectorLabel(c.connectorid),
          connectionId: c.connectionid || null,
          stateCode:    c.statecode,
          status:       D.formatted(c, 'statuscode') || (c.statecode === 0 ? 'Active' : 'Inactive'),
          owner:        D.formatted(c, '_ownerid_value') || '',
          solutions:    sols,
          introducedIn: sols.length ? sols[0].friendlyname : '',
        };
      });

      document.getElementById('state').classList.add('hidden');
      document.getElementById('grid').classList.remove('hidden');
      grid.setRows(rows);
    } catch (err) {
      console.error('[EF PPT] Connection references load failed:', err);
      showState('Failed to load connection references: ' + err.message, 'error');
    }
  }

  /** Accepts a bare connection id, or a Power Apps connection URL containing one. */
  function parseConnectionId(text) {
    const t = text.trim();
    const m = t.match(/\/connections\/[^/?#]+\/([^/?#]+)/i);
    return m ? decodeURIComponent(m[1]) : t;
  }

  async function saveConnection(row, text) {
    const id = parseConnectionId(text) || null;
    await D.request('/connectionreferences(' + row.id + ')', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: { connectionid: id },
    });
    row.connectionId = id;
  }
})();
