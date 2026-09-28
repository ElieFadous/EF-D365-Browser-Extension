/**
 * EF Power Platform Tools — shared D365 helpers for web-app tool pages.
 *
 * Tool pages run as iframes/tabs opened by launcher.js; every Web API call is
 * relayed through launcher.js's postMessage bridge (it runs same-origin in the
 * D365 page). Exposes window.EFD365.
 */
(function () {
  'use strict';

  const params = new URLSearchParams(location.search);

  function loadCfg() {
    try {
      const fromUrl = params.get('cfg');
      if (fromUrl) return JSON.parse(decodeURIComponent(escape(atob(fromUrl))));
    } catch (_) { /* fall through to localStorage */ }
    try { return JSON.parse(localStorage.getItem('ef_ppt_config')) || {}; }
    catch (_) { return {}; }
  }

  const cfg     = loadCfg();
  const envUrl  = (params.get('env') || '').replace(/\/$/, '');
  const envName = params.get('name') || envUrl;
  const paEnvId = params.get('paEnvId') || '';
  const apiBase = envUrl + '/api/data/' + ((cfg.settings && cfg.settings.apiVersion) || 'v9.2');

  const FORMATTED = 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"';

  /** Relays a Web API call through launcher.js. Resolves with the parsed body, rejects with an Error. */
  function request(path, opts) {
    opts = opts || {};
    const url = /^https?:/i.test(path) ? path : apiBase + path;
    return new Promise(function (resolve, reject) {
      const id = Math.random().toString(36).slice(2) + Date.now();
      const timer = setTimeout(function () {
        window.removeEventListener('message', onMsg);
        reject(new Error('Request timed out after 30s'));
      }, 30000);
      function onMsg(e) {
        const d = e.data;
        if (!d || d.__efppt !== 'fetch-result' || d.id !== id) return;
        clearTimeout(timer);
        window.removeEventListener('message', onMsg);
        if (d.ok) resolve(d.data);
        else reject(new Error(d.error || (d.data && d.data.error && d.data.error.message) || ('HTTP ' + d.status)));
      }
      window.addEventListener('message', onMsg);
      (window.opener || window.parent).postMessage({
        __efppt: 'fetch', id: id, url: url,
        method: opts.method || 'GET',
        headers: opts.headers || {},
        body: opts.body == null ? null : opts.body,
      }, '*');
    });
  }

  /** GETs a collection, following @odata.nextLink until exhausted. */
  async function fetchAll(path) {
    const out = [];
    let url = path;
    while (url) {
      const d = await request(url, { headers: { Prefer: FORMATTED + ',odata.maxpagesize=5000' } });
      (d.value || []).forEach(function (v) { out.push(v); });
      url = d['@odata.nextLink'] || null;
    }
    return out;
  }

  function formatted(record, attr) {
    return record[attr + '@OData.Community.Display.V1.FormattedValue'];
  }

  async function solutionMap() {
    const sols = await fetchAll('/solutions?$select=solutionid,friendlyname,uniquename,ismanaged,isvisible,version');
    const map = new Map();
    sols.forEach(function (s) { map.set(s.solutionid, s); });
    return map;
  }

  /**
   * Maps each component objectid to the user-facing solutions containing it,
   * oldest membership first — so [0] is the solution it was introduced in.
   * Filters by objectid rather than componenttype because newer component
   * types (e.g. connection references) have org-specific componenttype codes.
   */
  async function componentSolutions(objectIds, solMap) {
    const result = new Map();
    const ids = Array.from(new Set(objectIds.filter(Boolean)));
    const chunks = [];
    for (let i = 0; i < ids.length; i += 40) chunks.push(ids.slice(i, i + 40));

    const pages = await Promise.all(chunks.map(function (chunk) {
      const filter = chunk.map(function (id) { return 'objectid eq ' + id; }).join(' or ');
      return fetchAll('/solutioncomponents?$select=objectid,createdon,_solutionid_value&$filter=' + encodeURIComponent(filter));
    }));

    pages.forEach(function (comps) {
      comps.forEach(function (c) {
        const s = solMap.get(c._solutionid_value);
        if (!s || s.isvisible === false) return;
        const u = (s.uniquename || '').toLowerCase();
        if (u === 'default' || u === 'active') return;
        const list = result.get(c.objectid) || [];
        if (!list.some(function (x) { return x.solution.solutionid === s.solutionid; })) {
          list.push({ solution: s, createdon: c.createdon || '' });
        }
        result.set(c.objectid, list);
      });
    });

    result.forEach(function (list) {
      list.sort(function (a, b) { return String(a.createdon).localeCompare(String(b.createdon)); });
    });
    return result;
  }

  const _navCache = new Map();

  /** Single-valued navigation property for a lookup attribute (for @odata.bind). */
  async function navProperty(entity, attribute) {
    const key = entity + '.' + attribute;
    if (_navCache.has(key)) return _navCache.get(key);
    const d = await request("/EntityDefinitions(LogicalName='" + entity + "')/ManyToOneRelationships" +
      '?$select=ReferencingAttribute,ReferencingEntityNavigationPropertyName');
    const rel = (d.value || []).filter(function (r) { return r.ReferencingAttribute === attribute; })[0];
    const nav = (rel && rel.ReferencingEntityNavigationPropertyName) || attribute;
    _navCache.set(key, nav);
    return nav;
  }

  /** Link to a solution: maker portal when the Power Apps env id is known, classic explorer otherwise. */
  function solutionUrl(solutionId) {
    return paEnvId
      ? 'https://make.powerapps.com/environments/' + paEnvId + '/solutions/' + solutionId
      : envUrl + '/tools/solution/edit.aspx?id=' + solutionId;
  }

  function escHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  window.EFD365 = {
    cfg: cfg, envUrl: envUrl, envName: envName, paEnvId: paEnvId, apiBase: apiBase,
    request: request, fetchAll: fetchAll, formatted: formatted,
    solutionMap: solutionMap, componentSolutions: componentSolutions,
    navProperty: navProperty, solutionUrl: solutionUrl, escHtml: escHtml,
  };
})();
