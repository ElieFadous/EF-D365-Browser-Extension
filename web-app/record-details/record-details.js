/**
 * EF Power Platform Tools — Record Details (web-app)
 *
 * Shows metadata for the D365 record the launcher was opened from (created/
 * modified by/on, owner), lets you clone it (same environment or a different
 * one), and offers quick "Open In…" links to other configured environments.
 *
 * Runs as an iframe/tab opened by launcher.js. All D365 API calls go through
 * the postMessage fetch bridge — no direct fetch, no chrome.* APIs.
 *
 * Cross-environment clone reuses the same bridge Data Sync uses: launcher.js
 * routes each fetch by the URL's origin — same-origin directly, or via a
 * second tab connected through "Connect Target Environment" for a different
 * org. cloneRecord() below never needs to know which case it's in.
 */

// ─── Bridge fetch ───────────────────────────────────────────────────────────

/** Resolves with the full { ok, status, data, error } envelope — never rejects. */
const _bridgeRequest = (url, extraHeaders = {}, method = 'GET', body = null) =>
  new Promise((resolve) => {
    const id    = Math.random().toString(36).slice(2) + Date.now();
    const timer = setTimeout(() => {
      window.removeEventListener('message', onMsg);
      resolve({ ok: false, status: 0, error: 'Request timed out after 30s' });
    }, 30_000);
    function onMsg(e) {
      if (!e.data || e.data.__efppt !== 'fetch-result' || e.data.id !== id) return;
      clearTimeout(timer);
      window.removeEventListener('message', onMsg);
      resolve(e.data);
    }
    window.addEventListener('message', onMsg);
    (window.opener || window.parent).postMessage({ __efppt: 'fetch', id, url, method, headers: extraHeaders, body }, '*');
  });

/** Resolves with the parsed body on success, rejects with an Error otherwise. */
const _bridgeFetch = (url, extraHeaders = {}, method = 'GET', body = null) =>
  _bridgeRequest(url, extraHeaders, method, body).then(r => {
    if (r.ok) return r.data;
    throw new Error(r.error || (r.data && r.data.error && r.data.error.message) || `HTTP ${r.status}`);
  });

const _bridgeConnectTarget = (targetOrigin) =>
  new Promise((resolve, reject) => {
    const id    = Math.random().toString(36).slice(2) + Date.now();
    const timer = setTimeout(() => {
      window.removeEventListener('message', onMsg);
      reject(new Error('Timed out waiting for the target tab to connect.'));
    }, 130_000);
    function onMsg(e) {
      if (!e.data || e.data.__efppt !== 'connect-target-result' || e.data.id !== id) return;
      clearTimeout(timer);
      window.removeEventListener('message', onMsg);
      if (e.data.ok) resolve();
      else reject(new Error(e.data.error || 'Could not connect to the target environment.'));
    }
    window.addEventListener('message', onMsg);
    (window.opener || window.parent).postMessage({ __efppt: 'connect-target', id, targetOrigin }, '*');
  });

// ─── Config ───────────────────────────────────────────────────────────────────

function _loadConfig() {
  try {
    const fromUrl = new URLSearchParams(location.search).get('cfg');
    if (fromUrl) return JSON.parse(decodeURIComponent(escape(atob(fromUrl))));
  } catch (_) { /* fall through to localStorage */ }
  try { return JSON.parse(localStorage.getItem('ef_ppt_config')) ?? {}; }
  catch { return {}; }
}

// ─── State ────────────────────────────────────────────────────────────────────

// Derived from this script's own src (not location.href) so it's unaffected
// by clean-URL rewriting some static hosts apply to the .html document.
const BASE_URL =
  (document.currentScript && document.currentScript.src
    ? document.currentScript.src
    : ''
  ).replace(/\/record-details\/record-details\.js.*$/, '') || location.origin;

let environments   = [];
let settings        = {};
let currentEnv       = null; // { url, name, color?, warn? } — from config, or a fallback if unlisted
let pageEtn          = '';
let pageId           = '';
let entitySetName    = null;
let _connectedTargetOrigin = null;

const _entitySetNameCache = new Map();

// ─── Bootstrap ────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', init);

async function init() {
  const config = _loadConfig();
  environments = Array.isArray(config.environments) ? config.environments : [];
  settings = {
    apiVersion: 'v9.2',
    clonePrefix: '',
    cloneWhitelist: null,
    cloneLookupMode: 'skip',
    ...(config.settings ?? {}),
  };

  const params  = new URLSearchParams(location.search);
  const envUrl  = (params.get('env') ?? '').replace(/\/$/, '');
  const envName = params.get('name') ?? '';
  pageEtn = params.get('etn') ?? '';
  pageId  = (params.get('id') ?? '').replace(/[{}]/g, '');

  currentEnv = environments.find(e => (e.url ?? '').replace(/\/$/, '') === envUrl)
    ?? (envUrl ? { url: envUrl, name: envName || envUrl, color: '#1B3A6B' } : null);

  document.getElementById('env-url').textContent  = currentEnv?.url  ?? '';
  document.getElementById('env-name').textContent = currentEnv?.name ?? '';

  document.getElementById('clone-err-modal-close').addEventListener('click', () => {
    document.getElementById('clone-err-overlay').classList.add('hidden');
  });
  document.getElementById('clone-err-overlay').addEventListener('click', (e) => {
    if (e.target.id === 'clone-err-overlay') e.target.classList.add('hidden');
  });
  document.getElementById('btn-open-metadata').addEventListener('click', () => {
    window.open(buildToolUrl('metadata', { etn: pageEtn }), '_blank');
  });

  if (!currentEnv || !pageEtn || !pageId) {
    showState('Not on a Dynamics 365 record form.', 'info');
    return;
  }

  try {
    entitySetName = await fetchEntitySetName(currentEnv.url, pageEtn);
    if (!entitySetName) throw new Error('Could not resolve entity type.');
    const data = await fetchRecordDetails(currentEnv.url, entitySetName);
    renderRecordDetails(data);
  } catch (err) {
    showState(`Failed to load: ${err.message}`, 'error');
  }
}

function showState(msg, type) {
  const el = document.getElementById('state-container');
  el.textContent = msg;
  el.className = `state-msg state-${type}`;
  el.classList.remove('hidden');
  document.getElementById('rd-root').classList.add('hidden');
}

// ─── Tool URL builder (for "Open in Metadata Browser") ────────────────────────

function buildToolUrl(toolName, extraParams) {
  const params = new URLSearchParams();
  params.set('env', currentEnv.url);
  if (currentEnv.name)       params.set('name', currentEnv.name);
  if (currentEnv.powerAppsId) params.set('paEnvId', currentEnv.powerAppsId);
  params.set('_inModal', '1');
  const rawCfg = new URLSearchParams(location.search).get('cfg');
  if (rawCfg) params.set('cfg', rawCfg);
  if (extraParams) {
    Object.keys(extraParams).forEach(k => { if (extraParams[k] != null) params.set(k, extraParams[k]); });
  }
  return `${BASE_URL}/${toolName}/${toolName}.html?${params.toString()}`;
}

// ─── Entity set name resolver (works for any environment origin) ──────────────

async function fetchEntitySetName(baseUrl, etn) {
  const cacheKey = `${baseUrl}:${etn}`;
  if (_entitySetNameCache.has(cacheKey)) return _entitySetNameCache.get(cacheKey);
  try {
    const d = await _bridgeFetch(`${baseUrl}/api/data/${settings.apiVersion}/EntityDefinitions(LogicalName='${etn}')?$select=EntitySetName`);
    const name = d?.EntitySetName ?? null;
    if (name) _entitySetNameCache.set(cacheKey, name);
    return name;
  } catch (e) {
    console.error('[Record Details] Failed to resolve entity set name:', e);
    return null;
  }
}

// ─── Fetch + render record details ─────────────────────────────────────────────

async function fetchRecordDetails(baseUrl, entitySetName) {
  const url = `${baseUrl}/api/data/${settings.apiVersion}/${entitySetName}(${pageId})` +
    `?$select=createdon,modifiedon,_ownerid_value` +
    `&$expand=createdby($select=fullname,systemuserid),modifiedby($select=fullname,systemuserid)`;
  return _bridgeFetch(url, {
    'Prefer': 'odata.include-annotations="OData.Community.Display.V1.FormattedValue,Microsoft.Dynamics.CRM.lookuplogicalname"',
  });
}

function escHtml(str) {
  return String(str == null ? '' : str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** Formats an ISO 8601 timestamp as DD/MM/YYYY HH:mm (local time). */
function _formatDateTime(iso) {
  if (!iso) return '—';
  try {
    const d   = new Date(iso);
    const dd  = String(d.getDate()).padStart(2, '0');
    const mm  = String(d.getMonth() + 1).padStart(2, '0');
    const hh  = String(d.getHours()).padStart(2, '0');
    const min = String(d.getMinutes()).padStart(2, '0');
    return `${dd}/${mm}/${d.getFullYear()} ${hh}:${min}`;
  } catch {
    return iso;
  }
}

function _d365RecordUrl(etn, id) {
  if (!currentEnv?.url || !id) return null;
  return `${currentEnv.url}/main.aspx?pagetype=entityrecord&etn=${etn}&id=${id}`;
}

function copyBtn(text) {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'rd-copy-btn';
  btn.title = 'Copy';
  btn.innerHTML = `<svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3.5" y="2" width="8" height="10" rx="1"/><path d="M2 4.5H1.5a.5.5 0 0 0-.5.5v7a.5.5 0 0 0 .5.5h7a.5.5 0 0 0 .5-.5V12"/></svg>`;
  btn.addEventListener('click', () => {
    navigator.clipboard.writeText(text).then(() => {
      btn.classList.add('copied');
      setTimeout(() => btn.classList.remove('copied'), 1500);
    });
  });
  return btn;
}

/** Builds an rd-card with one or more copyable value rows. */
function buildCard(label, rows, wide) {
  const card = document.createElement('div');
  card.className = 'rd-card' + (wide ? ' rd-card--wide' : '');

  const lbl = document.createElement('span');
  lbl.className = 'rd-label';
  lbl.textContent = label;
  card.appendChild(lbl);

  rows.forEach(({ text, mono, link, icon }) => {
    if (text == null) return;
    const row = document.createElement('div');
    row.className = 'rd-value-row';

    if (icon) {
      const iconEl = document.createElement('span');
      iconEl.className = 'rd-icon';
      iconEl.innerHTML = icon;
      row.appendChild(iconEl);
    }

    let valEl;
    if (link) {
      valEl = document.createElement('a');
      valEl.className = mono ? 'rd-mono rd-link' : 'rd-value rd-link';
      valEl.href = link;
      valEl.target = '_blank';
      valEl.rel = 'noopener noreferrer';
    } else {
      valEl = document.createElement('span');
      valEl.className = mono ? 'rd-mono' : 'rd-value';
    }
    valEl.textContent = text;
    row.appendChild(valEl);
    row.appendChild(copyBtn(text));
    card.appendChild(row);
  });

  return card;
}

function renderRecordDetails(data) {
  document.getElementById('state-container').classList.add('hidden');
  document.getElementById('rd-root').classList.remove('hidden');

  const cardsEl = document.getElementById('rd-cards');
  cardsEl.innerHTML = '';

  const apiLink = `${currentEnv.url}/api/data/${settings.apiVersion}/${entitySetName}(${pageId})`;
  cardsEl.appendChild(buildCard('Record ID',   [{ text: pageId, mono: true, link: apiLink }]));
  cardsEl.appendChild(buildCard('Entity Type', [{ text: pageEtn }]));

  const createdBy  = data.createdby;
  const modifiedBy = data.modifiedby;

  cardsEl.appendChild(buildCard('Created On', [{ text: _formatDateTime(data.createdon) }]));
  cardsEl.appendChild(buildCard('Created By', [
    { text: createdBy?.fullname ?? '—', link: _d365RecordUrl('systemuser', createdBy?.systemuserid) },
    { text: createdBy?.systemuserid ?? null, mono: true },
  ]));

  cardsEl.appendChild(buildCard('Modified On', [{ text: _formatDateTime(data.modifiedon) }]));
  cardsEl.appendChild(buildCard('Modified By', [
    { text: modifiedBy?.fullname ?? '—', link: _d365RecordUrl('systemuser', modifiedBy?.systemuserid) },
    { text: modifiedBy?.systemuserid ?? null, mono: true },
  ]));

  const ownerName = data['_ownerid_value@OData.Community.Display.V1.FormattedValue'] ?? null;
  const ownerId   = data['_ownerid_value'] ?? null;
  const ownerType = data['_ownerid_value@Microsoft.Dynamics.CRM.lookuplogicalname'] ?? 'systemuser';
  if (ownerName || ownerId) {
    const isTeam   = ownerType === 'team';
    const userIcon = `<svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="7" cy="4.5" r="2.2"/><path d="M2 12.5c0-2.76 2.24-4.5 5-4.5s5 1.74 5 4.5"/></svg>`;
    const teamIcon = `<svg viewBox="0 0 18 13" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="4" cy="4" r="2"/><path d="M1 12c0-1.66 1.34-3 3-3s3 1.34 3 3"/><circle cx="9" cy="3.5" r="2.2"/><path d="M5.5 12c0-1.93 1.57-3.5 3.5-3.5s3.5 1.57 3.5 3.5"/><circle cx="14" cy="4" r="2"/><path d="M11 12c0-1.66 1.34-3 3-3s3 1.34 3 3"/></svg>`;
    cardsEl.appendChild(buildCard('Owner', [
      { text: ownerName ?? '—', icon: isTeam ? teamIcon : userIcon, link: _d365RecordUrl(isTeam ? 'team' : 'systemuser', ownerId) },
      { text: ownerId ?? null, mono: true },
    ]));
  }

  renderOpenInList();
  initCloneBar();
}

// ─── Open In list ───────────────────────────────────────────────────────────────

function isEnabledFor(env, featureId) {
  if (!Array.isArray(env.enabledFor) || env.enabledFor.length === 0) return true;
  return env.enabledFor.includes(featureId);
}

function renderOpenInList() {
  const list = document.getElementById('rd-open-in-list');
  list.innerHTML = '';

  const targets = environments.filter(env => env.url !== currentEnv.url && isEnabledFor(env, 'goto-open-in'));
  if (targets.length === 0) {
    list.innerHTML = '<div class="rd-open-in-empty">No other environments configured.</div>';
    return;
  }

  targets.forEach(env => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'rd-env-btn' + (env.warn ? ' rd-env-btn--warn' : '');
    btn.title = env.name;
    btn.style.setProperty('--env-color', env.color ?? '#1B3A6B');
    btn.innerHTML = `<span class="rd-env-btn__name">${escHtml(env.name)}</span>${env.warn ? '<span class="rd-env-btn__warn">Prod</span>' : ''}`;
    btn.addEventListener('click', () => {
      const url = `${env.url}/main.aspx?pagetype=entityrecord&etn=${encodeURIComponent(pageEtn)}&id=${encodeURIComponent(pageId)}`;
      window.open(url, '_blank', 'noopener');
    });
    list.appendChild(btn);
  });
}

// ─── Clone bar ──────────────────────────────────────────────────────────────────

function _originOf(url) { try { return new URL(url).origin; } catch { return ''; } }

function initCloneBar() {
  const bar = document.getElementById('clone-bar');
  const cloneAllowed = !settings.cloneWhitelist || settings.cloneWhitelist.includes((pageEtn || '').toLowerCase());
  if (!cloneAllowed) { bar.classList.add('hidden'); return; }
  bar.classList.remove('hidden');

  const select = document.getElementById('clone-target');
  select.innerHTML = '';
  const allCloneEnvs = [currentEnv, ...environments.filter(e => e.url !== currentEnv.url)];
  allCloneEnvs.forEach(env => select.appendChild(new Option(env.name, env.url)));
  select.value = currentEnv.url;

  select.addEventListener('change', updateConnectPanel);
  document.getElementById('btn-connect-target').addEventListener('click', onConnectTargetClick);
  document.getElementById('btn-clone').addEventListener('click', onCloneClick);

  updateConnectPanel();
}

function updateConnectPanel() {
  const select   = document.getElementById('clone-target');
  const panel    = document.getElementById('connect-target-panel');
  const cloneBtn = document.getElementById('btn-clone');
  const targetUrl = select.value;
  const isCross   = _originOf(targetUrl) !== _originOf(currentEnv.url);

  panel.classList.remove('alert--error');
  panel.classList.add('alert--warning');

  if (!isCross || _connectedTargetOrigin === _originOf(targetUrl)) {
    panel.classList.add('hidden');
    cloneBtn.disabled = false;
    return;
  }

  const tgtName = environments.find(e => e.url === targetUrl)?.name ?? targetUrl;
  document.getElementById('connect-target-text').textContent =
    `Cloning to ${tgtName} needs a live connection. Click Connect, then click the EF PPT bookmark in the new tab that opens — it will confirm here automatically once connected.`;
  panel.classList.remove('hidden');
  cloneBtn.disabled = true;
}

async function onConnectTargetClick() {
  const select  = document.getElementById('clone-target');
  const origin  = _originOf(select.value);
  const btn     = document.getElementById('btn-connect-target');
  const panel   = document.getElementById('connect-target-panel');

  btn.disabled = true;
  btn.textContent = 'Connecting…';
  try {
    await _bridgeConnectTarget(origin);
    _connectedTargetOrigin = origin;
    updateConnectPanel();
  } catch (err) {
    panel.classList.remove('hidden');
    panel.classList.remove('alert--warning');
    panel.classList.add('alert--error');
    document.getElementById('connect-target-text').textContent = err.message || 'Could not connect to the target environment.';
  } finally {
    btn.disabled = false;
    btn.textContent = 'Connect Target Environment';
  }
}

function showCloneErrorModal(text) {
  document.getElementById('clone-err-modal-body').textContent = text;
  document.getElementById('clone-err-overlay').classList.remove('hidden');
}

async function onCloneClick() {
  const btn      = document.getElementById('btn-clone');
  const select   = document.getElementById('clone-target');
  const resultEl = document.getElementById('clone-result');
  const targetUrl = select.value;
  const isCross   = _originOf(targetUrl) !== _originOf(currentEnv.url);

  if (btn.disabled) return;
  const origLabel = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = `<span class="mini-spinner"></span> ${isCross ? 'Copying…' : 'Cloning…'}`;
  resultEl.classList.add('hidden');
  resultEl.innerHTML = '';

  try {
    const { newId } = await cloneRecord(currentEnv.url, targetUrl, pageEtn, pageId, entitySetName);
    const newRecordUrl = `${targetUrl}/main.aspx?pagetype=entityrecord&etn=${pageEtn}&id=${newId}`;
    resultEl.innerHTML = `
      <div class="clone-success">
        <svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="12" height="12"><path d="M2 7l3.5 3.5L12 3"/></svg>
        ${isCross ? 'Copied!' : 'Cloned!'}&nbsp;
        <a href="${newRecordUrl}" target="_blank" rel="noopener noreferrer">Open record ↗</a>
      </div>`;
    resultEl.classList.remove('hidden');
  } catch (err) {
    resultEl.innerHTML = `
      <div class="clone-error">
        <span style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;">Error occurred.</span>
        <button type="button" class="clone-error-details">View details</button>
      </div>`;
    resultEl.classList.remove('hidden');
    resultEl.querySelector('.clone-error-details').addEventListener('click', () => showCloneErrorModal(err.message));
  } finally {
    btn.disabled = false;
    btn.innerHTML = origLabel;
  }
}

// ─── Clone engine ───────────────────────────────────────────────────────────────
//
// Same-env: POST a new record, then re-create every M2M association.
// Cross-env: PATCH (upsert) the SAME record ID on the target — self-healing
// against fields the target entity schema doesn't have — then reconcile M2M
// associations exactly. Every request goes through _bridgeFetch/_bridgeRequest;
// launcher.js decides whether that's a direct same-origin fetch or a relay
// through the connected target tab, so this code doesn't need to know which.

const CLONE_EXCLUDE = new Set([
  'createdon', 'modifiedon', 'overriddencreatedon',
  'createdby', 'modifiedby', 'createdonbehalfby', 'modifiedonbehalfby',
  'versionnumber', 'exchangerate', 'importsequencenumber',
  'timezoneruleversionnumber', 'utcconversiontimezonecode',
  'owningbusinessunit', 'owninguser', 'owningteam',
]);
const CLONE_SKIP_TYPES = new Set(['Virtual', 'EntityName', 'ManagedProperty', 'Uniqueidentifier']);

async function cloneRecord(sourceUrl, targetUrl, etn, id, sourceEntitySetName) {
  const isCross        = _originOf(targetUrl) !== _originOf(sourceUrl);
  const apiBase         = `${sourceUrl}/api/data/${settings.apiVersion}`;
  const targetApiBase   = `${targetUrl}/api/data/${settings.apiVersion}`;

  // ── 1. Entity definition — primary name/id attrs, writable attrs, M2M ──────
  const [def, attrData, m2oData] = await Promise.all([
    _bridgeFetch(
      `${apiBase}/EntityDefinitions(LogicalName='${etn}')?$select=PrimaryNameAttribute,PrimaryIdAttribute` +
      `&$expand=ManyToManyRelationships($select=SchemaName,Entity1LogicalName,Entity2LogicalName,Entity1NavigationPropertyName,Entity2NavigationPropertyName)`
    ),
    _bridgeFetch(`${apiBase}/EntityDefinitions(LogicalName='${etn}')/Attributes?$select=LogicalName,AttributeType,IsValidForCreate`),
    _bridgeFetch(`${apiBase}/EntityDefinitions(LogicalName='${etn}')/ManyToOneRelationships?$select=ReferencingAttribute,ReferencingEntityNavigationPropertyName`),
  ]);

  const primaryName   = def.PrimaryNameAttribute;
  const primaryIdAttr = def.PrimaryIdAttribute || `${etn}id`;
  const m2m           = def.ManyToManyRelationships ?? [];
  const attrs         = attrData.value ?? [];
  const navPropMap    = Object.fromEntries((m2oData.value ?? []).map(r => [r.ReferencingAttribute, r.ReferencingEntityNavigationPropertyName]));

  let primaryMaxLen = 100;
  if (primaryName) {
    try {
      const lenD = await _bridgeFetch(
        `${apiBase}/EntityDefinitions(LogicalName='${etn}')/Attributes/Microsoft.Dynamics.CRM.StringAttributeMetadata` +
        `?$select=LogicalName,MaxLength&$filter=LogicalName eq '${primaryName}'`
      );
      primaryMaxLen = lenD.value?.[0]?.MaxLength ?? 100;
    } catch (_) { /* keep default */ }
  }

  // ── 2. Full record (all fields + lookup annotations) ───────────────────────
  const record = await _bridgeFetch(`${apiBase}/${sourceEntitySetName}(${id})`, {
    'Prefer': 'odata.include-annotations="OData.Community.Display.V1.FormattedValue,Microsoft.Dynamics.CRM.lookuplogicalname"',
  });

  // ── 3. Build payload ─────────────────────────────────────────────────────────
  const EXCLUDE = new Set([primaryIdAttr, ...CLONE_EXCLUDE]);
  const scalarPayload = {};
  const lookupFields  = []; // { logicalName, guid, targetEtn }

  for (const attr of attrs) {
    if (!attr.IsValidForCreate) continue;
    const name = attr.LogicalName;
    if (EXCLUDE.has(name)) continue;
    if (CLONE_SKIP_TYPES.has(attr.AttributeType)) continue;

    if (attr.AttributeType === 'Lookup' || attr.AttributeType === 'Customer' || attr.AttributeType === 'Owner') {
      const guidKey   = `_${name}_value`;
      const guid      = record[guidKey];
      if (!guid) continue;
      const targetEtn = record[`${guidKey}@Microsoft.Dynamics.CRM.lookuplogicalname`];
      if (targetEtn) lookupFields.push({ logicalName: name, guid, targetEtn });
    } else {
      const val = record[name];
      if (val === null || val === undefined) continue;
      if (name === primaryName && !isCross) {
        const rawPrefix = (settings.clonePrefix ?? '').trim();
        const prefix    = rawPrefix ? `[${rawPrefix}] ` : '';
        const full      = prefix + String(val);
        scalarPayload[name] = full.length > primaryMaxLen ? full.slice(0, primaryMaxLen) : full;
      } else {
        scalarPayload[name] = val;
      }
    }
  }

  return isCross
    ? _cloneCrossEnv({ etn, id, apiBase, targetApiBase, sourceEntitySetName, scalarPayload, lookupFields, navPropMap, m2m, sourceUrl, targetUrl })
    : _cloneSameEnv({ etn, id, apiBase, sourceEntitySetName, primaryIdAttr, scalarPayload, lookupFields, navPropMap, m2m, sourceUrl });
}

async function _cloneSameEnv({ id, apiBase, sourceEntitySetName, primaryIdAttr, scalarPayload, lookupFields, navPropMap, m2m, sourceUrl }) {
  // Resolve entity set names for all lookup targets (parallel, cached)
  const uniqueTargetEtns = [...new Set(lookupFields.map(f => f.targetEtn))];
  const etnToSet = {};
  await Promise.all(uniqueTargetEtns.map(async t => { etnToSet[t] = await fetchEntitySetName(sourceUrl, t); }));

  const fullPayload = { ...scalarPayload };
  for (const { logicalName, guid, targetEtn } of lookupFields) {
    const targetSet = etnToSet[targetEtn];
    if (!targetSet) continue;
    const navProp = navPropMap[logicalName] ?? logicalName;
    fullPayload[`${navProp}@odata.bind`] = `/${targetSet}(${guid})`;
  }

  // Prefer: return=representation — the bridge doesn't forward response headers
  // (e.g. OData-EntityId), so the new ID must come back in the JSON body instead.
  const createHeaders = { 'Content-Type': 'application/json', 'MSCRM.SuppressDuplicateDetection': 'true', 'Prefer': 'return=representation' };
  const createResult  = await _bridgeRequest(`${apiBase}/${sourceEntitySetName}`, createHeaders, 'POST', fullPayload);
  if (!createResult.ok) {
    const errText = createResult.error || createResult.data?.error?.message || `HTTP ${createResult.status}`;
    throw new Error(`Create failed (${createResult.status}): ${errText}`);
  }
  const newId = createResult.data?.[primaryIdAttr] ?? null;
  if (!newId) throw new Error('Clone created but new record ID could not be determined.');

  // M2M: add all source associations to the new clone
  for (const rel of m2m) {
    await _syncM2mAdd({ rel, apiBase, entitySetName: sourceEntitySetName, sourceUrl, id, newId });
  }

  return { newId, entitySetName: sourceEntitySetName };
}

/** Adds every source-side association of `rel` to the newly cloned record (same env only). */
async function _syncM2mAdd({ rel, apiBase, entitySetName, sourceUrl, id, newId }) {
  // Figure out which side of the relationship our entity is on by checking
  // which nav prop actually resolves records back on the source record.
  const candidates = [
    { navProp: rel.Entity1NavigationPropertyName, relatedEtn: rel.Entity2LogicalName },
    { navProp: rel.Entity2NavigationPropertyName, relatedEtn: rel.Entity1LogicalName },
  ];

  for (const { navProp, relatedEtn } of candidates) {
    if (!navProp || !relatedEtn) continue;
    const relatedEntitySet = await fetchEntitySetName(sourceUrl, relatedEtn);
    if (!relatedEntitySet) continue;
    const pkField = `${relatedEtn}id`;

    let relatedIds = [];
    try {
      const d = await _bridgeFetch(`${apiBase}/${entitySetName}(${id})/${navProp}?$top=500&$select=${pkField}`);
      relatedIds = (d.value ?? []).map(r => r[pkField]).filter(Boolean);
    } catch (e) {
      // This nav prop doesn't apply to our side of the relationship — skip silently.
      continue;
    }

    for (const relatedId of relatedIds) {
      try {
        await _bridgeFetch(`${apiBase}/${entitySetName}(${newId})/${navProp}/$ref`, { 'Content-Type': 'application/json' }, 'POST', {
          '@odata.id': `${apiBase}/${relatedEntitySet}(${relatedId})`,
        });
      } catch (e) {
        console.error('[Record Details] M2M associate failed:', e);
      }
    }
    return; // matched side handled — don't also try the other nav prop
  }
}

async function _cloneCrossEnv({ etn, id, apiBase, targetApiBase, scalarPayload, lookupFields, navPropMap, m2m, sourceUrl, targetUrl }) {
  const sourceEntitySetName = await fetchEntitySetName(sourceUrl, etn);
  const targetEntitySetName = await fetchEntitySetName(targetUrl, etn);
  if (!targetEntitySetName) throw new Error(`Cannot resolve entity set for '${etn}' in target environment.`);

  // Target schema: writable attrs + nav prop map
  let tAttrSet = null, targetNavMap = {};
  try {
    const [attrD, m2oD] = await Promise.all([
      _bridgeFetch(`${targetApiBase}/EntityDefinitions(LogicalName='${etn}')/Attributes?$select=LogicalName,IsValidForCreate`),
      _bridgeFetch(`${targetApiBase}/EntityDefinitions(LogicalName='${etn}')/ManyToOneRelationships?$select=ReferencingAttribute,ReferencingEntityNavigationPropertyName`),
    ]);
    tAttrSet     = new Set((attrD.value ?? []).filter(a => a.IsValidForCreate).map(a => a.LogicalName));
    targetNavMap = Object.fromEntries((m2oD.value ?? []).map(r => [r.ReferencingAttribute, r.ReferencingEntityNavigationPropertyName]));
  } catch (e) {
    console.error('[Record Details] Failed to fetch target schema:', e);
  }

  // Validate scalar fields against target schema
  const validatedScalar = {};
  for (const [k, v] of Object.entries(scalarPayload)) {
    if (!tAttrSet || tAttrSet.has(k)) validatedScalar[k] = v;
  }

  // Rebuild lookup bindings using the target's own entity-set names + nav props
  const validatedLookups = {};
  for (const { logicalName, guid, targetEtn } of lookupFields) {
    if (tAttrSet && !tAttrSet.has(logicalName)) continue;
    const targetLookupSet = await fetchEntitySetName(targetUrl, targetEtn);
    if (!targetLookupSet) continue;
    const navProp = targetNavMap[logicalName] ?? navPropMap[logicalName] ?? logicalName;
    validatedLookups[`${navProp}@odata.bind`] = `/${targetLookupSet}(${guid})`;
  }

  // ── Upsert scalar fields — self-healing retry against fields the target
  // entity schema doesn't have (source-only custom fields, etc). ─────────────
  const writeHeaders = { 'Content-Type': 'application/json', 'MSCRM.SuppressDuplicateDetection': 'true' };
  let remaining  = { ...validatedScalar };
  let lastError  = null;
  for (let attempt = 0; attempt < 30 && Object.keys(remaining).length > 0; attempt++) {
    const r = await _bridgeRequest(`${targetApiBase}/${targetEntitySetName}(${id})`, writeHeaders, 'PATCH', remaining);
    if (r.ok) { remaining = {}; break; }
    const errText = r.error || r.data?.error?.message || `HTTP ${r.status}`;
    lastError = `Upsert failed (${r.status}): ${errText}`;

    const m = errText.match(/[`'"]?([a-z_][a-z0-9_]*)[`'"]?\s+(?:field\s+)?missing\s+from\s+target\s+entity/i)
           ?? errText.match(/missing\s+from\s+target\s+entity[^:]*:\s*([a-z_][a-z0-9_]*)/i);
    const badField = m?.[1];
    if (badField && badField in remaining) { delete remaining[badField]; continue; }
    throw new Error(lastError);
  }
  if (Object.keys(remaining).length > 0 && lastError) throw new Error(lastError);

  // ── PATCH each lookup individually (skip or fail per settings.cloneLookupMode) ──
  const failOnLookup = settings.cloneLookupMode === 'fail';
  for (const [bindKey, bindVal] of Object.entries(validatedLookups)) {
    const r = await _bridgeRequest(`${targetApiBase}/${targetEntitySetName}(${id})`, writeHeaders, 'PATCH', { [bindKey]: bindVal });
    if (!r.ok && failOnLookup) {
      throw new Error(`Lookup PATCH failed (${r.status}): ${r.error || r.data?.error?.message || `HTTP ${r.status}`}`);
    }
  }

  // ── Exact M2M sync in target (add missing, remove extra) ───────────────────
  for (const rel of m2m) {
    await _syncM2mExact({ rel, apiBase, targetApiBase, sourceEntitySetName, targetEntitySetName, sourceUrl, targetUrl, id });
  }

  return { newId: id, entitySetName: targetEntitySetName };
}

async function _syncM2mExact({ rel, apiBase, targetApiBase, sourceEntitySetName, targetEntitySetName, sourceUrl, targetUrl, id }) {
  const candidates = [
    { navProp: rel.Entity1NavigationPropertyName, relatedEtn: rel.Entity2LogicalName },
    { navProp: rel.Entity2NavigationPropertyName, relatedEtn: rel.Entity1LogicalName },
  ];

  for (const { navProp, relatedEtn } of candidates) {
    if (!navProp || !relatedEtn) continue;

    let sourceIds;
    try {
      const pkField = `${relatedEtn}id`;
      const d = await _bridgeFetch(`${apiBase}/${sourceEntitySetName}(${id})/${navProp}?$top=500&$select=${pkField}`);
      sourceIds = (d.value ?? []).map(r => r[pkField]).filter(Boolean);
    } catch (e) {
      continue; // wrong side of the relationship for our entity — try the other nav prop
    }

    const relatedSetSource = await fetchEntitySetName(sourceUrl, relatedEtn);
    const relatedSetTarget = await fetchEntitySetName(targetUrl, relatedEtn);
    if (!relatedSetSource || !relatedSetTarget) return;
    const pkField = `${relatedEtn}id`;

    let targetIds = [];
    try {
      const d = await _bridgeFetch(`${targetApiBase}/${targetEntitySetName}(${id})/${navProp}?$top=500&$select=${pkField}`);
      targetIds = (d.value ?? []).map(r => r[pkField]).filter(Boolean);
    } catch (e) {
      console.error('[Record Details] Failed to fetch target M2M IDs:', e);
    }

    const srcSet = new Set(sourceIds);
    const tgtSet = new Set(targetIds);

    for (const rid of sourceIds) {
      if (tgtSet.has(rid)) continue;
      try {
        await _bridgeFetch(`${targetApiBase}/${targetEntitySetName}(${id})/${navProp}/$ref`, { 'Content-Type': 'application/json' }, 'POST', {
          '@odata.id': `${targetApiBase}/${relatedSetTarget}(${rid})`,
        });
      } catch (e) { console.error('[Record Details] M2M associate failed:', e); }
    }
    for (const rid of targetIds) {
      if (srcSet.has(rid)) continue;
      try {
        await _bridgeFetch(`${targetApiBase}/${targetEntitySetName}(${id})/${navProp}(${rid})/$ref`, {}, 'DELETE');
      } catch (e) { console.error('[Record Details] M2M disassociate failed:', e); }
    }
    return; // matched side handled
  }
}
