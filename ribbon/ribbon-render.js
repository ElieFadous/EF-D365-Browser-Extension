/**
 * ribbon-render.js — All UI rendering for the Ribbon Buttons browser.
 *
 * Entity item clicks are communicated back to ribbon.js via a registered
 * callback (setEntityClickHandler) to avoid a circular import.
 */

'use strict';

import { state }                from './ribbon-state.js';
import { describeCondition }    from './ribbon-xml.js';
import { fetchEntityRibbonXml } from './ribbon-load.js';

// ─── Entity-click callback ────────────────────────────────────────────────

let _onEntityClick = null;

/** ribbon.js calls this once during init to wire the toggle handler. */
export function setEntityClickHandler(fn) { _onEntityClick = fn; }

// ─── Sidebar ──────────────────────────────────────────────────────────────

/**
 * Returns a Map of entity → record-count, sorted alphabetically with the
 * Application Ribbon entry pinned at the bottom.  Respects Hide Managed.
 */
export function getEntityGroups() {
  const hideManaged = document.getElementById('filter-managed').checked;
  const groups      = new Map();

  for (const rec of state.allRecords) {
    if (hideManaged && rec.isManaged) continue;
    groups.set(rec.entity, (groups.get(rec.entity) ?? 0) + 1);
  }

  return new Map(
    [...groups.entries()].sort(([a], [b]) => {
      if (a === '' && b !== '') return  1;
      if (b === '' && a !== '') return -1;
      return a.localeCompare(b);
    })
  );
}

export function renderEntityList() {
  const query  = document.getElementById('search-input').value.trim().toLowerCase();
  const groups = getEntityGroups();
  const list   = document.getElementById('entity-list');
  list.innerHTML = '';
  let visible = 0;

  for (const [entity, count] of groups) {
    const displayName = entity || '(Application Ribbon)';
    if (query && !displayName.toLowerCase().includes(query)) continue;
    visible++;

    const li = document.createElement('li');
    li.className = [
      'entity-item',
      entity === '' ? 'entity-item--app-ribbon' : '',
      state.selectedEntities.has(entity) ? 'active' : '',
    ].filter(Boolean).join(' ');
    li.setAttribute('role', 'option');
    li.innerHTML = `
      <span class="entity-item__name" title="${esc(displayName)}">${esc(displayName)}</span>
      <span class="entity-item__count">${count}</span>
    `;
    li.addEventListener('click', () => _onEntityClick?.(entity));
    list.appendChild(li);
  }

  document.getElementById('entity-count').textContent =
    `${visible} table${visible !== 1 ? 's' : ''}`;
}

export function updateSidebarActive() {
  document.querySelectorAll('.entity-item').forEach(li => {
    const nameEl = li.querySelector('.entity-item__name');
    const title  = nameEl?.title ?? '';
    const entity = title === '(Application Ribbon)' ? '' : title;
    li.classList.toggle('active', state.selectedEntities.has(entity));
  });
}

// ─── Detail header ─────────────────────────────────────────────────────────

export function renderDetailHeader() {
  const container = document.getElementById('ribbon-detail-header');
  const clearBtn  = document.getElementById('btn-clear-selection');

  clearBtn.classList.toggle('hidden', state.selectedEntities.size <= 1);

  if (state.selectedEntities.size === 1) {
    const entity      = [...state.selectedEntities][0];
    const displayName = entity || '(Application Ribbon)';
    const buttons     = _getButtonsForAllSelected();
    container.innerHTML = `
      <h2>${esc(displayName)}</h2>
      <span class="detail-sub">${buttons.length} button${buttons.length !== 1 ? 's' : ''}</span>
    `;
  } else {
    const buttons  = _getButtonsForAllSelected();
    const entities = [...state.selectedEntities];
    // Cap pills so the header never squeezes the button list out of view.
    const MAX_PILLS = 5;
    const shown     = entities.slice(0, MAX_PILLS);
    const extra     = entities.length - MAX_PILLS;
    const pillsHtml = shown
      .map(e => `<span class="entity-pill">${esc(e || '(App Ribbon)')}</span>`)
      .join('') + (extra > 0
        ? `<span class="entity-pill entity-pill--more">+${extra} more</span>`
        : '');
    container.innerHTML = `
      <h2>${state.selectedEntities.size} tables &middot; ${buttons.length} button${buttons.length !== 1 ? 's' : ''}</h2>
      <div class="entity-pills" style="margin-top:6px">${pillsHtml}</div>
    `;
  }
}

// ─── Button list ───────────────────────────────────────────────────────────

export function getButtonsForEntity(entity) {
  // Prefer compiled ribbon cache (full data); fall back to per-ribbondiff records.
  const cached = state.entityRibbonCache.get(entity);
  if (cached) return cached.buttons;
  return state.allRecords
    .filter(r => r.entity === entity)
    .flatMap(r => r.parsed?.buttons ?? []);
}

function _getButtonsForAllSelected() {
  const buttons = [];
  for (const entity of state.selectedEntities) {
    // Spread each button so we don't mutate cached objects.
    const eb = getButtonsForEntity(entity).map(b => ({ ...b, _entity: entity }));
    buttons.push(...eb);
  }
  return buttons;
}

export function renderButtonCards() {
  const multiSelect = state.selectedEntities.size > 1;
  const buttons     = _getButtonsForAllSelected();
  const query       = document.getElementById('btn-search').value.trim().toLowerCase();

  const filtered = query
    ? buttons.filter(b =>
        b.label.toLowerCase().includes(query) ||
        b.id.toLowerCase().includes(query)    ||
        b.command.toLowerCase().includes(query) ||
        (state.globalCommandDefs.get(b.command)?.action?.functionName ?? '').toLowerCase().includes(query)
      )
    : buttons;

  const sorted = filtered.slice().sort((a, b) => {
    if (multiSelect) {
      const ea = a._entity || '';
      const eb = b._entity || '';
      if (ea !== eb) return ea.localeCompare(eb);
    }
    return (a.label || a.id).localeCompare(b.label || b.id);
  });

  document.getElementById('btn-count').textContent =
    filtered.length === buttons.length
      ? `${buttons.length} button${buttons.length !== 1 ? 's' : ''}`
      : `${filtered.length} of ${buttons.length}`;

  const container = document.getElementById('button-list');
  container.innerHTML = '';

  if (!sorted.length) {
    container.innerHTML = buttons.length
      ? '<div class="no-results">No buttons match the filter.</div>'
      : '<div class="no-results">No custom ribbon buttons found for this table.</div>';
    return;
  }

  // Always group by entity.  Single selection = one group, expanded by default.
  const groups = new Map();
  for (const btn of sorted) {
    const key = btn._entity ?? '';
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(btn);
  }
  for (const [entityKey, groupBtns] of groups) {
    container.appendChild(_buildGroupSection(entityKey, groupBtns, !multiSelect));
  }
}

function _buildGroupSection(entity, buttons, startExpanded = false) {
  const displayName = entity || '(Application Ribbon)';

  const section = document.createElement('div');
  section.className = 'btn-group';

  const header = document.createElement('div');
  header.className = 'btn-group-header';
  header.innerHTML = `
    <span class="btn-group-toggle">${startExpanded ? '&#9660;' : '&#9654;'}</span>
    <span class="btn-group-entity">${esc(displayName)}</span>
    <span class="btn-group-count">${buttons.length} button${buttons.length !== 1 ? 's' : ''}</span>
  `;

  const body = document.createElement('div');
  body.className = `btn-group-body${startExpanded ? '' : ' hidden'}`;
  buttons.forEach(btn => body.appendChild(buildButtonCard(btn)));

  // XML button — real entities only (Application Ribbon has no entity name).
  if (entity) {
    const xmlBtn = document.createElement('button');
    xmlBtn.className = 'btn-entity-xml';
    xmlBtn.textContent = '</> XML';
    xmlBtn.title = 'Fetch full compiled ribbon XML for this table (all locations)';

    xmlBtn.addEventListener('click', async e => {
      e.stopPropagation(); // don't trigger the group collapse/expand
      xmlBtn.disabled    = true;
      xmlBtn.textContent = 'Loading\u2026';
      try {
        const xml = await fetchEntityRibbonXml(entity);
        document.dispatchEvent(new CustomEvent('ribbon:view-entity-xml', {
          detail: { label: displayName, xml }
        }));
      } catch (err) {
        console.error('[EF PPT]', err);
        document.dispatchEvent(new CustomEvent('ribbon:view-entity-xml', {
          detail: { label: displayName, xml: `\u26a0 ${err.message}` }
        }));
      } finally {
        xmlBtn.disabled    = false;
        xmlBtn.textContent = '</> XML';
      }
    });

    header.appendChild(xmlBtn);
  }

  header.addEventListener('click', () => {
    const collapsed = body.classList.toggle('hidden');
    header.querySelector('.btn-group-toggle').innerHTML = collapsed ? '&#9654;' : '&#9660;';
  });

  section.appendChild(header);
  section.appendChild(body);
  return section;
}

export function buildButtonCard(button) {
  const cmd = state.globalCommandDefs.get(button.command) ?? null;

  let actionHtml = '';
  if (cmd?.action) {
    const { functionName, library, params } = cmd.action;
    const cleanLib   = library.replace(/^\$webresource:/, '');
    const paramsHtml = params.length
      ? `<ul class="param-list">${params.map(p => {
          const short = p.type.replace('Parameter', '');
          return `<li><span class="param-badge param-badge--${esc(short)}">${esc(short)}</span>${esc(p.value)}</li>`;
        }).join('')}</ul>`
      : '';
    actionHtml = `
      <div class="card-section">
        <div class="card-section-title">Action</div>
        <div class="fn-name">${esc(functionName)}</div>
        ${cleanLib ? `<div class="fn-library">${esc(cleanLib)}</div>` : ''}
        ${paramsHtml}
      </div>`;
  } else if (cmd) {
    actionHtml = `
      <div class="card-section">
        <div class="card-section-title">Action</div>
        <span class="no-action">No JavaScript action configured</span>
      </div>`;
  } else if (button.command) {
    actionHtml = `
      <div class="card-section">
        <div class="card-section-title">Action</div>
        <span class="no-action">Command definition not found (ID: <code>${esc(button.command)}</code>)</span>
      </div>`;
  }

  const displayRulesHtml = _buildRulesSectionHtml(
    'Display Rules', cmd?.displayRuleRefs ?? [], state.globalDisplayRuleDefs
  );
  const enableRulesHtml = _buildRulesSectionHtml(
    'Enable Rules', cmd?.enableRuleRefs ?? [], state.globalEnableRuleDefs
  );
  const locationHtml = button.location
    ? `<div class="btn-card-location">${esc(button.location)}</div>`
    : '';

  const displayLabel   = (button.label && !button.label.startsWith('$LocLabels:')) ? button.label : button.id;
  const showSeparateId = displayLabel !== button.id;

  const div = document.createElement('div');
  div.className = 'btn-card';
  div.innerHTML = `
    <div class="btn-card-header">
      <div>
        <div class="btn-card-label">${esc(displayLabel)}</div>
        ${showSeparateId ? `<div class="btn-card-id">${esc(button.id)}</div>` : ''}
      </div>
      <div style="display:flex;align-items:center;gap:8px;flex-shrink:0">
        <span class="btn-type-badge">${esc(button.tagName)}</span>
      </div>
    </div>
    ${button.command ? `<div class="btn-card-command">Command: <code>${esc(button.command)}</code></div>` : ''}
    ${locationHtml}
    ${actionHtml}
    ${displayRulesHtml}
    ${enableRulesHtml}
  `;

  // "View XML" button — appended programmatically so the click handler can close
  // over the button object without serialising large XML into data attributes.
  const xmlBtn = document.createElement('button');
  xmlBtn.className = 'btn-view-xml';
  xmlBtn.textContent = '</> XML';
  xmlBtn.title = 'View raw XML for this button and its command definition';
  xmlBtn.addEventListener('click', e => {
    e.stopPropagation();
    document.dispatchEvent(new CustomEvent('ribbon:view-xml', {
      detail: {
        label:      displayLabel,
        buttonXml:  button._sourceXml ?? '',
        commandXml: state.globalCommandDefs.get(button.command)?.rawXml ?? '',
      }
    }));
  });
  // Insert the button into the card header's right side.
  div.querySelector('.btn-card-header > div:last-child').prepend(xmlBtn);

  return div;
}

function _buildRulesSectionHtml(title, ruleRefs, defsMap) {
  if (!ruleRefs.length) return '';
  const rulesHtml = ruleRefs.map(id => {
    const conditions = defsMap.get(id) ?? null;
    const condHtml   = conditions?.length
      ? conditions.map(c => `<li class="rule-condition">${esc(describeCondition(c))}</li>`).join('')
      : conditions !== null
        ? '<li class="rule-condition rule-condition--external">No conditions defined</li>'
        : '<li class="rule-condition rule-condition--external">Definition in another record</li>';
    return `<li class="rule-item"><span class="rule-id">${esc(id)}</span><ul class="rule-conditions">${condHtml}</ul></li>`;
  }).join('');
  return `
    <div class="card-section">
      <div class="card-section-title">${esc(title)}</div>
      <ul class="rule-list">${rulesHtml}</ul>
    </div>`;
}

// ─── Utility ──────────────────────────────────────────────────────────────

function esc(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
