/**
 * ribbon.js — Entry point for the Ribbon Buttons browser (web-app build).
 *
 * Responsibilities:
 *  • Bootstrap (solutions fetch)
 *  • Solution picker navigation
 *  • Entity selection (toggleEntity, selectAllVisible)
 *  • Wiring all event listeners
 *
 * Heavy lifting is delegated to the feature modules:
 *  ribbon-state.js   — shared mutable state
 *  ribbon-api.js     — D365 Web API calls (via the postMessage fetch bridge)
 *  ribbon-load.js    — ribbon data loading
 *  ribbon-render.js  — all UI rendering
 *  ribbon-export.js  — CSV export
 *  ribbon-xml.js     — XML parsing
 */

'use strict';

import { state, resetLoadState }         from './ribbon-state.js';
import { fetchAllPages }                 from './ribbon-api.js';
import { ensureAllRibbonDataLoaded,
         ensureEntityRibbonLoaded }       from './ribbon-load.js';
import { setEntityClickHandler,
         renderEntityList, renderDetailHeader, renderButtonCards,
         updateSidebarActive, getEntityGroups }  from './ribbon-render.js';
import { openExportModal, closeExportModal,
         setAllModalChecks, runExport }  from './ribbon-export.js';

// ─── Bootstrap ────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {
  const params = new URLSearchParams(location.search);
  state.envUrl  = params.get('env')  ?? '';
  state.envName = params.get('name') ?? 'Unknown';
  document.title = `Ribbon — ${state.envName}`;
  document.getElementById('env-name').textContent = state.envName;
  document.getElementById('env-url').textContent  = state.envUrl.replace('https://', '');

  if (!state.envUrl) {
    showGlobalError('No environment URL provided.\nOpen this page from the EF Power Platform Tools launcher.');
    return;
  }

  // Register the entity-click callback so ribbon-render.js can call back into us.
  setEntityClickHandler(toggleEntity);
  wireEventListeners();

  // Fetch the solution list (network goes through the postMessage bridge).
  setLoadingText('Loading solutions…');
  try {
    state.allSolutions = await fetchAllPages(
      `/solutions?$select=solutionid,uniquename,friendlyname,ismanaged,version` +
      `&$filter=isvisible eq true&$orderby=ismanaged asc,friendlyname asc`
    );
  } catch (e) {
    document.getElementById('env-link').href = state.envUrl;
    console.error('[EF PPT]', e);
    showGlobalError(`Failed to load solutions:\n${e.message}`);
    return;
  }

  // Show solution picker.
  document.getElementById('global-loading').classList.add('hidden');
  renderSolutionList();
  document.getElementById('solution-picker').style.display = 'flex';
});

// ─── Event wiring ─────────────────────────────────────────────────────────

function wireEventListeners() {
  // Solution picker
  document.getElementById('solution-search').addEventListener('input', renderSolutionList);
  document.getElementById('btn-load-all').addEventListener('click', () => loadForSolution(null));

  // Header navigation
  document.getElementById('btn-change-solution').addEventListener('click', returnToSolutionPicker);

  // Sidebar filters & search
  document.getElementById('search-input').addEventListener('input', renderEntityList);
  document.getElementById('filter-managed').addEventListener('change', () => {
    state.selectedEntities.clear();
    updateSidebarActive();
    renderEntityList();
    showEmptyState();
  });

  // Select all / clear selection
  document.getElementById('btn-select-all').addEventListener('click', selectAllVisible);
  document.getElementById('btn-clear-selection').addEventListener('click', () => {
    state.selectedEntities.clear();
    updateSidebarActive();
    showEmptyState();
  });

  // Button search in detail panel
  document.getElementById('btn-search').addEventListener('input', () => {
    if (state.selectedEntities.size > 0) renderButtonCards();
  });

  // Export modal
  document.getElementById('btn-export-csv').addEventListener('click',   openExportModal);
  document.getElementById('modal-close').addEventListener('click',      closeExportModal);
  document.getElementById('modal-cancel-btn').addEventListener('click', closeExportModal);
  document.getElementById('modal-export-btn').addEventListener('click', runExport);
  document.getElementById('select-all-btn').addEventListener('click',  () => setAllModalChecks(true));
  document.getElementById('select-none-btn').addEventListener('click', () => setAllModalChecks(false));
  document.getElementById('export-modal').addEventListener('click', e => {
    if (e.target === e.currentTarget) closeExportModal();
  });

  // XML viewer modal — button-level (button + command sections)
  document.addEventListener('ribbon:view-xml', e => {
    const { label, buttonXml, commandXml } = e.detail;
    openXmlModal(label || 'Button XML', [
      { title: 'Button XML',  xml: buttonXml  },
      { title: 'Command XML', xml: commandXml },
    ]);
  });

  // XML viewer modal — table level (full entity ribbon)
  document.addEventListener('ribbon:view-entity-xml', e => {
    const { label, xml } = e.detail;
    openXmlModal(`${label} — Entity Ribbon XML`, [
      { title: 'Full Ribbon XML', xml },
    ]);
  });
  document.getElementById('xml-modal-close').addEventListener('click', closeXmlModal);
  document.getElementById('xml-modal').addEventListener('click', e => {
    if (e.target === e.currentTarget) closeXmlModal();
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeXmlModal();
  });
}

// ─── Solution picker ───────────────────────────────────────────────────────

function renderSolutionList() {
  const query    = document.getElementById('solution-search').value.trim().toLowerCase();
  const filtered = state.allSolutions.filter(s =>
    !query ||
    (s.friendlyname ?? '').toLowerCase().includes(query) ||
    (s.uniquename   ?? '').toLowerCase().includes(query)
  );

  const list = document.getElementById('solution-list');
  list.innerHTML = '';

  if (!filtered.length) {
    list.innerHTML = '<div class="sol-no-results">No solutions match your search.</div>';
    return;
  }

  filtered.forEach(sol => {
    const item = document.createElement('div');
    item.className = 'solution-item';
    item.innerHTML = `
      <div class="sol-info">
        <div class="sol-name">${esc(sol.friendlyname || sol.uniquename)}</div>
        <div class="sol-meta">${esc(sol.uniquename)} · v${esc(sol.version ?? '?')}</div>
      </div>
      <span class="sol-badge${sol.ismanaged ? '' : ' sol-badge--unmanaged'}">
        ${sol.ismanaged ? 'Managed' : 'Unmanaged'}
      </span>
    `;
    item.addEventListener('click', () => loadForSolution(sol));
    list.appendChild(item);
  });
}

function returnToSolutionPicker() {
  resetLoadState();

  document.getElementById('app-body').style.display         = 'none';
  document.getElementById('btn-change-solution').classList.add('hidden');
  document.getElementById('btn-export-csv').classList.add('hidden');
  document.getElementById('solution-search').value = '';
  renderSolutionList();
  document.getElementById('solution-picker').style.display = 'flex';
}

// ─── Load records for a solution ──────────────────────────────────────────

async function loadForSolution(solution) {
  state.selectedSolution = solution;
  const label = solution ? (solution.friendlyname || solution.uniquename) : 'All Tables';

  document.getElementById('solution-picker').style.display = 'none';
  document.getElementById('global-loading').classList.remove('hidden');

  try {
    setLoadingText('Loading ribbon records…');
    const allRaw = await fetchAllPages(`/ribbondiffs?$select=ribbondiffid,entity,ismanaged`);

    let filtered = allRaw;

    if (solution) {
      setLoadingText(`Getting solution entities for "${label}"…`);
      const [entityComponents, entityDefs] = await Promise.all([
        fetchAllPages(
          `/solutioncomponents?$select=objectid` +
          `&$filter=_solutionid_value eq ${solution.solutionid} and componenttype eq 1`
        ),
        fetchAllPages(`/EntityDefinitions?$select=LogicalName,MetadataId`),
      ]);

      const solutionMetadataIds = new Set(entityComponents.map(c => c.objectid));
      const solutionEntities    = new Set(
        entityDefs
          .filter(e => solutionMetadataIds.has(e.MetadataId))
          .map(e => e.LogicalName)
      );

      filtered = allRaw.filter(r =>
        r.entity === '' || r.entity == null || solutionEntities.has(r.entity)
      );
    }

    state.allRecords = filtered.map(r => ({
      id:        r.ribbondiffid,
      entity:    r.entity ?? '',
      isManaged: r.ismanaged ?? false,
      parsed:    null,
    }));
  } catch (e) {
    showGlobalError(`Failed to load ribbon records:\n${e.message}`);
    return;
  }

  // Show solution badge in sidebar
  const badge  = document.getElementById('sidebar-solution-badge');
  const nameEl = document.getElementById('sidebar-solution-name');
  if (solution) {
    nameEl.textContent  = solution.friendlyname || solution.uniquename;
    badge.style.display = 'block';
  } else {
    badge.style.display = 'none';
  }

  document.getElementById('global-loading').classList.add('hidden');
  document.getElementById('app-body').style.display = 'flex';
  document.getElementById('btn-change-solution').classList.remove('hidden');
  document.getElementById('btn-export-csv').classList.remove('hidden');

  showEmptyState();
  renderEntityList();
}

// ─── Entity selection ──────────────────────────────────────────────────────

async function toggleEntity(entity) {
  // Deselect if already active
  if (state.selectedEntities.has(entity)) {
    state.selectedEntities.delete(entity);
    updateSidebarActive();
    if (state.selectedEntities.size === 0) {
      showEmptyState();
    } else {
      renderDetailHeader();
      renderButtonCards();
    }
    return;
  }

  // Select and load
  state.selectedEntities.add(entity);
  updateSidebarActive();

  const needsLoad = entity !== ''
    ? !state.entityRibbonCache.has(entity)
    : !state.allRibbonDataLoaded;

  if (needsLoad) {
    document.getElementById('empty-state').classList.add('hidden');
    document.getElementById('ribbon-detail').classList.add('hidden');
    document.getElementById('detail-loading').classList.remove('hidden');

    if (entity !== '') {
      await ensureEntityRibbonLoaded(entity);
    } else {
      await ensureAllRibbonDataLoaded();
    }

    document.getElementById('detail-loading').classList.add('hidden');
  }

  document.getElementById('empty-state').classList.add('hidden');
  document.getElementById('ribbon-detail').classList.remove('hidden');

  // Reset the button search box when the first entity is picked
  if (state.selectedEntities.size === 1) {
    document.getElementById('btn-search').value = '';
  }

  renderDetailHeader();
  renderButtonCards();
}

/**
 * Select every entity currently visible in the sidebar (respects search
 * filter and Hide Managed checkbox), load all data if needed, then render.
 */
async function selectAllVisible() {
  const query  = document.getElementById('search-input').value.trim().toLowerCase();
  const groups = getEntityGroups();

  const toLoad = [];
  for (const [entity] of groups) {
    const displayName = entity || '(Application Ribbon)';
    if (query && !displayName.toLowerCase().includes(query)) continue;
    state.selectedEntities.add(entity);
    toLoad.push(entity);
  }

  updateSidebarActive();
  if (state.selectedEntities.size === 0) return;

  // Check if anything still needs fetching.
  const needsFetch = toLoad.some(e =>
    e !== '' ? !state.entityRibbonCache.has(e) : !state.allRibbonDataLoaded
  );

  if (needsFetch) {
    document.getElementById('empty-state').classList.add('hidden');
    document.getElementById('ribbon-detail').classList.add('hidden');
    document.getElementById('detail-loading').classList.remove('hidden');

    // Named entities use compiled ribbon; Application Ribbon uses the old path.
    const namedEntities = toLoad.filter(e => e !== '');
    const hasAppRibbon  = toLoad.includes('');
    await Promise.all([
      ...namedEntities.map(e => ensureEntityRibbonLoaded(e)),
      ...(hasAppRibbon ? [ensureAllRibbonDataLoaded()] : []),
    ]);

    document.getElementById('detail-loading').classList.add('hidden');
  }

  document.getElementById('empty-state').classList.add('hidden');
  document.getElementById('ribbon-detail').classList.remove('hidden');
  document.getElementById('btn-search').value = '';

  renderDetailHeader();
  renderButtonCards();
}

// ─── UI helpers ────────────────────────────────────────────────────────────

export function showEmptyState(msg) {
  if (msg) {
    const sub = document.querySelector('.empty-state__sub');
    if (sub) sub.textContent = msg;
  }
  document.getElementById('empty-state').classList.remove('hidden');
  document.getElementById('ribbon-detail').classList.add('hidden');
  document.getElementById('detail-loading').classList.add('hidden');
  document.getElementById('btn-clear-selection').classList.add('hidden');
  document.getElementById('ribbon-detail-header').innerHTML = '';
  document.getElementById('btn-count').textContent = '';
}

function setLoadingText(text) {
  document.getElementById('loading-text').textContent = text;
}

function showGlobalError(msg) {
  document.getElementById('error-text').textContent = msg;
  document.getElementById('global-loading').classList.add('hidden');
  document.getElementById('global-error').classList.remove('hidden');
}

// ─── XML viewer modal ──────────────────────────────────────────────────────

/**
 * Open the XML viewer modal.
 * @param {string} label   Modal title.
 * @param {Array<{title:string, xml:string}>} sections  One or more XML sections to show.
 */
function openXmlModal(label, sections) {
  document.getElementById('xml-modal-title').textContent = label;

  const body   = document.getElementById('xml-modal-body');
  const footer = document.getElementById('xml-modal-footer');
  body.innerHTML   = '';
  footer.innerHTML = '';

  // Content sections (no copy button inside).
  sections.forEach(({ title, xml }) => body.appendChild(_buildXmlSection(title, xml)));

  // Copy buttons pinned in footer — one per non-empty section.
  const nonEmpty = sections.filter(s => s.xml && s.xml.trim());
  nonEmpty.forEach(({ title, xml }) => {
    const copyBtn = document.createElement('button');
    copyBtn.className   = 'xml-copy-btn';
    copyBtn.textContent = nonEmpty.length > 1 ? `Copy ${title}` : 'Copy';
    const defaultLabel  = copyBtn.textContent;
    copyBtn.addEventListener('click', () => {
      navigator.clipboard.writeText(xml).then(() => {
        copyBtn.textContent = 'Copied!';
        setTimeout(() => { copyBtn.textContent = defaultLabel; }, 1500);
      }).catch(() => {});
    });
    footer.appendChild(copyBtn);
  });

  document.getElementById('xml-modal').classList.remove('hidden');
}

function closeXmlModal() {
  document.getElementById('xml-modal').classList.add('hidden');
}

function _buildXmlSection(title, xml) {
  const section = document.createElement('div');
  section.className = 'xml-section';

  const labelEl = document.createElement('div');
  labelEl.className = 'xml-section-label';
  labelEl.textContent = title;
  section.appendChild(labelEl);

  if (!xml || !xml.trim()) {
    const empty = document.createElement('div');
    empty.className = 'xml-empty';
    empty.textContent = 'Not available';
    section.appendChild(empty);
    return section;
  }

  const pre = document.createElement('pre');
  pre.className = 'xml-pre';
  pre.textContent = _formatXml(xml);
  section.appendChild(pre);

  // Copy button is added to the modal footer by openXmlModal — not here.
  return section;
}

/**
 * Minimal XML pretty-printer.  Adds newlines and indentation without
 * a dependency on DOMParser re-serialisation (which strips comments).
 */
function _formatXml(xml) {
  const INDENT = '  ';
  let depth  = 0;
  let result = '';

  // Normalise whitespace between tags first.
  const normalised = xml.trim().replace(/>\s+</g, '><');

  const tagRe = /(<[^>]+>)|([^<]+)/g;
  let match;
  while ((match = tagRe.exec(normalised)) !== null) {
    const tag  = match[1];
    const text = match[2];

    if (text) {
      result += INDENT.repeat(depth) + text.trim() + '\n';
      continue;
    }

    if (!tag) continue;

    const isClose    = tag.startsWith('</');
    const isSelfClose = tag.endsWith('/>');
    const isPI       = tag.startsWith('<?');
    const isComment  = tag.startsWith('<!--');

    if (isClose)    depth = Math.max(0, depth - 1);

    result += INDENT.repeat(depth) + tag + '\n';

    if (!isClose && !isSelfClose && !isPI && !isComment) {
      depth++;
    }
  }

  return result.trimEnd();
}

// ─── Utility ──────────────────────────────────────────────────────────────

function esc(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
