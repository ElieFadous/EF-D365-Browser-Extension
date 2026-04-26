/**
 * ribbon-export.js — CSV export modal and download logic.
 */

'use strict';

import { state }                    from './ribbon-state.js';
import { ensureAllRibbonDataLoaded } from './ribbon-load.js';
import { getEntityGroups,
         getButtonsForEntity }       from './ribbon-render.js';
import { describeCondition }         from './ribbon-xml.js';

// ─── Modal open / close ────────────────────────────────────────────────────

export function openExportModal() {
  const groups = getEntityGroups();
  const list   = document.getElementById('modal-entity-list');
  list.innerHTML = '';

  for (const [entity] of groups) {
    const displayName = entity || '(Application Ribbon)';
    const buttons     = getButtonsForEntity(entity);
    const isApp       = entity === '';

    const label = document.createElement('label');
    label.className = `modal-entity-item${isApp ? ' modal-entity-item--app' : ''}`;
    label.innerHTML = `
      <input type="checkbox" value="${esc(entity)}" checked />
      <span class="modal-entity-item__label" title="${esc(displayName)}">${esc(displayName)}</span>
      <span class="modal-entity-item__count">${buttons.length ? buttons.length + ' btn' : 'not loaded'}</span>
    `;
    list.appendChild(label);
  }

  document.getElementById('export-progress').textContent = '';
  document.getElementById('modal-export-btn').disabled = false;
  document.getElementById('export-modal').classList.remove('hidden');
}

export function closeExportModal() {
  document.getElementById('export-modal').classList.add('hidden');
}

export function setAllModalChecks(checked) {
  document.querySelectorAll('#modal-entity-list input[type="checkbox"]')
    .forEach(cb => { cb.checked = checked; });
}

// ─── Export execution ──────────────────────────────────────────────────────

export async function runExport() {
  const selected = [...document.querySelectorAll('#modal-entity-list input[type="checkbox"]:checked')]
    .map(cb => cb.value);

  const progressEl = document.getElementById('export-progress');
  const exportBtn  = document.getElementById('modal-export-btn');

  if (!selected.length) {
    progressEl.textContent = 'Select at least one table.';
    return;
  }

  exportBtn.disabled = true;

  // Ensure all ribbon data is available (CommandDefinitions needed for the CSV).
  if (!state.allRibbonDataLoaded) {
    progressEl.innerHTML = `<span class="spinner"></span> Loading ribbon data…`;
    try {
      await ensureAllRibbonDataLoaded();
    } catch (e) {
      console.error('[EF PPT]', e);
      progressEl.textContent = `Error: ${e.message}`;
      exportBtn.disabled = false;
      return;
    }
  }

  progressEl.textContent = 'Generating CSV…';

  const rows = [[
    'Entity', 'Button Label', 'Button ID', 'Type', 'Location',
    'Command ID', 'Function Name', 'Library', 'Parameters',
    'Display Rules', 'Enable Rules',
  ]];

  for (const entity of selected) {
    const entityLabel = entity || '(Application Ribbon)';
    const buttons = getButtonsForEntity(entity)
      .slice()
      .sort((a, b) => (a.label || a.id).localeCompare(b.label || b.id));

    for (const btn of buttons) {
      const cmd          = state.globalCommandDefs.get(btn.command) ?? null;
      const action       = cmd?.action ?? null;
      const displayLabel = (btn.label && !btn.label.startsWith('$LocLabels:')) ? btn.label : btn.id;

      const paramsText = action?.params.length
        ? action.params.map(p => `[${p.type.replace('Parameter', '')}] ${p.value}`).join('; ')
        : '';

      const ruleText = (refs, map) =>
        refs.map(id => {
          const conds = map.get(id);
          const desc  = conds?.length ? conds.map(c => describeCondition(c)).join(' & ') : '';
          return desc ? `${id}: ${desc}` : id;
        }).join(' | ');

      rows.push([
        entityLabel, displayLabel, btn.id, btn.tagName, btn.location, btn.command,
        action?.functionName ?? '',
        action?.library.replace(/^\$webresource:/, '') ?? '',
        paramsText,
        ruleText(cmd?.displayRuleRefs ?? [], state.globalDisplayRuleDefs),
        ruleText(cmd?.enableRuleRefs  ?? [], state.globalEnableRuleDefs),
      ]);
    }
  }

  _downloadCSV(rows, `ribbon-buttons-${state.envName.replace(/[^a-zA-Z0-9]/g, '_')}.csv`);

  progressEl.textContent = `Exported ${rows.length - 1} row${rows.length - 1 !== 1 ? 's' : ''}.`;
  exportBtn.disabled = false;
}

// ─── CSV download ──────────────────────────────────────────────────────────

function _downloadCSV(rows, filename) {
  const csv = rows.map(row =>
    row.map(cell => {
      const s = String(cell ?? '');
      return (s.includes(',') || s.includes('"') || s.includes('\n'))
        ? `"${s.replace(/"/g, '""')}"` : s;
    }).join(',')
  ).join('\r\n');

  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

// ─── Utility ──────────────────────────────────────────────────────────────

function esc(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
