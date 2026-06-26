/**
 * ribbon-state.js — Shared mutable state for the Ribbon Buttons browser.
 *
 * All modules import `state` and read/write its properties directly.
 * Using a single object reference means changes are always visible to all
 * importers without re-exporting primitive values.
 */

'use strict';

export const state = {
  envUrl:           '',
  envName:          '',

  /** Full list of visible solutions fetched on startup. */
  allSolutions:     [],
  selectedSolution: null,

  /**
   * Lightweight ribbondiff records for the chosen solution.
   * Each: { id, entity, isManaged, parsed }
   * `parsed` is null until loadRibbonData() has run.
   */
  allRecords:       [],
  xmlFieldName:     null,   // discovered at runtime (often 'rdx')

  /** Global maps populated once allRibbonDataLoaded = true. */
  globalCommandDefs:     new Map(),
  globalDisplayRuleDefs: new Map(),
  globalEnableRuleDefs:  new Map(),

  /**
   * Per-entity compiled ribbon cache populated by ensureEntityRibbonLoaded().
   * Key: entity logical name.  Value: { buttons, commandDefs, displayRuleDefs, enableRuleDefs }
   */
  entityRibbonCache: new Map(),

  /** Set of currently selected entity logical names ('' = Application Ribbon). */
  selectedEntities:    new Set(),

  /**
   * True after all ribbondiff records have been fetched and parsed.
   * Subsequent entity selections skip the API call.
   */
  allRibbonDataLoaded: false,
};

/** Reset everything that depends on a particular solution/environment load. */
export function resetLoadState() {
  state.allRecords           = [];
  state.xmlFieldName         = null;
  state.allRibbonDataLoaded  = false;
  state.entityRibbonCache.clear();
  state.selectedEntities.clear();
  state.globalCommandDefs.clear();
  state.globalDisplayRuleDefs.clear();
  state.globalEnableRuleDefs.clear();
}
