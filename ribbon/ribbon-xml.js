/**
 * ribbon-xml.js — Ribbon XML parsing.
 *
 * D365 ribbondiff records store ONE of two formats in their `rdx` field:
 *
 *   1. A bare fragment — a single XML element that is itself the thing being
 *      customised (e.g. <CustomAction>, <CommandDefinition>, <DisplayRule>).
 *      Detected when the root tag is in RDX_FRAGMENT_ROOTS.
 *
 *   2. A full RibbonDiffXml document (occasionally used by older tooling or
 *      the RetrieveEntityRibbon compiled-ribbon endpoint).
 *
 * Both formats are handled by parseRibbonXml().
 */

'use strict';

import { state } from './ribbon-state.js';

// ─── XML field discovery ───────────────────────────────────────────────────

/**
 * Scan a D365 OData record for the field that contains XML ribbon data.
 * The field is typically named `rdx` but varies across D365 versions.
 */
export function detectXmlField(obj) {
  // Check well-known candidates first (fast path)
  for (const key of ['rdx', 'ribbondiffxml', 'content', 'diffxml', 'xml', 'customizationxml']) {
    if (typeof obj[key] === 'string' && obj[key].trimStart().startsWith('<')) return key;
  }
  // Fallback: scan every string field
  for (const [key, val] of Object.entries(obj)) {
    if (typeof val === 'string' && val.trimStart().startsWith('<')) return key;
  }
  return null;
}

// ─── Global map management ────────────────────────────────────────────────

export function mergeIntoGlobalMaps(parsed) {
  parsed.commandDefs.forEach((def, id)   => state.globalCommandDefs.set(id, def));
  parsed.displayRuleDefs.forEach((c, id) => state.globalDisplayRuleDefs.set(id, c));
  parsed.enableRuleDefs.forEach((c, id)  => state.globalEnableRuleDefs.set(id, c));
}

export function emptyParsed() {
  return {
    buttons:         [],
    commandDefs:     new Map(),
    displayRuleDefs: new Map(),
    enableRuleDefs:  new Map(),
  };
}

// ─── Root fragment tags ────────────────────────────────────────────────────

const RDX_FRAGMENT_ROOTS = new Set([
  'CustomAction', 'HideCustomAction', 'CommandDefinition',
  'DisplayRule',  'EnableRule',        'LocLabel',
]);

// ─── Main entry point ──────────────────────────────────────────────────────

/**
 * Parse a ribbon XML string and return a `{ buttons, commandDefs,
 * displayRuleDefs, enableRuleDefs }` object.
 *
 * @param {string}  xmlStr         Raw XML string from the rdx field.
 * @param {boolean} filterStandard When true, Microsoft built-in IDs
 *                                 (Mscrm.*, _Default.*, etc.) are excluded.
 */
export function parseRibbonXml(xmlStr, filterStandard = false) {
  const parser = new DOMParser();
  const doc    = parser.parseFromString(xmlStr, 'text/xml');
  if (doc.querySelector('parsererror')) return emptyParsed();

  const root = doc.documentElement;

  if (RDX_FRAGMENT_ROOTS.has(root.tagName)) {
    return _parseFragment(root);
  }

  return _parseFullDoc(doc, filterStandard);
}

// ─── Fragment parser (one bare element per ribbondiff record) ──────────────

function _parseFragment(el) {
  const buttons         = [];
  const commandDefs     = new Map();
  const displayRuleDefs = new Map();
  const enableRuleDefs  = new Map();

  switch (el.tagName) {
    case 'CustomAction': {
      const location = el.getAttribute('Location') ?? '';
      el.querySelectorAll('CommandUIDefinition').forEach(cud => {
        _extractButtons(cud, location, buttons, () => false);
      });
      // Some records embed an inline CommandDefinition inside the CustomAction.
      el.querySelectorAll('CommandDefinition').forEach(cd => {
        const id = cd.getAttribute('Id') ?? '';
        if (id) commandDefs.set(id, _parseCommandDef(cd));
      });
      break;
    }
    case 'CommandDefinition': {
      const id = el.getAttribute('Id') ?? '';
      if (id) commandDefs.set(id, _parseCommandDef(el));
      break;
    }
    case 'DisplayRule': {
      const id = el.getAttribute('Id') ?? '';
      if (id) displayRuleDefs.set(id, _parseConditions(el));
      break;
    }
    case 'EnableRule': {
      const id = el.getAttribute('Id') ?? '';
      if (id) enableRuleDefs.set(id, _parseConditions(el));
      break;
    }
    // HideCustomAction, LocLabel — nothing to extract for button listing.
  }

  return { buttons, commandDefs, displayRuleDefs, enableRuleDefs };
}

// ─── Full-document parser (RibbonDiffXml or RetrieveEntityRibbon output) ───

function _parseFullDoc(doc, filterStandard) {
  const isStd = id =>
    filterStandard && /^(Mscrm\.|_Default\.|Mscrm_|Microsoft\.)/i.test(id);

  const buttons = [];

  // Use descendant selector so different wrapper depths are all handled.
  doc.querySelectorAll('CustomAction').forEach(ca => {
    const location = ca.getAttribute('Location') ?? '';
    ca.querySelectorAll('CommandUIDefinition').forEach(cud => {
      _extractButtons(cud, location, buttons, isStd);
    });
  });

  const commandDefs = new Map();
  doc.querySelectorAll('CommandDefinition').forEach(cd => {
    const id = cd.getAttribute('Id') ?? '';
    if (!id || isStd(id)) return;
    commandDefs.set(id, _parseCommandDef(cd));
  });

  const displayRuleDefs = new Map();
  // Only parse rule DEFINITIONS (they have condition children).
  // Rule REFERENCES inside <CommandDefinition><DisplayRules> are self-closing.
  doc.querySelectorAll('DisplayRule').forEach(rule => {
    if (!rule.children.length) return; // reference node, skip
    const id = rule.getAttribute('Id') ?? '';
    if (id) displayRuleDefs.set(id, _parseConditions(rule));
  });

  const enableRuleDefs = new Map();
  doc.querySelectorAll('EnableRule').forEach(rule => {
    if (!rule.children.length) return;
    const id = rule.getAttribute('Id') ?? '';
    if (id) enableRuleDefs.set(id, _parseConditions(rule));
  });

  return { buttons, commandDefs, displayRuleDefs, enableRuleDefs };
}

// ─── Shared helpers ────────────────────────────────────────────────────────

/** Serialise a parsed DOM element back to an XML string. */
function _serializeEl(el) {
  try {
    return new XMLSerializer().serializeToString(el);
  } catch {
    return '';
  }
}

function _extractButtons(cud, location, buttons, isStd = () => false) {
  ['Button', 'SplitButton', 'FlyoutAnchor'].forEach(tag => {
    cud.querySelectorAll(tag).forEach(el => {
      const id = el.getAttribute('Id') ?? '';
      if (id && !isStd(id)) {
        buttons.push({
          id,
          label:      el.getAttribute('LabelText') ?? '',
          command:    el.getAttribute('Command')   ?? '',
          location,
          tagName:    tag,
          _sourceXml: _serializeEl(el),   // serialised from the live DOM element
        });
      }
    });
  });
}

function _parseCommandDef(el) {
  const jsFn = el.querySelector('Actions > JavaScriptFunction');
  return {
    id:     el.getAttribute('Id') ?? '',
    rawXml: _serializeEl(el),            // serialised from the live DOM element
    action: jsFn ? {
      functionName: jsFn.getAttribute('FunctionName') ?? '',
      library:      jsFn.getAttribute('Library')      ?? '',
      params: Array.from(jsFn.children).map(p => ({
        type:  p.tagName,
        value: p.getAttribute('Value') ?? '',
      })),
    } : null,
    displayRuleRefs: _extractRuleIds(el, 'DisplayRules > DisplayRule'),
    enableRuleRefs:  _extractRuleIds(el, 'EnableRules > EnableRule'),
  };
}

function _extractRuleIds(parentEl, selector) {
  return Array.from(parentEl.querySelectorAll(selector))
    .map(el => el.getAttribute('Id') ?? '')
    .filter(Boolean);
}

function _parseConditions(ruleEl) {
  return Array.from(ruleEl.children).map(child => {
    const attrs = {};
    Array.from(child.attributes).forEach(a => { attrs[a.name] = a.value; });
    return { type: child.tagName, attrs };
  });
}

// ─── Compiled ribbon parser (RetrieveEntityRibbon output) ─────────────────

/**
 * Parse a full compiled ribbon returned by RetrieveEntityRibbon.
 *
 * Structure:  <RibbonDefinitions> → <RibbonDefinition> → <UI> → <Ribbon>
 *               → <Tabs> → <Tab> → <Groups> → <Group> → <Controls>
 *
 * Buttons are direct children of <Group><Controls>, NOT wrapped in a
 * <CustomAction>.  CommandDefinitions and RuleDefinitions are kept without
 * the Mscrm.* filter so that rules referenced by custom buttons are always
 * resolvable even when the command itself is a standard one.
 */
function _parseCompiledRibbon(doc) {
  const isStdBtn = id => /^(Mscrm\.|_Default\.|Mscrm_|Microsoft\.)/i.test(id);

  const buttons = [];

  // Select only direct children of Group>Controls to skip nested
  // Menu > MenuSection > Controls entries.
  doc.querySelectorAll('Group > Controls').forEach(controls => {
    const tab      = controls.closest('Tab');
    const location = tab?.getAttribute('Id') ?? '';
    Array.from(controls.children).forEach(el => {
      if (!['Button', 'SplitButton', 'FlyoutAnchor'].includes(el.tagName)) return;
      const id = el.getAttribute('Id') ?? '';
      if (!id || isStdBtn(id)) return;
      buttons.push({
        id,
        label:      el.getAttribute('LabelText') ?? '',
        command:    el.getAttribute('Command')   ?? '',
        location,
        tagName:    el.tagName,
        _sourceXml: _serializeEl(el),
      });
    });
  });

  // All CommandDefinitions — including Mscrm.* — so display/enable rules on
  // custom buttons can always be resolved.
  const commandDefs = new Map();
  doc.querySelectorAll('CommandDefinition').forEach(cd => {
    const id = cd.getAttribute('Id') ?? '';
    if (id) commandDefs.set(id, _parseCommandDef(cd));
  });

  const displayRuleDefs = new Map();
  doc.querySelectorAll('DisplayRule').forEach(rule => {
    if (!rule.children.length) return;
    const id = rule.getAttribute('Id') ?? '';
    if (id) displayRuleDefs.set(id, _parseConditions(rule));
  });

  const enableRuleDefs = new Map();
  doc.querySelectorAll('EnableRule').forEach(rule => {
    if (!rule.children.length) return;
    const id = rule.getAttribute('Id') ?? '';
    if (id) enableRuleDefs.set(id, _parseConditions(rule));
  });

  return { buttons, commandDefs, displayRuleDefs, enableRuleDefs };
}

/**
 * Public entry point for the compiled ribbon format returned by
 * RetrieveEntityRibbon (after ZIP decompression).
 */
export function parseCompiledRibbonXml(xmlStr) {
  const parser = new DOMParser();
  const doc    = parser.parseFromString(xmlStr, 'text/xml');
  if (doc.querySelector('parsererror')) return emptyParsed();
  return _parseCompiledRibbon(doc);
}

// ─── Condition descriptions ────────────────────────────────────────────────

export function describeCondition(cond) {
  const a = cond.attrs;
  switch (cond.type) {
    case 'EntityRule':              return `Table: ${a.EntityName || '(any)'}${a.AppliesTo ? ` — applies to: ${a.AppliesTo}` : ''}`;
    case 'SelectionCountRule':      return `Selection count: ${a.Minimum ?? '0'} – ${(!a.Maximum || a.Maximum === '0') ? '∞' : a.Maximum}`;
    case 'FormStateRule':           return `Form state: ${a.State ?? '?'}`;
    case 'EntityPrivilegeRule':     return `Privilege: ${a.PrivilegeType ?? '?'} on ${a.EntityName || '(entity)'}`;
    case 'CrmClientTypeRule':       return `Client type: ${a.Type ?? '?'}`;
    case 'SkuRule':                 return `SKU: ${a.Sku ?? '?'}`;
    case 'PageRule':                return `Page: ${a.AppliesTo ?? '?'}`;
    case 'RecordPrivilegeRule':     return `Record privilege: ${a.PrivilegeType ?? '?'}`;
    case 'MiscellaneousPrivilegeRule': return `Misc. privilege: ${a.PrivilegeName ?? '?'}`;
    case 'OutlookRenderTypeRule':   return `Outlook render type: ${a.Type ?? '?'}`;
    case 'OrRule':                  return `OR (combined conditions)`;
    case 'ValueRule':               return `Value: ${a.Field ?? ''} = ${a.Value ?? ''}`;
    case 'ReferencingAttributeRequiredRule': return `Referencing attribute is required`;
    case 'CommandClientTypeRule':   return `Command client type: ${a.Type ?? '?'}`;
    case 'HideForTabletExperienceRule': return `Hide for tablet experience`;
    default: {
      const pairs = Object.entries(a).map(([k, v]) => `${k}=${v}`).join(', ');
      return `${cond.type}${pairs ? `: ${pairs}` : ''}`;
    }
  }
}
