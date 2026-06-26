/**
 * ribbon-load.js — Ribbon data loading (web-app build).
 *
 * All network calls go through the postMessage fetch bridge (see ribbon-api.js).
 * Per-entity OData queries (?$filter=entity eq 'X') are used because:
 *
 *  • Per-entity queries are confirmed to return the `rdx` XML field in D365
 *    responses.
 *  • Unfiltered /ribbondiffs requests do NOT return `rdx` by default (D365
 *    omits large text columns from broad collection responses).
 *
 * Compiled ribbon XML (RetrieveEntityRibbon) is ZIP/deflate-compressed; the
 * iframe decompresses it locally with DecompressionStream.
 */

'use strict';

import { state }                                         from './ribbon-state.js';
import { d365Fetch, d365FetchSingle, d365FetchText,
         fetchAllPages }                                  from './ribbon-api.js';
import { detectXmlField, mergeIntoGlobalMaps,
         parseRibbonXml, parseCompiledRibbonXml,
         emptyParsed }                                    from './ribbon-xml.js';

// ─── Primary loader ────────────────────────────────────────────────────────

export async function ensureAllRibbonDataLoaded() {
  if (state.allRibbonDataLoaded) return;

  // Collect distinct entity values we need (tracked records + app ribbon).
  const entitySet = new Set(state.allRecords.map(r => r.entity));
  entitySet.add('');   // always fetch Application Ribbon records too
  const entities = [...entitySet];

  const rows = await _fetchEntitiesParallel(entities);

  // Discover XML field name.
  if (!state.xmlFieldName) {
    for (const row of rows) {
      const found = detectXmlField(row);
      if (found) { state.xmlFieldName = found; break; }
    }
  }

  const byId = new Map(rows.map(r => [r.ribbondiffid, r]));

  // Parse every tracked record.
  for (const rec of state.allRecords) {
    if (rec.parsed !== null) continue;
    const row    = byId.get(rec.id);
    const xmlStr = row && state.xmlFieldName ? (row[state.xmlFieldName] ?? null) : null;
    rec.parsed   = xmlStr ? parseRibbonXml(xmlStr) : emptyParsed();
    mergeIntoGlobalMaps(rec.parsed);
  }

  // Merge untracked rows — these contain CommandDefinitions, DisplayRules and
  // EnableRules that may be referenced by the tracked buttons.
  const trackedIds = new Set(state.allRecords.map(r => r.id));
  for (const row of rows) {
    if (trackedIds.has(row.ribbondiffid)) continue;
    const xmlStr = state.xmlFieldName ? (row[state.xmlFieldName] ?? null) : null;
    if (xmlStr) mergeIntoGlobalMaps(parseRibbonXml(xmlStr));
  }

  state.allRibbonDataLoaded = true;
}

// ─── Per-entity compiled ribbon loader ────────────────────────────────────

/**
 * Fetch the full compiled ribbon for one named entity via RetrieveEntityRibbon,
 * parse it, and cache the result.  Subsequent calls for the same entity return
 * immediately from the cache.
 *
 * Uses the compiled ribbon XML (ZIP-compressed) rather than per-ribbondiff OData
 * records, which gives complete CommandDefinitions and RuleDefinitions in a
 * single request.
 */
export async function ensureEntityRibbonLoaded(entity) {
  if (state.entityRibbonCache.has(entity)) return;

  const xml    = await fetchEntityRibbonXml(entity);
  const parsed = parseCompiledRibbonXml(xml);
  mergeIntoGlobalMaps(parsed);
  state.entityRibbonCache.set(entity, parsed);
}

// ─── Parallel per-entity fetch (via bridge) ────────────────────────────────

/**
 * Fire all entity-scoped OData queries in parallel through the fetch bridge.
 * Per-entity queries reliably return the `rdx` XML field even when the
 * unfiltered /ribbondiffs endpoint omits it.
 */
async function _fetchEntitiesParallel(entities) {
  async function fetchEntityRows(entity) {
    // Build filter.  '' / null both mean Application Ribbon.
    const filter = (entity === '' || entity == null)
      ? `entity eq null or entity eq ''`
      : `entity eq '${entity}'`;
    return fetchAllPages(`/ribbondiffs?$filter=${encodeURIComponent(filter)}`);
  }

  const perEntity = await Promise.all(
    entities.map(e => fetchEntityRows(e).catch(() => []))
  );

  // Deduplicate by ribbondiffid (same record may appear under multiple entities).
  const seen = new Set();
  return perEntity.flat().filter(r => {
    if (seen.has(r.ribbondiffid)) return false;
    seen.add(r.ribbondiffid);
    return true;
  });
}

// ─── Per-record fallback ───────────────────────────────────────────────────

/**
 * Four-approach fallback for loading a single record's XML when the batch
 * approach is unavailable (e.g. during export before full load).
 */
export async function loadRecordXml(rec) {
  if (rec.parsed !== null) return;
  let xmlStr = null;

  // Approach 1: single-record GET (blocked on some D365 instances — HTTP 400).
  try {
    const full = await d365FetchSingle(`${state.envUrl}/api/data/v9.2/ribbondiffs(${rec.id})`);
    if (!state.xmlFieldName) state.xmlFieldName = detectXmlField(full);
    xmlStr = state.xmlFieldName ? (full[state.xmlFieldName] ?? null) : null;
  } catch (_) { /* fall through */ }

  // Approach 2: collection filter (RetrieveMultiple — always supported).
  if (!xmlStr) {
    try {
      const rows = await fetchAllPages(`/ribbondiffs?$filter=ribbondiffid eq ${rec.id}`);
      if (rows.length) {
        if (!state.xmlFieldName) state.xmlFieldName = detectXmlField(rows[0]);
        xmlStr = state.xmlFieldName ? (rows[0][state.xmlFieldName] ?? null) : null;
      }
    } catch (_) { /* fall through */ }
  }

  // Approach 3: $value stream.
  if (!xmlStr) {
    try {
      const txt = await d365FetchText(
        `${state.envUrl}/api/data/v9.2/ribbondiffs(${rec.id})/ribbondiffxml/$value`
      );
      if (txt?.trimStart().startsWith('<')) xmlStr = txt;
    } catch (_) { /* fall through */ }
  }

  // Approach 4: compiled ribbon via RetrieveEntityRibbon (entity-specific only).
  if (!xmlStr && rec.entity !== '') {
    try { xmlStr = await _fetchEntityRibbonXml(rec.entity); } catch (_) { console.error('[EF PPT]', 'Entity ribbon fetch failed:', _); }
  }

  rec.parsed = xmlStr ? parseRibbonXml(xmlStr, /* filterStandard= */ true) : emptyParsed();
  mergeIntoGlobalMaps(rec.parsed);
}

// ─── RetrieveEntityRibbon helpers ─────────────────────────────────────────

/**
 * Public: fetch the full compiled ribbon XML for an entity using the 'All'
 * location filter.  The JSON (with base64 CompressedEntityXml) is fetched via
 * the bridge, then decompressed locally in this iframe.
 */
export async function fetchEntityRibbonXml(entityLogicalName) {
  const url =
    `${state.envUrl}/api/data/v9.2/RetrieveEntityRibbon` +
    `(EntityName='${encodeURIComponent(entityLogicalName)}',` +
    `RibbonLocationFilter=Microsoft.Dynamics.CRM.RibbonLocationFilters'All')`;

  const data = await d365FetchSingle(url);
  if (!data?.CompressedEntityXml) throw new Error('No CompressedEntityXml in response');
  return _decompressXml(data.CompressedEntityXml);
}

/** Private fallback (used by loadRecordXml, filter=0 for speed). */
async function _fetchEntityRibbonXml(entityLogicalName) {
  const result = await d365FetchSingle(
    `${state.envUrl}/api/data/v9.2/RetrieveEntityRibbon` +
    `(EntityName='${encodeURIComponent(entityLogicalName)}',RibbonLocationFilter=0)`
  );
  if (!result?.CompressedEntityXml) throw new Error('No CompressedEntityXml in response');
  return _decompressXml(result.CompressedEntityXml);
}

async function _decompressXml(base64) {
  const b64   = base64.replace(/\s/g, '');
  const bin   = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);

  // PKZIP archive (D365 default format for CompressedEntityXml).
  const sig = (bytes[0] << 24 | bytes[1] << 16 | bytes[2] << 8 | bytes[3]) >>> 0;
  if (sig === 0x504b0304) {
    const view           = new DataView(bytes.buffer, bytes.byteOffset);
    const method         = view.getUint16(8,  true);
    const compressedSize = view.getUint32(18, true);
    const fileNameLen    = view.getUint16(26, true);
    const extraLen       = view.getUint16(28, true);
    const dataStart      = 30 + fileNameLen + extraLen;
    const payload        = bytes.slice(dataStart,
      compressedSize ? dataStart + compressedSize : undefined);

    if (method === 0) return new TextDecoder('utf-8').decode(payload);
    if (method === 8) {
      const buf = await new Response(
        new Response(payload).body.pipeThrough(new DecompressionStream('deflate-raw'))
      ).arrayBuffer();
      return new TextDecoder('utf-8').decode(buf);
    }
    throw new Error(`Unsupported ZIP compression method: ${method}`);
  }

  // Plain XML fallback.
  const asText = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
  if (asText.trimStart().startsWith('<')) return asText;

  // gzip / zlib / raw deflate fallback.
  for (const fmt of ['gzip', 'deflate', 'deflate-raw']) {
    try {
      const buf = await new Response(
        new Response(bytes).body.pipeThrough(new DecompressionStream(fmt))
      ).arrayBuffer();
      return new TextDecoder('utf-8').decode(buf);
    } catch (_) { /* try next format */ }
  }
  throw new Error('Could not decompress CompressedEntityXml');
}
