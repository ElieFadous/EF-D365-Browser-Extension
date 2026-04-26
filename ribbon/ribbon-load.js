/**
 * ribbon-load.js — Ribbon data loading.
 *
 * ensureAllRibbonDataLoaded() is the single entry point.  It fires one
 * executeScript call into the proxy tab which runs ALL entity-scoped fetches
 * in parallel.  This approach is used because:
 *
 *  • Per-entity OData queries (?$filter=entity eq 'X') are confirmed to return
 *    the `rdx` XML field in D365 responses.
 *  • Unfiltered /ribbondiffs requests do NOT return `rdx` by default (D365
 *    omits large text columns from broad collection responses).
 *  • A single executeScript avoids N round-trips through the extension bus.
 */

'use strict';

import { state }                                         from './ribbon-state.js';
import { d365FetchSingle, d365FetchText, fetchAllPages } from './ribbon-api.js';
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

// ─── Parallel per-entity fetch (runs inside proxy tab) ────────────────────

/**
 * Fire all entity-scoped OData queries in parallel inside one executeScript
 * call.  Per-entity queries reliably return the `rdx` XML field even when the
 * unfiltered /ribbondiffs endpoint omits it.
 */
async function _fetchEntitiesParallel(entities) {
  if (state.proxyTabId === null) throw new Error('Proxy tab not available.');

  const [injection] = await chrome.scripting.executeScript({
    target: { tabId: state.proxyTabId },
    func: async (baseUrl, entityList) => {
      const TIMEOUT   = 30000;
      const BASE      = `${baseUrl}/api/data/v9.2`;
      const HEADERS   = {
        Accept:           'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version':    '4.0',
      };

      async function fetchJson(url) {
        const ctrl  = new AbortController();
        const timer = setTimeout(() => ctrl.abort(), TIMEOUT);
        try {
          const res = await fetch(url, { headers: HEADERS, signal: ctrl.signal });
          clearTimeout(timer);
          if (!res.ok) throw new Error(`HTTP ${res.status}`);
          return res.json();
        } catch (e) {
          clearTimeout(timer);
          throw e;
        }
      }

      async function fetchEntityRows(entity) {
        // Build filter.  '' / null both mean Application Ribbon.
        const filter = (entity === '' || entity == null)
          ? `entity eq null or entity eq ''`
          : `entity eq '${entity}'`;
        const url0 = `${BASE}/ribbondiffs?$filter=${encodeURIComponent(filter)}`;
        const rows = [];
        let url = url0;
        while (url) {
          const data = await fetchJson(url);
          rows.push(...(data.value ?? []));
          url = data['@odata.nextLink'] ?? null;
        }
        return rows;
      }

      try {
        const perEntity = await Promise.all(
          entityList.map(e => fetchEntityRows(e).catch(() => []))
        );
        return { data: perEntity.flat() };
      } catch (e) {
        console.error('[EF PPT proxy]', e);
        return { error: e.message };
      }
    },
    args: [state.envUrl, entities],
  });

  if (injection.result?.error) throw new Error(injection.result.error);

  // Deduplicate by ribbondiffid (same record may appear under multiple entities).
  const seen = new Set();
  return (injection.result.data ?? []).filter(r => {
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
 * location filter.  Fetch AND decompress run inside the proxy tab so that
 * DecompressionStream has full browser-tab APIs and no large binary blob
 * passes through the message channel.
 */
export async function fetchEntityRibbonXml(entityLogicalName) {
  if (state.proxyTabId === null) throw new Error('Proxy tab not available.');

  const [injection] = await chrome.scripting.executeScript({
    target: { tabId: state.proxyTabId },
    func: async (baseUrl, entity) => {
      const url =
        `${baseUrl}/api/data/v9.2/RetrieveEntityRibbon` +
        `(EntityName='${encodeURIComponent(entity)}',` +
        `RibbonLocationFilter=Microsoft.Dynamics.CRM.RibbonLocationFilters'All')`;
      try {
        const ctrl  = new AbortController();
        const timer = setTimeout(() => ctrl.abort(), 60000);
        const res   = await fetch(url, {
          headers: {
            'Accept':           'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version':    '4.0',
          },
          signal: ctrl.signal,
        });
        clearTimeout(timer);
        if (!res.ok) {
          const body = await res.text().catch(() => '');
          return { error: `HTTP ${res.status}: ${body.slice(0, 400)}` };
        }
        const data = await res.json();
        if (!data.CompressedEntityXml) return { error: 'No CompressedEntityXml in response' };

        // Decompress here in the tab — full browser API access, no message-channel limits.
        // Use pipeThrough so decompression errors are raised as rejections, not stream stalls.
        const b64   = data.CompressedEntityXml.replace(/\s/g, ''); // strip .NET line-breaks
        const bin   = atob(b64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);

        // D365 wraps the ribbon XML in a PKZIP archive (magic 50 4b 03 04).
        // Parse the local file header manually and decompress the first entry.
        const sig = (bytes[0] << 24 | bytes[1] << 16 | bytes[2] << 8 | bytes[3]) >>> 0;
        if (sig === 0x504b0304) {
          // Local file header layout:
          //   0–3   signature
          //   4–5   version needed
          //   6–7   general purpose bit flag
          //   8–9   compression method  (0=stored, 8=deflate)
          //  10–11  last mod time
          //  12–13  last mod date
          //  14–17  CRC-32
          //  18–21  compressed size
          //  22–25  uncompressed size
          //  26–27  file name length
          //  28–29  extra field length
          //  30+    file name, extra field, then compressed data
          const view             = new DataView(bytes.buffer, bytes.byteOffset);
          const method           = view.getUint16(8,  true);
          const compressedSize   = view.getUint32(18, true);
          const fileNameLen      = view.getUint16(26, true);
          const extraLen         = view.getUint16(28, true);
          const dataStart        = 30 + fileNameLen + extraLen;
          const payload          = bytes.slice(dataStart,
            compressedSize ? dataStart + compressedSize : undefined);

          if (method === 0) {
            // Stored — no compression.
            return { xml: new TextDecoder('utf-8').decode(payload) };
          }
          if (method === 8) {
            // Deflate (raw — ZIP entries have no zlib wrapper).
            try {
              const buf = await new Response(
                new Response(payload).body.pipeThrough(new DecompressionStream('deflate-raw'))
              ).arrayBuffer();
              return { xml: new TextDecoder('utf-8').decode(buf) };
            } catch (e) {
              console.error('[EF PPT proxy]', e);
              return { error: `ZIP deflate decompression failed: ${e.message}` };
            }
          }
          return { error: `Unsupported ZIP compression method: ${method}` };
        }

        // Fallback: maybe the data is plain (uncompressed) XML.
        const asText = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
        if (asText.trimStart().startsWith('<')) return { xml: asText };

        // Fallback: standard gzip / zlib / raw deflate.
        for (const fmt of ['gzip', 'deflate', 'deflate-raw']) {
          try {
            const buf = await new Response(
              new Response(bytes).body.pipeThrough(new DecompressionStream(fmt))
            ).arrayBuffer();
            return { xml: new TextDecoder('utf-8').decode(buf) };
          } catch (_) { /* try next format */ }
        }

        const magic = Array.from(bytes.slice(0, 8))
          .map(b => b.toString(16).padStart(2, '0')).join(' ');
        return { error: `Could not decompress ribbon XML. Magic bytes: [${magic}]` };
      } catch (e) {
        return { error: e.message };
      }
    },
    args: [state.envUrl, entityLogicalName],
  });

  if (injection.result?.error) throw new Error(injection.result.error);
  return injection.result?.xml ?? '';
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
