/**
 * ribbon-api.js — D365 Web API helpers (web-app build).
 *
 * In the web-app the Ribbon tool always runs as an iframe hosted on github.io,
 * which cannot call the D365 Web API directly (cross-origin).  All network
 * calls go through the postMessage fetch bridge implemented in launcher.js,
 * which runs in the D365 page context (same origin, credentials included).
 *
 * Bridge protocol:
 *   request:  { __efppt: 'fetch', id, url, method, headers, isText, body }
 *   reply:    { __efppt: 'fetch-result', id, ok, status, data, error }
 */

'use strict';

import { state } from './ribbon-state.js';

const _inModal = true;

// ─── Bridge fetch ───────────────────────────────────────────────────────────

function _bridgeFetch(url, extraHeaders = {}, isText = false) {
  return new Promise((resolve, reject) => {
    const id    = Math.random().toString(36).slice(2) + Date.now();
    const timer = setTimeout(() => {
      window.removeEventListener('message', onMsg);
      reject(new Error('Request timed out'));
    }, 30_000);
    function onMsg(e) {
      if (!e.data || e.data.__efppt !== 'fetch-result' || e.data.id !== id) return;
      clearTimeout(timer);
      window.removeEventListener('message', onMsg);
      if (e.data.ok) resolve(e.data.data);
      else reject(new Error(e.data.error || 'HTTP ' + e.data.status));
    }
    window.addEventListener('message', onMsg);
    window.parent.postMessage({
      __efppt: 'fetch', id, url,
      method: 'GET',
      headers: extraHeaders,
      isText,
      body: null,
    }, '*');
  });
}

// ─── OData collection fetch (handles @odata.nextLink pagination) ───────────

export async function fetchAllPages(path) {
  const results = [];
  let url = path.startsWith('http') ? path : `${state.envUrl}/api/data/v9.2${path}`;
  while (url) {
    const data = await d365Fetch(url);
    results.push(...(data.value ?? []));
    url = data['@odata.nextLink'] ?? null;
  }
  return results;
}

// ─── Core fetch helpers (bridged) ──────────────────────────────────────────

/** Fetch JSON from D365 (collection or single record returning JSON). */
export async function d365Fetch(fullUrl) {
  return _bridgeFetch(fullUrl, {
    'Prefer': 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
  });
}

/** Fetch a single JSON record (used for unbound functions, single-entity GETs). */
export async function d365FetchSingle(fullUrl) {
  return _bridgeFetch(fullUrl, {});
}

/** Fetch plain text (used for $value endpoints). */
export async function d365FetchText(fullUrl) {
  // launcher.js auto-detects non-JSON responses and returns the raw text;
  // the isText flag is forwarded for clarity.
  return _bridgeFetch(fullUrl, { 'Accept': 'text/plain, application/xml, */*' }, true);
}
