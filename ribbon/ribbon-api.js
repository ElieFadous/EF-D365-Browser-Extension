/**
 * ribbon-api.js — D365 Web API helpers.
 *
 * All network calls are injected into the proxy tab via chrome.scripting so
 * they run in the same origin as D365.  Each injected function uses its own
 * AbortController so a hung request never blocks the UI permanently.
 *
 * Re-usable outside the Ribbon module — just pass the correct proxyTabId via
 * the shared state object.
 */

'use strict';

import { state } from './ribbon-state.js';

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

// ─── Core injected fetch helpers ──────────────────────────────────────────

/** Fetch JSON from D365 (collection or single record returning JSON). */
export async function d365Fetch(fullUrl) {
  _requireProxy();
  const [injection] = await chrome.scripting.executeScript({
    target: { tabId: state.proxyTabId },
    func: async (url) => {
      try {
        const ctrl  = new AbortController();
        const timer = setTimeout(() => ctrl.abort(), 30000);
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
          return { error: `HTTP ${res.status} ${res.statusText}: ${body.slice(0, 400)}` };
        }
        return { data: await res.json() };
      } catch (e) {
        if (e.name === 'AbortError') return { error: 'Request timed out after 30 s.' };
        console.error('[EF PPT proxy]', e);
        return { error: e.message };
      }
    },
    args: [fullUrl],
  });
  if (injection.result?.error) throw new Error(injection.result.error);
  return injection.result.data;
}

/** Fetch a single JSON record (used for unbound functions, single-entity GETs). */
export async function d365FetchSingle(fullUrl) {
  _requireProxy();
  const [injection] = await chrome.scripting.executeScript({
    target: { tabId: state.proxyTabId },
    func: async (url) => {
      try {
        const ctrl  = new AbortController();
        const timer = setTimeout(() => ctrl.abort(), 30000);
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
          return { error: `HTTP ${res.status} ${res.statusText}: ${body.slice(0, 400)}` };
        }
        return { data: await res.json() };
      } catch (e) {
        if (e.name === 'AbortError') return { error: 'Request timed out after 30 s.' };
        console.error('[EF PPT proxy]', e);
        return { error: e.message };
      }
    },
    args: [fullUrl],
  });
  if (injection.result?.error) throw new Error(injection.result.error);
  return injection.result.data;
}

/** Fetch plain text (used for $value endpoints). */
export async function d365FetchText(fullUrl) {
  _requireProxy();
  const [injection] = await chrome.scripting.executeScript({
    target: { tabId: state.proxyTabId },
    func: async (url) => {
      try {
        const ctrl  = new AbortController();
        const timer = setTimeout(() => ctrl.abort(), 30000);
        const res   = await fetch(url, {
          headers: {
            'Accept':           'text/plain, application/xml, */*',
            'OData-MaxVersion': '4.0',
            'OData-Version':    '4.0',
          },
          signal: ctrl.signal,
        });
        clearTimeout(timer);
        if (!res.ok) {
          const body = await res.text().catch(() => '');
          return { error: `HTTP ${res.status} ${res.statusText}: ${body.slice(0, 400)}` };
        }
        return { data: await res.text() };
      } catch (e) {
        if (e.name === 'AbortError') return { error: 'Request timed out after 30 s.' };
        console.error('[EF PPT proxy]', e);
        return { error: e.message };
      }
    },
    args: [fullUrl],
  });
  if (injection.result?.error) throw new Error(injection.result.error);
  return injection.result.data;
}

// ─── Extension messaging ───────────────────────────────────────────────────

export function sendMessage(msg) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(msg, response => {
      if (chrome.runtime.lastError) return reject(new Error(chrome.runtime.lastError.message));
      if (response?.error)          return reject(new Error(response.error));
      resolve(response);
    });
  });
}

// ─── Internal ─────────────────────────────────────────────────────────────

function _requireProxy() {
  if (state.proxyTabId === null) throw new Error('Proxy tab not available.');
}
