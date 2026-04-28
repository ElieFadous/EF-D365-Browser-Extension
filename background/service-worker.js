/**
 * EF Power Platform Tools — Background Service Worker
 *
 * Handles all privileged operations:
 *   • GOTO_ENVIRONMENT  — resolves app IDs and opens cross-env record URLs
 */

// ─── Extension icon theming ───────────────────────────────────────────────────

/** WCAG relative-luminance contrast helper — returns '#000' or '#fff'. */
function _contrastFg(hexColor) {
  const h  = hexColor.replace('#', '');
  const lin = x => {
    const c = parseInt(x, 16) / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  };
  const L = 0.2126 * lin(h.slice(0,2))
          + 0.7152 * lin(h.slice(2,4))
          + 0.0722 * lin(h.slice(4,6));
  return L > 0.179 ? '#000' : '#fff';
}

/** Draw the Angular D + Bolt icon using Canvas Path2D.
 *  The D shape is filled with the environment's colour; the bolt is always gold.
 *  Path data is based on the icon's 64×64 viewBox — scaled to `size`. */
function _createIconImageData(color, size) {
  const canvas = new OffscreenCanvas(size, size);
  const ctx    = canvas.getContext('2d');
  const s      = size / 64; // scale factor: paths are defined on a 64×64 grid

  ctx.clearRect(0, 0, size, size);

  // ── Angular D shape (fill-rule evenodd creates the inner cutout) ──────────
  ctx.fillStyle = color;
  const dPath = new Path2D();
  // Outer hexagonal-D contour
  dPath.moveTo(10*s, 6*s); dPath.lineTo(46*s, 6*s); dPath.lineTo(58*s, 18*s);
  dPath.lineTo(58*s, 46*s); dPath.lineTo(46*s, 58*s); dPath.lineTo(10*s, 58*s);
  dPath.closePath();
  // Inner cutout (evenodd makes this transparent)
  dPath.moveTo(22*s, 17*s); dPath.lineTo(40*s, 17*s); dPath.lineTo(50*s, 27*s);
  dPath.lineTo(50*s, 37*s); dPath.lineTo(40*s, 47*s); dPath.lineTo(22*s, 47*s);
  dPath.closePath();
  ctx.fill(dPath, 'evenodd');

  // ── Gold lightning bolt ───────────────────────────────────────────────────
  ctx.fillStyle = '#FFB900';
  const boltPath = new Path2D();
  boltPath.moveTo(36*s, 21*s); boltPath.lineTo(28*s, 34*s);
  boltPath.lineTo(33*s, 34*s); boltPath.lineTo(27*s, 47*s);
  boltPath.lineTo(43*s, 31*s); boltPath.lineTo(37*s, 31*s);
  boltPath.closePath();
  ctx.fill(boltPath);

  return ctx.getImageData(0, 0, size, size);
}

/**
 * Match a tab URL to a configured environment.
 * Handles both D365 origins and make.powerapps.com / make.powerautomate.com by powerAppsId.
 */
function _findEnvForUrl(url, envs) {
  if (!url) return null;
  try {
    const parsed = new URL(url);
    // 1. Exact D365 origin match.
    const byOrigin = envs.find(e => e.url === parsed.origin);
    if (byOrigin) return byOrigin;
    // 2. make.powerapps.com or make.powerautomate.com — match by powerAppsId.
    if (parsed.hostname === 'make.powerapps.com' || parsed.hostname === 'make.powerautomate.com') {
      const m = parsed.pathname.match(/^\/environments\/([^/]+)/i);
      if (m) {
        const paId = m[1].toLowerCase();
        return envs.find(e => e.powerAppsId && e.powerAppsId.toLowerCase() === paId) ?? null;
      }
    }
    return null;
  } catch (_) {
    return null;
  }
}

/**
 * Read the environment list from storage, match the given tab URL, and
 * repaint the browser action icon with that environment's configured colour.
 * Falls back to neutral gray (#6b7280) when no match is found.
 */
async function updateIconForTab(tabId) {
  try {
    const tab = await chrome.tabs.get(tabId);
    if (!tab?.url) return;

    const stored = await chrome.storage.local.get('environments');
    const envs   = Array.isArray(stored.environments) ? stored.environments : [];

    const match = _findEnvForUrl(tab.url, envs);
    const color = match?.color ?? '#6b7280';  // neutral gray — no recognised environment

    await chrome.action.setIcon({
      imageData: {
        16: _createIconImageData(color, 16),
        32: _createIconImageData(color, 32),
      }
    });
  } catch (_) { /* tab may have closed — ignore */ }
}

/**
 * If the environment has `highlightTab: true`, inject a fixed overlay div that
 * draws a colored border around the entire viewport.
 * Uses executeScript + a real DOM element (more reliable than ::before on complex SPAs).
 * Called on every completed tab navigation so the border is present even in background tabs.
 */
async function _injectTabHighlight(tabId, url) {
  try {
    if (!url || !url.startsWith('https://')) return;
    const stored = await chrome.storage.local.get('environments');
    const envs   = Array.isArray(stored.environments) ? stored.environments : [];
    const env    = _findEnvForUrl(url, envs);
    if (!env?.highlightTab || !env?.color) return;

    await chrome.scripting.executeScript({
      target: { tabId },
      func: (color) => {
        const ID = '__ef-ppt-env-border__';
        if (document.getElementById(ID)) return; // guard against double-injection
        const el = document.createElement('div');
        el.id = ID;
        el.style.cssText = [
          'position:fixed',
          'inset:0',
          `box-shadow:inset 0 0 0 4px ${color}`,
          'pointer-events:none',
          'z-index:2147483647',
          'border-radius:0',
        ].join(';');
        document.documentElement.appendChild(el);
      },
      args: [env.color],
    });
  } catch (_) { /* non-injectable pages (chrome://, devtools, etc.) — silently ignore */ }
}

// Repaint icon whenever the user switches tabs.
chrome.tabs.onActivated.addListener(({ tabId }) => updateIconForTab(tabId));

// Repaint icon and inject highlight border whenever any tab finishes navigating.
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (changeInfo.status !== 'complete') return;
  if (tab.active) updateIconForTab(tabId);
  _injectTabHighlight(tabId, tab.url);
});

// ─── Track the last focused browser window ────────────────────────────────────
//
// The popup cannot reliably determine its own window ID — Edge may shift focus
// away from the InPrivate window the moment the popup overlay appears.
// The service worker has no such problem: it listens to onFocusChanged
// continuously and records the most recently focused NORMAL window (InPrivate
// or regular) to chrome.storage.session before any popup interference.

chrome.windows.onFocusChanged.addListener(async (windowId) => {
  if (windowId === chrome.windows.WINDOW_ID_NONE) return; // focus left Edge entirely
  try {
    const win = await chrome.windows.get(windowId);
    if (win.type !== 'normal') return; // ignore popups, devtools, etc.
    // Persist to session storage so the value survives brief SW suspension.
    await chrome.storage.session.set({
      _efppt_winId:        win.id,
      _efppt_winIncognito: win.incognito ?? false,
    });
  } catch { /* window may have closed before we could read it */ }
});

// ─── Message Router ───────────────────────────────────────────────────────────

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  switch (message.type) {

    case 'GOTO_ENVIRONMENT':
      handleGoTo(message)
        .then(r  => sendResponse(r))
        .catch(e => sendResponse({ error: e.message }));
      return true;

    case 'GET_FOCUSED_WINDOW':
      chrome.storage.session
        .get(['_efppt_winId', '_efppt_winIncognito'])
        .then(s => {
          sendResponse({ windowId: s._efppt_winId ?? null, incognito: s._efppt_winIncognito ?? false });
        })
        .catch(() => sendResponse({ windowId: null, incognito: false }));
      return true;

    case 'OPEN_TAB_FROM_MODAL': {
      // Sent by the modal pop-out button (content script / isolated world).
      // Open the tool URL as a dedicated tab in the same window as the originating D365 tab.
      const windowId = sender.tab?.windowId ?? null;
      const createOpts = { url: message.url, active: true };
      if (windowId !== null) createOpts.windowId = windowId;
      chrome.tabs.create(createOpts)
        .then(() => sendResponse({ ok: true }))
        .catch(e => sendResponse({ error: e.message }));
      return true;
    }

    case 'OPEN_PROXY_TAB':
      getOrOpenProxyTab(message.env, message.windowId ?? null)
        .then(result => sendResponse(result))
        .catch(e     => sendResponse({ error: e.message }));
      return true;

    case 'CLOSE_PROXY_TAB':
      // Only close the tab if WE opened it — do not close a tab the user already had open.
      if (!message.reused) chrome.tabs.remove(message.tabId).catch(() => {});
      sendResponse({ ok: true });
      return false;

    case 'INJECT_TOOL':
      handleInjectTool(message)
        .then(r  => sendResponse(r))
        .catch(e => sendResponse({ error: e.message }));
      return true;

    default:
      return false;
  }
});

// ─── Tool modal injection ─────────────────────────────────────────────────────

/**
 * Injects the EF modal shell into the target tab (idempotent), then tells it
 * to open the requested tool. If the modal for that tool already exists in the
 * tab it is restored (un-minimised) rather than duplicated.
 */
async function handleInjectTool({ tabId, toolName, toolUrl, toolTitle }) {
  // Detect incognito so the modal can hide the pop-out button
  // (chrome-extension:// URLs are blocked in InPrivate tabs).
  const tab = await chrome.tabs.get(tabId).catch(() => null);
  const isIncognito = tab?.incognito ?? false;

  // In InPrivate mode, tell the tool page via a URL flag so it routes fetches
  // through the modal bridge (extension iframes can't access InPrivate cookies).
  // In regular mode the iframe fetches D365 APIs directly — no bridge needed.
  const effectiveUrl = isIncognito
    ? toolUrl + (toolUrl.includes('?') ? '&' : '?') + 'incognito=1'
    : toolUrl;

  // Step 1: ensure the modal shell is installed in the tab.
  // executeScript with files[] is idempotent — modal-shell.js bails immediately
  // if window.__EFPPT_Modal is already defined.
  await chrome.scripting.executeScript({
    target: { tabId },
    files:  ['modal/modal-shell.js'],
  });

  // Step 2: open (or restore) the specific tool modal.
  await chrome.scripting.executeScript({
    target: { tabId },
    func:   function (cfg) { window.__EFPPT_Modal && window.__EFPPT_Modal.open(cfg); },
    args:   [{ toolName, toolUrl: effectiveUrl, toolTitle, isIncognito }],
  });

  return { ok: true };
}

// ─── Shared helper: open a tab and wait for it to finish loading ──────────────

/**
 * Opens a new background tab to `url` and resolves with the tab ID once the
 * tab reaches `complete` status on the expected origin (`expectedOrigin`).
 *
 * After the tab reaches `complete` at the expected origin we wait one second
 * and re-verify the URL. This guards against D365's SPA triggering a silent
 * token-refresh redirect (to login.microsoftonline.com) immediately after the
 * initial load — which would cause executeScript to fail with a host-permission
 * error even though the tab briefly appeared to be on the right domain.
 *
 * Rejects if:
 *  • The tab redirects away from the expected origin (auth required).
 *  • 30 seconds elapse without reaching `complete`.
 *
 * @param {string}  expectedOrigin  Base URL (origin) that the tab must stay on.
 * @param {string}  [startUrl]      URL to open; defaults to expectedOrigin.
 * @param {boolean} [keepOpen]      If true, the tab is NOT closed on failure.
 * @param {number}  [windowId]      If provided, open the tab in this window
 *                                  (preserves incognito context of the popup).
 */
function openAndWaitForTab(expectedOrigin, startUrl = expectedOrigin, keepOpen = false, windowId = null) {
  return new Promise((resolve, reject) => {
    let tabId   = null;
    let settled = false;

    const settle = (value, err) => {
      if (settled) return;
      settled = true;
      chrome.tabs.onUpdated.removeListener(onUpdated);
      if (err) {
        if (!keepOpen && tabId !== null) chrome.tabs.remove(tabId).catch(() => {});
        reject(err);
      } else {
        resolve(value);
      }
    };

    const onUpdated = (id, changeInfo, tab) => {
      if (id !== tabId || changeInfo.status !== 'complete') return;

      if (!tab.url?.startsWith(expectedOrigin)) {
        settle(null, new Error(
          `You are not signed in to ${expectedOrigin}. ` +
          `Please open that environment in Edge and sign in, then try again.`
        ));
        return;
      }

      // Tab reached 'complete' at the expected origin.
      // Wait 1 s and re-verify: D365's SPA may trigger a silent auth-token
      // refresh redirect immediately after the initial load completes.
      setTimeout(async () => {
        try {
          const latest = await chrome.tabs.get(tabId);
          if (latest.url?.startsWith(expectedOrigin)) {
            settle(tabId); // Still on the right domain — good to use.
          } else {
            settle(null, new Error(
              `You are not signed in to ${expectedOrigin}. ` +
              `Please open that environment in Edge and sign in, then try again.`
            ));
          }
        } catch {
          // Tab was closed in the 1 s window — let executeScript surface the error.
          settle(tabId);
        }
      }, 1_000);
    };

    chrome.tabs.onUpdated.addListener(onUpdated);

    const createOpts = { url: startUrl, active: false };
    if (windowId !== null) createOpts.windowId = windowId;

    chrome.tabs.create(createOpts)
      .then(tab => { tabId = tab.id; })
      .catch(err => {
        chrome.tabs.onUpdated.removeListener(onUpdated);
        reject(err);
      });

    setTimeout(
      () => settle(null, new Error('Timeout: environment took too long to respond.')),
      30_000
    );
  });
}

// ─── Proxy tab: reuse existing tab or open a new background one ───────────────

/**
 * Returns { tabId, reused: true }  if an existing tab on envUrl is found.
 * Returns { tabId, reused: false } after opening a new background tab.
 *
 * @param {string} envUrl    Environment origin URL.
 * @param {number} windowId  Chrome window ID from the popup (for incognito support).
 */
async function getOrOpenProxyTab(envUrl, windowId = null) {
  const existing = await chrome.tabs.query({
    url:       `${envUrl}/*`,
    status:    'complete',
    discarded: false,
  });
  if (existing.length > 0) {
    return { tabId: existing[0].id, reused: true };
  }
  const tabId = await openAndWaitForTab(envUrl, envUrl, false, windowId);
  return { tabId, reused: false };
}

// ─── Go To Handler ────────────────────────────────────────────────────────────

async function handleGoTo({ currentTabId, currentUrl, targetEnv, windowId = null }) {
  const url    = new URL(currentUrl);
  const params = new URLSearchParams(url.search);
  const appId  = params.get('appid');

  if (!appId) {
    throw new Error('No App ID found in the current URL. Please open a model-driven app first.');
  }

  // 1. Resolve unique name from current environment (same-origin via content script).
  const uniqueName = await resolveUniqueName(currentTabId, url.origin, appId);

  // 2. Resolve app ID in the target environment.
  const targetAppId = await resolveTargetAppId(targetEnv, uniqueName, windowId);

  // 3. Reconstruct URL.
  params.set('appid', targetAppId);
  const targetUrl = `${targetEnv.url}${url.pathname}?${params.toString()}`;

  // 4. Open in a new tab — use the same window as the popup (preserves incognito context).
  const createOpts = { url: targetUrl, active: true };
  if (windowId !== null) createOpts.windowId = windowId;
  await chrome.tabs.create(createOpts);
  return { success: true, url: targetUrl };
}

// ─── Step 1: Resolve unique name in the current environment ──────────────────

async function resolveUniqueName(tabId, baseUrl, appId) {
  const apiUrl = `${baseUrl}/api/data/v9.2/appmodules(${appId})?$select=uniquename`;

  const [injection] = await chrome.scripting.executeScript({
    target: { tabId },
    func: async (url) => {
      try {
        const res = await fetch(url, {
          headers: {
            'Accept':           'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version':    '4.0'
          }
        });
        if (!res.ok) return { error: `API error ${res.status}: ${res.statusText}` };
        const data = await res.json();
        if (!data.uniquename) return { error: 'uniquename field missing in API response.' };
        return { uniquename: data.uniquename };
      } catch (e) {
        return { error: e.message };
      }
    },
    args: [apiUrl]
  });

  if (injection.result?.error) throw new Error(injection.result.error);
  return injection.result.uniquename;
}

// ─── Step 2: Resolve app ID in the target environment ────────────────────────

async function resolveTargetAppId(targetEnv, uniqueName, windowId = null) {
  // Strategy A — direct service-worker fetch (fast, avoids extra tab).
  try {
    const appId = await fetchAppIdDirect(targetEnv.url, uniqueName);
    if (appId) return appId;
  } catch (e) {}

  // Strategy B — background tab injection (reliable fallback).
  return fetchAppIdViaBackgroundTab(targetEnv, uniqueName, windowId);
}

async function fetchAppIdDirect(baseUrl, uniqueName) {
  const apiUrl = buildAppModulesUrl(baseUrl, uniqueName);
  const res    = await fetch(apiUrl, {
    credentials: 'include',
    headers: {
      'Accept':           'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version':    '4.0'
    }
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const data  = await res.json();
  const appId = data.value?.[0]?.appmoduleid;
  if (!appId) throw new Error('App not found in response');
  return appId;
}

async function fetchAppIdViaBackgroundTab(targetEnv, uniqueName, windowId = null) {
  const apiUrl = buildAppModulesUrl(targetEnv.url, uniqueName);

  // Open a background tab and wait for the environment to load.
  const tabId = await openAndWaitForTab(targetEnv.url, targetEnv.url, false, windowId);

  try {
    const [injection] = await chrome.scripting.executeScript({
      target: { tabId },
      func: async (url) => {
        try {
          const res = await fetch(url, {
            headers: {
              'Accept':           'application/json',
              'OData-MaxVersion': '4.0',
              'OData-Version':    '4.0'
            }
          });
          if (!res.ok) return { error: `HTTP ${res.status}: ${res.statusText}` };
          const data  = await res.json();
          const appId = data.value?.[0]?.appmoduleid;
          if (!appId) return { error: 'App not found in target environment. The app may not exist there.' };
          return { appmoduleid: appId };
        } catch (e) {
          return { error: e.message };
        }
      },
      args: [apiUrl]
    });

    if (injection.result?.error) throw new Error(injection.result.error);
    return injection.result.appmoduleid;
  } finally {
    // Always clean up the temporary tab.
    chrome.tabs.remove(tabId).catch(() => {});
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function buildAppModulesUrl(baseUrl, uniqueName) {
  return `${baseUrl}/api/data/v9.2/appmodules?$filter=uniquename eq '${uniqueName}'&$select=appmoduleid`;
}
