/**
 * modal/modal-shell.js
 *
 * Injected on-demand into D365 tabs via chrome.scripting.executeScript.
 * Defines window.__EFPPT_Modal — idempotent, safe to inject multiple times.
 *
 * Each tool gets its own modal keyed by toolName so multiple tools can
 * coexist independently (each minimised to its own FAB button).
 */

'use strict';

(function () {
  if (window.__EFPPT_Modal) return; // Already installed — do nothing.

  // ── Styles (applied inside each shadow root) ─────────────────────────────────

  var CSS = '\
.efppt-bd{\
  position:fixed;inset:0;\
  background:rgba(0,0,0,.52);\
  backdrop-filter:blur(5px);-webkit-backdrop-filter:blur(5px);\
  transition:opacity .25s ease;\
  pointer-events:auto;\
  z-index:1;\
}\
.efppt-bd.gone{opacity:0;pointer-events:none;}\
\
.efppt-modal{\
  position:fixed;top:1vh;left:1vw;width:98vw;height:98vh;\
  background:#fff;\
  border-radius:10px;\
  box-shadow:0 24px 80px rgba(0,0,0,.55);\
  display:flex;flex-direction:column;\
  overflow:hidden;\
  transition:transform .28s cubic-bezier(.4,0,.2,1),opacity .25s ease;\
  transform-origin:bottom right;\
  pointer-events:auto;\
  z-index:2;\
}\
.efppt-modal.gone{\
  transform:scale(.04);\
  opacity:0;\
  pointer-events:none;\
}\
\
.efppt-tb{\
  display:flex;align-items:center;\
  height:44px;min-height:44px;\
  padding:0 10px 0 14px;\
  background:linear-gradient(135deg,#1a1a2e 0%,#16213e 100%);\
  flex-shrink:0;gap:10px;\
  user-select:none;\
}\
.efppt-logo{\
  font-size:11px;font-weight:800;\
  color:#fff;\
  background:rgba(255,255,255,.18);\
  border-radius:4px;\
  padding:2px 6px;\
  letter-spacing:.5px;\
  flex-shrink:0;\
  font-family:system-ui,-apple-system,sans-serif;\
}\
.efppt-tb-title{\
  flex:1;\
  font-size:13px;font-weight:600;\
  color:#e8eaf6;\
  letter-spacing:.2px;\
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;\
  font-family:system-ui,-apple-system,sans-serif;\
}\
.efppt-actions{display:flex;gap:4px;flex-shrink:0;}\
.efppt-btn{\
  display:flex;align-items:center;justify-content:center;\
  width:30px;height:30px;\
  border:none;border-radius:6px;\
  cursor:pointer;\
  background:rgba(255,255,255,.1);\
  color:#e8eaf6;\
  transition:background .15s;\
  flex-shrink:0;\
}\
.efppt-btn:hover{background:rgba(255,255,255,.22);}\
.efppt-btn-close:hover{background:#c0392b;}\
.efppt-btn svg{width:14px;height:14px;display:block;}\
\
.efppt-content{flex:1;position:relative;overflow:hidden;}\
.efppt-frame{position:absolute;inset:0;width:100%;height:100%;border:none;}\
\
.efppt-fab{\
  position:fixed;bottom:20px;right:20px;\
  display:flex;align-items:center;gap:8px;\
  padding:10px 16px 10px 12px;\
  background:linear-gradient(135deg,#1a1a2e 0%,#16213e 100%);\
  color:#e8eaf6;\
  border:none;border-radius:28px;\
  cursor:pointer;\
  font-size:13px;font-weight:600;\
  font-family:system-ui,-apple-system,sans-serif;\
  box-shadow:0 4px 20px rgba(0,0,0,.35);\
  pointer-events:auto;\
  z-index:3;\
  transition:opacity .2s ease,transform .2s ease,box-shadow .15s ease;\
  white-space:nowrap;\
}\
.efppt-fab:hover{box-shadow:0 8px 28px rgba(0,0,0,.45);transform:translateY(-1px);}\
.efppt-fab.gone{opacity:0;pointer-events:none;transform:translateY(8px);}\
.efppt-fab-logo{\
  font-size:10px;font-weight:800;\
  color:#fff;\
  background:rgba(255,255,255,.2);\
  border-radius:4px;\
  padding:2px 5px;\
  letter-spacing:.3px;\
}';

  // ── SVG icons ─────────────────────────────────────────────────────────────────

  var SVG_MINIMIZE =
    '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round">' +
    '<line x1="2" y1="8" x2="14" y2="8"/>' +
    '</svg>';

  // Pop-out: classic "open in new tab" arrow-out-of-box icon
  var SVG_POPOUT =
    '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">' +
    '<path d="M6.5 3H3a1 1 0 0 0-1 1v9a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V9.5"/>' +
    '<polyline points="9.5,2 14,2 14,6.5"/>' +
    '<line x1="8.5" y1="7.5" x2="14" y2="2"/>' +
    '</svg>';

  var SVG_CLOSE =
    '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round">' +
    '<line x1="3" y1="3" x2="13" y2="13"/>' +
    '<line x1="13" y1="3" x2="3" y2="13"/>' +
    '</svg>';

  // ── Factory ───────────────────────────────────────────────────────────────────

  function open(cfg) {
    var toolName  = cfg.toolName  || 'tool';
    var toolUrl   = cfg.toolUrl;
    var toolTitle = cfg.toolTitle || 'EF Tool';
    var HOST_ID   = '__efppt-modal-' + toolName;

    // If the modal already exists in the DOM → restore it (un-minimise).
    var existing = document.getElementById(HOST_ID);
    if (existing && typeof existing._efpptRestore === 'function') {
      existing._efpptRestore();
      return;
    }

    // ── Shadow host ─────────────────────────────────────────────────────────────
    var host = document.createElement('div');
    host.id = HOST_ID;
    // pointer-events:none on the host so the shadow DOM controls interactivity
    // per-element. When minimised only the FAB captures events.
    host.style.cssText = 'position:fixed;inset:0;z-index:2147483647;pointer-events:none;';
    document.body.appendChild(host);

    var shadow = host.attachShadow({ mode: 'open' });

    // Styles
    var styleEl = document.createElement('style');
    styleEl.textContent = CSS;
    shadow.appendChild(styleEl);

    // ── Backdrop ────────────────────────────────────────────────────────────────
    var bd = document.createElement('div');
    bd.className = 'efppt-bd';
    shadow.appendChild(bd);

    // ── Modal ───────────────────────────────────────────────────────────────────
    var modal = document.createElement('div');
    modal.className = 'efppt-modal';
    shadow.appendChild(modal);

    // Title bar
    var tb = document.createElement('div');
    tb.className = 'efppt-tb';

    var logo = document.createElement('div');
    logo.className = 'efppt-logo';
    logo.textContent = 'EF';

    var titleSpan = document.createElement('span');
    titleSpan.className = 'efppt-tb-title';
    titleSpan.textContent = toolTitle;

    var actions = document.createElement('div');
    actions.className = 'efppt-actions';

    var btnMin = document.createElement('button');
    btnMin.className = 'efppt-btn efppt-btn-minimize';
    btnMin.title = 'Minimize';
    btnMin.innerHTML = SVG_MINIMIZE;

    // Pop-out button: hidden in InPrivate because chrome-extension:// URLs are
    // blocked by Edge/Chrome in incognito tabs — there is no workaround.
    var btnPopout = null;
    if (!cfg.isIncognito) {
      btnPopout = document.createElement('button');
      btnPopout.className = 'efppt-btn efppt-btn-popout';
      btnPopout.title = 'Open in dedicated tab';
      btnPopout.innerHTML = SVG_POPOUT;
    }

    var btnClose = document.createElement('button');
    btnClose.className = 'efppt-btn efppt-btn-close';
    btnClose.title = 'Close';
    btnClose.innerHTML = SVG_CLOSE;

    actions.appendChild(btnMin);
    if (btnPopout) actions.appendChild(btnPopout);
    actions.appendChild(btnClose);
    tb.appendChild(logo);
    tb.appendChild(titleSpan);
    tb.appendChild(actions);
    modal.appendChild(tb);

    // Content area + iframe
    var contentArea = document.createElement('div');
    contentArea.className = 'efppt-content';

    var frame = document.createElement('iframe');
    frame.className = 'efppt-frame';
    frame.src = toolUrl;
    frame.setAttribute('allow', 'same-origin');
    contentArea.appendChild(frame);
    modal.appendChild(contentArea);

    // ── FAB (minimised button) ──────────────────────────────────────────────────
    var fab = document.createElement('button');
    fab.className = 'efppt-fab gone'; // hidden initially

    var fabLogo = document.createElement('div');
    fabLogo.className = 'efppt-fab-logo';
    fabLogo.textContent = 'EF';

    var fabLabel = document.createElement('span');
    fabLabel.textContent = toolTitle;

    fab.appendChild(fabLogo);
    fab.appendChild(fabLabel);
    shadow.appendChild(fab);

    // ── State ───────────────────────────────────────────────────────────────────
    var minimised = false;

    function minimize() {
      if (minimised) return;
      minimised = true;
      modal.classList.add('gone');
      bd.classList.add('gone');
      fab.classList.remove('gone');
    }

    function restore() {
      if (!minimised) return;
      minimised = false;
      modal.classList.remove('gone');
      bd.classList.remove('gone');
      fab.classList.add('gone');
    }

    function close() {
      document.removeEventListener('keydown', onKey, true);
      host.remove();
    }

    // Expose restore so a second open() call un-minimises instead of duplicating.
    host._efpptRestore = restore;

    // ── Events ──────────────────────────────────────────────────────────────────
    btnMin.addEventListener('click', minimize);
    if (btnPopout) {
      btnPopout.addEventListener('click', function () {
        // Ask the service worker to open the tool URL as a dedicated tab in this
        // same window.  The service worker has the full chrome.tabs API; the
        // content-script context here does not.
        try {
          chrome.runtime.sendMessage({ type: 'OPEN_TAB_FROM_MODAL', url: toolUrl });
        } catch (_) {
          // Fallback: let the browser open it (may be blocked by popup-blocker).
          window.open(toolUrl, '_blank');
        }
        close();
      });
    }
    btnClose.addEventListener('click', close);
    fab.addEventListener('click', restore);

    // ESC minimises (not closes — avoids accidental data loss).
    function onKey(e) {
      if (e.key !== 'Escape') return;
      if (!document.getElementById(HOST_ID)) {
        // Modal was closed another way — clean up listener.
        document.removeEventListener('keydown', onKey, true);
        return;
      }
      if (!minimised) {
        e.stopPropagation();
        minimize();
      }
    }
    document.addEventListener('keydown', onKey, true);
  }

  // ── Fetch bridge ─────────────────────────────────────────────────────────────
  //
  // Extension iframes (chrome-extension://) don't share the InPrivate session
  // cookie jar — their fetch() calls go out unauthenticated (401).
  //
  // Solution: the tool page detects it is inside a modal (window !== top) and
  // instead of calling fetch() directly, it postMessages a request here.
  // This listener runs in the D365 tab's content-script context, which DOES
  // have InPrivate session cookies, performs the actual fetch, and posts the
  // result back to the extension iframe.
  //
  // Message shape (iframe → D365 tab):
  //   { __efppt: 'fetch', id, url, method?, headers?, body? }
  //
  // Reply shape (D365 tab → iframe):
  //   { __efppt: 'fetch-result', id, ok, status, data?, error? }

  var _extOrigin = '';
  try { _extOrigin = chrome.runtime.getURL('').replace(/\/$/, ''); } catch (_) {}

  window.addEventListener('message', function (e) {
    if (!e.data || e.data.__efppt !== 'fetch') return;
    // Only accept requests from our own extension pages.
    if (_extOrigin && e.origin !== _extOrigin) return;

    var req    = e.data;
    var source = e.source;
    var replyOrigin = _extOrigin || '*';

    var fetchOpts = {
      method:      req.method || 'GET',
      credentials: 'include',
      headers: Object.assign(
        { 'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' },
        req.headers || {}
      ),
    };
    if (req.body != null) {
      fetchOpts.body = typeof req.body === 'string' ? req.body : JSON.stringify(req.body);
      if (!fetchOpts.headers['Content-Type']) {
        fetchOpts.headers['Content-Type'] = 'application/json';
      }
    }

    fetch(req.url, fetchOpts)
      .then(function (res) {
        var ok = res.ok, status = res.status, statusText = res.statusText;
        if (status === 204) {
          source.postMessage(
            { __efppt: 'fetch-result', id: req.id, ok: true, status: 204, data: null },
            replyOrigin
          );
          return;
        }
        res.text().then(function (text) {
          var data = null;
          try { data = JSON.parse(text); } catch (_) {}
          source.postMessage({
            __efppt:    'fetch-result',
            id:         req.id,
            ok:         ok,
            status:     status,
            statusText: statusText,
            data:       ok ? data : null,
            error:      ok ? null : (text.slice(0, 400) || String(status)),
          }, replyOrigin);
        });
      })
      .catch(function (err) {
        source.postMessage(
          { __efppt: 'fetch-result', id: req.id, ok: false, error: err.message },
          replyOrigin
        );
      });
  }, false);

  // ── Public API ────────────────────────────────────────────────────────────────

  window.__EFPPT_Modal = { open: open };

})();
