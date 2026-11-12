'use strict';

import { state }     from './ribbon-state.js';
import { d365Fetch } from './ribbon-api.js';

const IMAGE_TYPES = { 5: 'image/png', 6: 'image/jpeg', 7: 'image/gif', 10: 'image/x-icon', 11: 'image/svg+xml' };
const _cache = new Map();
let _errorHooked = false;

export function iconSource(value) {
  const v = String(value ?? '').trim();
  if (!v) return null;
  let m = /^\$webresource:(.+)$/i.exec(v);
  if (m) return { kind: 'wr', name: m[1].trim(), raw: v };
  m = /^\/?WebResources\/(.+)$/i.exec(v);
  if (m) return { kind: 'wr', name: decodeURIComponent(m[1].split('?')[0]), raw: v };
  if (/^https?:\/\//i.test(v)) return { kind: 'url', url: v, raw: v };
  if (v.startsWith('/')) return { kind: 'url', url: state.envUrl + v, raw: v };
  if (/^\$Resources:/i.test(v)) return { kind: 'resource', raw: v };
  return { kind: 'symbol', name: v, raw: v };
}

export function iconSources(button) {
  const icon = button.icon ?? {};
  return [
    ['Modern', icon.modern],
    ['32×32',  icon.img32],
    ['16×16',  icon.img16],
  ].map(([slot, value]) => ({ slot, src: iconSource(value) })).filter(x => x.src);
}

function _primary(button) {
  const all = iconSources(button).map(x => x.src);
  return all.find(s => s.kind === 'wr' || s.kind === 'url') ?? all.find(s => s.kind === 'symbol') ?? null;
}

function _esc(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

export function iconHtml(button) {
  const src = _primary(button);
  if (!src) return '<span class="btn-icon btn-icon--none" title="No icon set">&#9633;</span>';
  if (src.kind === 'url') {
    const symbol = iconSources(button).map(x => x.src).find(s => s.kind === 'symbol');
    return `<span class="btn-icon" title="${_esc(src.raw)}"><img src="${_esc(src.url)}" alt="" loading="lazy" data-icon-url${symbol ? ` data-fallback="${_esc(symbol.name)}"` : ''}></span>`;
  }
  if (src.kind === 'symbol') {
    return `<span class="btn-icon btn-icon--symbol" title="Modern icon: ${_esc(src.name)}">${_esc(src.name.slice(0, 2))}</span>`;
  }
  const key = src.name.toLowerCase();
  const hit = _cache.get(key);
  if (hit && hit.src) return `<span class="btn-icon" title="${_esc(src.name)}"><img src="${_esc(hit.src)}" alt=""></span>`;
  if (hit && hit.done) return `<span class="btn-icon btn-icon--missing" title="${_esc(hit.reason)}: ${_esc(src.name)}">?</span>`;
  return `<span class="btn-icon btn-icon--loading" data-wr="${_esc(src.name)}" title="${_esc(src.name)}"></span>`;
}

export function iconSourcesHtml(button) {
  const list = iconSources(button);
  if (!list.length) return '';
  const describe = src =>
    src.kind === 'wr'       ? `<code>${_esc(src.name)}</code> <span class="icon-kind">web resource</span>` :
    src.kind === 'symbol'   ? `<code>${_esc(src.name)}</code> <span class="icon-kind">modern icon</span>` :
    src.kind === 'resource' ? `<code>${_esc(src.raw)}</code> <span class="icon-kind">built-in</span>` :
                              `<code>${_esc(src.raw)}</code> <span class="icon-kind">image path</span>`;
  return `
    <div class="card-section">
      <div class="card-section-title">Icon</div>
      <ul class="icon-list">${list.map(x => `<li><span class="icon-slot">${_esc(x.slot)}</span>${describe(x.src)}</li>`).join('')}</ul>
    </div>`;
}

async function _loadWebResources(names) {
  const todo = [...new Set(names.map(n => n.toLowerCase()))].filter(n => !_cache.has(n));
  if (!todo.length) return;
  todo.forEach(n => _cache.set(n, { done: false }));
  const chunks = [];
  for (let i = 0; i < todo.length; i += 15) chunks.push(todo.slice(i, i + 15));
  await Promise.all(chunks.map(async chunk => {
    const filter = chunk.map(n => `name eq '${n.replace(/'/g, "''")}'`).join(' or ');
    let rows = [];
    try {
      const data = await d365Fetch(
        `${state.envUrl}/api/data/v9.2/webresourceset?$select=name,webresourcetype,content&$filter=${encodeURIComponent(filter)}`
      );
      rows = data.value ?? [];
    } catch (e) {
      console.warn('[EF PPT] Could not load ribbon icons:', e);
      chunk.forEach(n => _cache.set(n, { done: true, reason: `Couldn't load (${e.message})` }));
      return;
    }
    const byName = new Map(rows.map(r => [String(r.name).toLowerCase(), r]));
    chunk.forEach(n => {
      const row = byName.get(n);
      const mime = row ? IMAGE_TYPES[row.webresourcetype] : null;
      if (row && mime && row.content) _cache.set(n, { done: true, src: `data:${mime};base64,${row.content}` });
      else _cache.set(n, { done: true, reason: row ? 'Web resource is not an image' : 'Web resource not found' });
    });
  }));
}

export async function hydrateIcons(root = document) {
  if (!_errorHooked) {
    _errorHooked = true;
    document.addEventListener('error', e => {
      const img = e.target;
      if (!(img instanceof HTMLImageElement) || !img.hasAttribute('data-icon-url')) return;
      const span = img.parentElement;
      const fallback = img.dataset.fallback;
      if (fallback) {
        span.classList.add('btn-icon--symbol');
        span.title = `Modern icon: ${fallback}`;
        span.textContent = fallback.slice(0, 2);
        return;
      }
      span.classList.add('btn-icon--missing');
      span.title = `Couldn't load ${span.title}`;
      span.textContent = '?';
    }, true);
  }
  const names = [...root.querySelectorAll('[data-wr]')].map(el => el.dataset.wr);
  if (!names.length) return;
  await _loadWebResources(names);
  document.querySelectorAll('[data-wr]').forEach(el => {
    const hit = _cache.get(el.dataset.wr.toLowerCase());
    if (!hit || !hit.done) return;
    el.removeAttribute('data-wr');
    el.classList.remove('btn-icon--loading');
    if (hit.src) {
      el.innerHTML = '';
      const img = document.createElement('img');
      img.alt = '';
      img.src = hit.src;
      el.appendChild(img);
    } else {
      el.classList.add('btn-icon--missing');
      el.title = `${hit.reason}: ${el.title}`;
      el.textContent = '?';
    }
  });
}
