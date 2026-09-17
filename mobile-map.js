/* Leaflet integration. Remove the old drawing click/touchend listeners before installing. */
(() => {
  'use strict';
  // No preventDefault/capture on the map: Leaflet continues handling pan/pinch.
  function installTapGate(element, { enabled, onTap, delay = 300 }) {
    const controller = new AbortController(), options = { signal: controller.signal };
    const pointers = new Map(); let pending = null, cooldown = 0;
    const cancel = () => { clearTimeout(pending); pending = null; };
    const excluded = event => event.target.closest?.('.leaflet-control,.leaflet-marker-icon,.leaflet-popup,button,a,input,select,textarea,[data-map-interactive]');
    element.addEventListener('pointerdown', event => {
      if (!enabled() || event.button !== 0 || excluded(event)) return;
      const record = { x: event.clientX, y: event.clientY, time: performance.now(), bad: false };
      // A second tap within the settling window cancels the first; no double-tap points.
      if (pending) { cancel(); record.bad = true; }
      pointers.set(event.pointerId, record);
      if (pointers.size > 1) { cancel(); pointers.forEach(p => p.bad = true); }
    }, options);
    window.addEventListener('pointermove', event => {
      const p = pointers.get(event.pointerId); if (!p) return;
      if (Math.hypot(event.clientX - p.x, event.clientY - p.y) > 10) p.bad = true;
    }, options);
    window.addEventListener('pointerup', event => {
      const p = pointers.get(event.pointerId); if (!p) return; pointers.delete(event.pointerId);
      const now = performance.now();
      if (!enabled() || p.bad || pointers.size || now - p.time > 650 || now < cooldown || Math.hypot(event.clientX - p.x, event.clientY - p.y) > 10 || !element.contains(event.target) || excluded(event)) return;
      const point = { clientX: event.clientX, clientY: event.clientY };
      pending = setTimeout(() => { pending = null; if (enabled() && !pointers.size) onTap(point); }, delay);
    }, options);
    window.addEventListener('pointercancel', event => { pointers.delete(event.pointerId); pointers.forEach(p => p.bad = true); cancel(); cooldown = performance.now() + 350; }, options);
    element.addEventListener('wheel', () => { cancel(); pointers.forEach(p => p.bad = true); }, { ...options, passive: true });
    const reset = () => { cancel(); pointers.clear(); cooldown = performance.now() + 100; };
    window.addEventListener('blur', reset, options);
    return { reset, destroy() { reset(); controller.abort(); } };
  }
  function connectLeaflet({ map, L, saveRoute, saveNote, area = {} }) {
    let mode = null, layerId, points = [], line, busy = false;
    function render() {
      if (line) line.setLatLngs(points);
      else if (points.length) line = L.polyline(points, { color: '#245e49', weight: 5, interactive: false }).addTo(map);
      window.FieldMobile.setRouteCount(points.length);
    }
    function clear() { points = []; if (line) { map.removeLayer(line); line = null; } window.FieldMobile.setRouteCount(0); }
    const gate = installTapGate(map.getContainer(), {
      enabled: () => !busy && ['route', 'note'].includes(mode) && !document.querySelector('dialog[open]'),
      onTap: event => {
        const latlng = map.mouseEventToLatLng(event);
        if (mode === 'route') { points.push({ lat: latlng.lat, lng: latlng.lng }); render(); }
        if (mode === 'note') window.FieldMobile.openNote({ coordinate: { lat: latlng.lat, lng: latlng.lng }, layerId,
          onSave: async note => {
            if (typeof saveNote !== 'function') throw new Error('尚未接入筆記儲存。');
            await saveNote(note);
          }
        });
      }
    });
    const doubleClickInitially = map.doubleClickZoom?.enabled();
    function restoreDoubleClick() { if (doubleClickInitially) map.doubleClickZoom.enable(); }
    const modeListener = event => {
      if (!event.detail.mode) { mode = null; gate.reset(); clear(); restoreDoubleClick(); }
    };
    document.addEventListener('field:mode', modeListener);
    window.FieldUI.connect({
      'start-edit': async payload => {
        if (mode || busy) throw new Error('請先結束目前編輯。');
        if (payload.mode === 'area') {
          if (!area.start) throw new Error('尚未接入範圍繪製。');
          await area.start(payload);
        }
        clear(); gate.reset(); mode = payload.mode; layerId = payload.layerId;
        if (['route', 'note'].includes(mode)) map.doubleClickZoom?.disable();
      },
      'undo-edit': () => { if (mode === 'route' && !busy) { gate.reset(); points.pop(); render(); } },
      'finish-edit': async () => {
        if (busy) throw new Error('正在儲存，請稍候。');
        gate.reset();
        if (mode === 'route' && points.length < 2) throw new Error('請至少加入兩個路線點。');
        busy = true;
        try {
          if (mode === 'route') {
            if (typeof saveRoute !== 'function') throw new Error('尚未接入路線儲存。');
            await saveRoute({ layerId, points: points.map(point => ({ ...point })) });
          } else if (mode === 'area' && area.finish) await area.finish();
          else throw new Error('請在筆記表單內儲存，或確認繪製功能已接入。');
          clear(); mode = null; restoreDoubleClick();
        } finally { busy = false; }
      },
      'cancel-edit': async () => {
        if (busy) throw new Error('正在儲存，請稍候。');
        if (mode === 'area') { if (!area.cancel) throw new Error('尚未接入範圍取消。'); await area.cancel(); }
        mode = null; gate.reset(); clear(); restoreDoubleClick();
      }
    });
    return { destroy() { gate.destroy(); clear(); mode = null; restoreDoubleClick(); document.removeEventListener('field:mode', modeListener); } };
  }
  window.FieldMobileMap = { installTapGate, connectLeaflet };
})();
