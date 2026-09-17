/* Shared, read-only export snapshot / bounds / numbered map / PDF pipeline. */
(() => {
  'use strict';
  const clean = value => {
    const text = String(value ?? '').trim();
    return !text || /未填寫|未記錄/.test(text) ? '' : text;
  };
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const color = (v, fallback = '#245e49') => /^#[\da-f]{6}$/i.test(v || '') ? v : fallback;
  const validPoint = p => Array.isArray(p) && p.length >= 2 && p.every(Number.isFinite) && Math.abs(p[0]) <= 90 && Math.abs(p[1]) <= 180;
  const timestamp = n => Number(n.createdAtMs) || (n.createdAt ? Date.parse(n.createdAt) : 0) || Infinity;
  const areaPoints = area => area.points || (Array.isArray(area.bounds) && area.bounds.length===2 ? [area.bounds[0],[area.bounds[0][0],area.bounds[1][1]],area.bounds[1],[area.bounds[1][0],area.bounds[0][1]]] : []);
  function snapshot(input) {
    const copy = JSON.parse(JSON.stringify(input));
    const order = new Map(copy.sessions.map((s, i) => [s.id, i]));
    copy.notes.sort((a,b) => (order.get(a.sessionId) ?? Infinity) - (order.get(b.sessionId) ?? Infinity) || timestamp(a) - timestamp(b) || String(a.id).localeCompare(String(b.id)));
    copy.notes.forEach((n,i) => { n.number = String(i + 1).padStart(2,'0'); n.sections = Object.fromEntries(Object.entries(n.sections || { observation:n.description }).map(([k,v]) => [k,clean(v)]).filter(([,v]) => v)); n.time = clean(n.time); n.author = clean(n.author); });
    return copy;
  }
  function targetBounds(data) {
    // Polygon/BBox coordinates use Leaflet [latitude, longitude].
    const areas = data.areas.map(a=>({...a,points:areaPoints(a)})).filter(a => a.visible !== false && a.points.length >= 3 && a.points.every(validPoint));
    const selected = areas.find(a => a.id === data.activeAreaId);
    const points = (selected ? [selected] : areas).flatMap(a => a.points);
    const bounds = L.latLngBounds(points.length ? points : [...data.routes.flatMap(r => r.coordinates.filter(validPoint)), ...data.notes.map(n => n.coordinates).filter(validPoint)]);
    if (!bounds.isValid()) throw Error('沒有有效的範圍、路線或筆記座標，無法輸出地圖。');
    return bounds;
  }
  function legendHtml(data) {
    const items = [...new Map(data.routes.map(r => {
      const label = data.sessions.find(s => s.id === r.sessionId)?.name || r.name || '踏查路線';
      return [label + color(r.color), { label, color:color(r.color) }];
    })).values()];
    return `<strong>圖面圖例</strong>${data.areas.length ? '<div><span class="fe-line fe-boundary"></span>框選研究範圍</div>' : ''}${items.map(i=>`<div><span class="fe-line" style="border-color:${i.color}"></span>${esc(i.label)}</div>`).join('')}<div><span class="fe-dot">01</span>踏查筆記（對應清單編號）</div>`;
  }
  let capturing = false;
  async function captureFieldMap(bounds, { data, tile, onProgress = () => {} } = {}) {
    if (capturing) throw Error('地圖正在擷取中，請稍候。');
    if (!window.html2canvas) throw Error('html2canvas 尚未載入，請稍後再試。');
    if (!data || !tile?.url) throw Error('請提供匯出快照與底圖設定。');
    capturing = true; let map, stage;
    try {
      const target = bounds ? L.latLngBounds(bounds) : targetBounds(data);
      if (!target.isValid()) throw Error('地圖範圍無效。');
      const sw = L.CRS.EPSG3857.project(target.getSouthWest()), ne = L.CRS.EPSG3857.project(target.getNorthEast());
      const ratio = Math.max(.7, Math.min(1.55, Math.abs(ne.x-sw.x) / (Math.abs(ne.y-sw.y) || 1)));
      const sw17=L.CRS.EPSG3857.latLngToPoint(target.getSouthWest(),17),ne17=L.CRS.EPSG3857.latLngToPoint(target.getNorthEast(),17);
      const extent17=Math.max(Math.abs(ne17.x-sw17.x),Math.abs(ne17.y-sw17.y)*ratio);
      const width = Math.round(Math.max(480,Math.min(1000,extent17/.82))), mapHeight = Math.round(width / ratio);
      stage = document.createElement('div');stage.className='fe-capture';
      // Fixed dimensions independent of the editor viewport or mobile orientation.
      Object.assign(stage.style,{position:'fixed',left:'0',top:'0',width:width+'px',zIndex:'-1',background:'#fff'});
      const host=document.createElement('div');host.style.cssText=`width:${width}px;height:${mapHeight}px;position:relative`;
      stage.append(host);document.body.append(stage);
      map=L.map(host,{zoomControl:false,attributionControl:true,zoomSnap:0,fadeAnimation:false,zoomAnimation:false,markerZoomAnimation:false,preferCanvas:true,dragging:false,touchZoom:false,scrollWheelZoom:false,doubleClickZoom:false,keyboard:false,boxZoom:false});
      data.areas.forEach(a=>{const points=areaPoints(a);if(a.visible!==false&&points.length>=3&&points.every(validPoint))L.polygon(points,{color:color(a.color,'#dc2626'),weight:3,dashArray:'8 8',fill:false,interactive:false}).addTo(map)});
      data.routes.forEach(r=>{const points=r.coordinates.filter(validPoint);if(points.length>=2)L.polyline(points,{color:color(r.color),weight:Math.max(3,Number(r.weight)||5),interactive:false}).addTo(map)});
      data.notes.forEach(n=>{if(validPoint(n.coordinates))L.marker(n.coordinates,{interactive:false,icon:L.divIcon({className:'fe-marker',html:`<span style="background:${color(n.color)}">${esc(n.number)}</span>`,iconSize:[30,30],iconAnchor:[15,15]})}).addTo(map)});
      map.fitBounds(target,{padding:[40,40],maxZoom:17,animate:false});
      // 40px initial padding, then normalize the limiting projected dimension to 82%.
      const a=map.latLngToContainerPoint(target.getSouthWest()),b=map.latLngToContainerPoint(target.getNorthEast());
      const fraction=Math.max(Math.abs(b.x-a.x)/width,Math.abs(b.y-a.y)/mapHeight);
      if(fraction>0)map.setZoom(Math.min(17,map.getZoom()+Math.log2(.82/fraction)),{animate:false});
      L.control.scale({position:'bottomleft',imperial:false,maxWidth:180}).addTo(map);
      const north=L.control({position:'topright'});north.onAdd=()=>{const d=L.DomUtil.create('div','fe-north');d.textContent='N ↑';return d};north.addTo(map);
      const legend=L.control({position:'bottomright'});legend.onAdd=()=>{const d=L.DomUtil.create('div','fe-legend');d.innerHTML=legendHtml(data);return d};legend.addTo(map);
      // All entries also appear below the map; large legends never cover the whole field.
      const floatingLegend=host.querySelector('.fe-legend');
      if(floatingLegend.scrollHeight>floatingLegend.clientHeight){
        const legendFull=document.createElement('div');legendFull.className='fe-legend-full';legendFull.innerHTML=legendHtml(data);stage.append(legendFull);
        floatingLegend.innerHTML='<strong>圖面圖例</strong><div>完整圖層圖例見圖下方</div>';
      }
      const tiles=L.tileLayer(tile.url,{...tile.options,crossOrigin:'anonymous',attribution:tile.attribution || tile.options?.attribution || '',maxZoom:tile.options?.maxZoom || 19});
      onProgress('正在載入底圖與路線…');
      await new Promise((resolve,reject)=>{
        const timer=setTimeout(()=>done(Error('底圖載入逾時，請檢查網路後重試。')),20000);
        const done=error=>{clearTimeout(timer);tiles.off('load',loaded);tiles.off('tileerror',failed);error?reject(error):resolve()};
        const loaded=()=>done(),failed=()=>done(Error('底圖圖片載入或 CORS 驗證失敗；未產生缺圖文件。'));
        tiles.once('load',loaded);tiles.once('tileerror',failed);tiles.addTo(map);
      });
      await Promise.all([...host.querySelectorAll('img')].map(img=>img.decode ? img.decode() : Promise.resolve()));
      await document.fonts?.ready;
      await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
      const p=map.latLngToContainerPoint(target.getSouthWest()),q=map.latLngToContainerPoint(target.getNorthEast());
      const occupancy=Math.max(Math.abs(q.x-p.x)/width,Math.abs(q.y-p.y)/mapHeight);
      const clipped=data.notes.filter(n=>validPoint(n.coordinates)&&!map.getBounds().contains(n.coordinates)).map(n=>n.number);
      onProgress('正在擷取高解析度地圖…');
      const canvas=await html2canvas(stage,{useCORS:true,allowTaint:false,scale:2000/width,backgroundColor:'#fff',logging:false,width,height:stage.offsetHeight,windowWidth:Math.max(width,innerWidth),windowHeight:Math.max(stage.offsetHeight,innerHeight)});
      const png=canvas.toDataURL('image/png'),blob=await new Promise(resolve=>canvas.toBlob(resolve,'image/png'));
      if(!blob)throw Error('無法產生地圖 PNG。');
      return {png,data:new Uint8Array(await blob.arrayBuffer()),width:canvas.width,height:canvas.height,occupancy,clippedNotes:clipped,zoom:map.getZoom()};
    } finally {map?.remove();stage?.remove();capturing=false;}
  }
  const inline = n => [clean(n.time), validPoint(n.coordinates) ? `${n.coordinates[0].toFixed(6)}, ${n.coordinates[1].toFixed(6)}` : '',clean(n.author)].filter(Boolean).join(' · ');
  let printing=false;
  async function generatePdf(data, meta, capture, {mapOnly=false,preview=false}={}) {
    if(printing)throw Error('已有列印工作，請先關閉列印視窗。');printing=true;
    const root=document.getElementById('print-report-container');
    const labels={observation:'現場觀察',interview:'訪談紀錄',questions:'待查問題'};
    let style;
    const cleanup=()=>{document.body.classList.remove('report-printing');root.replaceChildren();style?.remove();printing=false;window.removeEventListener('afterprint',cleanup)};
    try {
      const caption='圖 1：踏查範圍與路線軌跡總覽';
      root.className='fe-report';root.style.width='auto';
      root.innerHTML=`<section class="fe-cover"><h1>${esc(meta.name)}</h1><p>${[clean(meta.dates),clean(meta.team),clean(meta.weather),clean(meta.remarks)].filter(Boolean).map(esc).join(' · ')}</p><img class="fe-overview" src="${capture.png}" alt="${caption}"><p>${caption}</p>${capture.clippedNotes.length?`<p>依研究範圍聚焦；圖外筆記：${capture.clippedNotes.map(esc).join('、')}</p>`:''}</section>${mapOnly?`<section class="fe-note-index"><h2>筆記編號對照</h2>${data.notes.map(n=>`<p>${n.number} · ${esc(n.title)}</p>`).join('')}</section>`:data.notes.map(n=>`<article class="fe-note"><h2>${n.number} · ${esc(n.title)}</h2><p class="fe-meta">${esc(inline(n))}</p>${[['訪談對象',n.interviewee],['現況分類',n.category]].filter(([,v])=>clean(v)).map(([k,v])=>`<p>${k}：${esc(clean(v))}</p>`).join('')}${Object.entries(n.sections).filter(([,v])=>clean(v)).map(([k,v])=>`<h3>${labels[k]||esc(k)}</h3><p class="fe-text">${esc(v)}</p>`).join('')}${n.photos.map((src,i)=>`<figure><img src="${esc(src)}" alt="${esc(n.title)}"><figcaption>圖 ${n.number}-${i+1}：${esc(n.title)}</figcaption></figure>`).join('')}</article>`).join('')}`;
      await Promise.all([...root.querySelectorAll('img')].map(img=>new Promise(resolve=>{const done=()=>{clearTimeout(timer);if(!img.naturalWidth){const p=document.createElement('p');p.textContent='照片載入失敗：'+img.alt;img.replaceWith(p)}resolve()};const timer=setTimeout(done,15000);if(img.complete)done();else img.onload=img.onerror=done})));
      style=document.createElement('style');style.textContent=`@media print{@page{size:A4 ${meta.orientation==='landscape'?'landscape':'portrait'};margin:12mm}.fe-overview{max-height:${meta.orientation==='landscape'?'145mm':'225mm'}}}`;document.head.append(style);
      document.body.classList.add('report-printing');window.addEventListener('afterprint',cleanup);
      await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
      if(preview)return {cleanup,root};
      window.print();
    }catch(error){cleanup();throw error}
  }
  window.FieldExport={clean,esc,color,validPoint,snapshot,targetBounds,captureFieldMap,inline,generatePdf};
  window.captureFieldMap=captureFieldMap;
})();
