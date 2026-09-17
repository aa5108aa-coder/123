/* Browser UMD dependencies: docx 8.5.0, file-saver 2.0.5. */
(() => {
  'use strict';
  const MAX_BYTES = 15 * 1024 * 1024;
  const text = value => String(value ?? '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  function date(value) {
    if (value == null || value === '') return null;
    if (typeof value.toDate === 'function') value = value.toDate();
    else if (typeof value === 'object' && Number.isFinite(value.seconds)) value = value.seconds * 1000;
    const result = value instanceof Date ? value : new Date(value);
    return Number.isFinite(result.getTime()) ? result : null;
  }
  function stamp(value, timeZone, dateOnly = false) {
    const d = date(value); if (!d) return '未記錄';
    return new Intl.DateTimeFormat('zh-TW', {
      timeZone, year: 'numeric', month: '2-digit', day: '2-digit',
      ...(dateOnly ? {} : { hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
    }).format(d);
  }
  function filenamePart(value) {
    return text(value).replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').replace(/[. ]+$/g, '').slice(0, 80) || '未命名';
  }
  function normalize(input) {
    if (!input || !Array.isArray(input.layers)) throw new Error('手札資料必須包含 layers 陣列。');
    return {
      name: text(input.name).trim() || '未命名專案',
      surveyDates: (Array.isArray(input.surveyDates) ? input.surveyDates : [input.surveyDates]).filter(v => v != null && v !== '').map(v => date(v) || text(v)),
      layers: input.layers.map((layer, index) => ({
        name: text(layer.name) || `圖層 ${index + 1}`, order: Number.isFinite(layer.order) ? layer.order : index, index,
        notes: (Array.isArray(layer.notes) ? layer.notes : []).map((note, position) => ({
          title: text(note.title) || '未命名筆記', createdAt: date(note.createdAt), position,
          content: text(note.content), longitude: note.longitude, latitude: note.latitude,
          metadata: Object.entries(note.metadata || {}).map(([key, value]) => [text(key), text(value)]),
          photos: (Array.isArray(note.photos) ? note.photos : []).map(photo =>
            typeof photo === 'string' || photo instanceof Blob ? { source: photo, caption: '' } : { source: photo.source, caption: text(photo.caption) })
        })).sort((a, b) => (a.createdAt?.getTime() ?? Infinity) - (b.createdAt?.getTime() ?? Infinity) || a.position - b.position)
      })).sort((a, b) => a.order - b.order || a.index - b.index)
    };
  }
  async function imageBlob(source) {
    if (source instanceof Blob) {
      if (source.size > MAX_BYTES) throw new Error('照片超過 15 MB');
      return source;
    }
    const url = new URL(source, document.baseURI);
    if (!['https:', 'http:', 'blob:', 'data:'].includes(url.protocol)) throw new Error('不支援的圖片網址');
    const controller = new AbortController(), timeout = setTimeout(() => controller.abort(), 15000);
    try {
      const response = await fetch(url.href, { signal: controller.signal, mode: 'cors', credentials: 'same-origin' });
      if (!response.ok) throw new Error(`圖片回應 HTTP ${response.status}`);
      if (Number(response.headers.get('content-length')) > MAX_BYTES) throw new Error('照片超過 15 MB');
      // 串流限制實際下載量，不能僅信任 Content-Length。
      if (!response.body) {
        const blob = await response.blob(); if (blob.size > MAX_BYTES) throw new Error('照片超過 15 MB'); return blob;
      }
      const reader = response.body.getReader(), chunks = []; let size = 0;
      try {
        while (true) {
          const { value, done } = await reader.read(); if (done) break;
          size += value.byteLength;
          if (size > MAX_BYTES) { await reader.cancel(); throw new Error('照片超過 15 MB'); }
          chunks.push(value);
        }
      } finally { reader.releaseLock(); }
      return new Blob(chunks, { type: response.headers.get('content-type') || '' });
    } finally { clearTimeout(timeout); }
  }
  async function prepareImage(source) {
    const blob = await imageBlob(source);
    const url = URL.createObjectURL(blob), image = new Image();
    try {
      await new Promise((resolve, reject) => {
        const timeout = setTimeout(() => { image.src = ''; reject(new Error('圖片解碼逾時')); }, 15000);
        image.onload = () => { clearTimeout(timeout); resolve(); };
        image.onerror = () => { clearTimeout(timeout); reject(new Error('圖片格式無法解碼')); };
        image.src = url;
      });
      const w = image.naturalWidth, h = image.naturalHeight;
      if (!w || !h || w * h > 40000000) throw new Error('圖片尺寸無效或超過 4,000 萬像素');
      // 顯示最大 480×560，保留比例；實際嵌入最多兩倍解析度。
      const ratio = Math.min(1, 480 / w, 560 / h);
      const width = Math.max(1, Math.round(w * ratio)), height = Math.max(1, Math.round(h * ratio));
      const canvas = document.createElement('canvas');
      canvas.width = Math.min(w, width * 2); canvas.height = Math.min(h, height * 2);
      const ctx = canvas.getContext('2d'); if (!ctx) throw new Error('瀏覽器無法處理圖片');
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
      const output = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', .88));
      canvas.width = canvas.height = 0;
      if (!output) throw new Error('圖片轉換失敗');
      return { data: new Uint8Array(await output.arrayBuffer()), width, height };
    } finally { URL.revokeObjectURL(url); image.onload = image.onerror = null; }
  }
  function coordinates(note) {
    const valid = v => v !== null && v !== undefined && v !== '' && Number.isFinite(Number(v));
    if (!valid(note.longitude) || !valid(note.latitude) || Math.abs(Number(note.longitude)) > 180 || Math.abs(Number(note.latitude)) > 90) return '未記錄或座標無效';
    return `經度 ${Number(note.longitude).toFixed(6)} / 緯度 ${Number(note.latitude).toFixed(6)}`;
  }
  let exporting = false;
  async function exportDocx(input, { onProgress = () => {}, timeZone = 'Asia/Taipei', download = true } = {}) {
    if (exporting) throw new Error('Word 文件正在生成中，請稍候。');
    const d = window.docx;
    if (!d?.Document || !d?.Packer) throw new Error('docx 未載入，請檢查網路或 CDN。');
    if (download && typeof window.saveAs !== 'function') throw new Error('file-saver 未載入，無法下載。');
    exporting = true;
    const report = (percent, message) => { try { onProgress({ percent, message }); } catch (error) { console.warn('進度顯示失敗', error); } };
    try {
      const project = normalize(input), generatedAt = new Date(), warnings = [];
      const { Document, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ImageRun, Packer } = d;
      const paragraph = (value, options = {}) => new Paragraph({ children: [new TextRun(text(value))], spacing: { after: 140, line: 320 }, ...options });
      const heading = (value, level) => paragraph(value, { heading: level, keepNext: true });
      const children = [heading(project.name, HeadingLevel.TITLE), paragraph('田野手札'),
        paragraph(`踏查日期：${project.surveyDates.map(v => v instanceof Date ? stamp(v, timeZone, true) : v).join('、') || '未記錄'}`),
        paragraph(`產出時間：${stamp(generatedAt, timeZone)}（${timeZone}）`), heading('踏查筆記', HeadingLevel.HEADING_1)];
      const count = project.layers.reduce((sum, layer) => sum + layer.notes.length, 0);
      const total = project.layers.reduce((sum, layer) => sum + layer.notes.reduce((n, note) => n + 1 + note.photos.length, 0), 0);
      let done = 0;
      const progress = message => report(Math.round(5 + (done / Math.max(1, total)) * 80), message);
      report(2, '正在整理圖層與筆記...');
      if (!count) children.push(paragraph('此專案尚無踏查筆記。'));
      for (const layer of project.layers) {
        children.push(heading(layer.name, HeadingLevel.HEADING_2));
        if (!layer.notes.length) children.push(paragraph('此圖層尚無筆記。'));
        for (const note of layer.notes) {
          children.push(heading(note.title, HeadingLevel.HEADING_3));
          const rows = [['記錄時間', `${stamp(note.createdAt, timeZone)}`], ['經緯度座標', coordinates(note)], ...note.metadata];
          children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE },
            columnWidths: [2100, 7538], rows: rows.map(([key, value]) => new TableRow({
              // 保留長內容跨頁能力，避免訪談對象或分類文字被截斷。
              children: [key, value].map((value, i) => new TableCell({
                width: { size: i ? 7538 : 2100, type: WidthType.DXA },
                ...(i ? {} : { shading: { fill: 'EDF2EF' } }),
                margins: { top: 100, bottom: 100, left: 120, right: 120 },
                children: text(value || '未記錄').split(/\r?\n/).map(line => paragraph(line))
              }))
            })) }));
          children.push(paragraph('觀察／訪談內容', { keepNext: true, spacing: { before: 180, after: 100 } }));
          text(note.content || '未填寫').split(/\r?\n/).forEach(line => children.push(paragraph(line)));
          for (const [index, photo] of note.photos.entries()) {
            progress(`正在處理「${note.title}」照片 ${index + 1}/${note.photos.length}...`);
            try {
              const img = await prepareImage(photo.source);
              children.push(new Paragraph({ children: [new ImageRun({ data: img.data, transformation: { width: img.width, height: img.height } })], spacing: { before: 160, after: 100 }, keepNext: !!photo.caption }));
              if (photo.caption) children.push(paragraph(photo.caption));
            } catch (error) {
              const message = `${layer.name}／${note.title}：照片 ${index + 1} 未嵌入（${error.message}）`;
              warnings.push(message); children.push(paragraph(`[${message}]`));
            }
            done++; progress('照片處理完成');
          }
          done++; progress(`已整理「${note.title}」`);
          await new Promise(resolve => setTimeout(resolve, 0));
        }
      }
      report(90, '正在封裝 Word 文件...');
      const doc = new Document({ title: `${project.name} 田野手札`, creator: '南台南田野踏查',
        styles: { default: { document: { run: { font: 'Microsoft JhengHei', size: 22, color: '000000' } } } },
        sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } }, children }]
      });
      const blob = await Packer.toBlob(doc);
      const day = new Intl.DateTimeFormat('en-CA', { timeZone, year: 'numeric', month: '2-digit', day: '2-digit' }).format(generatedAt).replaceAll('/', '-');
      const filename = `田野手札_${filenamePart(project.name)}_${day}.docx`;
      if (download) window.saveAs(blob, filename);
      report(100, download ? '文件已生成，已觸發下載。' : '文件已生成。');
      return { blob, filename, warnings, noteCount: count };
    } finally { exporting = false; }
  }
  window.FieldWord = Object.freeze({ exportDocx, normalize });
})();
