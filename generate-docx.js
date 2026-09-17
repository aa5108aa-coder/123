/* docx 8.5.0 + file-saver; shares the exact PDF map snapshot and numbering. */
(() => {
  'use strict';
  let busy=false;
  async function generateDocx({data,meta,tile}, {onProgress=()=>{},download=true}={}) {
    if(busy)throw Error('Word 正在生成中。');
    if(!window.docx?.Packer || (download && !window.saveAs))throw Error('Word 匯出套件尚未載入，請稍後重試。');
    busy=true;
    const progress=(percent,message)=>{try{onProgress({percent,message})}catch(error){console.warn(error)}};
    try {
      const {clean,inline,targetBounds,captureFieldMap}=FieldExport;
      progress(5,'正在生成 Word 文件…');
      const mapImage=await captureFieldMap(targetBounds(data),{data,tile,onProgress:message=>progress(15,message)});
      progress(30,'地圖已完成，正在整理筆記…');
      const {Document,Paragraph,TextRun,ImageRun,HeadingLevel,AlignmentType,Table,TableRow,TableCell,WidthType,Packer}=docx;
      const p=(text,options={})=>new Paragraph({children:[new TextRun(String(text))],spacing:{after:120,line:300},...options});
      const heading=(text,level,options={})=>p(text,{heading:level,keepNext:true,...options});
      const generatedAt=new Date(),children=[heading(meta.name||'田野手札',HeadingLevel.TITLE)];
      const header=[clean(meta.dates),clean(meta.team)].filter(Boolean).join(' · ');if(header)children.push(p(header));
      children.push(p('產出時間：'+generatedAt.toLocaleString('zh-TW',{timeZone:'Asia/Taipei'})));
      const ratio=mapImage.width/mapImage.height,width=Math.min(600,700*ratio),height=width/ratio;
      children.push(new Paragraph({alignment:AlignmentType.CENTER,keepNext:true,children:[new ImageRun({data:mapImage.data,transformation:{width:Math.round(width),height:Math.round(height)}})]}));
      children.push(p('圖 1：踏查範圍與路線軌跡總覽',{alignment:AlignmentType.CENTER}));
      if(mapImage.clippedNotes.length)children.push(p('依研究範圍聚焦；圖外筆記：'+mapImage.clippedNotes.join('、')));
      children.push(heading('踏查筆記',HeadingLevel.HEADING_1,{pageBreakBefore:true}));
      const warnings=[];let completed=0;
      const total=data.notes.reduce((n,note)=>n+1+note.photos.length,0);
      const update=message=>progress(Math.round(30+55*completed/Math.max(1,total)),message);
      if(!data.notes.length)children.push(p('此匯出範圍沒有踏查筆記。'));
      for(const session of data.sessions) {
        const notes=data.notes.filter(n=>n.sessionId===session.id);if(!notes.length)continue;
        children.push(heading(session.name||session.date||'踏查圖層',HeadingLevel.HEADING_2));
        for(const note of notes) {
          children.push(heading(`${note.number} · ${note.title||'踏查筆記'}`,HeadingLevel.HEADING_3));
          const metadata=inline(note);if(metadata)children.push(p(metadata,{spacing:{after:100},children:[new TextRun({text:metadata,size:18,color:'475569'})]}));
          const rows=[['訪談對象',note.interviewee],['現況分類',note.category]].map(([key,value])=>[key,clean(value)]).filter(([,value])=>value);
          if(rows.length)children.push(new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:rows.map(([key,value])=>new TableRow({children:[key,value].map((text,i)=>new TableCell({width:{size:i?78:22,type:WidthType.PERCENTAGE},children:[p(text)],margins:{top:70,bottom:70,left:100,right:100}}))}))}));
          for(const [key,value] of Object.entries(note.sections)) {
            const content=clean(value);if(!content)continue;
            children.push(p(({observation:'現場觀察',interview:'訪談紀錄',questions:'待查問題'})[key]||key,{keepNext:true}));
            content.split(/\r?\n/).forEach(line=>children.push(p(line)));
          }
          for(const [i,source] of note.photos.entries()) {
            update(`正在處理筆記 ${note.number} 照片 ${i+1}…`);
            const caption=`圖 ${note.number}-${i+1}：${note.title||'踏查照片'}`;
            try {
              const img=await FieldWord.prepareImage(source);
              children.push(new Paragraph({alignment:AlignmentType.CENTER,keepNext:true,children:[new ImageRun({data:img.data,transformation:{width:img.width,height:img.height}})]}));
              children.push(p(caption,{alignment:AlignmentType.CENTER}));
            } catch(error) {warnings.push(caption+'：'+error.message);children.push(p(caption+'（照片載入失敗）'));}
            completed++;update('正在排版照片…');
          }
          completed++;update('正在排版筆記…');
          await new Promise(resolve=>setTimeout(resolve,0));
        }
      }
      progress(92,'正在封裝 Word 文件…');
      const document=new Document({title:meta.name||'田野手札',creator:'南台南田野踏查',styles:{default:{document:{run:{font:'Microsoft JhengHei',size:22,color:'000000'}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1134,bottom:1134,left:1134,right:1134}}},children}]});
      const blob=await Packer.toBlob(document),date=generatedAt.toLocaleDateString('sv-SE',{timeZone:'Asia/Taipei'});
      const filename=`田野手札_${String(meta.name||'未命名').replace(/[<>:"/\\|?*\x00-\x1f]/g,'_').slice(0,80)}_${date}.docx`;
      if(download)saveAs(blob,filename);
      progress(100,'文件已生成，'+(download?'已觸發下載。':'可供下載。'));
      return {blob,filename,warnings,mapImage};
    } finally {busy=false;}
  }
  window.generateDocx=generateDocx;
})();
