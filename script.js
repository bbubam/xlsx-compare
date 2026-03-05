let fileA=null,fileB=null,wbA=null,wbB=null,allChanges=[],sheetDiffs={};

function resetApp(){
  fileA=null;fileB=null;wbA=null;wbB=null;allChanges=[];sheetDiffs={};
  ['f1','f2'].forEach(id=>document.getElementById(id).value='');
  ['fn1','fn2'].forEach(id=>document.getElementById(id).textContent='Drop or click to select');
  document.getElementById('cmpBtn').disabled=true;
  document.getElementById('viewTabs').style.display='none';
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.view-tab').forEach((t,i)=>t.classList.toggle('active',i===0));
  document.getElementById('placeholder').style.display='flex';
}

function setupDZ(dzId,inputId,fnId,slot){
  const dz=document.getElementById(dzId),inp=document.getElementById(inputId);
  inp.addEventListener('change',e=>{if(e.target.files[0])setFile(slot,e.target.files[0],fnId)});
  dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('over')});
  dz.addEventListener('dragleave',()=>dz.classList.remove('over'));
  dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('over');const f=e.dataTransfer.files[0];if(f?.name.endsWith('.xlsx'))setFile(slot,f,fnId)});
}
function setFile(slot,file,fnId){
  if(slot==='A')fileA=file;else fileB=file;
  document.getElementById(fnId).textContent=file.name;
  const ready=fileA&&fileB;
  document.getElementById('cmpBtn').disabled=!ready;
  if(ready)document.getElementById('cmpBtn').click();
}
setupDZ('dz1','f1','fn1','A');
setupDZ('dz2','f2','fn2','B');

document.querySelectorAll('.view-tab').forEach(t=>t.addEventListener('click',()=>switchView(t.dataset.view)));
function switchView(v){
  document.querySelectorAll('.view-tab').forEach(t=>t.classList.toggle('active',t.dataset.view===v));
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));
  document.getElementById('p'+v[0].toUpperCase()+v.slice(1)).classList.add('active');
}

async function readWB(file){const buf=await file.arrayBuffer();return XLSX.read(buf,{type:'array',cellFormula:true});}

function wsToRows(ws){
  if(!ws['!ref'])return[];
  const rng=XLSX.utils.decode_range(ws['!ref']),rows=[];
  for(let r=rng.s.r;r<=rng.e.r;r++){
    const row=[];
    for(let c=rng.s.c;c<=rng.e.c;c++){
      const cell=ws[XLSX.utils.encode_cell({r,c})];
      row.push(cell?(cell.f?'='+cell.f:(cell.v!==undefined?String(cell.v):'')):'');
    }
    rows.push(row);
  }
  return rows;
}

// Auto-detect best key: find smallest set of columns that is unique across all rows
function detectKey(rows){
  if(rows.length<=1)return null; // header only
  const data=rows.slice(1);
  const numCols=data[0]?.length||0;
  // try each single column first
  for(let c=0;c<numCols;c++){
    const vals=data.map(r=>r[c]??'');
    if(new Set(vals).size===vals.length) return [c];
  }
  // try pairs
  for(let a=0;a<numCols;a++)for(let b=a+1;b<numCols;b++){
    const vals=data.map(r=>(r[a]??'')+'|'+(r[b]??''));
    if(new Set(vals).size===vals.length) return [a,b];
  }
  // fallback: entire row
  return null;
}

function rowKey(row,keyCols){
  if(!keyCols)return row.join('\x00');
  return keyCols.map(c=>row[c]??'').join('\x00');
}

document.getElementById('cmpBtn').addEventListener('click',async()=>{
  [wbA,wbB]=await Promise.all([readWB(fileA),readWB(fileB)]);
  const sA=new Set(wbA.SheetNames),sB=new Set(wbB.SheetNames);
  const allSheets=[...new Set([...wbA.SheetNames,...wbB.SheetNames])];
  allChanges=[];sheetDiffs={};

  allSheets.forEach(name=>{
    const status=!sA.has(name)?'added':!sB.has(name)?'deleted':'changed';
    const rowsA=sA.has(name)?wsToRows(wbA.Sheets[name]):[];
    const rowsB=sB.has(name)?wsToRows(wbB.Sheets[name]):[];
    const headers=(rowsA[0]||rowsB[0]||[]);
    const numCols=Math.max(headers.length,...(rowsA[1]?[rowsA[1].length]:[]),...(rowsB[1]?[rowsB[1].length]:[]));

    // detect key from whichever file has more rows
    const keyCols=detectKey(rowsA.length>=rowsB.length?rowsA:rowsB);

    const mapA=new Map(),mapB=new Map();
    for(let i=1;i<rowsA.length;i++){const k=rowKey(rowsA[i],keyCols);if(!mapA.has(k))mapA.set(k,{row:rowsA[i],i});}
    for(let i=1;i<rowsB.length;i++){const k=rowKey(rowsB[i],keyCols);if(!mapB.has(k))mapB.set(k,{row:rowsB[i],i});}

    const rowDiffs=[];
    // header always first
    rowDiffs.push({rowType:'header',rowA:headers,rowB:headers,changedCols:new Set()});

    let addedRows=0,deletedRows=0,modifiedRows=0;

    mapA.forEach(({row:rowA},k)=>{
      if(mapB.has(k)){
        const rowB=mapB.get(k).row;
        const changedCols=new Set();
        for(let c=0;c<Math.max(rowA.length,rowB.length);c++){
          if((rowA[c]??'')!==(rowB[c]??''))changedCols.add(c);
        }
        if(changedCols.size){
          rowDiffs.push({rowType:'modified',rowA,rowB,changedCols});
          modifiedRows++;
          changedCols.forEach(c=>allChanges.push({sheet:name,type:'modified',col:c,header:headers[c]??c,va:rowA[c]??'',vb:rowB[c]??'',key:k}));
        } else {
          rowDiffs.push({rowType:'same',rowA,rowB,changedCols:new Set()});
        }
      } else {
        rowDiffs.push({rowType:'deleted',rowA,rowB:null,changedCols:new Set()});
        deletedRows++;
        allChanges.push({sheet:name,type:'deleted',col:-1,header:'(row)',va:rowA.join('\t'),vb:'',key:k});
      }
    });

    mapB.forEach(({row:rowB},k)=>{
      if(!mapA.has(k)){
        rowDiffs.push({rowType:'added',rowA:null,rowB,changedCols:new Set()});
        addedRows++;
        allChanges.push({sheet:name,type:'added',col:-1,header:'(row)',va:'',vb:rowB.join('\t'),key:k});
      }
    });

    sheetDiffs[name]={rowDiffs,status,headers,numCols,addedRows,deletedRows,modifiedRows};
  });

  buildSummary(allSheets);
  buildGrid(allSheets);
  buildChanges(allSheets);
  document.getElementById('placeholder').style.display='none';
  document.getElementById('viewTabs').style.display='flex';
  switchView('summary');
});

// ── 1. SUMMARY ──
function buildSummary(sheets){
  const totAdd=allChanges.filter(c=>c.type==='added').length;
  const totDel=allChanges.filter(c=>c.type==='deleted').length;
  const totMod=sheets.reduce((s,n)=>s+(sheetDiffs[n].modifiedRows||0),0);
  const changedSheets=sheets.filter(n=>{const d=sheetDiffs[n];return d.addedRows||d.deletedRows||d.modifiedRows||d.status!=='changed';}).length;
  document.getElementById('sumOverview').innerHTML=
    `<div class="ov-stat ov-sh"><span class="big">${changedSheets}</span><span>sheets<br>changed</span></div>`+
    `<div class="ov-stat ov-add"><span class="big">${totAdd}</span><span>rows<br>added</span></div>`+
    `<div class="ov-stat ov-del"><span class="big">${totDel}</span><span>rows<br>deleted</span></div>`+
    `<div class="ov-stat ov-mod"><span class="big">${totMod}</span><span>rows<br>modified</span></div>`;

  const cards=document.getElementById('sumCards');
  cards.innerHTML='';
  sheets.forEach(name=>{
    const {addedRows,deletedRows,modifiedRows,status}=sheetDiffs[name];
    const total=addedRows+deletedRows+modifiedRows;
    const cardCls=status==='added'?'c-add':status==='deleted'?'c-del':total?'c-mod':'c-ok';
    const statusBadge=status==='added'?'<span class="badge b-add">New Sheet</span>':status==='deleted'?'<span class="badge b-del">Deleted Sheet</span>':total?'<span class="badge b-mod">Changed</span>':'<span class="b-none">Unchanged</span>';
    const clickable=total||status!=='changed'?`onclick="switchView('grid');activateSheet('${esc(name)}')" style="cursor:pointer"`:'style="cursor:default"';
    cards.innerHTML+=`<div class="card ${cardCls}" ${clickable}>
      <div class="card-top"><span class="card-name" title="${esc(name)}">${esc(name)}</span>${statusBadge}</div>
      <div class="card-stats">
        <div class="stat s-add"><span class="num">${addedRows}</span><span class="lbl">Added</span></div>
        <div class="stat s-del"><span class="num">${deletedRows}</span><span class="lbl">Deleted</span></div>
        <div class="stat s-mod"><span class="num">${modifiedRows}</span><span class="lbl">Modified</span></div>
        <div class="stat s-tot"><span class="num">${total}</span><span class="lbl">Total</span></div>
      </div>
      <div class="card-footer"><span>${total?'Click to view in grid':'No changes'}</span></div>
    </div>`;
  });
}

// ── 2. SHEET GRID ──
let hideUnchanged=false,currentSheet=null,lockedLeft=new Set(),lockedRight=new Set(),lockedRows=new Set();

function buildGrid(sheets){
  const tabBar=document.getElementById('sheetTabs'),wrap=document.getElementById('gridWrap');
  tabBar.innerHTML='';wrap.innerHTML='';
  sheets.forEach((name,i)=>{
    const {addedRows,deletedRows,modifiedRows,status}=sheetDiffs[name];
    const total=addedRows+deletedRows+modifiedRows;
    const cls=status==='added'?'sadd':status==='deleted'?'sdel':total?'chg':'';
    const tab=document.createElement('div');
    tab.className='stab '+cls;tab.textContent=name;tab.dataset.sheet=name;
    if(i===0)tab.classList.add('active');
    tab.addEventListener('click',()=>activateSheet(name));
    tabBar.appendChild(tab);
    const panel=document.createElement('div');
    panel.id='sp-'+name;panel.style.display=i===0?'block':'none';
    wrap.appendChild(panel);
  });
  lockedLeft=new Set();lockedRight=new Set();lockedRows=new Set();
  currentSheet=sheets[0];
  refreshGrid();
}

function activateSheet(name){
  currentSheet=name;lockedLeft=new Set();lockedRight=new Set();lockedRows=new Set();
  document.querySelectorAll('.stab').forEach(t=>t.classList.toggle('active',t.dataset.sheet===name));
  document.querySelectorAll('#gridWrap > div').forEach(p=>p.style.display=p.id==='sp-'+name?'block':'none');
  refreshGrid();
}

function toggleHideUnchanged(){
  hideUnchanged=!hideUnchanged;
  document.getElementById('hideUnchangedBtn').classList.toggle('active',hideUnchanged);
  refreshGrid();
}

function clearLocks(){lockedLeft=new Set();lockedRight=new Set();lockedRows=new Set();refreshGrid();}

function refreshGrid(){
  if(!currentSheet)return;
  const panel=document.getElementById('sp-'+currentSheet);
  if(panel){panel.innerHTML=renderGrid(currentSheet);attachGridEvents(panel);}
}

function renderGrid(name){
  const {rowDiffs,status,headers,numCols}=sheetDiffs[name];
  if(status==='added')return`<div class="no-chg" style="color:#5ddb9e">Sheet added in File B</div>`;
  if(status==='deleted')return`<div class="no-chg" style="color:#ff6b6b">Sheet only in File A (deleted)</div>`;
  if(!rowDiffs.length)return`<div class="no-chg">Empty sheet</div>`;

  const DEF_COL=120,RN_W=50;
  const visibleRows=hideUnchanged?rowDiffs.filter(r=>r.rowType!=='same'):rowDiffs;

  // changed cols set (for hide unchanged cols)
  const changedColsSet=new Set();
  rowDiffs.forEach(rd=>{if(rd.changedCols)rd.changedCols.forEach(c=>changedColsSet.add(c));});

  const cols=[];
  for(let c=0;c<numCols;c++){
    if(hideUnchanged&&!lockedLeft.has(c)&&!lockedRight.has(c)&&!changedColsSet.has(c))continue;
    cols.push(c);
  }

  // left/right sticky offsets
  const leftOffMap={};let runL=RN_W;
  cols.forEach((c,ci)=>{if(lockedLeft.has(ci)){leftOffMap[ci]=runL;runL+=DEF_COL;}});
  const rightOffMap={};let runR=0;
  for(let i=cols.length-1;i>=0;i--){if(lockedRight.has(i)){rightOffMap[i]=runR;runR+=DEF_COL;}}

  let html=`<table><thead><tr>`;
  html+=`<th class="rn" style="left:0;z-index:3;width:${RN_W}px">#</th>`;
  cols.forEach((c,ci)=>{
    const isL=lockedLeft.has(ci),isR=lockedRight.has(ci);
    const cls=isL?'cl':isR?'cr-pin':'';
    const sticky=isL?` style="position:sticky;left:${leftOffMap[ci]}px;z-index:3"`:isR?` style="position:sticky;right:${rightOffMap[ci]}px;z-index:3"`:'';
    const thChg=changedColsSet.has(c)?' th-chg':'';
    html+=`<th class="${cls}${thChg}" data-ci="${ci}"${sticky}>${esc(headers[c]||XLSX.utils.encode_col(c))}<span class="resize-h" data-rh="${ci}"></span></th>`;
  });
  html+='</tr></thead><tbody>';

  visibleRows.forEach((rd,ri)=>{
    const {rowType,rowA,rowB,changedCols}=rd;
    if(rowType==='header')return; // already in thead

    const isLockedRow=lockedRows.has(ri);
    const rowCls=isLockedRow?'row-locked':'';

    if(rowType==='deleted'){
      html+=`<tr class="${rowCls} rr-row" data-ri="${ri}"><td class="rn rn-chg" data-ri="${ri}" style="left:0">${ri}<span class="resize-v" data-rv="${ri}"></span></td>`;
      cols.forEach((c,ci)=>{
        const isL=lockedLeft.has(ci),isR=lockedRight.has(ci);
        let cls='cr'+(isL?' cl':isR?' cr-pin':'');
        const sticky=isL?` style="position:sticky;left:${leftOffMap[ci]}px"`:isR?` style="position:sticky;right:${rightOffMap[ci]}px"`:'';
        html+=`<td class="${cls}"${sticky}>${esc(rowA[c]??'')}</td>`;
      });
      html+='</tr>';
    } else if(rowType==='added'){
      html+=`<tr class="${rowCls} ra-row" data-ri="${ri}"><td class="rn rn-chg" data-ri="${ri}" style="left:0">${ri}<span class="resize-v" data-rv="${ri}"></span></td>`;
      cols.forEach((c,ci)=>{
        const isL=lockedLeft.has(ci),isR=lockedRight.has(ci);
        let cls='ca'+(isL?' cl':isR?' cr-pin':'');
        const sticky=isL?` style="position:sticky;left:${leftOffMap[ci]}px"`:isR?` style="position:sticky;right:${rightOffMap[ci]}px"`:'';
        html+=`<td class="${cls}"${sticky}>${esc(rowB[c]??'')}</td>`;
      });
      html+='</tr>';
    } else if(rowType==='modified'){
      html+=`<tr class="${rowCls}" data-ri="${ri}"><td class="rn rn-chg" data-ri="${ri}" style="left:0">${ri}<span class="resize-v" data-rv="${ri}"></span></td>`;
      cols.forEach((c,ci)=>{
        const isL=lockedLeft.has(ci),isR=lockedRight.has(ci);
        const sticky=isL?` style="position:sticky;left:${leftOffMap[ci]}px"`:isR?` style="position:sticky;right:${rightOffMap[ci]}px"`:'';
        if(changedCols.has(c)){
          let cls='cm'+(isL?' cl':isR?' cr-pin':'');
          html+=`<td class="${cls}"${sticky}><span class="diff-old">${esc(rowA[c]??'')}</span><span class="diff-arr">→</span><span class="diff-new">${esc(rowB[c]??'')}</span></td>`;
        } else {
          let cls=isL?'cl':isR?'cr-pin':'';
          html+=`<td class="${cls}"${sticky}>${esc(rowB[c]??'')}</td>`;
        }
      });
      html+='</tr>';
    } else {
      // same
      html+=`<tr class="${rowCls}" data-ri="${ri}"><td class="rn" data-ri="${ri}" style="left:0">${ri}<span class="resize-v" data-rv="${ri}"></span></td>`;
      cols.forEach((c,ci)=>{
        const isL=lockedLeft.has(ci),isR=lockedRight.has(ci);
        const cls=isL?'cl':isR?'cr-pin':'';
        const sticky=isL?` style="position:sticky;left:${leftOffMap[ci]}px"`:isR?` style="position:sticky;right:${rightOffMap[ci]}px"`:'';
        html+=`<td class="${cls}"${sticky}>${esc(rowB[c]??'')}</td>`;
      });
      html+='</tr>';
    }
  });
  html+='</tbody></table>';
  return html;
}

function attachGridEvents(panel){
  const table=panel.querySelector('table');
  if(!table)return;

  table.querySelectorAll('th[data-ci]').forEach(th=>{
    th.addEventListener('click',e=>{
      if(e.target.classList.contains('resize-h'))return;
      const ci=+th.dataset.ci;
      if(e.shiftKey){lockedRight.has(ci)?lockedRight.delete(ci):lockedRight.add(ci);lockedLeft.delete(ci);}
      else{lockedLeft.has(ci)?lockedLeft.delete(ci):lockedLeft.add(ci);lockedRight.delete(ci);}
      refreshGrid();
    });
  });

  table.querySelectorAll('td.rn[data-ri]').forEach(td=>{
    td.addEventListener('click',e=>{
      if(e.target.classList.contains('resize-v'))return;
      const ri=+td.dataset.ri;
      lockedRows.has(ri)?lockedRows.delete(ri):lockedRows.add(ri);
      refreshGrid();
    });
  });

  table.querySelectorAll('.resize-h').forEach(handle=>{
    handle.addEventListener('mousedown',e=>{
      e.preventDefault();e.stopPropagation();
      const ci=+handle.dataset.rh;
      const th=table.querySelector(`th[data-ci="${ci}"]`);
      const startX=e.clientX,startW=th.getBoundingClientRect().width;
      handle.classList.add('dragging');
      const onMove=mv=>{
        const w=Math.max(40,startW+(mv.clientX-startX));
        table.querySelectorAll(`tr>*:nth-child(${ci+2})`).forEach(el=>{el.style.minWidth=w+'px';el.style.maxWidth=w+'px';el.style.width=w+'px';});
      };
      const onUp=()=>{handle.classList.remove('dragging');document.removeEventListener('mousemove',onMove);document.removeEventListener('mouseup',onUp);};
      document.addEventListener('mousemove',onMove);document.addEventListener('mouseup',onUp);
    });
  });

  table.querySelectorAll('.resize-v').forEach(handle=>{
    handle.addEventListener('mousedown',e=>{
      e.preventDefault();e.stopPropagation();
      const ri=handle.dataset.rv;
      const tr=table.querySelector(`tr[data-ri="${ri}"]`);
      const startY=e.clientY,startH=tr.offsetHeight||31;
      handle.classList.add('dragging');
      const onMove=e=>{const h=Math.max(20,startH+(e.clientY-startY));tr.style.height=h+'px';tr.querySelectorAll('td').forEach(td=>td.style.height=h+'px');};
      const onUp=()=>{handle.classList.remove('dragging');document.removeEventListener('mousemove',onMove);document.removeEventListener('mouseup',onUp);};
      document.addEventListener('mousemove',onMove);document.addEventListener('mouseup',onUp);
    });
  });
}

// ── 3. ALL CHANGES ──
function buildChanges(sheets){
  const sel=document.getElementById('filterSheet');
  sel.innerHTML='<option value="">All sheets</option>';
  sheets.forEach(n=>sel.innerHTML+=`<option value="${esc(n)}">${esc(n)}</option>`);
  renderChanges();
  document.getElementById('filterText').oninput=renderChanges;
  document.getElementById('filterType').onchange=renderChanges;
  document.getElementById('filterSheet').onchange=renderChanges;
}

function renderChanges(){
  const txt=document.getElementById('filterText').value.toLowerCase();
  const type=document.getElementById('filterType').value;
  const sheet=document.getElementById('filterSheet').value;
  const filtered=allChanges.filter(c=>{
    if(type&&c.type!==type)return false;
    if(sheet&&c.sheet!==sheet)return false;
    if(txt&&![c.sheet,c.header,c.va,c.vb,c.key].some(v=>String(v).toLowerCase().includes(txt)))return false;
    return true;
  });
  document.getElementById('chgCount').textContent=`${filtered.length} change${filtered.length!==1?'s':''}`;
  const list=document.getElementById('chgList');
  list.innerHTML=filtered.map(c=>{
    const typeBadge=c.type==='added'?'<span class="badge b-add">added</span>':c.type==='deleted'?'<span class="badge b-del">deleted</span>':'<span class="badge b-mod">modified</span>';
    const oldVal=c.type==='added'?'<span style="color:#555">—</span>':`<span class="val-old">${esc(c.va)}</span>`;
    const newVal=c.type==='deleted'?'<span style="color:#555">—</span>':`<span class="val-new">${esc(c.vb)}</span>`;
    return`<div class="chg-item"><span>${esc(c.sheet)}</span><span style="color:#569cd6">${esc(c.header)}</span>${typeBadge}${oldVal}${newVal}</div>`;
  }).join('');
}

function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}
