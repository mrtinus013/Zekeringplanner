const POI_SOURCES = [
  ['https://raw.githubusercontent.com/POIenexis/POI-zoeker/main/poi_e_station.csv','Station'],
  ['https://raw.githubusercontent.com/POIenexis/POI-zoeker/main/poi_e_verdeelkast.csv','Verdeelkast']
];
const LS = { rows:'zp_rows_v2', meta:'zp_meta_v2', settings:'zp_settings_v2', mapping:'zp_mapping_v2' };
const $ = s => document.querySelector(s);
const state = { poi:[], poiByCode:new Map(), rows:[], groups:[], start:null, workbook:null, currentSheet:null, mapping:null, planned:[], nearbyActions:[], pendingFilename:'' };

function toast(t){const x=$('#toast');x.textContent=t;x.classList.add('show');clearTimeout(x._t);x._t=setTimeout(()=>x.classList.remove('show'),1900)}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function norm(s){return String(s??'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'')}
function csvLine(s){let a=[],v='',z=false;for(let i=0;i<s.length;i++){const c=s[i];if(c==='"'){if(z&&s[i+1]==='"'){v+='"';i++}else z=!z}else if(c===';'&&!z){a.push(v);v=''}else v+=c}a.push(v);return a}
function coord(v,k){v=String(v??'').trim().replace(/\s/g,'').replace(',','.');let n=Number(v),ok=x=>k==='lat'?x>=50&&x<=54:x>=3&&x<=8;if(ok(n))return n;let d=v.replace(/[^\d-]/g,''),neg=d[0]==='-',x=d.replace('-','');for(let p=1;p<=2;p++){n=Number((neg?'-':'')+x.slice(0,p)+'.'+x.slice(p));if(ok(n))return n}return NaN}
function hav(a,b,c,d){const R=6371,x=(c-a)*Math.PI/180,y=(d-b)*Math.PI/180,z=Math.sin(x/2)**2+Math.cos(a*Math.PI/180)*Math.cos(c*Math.PI/180)*Math.sin(y/2)**2;return 2*R*Math.asin(Math.sqrt(z))}
function val(cell){const v=cell?.value;if(v==null)return '';if(typeof v==='object'){if('text' in v)return String(v.text??'');if('result' in v)return String(v.result??'');if(Array.isArray(v.richText))return v.richText.map(x=>x.text).join('');if('formula' in v&&cell.result!=null)return String(cell.result)}return String(v)}
function fuseLabel(v){let s=String(v??'').trim();if(!s)return '';const m=s.replace(',','.').match(/^(?:.*?)(\d+(?:\.\d+)?)(?:\s*[aA])?$/);if(!m)return s;const n=Number(m[1]);return Number.isFinite(n)?`${Number.isInteger(n)?n:n.toFixed(1)}A`:s}
function fuseSort([a],[b]){const na=parseFloat(a),nb=parseFloat(b),fa=Number.isFinite(na),fb=Number.isFinite(nb);if(fa&&fb)return na-nb;if(fa)return -1;if(fb)return 1;return String(a).localeCompare(String(b),'nl')}
function getSettings(){return Object.assign({defaultQty:3,theme:'dark'},JSON.parse(localStorage.getItem(LS.settings)||'{}'))}
function saveSettings(patch){const s=Object.assign(getSettings(),patch);localStorage.setItem(LS.settings,JSON.stringify(s));return s}

async function loadPoi(){
  if(!state.rows.length)$('#fileInfo').textContent='POI-locaties laden…';
  try{
    const all=[];let failures=0;
    await Promise.all(POI_SOURCES.map(async([url,type])=>{
      try{
        const r=await fetch(url,{cache:'no-cache'});if(!r.ok)throw new Error(type);
        const lines=(await r.text()).split(/\r?\n/).filter(Boolean),h=csvLine(lines.shift());
        const ni=h.findIndex(x=>/naam/i.test(x)),ai=h.findIndex(x=>/^(latitude|lat)$/i.test(x.trim())),oi=h.findIndex(x=>/^(longitude|lon)$/i.test(x.trim()));
        for(const line of lines){const c=csvLine(line),lat=coord(c[ai],'lat'),lon=coord(c[oi],'lon');if(!isFinite(lat)||!isFinite(lon))continue;const raw=c[ni]||'',p=raw.split(',').map(x=>x.trim());all.push({type,raw,code:p[0]||'',name:p[1]||'',street:p[2]||'',house:p[3]||'',zip:p[4]||'',city:p[5]||'',lat,lon,search:norm(raw+' '+type)})}
      }catch(e){failures++;console.error(e)}
    }));
    const unique=new Map();for(const i of all){const k=norm(i.code);if(!k)continue;if(!unique.has(k)||i.type==='Station')unique.set(k,i)}
    state.poi=[...unique.values()];state.poiByCode=unique;rebuildGroups();updateSummary();if(failures)toast('Een POI-bron kon niet laden');
  }catch(e){toast('POI-data kon niet worden geladen');console.error(e)}
}

function fillStatus(cell){
  const f=cell?.fill;if(!f||f.type!=='pattern')return null;
  const c=f.fgColor||f.bgColor||{};
  // Dit bestand gebruikt Accent 6 (theme 9) voor groen; ook de donkere tint daarvan is gereed.
  if(Number(c.theme)===9)return 'done';
  // Accent 4 wordt in veel Excel-thema's als geel gebruikt.
  if(Number(c.theme)===7)return 'action';
  let argb=c.argb;if(!argb)return null;argb=String(argb).replace('#','').toUpperCase();if(argb.length===8)argb=argb.slice(2);if(argb.length!==6)return null;
  if(['92D050','00B050','70AD47','548235'].includes(argb))return 'done';
  if(['FFFF00','FFD966','FFC000'].includes(argb))return 'action';
  const r=parseInt(argb.slice(0,2),16)/255,g=parseInt(argb.slice(2,4),16)/255,b=parseInt(argb.slice(4,6),16)/255,max=Math.max(r,g,b),min=Math.min(r,g,b),d=max-min;let h=0;
  if(d){if(max===r)h=((g-b)/d)%6;else if(max===g)h=(b-r)/d+2;else h=(r-g)/d+4;h*=60;if(h<0)h+=360}
  const s=max===0?0:d/max,v=max;if(s<.20||v<.35)return null;if(h>=75&&h<=165)return 'done';if(h>=38&&h<=70)return 'action';return null
}
function rowStatus(row,statusText='',primaryCell=null){
  // Kleur is leidend; in de aangeleverde werklijst is G groen/geel/ongekleurd.
  const primary=fillStatus(primaryCell);if(primary)return primary;
  let action=false;for(let i=1;i<=Math.min(row.cellCount,40);i++){const s=fillStatus(row.getCell(i));if(s==='done')return 'done';if(s==='action')action=true}if(action)return 'action';
  const t=norm(statusText);if(/gereed|gedaan|klaar|done|voltooid/.test(t))return 'done';if(/actie|wacht|uitzoek|control|probleem|afwijk|plan nodig/.test(t))return 'action';return 'open'
}
function headerScore(row){let s=0;row.eachCell({includeEmpty:false},c=>{const t=norm(val(c));if(t==='tewisselenstation')s+=30;else if(t.includes('station'))s+=7;if(t.includes('richting')||t.includes('veld'))s+=5;if(t.includes('zekering')||t.includes('smeltveiligheid')||t.includes('ampere'))s+=6;if(t.includes('status')||t.includes('opmerking')||t.includes('actie'))s+=2});return s}
function worksheetScore(ws){let best=0;for(let r=1;r<=Math.min(ws.rowCount,40);r++)best=Math.max(best,headerScore(ws.getRow(r)));if(norm(ws.name)==='data')best+=20;return best}
function detectHeader(ws){let best={row:1,score:-1};for(let r=1;r<=Math.min(ws.rowCount,50);r++){const s=headerScore(ws.getRow(r));if(s>best.score)best={row:r,score:s}}return best.row}
function headersFor(ws,rowNum){const row=ws.getRow(rowNum),a=[];for(let i=1;i<=Math.max(row.cellCount,1);i++){const text=val(row.getCell(i)).trim();if(text)a.push({col:i,text})}return a}
function detectMapping(ws,headerRow){
  const hs=headersFor(ws,headerRow),findExact=names=>{for(const h of hs)if(names.includes(norm(h.text)))return h.col;return 0},pick=(tests,exclude=[])=>{for(const h of hs){const n=norm(h.text);if(exclude.some(x=>n.includes(x)))continue;if(tests.some(x=>x.test(n)))return h.col}return 0};
  // Specifiek voor GFF_zekeringwissel: G/H/I/J zijn de operationele kolommen.
  let station=findExact(['tewisselenstation']);
  if(!station)station=findExact(['stationnummer','stationsnummer','station','objectnummer','objectnr']);
  if(!station)station=pick([/station/,/locatiecode/,/^code$/,/^nummer$/],[/richting/]);
  const direction=findExact(['lsveldnummer','richtingnummer','richting','richtingnr'])||pick([/richtingnummer/,/^richting$/, /richtingnr/,/^veld$/, /veldnummer/,/groepnummer/]);
  const fuse=findExact(['lssmeltveiligheidinom','zekeringwaarde','zekering'])||pick([/zekeringwaarde/,/zekering/,/smeltveiligheid/,/ampere/,/nominaal/,/^waarde$/]);
  const status=findExact(['opmerking','status'])||pick([/^status$/, /opmerking/,/actie/,/gereed/]);
  const qty=findExact(['aantal','stuks','hoeveelheid','qty'])||pick([/^aantal$/, /stuks/,/hoeveel/,/^qty$/]);
  return {sheet:ws.name,headerRow,stationCol:station,directionCol:direction,fuseCol:fuse,qtyCol:qty,statusCol:status};
}
function colText(ws,row,col){return col?val(ws.getRow(row).getCell(col)).trim():''}
function parseWorksheet(ws,m){
  const rows=[];let lastStation='';
  for(let r=m.headerRow+1;r<=ws.rowCount;r++){
    const row=ws.getRow(r);let station=colText(ws,r,m.stationCol);const direction=colText(ws,r,m.directionCol),fuse=fuseLabel(colText(ws,r,m.fuseCol)),qtyTxt=colText(ws,r,m.qtyCol),statusTxt=colText(ws,r,m.statusCol);
    if(station)lastStation=station;else if((direction||fuse)&&lastStation)station=lastStation;
    if(!station||(!direction&&!fuse))continue;
    const qtyNum=Number(String(qtyTxt).replace(',','.'));const status=rowStatus(row,statusTxt,row.getCell(m.stationCol));
    rows.push({id:`${r}-${norm(station)}-${norm(direction)}-${norm(fuse)}`,row:r,station:station.trim(),direction:direction||'—',fuse,status,qty:Number.isFinite(qtyNum)&&qtyNum>0?qtyNum:null,statusText:statusTxt});
  }
  return rows;
}

async function importFile(file){
  if(typeof ExcelJS==='undefined'){toast('Excel-lezer kon niet laden');return}
  try{
    $('#fileInfo').textContent='Excel wordt gelezen…';const wb=new ExcelJS.Workbook();await wb.xlsx.load(await file.arrayBuffer());state.workbook=wb;state.pendingFilename=file.name;
    const sheets=wb.worksheets.slice().sort((a,b)=>worksheetScore(b)-worksheetScore(a)),ws=sheets[0];if(!ws)throw new Error('Geen werkblad gevonden');
    const headerRow=detectHeader(ws),m=detectMapping(ws,headerRow);state.currentSheet=ws;state.mapping=m;
    if(!m.stationCol||!m.fuseCol){openMapping();toast('Controleer de Excel-kolommen');return}
    applyParsedRows(ws,m,file.name);
  }catch(e){console.error(e);$('#fileInfo').textContent='Excel kon niet worden gelezen';toast('Excel kon niet worden gelezen')}
}
function applyParsedRows(ws,m,filename){
  const rows=parseWorksheet(ws,m);state.rows=rows;state.mapping=m;localStorage.setItem(LS.rows,JSON.stringify(rows));localStorage.setItem(LS.mapping,JSON.stringify(m));localStorage.setItem(LS.meta,JSON.stringify({filename:filename||JSON.parse(localStorage.getItem(LS.meta)||'{}').filename||'Excel',loadedAt:new Date().toISOString(),sheet:ws.name}));
  rebuildGroups();updateSummary();if($('#mappingDlg').open)$('#mappingDlg').close();const open=rows.filter(r=>r.status==='open').length,action=rows.filter(r=>r.status==='action').length,done=rows.filter(r=>r.status==='done').length;toast(`${open} open · ${action} actie · ${done} gereed`)
}
function extractCode(s){const t=String(s||'').trim();const m=t.match(/\b\d{1,4}\.(?:(?:SK|VK)\s*)?\d+\b/i)||t.match(/\b\d{1,4}[.\-\s]\d{2,5}\b/);return m?m[0].replace(/\s/g,''):t.split(/[;,]/)[0].trim()}
function matchPoi(station){
  const code=extractCode(station),k=norm(code);if(state.poiByCode.has(k))return state.poiByCode.get(k);
  const candidates=state.poi.filter(p=>norm(p.code).endsWith(k)||k.endsWith(norm(p.code)));if(candidates.length===1)return candidates[0];
  const digits=String(code).replace(/\D/g,'');if(digits){const same=state.poi.filter(p=>String(p.code).replace(/\D/g,'')===digits);if(same.length===1)return same[0]}return null
}
function rebuildGroups(){
  const map=new Map();for(const row of state.rows){const key=norm(extractCode(row.station));if(!key)continue;if(!map.has(key))map.set(key,{key,station:row.station,rows:[]});map.get(key).rows.push(row)}
  state.groups=[...map.values()].map(g=>{g.poi=matchPoi(g.station);g.open=g.rows.filter(r=>r.status==='open').length;g.action=g.rows.filter(r=>r.status==='action').length;g.done=g.rows.filter(r=>r.status==='done').length;return g});renderUnmatched();
}
function updateSummary(){
  const meta=JSON.parse(localStorage.getItem(LS.meta)||'{}');if(state.rows.length){$('#fileInfo').textContent=`${meta.filename||'Werklijst'} · ${meta.sheet||state.mapping?.sheet||'werkblad'} · lokaal geladen`;$('#reimportBtn').classList.remove('hidden')}else{$('#fileInfo').textContent=state.poi.length?'Nog geen werklijst geladen':'POI-locaties laden…';$('#reimportBtn').classList.add('hidden')}
  const chips=$('#summaryChips');if(!state.rows.length){chips.classList.add('hidden');return}const open=state.rows.filter(r=>r.status==='open').length,action=state.rows.filter(r=>r.status==='action').length,done=state.rows.filter(r=>r.status==='done').length,work=state.groups.filter(g=>g.open||g.action).length,unmatched=state.groups.filter(g=>(g.open||g.action)&&!g.poi).length;chips.innerHTML=`<span class="chip open">${open} open</span><span class="chip action">${action} actie nodig</span><span class="chip done">${done} gereed</span><span class="chip">${work} locaties over</span>${unmatched?`<span class="chip bad">${unmatched} niet gekoppeld</span>`:''}`;chips.classList.remove('hidden');
}
function renderUnmatched(){const u=state.groups.filter(g=>(g.open||g.action)&&!g.poi);$('#unmatchedCount').textContent=u.length?`(${u.length})`:'';$('#unmatchedList').innerHTML=u.slice(0,80).map(g=>`<div>${esc(g.station)} · ${g.open} open · ${g.action} actie</div>`).join('');$('#unmatchedSection').classList.toggle('hidden',!u.length)}

function searchPoi(term){const q=norm(term);if(q.length<1)return [];return state.poi.map(p=>{let score=-1;const c=norm(p.code),n=norm(p.name),city=norm(p.city);if(c===q)score=100;if(c.startsWith(q))score=Math.max(score,80);if(c.includes(q))score=Math.max(score,65);if(n.startsWith(q))score=Math.max(score,55);if(n.includes(q)||city.includes(q)||p.search.includes(q))score=Math.max(score,35);return{p,score}}).filter(x=>x.score>=0).sort((a,b)=>b.score-a.score).slice(0,14).map(x=>x.p)}
function renderSuggestions(){const term=$('#stationSearch').value.trim(),box=$('#suggestions');if(!term){box.classList.add('hidden');box.innerHTML='';return}const a=searchPoi(term);box.innerHTML=a.map((p,i)=>`<button class="suggestion" data-i="${i}"><div><code>${esc(p.code)}</code><b>${esc(p.name||p.type)}</b><small>${esc(p.type)} · ${esc([p.street,p.house,p.city].filter(Boolean).join(' '))}</small></div><span>›</span></button>`).join('')||'<div class="suggestion"><small>Geen locatie gevonden</small></div>';box.classList.remove('hidden');box.querySelectorAll('button').forEach((b,i)=>b.onclick=()=>selectStart(a[i]))}
function selectStart(p,label){state.start={lat:p.lat,lon:p.lon,poi:p,label:label||`${p.code} ${p.name}`};$('#selectedStart').classList.remove('empty');$('#selectedStart').innerHTML=`<b>${esc(label||p.code+' '+p.name)}</b><small>${esc(p.type)} · ${esc([p.street,p.house,p.zip,p.city].filter(Boolean).join(' '))}</small>`;$('#suggestions').classList.add('hidden');$('#stationSearch').value='';$('#planBtn').disabled=!state.rows.length}
function useGeo(){if(!navigator.geolocation)return toast('Locatie niet beschikbaar');navigator.geolocation.getCurrentPosition(pos=>{state.start={lat:pos.coords.latitude,lon:pos.coords.longitude,poi:null,label:'Mijn locatie'};$('#selectedStart').classList.remove('empty');$('#selectedStart').innerHTML='<b>Mijn locatie</b><small>Startpunt op basis van GPS</small>';$('#planBtn').disabled=!state.rows.length;toast('Locatie actief')},()=>toast('Geef locatie-toegang in de browser'),{enableHighAccuracy:true,timeout:12000,maximumAge:60000})}

function eligibleRows(g,includeAction){return g.rows.filter(r=>r.status==='open'||(includeAction&&r.status==='action'))}
function plan(){
  if(!state.start)return toast('Kies eerst een startpunt');if(!state.rows.length)return toast('Laad eerst de Excel');const radius=Number($('#radius').value),max=Number($('#maxStations').value),includeAction=$('#includeAction').checked;
  const nearby=state.groups.filter(g=>g.poi&&(g.open||g.action)).map(g=>Object.assign({},g,{dist:hav(state.start.lat,state.start.lon,g.poi.lat,g.poi.lon)})).filter(g=>g.dist<=radius).sort((a,b)=>a.dist-b.dist);
  const work=nearby.filter(g=>eligibleRows(g,includeAction).length).slice(0,max),actions=nearby.filter(g=>g.action>0).slice(0,60);state.planned=work;state.nearbyActions=actions;renderPlan(work,actions,includeAction);if(!work.length)toast(includeAction?'Geen resterende locaties binnen deze straal':'Geen open wissels binnen deze straal')
}
function navUrl(p){return `https://www.google.com/maps/dir/?api=1&destination=${p.lat},${p.lon}`}
function distanceLabel(d){return d<1?Math.round(d*1000)+' m':d.toFixed(1)+' km'}
function actionLines(g){return g.rows.filter(r=>r.status==='action').map(r=>`<div class="action-line"><b>R${esc(r.direction)} · ${esc(r.fuse||'geen waarde')}</b>${r.statusText?`<span>${esc(r.statusText)}</span>`:''}</div>`).join('')}
function renderPlan(groups,actions,includeAction){
  $('#resultsSection').classList.remove('hidden');$('#shoppingSection').classList.remove('hidden');$('#resultCount').textContent=`${groups.length} locatie${groups.length===1?'':'s'}`;
  $('#stationResults').innerHTML=groups.map(g=>{const active=eligibleRows(g,includeAction),warnings=g.rows.filter(r=>r.status==='action');return `<article class="station-card"><div class="station-top"><div><div class="station-code">${esc(g.poi.code||g.station)}</div><div class="station-name">${esc(g.poi.type)} · ${esc(g.poi.name||'')} · ${esc(g.poi.city||'')}</div></div><div class="distance">${distanceLabel(g.dist)}</div></div><div class="direction-list">${active.map(r=>`<span class="fuse-tag ${r.status==='action'?'action':''}">R${esc(r.direction)} · ${esc(r.fuse||'geen waarde')}</span>`).join('')}</div>${!includeAction&&warnings.length?`<div class="warnline">⚠ ${warnings.length} gele regel${warnings.length===1?'':'s'} — zie Actie nodig hieronder</div>`:''}<div class="station-actions"><a class="navlink" target="_blank" rel="noopener" href="${navUrl(g.poi)}">↗ Navigeer</a></div></article>`}).join('')||'<div class="muted">Geen open wissels gevonden.</div>';

  if(!includeAction&&actions.length){$('#actionsSection').classList.remove('hidden');$('#actionCount').textContent=`${actions.length} locatie${actions.length===1?'':'s'}`;$('#actionResults').innerHTML=actions.map(g=>`<article class="station-card action-card"><div class="station-top"><div><div class="station-code">${esc(g.poi.code||g.station)}</div><div class="station-name">${esc(g.poi.type)} · ${esc(g.poi.name||'')} · ${esc(g.poi.city||'')}</div></div><div class="distance action-distance">${distanceLabel(g.dist)}</div></div><div class="action-lines">${actionLines(g)}</div><div class="station-actions"><a class="navlink" target="_blank" rel="noopener" href="${navUrl(g.poi)}">↗ Navigeer</a></div></article>`).join('')}else{$('#actionsSection').classList.add('hidden');$('#actionResults').innerHTML=''}

  const counts=new Map(),settings=getSettings();let total=0,directions=0,unknown=0;
  for(const g of groups)for(const r of eligibleRows(g,includeAction)){if(!r.fuse)continue;const qty=r.qty||settings.defaultQty||3;counts.set(r.fuse,(counts.get(r.fuse)||0)+qty);total+=qty;directions++;if(!/\d/.test(r.fuse))unknown+=qty}
  const sorted=[...counts.entries()].sort(fuseSort);$('#shoppingList').innerHTML=sorted.map(([f,q])=>`<div class="shopping-row"><b>${esc(f)}</b><strong>${q}×</strong></div>`).join('')+`<div class="shopping-total">Totaal: ${total} zekeringen · ${directions} richtingen${unknown?` · ${unknown} zonder ampèrewaarde`:''}</div>`;$('#shoppingNote').textContent=`Deze Excel heeft geen aantalkolom. Daarom rekent de app met ${settings.defaultQty||3} zekeringen per richting. Groene regels zijn altijd uitgesloten.${includeAction?' Gele regels zijn nu meegenomen.':' Gele regels staan apart bij Actie nodig.'}`;
}
function shoppingText(){const includeAction=$('#includeAction').checked,s=getSettings(),counts=new Map();for(const g of state.planned)for(const r of eligibleRows(g,includeAction)){if(!r.fuse)continue;const q=r.qty||s.defaultQty||3;counts.set(r.fuse,(counts.get(r.fuse)||0)+q)}let total=0;const lines=[...counts.entries()].sort(fuseSort).map(([f,q])=>{total+=q;return `${q}x ${f}`});return `Zekeringen meenemen\n${lines.join('\n')}\n\nTotaal: ${total} stuks`}

function fillSelect(sel,headers,selected,allowEmpty=true){sel.innerHTML=(allowEmpty?'<option value="0">— niet gebruiken —</option>':'')+headers.map(h=>`<option value="${h.col}" ${h.col===Number(selected)?'selected':''}>${h.col}: ${esc(h.text)}</option>`).join('')}
function openMapping(){if(!state.workbook){toast('Kies eerst opnieuw de Excel');return}const dlg=$('#mappingDlg'),sheetSel=$('#sheetSelect');sheetSel.innerHTML=state.workbook.worksheets.map(ws=>`<option value="${esc(ws.name)}" ${ws.name===state.mapping?.sheet?'selected':''}>${esc(ws.name)}</option>`).join('');const ws=state.workbook.getWorksheet(state.mapping?.sheet)||state.workbook.worksheets[0];prepareMappingForSheet(ws,state.mapping?.headerRow||detectHeader(ws),state.mapping);dlg.showModal()}
function prepareMappingForSheet(ws,headerRow,current){state.currentSheet=ws;$('#headerRow').value=headerRow;const auto=current&&current.sheet===ws.name?current:detectMapping(ws,headerRow),hs=headersFor(ws,headerRow);fillSelect($('#stationCol'),hs,auto.stationCol,false);fillSelect($('#directionCol'),hs,auto.directionCol);fillSelect($('#fuseCol'),hs,auto.fuseCol,false);fillSelect($('#qtyCol'),hs,auto.qtyCol);fillSelect($('#statusCol'),hs,auto.statusCol)}
function mappingFromForm(){return {sheet:$('#sheetSelect').value,headerRow:Number($('#headerRow').value),stationCol:Number($('#stationCol').value),directionCol:Number($('#directionCol').value),fuseCol:Number($('#fuseCol').value),qtyCol:Number($('#qtyCol').value),statusCol:Number($('#statusCol').value)}}

function applyTheme(v){const r=v==='system'?(matchMedia('(prefers-color-scheme: light)').matches?'light':'dark'):v;document.body.classList.toggle('light',r==='light');document.querySelector('meta[name="theme-color"]')?.setAttribute('content',r==='light'?'#edf3f5':'#07131b');$('#themeBtn').textContent=r==='light'?'☾':'☼'}
function loadLocal(){try{state.rows=JSON.parse(localStorage.getItem(LS.rows)||'[]');state.mapping=JSON.parse(localStorage.getItem(LS.mapping)||'null')}catch{state.rows=[]}rebuildGroups();updateSummary();const s=getSettings();$('#defaultQty').value=s.defaultQty;$('#themePref').value=s.theme;applyTheme(s.theme);$('#planBtn').disabled=!state.rows.length||!state.start}

$('#excelFile').onchange=e=>{const f=e.target.files?.[0];if(f)importFile(f)};
$('#reimportBtn').onclick=()=>$('#excelFile').click();
$('#stationSearch').oninput=renderSuggestions;$('#clearStation').onclick=()=>{$('#stationSearch').value='';renderSuggestions()};$('#geoBtn').onclick=useGeo;$('#planBtn').onclick=plan;
$('#radius').onchange=()=>state.planned.length&&plan();$('#maxStations').onchange=()=>state.planned.length&&plan();$('#includeAction').onchange=()=>state.start&&state.rows.length&&plan();
$('#copyShopping').onclick=async()=>{try{await navigator.clipboard.writeText(shoppingText());toast('Boodschappenlijst gekopieerd')}catch{toast('Kopiëren niet gelukt')}};
$('#settingsBtn').onclick=()=>{$('#settingsDlg').showModal()};document.querySelectorAll('[data-close]').forEach(b=>b.onclick=()=>document.getElementById(b.dataset.close).close());
$('#defaultQty').onchange=e=>{const n=Math.max(1,Math.min(12,Number(e.target.value)||3));e.target.value=n;saveSettings({defaultQty:n});if(state.planned.length)plan()};
$('#themePref').onchange=e=>{saveSettings({theme:e.target.value});applyTheme(e.target.value)};$('#themeBtn').onclick=()=>{const now=document.body.classList.contains('light')?'dark':'light';saveSettings({theme:now});$('#themePref').value=now;applyTheme(now)};
$('#clearDataBtn').onclick=()=>{if(!confirm('Lokale werklijst op dit apparaat wissen?'))return;[LS.rows,LS.meta,LS.mapping].forEach(k=>localStorage.removeItem(k));state.rows=[];state.groups=[];state.planned=[];state.nearbyActions=[];updateSummary();$('#resultsSection').classList.add('hidden');$('#actionsSection').classList.add('hidden');$('#shoppingSection').classList.add('hidden');$('#settingsDlg').close();toast('Lokale werklijst gewist')};
$('#mappingBtn').onclick=()=>{openMapping();$('#settingsDlg').close()};$('#sheetSelect').onchange=e=>{const ws=state.workbook.getWorksheet(e.target.value);prepareMappingForSheet(ws,detectHeader(ws),null)};$('#headerRow').onchange=e=>{const ws=state.workbook.getWorksheet($('#sheetSelect').value);prepareMappingForSheet(ws,Number(e.target.value),null)};
$('#applyMapping').onclick=()=>{const m=mappingFromForm(),ws=state.workbook.getWorksheet(m.sheet);if(!m.stationCol||!m.fuseCol)return toast('Station en zekeringwaarde zijn verplicht');applyParsedRows(ws,m,state.pendingFilename||JSON.parse(localStorage.getItem(LS.meta)||'{}').filename||'Excel')};

loadLocal();loadPoi();
