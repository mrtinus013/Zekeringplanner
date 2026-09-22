const POI_SOURCES = [
  ['https://raw.githubusercontent.com/POIenexis/POI-zoeker/main/poi_e_station.csv','Station'],
  ['https://raw.githubusercontent.com/POIenexis/POI-zoeker/main/poi_e_verdeelkast.csv','Verdeelkast']
];
const LS = {
  rows:'zp_rows_v3', meta:'zp_meta_v3', settings:'zp_settings_v3', mapping:'zp_mapping_v3',
  completed:'zp_completed_v3', route:'zp_route_v4'
};
const $ = s => document.querySelector(s);
const state = {
  poi:[], poiByCode:new Map(), rows:[], groups:[], start:null, workbook:null, currentSheet:null,
  mapping:null, nearbyCandidates:[], nearbyActions:[], nearbyEngineering:[], nearbyMaintenance:[],
  routeDraft:[], routeOrder:[], routeFinalized:false, routeSnapshot:null, pendingFilename:'', completed:{}, searchDone:false
};

function syncSearchButton(){
  const btn=$('#searchNearbyBtn');if(!btn)return;
  const hasRows=state.rows.length>0,hasPoi=state.poi.length>0,hasStart=!!state.start;
  btn.disabled=!(hasRows&&hasPoi&&hasStart);
  btn.title=!hasRows?'Laad eerst de Excel':!hasPoi?'POI-locaties worden nog geladen':!hasStart?'Kies eerst je locatie of een station/postcode':'';
}

function toast(t){const x=$('#toast');x.textContent=t;x.classList.add('show');clearTimeout(x._t);x._t=setTimeout(()=>x.classList.remove('show'),1900)}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function norm(s){return String(s??'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'')}
function csvLine(s){let a=[],v='',z=false;for(let i=0;i<s.length;i++){const c=s[i];if(c==='"'){if(z&&s[i+1]==='"'){v+='"';i++}else z=!z}else if(c===';'&&!z){a.push(v);v=''}else v+=c}a.push(v);return a}
function coord(v,k){v=String(v??'').trim().replace(/\s/g,'').replace(',','.');let n=Number(v),ok=x=>k==='lat'?x>=50&&x<=54:x>=3&&x<=8;if(ok(n))return n;let d=v.replace(/[^\d-]/g,''),neg=d[0]==='-',x=d.replace('-','');for(let p=1;p<=2;p++){n=Number((neg?'-':'')+x.slice(0,p)+'.'+x.slice(p));if(ok(n))return n}return NaN}
function hav(a,b,c,d){const R=6371,x=(c-a)*Math.PI/180,y=(d-b)*Math.PI/180,z=Math.sin(x/2)**2+Math.cos(a*Math.PI/180)*Math.cos(c*Math.PI/180)*Math.sin(y/2)**2;return 2*R*Math.asin(Math.sqrt(z))}
function val(cell){const v=cell?.value;if(v==null)return '';if(typeof v==='object'){if('text' in v)return String(v.text??'');if('result' in v)return String(v.result??'');if(Array.isArray(v.richText))return v.richText.map(x=>x.text).join('');if('formula' in v&&cell.result!=null)return String(cell.result)}return String(v)}
function fuseLabel(v){let s=String(v??'').trim();if(!s)return '';if(/^\d+(?:[.,]\d+)?$/.test(s))return `${s.replace(',','.')}A`;return s}
function fuseAmps(v){const m=String(v??'').replace(',','.').match(/\b(\d+(?:\.\d+)?)\s*A?\b/i);return m?Number(m[1]):null}
function fuseSort([a],[b]){const na=fuseAmps(a),nb=fuseAmps(b),fa=Number.isFinite(na),fb=Number.isFinite(nb);if(fa&&fb)return na-nb;if(fa)return -1;if(fb)return 1;return String(a).localeCompare(String(b),'nl')}
function getSettings(){return Object.assign({defaultQty:3,theme:'dark',kevinEmail:''},JSON.parse(localStorage.getItem(LS.settings)||'{}'))}
function saveSettings(patch){const s=Object.assign(getSettings(),patch);localStorage.setItem(LS.settings,JSON.stringify(s));return s}
function loadCompleted(){try{state.completed=JSON.parse(localStorage.getItem(LS.completed)||'{}')||{}}catch{state.completed={}}}
function saveCompleted(){localStorage.setItem(LS.completed,JSON.stringify(state.completed))}
function isLocalDone(r){return !!state.completed[r.id]}
function effectiveOpen(r){return r.status==='open'&&!isLocalDone(r)}
function routeRows(g){return g.rows.filter(effectiveOpen)}

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
    state.poi=[...unique.values()];state.poiByCode=unique;rebuildGroups();updateSummary();restoreRouteAfterGroups();syncSearchButton();if(failures)toast('Een POI-bron kon niet laden');
  }catch(e){toast('POI-data kon niet worden geladen');console.error(e);syncSearchButton()}
}

function fillStatus(cell){
  const f=cell?.fill;if(!f||f.type!=='pattern')return null;
  const c=f.fgColor||f.bgColor||{};
  if(Number(c.theme)===9)return 'done';
  if(Number(c.theme)===7)return 'engineering';
  let argb=c.argb;if(!argb)return null;argb=String(argb).replace('#','').toUpperCase();if(argb.length===8)argb=argb.slice(2);if(argb.length!==6)return null;
  if(['92D050','00B050','70AD47','548235'].includes(argb))return 'done';
  if(['FFFF00','FFD966','FFF2CC'].includes(argb))return 'action';
  if(['FFC000','F4B183','ED7D31','FFCC99'].includes(argb))return 'engineering';
  if(['00B0F0','5B9BD5','4472C4','9DC3E6'].includes(argb))return 'maintenance';
  const r=parseInt(argb.slice(0,2),16)/255,g=parseInt(argb.slice(2,4),16)/255,b=parseInt(argb.slice(4,6),16)/255,max=Math.max(r,g,b),min=Math.min(r,g,b),d=max-min;let h=0;
  if(d){if(max===r)h=((g-b)/d)%6;else if(max===g)h=(b-r)/d+2;else h=(r-g)/d+4;h*=60;if(h<0)h+=360}
  const s=max===0?0:d/max,v=max;if(s<.22||v<.35)return null;
  if(h>=75&&h<=165)return 'done';if(h>=48&&h<70)return 'action';if(h>=25&&h<48)return 'engineering';if(h>=185&&h<=225)return 'maintenance';return null
}
function rowStatus(row,employeeText='',statusText=''){
  if(String(employeeText||'').trim())return 'done';
  const found=[];for(let i=1;i<=Math.min(row.cellCount,40);i++){const s=fillStatus(row.getCell(i));if(s)found.push(s)}
  if(found.includes('done'))return 'done';if(found.includes('action'))return 'action';if(found.includes('engineering'))return 'engineering';if(found.includes('maintenance'))return 'maintenance';
  const t=norm(statusText);if(/gereed|gedaan|klaar|done|voltooid/.test(t))return 'done';if(/actie|wacht|uitzoek|control|probleem|afwijk/.test(t))return 'action';return 'open'
}
function headerScore(row){let s=0;row.eachCell({includeEmpty:false},c=>{const t=norm(val(c));if(t==='tewisselenstation'||t==='netstationlskast')s+=30;else if(t.includes('station'))s+=7;if(t.includes('richting')||t.includes('veld'))s+=6;if(t.includes('zekering')||t.includes('smeltveiligheid')||t==='huidigewaarde'||t.includes('ampere'))s+=7;if(t.includes('opmerking')||t.includes('medewerker'))s+=2});return s}
function worksheetScore(ws){let best=0;for(let r=1;r<=Math.min(ws.rowCount,40);r++)best=Math.max(best,headerScore(ws.getRow(r)));if(norm(ws.name)==='data')best+=30;return best}
function detectHeader(ws){let best={row:1,score:-1};for(let r=1;r<=Math.min(ws.rowCount,50);r++){const s=headerScore(ws.getRow(r));if(s>best.score)best={row:r,score:s}}return best.row}
function headersFor(ws,rowNum){const row=ws.getRow(rowNum),a=[];for(let i=1;i<=Math.max(row.cellCount,1);i++){const text=val(row.getCell(i)).trim();if(text)a.push({col:i,text})}return a}
function detectMapping(ws,headerRow){
  const hs=headersFor(ws,headerRow),findExact=names=>{for(const h of hs)if(names.includes(norm(h.text)))return h.col;return 0},pick=(tests,exclude=[])=>{for(const h of hs){const n=norm(h.text);if(exclude.some(x=>n.includes(x)))continue;if(tests.some(x=>x.test(n)))return h.col}return 0};
  let station=findExact(['netstationlskast','tewisselenstation']);if(!station)station=findExact(['stationnummer','stationsnummer','station','objectnummer','objectnr']);if(!station)station=pick([/station/,/locatiecode/,/^code$/,/^nummer$/],[/richting/]);
  const direction=findExact(['lsveldnummer','richtingnummer','richting','richtingnr'])||pick([/richtingnummer/,/^richting$/, /richtingnr/,/^veld$/, /veldnummer/,/groepnummer/]);
  const fuse=findExact(['huidigewaarde','lssmeltveiligheidinom','zekeringwaarde','zekering'])||pick([/zekeringwaarde/,/zekering/,/smeltveiligheid/,/ampere/,/nominaal/,/^waarde$/]);
  const qty=findExact(['aantal','stuks','hoeveelheid','qty'])||pick([/^aantal$/, /stuks/,/hoeveel/,/^qty$/]);
  const specialism=findExact(['opmeringspecialisme','opmerkingspecialisme'])||pick([/specialisme/]);
  const engineering=findExact(['opmerkingenvanuitengineering','opmerkingvanuitengineering'])||pick([/engineering/]);
  const employee=findExact(['zekeringvervangenmedewerker','medewerker'])||pick([/vervangenmedewerker/,/^medewerker$/]);
  const sourceStation=findExact(['netstation']),cableGroup=findExact(['kabelgroep']),city=findExact(['woonplaats']),municipality=findExact(['gemeente']),netType=findExact(['netwerktype']);
  const status=findExact(['opmerking','status'])||pick([/^status$/, /^opmerking$/,/actie/,/gereed/]);
  return {sheet:ws.name,headerRow,stationCol:station,directionCol:direction,fuseCol:fuse,qtyCol:qty,statusCol:status,specialismCol:specialism,engineeringCol:engineering,employeeCol:employee,sourceStationCol:sourceStation,cableGroupCol:cableGroup,cityCol:city,municipalityCol:municipality,netTypeCol:netType};
}
function colText(ws,row,col){return col?val(ws.getRow(row).getCell(col)).trim():''}
function makeRowId(o){return [o.station,o.direction,o.fuse,o.cableGroup,o.sourceStation].map(norm).join('|')}
function parseWorksheet(ws,m){
  const rows=[];let lastStation='';
  for(let r=m.headerRow+1;r<=ws.rowCount;r++){
    const row=ws.getRow(r);let station=colText(ws,r,m.stationCol);const direction=colText(ws,r,m.directionCol),fuse=fuseLabel(colText(ws,r,m.fuseCol)),qtyTxt=colText(ws,r,m.qtyCol);
    const specialism=colText(ws,r,m.specialismCol),engineering=colText(ws,r,m.engineeringCol),employee=colText(ws,r,m.employeeCol),legacyStatus=colText(ws,r,m.statusCol);
    if(station)lastStation=station;else if((direction||fuse)&&lastStation)station=lastStation;if(!station||(!direction&&!fuse))continue;
    const item={row:r,station:station.trim(),direction:direction||'—',fuse,qty:null,specialism,engineering,employee,statusText:specialism||legacyStatus,sourceStation:colText(ws,r,m.sourceStationCol),cableGroup:colText(ws,r,m.cableGroupCol),city:colText(ws,r,m.cityCol),municipality:colText(ws,r,m.municipalityCol),netType:colText(ws,r,m.netTypeCol)};
    const qtyNum=Number(String(qtyTxt).replace(',','.'));item.qty=Number.isFinite(qtyNum)&&qtyNum>0?qtyNum:null;item.status=rowStatus(row,employee,item.statusText);item.id=makeRowId(item)||`${r}-${norm(station)}-${norm(direction)}-${norm(fuse)}`;rows.push(item);
  }
  return rows;
}

async function importFile(file){
  if(typeof ExcelJS==='undefined'){toast('Excel-lezer kon niet laden');return}
  try{
    $('#fileInfo').textContent='Excel wordt gelezen…';const wb=new ExcelJS.Workbook();await wb.xlsx.load(await file.arrayBuffer());state.workbook=wb;state.pendingFilename=file.name;
    const sheets=wb.worksheets.slice().sort((a,b)=>worksheetScore(b)-worksheetScore(a)),ws=sheets[0];if(!ws)throw new Error('Geen werkblad gevonden');
    const headerRow=detectHeader(ws),m=detectMapping(ws,headerRow);state.currentSheet=ws;state.mapping=m;if(!m.stationCol||!m.fuseCol){openMapping();toast('Controleer de Excel-kolommen');return}applyParsedRows(ws,m,file.name);
  }catch(e){console.error(e);$('#fileInfo').textContent='Excel kon niet worden gelezen';toast('Excel kon niet worden gelezen')}
}
function applyParsedRows(ws,m,filename){
  const rows=parseWorksheet(ws,m);state.rows=rows;state.mapping=m;localStorage.setItem(LS.rows,JSON.stringify(rows));localStorage.setItem(LS.mapping,JSON.stringify(m));localStorage.setItem(LS.meta,JSON.stringify({filename:filename||JSON.parse(localStorage.getItem(LS.meta)||'{}').filename||'Excel',loadedAt:new Date().toISOString(),sheet:ws.name}));
  rebuildGroups();updateSummary();renderCompleted();restoreRouteAfterGroups();syncSearchButton();if($('#mappingDlg').open)$('#mappingDlg').close();const c=statusCounts();toast(`${c.open} open · ${c.action} geel · ${c.engineering} engineering · ${c.maintenance} Henri · ${c.done} gereed`)
}
function extractCode(s){const t=String(s||'').trim();const m=t.match(/\b\d{1,4}\.(?:(?:SK|VK)\s*)?\d+\b/i)||t.match(/\b\d{1,4}[.\-\s]\d{2,5}\b/);return m?m[0].replace(/\s/g,''):t.split(/[;,]/)[0].trim()}
function matchPoi(station){const code=extractCode(station),k=norm(code);if(state.poiByCode.has(k))return state.poiByCode.get(k);const candidates=state.poi.filter(p=>norm(p.code).endsWith(k)||k.endsWith(norm(p.code)));if(candidates.length===1)return candidates[0];const digits=String(code).replace(/\D/g,'');if(digits){const same=state.poi.filter(p=>String(p.code).replace(/\D/g,'')===digits);if(same.length===1)return same[0]}return null}
function rebuildGroups(){
  const map=new Map();for(const row of state.rows){const key=norm(extractCode(row.station));if(!key)continue;if(!map.has(key))map.set(key,{key,station:row.station,rows:[]});map.get(key).rows.push(row)}
  state.groups=[...map.values()].map(g=>{g.poi=matchPoi(g.station);g.open=g.rows.filter(effectiveOpen).length;g.action=g.rows.filter(r=>r.status==='action').length;g.engineering=g.rows.filter(r=>r.status==='engineering').length;g.maintenance=g.rows.filter(r=>r.status==='maintenance').length;g.done=g.rows.filter(r=>r.status==='done').length;g.localDone=g.rows.filter(isLocalDone).length;return g});renderUnmatched();
}
function statusCounts(){return {open:state.rows.filter(effectiveOpen).length,action:state.rows.filter(r=>r.status==='action').length,engineering:state.rows.filter(r=>r.status==='engineering').length,maintenance:state.rows.filter(r=>r.status==='maintenance').length,done:state.rows.filter(r=>r.status==='done').length,localDone:state.rows.filter(r=>isLocalDone(r)&&r.status==='open').length}}
function updateSummary(){
  const meta=JSON.parse(localStorage.getItem(LS.meta)||'{}');if(state.rows.length){$('#fileInfo').textContent=`${meta.filename||'Werklijst'} · ${meta.sheet||state.mapping?.sheet||'werkblad'} · lokaal geladen`;$('#reimportBtn').classList.remove('hidden')}else{$('#fileInfo').textContent=state.poi.length?'Nog geen werklijst geladen':'POI-locaties laden…';$('#reimportBtn').classList.add('hidden')}
  const chips=$('#summaryChips');if(!state.rows.length){chips.classList.add('hidden');return}const c=statusCounts(),work=state.groups.filter(g=>g.open||g.action||g.engineering||g.maintenance).length,unmatched=state.groups.filter(g=>(g.open||g.action||g.engineering||g.maintenance)&&!g.poi).length;chips.innerHTML=`<span class="chip open">${c.open} open</span><span class="chip action">${c.action} geel</span><span class="chip engineering">${c.engineering} engineering</span><span class="chip maintenance">${c.maintenance} Henri</span><span class="chip done">${c.done} gereed Excel</span>${c.localDone?`<span class="chip localdone">${c.localDone} afgevinkt</span>`:''}<span class="chip">${work} locaties actief</span>${unmatched?`<span class="chip bad">${unmatched} niet gekoppeld</span>`:''}`;chips.classList.remove('hidden');
}
function renderUnmatched(){const u=state.groups.filter(g=>(g.open||g.action||g.engineering||g.maintenance)&&!g.poi);$('#unmatchedCount').textContent=u.length?`(${u.length})`:'';$('#unmatchedList').innerHTML=u.slice(0,100).map(g=>`<div>${esc(g.station)} · ${g.open} open · ${g.action} geel · ${g.engineering} engineering · ${g.maintenance} Henri</div>`).join('');$('#unmatchedSection').classList.toggle('hidden',!u.length)}

function searchPoi(term){const q=norm(term);if(q.length<1)return [];return state.poi.map(p=>{let score=-1;const c=norm(p.code),n=norm(p.name),city=norm(p.city);if(c===q)score=100;if(c.startsWith(q))score=Math.max(score,80);if(c.includes(q))score=Math.max(score,65);if(n.startsWith(q))score=Math.max(score,55);if(n.includes(q)||city.includes(q)||p.search.includes(q))score=Math.max(score,35);return{p,score}}).filter(x=>x.score>=0).sort((a,b)=>b.score-a.score).slice(0,14).map(x=>x.p)}
function renderSuggestions(){const term=$('#stationSearch').value.trim(),box=$('#suggestions');if(!term){box.classList.add('hidden');box.innerHTML='';return}const a=searchPoi(term);box.innerHTML=a.map((p,i)=>`<button class="suggestion" data-i="${i}"><div><code>${esc(p.code)}</code><b>${esc(p.name||p.type)}</b><small>${esc(p.type)} · ${esc([p.street,p.house,p.city].filter(Boolean).join(' '))}</small></div><span>›</span></button>`).join('')||'<div class="suggestion"><small>Geen locatie gevonden</small></div>';box.classList.remove('hidden');box.querySelectorAll('button').forEach((b,i)=>b.onclick=()=>selectStart(a[i]))}
function selectStart(p,label){state.start={lat:p.lat,lon:p.lon,poi:p,label:label||`${p.code} ${p.name}`};$('#selectedStart').classList.remove('empty');$('#selectedStart').innerHTML=`<b>${esc(label||p.code+' '+p.name)}</b><small>${esc(p.type)} · ${esc([p.street,p.house,p.zip,p.city].filter(Boolean).join(' '))}</small>`;$('#suggestions').classList.add('hidden');$('#stationSearch').value='';syncSearchButton()}
function useGeo(){if(!navigator.geolocation)return toast('Locatie niet beschikbaar');navigator.geolocation.getCurrentPosition(pos=>{state.start={lat:pos.coords.latitude,lon:pos.coords.longitude,poi:null,label:'Mijn locatie'};$('#selectedStart').classList.remove('empty');$('#selectedStart').innerHTML='<b>Mijn locatie</b><small>Zoek- en startpunt op basis van GPS</small>';syncSearchButton();toast(state.poi.length?'Locatie actief':'Locatie actief · POI-data wordt nog geladen')},()=>toast('Geef locatie-toegang in je browser'),{enableHighAccuracy:true,timeout:12000,maximumAge:60000})}

function distanceLabel(d){return d<1?Math.round(d*1000)+' m':d.toFixed(1)+' km'}
function navUrl(p){return `https://www.google.com/maps/dir/?api=1&destination=${p.lat},${p.lon}`}
function rowMeta(r){return [r.cableGroup,r.sourceStation&&r.sourceStation!==r.station?`voeding ${r.sourceStation}`:''].filter(Boolean).join(' · ')}
function fuseSummary(g){const rows=routeRows(g),m=new Map();for(const r of rows){const k=r.fuse||'waarde onbekend';m.set(k,(m.get(k)||0)+1)}return [...m.entries()].sort(fuseSort).map(([f,n])=>`${n}× ${f}`).join(' · ')}
function isSelected(key){return state.routeDraft.includes(key)}

function searchNearby(){
  if(!state.start)return toast('Kies eerst een zoekpunt');if(!state.rows.length)return toast('Laad eerst de Excel');if(!state.poi.length)return toast('POI-locaties zijn nog niet geladen');
  const radius=Number($('#radius').value);
  const nearby=state.groups.filter(g=>g.poi&&(g.open||g.action||g.engineering||g.maintenance)).map(g=>Object.assign({},g,{dist:hav(state.start.lat,state.start.lon,g.poi.lat,g.poi.lon)})).filter(g=>g.dist<=radius).sort((a,b)=>a.dist-b.dist);
  state.nearbyCandidates=nearby.filter(g=>routeRows(g).length);state.nearbyActions=nearby.filter(g=>g.action>0).slice(0,100);state.nearbyEngineering=nearby.filter(g=>g.engineering>0).slice(0,100);state.nearbyMaintenance=nearby.filter(g=>g.maintenance>0).slice(0,100);state.searchDone=true;
  renderNearby();renderStatusSection('#actionsSection','#actionCount','#actionResults',state.nearbyActions,'action');renderStatusSection('#engineeringSection','#engineeringCount','#engineeringResults',state.nearbyEngineering,'engineering');renderStatusSection('#maintenanceSection','#maintenanceCount','#maintenanceResults',state.nearbyMaintenance,'maintenance');
  if(!state.nearbyCandidates.length)toast('Geen normale open wissels binnen deze straal')
}
function renderNearby(){
  const groups=state.nearbyCandidates;$('#resultsSection').classList.toggle('hidden',!state.searchDone);$('#routeBuilderSection').classList.toggle('hidden',!state.searchDone);$('#resultCount').textContent=`${groups.length} locatie${groups.length===1?'':'s'}`;
  $('#stationResults').innerHTML=groups.length?groups.map(g=>{const selected=isSelected(g.key);return `<article class="station-card candidate-card ${selected?'selected-candidate':''}"><div class="station-top"><div><div class="station-code">${esc(g.poi.code||g.station)}</div><div class="station-name">${esc(g.poi.type)} · ${esc(g.poi.name||'')} · ${esc(g.poi.city||'')}</div></div><div class="distance">${distanceLabel(g.dist)}</div></div><div class="candidate-fuses">${esc(fuseSummary(g))}</div><div class="candidate-meta">${routeRows(g).length} richting${routeRows(g).length===1?'':'en'} open</div><div class="station-actions"><button type="button" class="route-toggle ${selected?'remove':''}" data-group="${esc(g.key)}">${selected?'✓ In route — verwijder':'＋ Voeg toe aan route'}</button><a class="navlink" target="_blank" rel="noopener" href="${navUrl(g.poi)}">↗ Bekijk locatie</a></div></article>`}).join(''):'<div class="muted">Geen open wissels gevonden.</div>';
  $('#stationResults').querySelectorAll('.route-toggle').forEach(b=>b.addEventListener('click',()=>toggleRouteGroup(b.dataset.group)));renderRouteBuilder();
}
function toggleRouteGroup(key){if(isSelected(key))state.routeDraft=state.routeDraft.filter(x=>x!==key);else state.routeDraft.push(key);markRouteDirty();saveRouteState();renderNearby()}
function markRouteDirty(){state.routeFinalized=false;state.routeOrder=[];state.routeSnapshot=null;$('#routeSection').classList.add('hidden');$('#shoppingSection').classList.add('hidden')}
function clearRoute(){state.routeDraft=[];state.routeOrder=[];state.routeFinalized=false;state.routeSnapshot=null;saveRouteState();renderNearby();renderRouteBuilder();$('#routeSection').classList.add('hidden');$('#shoppingSection').classList.add('hidden');toast('Route gewist')}
function groupForKey(key){return state.groups.find(g=>g.key===key&&g.poi)}
function routeDraftGroups(){return state.routeDraft.map(groupForKey).filter(Boolean).filter(g=>routeRows(g).length)}
function routeOrderedGroups(){const keys=state.routeFinalized&&state.routeOrder.length?state.routeOrder:state.routeDraft;return keys.map(groupForKey).filter(Boolean)}
function renderRouteBuilder(){
  const groups=routeDraftGroups(),count=groups.length;$('#routeDraftCount').textContent=`${count} locatie${count===1?'':'s'} gekozen`;$('#routeDraftHint').textContent=count?'Klaar? Laat de app nu de rijvolgorde bepalen.':'Voeg stations uit de lijst hierboven toe.';$('#optimizeRouteBtn').disabled=!count||!state.start;
  $('#routeDraftChips').innerHTML=groups.map(g=>`<button type="button" class="draft-chip" data-group="${esc(g.key)}">${esc(g.poi?.code||g.station)} <span>×</span></button>`).join('');$('#routeDraftChips').querySelectorAll('.draft-chip').forEach(b=>b.addEventListener('click',()=>toggleRouteGroup(b.dataset.group)));
}

function pointDistance(a,b){return hav(a.lat,a.lon,b.poi.lat,b.poi.lon)}
function routeDistance(groups){if(!groups.length||!state.start)return 0;let total=hav(state.start.lat,state.start.lon,groups[0].poi.lat,groups[0].poi.lon);for(let i=1;i<groups.length;i++)total+=hav(groups[i-1].poi.lat,groups[i-1].poi.lon,groups[i].poi.lat,groups[i].poi.lon);return total}
function nearestNeighbor(groups){
  const remaining=[...groups],out=[];let cur={lat:state.start.lat,lon:state.start.lon};
  while(remaining.length){let best=0,bestD=Infinity;for(let i=0;i<remaining.length;i++){const d=hav(cur.lat,cur.lon,remaining[i].poi.lat,remaining[i].poi.lon);if(d<bestD){bestD=d;best=i}}const next=remaining.splice(best,1)[0];out.push(next);cur={lat:next.poi.lat,lon:next.poi.lon}}
  return out;
}
function twoOpt(groups){
  let best=[...groups],bestD=routeDistance(best),improved=true,passes=0;
  while(improved&&passes<8){improved=false;passes++;for(let i=0;i<best.length-1;i++){for(let k=i+1;k<best.length;k++){const candidate=[...best.slice(0,i),...best.slice(i,k+1).reverse(),...best.slice(k+1)],d=routeDistance(candidate);if(d+0.01<bestD){best=candidate;bestD=d;improved=true}}}}
  return best;
}
function createShoppingSnapshot(groups){
  const settings=getSettings(),counts={},unknown=[];let total=0,directions=0;
  for(const g of groups)for(const r of routeRows(g)){const qty=r.qty||settings.defaultQty||3;directions++;if(!isCountableFuse(r)){unknown.push({station:g.poi?.code||g.station,direction:r.direction,fuse:r.fuse||'onbekend'});continue}counts[r.fuse]=(counts[r.fuse]||0)+qty;total+=qty}
  return {counts,total,directions,unknown,createdAt:new Date().toISOString()};
}
function optimizeRoute(){
  if(!state.start)return toast('Kies eerst een startpunt');const groups=routeDraftGroups();if(!groups.length)return toast('Voeg eerst stations toe aan je route');
  const optimized=twoOpt(nearestNeighbor(groups));state.routeOrder=optimized.map(g=>g.key);state.routeDraft=[...state.routeOrder];state.routeFinalized=true;state.routeSnapshot=createShoppingSnapshot(optimized);saveRouteState();renderRoute();renderShoppingSnapshot();toast('Routevolgorde gemaakt')
}
function routeLegs(groups){let prev={lat:state.start.lat,lon:state.start.lon},cum=0;return groups.map((g,i)=>{const leg=hav(prev.lat,prev.lon,g.poi.lat,g.poi.lon);cum+=leg;prev={lat:g.poi.lat,lon:g.poi.lon};return {g,index:i,leg,cum}})}
function workLine(r){const done=isLocalDone(r);return `<div class="work-line ${done?'line-done':''}"><div><b>R${esc(r.direction)} · ${esc(r.fuse||'waarde onbekend')}</b>${rowMeta(r)?`<small>${esc(rowMeta(r))}</small>`:''}</div>${done?'<span class="done-label">✓ Gedaan</span>':`<button type="button" class="done-btn" data-rid="${esc(r.id)}">✓ Gedaan</button>`}</div>`}
function renderRoute(){
  if(!state.routeFinalized){$('#routeSection').classList.add('hidden');return}const groups=routeOrderedGroups();if(!groups.length){$('#routeSection').classList.add('hidden');return}
  $('#routeSection').classList.remove('hidden');const legs=routeLegs(groups),total=legs.at(-1)?.cum||0,remaining=groups.filter(g=>routeRows(g).length).length;$('#routeSummary').innerHTML=`<b>${groups.length} stops</b><span>≈ ${total.toFixed(1)} km op ligging</span><small>${remaining} locatie${remaining===1?'':'s'} nog open</small>`;
  $('#routeResults').innerHTML=legs.map(({g,index,leg,cum})=>{const active=routeRows(g),allOriginalOpen=g.rows.filter(r=>r.status==='open'),stationDone=!active.length&&allOriginalOpen.length>0;return `<article class="route-card ${stationDone?'route-card-done':''}" data-group="${esc(g.key)}"><div class="route-number">${index+1}</div><div class="route-card-body"><div class="station-top"><div><div class="station-code">${esc(g.poi.code||g.station)}</div><div class="station-name">${esc(g.poi.name||g.poi.type)} · ${esc(g.poi.city||'')}</div></div><div class="route-leg"><b>${distanceLabel(leg)}</b><small>vanaf vorige</small></div></div>${stationDone?'<div class="station-finished">✓ Deze locatie is afgevinkt</div>':`<div class="work-list">${active.map(workLine).join('')}</div>`}<div class="route-card-actions"><a class="navlink primary-nav" target="_blank" rel="noopener" href="${navUrl(g.poi)}">↗ Navigeer</a><button type="button" class="order-btn" data-dir="-1" data-group="${esc(g.key)}" ${index===0?'disabled':''}>↑</button><button type="button" class="order-btn" data-dir="1" data-group="${esc(g.key)}" ${index===groups.length-1?'disabled':''}>↓</button><button type="button" class="route-remove" data-group="${esc(g.key)}">Verwijder</button></div></div></article>`}).join('');
  $('#routeResults').querySelectorAll('.done-btn').forEach(b=>b.addEventListener('click',()=>completeRow(b.dataset.rid)));$('#routeResults').querySelectorAll('.order-btn').forEach(b=>b.addEventListener('click',()=>moveRoute(b.dataset.group,Number(b.dataset.dir))));$('#routeResults').querySelectorAll('.route-remove').forEach(b=>b.addEventListener('click',()=>removeFromFinalRoute(b.dataset.group)));updateNextButton(groups);
}
function moveRoute(key,dir){const i=state.routeOrder.indexOf(key),j=i+dir;if(i<0||j<0||j>=state.routeOrder.length)return;[state.routeOrder[i],state.routeOrder[j]]=[state.routeOrder[j],state.routeOrder[i]];state.routeDraft=[...state.routeOrder];saveRouteState();renderRoute()}
function removeFromFinalRoute(key){state.routeDraft=state.routeDraft.filter(x=>x!==key);markRouteDirty();saveRouteState();renderNearby();renderRouteBuilder();toast('Station verwijderd — maak de route opnieuw')}
function updateNextButton(groups){const next=groups.find(g=>routeRows(g).length);const b=$('#navigateNextBtn');if(!next){b.disabled=true;b.textContent='✓ Route klaar';b.onclick=null;return}b.disabled=false;b.textContent=`↗ Volgende: ${next.poi.code||next.station}`;b.onclick=()=>window.open(navUrl(next.poi),'_blank','noopener')}

function isCountableFuse(r){return Number.isFinite(fuseAmps(r.fuse))&&!/restbak|onbekend/i.test(r.fuse)}
function renderShoppingSnapshot(){
  const s=state.routeSnapshot,sec=$('#shoppingSection');if(!state.routeFinalized||!s){sec.classList.add('hidden');return}sec.classList.remove('hidden');const sorted=Object.entries(s.counts||{}).sort(fuseSort);$('#shoppingList').innerHTML=(sorted.length?sorted.map(([f,q])=>`<div class="shopping-row"><b>${esc(f)}</b><strong>${q}×</strong></div>`).join(''):'<div class="muted">Geen bekende zekeringwaarden te tellen.</div>')+`<div class="shopping-total">Totaal: ${s.total} zekeringen · ${s.directions} richtingen${s.unknown?.length?` · ${s.unknown.length} onbekend/restbak controleren`:''}</div>`;$('#shoppingNote').textContent=`Deze lijst is vastgelegd toen je de route maakte. Zo blijft zichtbaar wat je vóór vertrek mee moest nemen. Standaard ${getSettings().defaultQty||3} zekeringen per richting als Excel geen aantal bevat.`
}
function shoppingText(){const s=state.routeSnapshot;if(!s)return 'Nog geen route gemaakt.';const lines=Object.entries(s.counts||{}).sort(fuseSort).map(([f,q])=>`${q}x ${f}`);if(s.unknown?.length){lines.push('', 'Controleren:',...s.unknown.map(x=>`${x.station} R${x.direction}: ${x.fuse}`))}return `Zekeringen meenemen voor route\n${lines.join('\n')}\n\n${s.directions} richtingen · ${s.total} bekende zekeringen`}

function actionLines(g){return g.rows.filter(r=>r.status==='action').map(r=>`<div class="status-line yellow-line"><b>R${esc(r.direction)} · ${esc(r.fuse||'geen waarde')}</b><span>${esc(r.specialism||r.statusText||'Geen toelichting ingevuld')}</span><strong>✉ Mail Kevin</strong></div>`).join('')}
function engineeringLines(g){return g.rows.filter(r=>r.status==='engineering').map(r=>`<div class="status-line orange-line"><b>R${esc(r.direction)} · ${esc(r.fuse||'geen waarde')}</b><span>${esc(r.engineering||'Engineering-opmerking ontbreekt')}</span></div>`).join('')}
function maintenanceLines(g){return g.rows.filter(r=>r.status==='maintenance').map(r=>`<div class="status-line blue-line"><b>R${esc(r.direction)} · ${esc(r.fuse||'geen waarde')}</b><span>Wordt opgepakt door Henri van der Vleuten tijdens onderhoud.</span></div>`).join('')}
function mailKevinForGroup(g){const rows=g.rows.filter(r=>r.status==='action'),settings=getSettings(),subject=`Zekeringwissel ${g.poi?.code||g.station} - actie nodig`,body=[`Hoi Kevin,`,``,`Bij ${g.poi?.code||g.station} zijn de volgende zekeringwissels niet uitgevoerd:`,...rows.map(r=>`- richting ${r.direction}, ${r.fuse||'waarde onbekend'}: ${r.specialism||r.statusText||'geen toelichting'}`),``,`Groet`].join('\n');location.href=`mailto:${encodeURIComponent(settings.kevinEmail||'')}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`}
function renderStatusSection(sectionId,countId,listId,groups,kind){
  const sec=$(sectionId);if(!groups.length){sec.classList.add('hidden');$(listId).innerHTML='';return}sec.classList.remove('hidden');$(countId).textContent=`${groups.length} locatie${groups.length===1?'':'s'}`;
  $(listId).innerHTML=groups.map(g=>{const lines=kind==='action'?actionLines(g):kind==='engineering'?engineeringLines(g):maintenanceLines(g),cls=kind==='action'?'action-card':kind==='engineering'?'engineering-card':'maintenance-card';return `<article class="station-card ${cls}"><div class="station-top"><div><div class="station-code">${esc(g.poi.code||g.station)}</div><div class="station-name">${esc(g.poi.type)} · ${esc(g.poi.name||'')} · ${esc(g.poi.city||'')}</div></div><div class="distance ${kind}-distance">${distanceLabel(g.dist)}</div></div><div class="status-lines">${lines}</div><div class="station-actions"><a class="navlink" target="_blank" rel="noopener" href="${navUrl(g.poi)}">↗ Navigeer</a>${kind==='action'?`<button type="button" class="navlink mail-kevin" data-group="${esc(g.key)}">✉ Maak mail</button>`:''}</div></article>`}).join('');
  if(kind==='action')$(listId).querySelectorAll('.mail-kevin').forEach(b=>b.addEventListener('click',()=>{const g=groups.find(x=>x.key===b.dataset.group);if(g)mailKevinForGroup(g)}))
}

function completeRow(id){
  const r=state.rows.find(x=>x.id===id);if(!r||r.status!=='open'||isLocalDone(r))return;const settings=getSettings();state.completed[id]={id,station:r.station,direction:r.direction,fuse:r.fuse,qty:r.qty||settings.defaultQty||3,cableGroup:r.cableGroup||'',sourceStation:r.sourceStation||'',city:r.city||'',completedAt:new Date().toISOString()};saveCompleted();rebuildGroups();updateSummary();renderCompleted();refreshVisibleData();toast(`${r.station} R${r.direction} afgevinkt`)
}
function undoCompleted(id){delete state.completed[id];saveCompleted();rebuildGroups();updateSummary();renderCompleted();refreshVisibleData();toast('Afvinken hersteld')}
function refreshVisibleData(){if(state.searchDone&&state.start)searchNearby();if(state.routeFinalized){renderRoute();renderShoppingSnapshot()}renderRouteBuilder()}
function dateKey(d){const x=new Date(d);return `${x.getFullYear()}-${String(x.getMonth()+1).padStart(2,'0')}-${String(x.getDate()).padStart(2,'0')}`}
function completedItems(){const mode=$('#completedFilter')?.value||'today',all=Object.values(state.completed).sort((a,b)=>new Date(b.completedAt)-new Date(a.completedAt));if(mode==='all')return all;const today=dateKey(new Date());return all.filter(x=>dateKey(x.completedAt)===today)}
function completedText(){const items=completedItems(),counts=new Map(),stations=new Set();for(const x of items){stations.add(x.station);if(Number.isFinite(fuseAmps(x.fuse))&&!/restbak|onbekend/i.test(x.fuse))counts.set(x.fuse,(counts.get(x.fuse)||0)+(x.qty||0))}const lines=items.map(x=>`${x.station} - richting ${x.direction} - ${x.fuse||'waarde onbekend'} - ${new Date(x.completedAt).toLocaleTimeString('nl-NL',{hour:'2-digit',minute:'2-digit'})}`);const fuse=[...counts.entries()].sort(fuseSort).map(([f,q])=>`${q}x ${f}`);return `Uitgevoerd overzicht\n${lines.join('\n')}\n\n${items.length} richtingen · ${stations.size} locaties${fuse.length?`\nGebruikt:\n${fuse.join('\n')}`:''}`}
function renderCompleted(){
  const sec=$('#completedSection');if(!sec)return;const items=completedItems();$('#completedCount').textContent=`${items.length} richting${items.length===1?'':'en'}`;sec.classList.toggle('empty-completed',!items.length);const counts=new Map(),stations=new Set();for(const x of items){stations.add(x.station);if(Number.isFinite(fuseAmps(x.fuse))&&!/restbak|onbekend/i.test(x.fuse))counts.set(x.fuse,(counts.get(x.fuse)||0)+(x.qty||0))}
  $('#completedSummary').innerHTML=`<b>${stations.size} locatie${stations.size===1?'':'s'}</b><span>${items.length} richtingen uitgevoerd</span>${counts.size?`<small>${[...counts.entries()].sort(fuseSort).map(([f,q])=>`${q}× ${esc(f)}`).join(' · ')}</small>`:''}`;
  $('#completedList').innerHTML=items.length?items.map(x=>`<div class="completed-row"><div><b>${esc(x.station)} · R${esc(x.direction)}</b><span>${esc(x.fuse||'waarde onbekend')}${x.cableGroup?` · ${esc(x.cableGroup)}`:''}</span><small>${new Date(x.completedAt).toLocaleString('nl-NL',{dateStyle:'short',timeStyle:'short'})}</small></div><button type="button" class="undo-btn" data-rid="${esc(x.id)}">↶ Herstel</button></div>`).join(''):'<div class="muted">Nog niets afgevinkt in deze periode.</div>';$('#completedList').querySelectorAll('.undo-btn').forEach(b=>b.addEventListener('click',()=>undoCompleted(b.dataset.rid)))
}

function saveRouteState(){localStorage.setItem(LS.route,JSON.stringify({draft:state.routeDraft,order:state.routeOrder,finalized:state.routeFinalized,snapshot:state.routeSnapshot,start:state.start?{lat:state.start.lat,lon:state.start.lon,label:state.start.label,poi:state.start.poi}:null}))}
function loadRouteState(){try{const r=JSON.parse(localStorage.getItem(LS.route)||'null');if(!r)return;state.routeDraft=Array.isArray(r.draft)?r.draft:[];state.routeOrder=Array.isArray(r.order)?r.order:[];state.routeFinalized=!!r.finalized;state.routeSnapshot=r.snapshot||null;if(r.start&&Number.isFinite(r.start.lat)&&Number.isFinite(r.start.lon))state.start=r.start}catch{}}
function restoreRouteAfterGroups(){
  const valid=new Set(state.groups.filter(g=>g.poi).map(g=>g.key));state.routeDraft=state.routeDraft.filter(k=>valid.has(k));state.routeOrder=state.routeOrder.filter(k=>valid.has(k));if(state.routeFinalized&&!state.routeOrder.length)state.routeFinalized=false;
  if(state.start){$('#selectedStart').classList.remove('empty');$('#selectedStart').innerHTML=`<b>${esc(state.start.label||'Opgeslagen startpunt')}</b><small>Opgeslagen startpunt</small>`;syncSearchButton()}renderRouteBuilder();if(state.routeFinalized){renderRoute();renderShoppingSnapshot()}
}
function clearSavedRoute(){state.routeDraft=[];state.routeOrder=[];state.routeFinalized=false;state.routeSnapshot=null;localStorage.removeItem(LS.route);renderRouteBuilder();$('#routeSection').classList.add('hidden');$('#shoppingSection').classList.add('hidden');if(state.searchDone)renderNearby();toast('Opgeslagen route gewist')}

function fillSelect(sel,headers,selected,allowEmpty=true){sel.innerHTML=(allowEmpty?'<option value="0">— niet gebruiken —</option>':'')+headers.map(h=>`<option value="${h.col}" ${h.col===Number(selected)?'selected':''}>${h.col}: ${esc(h.text)}</option>`).join('')}
function openMapping(){if(!state.workbook){toast('Kies eerst opnieuw de Excel');return}const dlg=$('#mappingDlg'),sheetSel=$('#sheetSelect');sheetSel.innerHTML=state.workbook.worksheets.map(ws=>`<option value="${esc(ws.name)}" ${ws.name===state.mapping?.sheet?'selected':''}>${esc(ws.name)}</option>`).join('');const ws=state.workbook.getWorksheet(state.mapping?.sheet)||state.workbook.worksheets[0];prepareMappingForSheet(ws,state.mapping?.headerRow||detectHeader(ws),state.mapping);dlg.showModal()}
function prepareMappingForSheet(ws,headerRow,current){state.currentSheet=ws;$('#headerRow').value=headerRow;const auto=current&&current.sheet===ws.name?current:detectMapping(ws,headerRow),hs=headersFor(ws,headerRow);fillSelect($('#stationCol'),hs,auto.stationCol,false);fillSelect($('#directionCol'),hs,auto.directionCol);fillSelect($('#fuseCol'),hs,auto.fuseCol,false);fillSelect($('#qtyCol'),hs,auto.qtyCol)}
function mappingFromForm(){const base=detectMapping(state.workbook.getWorksheet($('#sheetSelect').value),Number($('#headerRow').value));return Object.assign(base,{sheet:$('#sheetSelect').value,headerRow:Number($('#headerRow').value),stationCol:Number($('#stationCol').value),directionCol:Number($('#directionCol').value),fuseCol:Number($('#fuseCol').value),qtyCol:Number($('#qtyCol').value)})}

function applyTheme(v){const r=v==='system'?(matchMedia('(prefers-color-scheme: light)').matches?'light':'dark'):v;document.body.classList.toggle('light',r==='light');document.querySelector('meta[name="theme-color"]')?.setAttribute('content',r==='light'?'#edf3f5':'#07131b');$('#themeBtn').textContent=r==='light'?'☾':'☼'}
function loadLocal(){try{state.rows=JSON.parse(localStorage.getItem(LS.rows)||'[]');state.mapping=JSON.parse(localStorage.getItem(LS.mapping)||'null')}catch{state.rows=[]}loadCompleted();loadRouteState();rebuildGroups();updateSummary();const s=getSettings();$('#defaultQty').value=s.defaultQty;$('#themePref').value=s.theme;$('#kevinEmail').value=s.kevinEmail||'';applyTheme(s.theme);syncSearchButton();renderCompleted();restoreRouteAfterGroups()}

$('#excelFile').onchange=e=>{const f=e.target.files?.[0];if(f)importFile(f)};
$('#reimportBtn').onclick=()=>$('#excelFile').click();
$('#stationSearch').oninput=renderSuggestions;$('#stationSearch').addEventListener('keydown',e=>{if(e.key!=='Enter')return;const a=searchPoi($('#stationSearch').value.trim());if(a.length){e.preventDefault();selectStart(a[0])}});$('#clearStation').onclick=()=>{$('#stationSearch').value='';renderSuggestions()};$('#geoBtn').onclick=useGeo;$('#searchNearbyBtn').onclick=searchNearby;$('#radius').onchange=()=>state.start&&state.rows.length&&state.poi.length&&searchNearby();
$('#optimizeRouteBtn').onclick=optimizeRoute;$('#reoptimizeBtn').onclick=optimizeRoute;$('#clearRouteBtn').onclick=clearRoute;
$('#copyShopping').onclick=async()=>{try{await navigator.clipboard.writeText(shoppingText());toast('Zekeringlijst gekopieerd')}catch{toast('Kopiëren niet gelukt')}};
$('#copyCompleted').onclick=async()=>{try{await navigator.clipboard.writeText(completedText());toast('Uitgevoerd-overzicht gekopieerd')}catch{toast('Kopiëren niet gelukt')}};$('#completedFilter').onchange=renderCompleted;
$('#settingsBtn').onclick=()=>{$('#settingsDlg').showModal()};document.querySelectorAll('[data-close]').forEach(b=>b.onclick=()=>document.getElementById(b.dataset.close).close());
$('#defaultQty').onchange=e=>{const n=Math.max(1,Math.min(12,Number(e.target.value)||3));e.target.value=n;saveSettings({defaultQty:n});if(state.routeFinalized){state.routeSnapshot=createShoppingSnapshot(routeOrderedGroups());saveRouteState();renderShoppingSnapshot()}renderCompleted()};
$('#kevinEmail').onchange=e=>saveSettings({kevinEmail:e.target.value.trim()});$('#themePref').onchange=e=>{saveSettings({theme:e.target.value});applyTheme(e.target.value)};$('#themeBtn').onclick=()=>{const now=document.body.classList.contains('light')?'dark':'light';saveSettings({theme:now});$('#themePref').value=now;applyTheme(now)};
$('#clearDataBtn').onclick=()=>{if(!confirm('Lokale werklijst op dit apparaat wissen? Je afgevinkte historie blijft bewaard.'))return;[LS.rows,LS.meta,LS.mapping].forEach(k=>localStorage.removeItem(k));state.rows=[];state.groups=[];state.nearbyCandidates=[];updateSummary();syncSearchButton();['resultsSection','routeBuilderSection','routeSection','actionsSection','engineeringSection','maintenanceSection','shoppingSection','unmatchedSection'].forEach(id=>$('#'+id)?.classList.add('hidden'));$('#settingsDlg').close();toast('Lokale werklijst gewist')};
$('#clearCompletedBtn').onclick=()=>{if(!confirm('Alle lokaal afgevinkte historie wissen?'))return;state.completed={};saveCompleted();rebuildGroups();updateSummary();renderCompleted();refreshVisibleData();toast('Afvinkhistorie gewist')};
$('#clearSavedRouteBtn').onclick=()=>{if(!confirm('De opgeslagen route op dit apparaat wissen?'))return;clearSavedRoute()};
$('#mappingBtn').onclick=()=>{openMapping();$('#settingsDlg').close()};$('#sheetSelect').onchange=e=>{const ws=state.workbook.getWorksheet(e.target.value);prepareMappingForSheet(ws,detectHeader(ws),null)};$('#headerRow').onchange=e=>{const ws=state.workbook.getWorksheet($('#sheetSelect').value);prepareMappingForSheet(ws,Number(e.target.value),null)};
$('#applyMapping').onclick=()=>{const m=mappingFromForm(),ws=state.workbook.getWorksheet(m.sheet);if(!m.stationCol||!m.fuseCol)return toast('Station en zekeringwaarde zijn verplicht');applyParsedRows(ws,m,state.pendingFilename||JSON.parse(localStorage.getItem(LS.meta)||'{}').filename||'Excel')};

loadLocal();loadPoi();
