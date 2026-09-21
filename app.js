const POI_SOURCES = [
  ['https://raw.githubusercontent.com/POIenexis/POI-zoeker/main/poi_e_station.csv','Station'],
  ['https://raw.githubusercontent.com/POIenexis/POI-zoeker/main/poi_e_verdeelkast.csv','Verdeelkast']
];
const LS = {
  rows:'zp_rows_v3', meta:'zp_meta_v3', settings:'zp_settings_v3', mapping:'zp_mapping_v3',
  completed:'zp_completed_v3'
};
const $ = s => document.querySelector(s);
const state = {
  poi:[], poiByCode:new Map(), rows:[], groups:[], start:null, workbook:null, currentSheet:null,
  mapping:null, planned:[], nearbyActions:[], nearbyEngineering:[], nearbyMaintenance:[],
  pendingFilename:'', completed:{}
};

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
  if(Number(c.theme)===9)return 'done';        // Excel Accent 6 = groen
  if(Number(c.theme)===7)return 'engineering';// Excel Accent 4 = oranje
  let argb=c.argb;if(!argb)return null;argb=String(argb).replace('#','').toUpperCase();if(argb.length===8)argb=argb.slice(2);if(argb.length!==6)return null;
  if(['92D050','00B050','70AD47','548235'].includes(argb))return 'done';
  if(['FFFF00','FFD966','FFF2CC'].includes(argb))return 'action';
  if(['FFC000','F4B183','ED7D31','FFFFCC99'].includes(argb))return 'engineering';
  if(['00B0F0','5B9BD5','4472C4','9DC3E6'].includes(argb))return 'maintenance';
  const r=parseInt(argb.slice(0,2),16)/255,g=parseInt(argb.slice(2,4),16)/255,b=parseInt(argb.slice(4,6),16)/255,max=Math.max(r,g,b),min=Math.min(r,g,b),d=max-min;let h=0;
  if(d){if(max===r)h=((g-b)/d)%6;else if(max===g)h=(b-r)/d+2;else h=(r-g)/d+4;h*=60;if(h<0)h+=360}
  const s=max===0?0:d/max,v=max;if(s<.22||v<.35)return null;
  if(h>=75&&h<=165)return 'done';if(h>=48&&h<70)return 'action';if(h>=25&&h<48)return 'engineering';if(h>=185&&h<=225)return 'maintenance';return null
}
function rowStatus(row,employeeText='',statusText=''){
  // In de actuele lijst is kolom K "Zekering vervangen medewerker". Een naam betekent gereed.
  if(String(employeeText||'').trim())return 'done';
  const found=[];for(let i=1;i<=Math.min(row.cellCount,40);i++){const s=fillStatus(row.getCell(i));if(s)found.push(s)}
  if(found.includes('done'))return 'done';
  if(found.includes('action'))return 'action';
  if(found.includes('engineering'))return 'engineering';
  if(found.includes('maintenance'))return 'maintenance';
  const t=norm(statusText);if(/gereed|gedaan|klaar|done|voltooid/.test(t))return 'done';if(/actie|wacht|uitzoek|control|probleem|afwijk/.test(t))return 'action';return 'open'
}
function headerScore(row){let s=0;row.eachCell({includeEmpty:false},c=>{const t=norm(val(c));if(t==='tewisselenstation'||t==='netstationlskast')s+=30;else if(t.includes('station'))s+=7;if(t.includes('richting')||t.includes('veld'))s+=6;if(t.includes('zekering')||t.includes('smeltveiligheid')||t==='huidigewaarde'||t.includes('ampere'))s+=7;if(t.includes('opmerking')||t.includes('medewerker'))s+=2});return s}
function worksheetScore(ws){let best=0;for(let r=1;r<=Math.min(ws.rowCount,40);r++)best=Math.max(best,headerScore(ws.getRow(r)));if(norm(ws.name)==='data')best+=30;return best}
function detectHeader(ws){let best={row:1,score:-1};for(let r=1;r<=Math.min(ws.rowCount,50);r++){const s=headerScore(ws.getRow(r));if(s>best.score)best={row:r,score:s}}return best.row}
function headersFor(ws,rowNum){const row=ws.getRow(rowNum),a=[];for(let i=1;i<=Math.max(row.cellCount,1);i++){const text=val(row.getCell(i)).trim();if(text)a.push({col:i,text})}return a}
function detectMapping(ws,headerRow){
  const hs=headersFor(ws,headerRow),findExact=names=>{for(const h of hs)if(names.includes(norm(h.text)))return h.col;return 0},pick=(tests,exclude=[])=>{for(const h of hs){const n=norm(h.text);if(exclude.some(x=>n.includes(x)))continue;if(tests.some(x=>x.test(n)))return h.col}return 0};
  let station=findExact(['netstationlskast','tewisselenstation']);
  if(!station)station=findExact(['stationnummer','stationsnummer','station','objectnummer','objectnr']);
  if(!station)station=pick([/station/,/locatiecode/,/^code$/,/^nummer$/],[/richting/]);
  const direction=findExact(['lsveldnummer','richtingnummer','richting','richtingnr'])||pick([/richtingnummer/,/^richting$/, /richtingnr/,/^veld$/, /veldnummer/,/groepnummer/]);
  const fuse=findExact(['huidigewaarde','lssmeltveiligheidinom','zekeringwaarde','zekering'])||pick([/zekeringwaarde/,/zekering/,/smeltveiligheid/,/ampere/,/nominaal/,/^waarde$/]);
  const qty=findExact(['aantal','stuks','hoeveelheid','qty'])||pick([/^aantal$/, /stuks/,/hoeveel/,/^qty$/]);
  const specialism=findExact(['opmeringspecialisme','opmerkingspecialisme'])||pick([/specialisme/]);
  const engineering=findExact(['opmerkingenvanuitengineering','opmerkingvanuitengineering'])||pick([/engineering/]);
  const employee=findExact(['zekeringvervangenmedewerker','medewerker'])||pick([/vervangenmedewerker/,/^medewerker$/]);
  const sourceStation=findExact(['netstation']);
  const cableGroup=findExact(['kabelgroep']);
  const city=findExact(['woonplaats']);
  const municipality=findExact(['gemeente']);
  const netType=findExact(['netwerktype']);
  // Oud bestand: één opmerkingenkolom.
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
    if(station)lastStation=station;else if((direction||fuse)&&lastStation)station=lastStation;
    if(!station||(!direction&&!fuse))continue;
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
    const headerRow=detectHeader(ws),m=detectMapping(ws,headerRow);state.currentSheet=ws;state.mapping=m;
    if(!m.stationCol||!m.fuseCol){openMapping();toast('Controleer de Excel-kolommen');return}
    applyParsedRows(ws,m,file.name);
  }catch(e){console.error(e);$('#fileInfo').textContent='Excel kon niet worden gelezen';toast('Excel kon niet worden gelezen')}
}
function applyParsedRows(ws,m,filename){
  const rows=parseWorksheet(ws,m);state.rows=rows;state.mapping=m;localStorage.setItem(LS.rows,JSON.stringify(rows));localStorage.setItem(LS.mapping,JSON.stringify(m));localStorage.setItem(LS.meta,JSON.stringify({filename:filename||JSON.parse(localStorage.getItem(LS.meta)||'{}').filename||'Excel',loadedAt:new Date().toISOString(),sheet:ws.name}));
  rebuildGroups();updateSummary();renderCompleted();if($('#mappingDlg').open)$('#mappingDlg').close();const c=statusCounts();toast(`${c.open} open · ${c.action} geel · ${c.engineering} engineering · ${c.maintenance} Henri · ${c.done} gereed`)
}
function extractCode(s){const t=String(s||'').trim();const m=t.match(/\b\d{1,4}\.(?:(?:SK|VK)\s*)?\d+\b/i)||t.match(/\b\d{1,4}[.\-\s]\d{2,5}\b/);return m?m[0].replace(/\s/g,''):t.split(/[;,]/)[0].trim()}
function matchPoi(station){
  const code=extractCode(station),k=norm(code);if(state.poiByCode.has(k))return state.poiByCode.get(k);
  const candidates=state.poi.filter(p=>norm(p.code).endsWith(k)||k.endsWith(norm(p.code)));if(candidates.length===1)return candidates[0];
  const digits=String(code).replace(/\D/g,'');if(digits){const same=state.poi.filter(p=>String(p.code).replace(/\D/g,'')===digits);if(same.length===1)return same[0]}return null
}
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
function selectStart(p,label){state.start={lat:p.lat,lon:p.lon,poi:p,label:label||`${p.code} ${p.name}`};$('#selectedStart').classList.remove('empty');$('#selectedStart').innerHTML=`<b>${esc(label||p.code+' '+p.name)}</b><small>${esc(p.type)} · ${esc([p.street,p.house,p.zip,p.city].filter(Boolean).join(' '))}</small>`;$('#suggestions').classList.add('hidden');$('#stationSearch').value='';$('#planBtn').disabled=!state.rows.length}
function useGeo(){if(!navigator.geolocation)return toast('Locatie niet beschikbaar');navigator.geolocation.getCurrentPosition(pos=>{state.start={lat:pos.coords.latitude,lon:pos.coords.longitude,poi:null,label:'Mijn locatie'};$('#selectedStart').classList.remove('empty');$('#selectedStart').innerHTML='<b>Mijn locatie</b><small>Startpunt op basis van GPS</small>';$('#planBtn').disabled=!state.rows.length;toast('Locatie actief')},()=>toast('Geef locatie-toegang in je browser'),{enableHighAccuracy:true,timeout:12000,maximumAge:60000})}

function eligibleRows(g,includeAction){return g.rows.filter(r=>effectiveOpen(r)||(includeAction&&r.status==='action'&&!isLocalDone(r)))}
function plan(){
  if(!state.start)return toast('Kies eerst een startpunt');if(!state.rows.length)return toast('Laad eerst de Excel');const radius=Number($('#radius').value),max=Number($('#maxStations').value),includeAction=$('#includeAction').checked;
  const nearby=state.groups.filter(g=>g.poi&&(g.open||g.action||g.engineering||g.maintenance)).map(g=>Object.assign({},g,{dist:hav(state.start.lat,state.start.lon,g.poi.lat,g.poi.lon)})).filter(g=>g.dist<=radius).sort((a,b)=>a.dist-b.dist);
  const work=nearby.filter(g=>eligibleRows(g,includeAction).length).slice(0,max),actions=nearby.filter(g=>g.action>0).slice(0,80),engineering=nearby.filter(g=>g.engineering>0).slice(0,80),maintenance=nearby.filter(g=>g.maintenance>0).slice(0,80);
  state.planned=work;state.nearbyActions=actions;state.nearbyEngineering=engineering;state.nearbyMaintenance=maintenance;renderPlan(work,actions,engineering,maintenance,includeAction);if(!work.length)toast(includeAction?'Geen resterende locaties binnen deze straal':'Geen open wissels binnen deze straal')
}
function navUrl(p){return `https://www.google.com/maps/dir/?api=1&destination=${p.lat},${p.lon}`}
function distanceLabel(d){return d<1?Math.round(d*1000)+' m':d.toFixed(1)+' km'}
function rowMeta(r){return [r.cableGroup,r.sourceStation&&r.sourceStation!==r.station?`voeding ${r.sourceStation}`:''].filter(Boolean).join(' · ')}
function workLine(r){return `<div class="work-line"><div><b>R${esc(r.direction)} · ${esc(r.fuse||'waarde onbekend')}</b>${rowMeta(r)?`<small>${esc(rowMeta(r))}</small>`:''}</div><button type="button" class="done-btn" data-rid="${esc(r.id)}">✓ Gedaan</button></div>`}
function actionLines(g){return g.rows.filter(r=>r.status==='action').map(r=>`<div class="status-line yellow-line"><b>R${esc(r.direction)} · ${esc(r.fuse||'geen waarde')}</b><span>${esc(r.specialism||r.statusText||'Geen toelichting ingevuld')}</span><strong>✉ Mail Kevin</strong></div>`).join('')}
function engineeringLines(g){return g.rows.filter(r=>r.status==='engineering').map(r=>`<div class="status-line orange-line"><b>R${esc(r.direction)} · ${esc(r.fuse||'geen waarde')}</b><span>${esc(r.engineering||'Engineering-opmerking ontbreekt')}</span></div>`).join('')}
function maintenanceLines(g){return g.rows.filter(r=>r.status==='maintenance').map(r=>`<div class="status-line blue-line"><b>R${esc(r.direction)} · ${esc(r.fuse||'geen waarde')}</b><span>Wordt opgepakt door Henri van der Vleuten tijdens onderhoud.</span></div>`).join('')}
function mailKevinForGroup(g){
  const rows=g.rows.filter(r=>r.status==='action'),settings=getSettings(),subject=`Zekeringwissel ${g.poi?.code||g.station} - actie nodig`,body=[`Hoi Kevin,`,``,`Bij ${g.poi?.code||g.station} zijn de volgende zekeringwissels niet uitgevoerd:`,...rows.map(r=>`- richting ${r.direction}, ${r.fuse||'waarde onbekend'}: ${r.specialism||r.statusText||'geen toelichting'}`),``,`Groet`].join('\n');
  location.href=`mailto:${encodeURIComponent(settings.kevinEmail||'')}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}
function renderStatusSection(sectionId,countId,listId,groups,kind){
  const sec=$(sectionId);if(!groups.length){sec.classList.add('hidden');$(listId).innerHTML='';return}sec.classList.remove('hidden');$(countId).textContent=`${groups.length} locatie${groups.length===1?'':'s'}`;
  $(listId).innerHTML=groups.map(g=>{const lines=kind==='action'?actionLines(g):kind==='engineering'?engineeringLines(g):maintenanceLines(g),cls=kind==='action'?'action-card':kind==='engineering'?'engineering-card':'maintenance-card';return `<article class="station-card ${cls}"><div class="station-top"><div><div class="station-code">${esc(g.poi.code||g.station)}</div><div class="station-name">${esc(g.poi.type)} · ${esc(g.poi.name||'')} · ${esc(g.poi.city||'')}</div></div><div class="distance ${kind}-distance">${distanceLabel(g.dist)}</div></div><div class="status-lines">${lines}</div><div class="station-actions"><a class="navlink" target="_blank" rel="noopener" href="${navUrl(g.poi)}">↗ Navigeer</a>${kind==='action'?`<button type="button" class="navlink mail-kevin" data-group="${esc(g.key)}">✉ Maak mail</button>`:''}</div></article>`}).join('');
  if(kind==='action')$(listId).querySelectorAll('.mail-kevin').forEach(b=>b.addEventListener('click',()=>{const g=groups.find(x=>x.key===b.dataset.group);if(g)mailKevinForGroup(g)}))
}
function renderPlan(groups,actions,engineering,maintenance,includeAction){
  $('#resultsSection').classList.remove('hidden');$('#shoppingSection').classList.remove('hidden');$('#resultCount').textContent=`${groups.length} locatie${groups.length===1?'':'s'}`;
  $('#stationResults').innerHTML=groups.map(g=>{const active=eligibleRows(g,includeAction),warnings=g.rows.filter(r=>r.status==='action');return `<article class="station-card"><div class="station-top"><div><div class="station-code">${esc(g.poi.code||g.station)}</div><div class="station-name">${esc(g.poi.type)} · ${esc(g.poi.name||'')} · ${esc(g.poi.city||'')}</div></div><div class="distance">${distanceLabel(g.dist)}</div></div><div class="work-list">${active.map(workLine).join('')}</div>${!includeAction&&warnings.length?`<div class="warnline">⚠ ${warnings.length} gele regel${warnings.length===1?'':'s'} — niet meenemen; mail Kevin</div>`:''}<div class="station-actions"><a class="navlink" target="_blank" rel="noopener" href="${navUrl(g.poi)}">↗ Navigeer</a></div></article>`}).join('')||'<div class="muted">Geen open wissels gevonden.</div>';
  $('#stationResults').querySelectorAll('.done-btn').forEach(b=>b.addEventListener('click',()=>completeRow(b.dataset.rid)));
  renderStatusSection('#actionsSection','#actionCount','#actionResults',actions,'action');
  renderStatusSection('#engineeringSection','#engineeringCount','#engineeringResults',engineering,'engineering');
  renderStatusSection('#maintenanceSection','#maintenanceCount','#maintenanceResults',maintenance,'maintenance');
  renderShopping(groups,includeAction);
}
function isCountableFuse(r){return Number.isFinite(fuseAmps(r.fuse))&&!/restbak|onbekend/i.test(r.fuse)}
function renderShopping(groups,includeAction){
  const counts=new Map(),settings=getSettings();let total=0,directions=0,unknown=0;
  for(const g of groups)for(const r of eligibleRows(g,includeAction)){const qty=r.qty||settings.defaultQty||3;directions++;if(!isCountableFuse(r)){unknown++;continue}counts.set(r.fuse,(counts.get(r.fuse)||0)+qty);total+=qty}
  const sorted=[...counts.entries()].sort(fuseSort);$('#shoppingList').innerHTML=(sorted.length?sorted.map(([f,q])=>`<div class="shopping-row"><b>${esc(f)}</b><strong>${q}×</strong></div>`).join(''):'<div class="muted">Geen zekeringen te tellen.</div>')+`<div class="shopping-total">Totaal: ${total} zekeringen · ${directions} richtingen${unknown?` · ${unknown} richting${unknown===1?'':'en'} met onbekende/restbak-waarde`:''}</div>`;$('#shoppingNote').textContent=`De app rekent met ${settings.defaultQty||3} zekeringen per richting als er geen aantalkolom is. Afgevinkte regels verdwijnen direct uit deze lijst. Geel, Engineering en Henri worden standaard niet meegeteld.`
}
function shoppingText(){const includeAction=$('#includeAction').checked,s=getSettings(),counts=new Map();let unknown=0;for(const g of state.planned)for(const r of eligibleRows(g,includeAction)){if(!isCountableFuse(r)){unknown++;continue}const q=r.qty||s.defaultQty||3;counts.set(r.fuse,(counts.get(r.fuse)||0)+q)}let total=0;const lines=[...counts.entries()].sort(fuseSort).map(([f,q])=>{total+=q;return `${q}x ${f}`});if(unknown)lines.push(`${unknown} richting(en) met onbekende/restbak-waarde controleren`);return `Zekeringen meenemen\n${lines.join('\n')}\n\nTotaal bekende zekeringen: ${total} stuks`}

function completeRow(id){
  const r=state.rows.find(x=>x.id===id);if(!r||r.status!=='open'||isLocalDone(r))return;
  const settings=getSettings();state.completed[id]={id,station:r.station,direction:r.direction,fuse:r.fuse,qty:r.qty||settings.defaultQty||3,cableGroup:r.cableGroup||'',sourceStation:r.sourceStation||'',city:r.city||'',completedAt:new Date().toISOString()};saveCompleted();rebuildGroups();updateSummary();renderCompleted();if(state.start)plan();toast(`${r.station} R${r.direction} afgevinkt`)
}
function undoCompleted(id){delete state.completed[id];saveCompleted();rebuildGroups();updateSummary();renderCompleted();if(state.start)plan();toast('Afvinken hersteld')}
function dateKey(d){const x=new Date(d);return `${x.getFullYear()}-${String(x.getMonth()+1).padStart(2,'0')}-${String(x.getDate()).padStart(2,'0')}`}
function completedItems(){const mode=$('#completedFilter')?.value||'today',all=Object.values(state.completed).sort((a,b)=>new Date(b.completedAt)-new Date(a.completedAt));if(mode==='all')return all;const today=dateKey(new Date());return all.filter(x=>dateKey(x.completedAt)===today)}
function completedText(){const items=completedItems(),counts=new Map(),stations=new Set();for(const x of items){stations.add(x.station);if(Number.isFinite(fuseAmps(x.fuse))&&!/restbak|onbekend/i.test(x.fuse))counts.set(x.fuse,(counts.get(x.fuse)||0)+(x.qty||0))}const lines=items.map(x=>`${x.station} - richting ${x.direction} - ${x.fuse||'waarde onbekend'} - ${new Date(x.completedAt).toLocaleTimeString('nl-NL',{hour:'2-digit',minute:'2-digit'})}`);const fuse=[...counts.entries()].sort(fuseSort).map(([f,q])=>`${q}x ${f}`);return `Uitgevoerd overzicht\n${lines.join('\n')}\n\n${items.length} richtingen · ${stations.size} locaties${fuse.length?`\nGebruikt:\n${fuse.join('\n')}`:''}`}
function renderCompleted(){
  const sec=$('#completedSection');if(!sec)return;const items=completedItems();$('#completedCount').textContent=`${items.length} richting${items.length===1?'':'en'}`;sec.classList.toggle('empty-completed',!items.length);
  const counts=new Map(),stations=new Set();for(const x of items){stations.add(x.station);if(Number.isFinite(fuseAmps(x.fuse))&&!/restbak|onbekend/i.test(x.fuse))counts.set(x.fuse,(counts.get(x.fuse)||0)+(x.qty||0))}
  $('#completedSummary').innerHTML=`<b>${stations.size} locatie${stations.size===1?'':'s'}</b><span>${items.length} richtingen uitgevoerd</span>${counts.size?`<small>${[...counts.entries()].sort(fuseSort).map(([f,q])=>`${q}× ${esc(f)}`).join(' · ')}</small>`:''}`;
  $('#completedList').innerHTML=items.length?items.map(x=>`<div class="completed-row"><div><b>${esc(x.station)} · R${esc(x.direction)}</b><span>${esc(x.fuse||'waarde onbekend')}${x.cableGroup?` · ${esc(x.cableGroup)}`:''}</span><small>${new Date(x.completedAt).toLocaleString('nl-NL',{dateStyle:'short',timeStyle:'short'})}</small></div><button type="button" class="undo-btn" data-rid="${esc(x.id)}">↶ Herstel</button></div>`).join(''):'<div class="muted">Nog niets afgevinkt in deze periode.</div>';
  $('#completedList').querySelectorAll('.undo-btn').forEach(b=>b.addEventListener('click',()=>undoCompleted(b.dataset.rid)))
}

function fillSelect(sel,headers,selected,allowEmpty=true){sel.innerHTML=(allowEmpty?'<option value="0">— niet gebruiken —</option>':'')+headers.map(h=>`<option value="${h.col}" ${h.col===Number(selected)?'selected':''}>${h.col}: ${esc(h.text)}</option>`).join('')}
function openMapping(){if(!state.workbook){toast('Kies eerst opnieuw de Excel');return}const dlg=$('#mappingDlg'),sheetSel=$('#sheetSelect');sheetSel.innerHTML=state.workbook.worksheets.map(ws=>`<option value="${esc(ws.name)}" ${ws.name===state.mapping?.sheet?'selected':''}>${esc(ws.name)}</option>`).join('');const ws=state.workbook.getWorksheet(state.mapping?.sheet)||state.workbook.worksheets[0];prepareMappingForSheet(ws,state.mapping?.headerRow||detectHeader(ws),state.mapping);dlg.showModal()}
function prepareMappingForSheet(ws,headerRow,current){state.currentSheet=ws;$('#headerRow').value=headerRow;const auto=current&&current.sheet===ws.name?current:detectMapping(ws,headerRow),hs=headersFor(ws,headerRow);fillSelect($('#stationCol'),hs,auto.stationCol,false);fillSelect($('#directionCol'),hs,auto.directionCol);fillSelect($('#fuseCol'),hs,auto.fuseCol,false);fillSelect($('#qtyCol'),hs,auto.qtyCol)}
function mappingFromForm(){const base=detectMapping(state.workbook.getWorksheet($('#sheetSelect').value),Number($('#headerRow').value));return Object.assign(base,{sheet:$('#sheetSelect').value,headerRow:Number($('#headerRow').value),stationCol:Number($('#stationCol').value),directionCol:Number($('#directionCol').value),fuseCol:Number($('#fuseCol').value),qtyCol:Number($('#qtyCol').value)})}

function applyTheme(v){const r=v==='system'?(matchMedia('(prefers-color-scheme: light)').matches?'light':'dark'):v;document.body.classList.toggle('light',r==='light');document.querySelector('meta[name="theme-color"]')?.setAttribute('content',r==='light'?'#edf3f5':'#07131b');$('#themeBtn').textContent=r==='light'?'☾':'☼'}
function loadLocal(){try{state.rows=JSON.parse(localStorage.getItem(LS.rows)||'[]');state.mapping=JSON.parse(localStorage.getItem(LS.mapping)||'null')}catch{state.rows=[]}loadCompleted();rebuildGroups();updateSummary();const s=getSettings();$('#defaultQty').value=s.defaultQty;$('#themePref').value=s.theme;$('#kevinEmail').value=s.kevinEmail||'';applyTheme(s.theme);$('#planBtn').disabled=!state.rows.length||!state.start;renderCompleted()}

$('#excelFile').onchange=e=>{const f=e.target.files?.[0];if(f)importFile(f)};
$('#reimportBtn').onclick=()=>$('#excelFile').click();
$('#stationSearch').oninput=renderSuggestions;$('#clearStation').onclick=()=>{$('#stationSearch').value='';renderSuggestions()};$('#geoBtn').onclick=useGeo;$('#planBtn').onclick=plan;
$('#radius').onchange=()=>state.start&&state.rows.length&&plan();$('#maxStations').onchange=()=>state.start&&state.rows.length&&plan();$('#includeAction').onchange=()=>state.start&&state.rows.length&&plan();
$('#copyShopping').onclick=async()=>{try{await navigator.clipboard.writeText(shoppingText());toast('Boodschappenlijst gekopieerd')}catch{toast('Kopiëren niet gelukt')}};
$('#copyCompleted').onclick=async()=>{try{await navigator.clipboard.writeText(completedText());toast('Uitgevoerd-overzicht gekopieerd')}catch{toast('Kopiëren niet gelukt')}};
$('#completedFilter').onchange=renderCompleted;
$('#settingsBtn').onclick=()=>{$('#settingsDlg').showModal()};document.querySelectorAll('[data-close]').forEach(b=>b.onclick=()=>document.getElementById(b.dataset.close).close());
$('#defaultQty').onchange=e=>{const n=Math.max(1,Math.min(12,Number(e.target.value)||3));e.target.value=n;saveSettings({defaultQty:n});if(state.start&&state.rows.length)plan();renderCompleted()};
$('#kevinEmail').onchange=e=>saveSettings({kevinEmail:e.target.value.trim()});
$('#themePref').onchange=e=>{saveSettings({theme:e.target.value});applyTheme(e.target.value)};$('#themeBtn').onclick=()=>{const now=document.body.classList.contains('light')?'dark':'light';saveSettings({theme:now});$('#themePref').value=now;applyTheme(now)};
$('#clearDataBtn').onclick=()=>{if(!confirm('Lokale werklijst op dit apparaat wissen? Je afgevinkte historie blijft bewaard.'))return;[LS.rows,LS.meta,LS.mapping].forEach(k=>localStorage.removeItem(k));state.rows=[];state.groups=[];state.planned=[];state.nearbyActions=[];state.nearbyEngineering=[];state.nearbyMaintenance=[];updateSummary();['resultsSection','actionsSection','engineeringSection','maintenanceSection','shoppingSection','unmatchedSection'].forEach(id=>$('#'+id)?.classList.add('hidden'));$('#settingsDlg').close();toast('Lokale werklijst gewist')};
$('#clearCompletedBtn').onclick=()=>{if(!confirm('Alle lokaal afgevinkte historie wissen?'))return;state.completed={};saveCompleted();rebuildGroups();updateSummary();renderCompleted();if(state.start&&state.rows.length)plan();toast('Afvinkhistorie gewist')};
$('#mappingBtn').onclick=()=>{openMapping();$('#settingsDlg').close()};$('#sheetSelect').onchange=e=>{const ws=state.workbook.getWorksheet(e.target.value);prepareMappingForSheet(ws,detectHeader(ws),null)};$('#headerRow').onchange=e=>{const ws=state.workbook.getWorksheet($('#sheetSelect').value);prepareMappingForSheet(ws,Number(e.target.value),null)};
$('#applyMapping').onclick=()=>{const m=mappingFromForm(),ws=state.workbook.getWorksheet(m.sheet);if(!m.stationCol||!m.fuseCol)return toast('Station en zekeringwaarde zijn verplicht');applyParsedRows(ws,m,state.pendingFilename||JSON.parse(localStorage.getItem(LS.meta)||'{}').filename||'Excel')};

loadLocal();loadPoi();
