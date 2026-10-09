(function(){const module={exports:{}};
'use strict';
// Pure planner. No workbook mutation, external IO, or first-match lookup.
const HEADERS = ['PROJECT KEY', 'MÃ CT/DA', 'ĐỊA CHỈ THEO CTDA'];
const SHEETS = ['ĐỀ NGHỊ MUA', 'ĐỀ NGHỊ SX', 'ĐỀ NGHỊ XUẤT'];
const text = x => String(x ?? '').trim();
function resolve(headers) {
  return Object.fromEntries(HEADERS.map((h,i) => {
    const matches = headers.map((v,j) => v === h ? j : -1).filter(j => j >= 0);
    if (matches.length !== 1) throw Error('HEADER_NOT_UNIQUE: ' + h);
    return [['key','display','address'][i], matches[0]];
  }));
}
function catalog(rows) {
  const byKey = new Map(), byDisplay = new Map();
  for (const r of rows) {
    if (text(r.source) !== 'CTDA') continue;
    const key = text(r.key), display = text(r.display), address = text(r.address);
    if (!key || !display) throw Error('CATALOG_MISSING_ID');
    if([key,display,address].some(v=>v.startsWith('='))) throw Error('FORMULA_LIKE_CATALOG_VALUE');
    if (!byKey.has(key)) byKey.set(key,{key,displays:new Set(),addresses:new Set()});
    const p = byKey.get(key); p.displays.add(display); if(address) p.addresses.add(address);
  }
  for (const p of byKey.values()) {
    if(p.displays.size !== 1) throw Error('AMBIGUOUS_DISPLAY: '+p.key);
    p.display = [...p.displays][0];
    if(byDisplay.has(p.display)) throw Error('DUPLICATE_DISPLAY: '+p.display);
    byDisplay.set(p.display,p);
  }
  if(!byKey.size) throw Error('EMPTY_CATALOG');
  return {byKey,byDisplay};
}
function plan(row, changed, index, owned = null) {
  const out={patch:{},warnings:[],ownership:null};
  const key=text(row.key), display=text(row.display), address=text(row.address);
  const k=changed.includes('key'), d=changed.includes('display');
  if(!k && !d) return out; // Manual address edit relinquishes ownership.
  let p;
  if(k && d) {
    if(!key && !display) p=null;
    else { p=index.byKey.get(key); if(!p || p.display!==display) {out.warnings.push('CONFLICTING_PROJECT_INPUT');return out;} }
  } else p=(k ? index.byKey.get(key) : index.byDisplay.get(display));
  const blank=k&&!d&&!key || d&&!k&&!display || k&&d&&!key&&!display;
  if(!p && !blank) {out.warnings.push('UNKNOWN_PROJECT');return out;}
  out.patch.key=p?p.key:''; out.patch.display=p?p.display:'';
  // Ownership is session-local and accepted only if the non-edited identity still matches.
  const auto=owned && address===owned.address && (!k?key===owned.key:true) && (!d?display===owned.display:true);
  const explicitAddress=changed.includes('address');
  if(explicitAddress) {out.warnings.push('MANUAL_ADDRESS_PRESERVED');}
  else if(!address || auto) {
    out.patch.address=p && p.addresses.size===1 ? [...p.addresses][0] : '';
    if(p && p.addresses.size!==1) out.warnings.push(p.addresses.size?'AMBIGUOUS_ADDRESS':'MISSING_ADDRESS');
    out.ownership=p?{key:p.key,display:p.display,address:out.patch.address}:null;
  } else out.warnings.push('EXISTING_ADDRESS_PRESERVED_CHECK_PROJECT');
  for(const f of Object.keys(out.patch)) if(out.patch[f]===row[f]) delete out.patch[f];
  return out;
}
module.exports={HEADERS,SHEETS,resolve,catalog,plan};
function approvalColumns(headers, sheetName) {
 const find=h=>{const a=headers.map((v,i)=>v===h?i:-1).filter(i=>i>=0);if(a.length!==1)throw Error('HEADER_NOT_UNIQUE:'+h);return a[0];};
 return {status:find('TÌNH TRẠNG'),date:find(sheetName==='ĐỀ NGHỊ XUẤT'?'NGÀY DUYỆT NGUỒN':'NGÀY DUYỆT')};
}
function vietnamSerial(now=new Date()) {
 const parts=new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Ho_Chi_Minh',year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(now);
 const v=Object.fromEntries(parts.map(p=>[p.type,p.value]));
 return Math.round((Date.UTC(+v.year,+v.month-1,+v.day)-Date.UTC(1899,11,30))/86400000);
}
function approvalPlan(row, changed, previousStatus, now=new Date()) {
 if(changed.includes('date'))return {}; // Explicit manual dates win over simultaneous status pastes.
 if(!changed.includes('status') || previousStatus===undefined || text(row.status)===text(previousStatus))return {};
 if(text(row.status)==='Đã duyệt')return {date:vietnamSerial(now)};
 if(text(row.status)==='Chưa duyệt')return {date:''};
 return {};
}
Object.assign(module.exports,{approvalColumns,vietnamSerial,approvalPlan});
// Recover catalog-derived ownership without writing any workbook cell.
// Exact full identity + one unique address; custom/ambiguous values remain manual.
function seedOwnership(row,index) {
 const p=index.byKey.get(text(row.key));
 if(!p || text(row.display)!==p.display || p.addresses.size!==1)return null;
 const expected=[...p.addresses][0];
 if(!expected || text(row.address)!==expected)return null;
 return {key:p.key,display:p.display,address:text(row.address)};
}
module.exports.seedOwnership=seedOwnership;
module.exports.build='HCKT-20261009-R2';

globalThis.HcktProjectEngine=module.exports;})();
(function(){const module={exports:{}};
/* Local draft only. Caller must supply approved workbook identity guard and loaded engine.
   No startup edits. No settings, startup behavior, remote calls, baseline tables or UID seeding. */
async function startProjectOnly(Excel, engine, options) {
  if (!options || options.enabled !== true || typeof options.verifyIdentity !== 'function') throw Error('EXPLICIT_BINDING_REQUIRED');
  await options.verifyIdentity();
  const report=options.report || (()=>{}), owners=new Map(), subscriptions=[], statuses=new Map(), initialEnds=new Map(), processed=new Map(), journal=[];
  let stopped=false, paused=false, chain=Promise.resolve(), revision=0,pendingEvents=0;
  if (!options.supportsTriggerSource) throw Error('TRIGGER_SOURCE_SUPPORT_REQUIRED');
  const bounds=a=>{const m=String(a).split('!').pop().replace(/\$/g,'').match(/^([A-Z]+)([0-9]+)(?::([A-Z]+)([0-9]+))?$/i);if(!m)throw Error('UNSUPPORTED_EVENT_RANGE');const n=x=>[...x.toUpperCase()].reduce((v,c)=>v*26+c.charCodeAt(0)-64,0)-1;return{r0:+m[2]-1,r1:+(m[4]||m[2])-1,c0:n(m[1]),c1:n(m[3]||m[1])};};
  const fields=['key','display','address','status','date'];
  const readCatalog=async ctx=>{
    const sheet=ctx.workbook.worksheets.getItem('_A00_SOURCE_QA');
    const table=sheet.tables.getItem('tblA00IntegrationAudit');
    const head=table.getHeaderRowRange(), body=table.getDataBodyRange();
    head.load('values'); body.load('values'); await ctx.sync();
    const hs=head.values[0], names=['Source','ProjectKey','ProjectDisplay','Address'];
    const ix=names.map(n=>{const a=hs.reduce((a,v,i)=>v===n?a.concat(i):a,[]); if(a.length!==1)throw Error('SOURCE_HEADER:'+n);return a[0];});
    return engine.catalog(body.values.map(r=>({source:r[ix[0]],key:r[ix[1]],display:r[ix[2]],address:r[ix[3]]})));
  };
  const seedRows=(name,headers,used,source)=>{
    const col={...engine.resolve(headers),...engine.approvalColumns(headers,name)};
    initialEnds.set(name,used.rowIndex+used.values.length);
    for(let i=0;i<used.values.length;i++){
      const r=used.rowIndex+i;if(r<2)continue;
      const id=name+':'+r, row=Object.fromEntries(fields.map(f=>[f,used.values[i][col[f]-used.columnIndex]??'']));
      statuses.set(id,row.status);
      const owned=engine.seedOwnership(row,source);if(owned)owners.set(id,owned);else owners.delete(id);
    }
  };
  const process=async (event, expectedRevision)=>{
    if(stopped || paused || event.triggerSource==='ThisLocalAddin') return;
    if(event.source==='Remote') throw Error('REMOTE_EDIT_SINGLE_WRITER_STOP');
    if(event.changeType!=='RangeEdited') {owners.clear();statuses.clear();report('STRUCTURE_CHANGED_RECHECK');return true;}
    return await Excel.run(async ctx=>{
      await options.verifyIdentity();
      const sheet=ctx.workbook.worksheets.getItem(event.worksheetId);
      sheet.load('name'); const hit=sheet.getRange(event.address);hit.load(['rowIndex','columnIndex','rowCount','columnCount']);
      const header=sheet.getRange('A2:AZ2');header.load('values'); await ctx.sync();
      if(!engine.SHEETS.includes(sheet.name))return;
      const col={...engine.resolve(header.values[0]),...engine.approvalColumns(header.values[0],sheet.name)};
      const changed=fields.filter(f=>col[f]>=hit.columnIndex && col[f]<hit.columnIndex+hit.columnCount);
      if(!changed.length || hit.rowIndex+hit.rowCount<=2)return;
      if(hit.rowIndex<2)throw Error('HEADER_EDIT_STOP');
      if(hit.rowCount>5000)throw Error('BATCH_TOO_LARGE');
      const source=await readCatalog(ctx);
      const width=Math.max(...Object.values(col))+1;
      const range=sheet.getRangeByIndexes(hit.rowIndex,0,hit.rowCount,width);range.load(['values','formulas','numberFormat']);await ctx.sync();
      const pending=[];
      range.values.forEach((values,i)=>{
        const row=Object.fromEntries(fields.map(f=>[f,values[col[f]]]));
        const id=sheet.name+':'+(hit.rowIndex+i);
        if((processed.get(id)||0)>=event.sequence)return;
        const events=journal.filter(e=>e.worksheetId===event.worksheetId&&e.sequence>(processed.get(id)||0)&&e.sequence<=expectedRevision&&e.bounds.r0<=hit.rowIndex+i&&e.bounds.r1>=hit.rowIndex+i);
        let rowChanged=[...new Set(events.flatMap(e=>fields.filter(f=>col[f]>=e.bounds.c0&&col[f]<=e.bounds.c1)))];
        const identityEvents=events.filter(e=>['key','display'].some(f=>col[f]>=e.bounds.c0&&col[f]<=e.bounds.c1));
        if(identityEvents.length){const last=identityEvents[identityEvents.length-1];rowChanged=rowChanged.filter(f=>!['key','display'].includes(f)).concat(['key','display'].filter(f=>col[f]>=last.bounds.c0&&col[f]<=last.bounds.c1));}
        if(fields.some(f=>String(range.formulas[i][col[f]]||'').startsWith('=')))throw Error('FORMULA_CELL_STOP');
        const result=engine.plan(row,rowChanged,source,owners.get(id));
        let previous=statuses.has(id)?statuses.get(id):(hit.rowIndex+i>=(initialEnds.get(sheet.name)??Infinity)?'':undefined);
        const firstStatus=events.slice().reverse().find(e=>e.bounds.c0===col.status&&e.bounds.c1===col.status&&e.bounds.r0===e.bounds.r1&&e.details);if(firstStatus)previous=firstStatus.details.valueBefore;
        Object.assign(result.patch,engine.approvalPlan(row,rowChanged,previous));
        if(rowChanged.includes('status')&&previous===undefined)result.warnings.push('PREVIOUS_STATUS_UNKNOWN_NO_STAMP');
        pending.push({i,id,row,result,rowChanged,dateFormat:range.numberFormat?.[i]?.[col.date]});
      });
      // Global optimistic preflight before any write; only the three approved cells are compared.
      const check=sheet.getRangeByIndexes(hit.rowIndex,0,hit.rowCount,width);check.load(['values','formulas','numberFormat']);await ctx.sync();
      if(pending.some(p=>fields.some(f=>check.values[p.i][col[f]]!==p.row[f])))return false;
      if(pending.some(p=>fields.some(f=>String(check.formulas[p.i][col[f]]||'').startsWith('='))))throw Error('CONCURRENT_FORMULA_STOP');
      if(stopped||paused)return true;
      if(revision!==expectedRevision) return false;
      for(const p of pending) for(const [f,v] of Object.entries(p.result.patch)) {const cell=sheet.getCell(hit.rowIndex+p.i,col[f]);cell.values=[[v]];if(f==='date'&&typeof v==='number'&&check.numberFormat?.[p.i]?.[col.date]==='General')cell.numberFormat=[['dd/mm/yyyy']];}
      await ctx.sync();
      const verify=sheet.getRangeByIndexes(hit.rowIndex,0,hit.rowCount,width);verify.load('values');await ctx.sync();
      for(const p of pending) {
        if(Object.entries(p.result.patch).some(([f,v])=>verify.values[p.i][col[f]]!==v))throw Error('READBACK_MISMATCH');
        if(p.rowChanged.includes('status'))statuses.set(p.id,p.row.status);processed.set(p.id,expectedRevision);
        if(p.rowChanged.some(f=>['key','display','address'].includes(f))){if(p.result.ownership)owners.set(p.id,p.result.ownership);else owners.delete(p.id);}
        if(p.result.warnings.length)report({sheet:sheet.name,row:hit.rowIndex+p.i+1,warnings:p.result.warnings});
      }
      return true;
    });
  };
  try { await Excel.run(async ctx=>{
    const startupSource=await readCatalog(ctx); // Validate and seed memory only; no cell writes.
    for(const name of engine.SHEETS){
      const sheet=ctx.workbook.worksheets.getItem(name), header=sheet.getRange('A2:AZ2');header.load('values');const used=sheet.getUsedRange(true);used.load(['values','rowIndex','columnIndex']);await ctx.sync();seedRows(name,header.values[0],used,startupSource);
      const invalidate=()=>{++revision;paused=true;report({state:'PAUSED_RESEED',message:'Đang cập nhật sau khi sắp xếp; chờ thông báo sẵn sàng trước khi sửa.'});owners.clear();statuses.clear();processed.clear();journal.length=0;chain=chain.then(async()=>{await Excel.run(async fresh=>{const refreshedSource=await readCatalog(fresh);for(const sn of engine.SHEETS){const sh=fresh.workbook.worksheets.getItem(sn),h=sh.getRange('A2:AZ2'),u=sh.getUsedRange(true);h.load('values');u.load(['values','rowIndex','columnIndex']);await fresh.sync();seedRows(sn,h.values[0],u,refreshedSource);}});paused=false;if(!stopped)report({state:'READY_AFTER_STRUCTURE'});}).catch(e=>{stopped=true;report({stopped:true,error:String(e)});});};
      const token=sheet.onChanged.add(event=>{if(event.triggerSource==='ThisLocalAddin')return;if(event.source==='Remote'){stopped=true;++revision;report({stopped:true,error:'REMOTE_EDIT_SINGLE_WRITER_STOP'});return;}if(event.changeType!=='RangeEdited'){invalidate();return;}if(paused){stopped=true;++revision;report({stopped:true,error:'EDIT_DURING_STRUCTURE_RESEED_RESTART_AND_RESELECT_STATUS'});return;}event.sequence=++revision;try{event.bounds=bounds(event.address);}catch(e){stopped=true;report({stopped:true,error:String(e)});return;}journal.push(event);pendingEvents++;chain=chain.then(async()=>{for(let n=0;n<3;n++){if(await process(event,revision)!==false)return;}report({warning:'BUSY_ROW_RESELECT_TO_UPDATE',address:event.address});}).catch(e=>{stopped=true;report({error:String(e),stopped:true});}).finally(()=>{if(--pendingEvents===0){journal.length=0;processed.clear();}});});subscriptions.push(token);
      const sorted=sheet.onRowSorted.add(invalidate);subscriptions.push(sorted);
    }
    await ctx.sync();
  }); } catch(error) {stopped=true;for(const s of subscriptions){try{s.remove();await s.context.sync();}catch(_){}}throw error;}
  return {stop:async()=>{stopped=true;++revision;await chain;for(const s of subscriptions){s.remove();await s.context.sync();}owners.clear();statuses.clear();}};
}
if(typeof module!=='undefined')module.exports={startProjectOnly};

globalThis.HcktProjectHost=module.exports;})();
(function(){const module={exports:{}};
'use strict';
// Purchase-only planner. F is evidence, never an output. Persist every reservation before W.
const BUILD='HCKT-20261009-R3.1-CANDIDATE';
const HEADER='SỐ ĐỀ NGHỊ TỰ ĐỘNG';
const JOURNAL_HEADERS=['HCKT_PURCHASE_NUMBER_V1','ProjectKey','Sequence','RequestNumber','Kind','UTC','OperationId','Version'];
const GATE={required:['key','display','content','requester','submissionDate'],approvalRequired:false};
const str=x=>String(x??'').trim();
const formatNumber=(key,seq)=>key+'-'+String(seq).padStart(2,'0');
function parseNumber(value,key){
 if(typeof value!=='string')return null;
 const m=value.trim().match(/^(CT\d{3})-(\d{1,6})$/);
 if(!m||m[1].toUpperCase()!==key||+m[2]<1)return null;
 return {key,seq:+m[2],number:formatNumber(key,+m[2])};
}
function parseSource(value,key){
 if(typeof value!=='string')return null;
 const m=value.trim().match(/^(?:(?:ĐNMVT|DNMVT|DNMK|DNMN)\/(?:(?:XSX)[ -])?)?(?:CT)?(\d{3})-(\d{1,6})$/i);
 if(!m||'CT'+m[1]!==key||+m[2]<1)return null;
 return {key,seq:+m[2],number:formatNumber(key,+m[2])};
}
function validDate(v){
 if(typeof v==='number')return Number.isInteger(v)&&v>=36526&&v<=109574;
 const m=str(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);if(!m)return false;
 const d=new Date(Date.UTC(+m[3],+m[2]-1,+m[1]));return +m[3]>=2000&&+m[3]<=2199&&d.getUTCDate()===+m[1]&&d.getUTCMonth()===+m[2]-1;
}
function complete(row,catalog){const p=catalog.byKey.get(str(row.key));return !!p&&p.display===str(row.display)&&!!str(row.content)&&!!str(row.requester)&&validDate(row.submissionDate);}
function readJournal(values){
 const out=[];for(const r of values){if(!r.some(v=>str(v)))continue;const [uid,key,seq,number,kind,utc,op,version]=r;
 if(!['SEED','AUTO','MANUAL'].includes(kind)||!/^CT\d{3,}$/.test(key)||!Number.isInteger(seq)||seq<1||seq>999999||parseNumber(number,key)?.seq!==seq||!str(uid)||!str(op)||version!==1)throw Error('JOURNAL_CORRUPT_STOP');
 out.push({uid,key,seq,number:formatNumber(key,seq),kind,utc,op,version});}return out;
}
function serialize(e){return [e.uid,e.key,e.seq,e.number,e.kind,e.utc,e.op,1];}
function plan({rows,selected=[],journal=[],baseline=new Set(),catalog,now=new Date().toISOString(),operation='preview',uidFactory=()=>{throw Error('UID_FACTORY_REQUIRED');}}){
 const entries=[],writes=[],identityWrites=[],warnings=[],max=new Map(),used=new Map(),uids=new Map(),latest=new Map();
 const reserve=e=>{max.set(e.key,Math.max(max.get(e.key)||0,e.seq));const k=formatNumber(e.key,e.seq);if(!used.has(k))used.set(k,new Set());used.get(k).add(e.uid);if(e.kind!=='SEED')latest.set(e.uid,e);};
 journal.forEach(reserve);
 for(const r of rows){if(str(r.uid)){if(uids.has(str(r.uid)))throw Error('DUPLICATE_UID_STOP:'+r.uid);uids.set(str(r.uid),r);}}
 // Strict legacy evidence: duplicate historical references reserve once, never block a CT.
 for(const r of rows){const p=parseSource(r.source,str(r.key));if(!p)continue;if(!used.has(p.number))used.set(p.number,new Set());used.get(p.number).add('HISTORICAL:'+p.key);if(p.seq>(max.get(p.key)||0)){const e={...p,uid:'HISTORICAL:'+p.key,kind:'SEED',utc:now,op:operation};entries.push(e);reserve(e);}}
 // All extant W values are observed before allocating any selected row (multi-paste safe).
 const manualNumbers=new Map();
 for(const r of rows){if(!str(r.auto))continue;const p=parseNumber(r.auto,str(r.key));if(!p){warnings.push({row:r.row,code:'MANUAL_INVALID_PRESERVED'});continue;}
 if(str(r.auto)!==p.number)warnings.push({row:r.row,code:'MANUAL_NONCANONICAL_PRESERVED',canonical:p.number});
 if(!manualNumbers.has(p.number))manualNumbers.set(p.number,[]);manualNumbers.get(p.number).push(r);
 let uid=str(r.uid);if(!uid){if(str(r.source)||!complete(r,catalog)){warnings.push({row:r.row,code:'MANUAL_MISSING_UID_STOP'});continue;}uid=uidFactory();identityWrites.push({row:r.row,uid,previous:r});}
 if(used.has(p.number)&&[...used.get(p.number)].some(owner=>owner!==uid)){warnings.push({row:r.row,code:'MANUAL_RESERVED_NUMBER_CONFLICT',number:p.number});}
 const prev=latest.get(uid);if(prev&&prev.key!==r.key){warnings.push({row:r.row,code:'PROJECT_CHANGED_NUMBER_PRESERVED'});continue;}
 if(!journal.concat(entries).some(e=>e.uid===uid&&e.key===p.key&&e.seq===p.seq&&e.kind!=='SEED')){const e={...p,uid,kind:'MANUAL',utc:now,op:operation};entries.push(e);reserve(e);}}
 for(const [number,list] of manualNumbers)if(list.length>1)warnings.push({rows:list.map(r=>r.row),code:'DUPLICATE_MANUAL_PRESERVED',number});
 for(const r of rows){const prev=latest.get(str(r.uid));if(prev&&prev.key!==str(r.key))warnings.push({row:r.row,code:'PROJECT_CHANGED_NUMBER_PRESERVED'});}
 for(const index of selected){const r=rows.find(r=>r.row===index);if(!r)throw Error('SELECTION_OUTSIDE_TABLE');
 if(str(r.auto))continue;
 if(str(r.source)||baseline.has(str(r.uid))){warnings.push({row:r.row,code:'HISTORICAL_NO_BACKFILL'});continue;}
 if(!complete(r,catalog)){warnings.push({row:r.row,code:'INCOMPLETE_NO_NUMBER'});continue;}
 const uid=str(r.uid)||uidFactory();
 if(latest.has(uid)){warnings.push({row:r.row,code:'RESERVED_OR_CLEARED_NO_REISSUE'});continue;}
 if(warnings.some(w=>['MANUAL_MISSING_UID_STOP','DUPLICATE_MANUAL_PRESERVED','MANUAL_RESERVED_NUMBER_CONFLICT'].includes(w.code))){warnings.push({row:r.row,code:'MANUAL_CONFLICT_REVIEW_REQUIRED'});continue;}
 const key=str(r.key),seq=(max.get(key)||0)+1;if(seq>999999)throw Error('SEQUENCE_OVERFLOW');
 const e={uid,key,seq,number:formatNumber(key,seq),kind:'AUTO',utc:now,op:operation};entries.push(e);reserve(e);writes.push({row:r.row,uid,number:e.number,previous:r});
 }
 return {entries,writes,identityWrites,warnings,highWater:Object.fromEntries(max)};
}
module.exports={BUILD,HEADER,JOURNAL_HEADERS,GATE,formatNumber,parseNumber,parseSource,validDate,complete,readJournal,serialize,plan};

globalThis.HcktPurchaseEngine=module.exports;})();
(function(){const module={exports:{}};
'use strict';
// Single-writer, optimistic preflight. Office.js does NOT provide atomic compare-and-swap.
async function startPurchaseNumbering(Excel,E,projectEngine,options){
 if(options?.enabled!==true||typeof options.verifyIdentity!=='function')throw Error('EXPLICIT_BINDING_REQUIRED');
 await options.verifyIdentity();let stopped=false,revision=0,chain=Promise.resolve();let manualPending=false;const subscriptions=[],report=options.report||(()=>{});
 const uuid=()=>{if(!globalThis.crypto?.randomUUID)throw Error('SECURE_UUID_UNAVAILABLE');return globalThis.crypto.randomUUID();};
 const noInputFormulas=range=>{if(range.formulas.some(r=>[1,2,5,6,9,11,17,22].some(c=>String(r[c]||'').startsWith('='))))throw Error('NUMBER_INPUT_FORMULA_STOP');};
 const rangeValues=async(ctx,range)=>{range.load(['values','formulas','rowIndex','columnIndex','rowCount','columnCount']);await ctx.sync();return range;};
 const read=async(ctx,selection)=>{
  await options.verifyIdentity();const sh=ctx.workbook.worksheets.getItem('ĐỀ NGHỊ MUA'),table=sh.tables.getItem('tblCurrentPurchase');
  const body=await rangeValues(ctx,table.getDataBodyRange()),head=await rangeValues(ctx,table.getHeaderRowRange());
  if(head.rowIndex!==1||head.columnIndex!==0||![22,23].includes(head.columnCount)||head.values[0][1]!=='PROJECT KEY'||head.values[0][2]!=='MÃ CT/DA'||head.values[0][17]!=='SourceRowUID'||(head.columnCount===23&&head.values[0][22]!==E.HEADER))throw Error('PURCHASE_SCHEMA_MISMATCH');
  const rows=body.values.map((v,i)=>({row:body.rowIndex+i,key:v[1],display:v[2],source:v[5],content:v[6],requester:v[9],submissionDate:v[11],uid:v[17],auto:v[22]??''}));
  const state=ctx.workbook.worksheets.getItem('_HCKT_STATE');const extent=state.getUsedRange(true);extent.load(['rowIndex','rowCount']);await ctx.sync();const end=extent.rowIndex+extent.rowCount+1;if(end>100000)throw Error('STATE_EXTENT_TOO_LARGE');const journalAddress='AD3:AK'+Math.max(4,end);const base=await rangeValues(ctx,state.tables.getItem('tblHcktBaseline').getRange());
  if(base.values[0][0]!=='SourceRowUID'||base.values[0][10]!=='Origin')throw Error('BASELINE_SCHEMA_MISMATCH');
  const baseline=new Set(base.values.slice(1).filter(r=>r[10]==='HISTORICAL').map(r=>String(r[0])));
  if(!baseline.size)throw Error('BASELINE_MISSING_STOP');
  const jr=await rangeValues(ctx,state.getRange(journalAddress)),header=jr.values[0],journalEmpty=header.every(v=>v==='');
  if(!journalEmpty&&JSON.stringify(header)!==JSON.stringify(E.JOURNAL_HEADERS))throw Error('JOURNAL_SCHEMA_MISMATCH');
  if(journalEmpty&&jr.values.slice(1).some(r=>r.some(v=>v!=='')))throw Error('JOURNAL_OCCUPIED_STOP');
  const data=jr.values.slice(1);let last=-1;data.forEach((r,i)=>{if(r.some(v=>v!==''))last=i;});
  const journal=E.readJournal(data.slice(0,last+1));
  const catTable=ctx.workbook.worksheets.getItem('_A00_SOURCE_QA').tables.getItem('tblA00IntegrationAudit'),ch=await rangeValues(ctx,catTable.getHeaderRowRange()),cb=await rangeValues(ctx,catTable.getDataBodyRange());
  const ci=['Source','ProjectKey','ProjectDisplay','Address'].map(h=>{const hits=ch.values[0].map((v,i)=>v===h?i:-1).filter(i=>i>=0);if(hits.length!==1)throw Error('CATALOG_HEADER');return hits[0];});
  const catalog=projectEngine.catalog(cb.values.map(r=>({source:r[ci[0]],key:r[ci[1]],display:r[ci[2]],address:r[ci[3]]})));
  let selected=[];
  if(selection){const sel=ctx.workbook.getSelectedRange();sel.load(['rowIndex','rowCount']);const active=ctx.workbook.worksheets.getActiveWorksheet();active.load('name');await ctx.sync();if(active.name!=='ĐỀ NGHỊ MUA')throw Error('SELECT_PURCHASE_ROWS_ONLY');if(sel.rowCount>500)throw Error('MAX_500_SELECTED_ROWS');selected=Array.from({length:sel.rowCount},(_,i)=>sel.rowIndex+i);if(selected.some(r=>r<body.rowIndex||r>=body.rowIndex+body.rowCount))throw Error('SELECT_TABLE_DATA_ROWS_ONLY');}
  if(rows.some((r,i)=>[1,2,5,6,9,11,17,22].some(c=>String(body.formulas[i]?.[c]||'').startsWith('='))))throw Error('NUMBER_INPUT_FORMULA_STOP');
  return {sh,table,body,head,state,rows,baseline,journal,journalEmpty,journalNext:4+last,catalog,selected,journalAddress,journalSnapshot:JSON.stringify(jr.values),snapshot:JSON.stringify(body.values)};
 };
 const execute=async(mode)=>Excel.run(async ctx=>{
  if(stopped)throw Error('NUMBERING_STOPPED_RESTART_REQUIRED');const rev=revision,s=await read(ctx,mode!=='observe');
  let previewUid=0;const result=E.plan({...s,operation:mode==='preview'?'preview':uuid(),uidFactory:mode==='preview'?()=>'(UID mới '+(++previewUid)+')':uuid});
  if(mode==='preview'){report({numbering:'PREVIEW',gate:E.GATE,...result});return result;}
  const pending=result.entries,identities=result.writes.filter(w=>!String(w.previous.uid??'').trim()).concat(result.identityWrites);if(!pending.length&&!result.writes.length){report({numbering:'NO_CHANGE',warnings:result.warnings});return result;}
  if(s.journalNext+pending.length>99999)throw Error('JOURNAL_CAPACITY_STOP');
  // Compare full input + journal again. Sort/edit races invalidate the operation; never retry automatically.
  const bodyCheck=await rangeValues(ctx,s.table.getDataBodyRange());noInputFormulas(bodyCheck);const journalCheck=await rangeValues(ctx,s.state.getRange(s.journalAddress));
  if(bodyCheck.rowIndex!==s.body.rowIndex||JSON.stringify(bodyCheck.values)!==s.snapshot||JSON.stringify(journalCheck.values)!==s.journalSnapshot||stopped||revision!==rev)throw Error('CONCURRENT_CHANGE_REVIEW_AND_RETRY');
  await options.verifyIdentity();
  // Establish immutable row identity before reserving: uncertain reservations cannot acquire a new UID on retry.
  for(const w of identities)s.sh.getCell(w.row,17).values=[[w.uid]];
  if(identities.length){
   await ctx.sync();const identified=await rangeValues(ctx,s.table.getDataBodyRange());noInputFormulas(identified);
   const expected=JSON.parse(s.snapshot);for(const w of identities)expected[w.row-s.body.rowIndex][17]=w.uid;
   if(JSON.stringify(identified.values)!==JSON.stringify(expected)||stopped||revision!==rev)throw Error('UID_PREFLIGHT_FAILED_STOP');
   s.snapshot=JSON.stringify(expected);
  }
  if(s.journalEmpty)s.state.getRange('AD3:AK3').values=[E.JOURNAL_HEADERS];
  if(pending.length)s.state.getRangeByIndexes(s.journalNext,29,pending.length,8).values=pending.map(E.serialize);
  await ctx.sync(); // Durable reservation first. Any failure stops; no blind replay.
  const persisted=await rangeValues(ctx,s.state.getRangeByIndexes(s.journalNext,29,pending.length||1,8));
  if(pending.length&&JSON.stringify(persisted.values)!==JSON.stringify(pending.map(E.serialize)))throw Error('JOURNAL_READBACK_FAILED_STOP');
  if(stopped||revision!==rev)throw Error('RESERVED_ONLY_CONCURRENT_CHANGE_STOP');
  if(result.writes.length){
   const fresh=await rangeValues(ctx,s.table.getDataBodyRange());noInputFormulas(fresh);if(JSON.stringify(fresh.values)!==s.snapshot)throw Error('RESERVED_ONLY_INPUT_CHANGED_STOP');
   if(s.head.columnCount===22){const outside=await rangeValues(ctx,s.sh.getRangeByIndexes(1,22,s.body.rowCount+1,1));if(outside.values.some(r=>r.some(v=>v!==''))||outside.formulas.some(r=>r.some(v=>String(v||'').startsWith('='))))throw Error('W_OCCUPIED_STOP');s.table.columns.add(null,null,E.HEADER);await ctx.sync();}
   // Column creation is another await window: recheck row identities and formulas afterward.
   const finalCheck=await rangeValues(ctx,s.table.getDataBodyRange());noInputFormulas(finalCheck);
   const expectedFinal=JSON.parse(s.snapshot).map(r=>r.length===22?r.concat(''):r);
   if(JSON.stringify(finalCheck.values)!==JSON.stringify(expectedFinal)||stopped||revision!==rev)throw Error('RESERVED_ONLY_POST_SCHEMA_CHANGE_STOP');
   for(const w of result.writes)s.sh.getCell(w.row,22).values=[[w.number]];
   await ctx.sync();
   const verified=await rangeValues(ctx,s.table.getDataBodyRange());for(const w of result.writes){const row=verified.values[w.row-verified.rowIndex];if(row?.[17]!==w.uid||row?.[22]!==w.number)throw Error('NUMBER_READBACK_FAILED_STOP');}
  }
  report({numbering:'COMMITTED',count:result.writes.length,warnings:result.warnings});return result;
 });
 const queue=mode=>{const next=chain.then(()=>execute(mode));chain=next.catch(e=>{stopped=true;report({stopped:true,numbering:true,error:String(e)});});return next;};
 try{await Excel.run(async ctx=>{for(const name of ['ĐỀ NGHỊ MUA','_HCKT_STATE','_A00_SOURCE_QA']){const sh=ctx.workbook.worksheets.getItem(name);subscriptions.push(sh.onChanged.add(e=>{if(e.triggerSource==='ThisLocalAddin')return;++revision;if(e.source==='Remote'){stopped=true;report({stopped:true,numbering:true,error:'REMOTE_EDIT_SINGLE_WRITER_STOP'});return;}if(name==='ĐỀ NGHỊ MUA'&&e.changeType==='RangeEdited'){const a=String(e.address).split('!').pop().replace(/\$/g,'').match(/^([A-Z]+)\d+(?::([A-Z]+)\d+)?$/i);const column=x=>[...x.toUpperCase()].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);if(a&&column(a[1])<=23&&column(a[2]||a[1])>=23){if(manualPending){stopped=true;report({stopped:true,error:'RAPID_MANUAL_EDIT_REVIEW_REQUIRED'});return;}manualPending=true;queue('observe').catch(()=>{}).finally(()=>{manualPending=false;});}}}));subscriptions.push(sh.onRowSorted.add(()=>{++revision;}));}await ctx.sync();});}catch(e){stopped=true;for(const s of subscriptions){try{s.remove();await s.context.sync();}catch(_){}}throw e;}
 // Startup never seeds, creates a column, persists a setting, or issues numbers.
 return {preview:()=>queue('preview'),allocate:()=>queue('allocate'),stop:async()=>{stopped=true;++revision;await chain;for(const s of subscriptions){s.remove();await s.context.sync();}}};
}
if(typeof module!=='undefined')module.exports={startPurchaseNumbering};

globalThis.HcktPurchaseHost=module.exports;})();
