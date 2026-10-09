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
