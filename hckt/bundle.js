'use strict';const modules={},cache={};
modules["app/manual-request-rules.cjs"]=function(require,module,exports){
'use strict';
const clean=x=>String(x??'').trim();
const STATUS=['\u0110\u00e3 duy\u1ec7t','YC Gi\u1ea3i tr\u00ecnh','Ch\u01b0a tr\u00ecnh','Ch\u01b0a duy\u1ec7t'];
function serialVN(iso){const ms=Date.parse(iso);if(!Number.isFinite(ms))throw Error('INVALID_CLOCK');return Math.floor((ms+7*3600000)/86400000)+25569}
function project({row,changed,projects,approvers,profile}){
 const patch={},errors=[],warnings=[];
 if(changed.includes('B') || changed.includes('C')){
  const matches=projects.filter(p=>changed.includes('B')?p.key===clean(row.B):p.display===clean(row.C));
  if(matches.length!==1)errors.push('PROJECT_NOT_UNIQUE');
  else if(changed.includes('B') && changed.includes('C') && matches[0].display!==clean(row.C))errors.push('PROJECT_B_C_CONFLICT');
  else{patch.B=matches[0].key;patch.C=matches[0].display;if(clean(matches[0].address))patch[profile.address]=matches[0].address;else warnings.push('ADDRESS_MISSING_KEEP_EXISTING')}
 }
 if(changed.includes(profile.approver) && clean(row[profile.approver]) && !approvers.includes(clean(row[profile.approver])))errors.push('APPROVER_NOT_IN_THREE_BLD');
 return {patch:errors.length?{}:patch,errors,warnings};
}
function date({row,old,changed,profile,nowISO}){
 const patch={},errors=[],status=clean(row[profile.status]);
 if(!changed.includes(profile.status))return {patch,errors};
 if(!STATUS.includes(status)){errors.push('INVALID_STATUS');return {patch,errors}}
 if(!old){errors.push('STATUS_BASELINE_REQUIRED');return {patch,errors}}
 if(status===clean(old[profile.status]))return {patch,errors};
 if(status===STATUS[0])patch[profile.date]=serialVN(nowISO);
 else if(status===STATUS[3])patch[profile.date]='';
 // Other states preserve the entered date; date-only edits never cause writes.
 return {patch,errors};
}
module.exports={project,date,serialVN,STATUS};

};
modules["app/three-sheet-engine.cjs"]=function(require,module,exports){
'use strict';
const manual=require('./manual-request-rules.cjs'),xuat=require('../rules/export-rules.cjs');
const profiles={
 'ĐỀ NGHỊ MUA':{table:'tblCurrentPurchase',columns:22,uid:'R',status:'M',date:'N',approver:'O',address:'P',write:['B','C','P','N']},
 'ĐỀ NGHỊ SX':{table:'tblCurrentProduction',columns:26,uid:'W',status:'L',date:'M',approver:'N',address:'Q',write:['B','C','Q','M']},
 'ĐỀ NGHỊ XUẤT':{table:'tblCurrentExport',columns:26,uid:'W',status:'L',date:'M',approver:'N',address:'O',write:['B','C','D','O','M']}
};
function plan(sheet,snapshot,index,changed,nowISO){
 const profile=profiles[sheet],row=snapshot.rows[index];if(!profile)throw Error('SHEET_NOT_APPROVED');if(!row || !String(row.uid??'').trim())throw Error('ROW_UID_REQUIRED');
 if(snapshot.rows.filter(r=>r.uid===row.uid).length!==1)throw Error('DUPLICATE_UID');
 let result=profile.table==='tblCurrentExport'?xuat.proposal({row,changed,projects:snapshot.projects,contracts:snapshot.contracts,technicalNames:snapshot.technicalNames,approvers:snapshot.approvers}):manual.project({row,changed,projects:snapshot.projects,approvers:snapshot.approvers,profile});
 const dated=manual.date({row,old:snapshot.baseline.get(row.uid),changed,profile,nowISO});result.errors.push(...dated.errors);
 if(!result.errors.length)Object.assign(result.patch,dated.patch);else result.patch={};
 // F is manually controlled: no allocation, restore, validation against old F,
 // or optional flag can re-enable a write. Manual date edits are also allowed.
 for(const c of Object.keys(result.patch))if(c==='F' || !profile.write.includes(c))throw Error('WRITE_SCOPE_DENIED');
 return {...result,uid:row.uid,sheet,protectedRestore:[]};
}
module.exports={profiles,plan};

};
modules["app/three-sheet-controller.cjs"]=function(require,module,exports){
'use strict';
const {profiles,plan}=require('./three-sheet-engine.cjs');
function planBatch(sheet,snapshot,nowISO){
 const simulation={...snapshot,rows:snapshot.rows.map(r=>({...r})),baseline:new Map([...snapshot.baseline].map(([k,r])=>[k,{...r}]))};
 const grouped=new Map();for(const r of snapshot.changedRows){const cols=grouped.get(r.index)||new Set();for(const c of r.columns)cols.add(c);grouped.set(r.index,cols)}
 const plans=[];
 for(const [index,cols] of [...grouped].sort((a,b)=>a[0]-b[0])){
  const p=plan(sheet,simulation,index,[...cols],nowISO);p.index=index;plans.push(p);Object.assign(simulation.rows[index],p.patch);
 }
 return plans;
}
// Integration contract only: production Office/identity/storage adapters must pass
// this contract and real Excel QA. It does not assert Excel atomic transactions.
function controller(host){
 let ready=false,halted=false,queue=Promise.resolve(),handlers=[],generation=0,pending=[],timer=null,draining=false;
 const fail=e=>{ready=false;halted=true;host.invalidateSession?.();clearTimeout(timer);for(const x of pending)x.reject(e);pending=[];host.status('AUTOMATION_NOT_READY',e.message)};
 async function start(){
  if(ready || handlers.length)throw Error('ALREADY_REGISTERED');
  let startupAttempted=false;try{await host.guardQA();await host.claimSession();await host.bootstrapBounded(Object.keys(profiles));
   for(const sheet of Object.keys(profiles))handlers.push(await host.addHandler(sheet,event=>changed({...event,sheet})));
   startupAttempted=true;await host.setStartup(true);ready=true;halted=false;host.status('READY');
  }catch(e){fail(e);let cleanupFailed=false;if(startupAttempted)try{await (host.resetStartupVerified?host.resetStartupVerified():host.setStartup(false))}catch(_){cleanupFailed=true}const retained=[];for(const h of handlers)try{await host.removeHandler(h)}catch(_){retained.push(h);cleanupFailed=true}handlers=retained;try{await host.disposeAllHandlers?.()}catch(_){cleanupFailed=true}try{await host.releaseSession()}catch(_){cleanupFailed=true}if(cleanupFailed)host.status('REMOVAL_NOT_VERIFIED');throw e}
 }
 async function process(events){
  let batch=events,retries=0;
  while(true){
   await host.assertSession();
   const e={...batch[0].event,addresses:batch.map(x=>x.event.address).filter(Boolean)};
   if(e.kind!=='RangeEdited'){await host.reconcileStructure(e);return}
   const snapshot=await host.readChangedRows(e),plans=planBatch(e.sheet,snapshot,host.nowISO());
   await host.assertSession();if(halted)throw Error('STOPPED_BEFORE_COMMIT');
   const newer=pending.filter(x=>x.event.sheet===e.sheet && x.event.kind==='RangeEdited');
   if(newer.length){if(++retries>3)throw Error('EDIT_RETRY_LIMIT_REVIEW_REQUIRED');pending=pending.filter(x=>!newer.includes(x));batch.push(...newer);continue}
   try{await host.compareAndCommit(snapshot.revision,plans,{sheet:e.sheet,allowed:profiles[e.sheet].write})}
   catch(error){if(error.message==='LATEST_STATE_CHANGED' && ++retries<=3)continue;throw error}
   host.report(plans);return;
  }
 }
 async function drain(){
  if(draining)return;draining=true;
  try{while(pending.length && !halted){const first=pending.shift(),batch=[first];if(first.event.kind==='RangeEdited'){const matches=pending.filter(x=>x.event.sheet===first.event.sheet && x.event.kind==='RangeEdited');pending=pending.filter(x=>!matches.includes(x));batch.push(...matches)}
   try{await process(batch);for(const x of batch)x.resolve()}
   catch(error){fail(error);for(const x of batch)x.reject(error);for(const x of pending)x.reject(error);pending=[]}
  }}finally{draining=false}
 }
 async function changed(e){
  if(!ready || halted || e.internal)return;
  if(e.source==='Remote'){fail(Error('SECOND_SESSION_EDIT'));return}
  if(e.source!=='Local' || !profiles[e.sheet])return;
  const promise=new Promise((resolve,reject)=>pending.push({event:e,resolve,reject}));
  if(!draining){clearTimeout(timer);timer=setTimeout(()=>{queue=drain()},30)}
  return promise;
 }
 async function uninstall(){
  ready=false;halted=true;host.invalidateSession?.();clearTimeout(timer);const errors=[];
  for(const x of pending)x.resolve();if(pending.length)host.status('STOPPED_PENDING_EDITS_REQUIRE_REVIEW');pending=[];
  try{await (host.resetStartupVerified?host.resetStartupVerified():host.setStartup(false))}catch(e){errors.push(e)}
  const retained=[];for(const h of handlers)try{await host.removeHandler(h)}catch(e){retained.push(h);errors.push(e)}handlers=retained;try{await host.disposeAllHandlers?.()}catch(e){errors.push(e)}
  await queue.catch(()=>{});try{await host.releaseSession()}catch(e){errors.push(e)}
  if(errors.length){host.status('REMOVAL_NOT_VERIFIED',errors.map(x=>x.message).join('; '));throw Error('REMOVAL_NOT_VERIFIED')}
  host.status('REMOVED');
  // No business-data deletion or whole-state/named-range cleanup.
 }
 return {start,changed,uninstall,get ready(){return ready},get halted(){return halted}};
}
module.exports={controller,planBatch};

};
modules["app/office-three-sheet-host.cjs"]=function(require,module,exports){
'use strict';
const {profiles}=require('./three-sheet-engine.cjs');
const clean=x=>String(x??'').trim(),column=s=>[...s].reduce((v,c)=>v*26+c.charCodeAt(0)-64,0)-1;
function createOfficeHost(Excel,Office,cfg,ui){
 const handlerGroups=new Set(),cache=new Map(),snapshots=new Map(),metrics={fullReads:0,scopedReads:0,cellsRead:0};let sessionVersion=0,editVersion=0,active=false,dependencyAdded=false,state=[],register=[],catalog=[];
 const token=()=>({session:sessionVersion,edit:editVersion});
 function check(t,edits=false){if(!active || t.session!==sessionVersion)throw Error('LOCAL_SESSION_STOPPED');if(edits && t.edit!==editVersion)throw Error('LATEST_STATE_CHANGED')}
 const clone=x=>x.map(r=>r.slice()),yes=v=>v===true || v===1 || String(v).toUpperCase()==='TRUE';
 function decode(sheet,values){const p=profiles[sheet];return values.map((a,index)=>{const r={index,uid:clean(a[column(p.uid)])};a.forEach((v,i)=>r[String.fromCharCode(65+i)]=v);return r})}
 function baseline(sheet){const p=profiles[sheet];return new Map(state.filter(r=>clean(r[0])).map(r=>[clean(r[0]),{uid:r[0],B:r[1],C:r[2],[p.status]:r[3],[p.date]:r[4],[p.approver]:r[5],F:r[7],[p.address]:r[8],origin:r[10],E:r[11]}]))}
 function serialize(sheet,row,old){const p=profiles[sheet];return [row.uid,row.B||'',row.C||'',row[p.status]||'',row[p.date]??'',row[p.approver]||'',old?.[6]||'',row.F??'',row[p.address]||'',row.index+(cache.get(sheet)?.rowIndex??2)+1,old?.[10]||'HCKT_EVENT_QA',row.E||'']}
 function dictionaries(){
  const ct=catalog.filter(r=>r[0]==='CTDA'),keys=[...new Set(ct.map(r=>clean(r[2])).filter(Boolean))];
  const projects=keys.map(key=>{const r=ct.filter(r=>clean(r[2])===key),names=[...new Set(r.map(x=>clean(x[3])).filter(Boolean))],addresses=[...new Set(r.map(x=>clean(x[5])).filter(Boolean))];if(names.length!==1)throw Error('PROJECT_DISPLAY_NOT_UNIQUE');return {key,display:names[0],address:addresses.length===1?addresses[0]:''}});
  return {projects,contracts:ct.map(r=>({key:clean(r[2]),contract:clean(r[4])})),technicalNames:catalog.filter(r=>r[0]==='NSKT' && r[14]==='Phòng Kỹ thuật').map(r=>r[12]),approvers:catalog.filter(r=>r[0]==='NSKT' && r[14]==='Ban lãnh đạo' && yes(r[16])).map(r=>r[12])};
 }
 async function fullSheet(context,sheet,t){const p=profiles[sheet],range=context.workbook.worksheets.getItem(sheet).tables.getItem(p.table).getDataBodyRange();range.load('rowIndex,rowCount,columnCount');await context.sync();if(t)check(t,true);if(range.rowCount>10000 || range.columnCount!==p.columns)throw Error('SHEET_SCHEMA_OR_BOUNDS');range.load('values');await context.sync();if(t)check(t,true);const values=clone(range.values);cache.set(sheet,{values,rowIndex:range.rowIndex});metrics.fullReads++;metrics.cellsRead+=values.length*p.columns;return values}
 const host={metrics,status:(s,e)=>ui.status(s,e),report:p=>ui.report(p),
  async guardQA(){if(!cfg.enabled || cfg.singleSessionMode!==true || !cfg.approvedQAUrl || cfg.baselineTable!=='tblHcktBaseline' || cfg.registerTable!=='tblHcktRegister')throw Error('QA_CONFIGURATION_DISABLED');if(!Office.context.requirements.isSetSupported('ExcelApi','1.14') || !Office.context.requirements.isSetSupported('SharedRuntime','1.1'))throw Error('OFFICE_RUNTIME_UNSUPPORTED');const url=await new Promise((resolve,reject)=>Office.context.document.getFilePropertiesAsync(r=>r.status===Office.AsyncResultStatus.Succeeded?resolve(r.value.url):reject(Error('FILE_URL_UNAVAILABLE'))));if(url!==cfg.approvedQAUrl || !(decodeURIComponent(url).includes('/History/HCKT_EVENTS_QA_') || (cfg.qaWorkbookGuid && decodeURIComponent(url).toUpperCase().includes(cfg.qaWorkbookGuid.toUpperCase()))) || url.toUpperCase().includes(cfg.originalWorkbookGuid))throw Error('QA_DOCUMENT_ONLY')},
  async claimSession(){if(active)throw Error('LOCAL_SESSION_ALREADY_ACTIVE');if(handlerGroups.size)throw Error('HANDLER_CLEANUP_REQUIRED');sessionVersion++;active=true},invalidateSession(){active=false;sessionVersion++;snapshots.clear()},async assertSession(){if(!active)throw Error('LOCAL_SESSION_STOPPED')},async releaseSession(){host.invalidateSession()},
  async bootstrapBounded(sheets){await Excel.run(async context=>{
   const br=context.workbook.tables.getItem(cfg.baselineTable).getDataBodyRange(),rr=context.workbook.tables.getItem(cfg.registerTable).getDataBodyRange(),cr=context.workbook.tables.getItem('tblA00IntegrationAudit').getDataBodyRange();for(const r of [br,rr,cr])r.load('rowCount,columnCount');await context.sync();if(br.rowCount>30000 || rr.rowCount>30000 || cr.rowCount>5000 || br.columnCount!==12 || rr.columnCount!==9 || cr.columnCount!==20)throw Error('STATE_BOUNDS');for(const r of [br,rr,cr])r.load('values');await context.sync();state=clone(br.values);register=clone(rr.values);catalog=clone(cr.values);metrics.cellsRead+=state.length*12+register.length*9+catalog.length*20;const allUIDs=new Set(),pending=[];
   for(const sheet of sheets){const rows=decode(sheet,await fullSheet(context,sheet));for(const row of rows){if(!row.uid || allUIDs.has(row.uid))throw Error('EXISTING_UID_MISSING_OR_DUPLICATE');allUIDs.add(row.uid);if(!state.some(r=>r[0]===row.uid)){if(!cfg.allowInitialStateSeed)throw Error('QA_BASELINE_SEED_NOT_ENABLED');pending.push(serialize(sheet,row,['','','','','','','','','','','HistoricalQASeed','']))}}}
   if(pending.length){context.workbook.tables.getItem(cfg.baselineTable).rows.add(null,pending);await context.sync();state.push(...pending)}
   if(dictionaries().approvers.length!==3)throw Error('BLD_COUNT_NOT_THREE');
  })},
  async addHandler(sheet,callback){const group={results:[],removed:new Set(),dependencies:false};handlerGroups.add(group);
   try{return await Excel.run(async context=>{
    group.results.push(context.workbook.worksheets.getItem(sheet).onChanged.add(e=>{if(e.triggerSource==='ThisLocalAddin')return;editVersion++;callback({address:e.address,kind:e.changeType,source:e.source,internal:false}).catch(error=>ui.status('AUTOMATION_NOT_READY',error.message))}));
    if(!dependencyAdded){for(const name of ['_HCKT_STATE','_HCKT_LOOKUP','_A00_SOURCE_QA']){group.results.push(context.workbook.worksheets.getItem(name).onChanged.add(e=>{if(e.triggerSource==='ThisLocalAddin')return;host.invalidateSession();editVersion++;cache.clear();callback({source:'Remote',kind:'RangeEdited',address:e.address}).catch(()=>{});ui.status('DEPENDENCY_CHANGED_REVIEW_REQUIRED')}));group.dependencies=true;dependencyAdded=true}}
    await context.sync();return group;
   })}catch(error){try{await host.removeHandler(group)}catch(cleanup){ui.status('REMOVAL_NOT_VERIFIED',cleanup.message)}throw error}
  },
  async removeHandler(group){const errors=[];for(const r of group.results){if(group.removed.has(r))continue;try{await Excel.run(r.context,async context=>{r.remove();await context.sync()});group.removed.add(r)}catch(e){errors.push(e)}}
   if(!errors.length)handlerGroups.delete(group);dependencyAdded=[...handlerGroups].some(g=>g.dependencies && g.results.some(r=>!g.removed.has(r)));
   if(errors.length)throw Error('HANDLER_REMOVAL_NOT_VERIFIED');
  },
  async disposeAllHandlers(){const errors=[];for(const group of [...handlerGroups])try{await host.removeHandler(group)}catch(e){errors.push(e)}if(errors.length || handlerGroups.size)throw Error('HANDLER_REMOVAL_NOT_VERIFIED')},
  async setStartup(enabled){const expected=enabled?Office.StartupBehavior.load:Office.StartupBehavior.none;await Office.addin.setStartupBehavior(expected);if(await Office.addin.getStartupBehavior()!==expected)throw Error('STARTUP_VERIFICATION_FAILED')},
  async resetStartupVerified(){try{await Office.addin.setStartupBehavior(Office.StartupBehavior.none)}catch(_){}if(await Office.addin.getStartupBehavior()!==Office.StartupBehavior.none)throw Error('STARTUP_REMOVAL_NOT_VERIFIED')},
  nowISO:()=>new Date().toISOString(),
  async readChangedRows(e){const t=token();check(t);return Excel.run(async context=>{
   const p=profiles[e.sheet],sh=context.workbook.worksheets.getItem(e.sheet),range=sh.tables.getItem(p.table).getDataBodyRange();range.load('rowIndex,rowCount,columnCount');
   const edits=[...new Set(e.addresses?.length?e.addresses:[e.address])].map(address=>sh.getRange(address));for(const edit of edits)edit.load('rowIndex,rowCount,columnIndex,columnCount');await context.sync();check(t,true);let old=cache.get(e.sheet);
   if(range.rowIndex===old.rowIndex && range.rowCount>old.values.length){const first=old.rowIndex+old.values.length;await host.reconcileStructure({...e,kind:'RowInserted',address:'A'+(first+1)+':'+String.fromCharCode(64+p.columns)+(range.rowIndex+range.rowCount),prefixUIDs:old.values.map(v=>clean(v[column(p.uid)]))});check(t,true);old=cache.get(e.sheet)}
   if(range.rowIndex!==old.rowIndex || range.rowCount!==old.values.length || range.columnCount!==p.columns)throw Error('STRUCTURE_CHANGED_RECONCILE_REQUIRED');
   const changed=new Map(),intervals=[];for(const edit of edits){const lo=Math.max(edit.rowIndex,range.rowIndex),hi=Math.min(edit.rowIndex+edit.rowCount,range.rowIndex+range.rowCount);if(hi<=lo)continue;intervals.push({lo,hi});for(let r=lo;r<hi;r++){const cols=changed.get(r-range.rowIndex)||new Set();for(let c=edit.columnIndex;c<Math.min(edit.columnIndex+edit.columnCount,p.columns);c++)cols.add(String.fromCharCode(65+c));changed.set(r-range.rowIndex,cols)}}
   if(!changed.size)throw Error('EDIT_OUTSIDE_TABLE');const merged=[];for(const x of intervals.sort((a,b)=>a.lo-b.lo)){const last=merged.at(-1);if(last && x.lo<=last.hi)last.hi=Math.max(last.hi,x.hi);else merged.push({...x})}
   const ranges=merged.map(x=>({...x,range:sh.getRangeByIndexes(x.lo,0,x.hi-x.lo,p.columns)}));for(const x of ranges)x.range.load('values');await context.sync();check(t,true);const values=clone(old.values),raw=[];
   for(const x of ranges){x.range.values.forEach((v,i)=>{const index=x.lo-range.rowIndex+i;if(clean(v[column(p.uid)])!==clean(values[index][column(p.uid)]))throw Error('UID_ORDER_CHANGED_RECONCILE_REQUIRED');values[index]=v.slice()});raw.push({lo:x.lo,hi:x.hi,values:clone(x.range.values)});metrics.cellsRead+=(x.hi-x.lo)*p.columns}
   metrics.scopedReads++;for(const cols of changed.values())if(cols.has(p.uid))throw Error('UID_EDIT_REQUIRES_REVIEW');const revision=crypto.randomUUID();snapshots.set(revision,{token:t,e,raw,values,rowIndex:range.rowIndex,rowCount:range.rowCount});
   return {...dictionaries(),rows:decode(e.sheet,values),baseline:baseline(e.sheet),registeredNewUIDs:new Set(register.filter(r=>r[8]==='HCKT_EVENT_QA' && ['Registered','Issued'].includes(r[7])).map(r=>r[1])),historicalUIDs:new Set(state.filter(r=>r[10]!=='HCKT_EVENT_QA').map(r=>r[0])),revision,changedRows:[...changed].map(([index,cols])=>({index,columns:[...cols]}))};
  })},
  async compareAndCommit(revision,plans,scope){const snap=snapshots.get(revision);if(!snap || scope.sheet!==snap.e.sheet)throw Error('SNAPSHOT_REQUIRED');check(snap.token,true);await Excel.run(async context=>{
   const p=profiles[scope.sheet],sh=context.workbook.worksheets.getItem(scope.sheet),range=sh.tables.getItem(p.table).getDataBodyRange(),ranges=snap.raw.map(x=>({...x,range:sh.getRangeByIndexes(x.lo,0,x.hi-x.lo,p.columns)}));range.load('rowIndex,rowCount');for(const x of ranges)x.range.load('values');await context.sync();check(snap.token,true);metrics.cellsRead+=ranges.reduce((n,x)=>n+(x.hi-x.lo)*p.columns,0);if(range.rowIndex!==snap.rowIndex || range.rowCount!==snap.rowCount || ranges.some(x=>JSON.stringify(x.range.values)!==JSON.stringify(x.values)))throw Error('LATEST_STATE_CHANGED');if(!active)throw Error('LOCAL_SESSION_STOPPED');const next=clone(snap.values),nextState=clone(state),nextRegister=clone(register),bt=context.workbook.tables.getItem(cfg.baselineTable),br=bt.getDataBodyRange(),rt=context.workbook.tables.getItem(cfg.registerTable),rr=rt.getDataBodyRange();
   const stateChecks=[],registerChecks=[];br.load('rowCount');rr.load('rowCount');for(const plan of plans){const bi=state.findIndex(r=>r[0]===plan.uid);if(bi>=0){const range=br.getRow(bi);range.load('values');stateChecks.push({bi,range})}}
   await context.sync();check(snap.token,true);metrics.cellsRead+=stateChecks.length*12+registerChecks.length*9;if(!active)throw Error('LOCAL_SESSION_STOPPED');if(br.rowCount!==state.length || rr.rowCount!==register.length || stateChecks.some(x=>JSON.stringify(x.range.values[0])!==JSON.stringify(state[x.bi])) || registerChecks.some(x=>JSON.stringify(x.range.values[0])!==JSON.stringify(register[x.ri])))throw Error('LATEST_STATE_CHANGED');
   // Last read includes business rows AND metadata. No await may separate its
   // validation/session check from write staging and dispatch.
   range.load('rowIndex,rowCount');for(const x of ranges)x.range.load('values');br.load('rowCount');rr.load('rowCount');for(const x of [...stateChecks,...registerChecks])x.range.load('values');
   await context.sync();check(snap.token,true);metrics.cellsRead+=ranges.reduce((n,x)=>n+(x.hi-x.lo)*p.columns,0)+stateChecks.length*12+registerChecks.length*9;
   if(range.rowIndex!==snap.rowIndex || range.rowCount!==snap.rowCount || ranges.some(x=>JSON.stringify(x.range.values)!==JSON.stringify(x.values)) || br.rowCount!==state.length || rr.rowCount!==register.length || stateChecks.some(x=>JSON.stringify(x.range.values[0])!==JSON.stringify(state[x.bi])) || registerChecks.some(x=>JSON.stringify(x.range.values[0])!==JSON.stringify(register[x.ri])))throw Error('LATEST_STATE_CHANGED');
   for(const plan of plans){if(plan.uid!==decode(scope.sheet,[next[plan.index]])[0].uid || state.findIndex(r=>r[0]===plan.uid)<0)throw Error('BASELINE_OR_UID_MISMATCH');for(const c of Object.keys(plan.patch))if(!scope.allowed.includes(c) || !p.write.includes(c))throw Error('WRITE_SCOPE_DENIED')}
   check(snap.token,true);for(const plan of plans){const row=decode(scope.sheet,[next[plan.index]])[0];row.index=plan.index;for(const [c,value] of Object.entries(plan.patch)){if(!scope.allowed.includes(c) || !p.write.includes(c))throw Error('WRITE_SCOPE_DENIED');sh.getCell(snap.rowIndex+plan.index,column(c)).values=[[value]];next[plan.index][column(c)]=value;row[c]=value}
    if(plan.errors?.length)continue;
    const bi=nextState.findIndex(r=>r[0]===row.uid),record=serialize(scope.sheet,row,nextState[bi]);if(bi<0)throw Error('BASELINE_REQUIRED');br.getRow(bi).values=[record];nextState[bi]=record;
   }
   check(snap.token,true);await context.sync();check(snap.token,true);cache.set(scope.sheet,{values:next,rowIndex:snap.rowIndex});state=nextState;register=nextRegister;snapshots.delete(revision);
  })},
  async reconcileStructure(e){const t=token();check(t);await Excel.run(async context=>{const p=profiles[e.sheet],sh=context.workbook.worksheets.getItem(e.sheet),table=sh.tables.getItem(p.table),values=await fullSheet(context,e.sheet,t),rows=decode(e.sheet,values),edit=sh.getRange(e.address);edit.load('rowIndex,rowCount');await context.sync();check(t,true);const body=cache.get(e.sheet),seen=new Set();if(e.prefixUIDs && e.prefixUIDs.some((uid,i)=>rows[i]?.uid!==uid))throw Error('AUTOEXTEND_UID_PREFIX_CHANGED');for(const row of rows)if(row.uid){if(seen.has(row.uid))throw Error('COPIED_UID');seen.add(row.uid)}
   // Structural checks read all three UID columns from the workbook, not a
   // possibly stale per-sheet cache. Existing IDs are never rewritten.
   const otherBodies=Object.entries(profiles).filter(([name])=>name!==e.sheet).map(([name,profile])=>({profile,body:context.workbook.worksheets.getItem(name).tables.getItem(profile.table).getDataBodyRange()}));
   for(const x of otherBodies)x.body.load('rowCount,columnCount');await context.sync();check(t,true);for(const x of otherBodies){if(x.body.rowCount>10000 || x.body.columnCount!==x.profile.columns)throw Error('SHEET_SCHEMA_OR_BOUNDS');x.uids=x.body.getColumn(column(x.profile.uid));x.uids.load('values')}
   const finalBody=table.getDataBodyRange();finalBody.load('rowIndex,rowCount,values');await context.sync();check(t,true);
   for(const x of otherBodies){metrics.cellsRead+=x.body.rowCount;for(const r of x.uids.values){const uid=clean(r[0]);if(uid){if(seen.has(uid))throw Error('COPIED_UID_WORKBOOK');seen.add(uid)}}}
if(finalBody.rowIndex!==body.rowIndex || finalBody.rowCount!==values.length || JSON.stringify(finalBody.values)!==JSON.stringify(values))throw Error('LATEST_STATE_CHANGED');
   check(t,true);const pending=[],registrations=[];for(const row of rows)if(!row.uid){const absolute=body.rowIndex+row.index;if(e.kind!=='RowInserted' || absolute<edit.rowIndex || absolute>=edit.rowIndex+edit.rowCount)throw Error('UNVERIFIED_NEW_ROW');row.uid=crypto.randomUUID();if(seen.has(row.uid) || state.some(r=>clean(r[0])===row.uid))throw Error('GENERATED_UID_COLLISION');seen.add(row.uid);sh.getCell(absolute,column(p.uid)).values=[[row.uid]];body.values[row.index][column(p.uid)]=row.uid;const initial={...row,[p.status]:'',[p.date]:''};const requestUID=crypto.randomUUID(),record=serialize(e.sheet,initial);record[6]=requestUID;pending.push(record);registrations.push([requestUID,row.uid,row.B||'',row.E||'','','',new Date().toISOString(),'Registered','HCKT_EVENT_QA'])}
   if(pending.length){check(t,true);context.workbook.tables.getItem(cfg.baselineTable).rows.add(null,pending);context.workbook.tables.getItem(cfg.registerTable).rows.add(null,registrations);check(t,true);await context.sync();check(t,true);state.push(...pending);register.push(...registrations)}
  })}
 };
 return host;
}
module.exports={createOfficeHost,column};

};
modules["rules/export-rules.cjs"]=function(require,module,exports){
'use strict';
const clean=x=>String(x??'').trim();
function proposal({row,changed,projects,contracts,technicalNames,approvers}){
 const patch={},errors=[],warnings=[];
 if(changed.includes('B') || changed.includes('C')){
  const matches=projects.filter(p=>changed.includes('B')?p.key===clean(row.B):p.display===clean(row.C));
  if(matches.length!==1)errors.push('PROJECT_NOT_UNIQUE');
  else if(changed.includes('B') && changed.includes('C') && matches[0].display!==clean(row.C))errors.push('PROJECT_B_C_CONFLICT');
  else{
   const p=matches[0];patch.B=p.key;patch.C=p.display;
   if(clean(p.address))patch.O=p.address;else warnings.push('ADDRESS_MISSING_KEEP_EXISTING');
   const options=[...new Set(contracts.filter(c=>c.key===p.key).map(c=>clean(c.contract)).filter(Boolean))];
   if(options.length!==1)warnings.push('CONTRACT_SELECTION_REQUIRED_KEEP_EXISTING');
   else if(!clean(row.D))patch.D=options[0];
   else if(clean(row.D)!==options[0])warnings.push('EXISTING_CONTRACT_REVIEW_REQUIRED_KEEP_EXISTING');
  }
 }
 if(changed.includes('J') && clean(row.J) && !technicalNames.includes(clean(row.J)))errors.push('PREPARER_NOT_TECHNICAL');
 if(changed.includes('N') && clean(row.N) && !approvers.includes(clean(row.N)))errors.push('APPROVER_NOT_BLD');
 if(changed.includes('L') && clean(row.L) && !['Đã duyệt','YC Giải trình','Chưa trình','Chưa duyệt'].includes(clean(row.L)))errors.push('INVALID_STATUS');
 return {patch:errors.length?{}:patch,errors,warnings};
}
module.exports={proposal};

};
function normalize(path){const out=[];for(const x of path.split('/')){if(x==='..')out.pop();else if(x && x!=='.')out.push(x)}return out.join('/')}
function load(id){if(cache[id])return cache[id].exports;if(!modules[id])throw Error('MODULE_NOT_PACKAGED');const m={exports:{}};cache[id]=m;const dir=id.slice(0,id.lastIndexOf('/')+1);modules[id](name=>load(normalize(dir+name)),m,m.exports);return m.exports}
globalThis.HcktQA={...load('app/three-sheet-controller.cjs'),...load('app/office-three-sheet-host.cjs')};
