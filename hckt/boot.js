'use strict';
// Review-only proposal. Call ONLY from an explicit QA activation button.
// No private URL/GUID/data is hardcoded or fetched from public hosting.
const guid=x=>String(x??'').replace(/[{}]/g,'').toUpperCase();
const validGUID=x=>/^[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}$/.test(guid(x));
function url(value){const u=new URL(value);if(u.protocol!=='https:' || u.username || u.password)throw Error('QA_HTTPS_URL_REQUIRED');return u}
const path=u=>decodeURIComponent(u.pathname).replace(/\/+$/,'');
function bindCurrentQA(liveURL,binding){
 const qa=guid(binding.qaWorkbookGuid),main=guid(binding.originalWorkbookGuid);
 if(!validGUID(qa) || !validGUID(main) || qa===main)throw Error('QA_AND_MAIN_IDENTITIES_REQUIRED');
 const live=url(liveURL),document=url(binding.qaDocumentUrl),file=url(binding.qaWebDavUrl);
 if(document.host!==file.host || !path(file).includes('/History/HCKT_EVENTS_QA_') || !path(file).endsWith('.xlsx') || guid(document.searchParams.get('sourcedoc'))!==qa)throw Error('QA_BINDING_NOT_HISTORY_COPY');
 const liveGUID=guid(live.searchParams.get('sourcedoc'));
 if(liveGUID===main || decodeURIComponent(liveURL).toUpperCase().includes(main))throw Error('OFFICIAL_WORKBOOK_BLOCKED');
 const exactDoc=live.host===document.host && path(live)===path(document) && liveGUID===qa;
 const exactFile=live.host===file.host && path(live)===path(file) && !liveGUID;
 if(!exactDoc && !exactFile)throw Error('CURRENT_WORKBOOK_NOT_APPROVED_QA');
 return {enabled:true,singleSessionMode:true,approvedQAUrl:liveURL,qaWorkbookGuid:qa,originalWorkbookGuid:main,baselineTable:'tblHcktBaseline',registerTable:'tblHcktRegister',allowInitialStateSeed:true};
}
async function activateQA({Office,binding,start,stop,assertActive=()=>{}}){
 const liveURL=await new Promise((resolve,reject)=>Office.context.document.getFilePropertiesAsync(r=>r.status===Office.AsyncResultStatus.Succeeded?resolve(r.value.url):reject(Error('FILE_URL_UNAVAILABLE'))));
 assertActive();const cfg=bindCurrentQA(liveURL,binding),settings=Office.context.document.settings;
 const save=()=>new Promise((resolve,reject)=>settings.saveAsync(r=>r.status===Office.AsyncResultStatus.Succeeded?resolve():reject(Error('QA_SETTINGS_SAVE_FAILED'))));
 let startAttempted=false;
 try{
  settings.set('A00.HCKT.QA',cfg);await save();assertActive();
  startAttempted=true;await start(cfg);assertActive(); // Reviewed controller/host: metadata seed only.
  cfg.allowInitialStateSeed=false;settings.set('A00.HCKT.QA',{...cfg});await save();assertActive();
  return cfg;
 }catch(error){
  const cleanupErrors=[];if(startAttempted)try{await stop()}catch(e){cleanupErrors.push(e.message)}
  cfg.enabled=false;cfg.allowInitialStateSeed=false;settings.set('A00.HCKT.QA',{...cfg});try{await save()}catch(e){cleanupErrors.push(e.message)}
  if(cleanupErrors.length)throw Error('QA_ACTIVATION_FAILED_REMOVAL_NOT_VERIFIED: '+error.message+'; '+cleanupErrors.join('; '));
  throw error;
 }
}
globalThis.HcktQAActivation={bindCurrentQA,activateQA};

let activationVersion=0,activationPending=null,stopping=false;
/* global Office, Excel, HcktQA */
'use strict';

// Local, allowlisted diagnostic labels only. No URL, row data or telemetry.
const qaResources=[
 ['sheet','PURCHASE','ĐỀ NGHỊ MUA'],['sheet','PRODUCTION','ĐỀ NGHỊ SX'],
 ['sheet','EXPORT','ĐỀ NGHỊ XUẤT'],['sheet','STATE','_HCKT_STATE'],
 ['sheet','LOOKUP','_HCKT_LOOKUP'],['sheet','SOURCE','_A00_SOURCE_QA'],
 ['table','BASELINE','tblHcktBaseline'],['table','REGISTER','tblHcktRegister'],
 ['table','CATALOG','tblA00IntegrationAudit'],
 ['table','PURCHASE_ROWS','tblCurrentPurchase'],['table','PRODUCTION_ROWS','tblCurrentProduction'],
 ['table','EXPORT_ROWS','tblCurrentExport']];
const qaTrustedDiagnostics=new WeakMap();
const qaApplicationCodes=new Set([
 'QA_CONFIGURATION_DISABLED','OFFICE_RUNTIME_UNSUPPORTED','FILE_URL_UNAVAILABLE','QA_DOCUMENT_ONLY',
 'LOCAL_SESSION_ALREADY_ACTIVE','HANDLER_CLEANUP_REQUIRED','LOCAL_SESSION_STOPPED','LATEST_STATE_CHANGED',
 'STATE_BOUNDS','SHEET_SCHEMA_OR_BOUNDS','EXISTING_UID_MISSING_OR_DUPLICATE',
 'QA_BASELINE_SEED_NOT_ENABLED','PROJECT_DISPLAY_NOT_UNIQUE','BLD_COUNT_NOT_THREE',
 'HANDLER_REMOVAL_NOT_VERIFIED','STARTUP_VERIFICATION_FAILED','STARTUP_REMOVAL_NOT_VERIFIED'
]);
function qaTrustedError(message){const error=Error(message);qaTrustedDiagnostics.set(error,message);return error}
function qaDiagnostic(stage,error){
 const known=['ItemNotFound','InvalidArgument','GeneralException','AccessDenied','ApiNotFound','InvalidOperation'];
 const code=known.includes(error?.code)?error.code:'UNCLASSIFIED';
 // Do not display native message/debugInfo: they can include private content.
 return qaTrustedError('QA_DIAGNOSTIC stage='+stage+' code='+code);
}
async function qaResourcePreflight(Excel){
 for(const [kind,label,name] of qaResources){
  let missing=false;
  try{await Excel.run(async context=>{
   const items=kind==='sheet'?context.workbook.worksheets:context.workbook.tables;
   const item=items.getItemOrNullObject(name);item.load('isNullObject');
   await context.sync();missing=item.isNullObject;
  })}catch(error){throw qaDiagnostic('PREFLIGHT_'+label,error)}
  if(missing)throw qaTrustedError('QA_DIAGNOSTIC stage=PREFLIGHT_'+label+' code=MISSING_RESOURCE');
 }
}
function qaDiagnosticHost(host,Excel){
 const stages={guardQA:'IDENTITY_GUARD',claimSession:'CLAIM_SESSION',bootstrapBounded:'BOOTSTRAP',addHandler:'ADD_HANDLER',setStartup:'SET_STARTUP',resetStartupVerified:'RESET_STARTUP',removeHandler:'REMOVE_HANDLER',disposeAllHandlers:'DISPOSE_HANDLERS',releaseSession:'RELEASE_SESSION'};
 for(const [method,stage] of Object.entries(stages)){
  if(typeof host[method]!=='function')continue;
  const original=host[method].bind(host);
  host[method]=async(...args)=>{
   if(method==='bootstrapBounded')await qaResourcePreflight(Excel);
   try{return await original(...args)}catch(error){
    const trusted=qaTrustedDiagnostics.get(error);
    if(trusted)throw qaTrustedError(trusted);
    // Exact known application codes only; create a fresh bounded error.
    if(qaApplicationCodes.has(error?.message))throw Error(error.message);
    throw qaDiagnostic(stage,error);
   }
  };
 }
 return host;
}
globalThis.HcktQADiagnostics={qaResourcePreflight,qaDiagnosticHost};

let controller,host;
const view={status:(state,error)=>{document.getElementById('status').textContent=state;document.getElementById('details').textContent=error||''},report:plans=>{const warnings=plans.flatMap(p=>[...(p.errors||[]),...(p.warnings||[])]);view.status(warnings.length?'Có dòng cần đối soát; xem chi tiết':'Tự động đang hoạt động trên bản QA',warnings.join('\n'))}};
Office.onReady(async info=>{
 if(globalThis.__hcktThreeSheetStarted)return;globalThis.__hcktThreeSheetStarted=true;
 if(info.host!==Office.HostType.Excel){view.status('Mở trong Excel để kiểm thử');return}
 try{
  const abort=new AbortController(),timer=setTimeout(()=>abort.abort(),10000);let defaults;
  try{const response=await fetch('./config.json',{signal:abort.signal});if(!response.ok)throw Error('STATIC_CONFIG_UNAVAILABLE');defaults=await response.json()}finally{clearTimeout(timer)}
  // QA target and operator acknowledgement stay in this document's add-in
  // settings, never in publicly hosted assets. No business cells sent to host.
  const privateCfg=Office.context.document.settings.get('A00.HCKT.QA')||{};
  const cfg={...defaults,...privateCfg};
  host=qaDiagnosticHost(HcktQA.createOfficeHost(Excel,Office,cfg,view),Excel);controller=HcktQA.controller(host);await controller.start();
  document.getElementById('disable').disabled=false;
 }catch(error){view.status('Tự động chưa sẵn sàng',error.message);document.getElementById('disable').disabled=!controller}
 document.getElementById('disable').onclick=async()=>{ if(stopping)return;stopping=true;activationVersion++;host?.invalidateSession?.();view.status('Stopping QA; waiting for pending activation and cleanup');if(activationPending)await activationPending.catch(()=>{});
  const errors=[];
  if(controller)try{await controller.uninstall()}catch(e){errors.push(e.message)}
  try{const cfg=Office.context.document.settings.get('A00.HCKT.QA')||{};Office.context.document.settings.set('A00.HCKT.QA',{...cfg,enabled:false});await new Promise((resolve,reject)=>Office.context.document.settings.saveAsync(r=>r.status===Office.AsyncResultStatus.Succeeded?resolve():reject(Error('QA_SETTINGS_SAVE_FAILED'))))}catch(e){errors.push(e.message)}
  view.status(errors.length?'Chưa xác minh gỡ sạch; cần kiểm tra lại':'Đã dừng tự động; cần remove Add-in và mở lại để nghiệm thu',errors.join('\n'));stopping=false;document.getElementById('activateQA').disabled=errors.length>0 || controller?.ready===true;
 };

 const activate=document.getElementById('activateQA');let activating=false;
 activate.disabled=controller?.ready===true;
 activate.onclick=async()=>{
  if(activating || stopping || controller?.ready)return;activating=true;activate.disabled=true;const version=++activationVersion;
  try{
   const input=document.getElementById('qaBinding');const binding=JSON.parse(input.value);input.value='';
   activationPending=HcktQAActivation.activateQA({Office,binding,assertActive:()=>{if(version!==activationVersion)throw Error('QA_ACTIVATION_STOPPED')},start:async cfg=>{
    if(controller)await controller.uninstall();
    host=qaDiagnosticHost(HcktQA.createOfficeHost(Excel,Office,cfg,view),Excel);controller=HcktQA.controller(host);await controller.start();document.getElementById('disable').disabled=false;
   },stop:async()=>{if(controller)await controller.uninstall()}});
   await activationPending;
   view.status('QA active for this bound copy; request numbers remain manual.');
  }catch(error){view.status('QA activation not ready; review required',error.message)}
  finally{activating=false;activationPending=null;activate.disabled=stopping || controller?.ready===true}
 };
});

