/* global Office, Excel, HcktQA */
'use strict';
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
  host=HcktQA.createOfficeHost(Excel,Office,cfg,view);controller=HcktQA.controller(host);await controller.start();
  document.getElementById('disable').disabled=false;
 }catch(error){view.status('Tự động chưa sẵn sàng',error.message);document.getElementById('disable').disabled=!controller}
 document.getElementById('disable').onclick=async()=>{
  const errors=[];
  if(controller)try{await controller.uninstall()}catch(e){errors.push(e.message)}
  try{const cfg=Office.context.document.settings.get('A00.HCKT.QA')||{};Office.context.document.settings.set('A00.HCKT.QA',{...cfg,enabled:false});await new Promise((resolve,reject)=>Office.context.document.settings.saveAsync(r=>r.status===Office.AsyncResultStatus.Succeeded?resolve():reject(Error('QA_SETTINGS_SAVE_FAILED'))))}catch(e){errors.push(e.message)}
  view.status(errors.length?'Chưa xác minh gỡ sạch; cần kiểm tra lại':'Đã dừng tự động; cần remove Add-in và mở lại để nghiệm thu',errors.join('\n'));
 };
});
