/* No project data leaves Excel. Target identity stored only in this workbook's settings. */
'use strict';
const SETTING='KAI.HCKT.ProjectApproval.v1';let active=null, starting=false, actionBusy=false;
const view=(state,detail='')=>{document.getElementById('status').textContent=state;document.getElementById('details').textContent=detail;};
function identity(url){const u=new URL(url);if(u.protocol!=='https:'||!u.hostname.endsWith('.sharepoint.com'))throw Error('SHAREPOINT_URL_REQUIRED');const g=(u.searchParams.get('sourcedoc')||'').replace(/[{}]/g,'').toLowerCase();if(g){if(!/^[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$/.test(g))throw Error('INVALID_SOURCEDOC');return{origin:u.origin,id:'guid:'+g};}if(!/\.xlsx$/i.test(u.pathname))throw Error('CANONICAL_XLSX_URL_REQUIRED');return{origin:u.origin,id:'path:'+decodeURIComponent(u.pathname)};}
const save=()=>new Promise((ok,no)=>Office.context.document.settings.saveAsync(r=>r.status===Office.AsyncResultStatus.Succeeded?ok():no(Error(r.error.message))));
async function start(){
 if(starting)throw Error('START_ALREADY_IN_PROGRESS');starting=true;document.getElementById('disable').disabled=true;document.getElementById('activate').disabled=true;try{
 if(active){await active.stop();active=null;}
 const cfg=Office.context.document.settings.get(SETTING);
 if(!cfg||!cfg.enabled){view('Chưa kích hoạt');return;}
 if(!Office.context.requirements.isSetSupported('ExcelApi','1.14')||!Office.context.requirements.isSetSupported('SharedRuntime','1.1'))throw Error('EXCEL_REQUIREMENTS_UNAVAILABLE');
 const verifyIdentity=async()=>{const now=identity(Office.context.document.url);if(now.origin!==cfg.origin||now.id!==cfg.id)throw Error('WORKBOOK_IDENTITY_MISMATCH');};
 active=await HcktProjectHost.startProjectOnly(Excel,HcktProjectEngine,{enabled:true,supportsTriggerSource:true,verifyIdentity,report:r=>{if(r.stopped)view('Đã dừng an toàn',r.error==='EDIT_DURING_STRUCTURE_RESEED_RESTART_AND_RESELECT_STATUS'?'Có sửa dữ liệu trong lúc cập nhật sau sắp xếp. Khởi động lại, kiểm tra ngày của các dòng vừa sửa; sửa ngày trực tiếp hoặc chuyển trạng thái sang mục khác rồi chọn lại trạng thái mong muốn.':JSON.stringify(r));else if(r.state==='PAUSED_RESEED')view('Đang cập nhật sau sắp xếp; vui lòng chờ',r.message);else if(r.state==='READY_AFTER_STRUCTURE')view('Đang hoạt động','Đã cập nhật sau sắp xếp.');else view('Đang hoạt động; có nội dung cần kiểm tra',JSON.stringify(r));}});
 view('Đang hoạt động','Mã/tên dự án ↔ địa chỉ; ngày duyệt theo ngày Việt Nam.');
 }finally{starting=false;document.getElementById('disable').disabled=false;document.getElementById('activate').disabled=false;}
}
Office.onReady(async info=>{
 if(info.host!==Office.HostType.Excel){view('Hãy mở trong Excel');return;}
 document.getElementById('current').textContent=Office.context.document.url||'Không lấy được link bản hiện tại';
 document.getElementById('activate').onclick=async()=>{if(actionBusy||starting)return;actionBusy=true;try{
 const wanted=identity(document.getElementById('target').value), current=identity(Office.context.document.url);
 if(wanted.origin!==current.origin||wanted.id!==current.id)throw Error('TARGET_NOT_THIS_WORKBOOK');
 Office.context.document.settings.set(SETTING,{...wanted,enabled:true});await save();await start();
 await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
 view('Đang hoạt động; đã bật tự nạp');
 }catch(e){if(active){try{await active.stop();}catch(_){}active=null;}Office.context.document.settings.remove(SETTING);try{await save();await Office.addin.setStartupBehavior(Office.StartupBehavior.none);}catch(_){}view('Chưa sẵn sàng',String(e));}finally{actionBusy=false;}};
 document.getElementById('restart').onclick=()=>{if(actionBusy||starting)return;start().catch(e=>view('Đã dừng',String(e)));};
 document.getElementById('disable').onclick=async()=>{if(actionBusy||starting)return;actionBusy=true;try{if(active){await active.stop();active=null;}Office.context.document.settings.remove(SETTING);await save();await Office.addin.setStartupBehavior(Office.StartupBehavior.none);view('Đã tắt và bỏ tự nạp');}catch(e){view('Tắt chưa hoàn tất',String(e));}finally{actionBusy=false;}};
 try{await start();}catch(e){view('Đã dừng',String(e));}
});
