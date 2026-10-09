/* No project data leaves Excel. Target identity stored only in this workbook's settings. */
'use strict';
const SETTING='KAI.HCKT.ProjectApproval.v1';let numbering=null, active=null, starting=false, actionBusy=false;
const view=(state,detail='')=>{document.getElementById('status').textContent=state;document.getElementById('details').textContent='Build: '+(globalThis.HcktPurchaseEngine?.BUILD||'UNKNOWN')+'\n'+detail;};
function formatNumberingReport(r){
 const reasons={
  HISTORICAL_NO_BACKFILL:'Dòng lịch sử: giữ nguyên, không cấp bù số',INCOMPLETE_NO_NUMBER:'Chưa đủ dữ liệu: kiểm tra mã/tên CT, nội dung, người đề nghị và ngày trình',
  RESERVED_OR_CLEARED_NO_REISSUE:'Số đã được giữ hoặc đã xóa: cần đối chiếu, không cấp lại',MANUAL_INVALID_PRESERVED:'Giữ số sửa tay; định dạng chưa hợp lệ',
  MANUAL_NONCANONICAL_PRESERVED:'Giữ số sửa tay; định dạng chuẩn dùng ít nhất 2 chữ số',MANUAL_MISSING_UID_STOP:'Số sửa tay chưa có mã dòng hợp lệ; cần kiểm tra',
  MANUAL_RESERVED_NUMBER_CONFLICT:'Số sửa tay trùng số đã giữ; cần kiểm tra',DUPLICATE_MANUAL_PRESERVED:'Có số sửa tay trùng nhau; cần kiểm tra',
  PROJECT_CHANGED_NUMBER_PRESERVED:'Đã đổi công trình: giữ nguyên số cũ để đối chiếu',MANUAL_CONFLICT_REVIEW_REQUIRED:'Chưa cấp số vì có số sửa tay cần kiểm tra'
 };
 if(r.stopped)return 'Đã dừng cấp số.\n'+String(r.error||'Cần kiểm tra trước khi tiếp tục.').replace(/[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}/gi,'[mã dòng]');
 const proposed=r.writes||[],warnings=r.warnings||[];
 const lines=[r.numbering==='PREVIEW'?'Xem trước: '+proposed.length+' dòng có thể cấp số. Chưa ghi vào file.':r.numbering==='COMMITTED'?'Đã cấp số cho '+r.count+' dòng. Đã ghi nhận các số hợp lệ.':'Không có số mới cần ghi.'];
 if(r.numbering==='PREVIEW')for(const w of proposed.slice(0,30))lines.push('Dòng '+(w.row+1)+': '+w.number);
 if(proposed.length>30)lines.push('Còn '+(proposed.length-30)+' dòng có thể cấp số.');
 const unique=[...new Map(warnings.map(w=>[JSON.stringify([w.row,w.rows,w.code]),w])).values()];
 if(unique.length)lines.push('Cần kiểm tra: '+unique.length+' thông báo.');
 for(const w of unique.slice(0,30)){const location=w.rows?'Dòng '+w.rows.map(n=>n+1).join(', '):Number.isInteger(w.row)?'Dòng '+(w.row+1):'Lưu ý';lines.push(location+': '+(reasons[w.code]||'Cần kiểm tra dữ liệu ('+w.code+')'));}
 if(unique.length>30)lines.push('Còn '+(unique.length-30)+' thông báo; chọn ít dòng hơn để kiểm tra.');
 return lines.join('\n');
}
function identity(url){const u=new URL(url);if(u.protocol!=='https:'||!u.hostname.endsWith('.sharepoint.com'))throw Error('SHAREPOINT_URL_REQUIRED');const g=(u.searchParams.get('sourcedoc')||'').replace(/[{}]/g,'').toLowerCase();if(g){if(!/^[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$/.test(g))throw Error('INVALID_SOURCEDOC');return{origin:u.origin,id:'guid:'+g};}if(!/\.xlsx$/i.test(u.pathname))throw Error('CANONICAL_XLSX_URL_REQUIRED');return{origin:u.origin,id:'path:'+decodeURIComponent(u.pathname)};}
const save=()=>new Promise((ok,no)=>Office.context.document.settings.saveAsync(r=>r.status===Office.AsyncResultStatus.Succeeded?ok():no(Error(r.error.message))));
async function stopBoth(){const hosts=[numbering,active];numbering=null;active=null;let failure;for(const host of hosts)if(host){try{await host.stop();}catch(e){failure=failure||e;}}if(failure)throw failure;}
async function start(){
 if(starting)throw Error('START_ALREADY_IN_PROGRESS');starting=true;document.getElementById('disable').disabled=true;document.getElementById('activate').disabled=true;try{
 if(numbering){await numbering.stop();numbering=null;}
 if(active){await active.stop();active=null;}
 const cfg=Office.context.document.settings.get(SETTING);
 if(!cfg||!cfg.enabled){view('Chưa kích hoạt');return;}
 if(!Office.context.requirements.isSetSupported('ExcelApi','1.14')||!Office.context.requirements.isSetSupported('SharedRuntime','1.1'))throw Error('EXCEL_REQUIREMENTS_UNAVAILABLE');
 const verifyIdentity=async()=>{const now=identity(Office.context.document.url);if(now.origin!==cfg.origin||now.id!==cfg.id)throw Error('WORKBOOK_IDENTITY_MISMATCH');};
 active=await HcktProjectHost.startProjectOnly(Excel,HcktProjectEngine,{enabled:true,supportsTriggerSource:true,verifyIdentity,report:r=>{if(r.stopped)view('Đã dừng an toàn',r.error==='EDIT_DURING_STRUCTURE_RESEED_RESTART_AND_RESELECT_STATUS'?'Có sửa dữ liệu trong lúc cập nhật sau sắp xếp. Khởi động lại, kiểm tra ngày của các dòng vừa sửa; sửa ngày trực tiếp hoặc chuyển trạng thái sang mục khác rồi chọn lại trạng thái mong muốn.':JSON.stringify(r));else if(r.state==='PAUSED_RESEED')view('Đang cập nhật sau sắp xếp; vui lòng chờ',r.message);else if(r.state==='READY_AFTER_STRUCTURE')view('Đang hoạt động','Đã cập nhật sau sắp xếp.');else view('Đang hoạt động; có nội dung cần kiểm tra',JSON.stringify(r));}});
 numbering=await HcktPurchaseHost.startPurchaseNumbering(Excel,HcktPurchaseEngine,HcktProjectEngine,{enabled:true,verifyIdentity,report:r=>view(r.stopped?'Cấp số đã dừng an toàn':'ĐỀ NGHỊ MUA',formatNumberingReport(r))});
 view('Đang hoạt động','Mã/tên dự án ↔ địa chỉ; ngày duyệt theo ngày Việt Nam.');
 }catch(e){await stopBoth();throw e;}finally{starting=false;document.getElementById('disable').disabled=false;document.getElementById('activate').disabled=false;}
}
Office.onReady(async info=>{
 if(info.host!==Office.HostType.Excel){view('Hãy mở trong Excel');return;}
 for(const [id,method] of [['previewNumbers','preview'],['allocateNumbers','allocate']])document.getElementById(id).onclick=async()=>{if(actionBusy||starting)return;actionBusy=true;try{if(!numbering)throw Error('ACTIVATE_BOUND_WORKBOOK_FIRST');await numbering[method]();}catch(e){view('Cấp số cần kiểm tra',formatNumberingReport({stopped:true,error:String(e)}));}finally{actionBusy=false;}};
 document.getElementById('current').textContent=Office.context.document.url||'Không lấy được link bản hiện tại';
 document.getElementById('activate').onclick=async()=>{if(actionBusy||starting)return;actionBusy=true;try{
 const wanted=identity(document.getElementById('target').value), current=identity(Office.context.document.url);
 if(wanted.origin!==current.origin||wanted.id!==current.id)throw Error('TARGET_NOT_THIS_WORKBOOK');
 Office.context.document.settings.set(SETTING,{...wanted,enabled:true});await save();await start();
 await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
 view('Đang hoạt động; đã bật tự nạp');
 }catch(e){try{await stopBoth();}catch(_){}Office.context.document.settings.remove(SETTING);try{await save();await Office.addin.setStartupBehavior(Office.StartupBehavior.none);}catch(_){}view('Chưa sẵn sàng',String(e));}finally{actionBusy=false;}};
 document.getElementById('restart').onclick=()=>{if(actionBusy||starting)return;start().catch(e=>view('Đã dừng',String(e)));};
 document.getElementById('disable').onclick=async()=>{if(actionBusy||starting)return;actionBusy=true;try{if(numbering){await numbering.stop();numbering=null;}
 if(active){await active.stop();active=null;}Office.context.document.settings.remove(SETTING);await save();await Office.addin.setStartupBehavior(Office.StartupBehavior.none);view('Đã tắt và bỏ tự nạp');}catch(e){view('Tắt chưa hoàn tất',String(e));}finally{actionBusy=false;}};
 try{await start();}catch(e){view('Đã dừng',String(e));}
});
