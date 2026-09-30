const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const hist=$('#history'), msg=$('#msg'), flow=$('#flow'), flowStatus=$('#flowStatus');

function dashboardUrl(url='/api/dashboard/download'){
  const u=url||'/api/dashboard/download';
  return u+(u.includes('?')?'&':'?')+'t='+Date.now();
}
function showDashboardDownload(url='/api/dashboard/download'){
  const href=dashboardUrl(url);
  const area=$('#downloadArea'), dl=$('#dashboardDownload');
  if(area&&dl){
    dl.href=href;
    dl.setAttribute('download','Jaseer_MTD_Performance_Dashboard.xlsx');
    area.hidden=false;
    area.removeAttribute('hidden');
    area.style.cssText='display:block!important;padding:8px 10px;background:#091624;position:relative;z-index:50;';
    dl.style.cssText='display:block!important;text-align:center;padding:11px 14px;border-radius:8px;background:#16d77b;color:#04150d;text-decoration:none;font-weight:800;';
  }
  // Also place a second link directly inside chat history. This cannot be hidden by the chat grid.
  let inline=$('#inlineDashboardDownload');
  if(!inline){
    inline=document.createElement('a');
    inline.id='inlineDashboardDownload';
    inline.textContent='⬇ Download Complete Excel Dashboard';
    inline.style.cssText='display:block;margin:10px 0;padding:11px 14px;border-radius:8px;background:#16d77b;color:#04150d;text-decoration:none;font-weight:800;text-align:center;';
    hist.appendChild(inline);
  }
  inline.href=href;
  inline.setAttribute('download','Jaseer_MTD_Performance_Dashboard.xlsx');
  hist.scrollTop=hist.scrollHeight;
}
async function downloadDashboard(url='/api/dashboard/download'){
  showDashboardDownload(url);
  const href=dashboardUrl(url);
  try{
    const r=await fetch(href,{cache:'no-store'});
    if(!r.ok){let m='Excel download failed';try{const d=await r.json();m=d.error||m}catch{}throw Error(m)}
    const blob=await r.blob();
    if(!blob.size)throw Error('Generated Excel file is empty.');
    const obj=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=obj;a.download='Jaseer_MTD_Performance_Dashboard.xlsx';
    document.body.appendChild(a);a.click();a.remove();
    setTimeout(()=>URL.revokeObjectURL(obj),30000);
  }catch(e){
    add('j','Jaseer: Dashboard was created, but automatic download could not start. Use the green Download Complete Excel Dashboard button. ('+e.message+')');
  }
}
function add(cls,text){let p=document.createElement('p');p.className=cls;p.textContent=text;hist.appendChild(p);hist.scrollTop=hist.scrollHeight}
function resetRooms(){$$('.room').forEach(x=>{x.classList.remove('working','done');x.querySelector('em').textContent=x.dataset.dept==='Jaseer'?'Ready':'Idle'})}
function setRoom(name,state){let r=$(`.room[data-dept="${name}"]`);if(!r)return;r.classList.remove('working','done');r.classList.add(state==='Working'?'working':'done');r.querySelector('em').textContent=state}
function drawFlow(steps,idx){flow.innerHTML='';steps.forEach((s,i)=>{let n=document.createElement('div');n.className='node '+(i<idx?'done':i===idx?'work':'wait');n.innerHTML=`<b>${s}</b><small>● ${i<idx?'Completed':i===idx?'Working':'Waiting'}</small>`;flow.appendChild(n);if(i<steps.length-1){let a=document.createElement('span');a.className='arrow';a.textContent='→';flow.appendChild(a)}})}
async function ask(q){q=(q||msg.value).trim();if(!q)return;msg.value='';add('u','You: '+q);add('j','Jaseer: Understanding your instruction…');resetRooms();setRoom('Jaseer','Working');flowStatus.textContent='● Processing';
 try{let r=await fetch('/api/chat',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({message:q})});let raw=await r.text();let d;try{d=JSON.parse(raw)}catch{throw Error('Server returned a non-JSON response. Check Render logs.')}if(!r.ok)throw Error(d.error||'Request failed');let steps=d.steps;drawFlow(steps,0);for(let i=1;i<steps.length;i++){await new Promise(x=>setTimeout(x,430));drawFlow(steps,i);if(i===2)(d.active||[]).forEach(x=>setRoom(x,'Working'))}await new Promise(x=>setTimeout(x,430));drawFlow(steps,steps.length);(d.active||[]).forEach(x=>setRoom(x,'Completed'));setRoom('Jaseer','Completed');flowStatus.textContent='● Completed';hist.lastElementChild.remove();add('j','Jaseer: '+d.answer+(d.llm?'\n\nAI advisor: connected':'\n\nAI advisor: Python fallback'));
   if(d.download_url||d.dashboard_ready||/dashboard created/i.test(d.answer||'')){
     add('j','Jaseer: Excel dashboard is ready. Starting the download now…');
     await downloadDashboard(d.download_url||'/api/dashboard/download');
   }
   $('#source').textContent=d.source
 }
 catch(e){hist.lastElementChild.remove();add('j','Jaseer: '+e.message);flowStatus.textContent='● Error';setRoom('Jaseer','Completed')}}
$('#chatForm').addEventListener('submit',e=>{e.preventDefault();ask()});$$('[data-q]').forEach(b=>b.onclick=()=>ask(b.dataset.q));$$('.room').forEach(r=>r.onclick=()=>{let d=r.dataset.dept;if(d==='Jaseer'){msg.focus();return}ask(d==='Reporting'?'make MTD dashboard':`show ${d.toLowerCase()} summary`)})
const modal=$('#modal');$('#dataBtn').onclick=async()=>{modal.classList.add('show');let d=await (await fetch('/api/status',{cache:'no-store'})).json();$('#status').textContent=`LLM: ${d.llm_configured?'configured':'not configured'} | Google Drive: ${d.gdrive_configured?'folder ID configured':'not configured'} | Source: ${d.source}`};$('#close').onclick=()=>modal.classList.remove('show');
$('#file').addEventListener('change',()=>{let fs=[...$('#file').files];$('#selectedFiles').textContent=fs.length?`${fs.length} file(s): `+fs.map(x=>x.name).join(', '):'No files selected'});
$('#upload').onclick=async()=>{let fs=[...$('#file').files];if(!fs.length)return $('#uploadMsg').textContent='Choose one or more files first.';let fd=new FormData();fs.forEach(f=>fd.append('files',f));$('#uploadMsg').textContent=`Loading ${fs.length} file(s)…`;try{let r=await fetch('/api/upload',{method:'POST',body:fd});let raw=await r.text();let d;try{d=JSON.parse(raw)}catch{throw Error('Server returned a non-JSON response. Check Render logs.')}$('#uploadMsg').textContent=r.ok?`Loaded ${d.file_count} file(s) | Sheets: `+d.sheets.join(', '):'Error: '+d.error;if(r.ok)$('#source').textContent=d.source}catch(e){$('#uploadMsg').textContent='Error: '+e.message}};

fetch('/api/status',{cache:'no-store'}).then(r=>r.json()).then(d=>{if(d.dashboard_ready)showDashboardDownload(d.download_url||'/api/dashboard/download')}).catch(()=>{});
