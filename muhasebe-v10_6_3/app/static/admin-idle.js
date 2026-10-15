(()=>{
const body=document.body;if(!body.dataset.adminTimeout)return;
let deadline=Date.now()+Number(body.dataset.adminRemaining)*1000;
let pending=false,busy=false,lastSent=0,lastEvent=0,leaving=false;
function leave(){if(leaving)return;leaving=true;location.replace('/admin-login');}
async function request(url,options){try{const response=await fetch(url,{credentials:'same-origin',cache:'no-store',...options});if(response.status===401||response.redirected){leave();return;}if(!response.ok)return;const data=await response.json();if(typeof data.remaining==='number')deadline=Date.now()+data.remaining*1000;}catch{/* Server checks expiry even when this browser cannot reach it. */}}
async function check(){await request('/admin/idle-status');if(Date.now()>=deadline)leave();}
async function send(){if(!pending||busy||leaving)return;busy=true;pending=false;lastSent=Date.now();try{await request('/admin/activity',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({csrf_token:body.dataset.csrf,activity_age:Math.max(0,(Date.now()-lastEvent)/1000).toString()})});}finally{busy=false;}}
function activity(event){if(!event.isTrusted)return;if(Date.now()>=deadline){if(event.cancelable)event.preventDefault();check();return;}lastEvent=Date.now();pending=true;if(Date.now()-lastSent>=20000||deadline-Date.now()<25000)send();}
['pointerdown','pointermove','keydown','input','wheel','touchstart'].forEach(name=>document.addEventListener(name,activity,{capture:true,passive:name==='wheel'||name==='touchstart'}));
setInterval(()=>{if(Date.now()>=deadline)check();else if(pending&&Date.now()-lastSent>=20000)send();},1000);
setInterval(check,30000);
document.addEventListener('visibilitychange',()=>{if(!document.hidden)check();});
window.addEventListener('pageshow',check);
})();
