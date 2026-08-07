const API_URL=(window.APP_CONFIG&&window.APP_CONFIG.API_URL)||'';
const state={student:null,tasks:[],currentTask:null};
const $=s=>document.querySelector(s);
const esc=v=>String(v??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
const fmtDate=v=>v?new Date(v).toLocaleDateString('zh-TW'):'未設定';
const labels={pending:'等待老師審核',approved:'已完成並累計 1 篇',rejected:'老師退回修改','':'尚未繳交'};

document.addEventListener('DOMContentLoaded',()=>{
  $('#studentLoginForm').addEventListener('submit',login);
  $('#studentLogout').addEventListener('click',()=>{sessionStorage.removeItem('rp-student');location.reload()});
  $('#submissionClose').addEventListener('click',closeDialog);
  $('#submissionCancel').addEventListener('click',closeDialog);
  $('#submissionForm').addEventListener('submit',submitReview);
  const saved=JSON.parse(sessionStorage.getItem('rp-student')||'null');
  if(saved){$('#studentId').value=saved.studentId||'';$('#studentClass').value=saved.className||'';loadTasks(saved.studentId,saved.className);}
});
async function request(action,payload){if(!API_URL)throw new Error('尚未在 config.js 設定 Google Apps Script API_URL');const res=await fetch(API_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action,...payload})});const json=await res.json();if(!json.success)throw new Error(json.message||'操作失敗');return json.data;}
async function login(e){e.preventDefault();await loadTasks($('#studentId').value.trim(),$('#studentClass').value.trim());}
async function loadTasks(studentId,className){try{$('#loginMessage').textContent='資料載入中…';const data=await request('getStudentTasks',{studentId,className});state.student=data.student;state.tasks=data.tasks||[];sessionStorage.setItem('rp-student',JSON.stringify({studentId,className}));$('#studentLogin').hidden=true;$('#studentApp').hidden=false;render();}catch(err){$('#loginMessage').textContent=err.message;}}
function render(){const s=state.student;$('#studentWelcome').textContent=`${s.studentName}，今天想讀哪一本書？`;$('#studentMeta').textContent=`${s.className}｜學號 ${s.studentId}`;$('#taskTotal').textContent=state.tasks.length;$('#taskPending').textContent=state.tasks.filter(t=>t.submissionStatus==='pending').length;$('#taskApproved').textContent=state.tasks.filter(t=>t.submissionStatus==='approved').length;$('#taskRejected').textContent=state.tasks.filter(t=>t.submissionStatus==='rejected').length;$('#taskList').innerHTML=state.tasks.length?state.tasks.map(taskCard).join(''):'<article class="task-card"><h2>目前沒有指定任務</h2><p>老師尚未發布適合你年級或班級的閱讀任務。</p></article>';}
function taskCard(t){const status=t.submissionStatus||'';const button=status==='approved'?'<button disabled>已完成並累計</button>':`<button onclick="openSubmission('${esc(t.taskId)}')">${status==='rejected'?'修改並重新送出':status==='pending'?'查看繳交內容':'開始寫心得'}</button>`;return `<article class="task-card"><span class="scope">${esc(t.scopeLabel||t.scopeValue)}</span><h2>${esc(t.title)}</h2><div class="book">《${esc(t.bookTitle)}》</div><p>${esc(t.description||'請完成閱讀並撰寫心得。')}</p><div class="task-meta"><span>期限：${fmtDate(t.dueDate)}</span><span>狀態：<b class="badge ${status||'none'}">${labels[status]}</b></span></div>${button}</article>`;}
window.openSubmission=function(id){const t=state.tasks.find(x=>String(x.taskId)===String(id));if(!t)return;state.currentTask=t;$('#submissionTaskId').value=t.taskId;$('#submissionTitle').textContent=`《${t.bookTitle}》閱讀心得`;$('#taskDetail').innerHTML=`<strong>${esc(t.title)}</strong><br>${esc(t.description||'完成閱讀並撰寫心得。')}<br>繳交期限：${fmtDate(t.dueDate)}`;$('#reviewTitle').value=t.reviewTitle||`${t.bookTitle}閱讀心得`;$('#reviewPages').value=t.submittedPages||t.pages||'';$('#reviewContent').value=t.reviewContent||'';$('#reviewPublic').checked=String(t.isPublic).toLowerCase()==='true';const feedback=t.teacherFeedback||'';$('#teacherFeedback').hidden=!feedback;$('#teacherFeedback').innerHTML=feedback?`<strong>老師回饋</strong><br>${esc(feedback)}`:'';$('#submissionDialog').showModal();}
function closeDialog(){$('#submissionDialog').close();}
async function submitReview(e){e.preventDefault();try{const t=state.currentTask;await request('submitTaskReview',{studentId:state.student.studentId,className:state.student.className,taskId:t.taskId,title:$('#reviewTitle').value.trim(),pages:Number($('#reviewPages').value),reviewContent:$('#reviewContent').value.trim(),isPublic:$('#reviewPublic').checked});closeDialog();toast('心得已送出，等待老師審核');await loadTasks(state.student.studentId,state.student.className);}catch(err){toast(err.message,true);}}
function toast(text,error=false){const el=$('#toast');el.textContent=text;el.style.background=error?'#a23d32':'';el.classList.add('show');setTimeout(()=>el.classList.remove('show'),2600);}
