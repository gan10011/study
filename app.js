// app.js - 完整版（含搜索 + 错题导出）
// 主要改动：
// - tryLoadDefaultBank (./study.xlsx -> 回退 jsDelivr)
// - 模考交卷后记录 wrongQids 并在答题卡上标红
// - 添加题库全文搜索（题目/选项/解析）并作为临时 session 浏览
// - 错题本增加“导出错题 Excel”按钮
// 依赖：SheetJS (xlsx) 已在 index.html 引入

const LS_KEYS = {
  BANK: 'qb_bank_v1',
  PROGRESS_PREFIX: 'qb_progress_',
  WRONG: 'qb_wrong_v2',
  MOCK_HISTORY: 'qb_mock_hist_v1',
  CARD_COLLAPSED: 'qb_card_collapsed_v1',
  WRONG_AUTO_REMOVE: 'qb_wrong_auto_remove_v1'
};

function saveBank(bank){ localStorage.setItem(LS_KEYS.BANK, JSON.stringify(bank)); }
function loadBank(){ const s = localStorage.getItem(LS_KEYS.BANK); return s ? JSON.parse(s) : []; }

function saveProgress(key, data){ localStorage.setItem(LS_KEYS.PROGRESS_PREFIX + key, JSON.stringify(data)); }
function loadProgress(key){ const s = localStorage.getItem(LS_KEYS.PROGRESS_PREFIX + key); return s ? JSON.parse(s) : null; }
function clearProgress(key){ localStorage.removeItem(LS_KEYS.PROGRESS_PREFIX + key); }

function saveWrong(arr){ localStorage.setItem(LS_KEYS.WRONG, JSON.stringify(arr)); }
function loadWrong(){ const s = localStorage.getItem(LS_KEYS.WRONG); return s ? JSON.parse(s) : []; }

function saveMockHistory(arr){ localStorage.setItem(LS_KEYS.MOCK_HISTORY, JSON.stringify(arr)); }
function loadMockHistory(){ const s = localStorage.getItem(LS_KEYS.MOCK_HISTORY); return s ? JSON.parse(s) : []; }

function saveCardCollapsed(v){ localStorage.setItem(LS_KEYS.CARD_COLLAPSED, JSON.stringify(!!v)); }
function loadCardCollapsed(){ const s = localStorage.getItem(LS_KEYS.CARD_COLLAPSED); return s ? JSON.parse(s) : false; }

function saveWrongAutoRemove(v){ localStorage.setItem(LS_KEYS.WRONG_AUTO_REMOVE, JSON.stringify(!!v)); }
function loadWrongAutoRemove(){ const s = localStorage.getItem(LS_KEYS.WRONG_AUTO_REMOVE); return s ? JSON.parse(s) : false; }

function uid(){ return Date.now().toString(36) + Math.random().toString(36).slice(2,8); }

const state = { bank: [], view: null, session: null };

// DOM refs
const fileInput = document.getElementById('fileInput');
const clearBankBtn = document.getElementById('clearBank');
const bankCount = document.getElementById('bankCount');

const modeSeqBtn = document.getElementById('mode-seq');
const modeSpecialBtn = document.getElementById('mode-special');
const modeMockBtn = document.getElementById('mode-mock');

const answerCardWrap = document.getElementById('answerCardWrap');
const modeConfig = document.getElementById('modeConfig');
const questionHeader = document.getElementById('questionHeader');
const questionBody = document.getElementById('questionBody');
const questionFooter = document.getElementById('questionFooter');

const prevBtn = document.getElementById('prevBtn');
const nextBtn = document.getElementById('nextBtn');
const markBtn = document.getElementById('markBtn');
const submitMockBtn = document.getElementById('submitMock');
const examTimerDisplay = document.getElementById('examTimer');
const timerDisplay = document.getElementById('timerDisplay');

const statusMode = document.getElementById('currentMode');

// search controls (index.html must include #searchInput and #clearSearch)
const searchInput = document.getElementById('searchInput');
const clearSearchBtn = document.getElementById('clearSearch');

let wrongBtn = null;
let toggleCardBtn = null;
let examTimer = null;

fileInput && fileInput.addEventListener('change', onFileChange);
clearBankBtn && clearBankBtn.addEventListener('click', onClearBank);

modeSeqBtn && modeSeqBtn.addEventListener('click', () => enterMode('sequential'));
modeSpecialBtn && modeSpecialBtn.addEventListener('click', () => enterMode('special'));
modeMockBtn && modeMockBtn.addEventListener('click', () => enterMode('mock'));

prevBtn && prevBtn.addEventListener('click', ()=> navigateTo(state.session ? state.session.index - 1 : 0));
nextBtn && nextBtn.addEventListener('click', ()=> navigateTo(state.session ? state.session.index + 1 : 0));
markBtn && markBtn.addEventListener('click', toggleMark);
submitMockBtn && submitMockBtn.addEventListener('click', submitMock);

(function init(){
  try {
    state.bank = loadBank();
    renderBankCount();

    // try load default study.xlsx when local bank empty
    tryLoadDefaultBank();

    const modesNode = document.querySelector('.modes');
    if(modesNode){
      wrongBtn = document.createElement('button');
      wrongBtn.id = 'mode-wrong';
      wrongBtn.textContent = '错题本';
      wrongBtn.addEventListener('click', () => enterMode('wrong'));
      modesNode.appendChild(wrongBtn);

      toggleCardBtn = document.createElement('button');
      toggleCardBtn.id = 'toggleCardBtn';
      toggleCardBtn.style.marginLeft = '8px';
      toggleCardBtn.addEventListener('click', toggleAnswerCard);
      answerCardWrap.parentNode.insertBefore(toggleCardBtn, answerCardWrap.nextSibling);

      updateCardToggleLabel();
    }

    // search bindings
    if (searchInput) {
      searchInput.addEventListener('input', onSearchInput);
    }
    if (clearSearchBtn) {
      clearSearchBtn.addEventListener('click', ()=> {
        if (searchInput) searchInput.value = '';
        onSearchInput();
      });
    }

    renderEmptyView();
    console.log('初始化完成，题库数：', state.bank.length);
  } catch (err) {
    console.error('初始化错误：', err);
  }
})();

// ----------------- try load default bank (relative -> cdn fallback) -----------------
async function tryLoadDefaultBank(){
  try {
    const existing = loadBank();
    if (existing && existing.length > 0) {
      console.log('本地已有题库，默认 study.xlsx 加载跳过');
      return;
    }

    const tryFetch = async (url) => {
      try {
        const res = await fetch(url, { cache: 'no-store' });
        if (!res.ok) return { ok: false, status: res.status };
        const ab = await res.arrayBuffer();
        const wb = XLSX.read(ab, { type: 'array' });
        const first = wb.SheetNames[0];
        const ws = wb.Sheets[first];
        const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
        const parsed = parseRows(json);
        return { ok: true, parsed };
      } catch (err) {
        return { ok: false, err };
      }
    };

    console.log('尝试从 ./study.xlsx 加载默认题库...');
    let r = await tryFetch('./study.xlsx');
    if(!r.ok){
      console.log('相对路径加载失败，尝试使用 jsDelivr CDN 作为回退。');
      const cdnHard = 'https://cdn.jsdelivr.net/gh/gan10011/study@hard/study.xlsx';
      const cdnMain = 'https://cdn.jsdelivr.net/gh/gan10011/study@main/study.xlsx';
      r = await tryFetch(cdnHard);
      if(!r.ok){
        // 继续尝试 main
        r = await tryFetch(cdnMain);
        if(!r.ok){
          console.warn('通过 jsDelivr 也未能加载 study.xlsx：', r);
          return;
        }
      }
    }

    const parsed = r.parsed;
    if (parsed && parsed.length) {
      state.bank = parsed;
      saveBank(parsed);
      renderBankCount();
      console.log('默认题库已加载，题目数：', parsed.length);
      if(!state.view) renderEmptyView();
    } else {
      console.log('study.xlsx 解析后无数据');
    }
  } catch (err) {
    console.warn('加载默认题库 study.xlsx 失败：', err);
  }
}

// ----------------- file import -----------------
function onFileChange(e){
  const f = e.target.files[0];
  if(!f) return;
  const reader = new FileReader();
  reader.onload = function(ev){
    try {
      const data = ev.target.result;
      const wb = XLSX.read(data, {type:'array'});
      const firstName = wb.SheetNames[0];
      const ws = wb.Sheets[firstName];
      const json = XLSX.utils.sheet_to_json(ws, {defval: ''});
      const parsed = parseRows(json);
      state.bank = parsed;
      saveBank(parsed);
      renderBankCount();
      alert('题库已导入，共 ' + parsed.length + ' 题（已保存到本地）。');
      fileInput.value = '';
      console.log('导入成功，条目：', parsed.length);
    } catch (err) {
      console.error('解析 Excel 错误：', err);
      alert('导入失败，见控制台错误信息（F12）。');
    }
  };
  reader.readAsArrayBuffer(f);
}

// ----------------- parse rows -----------------
function parseRows(rows){
  const out = [];
  for(const r of rows){
    const rawType = r['题型'] || r['类型'] || r['Type'] || '';
    const type = (rawType || '').toString().trim();
    const content = (r['题目'] || r['问题'] || r['Question'] || '').toString().trim();
    const optRaw = (r['选项'] || r['选项内容'] || '').toString().trim();
    const answerRaw = (r['答案'] || r['Answer'] || '').toString().trim();
    const analysis = (r['解析'] || r['说明'] || '').toString().trim();

    let ttype = '判断题';
    if(/单选|选择/.test(type)) ttype = '单选题';
    else if(/判断|对错|是非/.test(type)) ttype = '判断题';
    else { if(optRaw && /\|/.test(optRaw)) ttype = '单选题'; }

    let options = [];
    if(optRaw){
      options = optRaw.split(/\||\n|；|;/).map(s => s.trim()).filter(Boolean);
    }else if(ttype === '判断题'){
      options = ['正确','错误'];
    }

    let answer = answerRaw;
    if(/^(对|正确|True|true)$/i.test(answerRaw)) answer = '正确';
    if(/^(错|错误|False|false)$/i.test(answerRaw)) answer = '错误';

    let opts = options;
    if(ttype === '单选题' && options.length){
      const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
      opts = options.map((t,i)=> ({key: letters[i] || String(i+1), text: t}));
    }

    out.push({
      id: r['序号'] ? String(r['序号']) : uid(),
      type: ttype,
      question: content,
      options: opts,
      rawOptions: options,
      answer: answer,
      analysis: analysis
    });
  }
  return out;
}

// ----------------- render / UI -----------------
function renderBankCount(){
  const all = state.bank.length;
  const wrong = loadWrong().length;
  bankCount.textContent = `${all}（错题 ${wrong}）`;
}

function renderEmptyView(){
  state.view = null;
  state.session = null;
  statusMode.textContent = '无';
  modeSeqBtn.classList.remove('active');
  modeSpecialBtn.classList.remove('active');
  modeMockBtn.classList.remove('active');
  if(wrongBtn) wrongBtn.classList.remove('active');

  modeConfig.innerHTML = '<div class="hint">请选择模式开始练习，先导入题库（Excel）或等待默认题库加载。</div>';
  answerCardWrap.innerHTML = '';
  questionHeader.innerHTML = '';
  questionBody.innerHTML = '';
  questionFooter.innerHTML = '';
  examTimerDisplay.style.display = 'none';
  submitMockBtn.style.display = 'none';
}

function updateCardToggleLabel(){
  const collapsed = loadCardCollapsed();
  if(!toggleCardBtn) return;
  toggleCardBtn.textContent = collapsed ? '展开答题卡' : '折叠答题卡';
}

function toggleAnswerCard(){
  const collapsed = !loadCardCollapsed();
  saveCardCollapsed(collapsed);
  updateCardToggleLabel();
  renderAnswerCard();
}

function enterMode(view){
  if(state.view === view) return;
  const key = modeProgressKey(view);
  const existing = loadProgress(key);
  let resume = false;
  if(existing && !existing.finished){
    resume = confirm('检测到该模式有未完成进度，是否恢复上次进度？（确定恢复，取消从头开始）');
  }
  state.view = view;
  modeSeqBtn.classList.toggle('active', view === 'sequential');
  modeSpecialBtn.classList.toggle('active', view === 'special');
  modeMockBtn.classList.toggle('active', view === 'mock');
  if(wrongBtn) wrongBtn.classList.toggle('active', view === 'wrong');
  statusMode.textContent = view === 'sequential' ? '顺序练习' : view === 'special' ? '专项练习' : view === 'mock' ? '模拟考试' : '错题本';

  if(view === 'sequential') prepareSequential(resume ? existing : null);
  else if(view === 'special') prepareSpecial(resume ? existing : null);
  else if(view === 'mock') prepareMock(resume ? existing : null);
  else if(view === 'wrong') prepareWrong(resume ? existing : null);
}

function modeProgressKey(view, extras=''){ return view + (extras ? ('_' + extras) : ''); }

function prepareSequential(existing){
  modeConfig.innerHTML = '<div>顺序练习：按题库顺序练习，答对自动下一题，答错将加入错题本</div>';
  const qlist = state.bank.slice();
  const key = modeProgressKey('sequential');
  if(existing) state.session = existing;
  else {
    state.session = { id: uid(), mode:'sequential', filter:{type:'all'}, qids: qlist.map(q=>q.id), index:0, answers:{}, marked:{}, startedAt: Date.now() };
    saveProgress(key, state.session);
  }
  renderAnswerCard(); renderCurrentQuestion();
}

function prepareSpecial(existing){
  modeConfig.innerHTML = `
    <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
      <label>题型：
        <select id="specialType"><option value="判断题">判断题</option><option value="单选题">单选题</option></select>
      </label>
      <label id="judgeFilterWrap" style="display:inline-block">
        判断筛选：
        <select id="judgeFilter"><option value="all">全部</option><option value="正确">练答案为“对/正确”的题</option><option value="错误">练答案为“错/错误”的题</option></select>
      </label>
      <button id="startSpecial">开始专项练习</button>
    </div>
  `;
  document.getElementById('specialType').addEventListener('change', (e)=>{
    const wrap = document.getElementById('judgeFilterWrap');
    wrap.style.display = e.target.value === '判断题' ? 'inline-block' : 'none';
  });
  document.getElementById('startSpecial').addEventListener('click', ()=>{
    const type = document.getElementById('specialType').value;
    const judgeFilter = document.getElementById('judgeFilter').value;
    startSpecialSession(type, judgeFilter, existing);
  });

  if(existing){ state.session = existing; renderAnswerCard(); renderCurrentQuestion(); }
  else { questionHeader.innerHTML = '<div>尚未开始专项练习，请选择题型并点击“开始专项练习”。</div>'; questionBody.innerHTML=''; questionFooter.innerHTML=''; answerCardWrap.innerHTML=''; }
}

function startSpecialSession(type, judgeFilter, existing){
  let list = state.bank.filter(q => q.type === type);
  if(type === '判断题' && judgeFilter && judgeFilter !== 'all'){
    list = list.filter(q => {
      const a = (q.answer || '').toString();
      return /正|对/i.test(judgeFilter) ? /正|对/i.test(a) : /错|误/i.test(a);
    });
  }
  if(list.length === 0){ alert('没有满足条件的题目。'); return; }
  const key = modeProgressKey('special', `${type}_${judgeFilter}`);
  const exist = loadProgress(key);
  if(exist && !exist.finished){
    const resume = confirm('检测到该专项筛选有未完成进度，是否恢复上次进度？（确定恢复，取消重新开始）');
    if(resume){ state.session = exist; renderAnswerCard(); renderCurrentQuestion(); return; } else clearProgress(key);
  }

  state.session = { id: uid(), mode:'special', filter:{type,judgeFilter}, qids: list.map(q=>q.id), index:0, answers:{}, marked:{}, startedAt: Date.now() };
  saveProgress(key, state.session);
  renderAnswerCard(); renderCurrentQuestion();
}

function prepareWrong(existing){
  const wrongList = loadWrong();
  const autoRemove = loadWrongAutoRemove();
  // 添加导出按钮
  modeConfig.innerHTML = `
    <div>错题本：当前错题 ${wrongList.length} 题 
      <button id="exportWrongBtn" style="margin-left:12px; padding:4px 12px; background:#4caf50; color:#fff; border:none; border-radius:4px; cursor:pointer;">📥 导出错题 Excel</button>
    </div>
    <div style="margin-top:6px"><label><input type="checkbox" id="autoRemoveWrong" ${autoRemove ? 'checked' : ''}/> 答对后自动移出错题本（可切换）</label></div>
  `;
  // 绑定导出事件
  document.getElementById('exportWrongBtn')?.addEventListener('click', exportWrongToExcel);
  const chk = document.getElementById('autoRemoveWrong');
  chk && chk.addEventListener('change', ()=>{ saveWrongAutoRemove(!!chk.checked); });

  if(wrongList.length === 0){ questionHeader.innerHTML = '<div>错题本为空。</div>'; questionBody.innerHTML=''; questionFooter.innerHTML=''; answerCardWrap.innerHTML=''; return; }

  const qids = wrongList.map(w => w.id);
  const key = modeProgressKey('wrong');
  const existingProgress = loadProgress(key);
  if(existingProgress && !existingProgress.finished){
    const resume = confirm('检测到错题本有未完成进度，是否恢复上次进度？');
    if(resume){ state.session = existingProgress; renderAnswerCard(); renderCurrentQuestion(); return; } else clearProgress(key);
  }

  state.session = { id: uid(), mode:'wrong', filter:{type:'wrong'}, qids, index:0, answers:{}, marked:{}, startedAt: Date.now() };
  saveProgress(key, state.session);
  renderAnswerCard(); renderCurrentQuestion();
}

// addWrongRecord, removeWrongEntry unchanged
function addWrongRecord(qid, selected, mode, correct){
  try{
    const q = getQuestionById(qid);
    if(!q) return;
    const arr = loadWrong();
    let item = arr.find(x=>x.id == qid);
    const now = Date.now();
    const histEntry = { time: now, mode: mode || state.view || '', selected: selected === undefined ? null : String(selected), correct: !!correct };
    if(!item){
      if(!correct){
        item = { id: qid, q: JSON.parse(JSON.stringify(q)), wrongCount: 1, history: [histEntry] };
        arr.push(item);
      } else {
        // 做对且无条目，什么也不做（避免新增）
      }
    } else {
      item.history = item.history || [];
      item.history.push(histEntry);
      if(!correct) item.wrongCount = (item.wrongCount || 0) + 1;
    }
    saveWrong(arr);
    renderBankCount();
  }catch(e){ console.error('addWrongRecord err', e); }
}

function removeWrongEntry(qid){
  try{
    let arr = loadWrong();
    arr = arr.filter(x => x.id != qid);
    saveWrong(arr);
    renderBankCount();
  }catch(e){ console.error('removeWrongEntry err', e); }
}

// ----------------- Mock exam -----------------
function prepareMock(existing){
  const history = loadMockHistory();
  let histHtml = '';
  if(history.length){
    histHtml = `<div style="margin-top:8px"><strong>历史考试</strong><ul id="mockHistList" style="padding-left:18px">` +
      history.map(h => `<li data-id="${h.id}" style="margin-bottom:6px">[${new Date(h.startedAt).toLocaleString()}] 得分:${h.score}/${h.total} <button class="viewMock" data-id="${h.id}">查看</button> <button class="redoMock" data-id="${h.id}">重做</button></li>`).join('') +
      `</ul></div>`;
  }
  modeConfig.innerHTML = '<div>模拟考试：70 道判断 + 30 道单选，总时长 60 分钟</div><div style="margin-top:8px"><button id="startMockBtn">开始 模拟考试</button></div>' + histHtml;
  const btn = document.getElementById('startMockBtn');
  btn && btn.addEventListener('click', ()=> startMockSession(existing));
  setTimeout(()=>{
    document.querySelectorAll('.viewMock').forEach(b=> b.addEventListener('click', ()=> reviewMock(b.dataset.id)));
    document.querySelectorAll('.redoMock').forEach(b=> b.addEventListener('click', ()=> redoMock(b.dataset.id)));
  }, 50);

  if(existing && confirm('检测到未完成的模拟考试，是否恢复？（确定恢复，取消重新开始）')){ startMockSession(existing, true); }
  else { questionHeader.innerHTML = '<div>尚未开始模拟考试，点击“开始 模拟考试”开始。</div>'; questionBody.innerHTML=''; questionFooter.innerHTML=''; answerCardWrap.innerHTML=''; }
}

function startMockSession(existing, forceResume=false){
  const key = modeProgressKey('mock');
  if(existing && forceResume){ state.session = existing; startExamTimer(); renderAnswerCard(); renderCurrentQuestion(); submitMockBtn.style.display = 'block'; examTimerDisplay.style.display='block'; return; }
  const judges = state.bank.filter(q=>q.type==='判断题');
  const singles = state.bank.filter(q=>q.type==='单选题');
  const pick = (arr,n)=>{ const a = arr.slice(); for(let i=a.length-1;i>0;i--){ const j=Math.floor(Math.random()*(i+1)); [a[i],a[j]]=[a[j],a[i]]; } return a.slice(0, Math.min(n,a.length)); };
  const examList = [].concat(pick(judges,70), pick(singles,30));
  if(examList.length===0){ alert('题库中没有合适题目，无法开始模拟考试。'); return; }
  state.session = { id: uid(), mode:'mock', filter:{examSize: examList.length}, qids: examList.map(q=>q.id), index:0, answers:{}, marked:{}, startedAt:Date.now(), durationMin:60, endsAt: Date.now()+60*60*1000 };
  saveProgress(key, state.session);
  startExamTimer(); renderAnswerCard(); renderCurrentQuestion(); submitMockBtn.style.display='block'; examTimerDisplay.style.display='block';
}

function startExamTimer(){
  if(examTimer) clearInterval(examTimer);
  if(!state.session || !state.session.endsAt) return;
  examTimer = setInterval(()=>{
    if(!state.session) return;
    const left = Math.max(0, Math.floor((state.session.endsAt - Date.now())/1000));
    const mm = String(Math.floor(left/60)).padStart(2,'0');
    const ss = String(left % 60).padStart(2,'0');
    timerDisplay.textContent = `${mm}:${ss}`;
    if(left <= 0){
      clearInterval(examTimer);
      alert('考试时间到，自动交卷。');
      submitMock();
    }
  }, 500);
}

function addMockHistory(record){ const arr = loadMockHistory(); arr.unshift(record); if(arr.length>50) arr.length=50; saveMockHistory(arr); }

function submitMock(){
  if(!state.session || state.session.mode !== 'mock') return;
  if(!confirm('确认交卷并查看结果？')) return;
  let total=0, correct=0;
  for(const qid of state.session.qids){
    total++; const q = getQuestionById(qid); const user = state.session.answers[qid]; if(user==null) continue; const ok = checkCorrect(q,user); if(ok) correct++;
  }
  clearInterval(examTimer);
  const endedAt = Date.now();
  const rec = { id: state.session.id, qids: state.session.qids.slice(), answers: Object.assign({}, state.session.answers), score: correct, total, startedAt: state.session.startedAt, endedAt };
  addMockHistory(rec);

  // 在交卷时把错题保存到错题本，并为每题写入历史（只为错题新增条目；对的仅补充已有条目）
  for(const qid of state.session.qids){
    const q = getQuestionById(qid);
    const user = state.session.answers[qid];
    const ok = (user != null) && checkCorrect(q,user);
    addWrongRecord(qid, user, 'mock', ok);
  }

  // collect wrong qids for UI highlighting in this session
  const wrongQids = [];
  for (const qid of state.session.qids) {
    const q = getQuestionById(qid);
    const user = state.session.answers[qid];
    const ok = (user != null) && checkCorrect(q, user);
    if (!ok) wrongQids.push(qid);
  }
  state.session.wrongQids = wrongQids;

  alert(`交卷完成：共 ${total} 题，答对 ${correct} 题，得分 ${correct} 分。进入答题回顾（显示正确答案与解析）。`);
  state.session.finished = true;
  saveProgress(modeProgressKey('mock'), state.session);
  renderAnswerCard(); renderCurrentQuestion();
}

function reviewMock(historyId){
  const history = loadMockHistory();
  const rec = history.find(h=>h.id==historyId);
  if(!rec){ alert('找不到历史记录'); return; }
  state.session = { id: rec.id + '_review', mode:'mock', qids: rec.qids.slice(), index:0, answers: Object.assign({}, rec.answers), marked:{}, startedAt: rec.startedAt, endedAt: rec.endedAt, finished:true };
  // calculate wrongQids for this review session to highlight
  state.session.wrongQids = state.session.qids.filter(qid => {
    const user = state.session.answers[qid];
    const q = getQuestionById(qid);
    return !(user != null && checkCorrect(q, user));
  });

  state.view = 'mock';
  modeSeqBtn.classList.remove('active'); modeSpecialBtn.classList.remove('active'); modeMockBtn.classList.add('active'); if(wrongBtn) wrongBtn.classList.remove('active');
  statusMode.textContent = '模拟考试（复卷）';
  renderAnswerCard(); renderCurrentQuestion();
  submitMockBtn.style.display='none'; examTimerDisplay.style.display='none';
}

function redoMock(historyId){
  const history = loadMockHistory();
  const rec = history.find(h=>h.id==historyId);
  if(!rec){ alert('找不到历史记录'); return; }
  state.session = { id: uid(), mode:'mock', filter:{examSize: rec.qids.length}, qids: rec.qids.slice(), index:0, answers:{}, marked:{}, startedAt:Date.now(), durationMin:60, endsAt: Date.now()+60*60*1000 };
  saveProgress(modeProgressKey('mock'), state.session);
  startExamTimer(); renderAnswerCard(); renderCurrentQuestion(); submitMockBtn.style.display='block'; examTimerDisplay.style.display='block';
}

// ----------------- render answer card -----------------
function renderAnswerCard(){
  if(!state.session){ answerCardWrap.innerHTML=''; updateCardToggleLabel(); return; }
  const total = state.session.qids.length;
  const arr = state.session.qids;
  const collapsed = loadCardCollapsed();
  const currentIndex = state.session.index || 0;
  const isMockFinished = state.session.mode === 'mock' && !!state.session.finished;

  if(collapsed){
    const cur = currentIndex;
    const winStart = Math.max(0, cur - 10);
    const winEnd = Math.min(total, cur + 9 + 1); // 前10，后9（含当前）
    const chipsHtml = arr.slice(winStart, winEnd).map((qid, idx) => {
      const globalIdx = winStart + idx;
      const answered = state.session.answers && (state.session.answers[qid] !== undefined && state.session.answers[qid] !== null);
      const marked = state.session.marked && state.session.marked[qid];
      const isWrongInSession = state.session && Array.isArray(state.session.wrongQids) && state.session.wrongQids.includes(qid);
      const classes = ['q-chip'];
      if(globalIdx === currentIndex) classes.push('current');
      if(answered) classes.push('answered');
      if(marked) classes.push('marked');
      if(isWrongInSession) classes.push('wrong');
      return `<div class="${classes.join(' ')}" data-idx="${globalIdx}" title="题 ${globalIdx+1}">${globalIdx+1}</div>`;
    }).join('');
    const info = `显示 ${winStart+1} ~ ${winEnd} / ${total}`;
    answerCardWrap.innerHTML = `<div style="padding:6px;background:#fff;border-radius:6px;display:flex;flex-direction:column;gap:6px"><div>${info}</div><div style="display:flex;gap:4px;flex-wrap:wrap">${chipsHtml}</div></div>`;
    answerCardWrap.querySelectorAll('.q-chip').forEach(n => n.addEventListener('click', ()=> navigateTo(Number(n.dataset.idx))));
  } else {
    const chips = arr.map((qid, idx) => {
      const answered = state.session.answers && (state.session.answers[qid] !== undefined && state.session.answers[qid] !== null);
      const marked = state.session.marked && state.session.marked[qid];
      const isWrongInSession = state.session && Array.isArray(state.session.wrongQids) && state.session.wrongQids.includes(qid);
      const classes = ['q-chip'];
      if(idx === currentIndex) classes.push('current');
      if(answered) classes.push('answered');
      if(marked) classes.push('marked');
      if(isWrongInSession) classes.push('wrong');
      return `<div class="${classes.join(' ')}" data-idx="${idx}" title="题 ${idx+1}">${idx+1}</div>`;
    }).join('');
    answerCardWrap.innerHTML = `<div class="answer-card">${chips}</div>`;
    answerCardWrap.querySelectorAll('.q-chip').forEach(n => n.addEventListener('click', ()=> navigateTo(Number(n.dataset.idx))));
  }
  updateCardToggleLabel();
}

// ----------------- render current question -----------------
function renderCurrentQuestion(){
  if(!state.session){ questionHeader.innerHTML=''; questionBody.innerHTML=''; questionFooter.innerHTML=''; return; }
  const idx = state.session.index;
  const qid = state.session.qids[idx];
  const q = getQuestionById(qid);
  if(!q){ questionHeader.innerHTML = '<div>找不到题目</div>'; return; }

  questionHeader.innerHTML = `第 ${idx+1} / ${state.session.qids.length} 题 <span style="margin-left:8px;color:#666">(${q.type})</span>`;
  const paraText = q.question || '';
  let para = `<div style="margin-bottom:8px">${escapeHtml(paraText)}</div>`;
  // if in search mode, highlight query
  if(state.view === 'search' && state.session && state.session.filter && state.session.filter.query){
    para = `<div style="margin-bottom:8px">${highlightMatch(escapeHtml(paraText), state.session.filter.query)}</div>`;
  }

  let optsHtml = '';
  if(q.type === '判断题'){
    optsHtml = `<div class="options"><div class="option" data-val="正确">正确</div><div class="option" data-val="错误">错误</div></div>`;
  } else {
    const opts = (q.options || []).map(o => {
      const key = o.key || '';
      const text = (o.text || o).toString();
      const showText = (state.view === 'search' && state.session && state.session.filter && state.session.filter.query) ? highlightMatch(escapeHtml(text), state.session.filter.query) : escapeHtml(text);
      return `<div class="option" data-val="${escapeHtml(key)}"><strong>${escapeHtml(key)}.</strong> ${showText}</div>`;
    }).join('');
    optsHtml = `<div class="options">${opts}</div>`;
  }
  questionBody.innerHTML = para + optsHtml;

  const isMock = state.session.mode === 'mock';
  const finished = !!state.session.finished;
  if(isMock) questionFooter.innerHTML = `<div><small>模拟考试/复卷模式：答题中不显示解析；交卷后或复卷显示解析。</small></div>`;
  else questionFooter.innerHTML = `<div><small>选择后将显示是否正确与解析（顺序模式答对自动下一题）。</small></div>`;

  questionBody.querySelectorAll('.option').forEach(n => n.addEventListener('click', ()=> handleAnswer(q, n.dataset.val)));

  // 错题统计（仅在非模考进行中，或已交卷/复卷时显示）
  const wrongItem = loadWrong().find(x => x.id == qid);
  if(wrongItem && (state.session.mode !== 'mock' || finished)){
    const hist = (wrongItem.history || []).slice(-10).reverse();
    const histHtml = `<div style="margin-top:8px;padding:8px;border-top:1px dashed #eee"><strong>错题统计：</strong>已错 ${wrongItem.wrongCount || 0} 次<br/><strong>最近记录（最多10条）：</strong><ul style="padding-left:18px">` +
      hist.map(h=>`<li>${new Date(h.time).toLocaleString()} ${h.mode ? '['+h.mode+']' : ''} ${h.selected != null ? '选: '+h.selected : ''} ${h.correct ? '<span style="color:green">正确</span>' : '<span style="color:red">错误</span>'}</li>`).join('') +
      `</ul></div>`;
    questionFooter.innerHTML += histHtml;
  }

  // restore previous selection
  const prev = state.session.answers[qid];
  if(prev != null){
    questionBody.querySelectorAll('.option').forEach(n => { if(n.getAttribute('data-val') == prev) n.classList.add('selected'); else n.classList.remove('selected'); });
    if(state.session.mode !== 'mock' || finished) showAnswerFeedback(q, prev);
  }

  // mark as wrong label if finished and this q was wrong in this session
  if(state.session && state.session.finished && Array.isArray(state.session.wrongQids) && state.session.wrongQids.includes(qid)){
    questionHeader.innerHTML += ' <span style="color:#b71c1c;margin-left:8px">【错题】</span>';
  }

  const isMarked = !!state.session.marked[qid];
  markBtn.textContent = isMarked ? '取消标记' : '标记';
  renderAnswerCard();
}

function getQuestionById(id){
  let q = state.bank.find(x=>x.id == id);
  if(!q){
    const wrong = loadWrong().find(x=>x.id==id);
    if(wrong) q = wrong.q;
  }
  return q;
}

// ----------------- answer handling -----------------
function handleAnswer(q, val){
  if(!state.session) return;
  const qid = q.id;
  state.session.answers[qid] = val;
  const pkey = state.session.mode === 'special' ? modeProgressKey('special', `${state.session.filter.type}_${state.session.filter.judgeFilter}`) : modeProgressKey(state.session.mode);
  saveProgress(pkey, state.session);

  const isCorrect = checkCorrect(q, val);

  // 在模拟考试进行中不在每题时立即写入错题库
  if(state.session.mode !== 'mock'){
    addWrongRecord(qid, val, state.session.mode || '', isCorrect);
  }

  // 在错题本模式中，答对自动移出（若设置）
  if(state.session.mode === 'wrong' && isCorrect && loadWrongAutoRemove()){
    removeWrongEntry(qid);
    const i = state.session.qids.indexOf(qid);
    if(i>=0){ state.session.qids.splice(i,1); if(state.session.index >= state.session.qids.length) state.session.index = Math.max(0, state.session.qids.length-1); saveProgress(modeProgressKey('wrong'), state.session); }
  }

  if(state.session.mode === 'sequential'){
    if(isCorrect){ renderAnswerFeedbackInline(q, true); setTimeout(()=> navigateTo(state.session.index+1), 600); return; }
    else { renderAnswerFeedbackInline(q, false); return; }
  } else if(state.session.mode === 'special' || state.session.mode === 'wrong' || state.session.mode === 'search'){
    renderAnswerFeedbackInline(q, isCorrect); return;
  } else if(state.session.mode === 'mock'){
    // 在模考中：如果不是最后一题自动跳转下一题；最后一题停留，等待交卷（不显示即时解析）
    if(state.session.index < state.session.qids.length - 1){
      navigateTo(state.session.index + 1);
    } else {
      // 最后一题停留
    }
    return;
  }
}

function checkCorrect(q, val){
  if(!q) return false;
  if(q.type === '判断题'){
    const ra = (q.answer||'').toString().trim();
    if(/正|对/i.test(ra)) return /正|对/i.test(val);
    if(/错|误/i.test(ra)) return /错|误/i.test(val);
    return ra === val;
  } else {
    const ra = (q.answer||'').toString().trim();
    if(val === ra) return true;
    const optByKey = (q.options||[]).find(o => (o.key||'')==ra);
    if(optByKey && (val == optByKey.key || val == optByKey.text)) return true;
    const optByText = (q.options||[]).find(o => (o.text||'')==ra);
    if(optByText && (val == optByText.key || val == optByText.text)) return true;
    return false;
  }
}

function renderAnswerFeedbackInline(q, isCorrect){
  const nodes = questionBody.querySelectorAll('.option');
  const user = state.session.answers[q.id];
  nodes.forEach(n=>{
    n.classList.remove('correct','wrong');
    const v = n.getAttribute('data-val');
    if(v == user){ if(isCorrect) n.classList.add('correct'); else n.classList.add('wrong'); }
    if(q.type === '单选题'){
      const ra = q.answer;
      if(ra) nodes.forEach(m=>{ if(m.getAttribute('data-val') == ra) m.classList.add('correct'); });
    } else {
      if(q.answer) nodes.forEach(m=>{ if(m.getAttribute('data-val') == q.answer) m.classList.add('correct'); });
    }
  });
  const anal = q.analysis ? `<div style="margin-top:10px;padding:8px;border-top:1px dashed #eee"><strong>解析：</strong><div>${escapeHtml(q.analysis)}</div></div>` : '';
  questionFooter.innerHTML = `<div>${isCorrect ? '<span style="color:green">回答正确</span>' : '<span style="color:red">回答错误</span>'}</div>${anal}<div style="margin-top:8px"><button id="goNextBtn">下一题</button></div>`;
  const goNextBtn = document.getElementById('goNextBtn');
  if(goNextBtn) goNextBtn.addEventListener('click', ()=> navigateTo(state.session.index + 1));
  renderAnswerCard();
}

function showAnswerFeedback(q, userVal){
  const nodes = questionBody.querySelectorAll('.option');
  const isCorrect = checkCorrect(q, userVal);
  nodes.forEach(n=>{
    const v = n.getAttribute('data-val');
    if(v == userVal){ if(isCorrect) n.classList.add('correct'); else n.classList.add('wrong'); }
    if(q.answer && (v == q.answer || (q.options||[]).find(o=>o.key==q.answer && o.key==v))) n.classList.add('correct');
  });
  const anal = q.analysis ? `<div style="margin-top:10px;padding:8px;border-top:1px dashed #eee"><strong>解析：</strong><div>${escapeHtml(q.analysis)}</div></div>` : '';
  questionFooter.innerHTML = `<div>${isCorrect ? '<span style="color:green">回答正确</span>' : '<span style="color:red">回答错误</span>'}</div>${anal}`;
}

// ----------------- navigation -----------------
function navigateTo(idx){
  if(!state.session) return;
  if(idx < 0) idx = 0;
  if(idx >= state.session.qids.length) idx = state.session.qids.length - 1;
  state.session.index = idx;
  const key = state.session.mode === 'special' ? modeProgressKey('special', `${state.session.filter.type}_${state.session.filter.judgeFilter}`) : modeProgressKey(state.session.mode);
  saveProgress(key, state.session);
  renderAnswerCard(); renderCurrentQuestion();
}

function toggleMark(){
  if(!state.session) return;
  const qid = state.session.qids[state.session.index];
  if(!qid) return;
  state.session.marked[qid] = !state.session.marked[qid];
  const key = state.session.mode === 'special' ? modeProgressKey('special', `${state.session.filter.type}_${state.session.filter.judgeFilter}`) : modeProgressKey(state.session.mode);
  saveProgress(key, state.session);
  renderCurrentQuestion();
}

function onClearBank(){
  if(!confirm('确认清空本地题库？此操作会删除保存在浏览器中的题库。')) return;
  localStorage.removeItem(LS_KEYS.BANK);
  state.bank = [];
  renderBankCount();
  renderEmptyView();
  alert('题库已清空。');
}

/* ---------------- search helpers ---------------- */
function highlightMatch(text, q) {
  if (!q) return text;
  const esc = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  try {
    return text.replace(new RegExp(esc, 'ig'), m => `<mark>${m}</mark>`);
  } catch(e) {
    return text;
  }
}

function escapeHtml(str) {
  if (!str && str !== 0) return '';
  return String(str).replace(/[&<>"']/g, function(m){ return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]); });
}

function onSearchInput(){
  if(!searchInput) return;
  const raw = searchInput.value || '';
  const q = raw.trim().toLowerCase();
  if (!q) {
    // 退出搜索：回到空视图（若之前为 search）
    if (state.view === 'search') {
      renderEmptyView();
      renderBankCount();
    }
    return;
  }

  // 全文匹配：题目 + rawOptions + 解析
  const results = state.bank.filter(item => {
    const hay = (item.question || '') + ' ' + (Array.isArray(item.rawOptions) ? item.rawOptions.join(' ') : '') + ' ' + (item.analysis || '');
    return hay.toString().toLowerCase().includes(q);
  });

  state.view = 'search';
  state.session = {
    id: 'search_' + uid(),
    mode: 'search',
    filter: { query: q },
    qids: results.map(r => r.id),
    index: 0,
    answers: {},
    marked: {},
    startedAt: Date.now()
  };

  modeConfig.innerHTML = `<div>搜索 "${escapeHtml(raw)}"：共 ${results.length} 条结果（题目/选项/解析匹配）</div>`;
  statusMode.textContent = `搜索: ${raw}`;
  renderAnswerCard();
  renderCurrentQuestion();
}

// ----------------- 导出错题到 Excel -----------------
function exportWrongToExcel() {
  try {
    const wrongList = loadWrong();
    if (!wrongList || wrongList.length === 0) {
      alert('错题本为空，无数据可导出。');
      return;
    }

    // 准备导出数据（每行一个错题记录）
    const rows = wrongList.map((item, index) => {
      const q = item.q;
      // 提取选项文本（原始选项或格式化）
      let optText = '';
      if (q.type === '判断题') {
        optText = '正确 / 错误';
      } else if (Array.isArray(q.options)) {
        optText = q.options.map(o => `${o.key || ''}. ${o.text || o}`).join('；');
      } else {
        optText = q.rawOptions ? q.rawOptions.join('；') : '';
      }

      // 最近错误时间（取最新的一条历史记录）
      const lastTime = item.history && item.history.length > 0 
        ? new Date(item.history[item.history.length - 1].time).toLocaleString()
        : '';

      return {
        '序号': index + 1,
        '题型': q.type,
        '题目': q.question || '',
        '选项': optText,
        '答案': q.answer || '',
        '解析': q.analysis || '',
        '错误次数': item.wrongCount || 0,
        '最近错误时间': lastTime,
        '历史记录': (item.history || []).map(h => 
          `${new Date(h.time).toLocaleString()} [${h.mode || ''}] 选择:${h.selected || ''} ${h.correct ? '✓' : '✗'}`
        ).join('；')
      };
    });

    // 创建工作簿
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, '错题本');

    // 生成 Excel 文件并下载
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `错题本_${new Date().toISOString().slice(0,10)}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);

    alert(`成功导出 ${rows.length} 条错题记录。`);
  } catch (err) {
    console.error('导出错题失败：', err);
    alert('导出失败，请查看控制台错误信息。');
  }
}