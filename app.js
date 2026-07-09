// app.js - tryLoadDefaultBank fallback to jsDelivr
// Minor addition: if fetch('./study.xlsx') fails, try CDN: https://cdn.jsdelivr.net/gh/gan10011/study@hard/study.xlsx

const LS_KEYS = {
  BANK: 'qb_bank_v1',
  PROGRESS_PREFIX: 'qb_progress_',
  WRONG: 'qb_wrong_v3',
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

// global state
const state = {
  bank: [],
  view: null, // 'sequential'|'special'|'mock'|'wrong'
  session: null // current session
};

// DOM
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

let wrongBtn = null;
let toggleCardBtn = null;
let examTimer = null;

// events
fileInput.addEventListener('change', onFileChange);
clearBankBtn.addEventListener('click', onClearBank);

modeSeqBtn.addEventListener('click', () => enterMode('sequential'));
modeSpecialBtn.addEventListener('click', () => enterMode('special'));
modeMockBtn.addEventListener('click', () => enterMode('mock'));

prevBtn.addEventListener('click', ()=> navigateTo(state.session ? state.session.index - 1 : 0));
nextBtn.addEventListener('click', ()=> navigateTo(state.session ? state.session.index + 1 : 0));
markBtn.addEventListener('click', toggleMark);
submitMockBtn.addEventListener('click', submitMock);

// init
(function init(){
  try {
    state.bank = loadBank();
    renderBankCount();

    // try load default study.xlsx (only when local storage has no bank)
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

    renderEmptyView();
    console.log('初始化完成，题库数：', state.bank.length);
  } catch (err) {
    console.error('初始化错误：', err);
  }
})();

// ----------------- Try load ./study.xlsx (only if local bank empty) -----------------
async function tryLoadDefaultBank(){
  try {
    const existing = loadBank();
    if (existing && existing.length > 0) {
      console.log('本地已有题库，默认 study.xlsx 加载跳过');
      return;
    }

    // attempt local relative path first
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
      } catch (err){
        return { ok: false, err };
      }
    };

    console.log('尝试从 ./study.xlsx 加载默认题库...');
    let r = await tryFetch('./study.xlsx');
    if(!r.ok){
      console.log('相对路径加载失败，尝试使用 jsDelivr CDN 作为回退。');
      const cdn = 'https://cdn.jsdelivr.net/gh/gan10011/study@hard/study.xlsx';
      r = await tryFetch(cdn);
      if(!r.ok){
        console.warn('通过 jsDelivr 也未能加载 study.xlsx：', r);
        return;
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

// ----------------- File import -----------------
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
      console.log('导入成功，题库条目：', parsed.length);
    } catch (err) {
      console.error('解析 Excel 错误：', err);
      alert('导入失败，控制台有错误信息（F12 打开）。');
    }
  };
  reader.readAsArrayBuffer(f);
}

// parse rows from sheet_to_json
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

// ----------------- Rendering & modes -----------------
function renderBankCount(){ bankCount.textContent = state.bank.length + '（错题 ' + loadWrong().length + '）'; }

function renderEmptyView(){
  state.view = null;
  state.session = null;
  statusMode.textContent = '无';
  modeSeqBtn.classList.remove('active');
  modeSpecialBtn.classList.remove('active');
  modeMockBtn.classList.remove('active');
  if(wrongBtn) wrongBtn.classList.remove('active');

  modeConfig.innerHTML = '<div class="hint">请选择模式开始练习，先导入题库（Excel）</div>';
  answerCardWrap.innerHTML = '';
  questionHeader.innerHTML = '';
  questionBody.innerHTML = '<div class="empty-placeholder">请选择模式开始练习，先导入题库（Excel）或等待默认题库加载。</div>';
  questionBody.style.minHeight = '0';
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

// sequential
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

// special
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

// wrong book
function prepareWrong(existing){
  const wrongList = loadWrong();
  const autoRemove = loadWrongAutoRemove();
  modeConfig.innerHTML = `<div>错题本：当前错题 ${wrongList.length} 题</div>
    <div style="margin-top:6px"><label><input type="checkbox" id="autoRemoveWrong" ${autoRemove ? 'checked' : ''}/> 答对后自动移出错题本</label></div>`;
  const chk = document.getElementById('autoRemoveWrong');
  chk.addEventListener('change', ()=>{ saveWrongAutoRemove(!!chk.checked); });

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

// wrong record: only add new entry when incorrect; always append history if existing
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
        // do not create new item when answer correct and no existing entry
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

// ----------------- Mock exam (history, review, redo) -----------------
function prepareMock(existing){
  const history = loadMockHistory();
  let histHtml = '';
  if(history.length){
    histHtml = `<div style="margin-top:8px"><strong>历史考试</strong><ul id="mockHistList" style="padding-left:18px">` +
      history.map(h => `<li data-id="${h.id}" style="margin-bottom:6px">[${new Date(h.startedAt).toLocaleString()}] 得分:${h.score}/${h.total} <button class="viewMock" data-id="${h.id}">查看</[...]