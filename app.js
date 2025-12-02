
/* app.js - Gerador de Simulados (vanilla JS)
 * Requisitos: carregar 'Simulados Faculdade.xlsx' via fetch na mesma pasta.
 * Fallback: input file permite selecionar uma planilha manualmente.
 * Dependências: SheetJS (XLSX) e html2pdf (in index.html via CDN).
 */

/* ---------- Helpers ---------- */
function escapeHtml(str){
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function uid(){
  return Date.now().toString(36) + Math.random().toString(36).slice(2,8);
}

/* ---------- App State ---------- */
const state = {
  allQuestions: [], // parsed from sheet
  subjects: [],     // distinct subjects
  currentQuiz: null,
  timer: null
};

/* ---------- DOM ---------- */
const subjectSelect = document.getElementById('subjectSelect');
const startBtn = document.getElementById('startBtn');
const quizArea = document.getElementById('quizArea');
const resultArea = document.getElementById('resultArea');
const historyList = document.getElementById('historyList');
const fileInput = document.getElementById('fileInput');
const numQInput = document.getElementById('numQ');
const shuffleCheckbox = document.getElementById('shuffleChoices');
const enableTimer = document.getElementById('enableTimer');
const timePerQ = document.getElementById('timePerQ');
const exportCsvBtn = document.getElementById('exportCsvBtn');
const exportPdfBtn = document.getElementById('exportPdfBtn');
const progressBar = document.getElementById('progressBar');
const darkToggle = document.getElementById('darkToggle');

/* ---------- Load XLSX via fetch (default filename) ---------- */
const DEFAULT_XLSX = 'Simulados Faculdade.xlsx';

async function fetchDefaultWorkbook(){
  try{
    const res = await fetch(DEFAULT_XLSX);
    if (!res.ok) throw new Error('arquivo não encontrado via fetch');
    const ab = await res.arrayBuffer();
    const wb = XLSX.read(ab, {type:'array'});
    return wb;
  }catch(err){
    console.warn('Não foi possível carregar planilha padrão:', err);
    return null;
  }
}

function parseSheetToQuestions(wb){
  const shtName = wb.SheetNames[0];
  const ws = wb.Sheets[shtName];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:''});
  const questions = [];
  for (let i=1;i<rows.length;i++){ // skip header row
    const r = rows[i];
    const subject = String(r[0]||'').trim();
    const enunciado = String(r[1]||'').trim();
    const optA = String(r[2]||'').trim();
    const optB = String(r[3]||'').trim();
    const optC = String(r[4]||'').trim();
    const optD = String(r[5]||'').trim();
    const gabarito = String(r[6]||'').trim().toUpperCase();
    if (!subject || !enunciado) continue;
    if (!['A','B','C','D'].includes(gabarito)){
      console.warn('Gabarito inválido na linha', i+1, '-> ignorando questão');
      continue;
    }
    questions.push({
      id: uid(),
      subject, enunciado,
      choices: {A:optA, B:optB, C:optC, D:optD},
      answer: gabarito
    });
  }
  return questions;
}

/* ---------- UI population ---------- */
function populateSubjects(){
  const subjects = Array.from(new Set(state.allQuestions.map(q=>q.subject))).sort();
  state.subjects = subjects;
  subjectSelect.innerHTML = '';
  const placeholder = document.createElement('option');
  placeholder.value='';
  placeholder.textContent = subjects.length? 'Escolha uma matéria...' : 'Nenhuma matéria encontrada';
  subjectSelect.appendChild(placeholder);
  subjects.forEach(s=>{
    const o = document.createElement('option');
    o.value = s;
    o.textContent = s;
    subjectSelect.appendChild(o);
  });
}

/* ---------- Quiz generation ---------- */
function sampleQuestions(subject, n){
  const pool = state.allQuestions.filter(q=>q.subject===subject);
  if (pool.length === 0) return [];
  const shuffled = pool.slice().sort(()=>Math.random()-0.5);
  return shuffled.slice(0, Math.min(n, pool.length)).map(q=>structuredClone(q));
}

function shuffleObjectChoices(q){
  const keys = Object.keys(q.choices);
  const entries = keys.map(k=>([k,q.choices[k]]));
  // randomize
  for (let i=entries.length-1;i>0;i--){
    const j = Math.floor(Math.random()*(i+1));
    [entries[i], entries[j]] = [entries[j], entries[i]];
  }
  // build new mapping A..D but keep track of original correct key
  const letters = ['A','B','C','D'];
  const newChoices = {};
  let newAnswer = null;
  entries.forEach(([,text],idx)=>{
    newChoices[letters[idx]] = text;
    // if this text equals original correct text, set newAnswer accordingly
    const origCorrectText = q.choices[q.answer];
    if (text === origCorrectText) newAnswer = letters[idx];
  });
  if (!newAnswer){
    // fallback: if duplicate texts or mismatch, try mapping by index of original answer
    const origIndex = letters.indexOf(q.answer);
    newAnswer = letters[origIndex] || 'A';
  }
  q.choices = newChoices;
  q.answer = newAnswer;
  return q;
}

/* ---------- Rendering quiz pages ---------- */
function renderQuiz(quiz){
  quizArea.innerHTML = ''; quizArea.hidden = false;
  resultArea.hidden = true;
  exportCsvBtn.disabled = true;
  exportPdfBtn.disabled = true;

  const q = quiz.questions[quiz.current];
  const container = document.createElement('div');
  container.className = 'question';
  const meta = document.createElement('div');
  meta.className = 'meta';
  meta.innerHTML = `<div class="progressSmall">Questão ${quiz.current+1}/${quiz.total}</div>
                    <div class="progressSmall">${quiz.subject}</div>`;
  container.appendChild(meta);

  const enun = document.createElement('div');
  enun.innerHTML = '<strong>' + escapeHtml(q.enunciado) + '</strong>';
  container.appendChild(enun);

  const ul = document.createElement('ul');
  ul.className = 'choices';
  ul.setAttribute('role','listbox');
  ul.tabIndex = 0;

  const letters = ['A','B','C','D'];
  letters.forEach(letter=>{
    const li = document.createElement('li');
    li.className = 'choice';
    li.setAttribute('data-choice', letter);
    li.tabIndex = 0;
    li.setAttribute('role','option');
    // use textContent to avoid injecting HTML
    li.textContent = letter + '. ' + q.choices[letter];
    // mark selected
    if (quiz.answers[q.id] === letter) li.classList.add('selected');
    li.addEventListener('click', ()=> {
      quiz.answers[q.id] = letter;
      // instant feedback optional (not enabled by default)
      // update UI selection
      Array.from(ul.children).forEach(c=>c.classList.remove('selected'));
      li.classList.add('selected');
    });
    li.addEventListener('keydown', (ev)=> {
      if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); li.click(); }
    });
    ul.appendChild(li);
  });

  container.appendChild(ul);

  // navigation
  const nav = document.createElement('div');
  nav.style.display='flex'; nav.style.gap='8px'; nav.style.marginTop='8px';
  const back = document.createElement('button'); back.textContent='Voltar';
  back.disabled = quiz.current === 0;
  back.addEventListener('click', ()=>{
    quiz.current = Math.max(0, quiz.current-1);
    renderQuiz(quiz);
  });
  const next = document.createElement('button'); next.textContent = (quiz.current === quiz.total-1)? 'Finalizar' : 'Próxima';
  next.addEventListener('click', ()=>{
    if (quiz.current === quiz.total-1){
      finishQuiz(quiz);
    } else {
      quiz.current = Math.min(quiz.total-1, quiz.current+1);
      renderQuiz(quiz);
    }
  });
  nav.appendChild(back); nav.appendChild(next);
  container.appendChild(nav);

  // progress bar
  progressBar.max = quiz.total;
  progressBar.value = quiz.current+1;

  quizArea.appendChild(container);
}

/* ---------- Finish and results ---------- */
function finishQuiz(quiz){
  // stop timer
  if (state.timer) { clearInterval(state.timer); state.timer = null; }
  const now = new Date();
  const timeSpent = quiz.timeElapsed || 0;
  // grading
  let correct = 0;
  const details = quiz.questions.map(q=>{
    const marked = quiz.answers[q.id] || null;
    const isCorrect = marked === q.answer;
    if (isCorrect) correct++;
    return {
      id: q.id, enunciado: q.enunciado,
      choices: q.choices, marked, correct: q.answer, isCorrect
    };
  });

  // show results UI
  quizArea.hidden = true;
  resultArea.hidden = false;
  resultArea.innerHTML = '';
  const summary = document.createElement('div');
  summary.innerHTML = `<h2>Resultado: ${correct} / ${quiz.total}</h2>
                       <div class="meta">Tempo: ${formatTime(timeSpent || 0)}</div>`;
  resultArea.appendChild(summary);

  details.forEach((d, idx)=>{
    const card = document.createElement('div');
    card.className = 'question';
    const en = document.createElement('div');
    en.innerHTML = '<strong>' + escapeHtml((idx+1)+'. '+d.enunciado) + '</strong>';
    card.appendChild(en);
    const ul = document.createElement('ul'); ul.className='choices';
    ['A','B','C','D'].forEach(l=>{
      const li = document.createElement('li');
      li.className = 'choice';
      li.textContent = l + '. ' + (d.choices[l] || '');
      if (l === d.correct) li.classList.add('correct');
      if (d.marked && l === d.marked && l !== d.correct) li.classList.add('wrong');
      // indicate user's selection
      if (d.marked === l){
        li.textContent += '  ← sua resposta';
      }
      ul.appendChild(li);
    });
    const verdict = document.createElement('div');
    verdict.className = 'meta';
    verdict.textContent = d.isCorrect? 'Acertou' : 'Errou';
    card.appendChild(ul);
    card.appendChild(verdict);
    resultArea.appendChild(card);
  });

  // save history
  const run = {
    id: uid(), datetime: new Date().toISOString(),
    subject: quiz.subject, total: quiz.total, correct, timeSpent,
    questions: details
  };
  saveHistoryItem(run);
  renderHistory();

  exportCsvBtn.disabled = false;
  exportPdfBtn.disabled = false;
  // attach export handlers (current result)
  exportCsvBtn.onclick = ()=> exportHistoryItemAsCsv(run);
  exportPdfBtn.onclick = ()=> exportRunAsPdf(run);
}

/* ---------- Storage ---------- */
function saveHistoryItem(item){
  const hist = JSON.parse(localStorage.getItem('simulados_history')||'[]');
  hist.unshift(item);
  localStorage.setItem('simulados_history', JSON.stringify(hist));
}
function getHistory(){ return JSON.parse(localStorage.getItem('simulados_history')||'[]'); }
function renderHistory(){
  const hist = getHistory();
  if (!hist.length) { historyList.innerHTML = '<div>Nenhuma tentativa registrada</div>'; return; }
  historyList.innerHTML = '';
  hist.forEach(h=>{
    const d = document.createElement('div');
    d.className = 'historyItem';
    d.innerHTML = `<strong>${escapeHtml(h.subject)}</strong> — ${escapeHtml(new Date(h.datetime).toLocaleString())} — ${h.correct}/${h.total}`;
    d.tabIndex = 0;
    d.addEventListener('click', ()=> showHistoryDetail(h));
    d.addEventListener('keydown', (e)=> { if (e.key==='Enter') showHistoryDetail(h); });
    historyList.appendChild(d);
  });
}
function showHistoryDetail(h){
  // render similarly to finishQuiz but from saved run
  quizArea.hidden = true; resultArea.hidden = false;
  resultArea.innerHTML = `<h2>Detalhes - ${escapeHtml(h.subject)} - ${escapeHtml(new Date(h.datetime).toLocaleString())}</h2>
                          <div class="meta">Acertos: ${h.correct}/${h.total} — Tempo: ${formatTime(h.timeSpent)}</div>`;
  h.questions.forEach((d,idx)=>{
    const card = document.createElement('div'); card.className='question';
    const en = document.createElement('div'); en.innerHTML = '<strong>' + escapeHtml((idx+1)+'. '+d.enunciado) + '</strong>';
    card.appendChild(en);
    const ul = document.createElement('ul'); ul.className='choices';
    ['A','B','C','D'].forEach(l=>{
      const li=document.createElement('li'); li.className='choice'; li.textContent = l + '. ' + (d.choices[l]||'');
      if (l===d.correct) li.classList.add('correct');
      if (d.marked && l===d.marked && l!==d.correct) li.classList.add('wrong');
      if (d.marked===l) li.textContent += '  ← sua resposta';
      ul.appendChild(li);
    });
    card.appendChild(ul);
    card.appendChild(Object.assign(document.createElement('div'),{className:'meta', textContent: d.isCorrect? 'Acertou' : 'Errou'}));
    resultArea.appendChild(card);
  });
}

/* ---------- Export helpers ---------- */
function exportHistoryItemAsCsv(run){
  const lines = [['questao','A','B','C','D','marcada','correta','acertou']];
  run.questions.forEach(q=>{
    lines.push([
      q.enunciado.replace(/\n/g,' '),
      q.choices.A||'', q.choices.B||'', q.choices.C||'', q.choices.D||'',
      q.marked||'', q.correct||'', q.isCorrect? '1':'0'
    ]);
  });
  const csv = lines.map(r => r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(',')).join('\n');
  const blob = new Blob([csv],{type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href=url; a.download = `simulado_${run.subject.replace(/\s+/g,'_')}_${run.datetime}.csv`; a.click();
  URL.revokeObjectURL(url);
}

function exportRunAsPdf(run){
  // create a printable element
  const el = document.createElement('div');
  el.style.padding='12px'; el.style.fontSize='12px';
  el.innerHTML = `<h2>${escapeHtml(run.subject)}</h2><div>${escapeHtml(new Date(run.datetime).toLocaleString())} — ${run.correct}/${run.total}</div><hr/>`;
  run.questions.forEach((q, i)=>{
    el.innerHTML += `<div><strong>${i+1}. ${escapeHtml(q.enunciado)}</strong></div>`;
    ['A','B','C','D'].forEach(l=>{
      const mark = (l===q.correct)? ' (correta)' : (q.marked===l? ' (sua)' : '');
      el.innerHTML += `<div>${escapeHtml(l + '. ' + (q.choices[l]||''))}${mark}</div>`;
    });
    el.innerHTML += '<hr/>';
  });
  // use html2pdf
  html2pdf().set({margin:10, filename:`simulado_${run.subject}.pdf`, html2canvas:{scale:1}}).from(el).save();
}

/* ---------- Timer ---------- */
function formatTime(seconds){
  if (!seconds) return '0s';
  const m = Math.floor(seconds/60); const s = seconds%60;
  return (m? m+'m ':'') + s + 's';
}

/* ---------- Main start flow ---------- */
async function startFlow(){
  const subject = subjectSelect.value;
  if (!subject){ alert('Escolha uma matéria.'); return; }
  const n = parseInt(numQInput.value) || 10;
  const shuffleChoices = shuffleCheckbox.checked;
  const enableT = enableTimer.checked;
  const tPerQ = parseInt(timePerQ.value) || 30;

  let questions = sampleQuestions(subject, n);
  if (!questions.length){ alert('Não há questões suficientes para esta matéria.'); return; }
  if (shuffleChoices){
    questions = questions.map(q=>shuffleObjectChoices(q));
  }
  // prepare quiz object
  const quiz = {
    id: uid(), subject, total: questions.length,
    questions, current:0, answers:{}, timeElapsed:0
  };
  state.currentQuiz = quiz;

  // timer
  if (enableT){
    const totalSeconds = tPerQ * quiz.total;
    quiz.timeLeft = totalSeconds;
    quiz.timeElapsed = 0;
    state.timer = setInterval(()=>{
      quiz.timeLeft--;
      quiz.timeElapsed++;
      if (quiz.timeLeft <= 0){
        clearInterval(state.timer); state.timer = null;
        finishQuiz(quiz);
      }
    }, 1000);
  } else {
    quiz.timeElapsed = 0;
  }

  renderQuiz(quiz);
}

/* ---------- Initialization ---------- */
async function init(){
  // dark toggle
  darkToggle.addEventListener('change', ()=> document.documentElement.classList.toggle('dark', darkToggle.checked));

  // try fetch default workbook
  let wb = await fetchDefaultWorkbook();
  if (wb){
    state.allQuestions = parseSheetToQuestions(wb);
    populateSubjects();
  } else {
    subjectSelect.innerHTML = '<option value="">Coloque a planilha na pasta ou faça upload</option>';
  }
  renderHistory();

  // start button
  startBtn.addEventListener('click', ()=> {
    // clear previous results area
    resultArea.innerHTML = '';
    quizArea.innerHTML = '';
    startFlow();
  });

  // file fallback
  fileInput.addEventListener('change', async (ev)=>{
    const f = ev.target.files[0];
    if (!f) return;
    const ab = await f.arrayBuffer();
    const wb2 = XLSX.read(ab, {type:'array'});
    state.allQuestions = parseSheetToQuestions(wb2);
    populateSubjects();
  });

  // keyboard accessibility: focus first select
  subjectSelect.focus();
}

init();
