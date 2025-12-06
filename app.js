// App config - replace with your own if needed
const SUPABASE_URL = 'https://oxhatfkwsgfxuxevuowd.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im94aGF0Zmt3c2dmeHV4ZXZ1b3dkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTU2NzY1MzksImV4cCI6MjA3MTI1MjUzOX0.bZ1cyQiONCrnjeUURfoMJpY6gfr8YgmYj57Dny__4QE';

const supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// Helpers
const $ = (q) => document.querySelector(q);
const $$ = (q) => Array.from(document.querySelectorAll(q));
const opt = (id, label) => {
  const o = document.createElement('option');
  o.value = id; o.textContent = label; return o;
};
const pad2 = (n) => (n != null ? String(n).padStart(2,'0') : '');

async function notify(msg, ok=true){
  console.log(msg);
  const bg = ok ? '#16a34a' /* green-600 */ : '#dc2626' /* red-600 */;
  let el = document.getElementById('toast');
  if (!el){ el = document.createElement('div'); el.id = 'toast'; document.body.appendChild(el); }
  // Centered toast styles
  Object.assign(el.style, {
    position: 'fixed',
    top: '50%',
    left: '50%',
    transform: 'translate(-50%, -50%)',
    zIndex: '9999',
    padding: '10px 16px',
    borderRadius: '8px',
    color: '#fff',
    backgroundColor: bg,
    boxShadow: '0 4px 12px rgba(0,0,0,0.2)',
    opacity: '1',
    transition: 'opacity 0.3s ease',
    textAlign: 'center',
    maxWidth: '90vw'
  });
  el.textContent = msg;
  try { clearTimeout(window.__toastTimeout); } catch {}
  window.__toastTimeout = setTimeout(()=>{ el.style.opacity = '0'; }, 1000);
  if(!ok){
    console.error(msg);
    // Handle expired session tokens
    const m = String(msg || '').toLowerCase();
    if (m.includes('jwt expired') || m.includes('token') && m.includes('expired')){
      try {
        const { data, error } = await supabase.auth.refreshSession();
        if (!error && data?.session){
          el.textContent = 'បានធ្វើឱ្យសម័យឡើងវិញ។ សូមព្យាយាមម្តងទៀត។';
          el.style.backgroundColor = '#16a34a';
          return;
        }
      } catch {}
      try { await supabase.auth.signOut(); } catch {}
      try { showAuth(); } catch {}
      el.textContent = 'សម័យបានផុតកំណត់។ សូមចូលគណនីម្តងទៀត។';
      el.style.backgroundColor = '#dc2626';
    }
    try { alert(msg); } catch {}
  }
}

// FINAL: Average(SPECIAL GR1, SPECIAL GR2)
// - Per-subject cells and Total are aggregated across Group1+SM1+Group2+SM2 exams
// - FINAL Average = mean of SPECIAL GR1 avg and SPECIAL GR2 avg
// - If only one special exists, use that; if none, show no data
async function runFinal(){
  const classSel = document.querySelector('#final-class');
  const tbody = document.querySelector('#table-final tbody');
  const thead = document.querySelector('#table-final thead');
  if (!classSel || !tbody || !thead) return;
  const classId = classSel.value;
  tbody.innerHTML = '';
  if (!classId){ thead.innerHTML = '<tr><th class="p-2 text-left">Student</th></tr>'; return; }
  // fetch exam IDs per group
  const [g1, sm1, g2, sm2] = await Promise.all([
    supabase.from('exams').select('id').eq('class_id', classId).eq('group_name', 'Group 1'),
    supabase.from('exams').select('id').eq('class_id', classId).eq('group_name', 'SM1'),
    supabase.from('exams').select('id').eq('class_id', classId).eq('group_name', 'Group 2'),
    supabase.from('exams').select('id').eq('class_id', classId).eq('group_name', 'SM2')
  ]);
  if (g1.error) return notify(g1.error.message, false);
  if (sm1.error) return notify(sm1.error.message, false);
  if (g2.error) return notify(g2.error.message, false);
  if (sm2.error) return notify(sm2.error.message, false);
  const idsG1 = (g1.data||[]).map(e=>e.id);
  const idsSM1 = (sm1.data||[]).map(e=>e.id);
  const idsG2 = (g2.data||[]).map(e=>e.id);
  const idsSM2 = (sm2.data||[]).map(e=>e.id);
  const haveAny = idsG1.length || idsSM1.length || idsG2.length || idsSM2.length;
  if (!haveAny){
    thead.innerHTML = '<tr><th class="p-2 text-left">Student</th></tr>';
    const tr = document.createElement('tr'); tr.innerHTML = '<td class="p-2 border">គ្មានការប្រឡងសម្រាប់ក្រុមទី១/ឆមាសទី១/ក្រុមទី២/ឆមាសទី២ទេ</td>'; tbody.appendChild(tr); return;
  }
  // subjects and students
  const [{ data: subjects, error: eSub }, { data: students, error: eStu }] = await Promise.all([
    supabase.from('subjects').select('id, name, max_score, display_no').order('display_no', { ascending: true }).order('id', { ascending: true }),
    supabase.from('students').select('id, display_no, first_name, last_name').eq('class_id', classId).order('display_no', { ascending: true }).order('id', { ascending: true })
  ]);
  if (eSub) return notify(eSub.message, false);
  if (eStu) return notify(eStu.message, false);
  const subs = subjects || [];
  const totalValue = subs.reduce((sum, s)=> sum + (s.max_score ? Number(s.max_score)/50 : 0), 0);
  // scores: union for per-subject totals, plus per-group exam totals for averages
  const unionIds = [...new Set([...idsG1, ...idsSM1, ...idsG2, ...idsSM2])];
  const [resAll, resG1, resSM1, resG2, resSM2] = await Promise.all([
    unionIds.length ? supabase.from('scores').select('student_id, subject_id, exam_id, score').in('exam_id', unionIds) : Promise.resolve({ data: [], error: null }),
    idsG1.length ? supabase.from('scores').select('student_id, exam_id, score').in('exam_id', idsG1) : Promise.resolve({ data: [], error: null }),
    idsSM1.length ? supabase.from('scores').select('student_id, exam_id, score').in('exam_id', idsSM1) : Promise.resolve({ data: [], error: null }),
    idsG2.length ? supabase.from('scores').select('student_id, exam_id, score').in('exam_id', idsG2) : Promise.resolve({ data: [], error: null }),
    idsSM2.length ? supabase.from('scores').select('student_id, exam_id, score').in('exam_id', idsSM2) : Promise.resolve({ data: [], error: null })
  ]);
  if (resAll.error) return notify(resAll.error.message, false);
  if (resG1.error) return notify(resG1.error.message, false);
  if (resSM1.error) return notify(resSM1.error.message, false);
  if (resG2.error) return notify(resG2.error.message, false);
  if (resSM2.error) return notify(resSM2.error.message, false);
  const allScores = resAll.data || [];
  const g1Scores = resG1.data || [];
  const sm1Scores = resSM1.data || [];
  const g2Scores = resG2.data || [];
  const sm2Scores = resSM2.data || [];
  // per-subject totals by group for SPECIAL computations
  const setG1 = new Set(idsG1);
  const setSM1 = new Set(idsSM1);
  const setG2 = new Set(idsG2);
  const setSM2 = new Set(idsSM2);
  const stuSubTotalsG1 = new Map(); // `${stu}-${sub}` => sum across G1
  const stuSubTotalsSM1 = new Map(); // `${stu}-${sub}` => sum across SM1
  const stuSubTotalsG2 = new Map(); // `${stu}-${sub}` => sum across G2
  const stuSubTotalsSM2 = new Map(); // `${stu}-${sub}` => sum across SM2
  allScores.forEach(r=>{
    const key = `${r.student_id}-${r.subject_id}`;
    if (setG1.has(r.exam_id)){
      stuSubTotalsG1.set(key, (stuSubTotalsG1.get(key) || 0) + Number(r.score || 0));
    }
    if (setSM1.has(r.exam_id)){
      stuSubTotalsSM1.set(key, (stuSubTotalsSM1.get(key) || 0) + Number(r.score || 0));
    }
    if (setG2.has(r.exam_id)){
      stuSubTotalsG2.set(key, (stuSubTotalsG2.get(key) || 0) + Number(r.score || 0));
    }
    if (setSM2.has(r.exam_id)){
      stuSubTotalsSM2.set(key, (stuSubTotalsSM2.get(key) || 0) + Number(r.score || 0));
    }
  });
  // per-exam totals for each group
  const mapTotals = (rows)=>{
    const m = new Map();
    rows.forEach(r=>{ const k = `${r.student_id}-${r.exam_id}`; m.set(k, (m.get(k)||0)+Number(r.score||0)); });
    return m;
  };
  const totG1 = mapTotals(g1Scores);
  const totSM1 = mapTotals(sm1Scores);
  const totG2 = mapTotals(g2Scores);
  const totSM2 = mapTotals(sm2Scores);
  // header
  const headCells = ['<th class="p-2 text-left">Student</th>']
    .concat(subs.map(s=>`<th class=\"p-2 text-left\">${s.name}</th>`))
    .concat(['<th class=\"p-2 text-left\">Total</th>','<th class=\"p-2 text-left\">Final Average</th>','<th class=\"p-2 text-left\">Rank</th>','<th class=\"p-2 text-left\">Grade</th>']);
  thead.innerHTML = `<tr>${headCells.join('')}</tr>`;
  // helper to compute normalized average for a set of exam IDs using provided totals map
  const avgFor = (examIds, totalsMap, stuId)=>{
    if (!examIds.length || totalValue <= 0) return null;
    let sumNorm = 0; let n = 0;
    examIds.forEach(exId=>{ const t = totalsMap.get(`${stuId}-${exId}`); if (t != null){ sumNorm += (t/totalValue); n++; } });
    return n ? (sumNorm / n) : null;
  };
  // rows
  const rows = students.map(stu=>{
    const examsCountG1 = idsG1.length;
    const examsCountSM1 = idsSM1.length;
    const examsCountG2 = idsG2.length;
    const examsCountSM2 = idsSM2.length;
    // per-subject: SPECIAL1 = mean(AVG G1, AVG SM1), SPECIAL2 = mean(AVG G2, AVG SM2), FINAL cell = mean(SPECIAL1, SPECIAL2)
    const subjectCells = subs.map(sub=>{
      const key = `${stu.id}-${sub.id}`;
      const sumG1 = stuSubTotalsG1.get(key);
      const sumSM1 = stuSubTotalsSM1.get(key);
      const sumG2 = stuSubTotalsG2.get(key);
      const sumSM2 = stuSubTotalsSM2.get(key);
      const avgSubG1 = (sumG1 != null && examsCountG1) ? (sumG1 / examsCountG1) : null;
      const avgSubSM1 = (sumSM1 != null && examsCountSM1) ? (sumSM1 / examsCountSM1) : null;
      const avgSubG2 = (sumG2 != null && examsCountG2) ? (sumG2 / examsCountG2) : null;
      const avgSubSM2 = (sumSM2 != null && examsCountSM2) ? (sumSM2 / examsCountSM2) : null;
      let special1 = null;
      if (avgSubG1 != null && avgSubSM1 != null) special1 = (avgSubG1 + avgSubSM1) / 2;
      else if (avgSubG1 != null) special1 = avgSubG1;
      else if (avgSubSM1 != null) special1 = avgSubSM1;
      let special2 = null;
      if (avgSubG2 != null && avgSubSM2 != null) special2 = (avgSubG2 + avgSubSM2) / 2;
      else if (avgSubG2 != null) special2 = avgSubG2;
      else if (avgSubSM2 != null) special2 = avgSubSM2;
      let cell = '';
      if (special1 != null && special2 != null) cell = ((special1 + special2) / 2).toFixed(2);
      else if (special1 != null) cell = special1.toFixed(2);
      else if (special2 != null) cell = special2.toFixed(2);
      return cell;
    });
    // totals per group (per-exam average across all subjects), then SPECIAL totals and FINAL total
    let sumTotalsG1 = 0, haveAnyG1 = false;
    idsG1.forEach(exId=>{ const t = totG1.get(`${stu.id}-${exId}`); if (t != null){ sumTotalsG1 += t; haveAnyG1 = true; } });
    const totalAvgG1 = (examsCountG1 && haveAnyG1) ? (sumTotalsG1 / examsCountG1) : null;
    let sumTotalsSM1 = 0, haveAnySM1 = false;
    idsSM1.forEach(exId=>{ const t = totSM1.get(`${stu.id}-${exId}`); if (t != null){ sumTotalsSM1 += t; haveAnySM1 = true; } });
    const totalAvgSM1 = (examsCountSM1 && haveAnySM1) ? (sumTotalsSM1 / examsCountSM1) : null;
    let sumTotalsG2 = 0, haveAnyG2 = false;
    idsG2.forEach(exId=>{ const t = totG2.get(`${stu.id}-${exId}`); if (t != null){ sumTotalsG2 += t; haveAnyG2 = true; } });
    const totalAvgG2 = (examsCountG2 && haveAnyG2) ? (sumTotalsG2 / examsCountG2) : null;
    let sumTotalsSM2 = 0, haveAnySM2 = false;
    idsSM2.forEach(exId=>{ const t = totSM2.get(`${stu.id}-${exId}`); if (t != null){ sumTotalsSM2 += t; haveAnySM2 = true; } });
    const totalAvgSM2 = (examsCountSM2 && haveAnySM2) ? (sumTotalsSM2 / examsCountSM2) : null;
    let totalSpecial1 = null;
    if (totalAvgG1 != null && totalAvgSM1 != null) totalSpecial1 = (totalAvgG1 + totalAvgSM1) / 2;
    else if (totalAvgG1 != null) totalSpecial1 = totalAvgG1;
    else if (totalAvgSM1 != null) totalSpecial1 = totalAvgSM1;
    let totalSpecial2 = null;
    if (totalAvgG2 != null && totalAvgSM2 != null) totalSpecial2 = (totalAvgG2 + totalAvgSM2) / 2;
    else if (totalAvgG2 != null) totalSpecial2 = totalAvgG2;
    else if (totalAvgSM2 != null) totalSpecial2 = totalAvgSM2;
    let totalDisplay = null;
    if (totalSpecial1 != null && totalSpecial2 != null) totalDisplay = (totalSpecial1 + totalSpecial2) / 2;
    else if (totalSpecial1 != null) totalDisplay = totalSpecial1;
    else if (totalSpecial2 != null) totalDisplay = totalSpecial2;
    // normalized averages (SPECIAL1 and SPECIAL2) for FINAL average
    const avgSpecial1 = (()=>{
      const a = avgFor(idsG1, totG1, stu.id);
      const b = avgFor(idsSM1, totSM1, stu.id);
      if (a!=null && b!=null) return (a+b)/2; else return (a!=null ? a : (b!=null ? b : null));
    })();
    const avgSpecial2 = (()=>{
      const a = avgFor(idsG2, totG2, stu.id);
      const b = avgFor(idsSM2, totSM2, stu.id);
      if (a!=null && b!=null) return (a+b)/2; else return (a!=null ? a : (b!=null ? b : null));
    })();
    let avgFinal = null;
    if (avgSpecial1!=null && avgSpecial2!=null) avgFinal = (avgSpecial1 + avgSpecial2) / 2;
    else if (avgSpecial1!=null) avgFinal = avgSpecial1;
    else if (avgSpecial2!=null) avgFinal = avgSpecial2;
    return { stu, subjectCells, totalDisplay, avgFinal };
  });
  // rank and render
  const avgs = rows.map(r=> (r.avgFinal==null? -Infinity : r.avgFinal)).filter(v=>v!==-Infinity).sort((a,b)=> b-a);
  const rankByAvg = new Map(); let prev = null; let currentRank = 0;
  avgs.forEach((t, i)=>{ if (prev===null || t<prev){ currentRank = i+1; prev = t; } if (!rankByAvg.has(t)) rankByAvg.set(t, currentRank); });
  rows.forEach(({stu, subjectCells, totalDisplay, avgFinal})=>{
    const tr = document.createElement('tr');
    const name = `${stu.first_name} ${stu.last_name}`;
    const tdName = `<td class=\"p-2 border\">${name}</td>`;
    const tdSubjects = subjectCells.map(v=>`<td class=\"p-2 border\">${v}</td>`).join('');
    const tdTotal = `<td class=\"p-2 border font-semibold\">${totalDisplay!=null ? totalDisplay.toFixed(2) : ''}</td>`;
    const tdAvg = `<td class=\"p-2 border\">${avgFinal!=null ? avgFinal.toFixed(2) : ''}</td>`;
    let grade = '';
    if (avgFinal != null){
      const a = avgFinal;
      if (a >= 45) grade = 'A';
      else if (a >= 40) grade = 'B';
      else if (a >= 35) grade = 'C';
      else if (a >= 30) grade = 'D';
      else if (a >= 25) grade = 'E';
      else grade = 'F';
    }
    const rank = avgFinal != null ? rankByAvg.get(avgFinal) : null;
    const tdRank = `<td class=\"p-2 border font-semibold\">${rank ?? ''}</td>`;
    const tdGrade = `<td class=\"p-2 border font-semibold\">${avgFinal!=null ? grade : 'គ្មានចំណាត់ថ្នាក់'}</td>`;
    tr.innerHTML = tdName + tdSubjects + tdTotal + tdAvg + tdRank + tdGrade;
    if (avgFinal == null){ tr.classList.add('text-green-600'); tr.classList.remove('text-red-600'); }
    else if (grade === 'F'){ tr.classList.add('text-red-600'); tr.classList.remove('text-green-600'); }
    tbody.appendChild(tr);
  });
}

// SPECIAL GR2: Average( AVG GR2 , SM2 )
// - Per-subject cells and Total are aggregated across BOTH Group 2 and SM2 exams
// - Average is the mean of: AVG over Group 2 exams and AVG over SM2 exams
// - If only one side exists, use that side; if none, show no data
async function runSpecialGr2_old(){
  const classSel = document.querySelector('#specialgr2-class');
  const tbody = document.querySelector('#table-specialgr2 tbody');
  const thead = document.querySelector('#table-specialgr2 thead');
  if (!classSel || !tbody || !thead) return;
  const classId = classSel.value;
  tbody.innerHTML = '';
  if (!classId){ thead.innerHTML = '<tr><th class="p-2 text-left">Student</th></tr>'; return; }
  // fetch exams for both groups
  const [{ data: exG2, error: e1 }, { data: exSM2, error: e2 }] = await Promise.all([
    supabase.from('exams').select('id').eq('class_id', classId).eq('group_name', 'Group 2'),
    supabase.from('exams').select('id').eq('class_id', classId).eq('group_name', 'SM2')
  ]);
  if (e1) return notify(e1.message, false);
  if (e2) return notify(e2.message, false);
  const idsG2 = (exG2 || []).map(e=>e.id);
  const idsSM2 = (exSM2 || []).map(e=>e.id);
  const haveAny = idsG2.length || idsSM2.length;
  if (!haveAny){
    thead.innerHTML = '<tr><th class="p-2 text-left">Student</th></tr>';
    const tr = document.createElement('tr'); tr.innerHTML = '<td class="p-2 border">គ្មានការប្រឡងក្រុមទី២ ឬឆមាសទី២ទេ</td>'; tbody.appendChild(tr); return;
  }
  const [{ data: subjects, error: eSub }, { data: students, error: eStu }] = await Promise.all([
    supabase.from('subjects').select('id, name, max_score, display_no').order('display_no', { ascending: true }).order('id', { ascending: true }),
    supabase.from('students').select('id, display_no, first_name, last_name').eq('class_id', classId).order('display_no', { ascending: true }).order('id', { ascending: true })
  ]);
  if (eSub) return notify(eSub.message, false);
  if (eStu) return notify(eStu.message, false);
  const subs = subjects || [];
  const totalValue = subs.reduce((sum, s)=> sum + (s.max_score ? Number(s.max_score)/50 : 0), 0);
  const unionIds = [...new Set([...idsG2, ...idsSM2])];
  const [resAll, resG2, resSM2] = await Promise.all([
    unionIds.length ? supabase.from('scores').select('student_id, subject_id, exam_id, score').in('exam_id', unionIds) : Promise.resolve({ data: [], error: null }),
    idsG2.length ? supabase.from('scores').select('student_id, exam_id, score').in('exam_id', idsG2) : Promise.resolve({ data: [], error: null }),
    idsSM2.length ? supabase.from('scores').select('student_id, exam_id, score').in('exam_id', idsSM2) : Promise.resolve({ data: [], error: null })
  ]);
  if (resAll.error) return notify(resAll.error.message, false);
  if (resG2.error) return notify(resG2.error.message, false);
  if (resSM2.error) return notify(resSM2.error.message, false);
  const allScores = resAll.data || [];
  const g2Scores = resG2.data || [];
  const sm2Scores = resSM2.data || [];
  const stuSubTotals = new Map();
  (allScores).forEach(r=>{
    const k = `${r.student_id}-${r.subject_id}`;
    stuSubTotals.set(k, (stuSubTotals.get(k) || 0) + Number(r.score || 0));
  });
  const totalsByStuExamG2 = new Map();
  g2Scores.forEach(r=>{
    const k = `${r.student_id}-${r.exam_id}`;
    totalsByStuExamG2.set(k, (totalsByStuExamG2.get(k) || 0) + Number(r.score || 0));
  });
  const totalsByStuExamSM2 = new Map();
  sm2Scores.forEach(r=>{
    const k = `${r.student_id}-${r.exam_id}`;
    totalsByStuExamSM2.set(k, (totalsByStuExamSM2.get(k) || 0) + Number(r.score || 0));
  });
  const headCells = ['<th class="p-2 text-left">Student</th>']
    .concat(subs.map(s=>`<th class=\"p-2 text-left\">${s.name}</th>`))
    .concat(['<th class=\"p-2 text-left\">Total</th>','<th class=\"p-2 text-left\">Average (AVG GR2 & SM2)</th>','<th class=\"p-2 text-left\">Rank</th>','<th class=\"p-2 text-left\">Grade</th>']);
  thead.innerHTML = `<tr>${headCells.join('')}</tr>`;
  const rows = students.map(stu=>{
    let total = 0;
    const subjectCells = subs.map(sub=>{
      const v = stuSubTotals.get(`${stu.id}-${sub.id}`);
      if (v != null) total += v;
      return v != null ? Number(v).toFixed(2) : '';
    });
    let avgG2 = null, avgSM2 = null;
    if (idsG2.length && totalValue > 0){
      let sumNorm = 0; let n = 0;
      idsG2.forEach(exId=>{
        const tot = totalsByStuExamG2.get(`${stu.id}-${exId}`);
        if (tot != null){ sumNorm += (tot / totalValue); n++; }
      });
      avgG2 = n ? (sumNorm / n) : null;
    }
    if (idsSM2.length && totalValue > 0){
      let sumNorm = 0; let n = 0;
      idsSM2.forEach(exId=>{
        const tot = totalsByStuExamSM2.get(`${stu.id}-${exId}`);
        if (tot != null){ sumNorm += (tot / totalValue); n++; }
      });
      avgSM2 = n ? (sumNorm / n) : null;
    }
    let avgFinal = null;
    if (avgG2 != null && avgSM2 != null) avgFinal = (avgG2 + avgSM2) / 2;
    else if (avgG2 != null) avgFinal = avgG2;
    else if (avgSM2 != null) avgFinal = avgSM2;
    return { stu, subjectCells, total, avgFinal };
  });
  const avgs = rows.map(r=> (r.avgFinal==null? -Infinity : r.avgFinal)).filter(v=>v!==-Infinity).sort((a,b)=> b-a);
  const rankByAvg = new Map(); let prev = null; let currentRank = 0;
  avgs.forEach((t, i)=>{ if (prev===null || t<prev){ currentRank = i+1; prev = t; } if (!rankByAvg.has(t)) rankByAvg.set(t, currentRank); });
  rows.forEach(({stu, subjectCells, total, avgFinal})=>{
    const tr = document.createElement('tr');
    const name = `${stu.first_name} ${stu.last_name}`;
    const tdName = `<td class=\"p-2 border\">${name}</td>`;
    const tdSubjects = subjectCells.map(v=>`<td class=\"p-2 border\">${v}</td>`).join('');
    const tdTotal = `<td class=\"p-2 border font-semibold\">${total ? total.toFixed(2) : ''}</td>`;
    const tdAvg = `<td class=\"p-2 border\">${avgFinal!=null ? avgFinal.toFixed(2) : ''}</td>`;
    let grade = '';
    if (avgFinal != null){
      const a = avgFinal;
      if (a >= 45) grade = 'A';
      else if (a >= 40) grade = 'B';
      else if (a >= 35) grade = 'C';
      else if (a >= 30) grade = 'D';
      else if (a >= 25) grade = 'E';
      else grade = 'F';
    }
    const rank = avgFinal != null ? rankByAvg.get(avgFinal) : null;
    const tdRank = `<td class=\"p-2 border font-semibold\">${rank ?? ''}</td>`;
    const tdGrade = `<td class=\"p-2 border font-semibold\">${avgFinal!=null ? grade : 'គ្មានចំណាត់ថ្នាក់'}</td>`;
    tr.innerHTML = tdName + tdSubjects + tdTotal + tdAvg + tdRank + tdGrade;
    if (avgFinal == null){ tr.classList.add('text-green-600'); tr.classList.remove('text-red-600'); }
    else if (grade === 'F'){ tr.classList.add('text-red-600'); tr.classList.remove('text-green-600'); }
    tbody.appendChild(tr);
  });
}

// AVG GR1: per-subject totals across Group 1 exams + total, average, rank, grade
async function runAvgGr1(){
  const classSel = document.querySelector('#avggr1-class');
  const tbody = document.querySelector('#table-avggr1 tbody');
  const thead = document.querySelector('#table-avggr1 thead');
  if (!classSel || !tbody || !thead) return;
  const classId = classSel.value;
  tbody.innerHTML = '';
  // guard
  if (!classId){ thead.innerHTML = '<tr><th class="p-2 text-left">Student</th></tr>'; return; }
  // get exams in Group 1 for class
  const { data: exams, error: eEx } = await supabase
    .from('exams')
    .select('id')
    .eq('class_id', classId)
    .eq('group_name', 'Group 1');
  if (eEx) return notify(eEx.message, false);
  const examIds = (exams || []).map(e=>e.id);
  if (!examIds.length){
    thead.innerHTML = '<tr><th class="p-2 text-left">Student</th></tr>';
    const tr = document.createElement('tr'); tr.innerHTML = '<td class="p-2 border">គ្មានការប្រឡងក្រុមទី១ទេ</td>'; tbody.appendChild(tr); return;
  }
  // subjects
  const { data: subjects, error: eSub } = await supabase
    .from('subjects')
    .select('id, name, max_score, display_no')
    .order('display_no', { ascending: true })
    .order('id', { ascending: true });
  if (eSub) return notify(eSub.message, false);
  const subs = subjects || [];
  const totalValue = subs.reduce((sum, s)=> sum + (s.max_score ? Number(s.max_score)/50 : 0), 0);
  // students
  const { data: students, error: eStu } = await supabase
    .from('students')
    .select('id, display_no, first_name, last_name')
    .eq('class_id', classId)
    .order('display_no', { ascending: true })
    .order('id', { ascending: true });
  if (eStu) return notify(eStu.message, false);
  // scores for these exams
  const { data: scores, error: eSc } = await supabase
    .from('scores')
    .select('student_id, subject_id, exam_id, score')
    .in('exam_id', examIds);
  if (eSc) return notify(eSc.message, false);
  // aggregate per student per subject, and per student per exam
  const stuSubTotals = new Map(); // `${stu}-${sub}` => sum
  const stuExamTotals = new Map(); // `${stu}-${exam}` => sum
  (scores || []).forEach(r=>{
    const k1 = `${r.student_id}-${r.subject_id}`;
    stuSubTotals.set(k1, (stuSubTotals.get(k1) || 0) + Number(r.score || 0));
    const k2 = `${r.student_id}-${r.exam_id}`;
    stuExamTotals.set(k2, (stuExamTotals.get(k2) || 0) + Number(r.score || 0));
  });
  // header
  const headCells = ['<th class="p-2 text-left">Student</th>']
    .concat(subs.map(s=>`<th class="p-2 text-left">${s.name}</th>`))
    .concat(['<th class="p-2 text-left">Total</th>','<th class="p-2 text-left">Average</th>','<th class="p-2 text-left">Rank</th>','<th class="p-2 text-left">Grade</th>']);
  thead.innerHTML = `<tr>${headCells.join('')}</tr>`;
  // compute averages first to rank later
  const rows = students.map(stu=>{
    let total = 0;
    const subjectCells = subs.map(sub=>{
      const key = `${stu.id}-${sub.id}`;
      const vSum = stuSubTotals.get(key);
      if (vSum != null) total += vSum; // keep total as sum across exams
      if (vSum == null) return '';
      const examsCountAll = examIds.length || 1; // avoid divide by 0
      const vAvg = vSum / examsCountAll;
      return Number(vAvg).toFixed(2);
    });
    // divide by number of Group 1 exams
    const examsCount = examIds.length;
    const avg = (examsCount && totalValue > 0) ? (total / (examsCount * totalValue)) : null;
    return { stu, subjectCells, total, examsCount, avg };
  });
  // rank by avg desc
  const avgs = rows.map(r=> (r.avg==null? -Infinity : r.avg)).filter(v=>v!==-Infinity).sort((a,b)=> b-a);
  const rankByAvg = new Map();
  let prev = null; let currentRank = 0;
  avgs.forEach((t, i)=>{ if (prev===null || t<prev){ currentRank = i+1; prev = t; } if (!rankByAvg.has(t)) rankByAvg.set(t, currentRank); });
  // render rows
  rows.forEach(({stu, subjectCells, total, examsCount, avg})=>{
    const tr = document.createElement('tr');
    const name = `${stu.first_name} ${stu.last_name}`;
    const tdName = `<td class="p-2 border">${name}</td>`;
    const tdSubjects = subjectCells.map(v=>`<td class=\"p-2 border\">${v}</td>`).join('');
    const tdTotal = `<td class=\"p-2 border font-semibold\">${examsCount ? (total / examsCount).toFixed(2) : ''}</td>`;
    const tdAvg = `<td class=\"p-2 border\">${avg!=null ? avg.toFixed(2) : ''}</td>`;
    // grade
    let grade = '';
    if (avg != null){
      const a = avg;
      if (a >= 45) grade = 'A';
      else if (a >= 40) grade = 'B';
      else if (a >= 35) grade = 'C';
      else if (a >= 30) grade = 'D';
      else if (a >= 25) grade = 'E';
      else grade = 'F';
    }
    const rank = avg != null ? rankByAvg.get(avg) : null;
    const tdRank = `<td class=\"p-2 border font-semibold\">${rank ?? ''}</td>`;
    const tdGrade = `<td class=\"p-2 border font-semibold\">${avg!=null ? grade : 'គ្មានចំណាត់'}</td>`;
    tr.innerHTML = tdName + tdSubjects + tdTotal + tdAvg + tdRank + tdGrade;
    if (avg == null){ tr.classList.add('text-green-600'); tr.classList.remove('text-red-600'); }
    else if (grade === 'F'){ tr.classList.add('text-red-600'); tr.classList.remove('text-green-600'); }
    tbody.appendChild(tr);
  });
}

async function runAvgGr2(){
  const classSel = document.querySelector('#avggr2-class');
  const tbody = document.querySelector('#table-avggr2 tbody');
  const thead = document.querySelector('#table-avggr2 thead');
  if (!classSel || !tbody || !thead) return;
  const classId = classSel.value;
  tbody.innerHTML = '';
  // guard
  if (!classId){ thead.innerHTML = '<tr><th class="p-2 text-left">Student</th></tr>'; return; }
  // get exams in Group 2 for class
  const { data: exams, error: eEx } = await supabase
    .from('exams')
    .select('id')
    .eq('class_id', classId)
    .eq('group_name', 'Group 2');
  if (eEx) return notify(eEx.message, false);
  const examIds = (exams || []).map(e=>e.id);
  if (!examIds.length){
    thead.innerHTML = '<tr><th class="p-2 text-left">Student</th></tr>';
    const tr = document.createElement('tr'); tr.innerHTML = '<td class="p-2 border">គ្មានការប្រឡងក្រុមទី២ទេ</td>'; tbody.appendChild(tr); return;
  }
  // subjects
  const { data: subjects, error: eSub } = await supabase
    .from('subjects')
    .select('id, name, max_score, display_no')
    .order('display_no', { ascending: true })
    .order('id', { ascending: true });
  if (eSub) return notify(eSub.message, false);
  const subs = subjects || [];
  const totalValue = subs.reduce((sum, s)=> sum + (s.max_score ? Number(s.max_score)/50 : 0), 0);
  // students
  const { data: students, error: eStu } = await supabase
    .from('students')
    .select('id, display_no, first_name, last_name')
    .eq('class_id', classId)
    .order('display_no', { ascending: true })
    .order('id', { ascending: true });
  if (eStu) return notify(eStu.message, false);
  // scores for these exams
  const { data: scores, error: eSc } = await supabase
    .from('scores')
    .select('student_id, subject_id, exam_id, score')
    .in('exam_id', examIds);
  if (eSc) return notify(eSc.message, false);
  // aggregate per student per subject, and per student per exam
  const stuSubTotals = new Map(); // `${stu}-${sub}` => sum
  const stuExamTotals = new Map(); // `${stu}-${exam}` => sum
  (scores || []).forEach(r=>{
    const k1 = `${r.student_id}-${r.subject_id}`;
    stuSubTotals.set(k1, (stuSubTotals.get(k1) || 0) + Number(r.score || 0));
    const k2 = `${r.student_id}-${r.exam_id}`;
    stuExamTotals.set(k2, (stuExamTotals.get(k2) || 0) + Number(r.score || 0));
  });
  // header
  const headCells = ['<th class="p-2 text-left">Student</th>']
    .concat(subs.map(s=>`<th class="p-2 text-left">${s.name}</th>`))
    .concat(['<th class="p-2 text-left">Total</th>','<th class="p-2 text-left">Average</th>','<th class="p-2 text-left">Rank</th>','<th class="p-2 text-left">Grade</th>']);
  thead.innerHTML = `<tr>${headCells.join('')}</tr>`;
  // compute averages first to rank later
  const rows = students.map(stu=>{
    let total = 0;
    const subjectCells = subs.map(sub=>{
      const key = `${stu.id}-${sub.id}`;
      const vSum = stuSubTotals.get(key);
      if (vSum != null) total += vSum; // keep total as sum across exams
      if (vSum == null) return '';
      const examsCountAll = examIds.length || 1; // avoid divide by 0
      const vAvg = vSum / examsCountAll;
      return Number(vAvg).toFixed(2);
    });
    // divide by number of Group 2 exams
    const examsCount = examIds.length;
    const avg = (examsCount && totalValue > 0) ? (total / (examsCount * totalValue)) : null;
    return { stu, subjectCells, total, examsCount, avg };
  });
  // rank by avg desc
  const avgs = rows.map(r=> (r.avg==null? -Infinity : r.avg)).filter(v=>v!==-Infinity).sort((a,b)=> b-a);
  const rankByAvg = new Map();
  let prev = null; let currentRank = 0;
  avgs.forEach((t, i)=>{ if (prev===null || t<prev){ currentRank = i+1; prev = t; } if (!rankByAvg.has(t)) rankByAvg.set(t, currentRank); });
  // render rows
  rows.forEach(({stu, subjectCells, total, examsCount, avg})=>{
    const tr = document.createElement('tr');
    const name = `${stu.first_name} ${stu.last_name}`;
    const tdName = `<td class="p-2 border">${name}</td>`;
    const tdSubjects = subjectCells.map(v=>`<td class=\"p-2 border\">${v}</td>`).join('');
    const tdTotal = `<td class=\"p-2 border font-semibold\">${examsCount ? (total / examsCount).toFixed(2) : ''}</td>`;
    const tdAvg = `<td class=\"p-2 border\">${avg!=null ? avg.toFixed(2) : ''}</td>`;
    // grade
    let grade = '';
    if (avg != null){
      const a = avg;
      if (a >= 45) grade = 'A';
      else if (a >= 40) grade = 'B';
      else if (a >= 35) grade = 'C';
      else if (a >= 30) grade = 'D';
      else if (a >= 25) grade = 'E';
      else grade = 'F';
    }
    const rank = avg != null ? rankByAvg.get(avg) : null;
    const tdRank = `<td class=\"p-2 border font-semibold\">${rank ?? ''}</td>`;
    const tdGrade = `<td class=\"p-2 border font-semibold\">${avg!=null ? grade : 'គ្មានចំណាត់'}</td>`;
    tr.innerHTML = tdName + tdSubjects + tdTotal + tdAvg + tdRank + tdGrade;
    if (avg == null){ tr.classList.add('text-green-600'); tr.classList.remove('text-red-600'); }
    else if (grade === 'F'){ tr.classList.add('text-red-600'); tr.classList.remove('text-green-600'); }
    tbody.appendChild(tr);
  });
}

// SPECIAL GR1: Average( AVG GR1 , SM1 )
// - Per-subject cells and Total are aggregated across BOTH Group 1 and SM1 exams
// - Average is the mean of: AVG over Group 1 exams and AVG over SM1 exams
// - If only one side exists, use that side; if none, show no data
async function runSpecialGr1(){
  const classSel = document.querySelector('#specialgr1-class');
  const tbody = document.querySelector('#table-specialgr1 tbody');
  const thead = document.querySelector('#table-specialgr1 thead');
  if (!classSel || !tbody || !thead) return;
  const classId = classSel.value;
  tbody.innerHTML = '';
  if (!classId){ thead.innerHTML = '<tr><th class="p-2 text-left">Student</th></tr>'; return; }
  // fetch exams for both groups
  const [{ data: exG1, error: e1 }, { data: exSM1, error: e2 }] = await Promise.all([
    supabase.from('exams').select('id').eq('class_id', classId).eq('group_name', 'Group 1'),
    supabase.from('exams').select('id').eq('class_id', classId).eq('group_name', 'SM1')
  ]);
  if (e1) return notify(e1.message, false);
  if (e2) return notify(e2.message, false);
  const idsG1 = (exG1 || []).map(e=>e.id);
  const idsSM1 = (exSM1 || []).map(e=>e.id);
  const haveAny = idsG1.length || idsSM1.length;
  if (!haveAny){
    thead.innerHTML = '<tr><th class="p-2 text-left">Student</th></tr>';
    const tr = document.createElement('tr'); tr.innerHTML = '<td class="p-2 border">គ្មានការប្រឡងក្រុមទី១ ឬឆមាសទី១ទេ</td>'; tbody.appendChild(tr); return;
  }
  // subjects and students
  const [{ data: subjects, error: eSub }, { data: students, error: eStu }] = await Promise.all([
    supabase.from('subjects').select('id, name, max_score, display_no').order('display_no', { ascending: true }).order('id', { ascending: true }),
    supabase.from('students').select('id, display_no, first_name, last_name').eq('class_id', classId).order('display_no', { ascending: true }).order('id', { ascending: true })
  ]);
  if (eSub) return notify(eSub.message, false);
  if (eStu) return notify(eStu.message, false);
  const subs = subjects || [];
  const totalValue = subs.reduce((sum, s)=> sum + (s.max_score ? Number(s.max_score)/50 : 0), 0);
  // load scores for both groups
  const unionIds = [...new Set([...idsG1, ...idsSM1])];
  const [resAll, resG1, resSM1] = await Promise.all([
    unionIds.length ? supabase.from('scores').select('student_id, subject_id, exam_id, score').in('exam_id', unionIds) : Promise.resolve({ data: [], error: null }),
    idsG1.length ? supabase.from('scores').select('student_id, exam_id, score').in('exam_id', idsG1) : Promise.resolve({ data: [], error: null }),
    idsSM1.length ? supabase.from('scores').select('student_id, exam_id, score').in('exam_id', idsSM1) : Promise.resolve({ data: [], error: null })
  ]);
  if (resAll.error) return notify(resAll.error.message, false);
  if (resG1.error) return notify(resG1.error.message, false);
  if (resSM1.error) return notify(resSM1.error.message, false);
  const allScores = resAll.data || [];
  const g1Scores = resG1.data || [];
  const sm1Scores = resSM1.data || [];
  // Build sets for group membership
  const setG1 = new Set(idsG1);
  const setSM1 = new Set(idsSM1);
  // Per-subject totals per group
  const stuSubTotalsG1 = new Map(); // `${stu}-${sub}` => sum across G1 exams
  const stuSubTotalsSM1 = new Map(); // `${stu}-${sub}` => sum across SM1 exams
  (allScores).forEach(r=>{
    const key = `${r.student_id}-${r.subject_id}`;
    if (setG1.has(r.exam_id)){
      stuSubTotalsG1.set(key, (stuSubTotalsG1.get(key) || 0) + Number(r.score || 0));
    }
    if (setSM1.has(r.exam_id)){
      stuSubTotalsSM1.set(key, (stuSubTotalsSM1.get(key) || 0) + Number(r.score || 0));
    }
  });
  // per-exam totals for each group (to compute per-student group averages)
  const totalsByStuExamG1 = new Map();
  g1Scores.forEach(r=>{
    const k = `${r.student_id}-${r.exam_id}`;
    totalsByStuExamG1.set(k, (totalsByStuExamG1.get(k) || 0) + Number(r.score || 0));
  });
  const totalsByStuExamSM1 = new Map();
  sm1Scores.forEach(r=>{
    const k = `${r.student_id}-${r.exam_id}`;
    totalsByStuExamSM1.set(k, (totalsByStuExamSM1.get(k) || 0) + Number(r.score || 0));
  });
  // header
  const headCells = ['<th class="p-2 text-left">Student</th>']
    .concat(subs.map(s=>`<th class=\"p-2 text-left\">${s.name}</th>`))
    .concat(['<th class=\"p-2 text-left\">Total</th>','<th class=\"p-2 text-left\">Average (AVG GR1 & SM1)</th>','<th class=\"p-2 text-left\">Rank</th>','<th class=\"p-2 text-left\">Grade</th>']);
  thead.innerHTML = `<tr>${headCells.join('')}</tr>`;
  // build rows
  const rows = students.map(stu=>{
    const examsCountG1 = idsG1.length;
    const examsCountSM1 = idsSM1.length;
    // per-subject: mean of group averages (AVG GR1 and SM1)
    let totalAvgDisplayG1 = null;
    let totalAvgDisplaySM1 = null;
    const subjectCells = subs.map(sub=>{
      const key = `${stu.id}-${sub.id}`;
      const sumG1 = stuSubTotalsG1.get(key);
      const sumSM1 = stuSubTotalsSM1.get(key);
      const avgSubG1 = (sumG1 != null && examsCountG1) ? (sumG1 / examsCountG1) : null;
      const avgSubSM1 = (sumSM1 != null && examsCountSM1) ? (sumSM1 / examsCountSM1) : null;
      let cell = '';
      if (avgSubG1 != null && avgSubSM1 != null) cell = ((avgSubG1 + avgSubSM1) / 2).toFixed(2);
      else if (avgSubG1 != null) cell = (avgSubG1).toFixed(2);
      else if (avgSubSM1 != null) cell = (avgSubSM1).toFixed(2);
      return cell;
    });
    // totals per group (per-exam average across all subjects)
    let sumTotalsG1 = 0; let haveAnyG1 = false;
    idsG1.forEach(exId=>{ const t = totalsByStuExamG1.get(`${stu.id}-${exId}`); if (t != null){ sumTotalsG1 += t; haveAnyG1 = true; } });
    const totalAvgG1 = (examsCountG1 && haveAnyG1) ? (sumTotalsG1 / examsCountG1) : null;
    let sumTotalsSM1 = 0; let haveAnySM1 = false;
    idsSM1.forEach(exId=>{ const t = totalsByStuExamSM1.get(`${stu.id}-${exId}`); if (t != null){ sumTotalsSM1 += t; haveAnySM1 = true; } });
    const totalAvgSM1 = (examsCountSM1 && haveAnySM1) ? (sumTotalsSM1 / examsCountSM1) : null;
    // displayed total: mean of the two groups' per-exam totals
    let totalDisplay = null;
    if (totalAvgG1 != null && totalAvgSM1 != null) totalDisplay = (totalAvgG1 + totalAvgSM1) / 2;
    else if (totalAvgG1 != null) totalDisplay = totalAvgG1;
    else if (totalAvgSM1 != null) totalDisplay = totalAvgSM1;
    // averages normalized by subject values
    const avgG1 = (totalAvgG1 != null && totalValue > 0) ? (totalAvgG1 / totalValue) : null;
    const avgSM1 = (totalAvgSM1 != null && totalValue > 0) ? (totalAvgSM1 / totalValue) : null;
    let avgFinal = null;
    if (avgG1 != null && avgSM1 != null) avgFinal = (avgG1 + avgSM1) / 2;
    else if (avgG1 != null) avgFinal = avgG1;
    else if (avgSM1 != null) avgFinal = avgSM1;
    return { stu, subjectCells, totalDisplay, avgFinal };
  });
  // rank by final average
  const avgs = rows.map(r=> (r.avgFinal==null? -Infinity : r.avgFinal)).filter(v=>v!==-Infinity).sort((a,b)=> b-a);
  const rankByAvg = new Map(); let prev = null; let currentRank = 0;
  avgs.forEach((t, i)=>{ if (prev===null || t<prev){ currentRank = i+1; prev = t; } if (!rankByAvg.has(t)) rankByAvg.set(t, currentRank); });
  // render
  rows.forEach(({stu, subjectCells, totalDisplay, avgFinal})=>{
    const tr = document.createElement('tr');
    const name = `${stu.first_name} ${stu.last_name}`;
    const tdName = `<td class=\"p-2 border\">${name}</td>`;
    const tdSubjects = subjectCells.map(v=>`<td class=\"p-2 border\">${v}</td>`).join('');
    const tdTotal = `<td class=\"p-2 border font-semibold\">${totalDisplay!=null ? totalDisplay.toFixed(2) : ''}</td>`;
    const tdAvg = `<td class=\"p-2 border\">${avgFinal!=null ? avgFinal.toFixed(2) : ''}</td>`;
    // grade based on avgFinal
    let grade = '';
    if (avgFinal != null){
      const a = avgFinal;
      if (a >= 45) grade = 'A';
      else if (a >= 40) grade = 'B';
      else if (a >= 35) grade = 'C';
      else if (a >= 30) grade = 'D';
      else if (a >= 25) grade = 'E';
      else grade = 'F';
    }
    const rank = avgFinal != null ? rankByAvg.get(avgFinal) : null;
    const tdRank = `<td class=\"p-2 border font-semibold\">${rank ?? ''}</td>`;
    const tdGrade = `<td class=\"p-2 border font-semibold\">${avgFinal!=null ? grade : 'គ្មានចំណាត់ថ្នាក់'}</td>`;
    tr.innerHTML = tdName + tdSubjects + tdTotal + tdAvg + tdRank + tdGrade;
    if (avgFinal == null){ tr.classList.add('text-green-600'); tr.classList.remove('text-red-600'); }
    else if (grade === 'F'){ tr.classList.add('text-red-600'); tr.classList.remove('text-green-600'); }
    tbody.appendChild(tr);
  });
}

async function runSpecialGr2(){
  const classSel = document.querySelector('#specialgr2-class');
  const tbody = document.querySelector('#table-specialgr2 tbody');
  const thead = document.querySelector('#table-specialgr2 thead');
  if (!classSel || !tbody || !thead) return;
  const classId = classSel.value;
  tbody.innerHTML = '';
  if (!classId){ thead.innerHTML = '<tr><th class="p-2 text-left">Student</th></tr>'; return; }
  // fetch exams for both groups
  const [{ data: exG2, error: e1 }, { data: exSM2, error: e2 }] = await Promise.all([
    supabase.from('exams').select('id').eq('class_id', classId).eq('group_name', 'Group 2'),
    supabase.from('exams').select('id').eq('class_id', classId).eq('group_name', 'SM2')
  ]);
  if (e1) return notify(e1.message, false);
  if (e2) return notify(e2.message, false);
  const idsG2 = (exG2 || []).map(e=>e.id);
  const idsSM2 = (exSM2 || []).map(e=>e.id);
  const haveAny = idsG2.length || idsSM2.length;
  if (!haveAny){
    thead.innerHTML = '<tr><th class="p-2 text-left">Student</th></tr>';
    const tr = document.createElement('tr'); tr.innerHTML = '<td class="p-2 border">No Group 2 or SM2 exams</td>'; tbody.appendChild(tr); return;
  }
  // subjects and students
  const [{ data: subjects, error: eSub }, { data: students, error: eStu }] = await Promise.all([
    supabase.from('subjects').select('id, name, max_score, display_no').order('display_no', { ascending: true }).order('id', { ascending: true }),
    supabase.from('students').select('id, display_no, first_name, last_name').eq('class_id', classId).order('display_no', { ascending: true }).order('id', { ascending: true })
  ]);
  if (eSub) return notify(eSub.message, false);
  if (eStu) return notify(eStu.message, false);
  const subs = subjects || [];
  const totalValue = subs.reduce((sum, s)=> sum + (s.max_score ? Number(s.max_score)/50 : 0), 0);
  // load scores for both groups
  const unionIds = [...new Set([...idsG2, ...idsSM2])];
  const [resAll, resG2, resSM2] = await Promise.all([
    unionIds.length ? supabase.from('scores').select('student_id, subject_id, exam_id, score').in('exam_id', unionIds) : Promise.resolve({ data: [], error: null }),
    idsG2.length ? supabase.from('scores').select('student_id, exam_id, score').in('exam_id', idsG2) : Promise.resolve({ data: [], error: null }),
    idsSM2.length ? supabase.from('scores').select('student_id, exam_id, score').in('exam_id', idsSM2) : Promise.resolve({ data: [], error: null })
  ]);
  if (resAll.error) return notify(resAll.error.message, false);
  if (resG2.error) return notify(resG2.error.message, false);
  if (resSM2.error) return notify(resSM2.error.message, false);
  const allScores = resAll.data || [];
  const g2Scores = resG2.data || [];
  const sm2Scores = resSM2.data || [];
  // Build sets for group membership
  const setG2 = new Set(idsG2);
  const setSM2 = new Set(idsSM2);
  // Per-subject totals per group
  const stuSubTotalsG2 = new Map(); // `${stu}-${sub}` => sum across G2 exams
  const stuSubTotalsSM2 = new Map(); // `${stu}-${sub}` => sum across SM2 exams
  (allScores).forEach(r=>{
    const key = `${r.student_id}-${r.subject_id}`;
    if (setG2.has(r.exam_id)){
      stuSubTotalsG2.set(key, (stuSubTotalsG2.get(key) || 0) + Number(r.score || 0));
    }
    if (setSM2.has(r.exam_id)){
      stuSubTotalsSM2.set(key, (stuSubTotalsSM2.get(key) || 0) + Number(r.score || 0));
    }
  });
  // per-exam totals for each group (to compute per-student group averages)
  const totalsByStuExamG2 = new Map();
  g2Scores.forEach(r=>{
    const k = `${r.student_id}-${r.exam_id}`;
    totalsByStuExamG2.set(k, (totalsByStuExamG2.get(k) || 0) + Number(r.score || 0));
  });
  const totalsByStuExamSM2 = new Map();
  sm2Scores.forEach(r=>{
    const k = `${r.student_id}-${r.exam_id}`;
    totalsByStuExamSM2.set(k, (totalsByStuExamSM2.get(k) || 0) + Number(r.score || 0));
  });
  // header
  const headCells = ['<th class="p-2 text-left">Student</th>']
    .concat(subs.map(s=>`<th class=\"p-2 text-left\">${s.name}</th>`))
    .concat(['<th class=\"p-2 text-left\">Total</th>','<th class=\"p-2 text-left\">Average (AVG GR2 & SM2)</th>','<th class=\"p-2 text-left\">Rank</th>','<th class=\"p-2 text-left\">Grade</th>']);
  thead.innerHTML = `<tr>${headCells.join('')}</tr>`;
  // build rows
  const rows = students.map(stu=>{
    const examsCountG2 = idsG2.length;
    const examsCountSM2 = idsSM2.length;
    // per-subject: mean of group averages (AVG GR2 and SM2)
    const subjectCells = subs.map(sub=>{
      const key = `${stu.id}-${sub.id}`;
      const sumG2 = stuSubTotalsG2.get(key);
      const sumSM2 = stuSubTotalsSM2.get(key);
      const avgSubG2 = (sumG2 != null && examsCountG2) ? (sumG2 / examsCountG2) : null;
      const avgSubSM2 = (sumSM2 != null && examsCountSM2) ? (sumSM2 / examsCountSM2) : null;
      let cell = '';
      if (avgSubG2 != null && avgSubSM2 != null) cell = ((avgSubG2 + avgSubSM2) / 2).toFixed(2);
      else if (avgSubG2 != null) cell = (avgSubG2).toFixed(2);
      else if (avgSubSM2 != null) cell = (avgSubSM2).toFixed(2);
      return cell;
    });
    // totals per group (per-exam average across all subjects)
    let sumTotalsG2 = 0; let haveAnyG2 = false;
    idsG2.forEach(exId=>{ const t = totalsByStuExamG2.get(`${stu.id}-${exId}`); if (t != null){ sumTotalsG2 += t; haveAnyG2 = true; } });
    const totalAvgG2 = (examsCountG2 && haveAnyG2) ? (sumTotalsG2 / examsCountG2) : null;
    let sumTotalsSM2 = 0; let haveAnySM2 = false;
    idsSM2.forEach(exId=>{ const t = totalsByStuExamSM2.get(`${stu.id}-${exId}`); if (t != null){ sumTotalsSM2 += t; haveAnySM2 = true; } });
    const totalAvgSM2 = (examsCountSM2 && haveAnySM2) ? (sumTotalsSM2 / examsCountSM2) : null;
    // displayed total: mean of the two groups' per-exam totals
    let totalDisplay = null;
    if (totalAvgG2 != null && totalAvgSM2 != null) totalDisplay = (totalAvgG2 + totalAvgSM2) / 2;
    else if (totalAvgG2 != null) totalDisplay = totalAvgG2;
    else if (totalAvgSM2 != null) totalDisplay = totalAvgSM2;
    // averages normalized by subject values
    const avgG2 = (totalAvgG2 != null && totalValue > 0) ? (totalAvgG2 / totalValue) : null;
    const avgSM2 = (totalAvgSM2 != null && totalValue > 0) ? (totalAvgSM2 / totalValue) : null;
    let avgFinal = null;
    if (avgG2 != null && avgSM2 != null) avgFinal = (avgG2 + avgSM2) / 2;
    else if (avgG2 != null) avgFinal = avgG2;
    else if (avgSM2 != null) avgFinal = avgSM2;
    return { stu, subjectCells, totalDisplay, avgFinal };
  });
  // rank by final average
  const avgs = rows.map(r=> (r.avgFinal==null? -Infinity : r.avgFinal)).filter(v=>v!==-Infinity).sort((a,b)=> b-a);
  const rankByAvg = new Map(); let prev = null; let currentRank = 0;
  avgs.forEach((t, i)=>{ if (prev===null || t<prev){ currentRank = i+1; prev = t; } if (!rankByAvg.has(t)) rankByAvg.set(t, currentRank); });
  // render
  rows.forEach(({stu, subjectCells, totalDisplay, avgFinal})=>{
    const tr = document.createElement('tr');
    const name = `${stu.first_name} ${stu.last_name}`;
    const tdName = `<td class=\"p-2 border\">${name}</td>`;
    const tdSubjects = subjectCells.map(v=>`<td class=\"p-2 border\">${v}</td>`).join('');
    const tdTotal = `<td class=\"p-2 border font-semibold\">${totalDisplay!=null ? totalDisplay.toFixed(2) : ''}</td>`;
    const tdAvg = `<td class=\"p-2 border\">${avgFinal!=null ? avgFinal.toFixed(2) : ''}</td>`;
    // grade based on avgFinal
    let grade = '';
    if (avgFinal != null){
      const a = avgFinal;
      if (a >= 45) grade = 'A';
      else if (a >= 40) grade = 'B';
      else if (a >= 35) grade = 'C';
      else if (a >= 30) grade = 'D';
      else if (a >= 25) grade = 'E';
      else grade = 'F';
    }
    const rank = avgFinal != null ? rankByAvg.get(avgFinal) : null;
    const tdRank = `<td class=\"p-2 border font-semibold\">${rank ?? ''}</td>`;
    const tdGrade = `<td class=\"p-2 border font-semibold\">${avgFinal!=null ? grade : 'គ្មានចំណាត់ថ្នាក់'}</td>`;
    tr.innerHTML = tdName + tdSubjects + tdTotal + tdAvg + tdRank + tdGrade;
    if (avgFinal == null){ tr.classList.add('text-green-600'); tr.classList.remove('text-red-600'); }
    else if (grade === 'F'){ tr.classList.add('text-red-600'); tr.classList.remove('text-green-600'); }
    tbody.appendChild(tr);
  });
}

function updateActionStates(){
  const saveAllBtn = document.querySelector('#btn-save-all-scores');
  const runBtn = document.querySelector('#btn-run-report');
  const scoreExamSel = document.querySelector('#score-exam');
  const reportClassSel = document.querySelector('#report-class');
  const reportExamSel = document.querySelector('#report-exam');

  if (saveAllBtn){
    const gridHasInputs = document.querySelectorAll('#table-scores tbody .score-input').length > 0;
    saveAllBtn.disabled = !gridHasInputs;
    saveAllBtn.classList.toggle('opacity-50', !gridHasInputs);
    saveAllBtn.classList.toggle('cursor-not-allowed', !gridHasInputs);
  }
  if (runBtn){
    const ok = !!(reportClassSel?.value) && !!(reportExamSel?.options.length) && !reportExamSel?.disabled;
    runBtn.disabled = !ok;
    runBtn.classList.toggle('opacity-50', !ok);
    runBtn.classList.toggle('cursor-not-allowed', !ok);
  }
}

// Loaders to populate selects and tables
async function loadClasses() {
  const { data: uData, error: uErr } = await supabase.auth.getUser();
  if (uErr) return notify(uErr.message, false);
  const user_id = uData?.user?.id;
  if (!user_id) return notify('Not signed in', false);
  const { data, error } = await supabase
    .from('classes')
    .select('id, display_no, name, year')
    .eq('user_id', user_id)
    .order('display_no', { ascending: true })
    .order('id', { ascending: true });
  if (error) return notify(error.message, false);
  // table
  const tbody = $('#table-classes tbody');
  tbody.innerHTML = '';
  data.forEach((c, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="p-2 border">${pad2(c.display_no ?? (idx + 1))}</td>
      <td class="p-2 border">${c.name}</td>
      <td class="p-2 border">${c.year}</td>
      <td class="p-2 border text-center space-x-2">
        <button data-id="${c.id}" class="btn-edit-class text-blue-600">Edit</button>
        <button data-id="${c.id}" class="btn-del-class text-red-600">Delete</button>
      </td>`;
    tbody.appendChild(tr);
  });
  // selects
  ['#student-class','#exam-class','#score-class','#report-class','#avggr1-class','#avggr2-class','#specialgr1-class','#specialgr2-class','#final-class'].forEach(sel=>{
    const s = $(sel);
    if(!s) return;
    s.innerHTML = '';
    data.forEach(c=>{
      if (sel === '#student-class'){
        const o = document.createElement('option');
        o.value = c.id; o.textContent = `${c.name} (${c.year})`;
        o.dataset.year = c.year;
        s.appendChild(o);
      } else {
        s.appendChild(opt(c.id, `${c.name} (${c.year})`));
      }
    });
    if (s.options.length && !s.value) s.selectedIndex = 0;
  });
  // Auto-load exams for currently selected classes in filters
  const scoreClassSel = document.querySelector('#score-class');
  const reportClassSel = document.querySelector('#report-class');
  if (scoreClassSel?.value) await loadExamsForClass(scoreClassSel.value, ['#score-exam']);
  if (reportClassSel?.value) await loadExamsForClass(reportClassSel.value, ['#report-exam']);
  document.querySelectorAll('#table-scores tbody .score-input').forEach(colorScoreInput);
  updateActionStates();
}

async function loadStudents() {
  const { data, error } = await supabase
    .from('students')
    .select('id, display_no, first_name, last_name, code, gender, photo_url, dob, class_id, classes(name, year)')
    .order('display_no', { ascending: true })
    .order('id', { ascending: true });
  if (error) return notify(error.message, false);
  const tbody = $('#table-students tbody');
  tbody.innerHTML = '';
  const fmt = (iso)=>{ if(!iso) return ''; const d=new Date(iso); if(isNaN(d)) return ''; return `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}/${d.getFullYear()}`; };
  data.forEach(s => {
    const tr = document.createElement('tr');
    const classYear = s.classes?.year != null ? Number(s.classes.year) : null;
    let ageStr = '';
    if (classYear && s.dob){
      const d = new Date(s.dob);
      if (!isNaN(d)){
        const ref = new Date(classYear, 9, 1);
        let age = ref.getFullYear() - d.getFullYear();
        const beforeBirthday = (ref.getMonth() < d.getMonth()) || (ref.getMonth() === d.getMonth() && ref.getDate() < d.getDate());
        if (beforeBirthday) age -= 1;
        if (Number.isFinite(age) && age >= 0) ageStr = String(age);
      }
    }
    tr.innerHTML = `
      <td class="p-2 border">${pad2(s.display_no ?? s.id)}</td>
      <td class="p-2 border">${s.code || ''}</td>
      <td class="p-2 border">${s.first_name} ${s.last_name}</td>
      <td class="p-2 border">${s.photo_url ? `<img src="${s.photo_url}" alt="" class="w-10 h-10 rounded object-cover" onerror="this.style.display='none'">` : ''}</td>
      <td class="p-2 border">${s.gender || ''}</td>
      <td class="p-2 border">${fmt(s.dob)}</td>
      <td class="p-2 border">${ageStr}</td>
      <td class="p-2 border">${s.classes ? `${s.classes.name} (${s.classes.year})` : ''}</td>
      <td class="p-2 border text-center space-x-2">
        <button data-id="${s.id}" class="btn-edit-student text-blue-600">Edit</button>
        <button data-id="${s.id}" class="btn-del-student text-red-600">Delete</button>
      </td>`;
    tbody.appendChild(tr);
  });
}

async function loadSubjects() {
  const { data, error } = await supabase
    .from('subjects')
    .select('id, display_no, name, max_score')
    .order('display_no', { ascending: true })
    .order('id', { ascending: true });
  if (error) return notify(error.message, false);
  const tbody = $('#table-subjects tbody');
  tbody.innerHTML = '';
  data.forEach((su, idx) => {
    const value = su.max_score ? (Number(su.max_score) / 50) : 0;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="p-2 border">${pad2(su.display_no ?? (idx + 1))}</td>
      <td class="p-2 border">${su.name}</td>
      <td class="p-2 border">${su.max_score ?? ''}</td>
      <td class="p-2 border">${value || value === 0 ? value : ''}</td>
      <td class="p-2 border text-center space-x-2">
        <button data-id="${su.id}" class="btn-edit-subject text-blue-600">Edit</button>
        <button data-id="${su.id}" class="btn-del-subject text-red-600">Delete</button>
      </td>`;
    tbody.appendChild(tr);
  });
}

async function loadExams() {
  const { data: uData, error: uErr } = await supabase.auth.getUser();
  if (uErr) return notify(uErr.message, false);
  const user_id = uData?.user?.id;
  if (!user_id) return notify('Not signed in', false);
  const { data, error } = await supabase
    .from('exams')
    .select('id, display_no, name, month, group_name, class_id, classes(name, year)')
    .eq('user_id', user_id)
    .order('display_no', { ascending: true })
    .order('id', { ascending: true });
  if (error) return notify(error.message, false);
  const tbody = $('#table-exams tbody');
  if (tbody){
    tbody.innerHTML = '';
    data.forEach(ex => {
      const tr = document.createElement('tr');
      const m = ex.month;
      const dm = new Date(m); const dd = String(dm.getDate()).padStart(2,'0'); const MM = String(dm.getMonth()+1).padStart(2,'0'); const yyyy = dm.getFullYear();
      tr.innerHTML = `
        <td class="p-2 border">${pad2(ex.display_no ?? ex.id)}</td>
        <td class="p-2 border">${ex.name}</td>
        <td class="p-2 border">${ex.group_name || ''}</td>
        <td class="p-2 border">${dd}/${MM}/${yyyy}</td>
        <td class="p-2 border">${ex.classes ? `${ex.classes.name} (${ex.classes.year})` : ''}</td>
        <td class="p-2 border text-center space-x-2">
          <button data-id="${ex.id}" class="btn-edit-exam text-blue-600">Edit</button>
          <button data-id="${ex.id}" class="btn-del-exam text-red-600">Delete</button>
        </td>`;
      tbody.appendChild(tr);
    });
  }
  return data;
}

// Populate exam selects filtered by class
async function loadExamsForClass(classId, selectIds){
  const { data: uData, error: uErr } = await supabase.auth.getUser();
  if (uErr) return notify(uErr.message, false);
  const user_id = uData?.user?.id;
  if (!user_id) return notify('Not signed in', false);
  const { data, error } = await supabase
    .from('exams')
    .select('id, display_no, name, month, class_id')
    .eq('user_id', user_id)
    .order('display_no', { ascending: true })
    .order('id', { ascending: true })
    .eq('class_id', classId)
    .order('month', { ascending: false });
  if (error) return notify(error.message, false);
  (selectIds || []).forEach(sel => {
    const s = document.querySelector(sel);
    if (!s) return;
    s.innerHTML = '';
    if (!data || !data.length){
      const placeholder = document.createElement('option');
      placeholder.value = '';
      placeholder.textContent = 'No exams for this class';
      placeholder.disabled = true; placeholder.selected = true;
      s.appendChild(placeholder);
      s.disabled = true;
    } else {
      // Special placeholders for Scores and Reports forms
      if (sel === '#score-exam' || sel === '#report-exam'){
        const ph = document.createElement('option');
        ph.value = '';
        ph.textContent = sel === '#score-exam' ? 'សូមជ្រើសរើសខែប្រឡង' : 'សូមជ្រើសរើសការប្រឡង';
        ph.disabled = false; ph.selected = true;
        s.appendChild(ph);
      }
      data.forEach(ex => s.appendChild(opt(ex.id, `${ex.name} (${new Date(ex.month).toLocaleDateString('en-GB')})`)));
      s.disabled = false;
      // Do not auto-select an exam for score/report selects; leave placeholder selected
      if (sel !== '#score-exam' && sel !== '#report-exam') s.selectedIndex = 0;
    }
  });
  return data;
}

function colorScoreInput(inp){
  const vRaw = (inp.value ?? '').toString().trim();
  const cleaned = vRaw.replace(/,/g, '');
  const num = cleaned === '' ? NaN : parseFloat(cleaned);
  const maxAttr = parseFloat(inp.getAttribute('max'));
  const hasMax = Number.isFinite(maxAttr);
  if (vRaw === '' || (!Number.isFinite(num)) || num === 0){
    inp.style.backgroundColor = '#fee2e2';
  } else if (hasMax && Number.isFinite(num) && num === maxAttr){
    inp.style.backgroundColor = '#dbeafe';
  } else {
    inp.style.backgroundColor = '';
  }
}

async function loadScoresGrid() {
  const scoreClassSel = $('#score-class');
  const scoreExamSel = $('#score-exam');
  if (scoreClassSel && !scoreClassSel.value && scoreClassSel.options.length) scoreClassSel.selectedIndex = 0;
  if (scoreExamSel && (!scoreExamSel.value || scoreExamSel.disabled || scoreExamSel.options.length === 0) && scoreClassSel?.value){
    await loadExamsForClass(scoreClassSel.value, ['#score-exam']);
  }
  if (scoreExamSel && !scoreExamSel.value && scoreExamSel.options.length) scoreExamSel.selectedIndex = 0;

  const classId = scoreClassSel?.value;
  const examId = scoreExamSel?.value;
  if (!classId) return notify('Select a class', false);
  if (scoreExamSel?.disabled) return notify('No exams for selected class. Create one in Exams tab.', false);
  if (!examId) return notify('សូមជ្រើសរើសខែប្រឡង', false);

  // Load all subjects with max_score
  const { data: uData0, error: uErr0 } = await supabase.auth.getUser();
  if (uErr0) return notify(uErr0.message, false);
  const user_id0 = uData0?.user?.id;
  if (!user_id0) return notify('Not signed in', false);
  const { data: subjects, error: esub } = await supabase.from('subjects').select('id, name, max_score').eq('user_id', user_id0).order('name');
  if (esub) return notify(esub.message, false);

  // Load students in class
  const { data: uData1, error: uErr1 } = await supabase.auth.getUser();
  if (uErr1) return notify(uErr1.message, false);
  const user_id1 = uData1?.user?.id;
  if (!user_id1) return notify('Not signed in', false);
  const { data: students, error: e1 } = await supabase
    .from('students')
    .select('id, first_name, last_name')
    .eq('user_id', user_id1)
    .eq('class_id', classId)
    .order('last_name');
  if (e1) return notify(e1.message, false);

  // Load existing scores for exam across all subjects and students in class
  const { data: uData2, error: uErr2 } = await supabase.auth.getUser();
  if (uErr2) return notify(uErr2.message, false);
  const user_id2 = uData2?.user?.id;
  if (!user_id2) return notify('Not signed in', false);
  const { data: existing, error: e2 } = await supabase
    .from('scores')
    .select('id, student_id, subject_id, score')
    .eq('user_id', user_id2)
    .eq('exam_id', examId);
  if (e2) return notify(e2.message, false);
  const map = new Map();
  existing.forEach(r => {
    const key = `${r.student_id}-${r.subject_id}`;
    map.set(key, { id: r.id, score: r.score });
  });

  // Render header: Student + each subject
  const headRow = document.getElementById('scores-header-row');
  headRow.innerHTML = '';
  const thStudent = document.createElement('th'); thStudent.className='p-2 text-left'; thStudent.textContent='Student'; headRow.appendChild(thStudent);
  subjects.forEach(sub => {
    const th = document.createElement('th'); th.className='p-2 text-left'; th.textContent=sub.name; headRow.appendChild(th);
  });

  // Render body
  const tbody = $('#table-scores tbody');
  tbody.innerHTML = '';
  students.forEach(s => {
    const tr = document.createElement('tr');
    tr.appendChild(Object.assign(document.createElement('td'), { className: 'p-2 border', textContent: `${s.first_name} ${s.last_name}` }));
    subjects.forEach(sub => {
      const key = `${s.id}-${sub.id}`;
      const td = document.createElement('td');
      td.className = 'p-2 border';
      const maxForSub = sub.max_score != null ? Number(sub.max_score) : 100;
      td.innerHTML = `<input type="number" step="0.01" min="0" max="${maxForSub}" class="score-input w-24 border rounded px-2 py-1" data-student="${s.id}" data-subject="${sub.id}" value="${map.get(key)?.score ?? ''}" title="Max ${maxForSub}">`;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  updateActionStates();
}

async function saveScore(studentId, examId, subjectId, scoreVal){
  const { data: uData, error: uErr } = await supabase.auth.getUser();
  if (uErr) return notify(uErr.message, false);
  const user_id = uData?.user?.id;
  if (!user_id) return notify('Not signed in', false);
  const { error } = await supabase.from('scores').upsert({
    student_id: studentId,
    exam_id: examId,
    subject_id: subjectId,
    score: scoreVal === '' ? null : Number(scoreVal),
    user_id
  }, { onConflict: 'student_id,exam_id,subject_id' });
  if (error) return notify(error.message, false); else notify('ការរក្សាទុកពិន្ទុបានជោគជ័យ!');
}

async function saveAllScores(){
  const btn = document.querySelector('#btn-save-all-scores');
  const scoreExamSel = $('#score-exam');
  const examId = scoreExamSel?.value;
  if (!examId) return notify('Select an exam first', false);
  const inputs = Array.from(document.querySelectorAll('#table-scores tbody .score-input'));
  if (!inputs.length) return notify('Nothing to save', false);

  // Build payload with validation and clamping per-subject max
  const payload = [];
  for (const inp of inputs){
    const student_id = inp.dataset.student;
    const subject_id = inp.dataset.subject;
    if (!student_id || !subject_id) continue;
    let v = inp.value;
    let score = null;
    const maxAllowed = parseFloat(inp.getAttribute('max')) || 100;
    if (v !== ''){
      const num = parseFloat(v);
      if (Number.isNaN(num)) return notify('Invalid score value detected', false);
      if (num > maxAllowed) return notify(`Score exceeds max (${maxAllowed}) for a subject`, false);
      score = Math.max(0, Math.min(maxAllowed, num));
    }
    payload.push({ student_id, subject_id, exam_id: examId, score });
  }
  if (!payload.length) return notify('Nothing to save', false);

  try{
    btn && (btn.disabled = true);
    const { data: uData, error: uErr } = await supabase.auth.getUser();
    if (uErr) return notify(uErr.message, false);
    const user_id = uData?.user?.id;
    if (!user_id) return notify('Not signed in', false);
    const payloadWithUser = payload.map(r => ({ ...r, user_id }));
    const { error } = await supabase
      .from('scores')
      .upsert(payloadWithUser, { onConflict: 'student_id,exam_id,subject_id' });
    if (error) return notify(error.message, false);
    notify('ការរក្សាទុកបានជោគជ័យ!');
    await loadScoresGrid();
    updateActionStates();
  } catch (e){
    notify(String(e), false);
  } finally {
    btn && (btn.disabled = false);
  }
}

// Report: matrix Students x Subjects (read-only)
async function runReport(){
  const reportClassSel = $('#report-class');
  const reportExamSel = $('#report-exam');
  if (reportClassSel && !reportClassSel.value && reportClassSel.options.length) reportClassSel.selectedIndex = 0;
  if (reportExamSel && !reportExamSel.value && reportExamSel.options.length) reportExamSel.selectedIndex = 0;

  const classId = reportClassSel?.value;
  const examId = reportExamSel?.value;
  if(!classId || !examId){
    const thead = document.querySelector('#table-report thead');
    if (thead) thead.innerHTML = '<tr><th class="p-2 text-left">Student</th></tr>';
    $('#table-report tbody').innerHTML = '';
    return;
  }

  // subjects (columns) with max_score to compute Value (max/50)
  const { data: subjects, error: esub } = await supabase.from('subjects').select('id, name, max_score').order('name');
  if (esub) return notify(esub.message, false);
  const totalValue = (subjects || []).reduce((sum, s)=> sum + (s.max_score ? Number(s.max_score)/50 : 0), 0);
  // students (rows)
  const { data: students, error: estu } = await supabase
    .from('students')
    .select('id, first_name, last_name')
    .eq('class_id', classId)
    .order('last_name');
  if (estu) return notify(estu.message, false);
  // scores map
  const { data: scores, error: esc } = await supabase
    .from('scores')
    .select('student_id, subject_id, score')
    .eq('exam_id', examId);
  if (esc) return notify(esc.message, false);
  const scoreMap = new Map(scores.map(r => [`${r.student_id}-${r.subject_id}`, r.score]));

  // compute per-student totals and ranks (1 = highest total)
  const studentTotals = new Map();
  students.forEach(s => {
    let tot = 0;
    subjects.forEach(sub => {
      const key = `${s.id}-${sub.id}`;
      const val = scoreMap.has(key) ? scoreMap.get(key) : null;
      if (val !== null && val !== undefined) tot += Number(val);
    });
    studentTotals.set(s.id, tot);
  });
  const totalsSorted = Array.from(studentTotals.values()).sort((a,b)=> b - a);
  const rankByTotal = new Map();
  let prev = null; let currentRank = 0;
  totalsSorted.forEach((t, i)=>{
    if (prev === null || t < prev){ currentRank = i + 1; prev = t; }
    if (!rankByTotal.has(t)) rankByTotal.set(t, currentRank);
  });

  // header into thead (no dependency on specific id)
  const thead = document.querySelector('#table-report thead');
  if (thead){
    const cells = ['<th class="p-2 text-left">Student</th>']
      .concat(subjects.map(s=>{
        const max = s.max_score != null ? Number(s.max_score) : 0;
        const maxLabel = Number.isInteger(max) ? String(max) : max.toFixed(2);
        const v = max ? (max/50) : 0;
        return `<th class=\"p-2 text-left\">${s.name}(${maxLabel}:${v.toFixed(2)})</th>`;
      }))
      .concat([`<th class=\"p-2 text-left\">Total (${totalValue.toFixed(2)})</th>`,`<th class=\"p-2 text-left\">Average</th>`,`<th class=\"p-2 text-left\">Rank</th>`,`<th class=\"p-2 text-left\">Grade</th>`]);
    thead.innerHTML = `<tr>${cells.join('')}</tr>`;
  }

  // body
  const tbody = $('#table-report tbody');
  tbody.innerHTML = '';
  if (!students.length || !subjects.length){
    const tr = document.createElement('tr');
    tr.innerHTML = `<td class=\"p-2 border\" colspan=\"${subjects.length+5}\">No data</td>`;
    tbody.appendChild(tr);
    return;
  }
  students.forEach(s => {
    const tr = document.createElement('tr');
    tr.appendChild(Object.assign(document.createElement('td'), { className:'p-2 border', textContent: `${s.first_name} ${s.last_name}` }));
    let total = 0; let count = 0;
    subjects.forEach(sub => {
      const td = document.createElement('td'); td.className='p-2 border';
      const key = `${s.id}-${sub.id}`;
      const val = scoreMap.has(key) ? scoreMap.get(key) : null;
      if (val !== null && val !== undefined){ total += Number(val); count++; }
      td.textContent = (val !== null && val !== undefined)
        ? `${Number(val).toFixed(2)}`
        : '';
      tr.appendChild(td);
    })
    const tdTotal = document.createElement('td'); tdTotal.className='p-2 border font-semibold'; tdTotal.textContent = count ? total.toFixed(2) : '';
    const tdAvg = document.createElement('td'); tdAvg.className='p-2 border';
    const avg = totalValue > 0 ? (total / totalValue).toFixed(2) : '';
    tdAvg.textContent = (count ? avg : '');
    // Grade based on average thresholds
    const tdGrade = document.createElement('td'); tdGrade.className='p-2 border font-semibold';
    let grade = '';
    if (count){
      const avgNum = totalValue > 0 ? (total / totalValue) : null;
      if (avgNum != null){
        if (avgNum >= 45) grade = 'A';
        else if (avgNum >= 40) grade = 'B';
        else if (avgNum >= 35) grade = 'C';
        else if (avgNum >= 30) grade = 'D';
        else if (avgNum >= 25) grade = 'E';
        else grade = 'F';
      }
    }
    tdGrade.textContent = count ? grade : 'គ្មានចំណាត់';
    const tdRank = document.createElement('td'); tdRank.className='p-2 border font-semibold';
    if (count){
      const rank = rankByTotal.get(total);
      tdRank.textContent = (rank ?? '');
      // Red for grade F, normal otherwise
      tr.classList.toggle('text-red-600', grade === 'F');
      tr.classList.remove('text-green-600');
    } else {
      tdRank.textContent = 'គ្មានចំណាត់';
      // Green for no-rank (no scores)
      tr.classList.add('text-green-600');
      tr.classList.remove('text-red-600');
    }
    tr.appendChild(tdTotal);
    tr.appendChild(tdAvg);
    tr.appendChild(tdRank);
    tr.appendChild(tdGrade);
    tbody.appendChild(tr);
  });
}

// Report: Group 1 averages per student (read-only)
async function runReportGroup1(){
  const reportClassSel = $('#report-class');
  const classId = reportClassSel?.value;
  const table = document.querySelector('#table-report-group1 tbody');
  if (!table) return; // table not present
  table.innerHTML = '';
  if (!classId){
    return;
  }
  // exams in Group 1 for this class
  const { data: exams, error: eEx } = await supabase
    .from('exams')
    .select('id')
    .eq('class_id', classId)
    .eq('group_name', 'Group 1');
  if (eEx) return notify(eEx.message, false);
  const examIds = (exams || []).map(x=>x.id);
  if (!examIds.length){
    const tr = document.createElement('tr'); tr.innerHTML = '<td class="p-2 border" colspan="3">No Group 1 exams</td>'; table.appendChild(tr); return;
  }
  // subjects and totalValue like main report
  const { data: subjects, error: eSub } = await supabase.from('subjects').select('id, max_score');
  if (eSub) return notify(eSub.message, false);
  const totalValue = (subjects || []).reduce((sum, s)=> sum + (s.max_score ? Number(s.max_score)/50 : 0), 0);
  // students in class
  const { data: students, error: eStu } = await supabase
    .from('students')
    .select('id, first_name, last_name')
    .eq('class_id', classId)
    .order('display_no', { ascending: true })
    .order('id', { ascending: true });
  if (eStu) return notify(eStu.message, false);
  // scores for these exams across subjects
  const { data: scores, error: eSc } = await supabase
    .from('scores')
    .select('student_id, subject_id, exam_id, score')
    .in('exam_id', examIds);
  if (eSc) return notify(eSc.message, false);
  // build per-student per-exam totals
  const subs = subjects || [];
  const scoresByStuExam = new Map(); // key: `${stu}-${exam}` => total
  (scores || []).forEach(r=>{
    const key = `${r.student_id}-${r.exam_id}`;
    const prev = scoresByStuExam.get(key) || 0;
    scoresByStuExam.set(key, prev + Number(r.score || 0));
  });
  // render rows
  students.forEach(s=>{
    let sumOfNormalized = 0; let examsCount = 0;
    examIds.forEach(exId=>{
      const key = `${s.id}-${exId}`;
      const tot = scoresByStuExam.get(key);
      if (tot != null){
        // normalize using totalValue (same subject weights as main report)
        if (totalValue > 0){ sumOfNormalized += (tot / totalValue); examsCount += 1; }
      }
    });
    const tr = document.createElement('tr');
    const name = `${s.first_name} ${s.last_name}`;
    const avg = examsCount ? (sumOfNormalized / examsCount) : null;
    tr.innerHTML = `
      <td class="p-2 border">${name}</td>
      <td class="p-2 border">${examsCount}</td>
      <td class="p-2 border font-semibold">${avg != null ? avg.toFixed(2) : ''}</td>`;
    table.appendChild(tr);
  });
}

// Auto-run report when class/exam changes (works with current HTML)
function bindReportAuto(){
  const reportClassSel = document.querySelector('#report-class');
  const reportExamSel = document.querySelector('#report-exam');
  reportClassSel?.addEventListener('change', async ()=>{
    const classId = reportClassSel.value;
    $('#table-report tbody').innerHTML = '';
    await loadExamsForClass(classId, ['#report-exam']);
    await runReport();
  });
  reportExamSel?.addEventListener('change', async ()=>{ await runReport(); });
  // AVG GR1 bindings
  const avgClassSel = document.querySelector('#avggr1-class');
  avgClassSel?.addEventListener('change', async ()=>{ await runAvgGr1(); });
  document.querySelector('#btn-run-avggr1')?.addEventListener('click', runAvgGr1);
  // AVG GR2 bindings
  const avg2ClassSel = document.querySelector('#avggr2-class');
  avg2ClassSel?.addEventListener('change', async ()=>{ await runAvgGr2(); });
  document.querySelector('#btn-run-avggr2')?.addEventListener('click', runAvgGr2);
  // SPECIAL GR1 bindings
  const spClassSel = document.querySelector('#specialgr1-class');
  spClassSel?.addEventListener('change', async ()=>{ await runSpecialGr1(); });
  document.querySelector('#btn-run-specialgr1')?.addEventListener('click', runSpecialGr1);
  // SPECIAL GR2 bindings
  const sp2ClassSel = document.querySelector('#specialgr2-class');
  sp2ClassSel?.addEventListener('change', async ()=>{ await runSpecialGr2(); });
  document.querySelector('#btn-run-specialgr2')?.addEventListener('click', runSpecialGr2);
  // FINAL bindings
  const fnClassSel = document.querySelector('#final-class');
  fnClassSel?.addEventListener('change', async ()=>{ await runFinal(); });
  document.querySelector('#btn-run-final')?.addEventListener('click', runFinal);
}

// Event bindings
function bindEvents(){
  // Classes
  $('#form-class').addEventListener('submit', async (e)=>{
    e.preventDefault();
    const form = $('#form-class');
    const btn = form.querySelector('button[type="submit"]');
    const name = $('#class-name').value.trim();
    const year = Number($('#class-year').value);
    const display_no_raw = document.getElementById('class-display')?.value?.trim();
    const display_no = display_no_raw ? Number(display_no_raw) : null;
    if(!name || !year) return notify('Name and year are required', false);
    if (display_no !== null && (!Number.isInteger(display_no) || display_no <= 0)) return notify('ID must be a positive integer', false);
    const editId = form?.dataset?.editId;
    if (display_no !== null){
      const { data: uD0, error: uE0 } = await supabase.auth.getUser(); if (uE0) return notify(uE0.message, false); const u0 = uD0?.user?.id; if (!u0) return notify('Not signed in', false);
      let qC = supabase.from('classes').select('id').eq('user_id', u0).eq('display_no', display_no);
      if (editId) qC = qC.neq('id', editId);
      const { data: exists, error: e1 } = await qC;
      if (e1) return notify(e1.message, false);
      if ((exists || []).length) return notify('ID already exists', false);
    }
    if (editId){
      const { error } = await supabase.from('classes').update({ name, year, display_no }).eq('id', editId);
      if (error) return notify(error.message, false);
      notify('ការរក្សាទុកបានជោគជ័យ!');
    } else {
      const { data: uData, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(uErr.message, false);
      const user_id = uData?.user?.id;
      if (!user_id) return notify('Not signed in', false);
      const { error } = await supabase.from('classes').insert({ name, year, display_no, user_id });
      if (error) return notify(error.message, false);
      notify('ការរក្សាទុកបានជោគជ័យ!');
    }
    form.dataset.editId = '';
    if (btn) btn.textContent = 'Save';
    $('#class-name').value=''; $('#class-year').value=''; const cd=document.getElementById('class-display'); if (cd) cd.value='';
    await loadClasses();
  });
  document.body.addEventListener('click', async (e)=>{
    const editC = e.target.closest('.btn-edit-class');
    if (editC){
      const id = editC.dataset.id;
      const { data, error } = await supabase.from('classes').select('id, display_no, name, year').eq('id', id).maybeSingle();
      if (error) return notify(error.message, false);
      if (data){
        const form = $('#form-class');
        form.dataset.editId = data.id;
        $('#class-name').value = data.name || '';
        $('#class-year').value = data.year || '';
        const cd=document.getElementById('class-display'); if (cd) cd.value = data.display_no ?? '';
        const btn = form.querySelector('button[type="submit"]'); if (btn) btn.textContent = 'Update';
      }
      return;
    }
    const delC = e.target.closest('.btn-del-class');
    if(delC){
      const id = delC.dataset.id;
      const { error } = await supabase.from('classes').delete().eq('id', id);
      if (error) return notify(error.message, false); await loadClasses(); return;
    }
  });

  // Classes: Import from Excel
  const importClassesBtn = document.getElementById('btn-import-classes');
  const importClassesFile = document.getElementById('file-import-classes');
  importClassesBtn?.addEventListener('click', ()=> importClassesFile?.click());
  importClassesFile?.addEventListener('change', async ()=>{
    const file = importClassesFile.files?.[0];
    if (!file) return;
    try{
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (!rows.length) return notify('Empty spreadsheet', false);
      const required = ['ID','Name','Year'];
      const hdr = Object.keys(rows[0] || {});
      const missing = required.filter(h=> !hdr.includes(h));
      if (missing.length){ return notify(`Missing headers: ${missing.join(', ')}`, false); }

      const toUpsert = [];
      for (const r of rows){
        const display_no = r['ID'] !== '' ? Number(r['ID']) : null;
        const name = String(r['Name']||'').trim();
        const yearRaw = r['Year'];
        const year = (yearRaw !== '' && yearRaw != null) ? Number(yearRaw) : null;
        if (!name || !year) continue;
        if (display_no !== null && (!Number.isInteger(display_no) || display_no <= 0)) return notify('ID must be a positive integer', false);
        if (!Number.isInteger(year)) return notify('Year must be an integer', false);
        toUpsert.push({ display_no, name, year });
      }
      if (!toUpsert.length) return notify('No valid rows to import', false);

      const { data: uData, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(uErr.message, false);
      const user_id = uData?.user?.id;
      if (!user_id) return notify('Not signed in', false);
      const withUser = toUpsert.map(r => ({ ...r, user_id }));
      const { error } = await supabase.from('classes').upsert(withUser);
      if (error) return notify(error.message, false);
      notify(`Imported ${toUpsert.length} classes`);
      await loadClasses();
    } catch(e){
      notify(`Import failed: ${String(e)}`, false);
    } finally {
      importClassesFile.value = '';
    }
  });

  // Classes: Export to Excel
  document.getElementById('btn-export-classes')?.addEventListener('click', async ()=>{
    try{
      const { data: uData, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(uErr.message, false);
      const user_id = uData?.user?.id;
      if (!user_id) return notify('Not signed in', false);
      const { data, error } = await supabase
        .from('classes')
        .select('display_no, name, year')
        .eq('user_id', user_id)
        .order('display_no', { ascending: true })
        .order('id', { ascending: true });
      if (error) return notify(error.message, false);
      const rows = (data||[]).map(c=>({ 'ID': c.display_no ?? '', 'Name': c.name ?? '', 'Year': c.year ?? '' }));
      const ws = XLSX.utils.json_to_sheet(rows, { header: ['ID','Name','Year'] });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Classes');
      XLSX.writeFile(wb, 'classes.xlsx');
      notify('Exported classes.xlsx');
    } catch(e){
      notify(`Export failed: ${String(e)}`, false);
    }
  });

  // Classes: Print list
  document.getElementById('btn-print-classes')?.addEventListener('click', async ()=>{
    try{
      const { data: uDataS, error: uErrS } = await supabase.auth.getUser();
      if (uErrS) return notify(uErrS.message, false);
      const user_idS = uDataS?.user?.id;
      if (!user_idS) return notify('Not signed in', false);
      const [{ data: settings }, { data: classes, error: cErr }] = await Promise.all([
        supabase.from('settings').select('school_name').eq('user_id', user_idS).maybeSingle(),
        supabase.from('classes').select('display_no, name, year').order('display_no', { ascending: true }).order('id', { ascending: true })
      ]);
      if (cErr) return notify(cErr.message, false);
      const schoolName = (settings?.school_name || document.getElementById('header-school')?.textContent || '').trim() || 'ឈ្មោះសាលា';
      const rows = (classes||[]).map(c=>({ id: (c.display_no ?? ''), name: c.name ?? '', year: c.year ?? '' }));
      let html = `<!doctype html><html><head><meta charset="utf-8"><title>Print Classes</title>
        <style>
          body { font-family: 'Khmer OS Siemreap', 'Noto Sans Khmer', Arial, sans-serif; }
          .page { padding: 24px; }
          .center { text-align: center; }
          .kh-header { font-family: 'Khmer OS Muol Light', 'Noto Serif Khmer', Georgia, serif; line-height: 1.4; }
          table { width: 100%; border-collapse: collapse; font-size: 12px; }
          th, td { border: 1px solid #000; padding: 4px 6px; }
          th { background: #f3f4f6; text-align: left; }
        </style>
      </head><body>`;
      html += `<div class="page">
        <div class="center kh-header">ព្រះរាជាណាចក្រកម្ពុជា</div>
        <div class="center kh-header">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
        <div class="center kh-header" style="margin:8px 0;">${schoolName}</div>
        <div class="center kh-header" style="margin-bottom:8px;">បញ្ជីថ្នាក់</div>
        ${renderTable(rows)}
      </div>`;
      function renderTable(arr){
        if (!arr || !arr.length) return '';
        const head = `<tr><th>ID</th><th>Name</th><th>Year</th></tr>`;
        let body = '';
        for (let i=0;i<arr.length;i++){
          const r = arr[i];
          body += `<tr><td>${r?.id ?? ''}</td><td>${r?.name ?? ''}</td><td>${r?.year ?? ''}</td></tr>`;
        }
        return `<table><thead>${head}</thead><tbody>${body}</tbody></table>`;
      }
      const w = window.open('', '_blank');
      if (!w) return notify('Popup blocked. Allow popups to print.', false);
      w.document.open(); w.document.write(html); w.document.close();
      w.focus();
      setTimeout(()=>{ try { w.print(); } catch {}; }, 300);
    } catch(e){
      notify(`Print failed: ${String(e)}`, false);
    }
  });

  // Classes: Delete all (with related students and their scores)
  document.getElementById('btn-delete-all-classes')?.addEventListener('click', async ()=>{
    try{
      const proceed = confirm('This will permanently delete ALL classes, their students, and related scores. Continue?');
      if (!proceed) return;
      // Load classes and students to cascade delete scores
      const { data: uData, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(uErr.message, false);
      const user_id = uData?.user?.id;
      if (!user_id) return notify('Not signed in', false);
      const [{ data: classes }, { data: students }] = await Promise.all([
        supabase.from('classes').select('id').eq('user_id', user_id),
        supabase.from('students').select('id, class_id').eq('user_id', user_id)
      ]);
      const classIds = (classes||[]).map(c=>c.id);
      const studentIds = (students||[]).map(s=>s.id);
      // Delete scores by student or exam in these classes
      for (let i = 0; i < studentIds.length; i += 1000){
        const group = studentIds.slice(i, i+1000);
        const { error: ds } = await supabase.from('scores').delete().in('student_id', group);
        if (ds) return notify(ds.message, false);
      }
      // Delete students
      for (let i = 0; i < studentIds.length; i += 1000){
        const group = studentIds.slice(i, i+1000);
        const { error: dstu } = await supabase.from('students').delete().in('id', group);
        if (dstu) return notify(dstu.message, false);
      }
      // Delete exams in those classes (and their scores)
      const { data: exams } = await supabase.from('exams').select('id').in('class_id', classIds).eq('user_id', user_id);
      const examIds = (exams||[]).map(e=>e.id);
      for (let i = 0; i < examIds.length; i += 1000){
        const group = examIds.slice(i, i+1000);
        const { error: ds2 } = await supabase.from('scores').delete().in('exam_id', group);
        if (ds2) return notify(ds2.message, false);
      }
      for (let i = 0; i < examIds.length; i += 1000){
        const group = examIds.slice(i, i+1000);
        const { error: de } = await supabase.from('exams').delete().in('id', group);
        if (de) return notify(de.message, false);
      }
      // Delete classes
      for (let i = 0; i < classIds.length; i += 1000){
        const group = classIds.slice(i, i+1000);
        const { error: dc } = await supabase.from('classes').delete().in('id', group);
        if (dc) return notify(dc.message, false);
      }
      notify('All classes (and related data) deleted');
      await loadClasses();
    } catch(e){
      notify(`Delete failed: ${String(e)}`, false);
    }
  });

  // Students: Print list
  document.getElementById('btn-print-students')?.addEventListener('click', async ()=>{
    try{
      const { data: uDataS, error: uErrS } = await supabase.auth.getUser();
      if (uErrS) return notify(uErrS.message, false);
      const user_idS = uDataS?.user?.id;
      if (!user_idS) return notify('Not signed in', false);
      const classFilter = document.getElementById('student-class')?.value || '';
      // Load settings (per-user) and students (optionally filtered by class)
      const settingsPromise = supabase.from('settings').select('school_name').eq('user_id', user_idS).maybeSingle();
      let studentQuery = supabase
        .from('students')
        .select('display_no, code, first_name, last_name, gender, dob, class_id, classes(name, year)')
        .eq('user_id', user_idS)
        .order('class_id', { ascending: true })
        .order('display_no', { ascending: true });
      if (classFilter) studentQuery = studentQuery.eq('class_id', classFilter);
      const [{ data: settings }, { data: students, error: sErr }] = await Promise.all([
        settingsPromise,
        studentQuery
      ]);
      if (sErr) return notify(sErr.message, false);
      const schoolName = (settings?.school_name || document.getElementById('header-school')?.textContent || '').trim() || 'ឈ្មោះសាលា';
      const fmt = (iso)=>{ if(!iso) return ''; const d = new Date(iso); if(isNaN(d)) return ''; return `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}/${d.getFullYear()}`; };
      const rows = (students||[]).map(s=>({
        id: pad2(s.display_no ?? ''),
        code: s.code || '',
        name: `${s.first_name} ${s.last_name}`,
        gender: s.gender || '',
        dob: fmt(s.dob),
        class: s.classes ? `${s.classes.name} (${s.classes.year})` : ''
      }));

      let html = `<!doctype html><html><head><meta charset="utf-8"><title>Print Students</title>
        <style>
          body { font-family: 'Khmer OS Siemreap', 'Noto Sans Khmer', Arial, sans-serif; }
          .page { padding: 24px; }
          .center { text-align: center; }
          .kh-header { font-family: 'Khmer OS Muol Light', 'Noto Serif Khmer', Georgia, serif; line-height: 1.4; }
          table { width: 100%; border-collapse: collapse; font-size: 12px; }
          th, td { border: 1px solid #000; padding: 4px 6px; }
          th { background: #f3f4f6; text-align: left; }
        </style>
      </head><body>`;
      html += `<div class="page">
        <div class="center kh-header">ព្រះរាជាណាចក្រកម្ពុជា</div>
        <div class="center kh-header">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
        <div class="center kh-header" style="margin:8px 0;">${schoolName}</div>
        <div class="center kh-header" style="margin-bottom:8px;">បញ្ជីសិស្ស</div>
        ${renderTable(rows)}
      </div>`;
      function renderTable(arr){
        if (!arr || !arr.length) return '';
        const head = `<tr><th>ID</th><th>CODE</th><th>NAME</th><th>GENDER</th><th>DOB</th><th>CLASS</th></tr>`;
        let body = '';
        for (let i=0;i<arr.length;i++){
          const r = arr[i];
          body += `<tr><td>${r?.id ?? ''}</td><td>${r?.code ?? ''}</td><td>${r?.name ?? ''}</td><td>${r?.gender ?? ''}</td><td>${r?.dob ?? ''}</td><td>${r?.class ?? ''}</td></tr>`;
        }
        return `<table><thead>${head}</thead><tbody>${body}</tbody></table>`;
      }
      const w = window.open('', '_blank');
      if (!w) return notify('Popup blocked. Allow popups to print.', false);
      w.document.open(); w.document.write(html); w.document.close();
      w.focus();
      setTimeout(()=>{ try { w.print(); } catch {}; }, 300);
    } catch(e){
      notify(`Print failed: ${String(e)}`, false);
    }
  });

  // Students
  $('#form-student').addEventListener('submit', async (e)=>{
    e.preventDefault();
    const form = $('#form-student');
    const btn = form.querySelector('button[type="submit"]');
    const first_name = $('#student-first').value.trim();
    const last_name = $('#student-last').value.trim();
    const code = $('#student-code').value.trim() || null;
    const gender = $('#student-gender') ? ($('#student-gender').value || null) : null;
    const photo_url = $('#student-photo') ? ($('#student-photo').value.trim() || null) : null;
    const class_id = $('#student-class').value || null;
    const dobStr = document.getElementById('student-dob')?.value?.trim();
    let dob = null; { const d = parseDobDDMMYYYY(dobStr); if (d) dob = `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`; }
    const display_no_raw = document.getElementById('student-display')?.value?.trim();
    const display_no = display_no_raw ? Number(display_no_raw) : null;
    if(!first_name || !last_name) return notify('First and last name required', false);
    if (display_no !== null){
      if (!Number.isInteger(display_no) || display_no <= 0) return notify('ID must be a positive integer', false);
      const editId = form?.dataset?.editId;
      const { data: uD1, error: uE1 } = await supabase.auth.getUser(); if (uE1) return notify(uE1.message, false); const u1 = uD1?.user?.id; if (!u1) return notify('Not signed in', false);
      let qS = supabase.from('students').select('id').eq('user_id', u1).eq('display_no', display_no);
      if (editId) qS = qS.neq('id', editId);
      const { data: exists, error: e1 } = await qS;
      if (e1) return notify(e1.message, false);
      if ((exists || []).length) return notify('ID already exists', false);
    }
    const editId = form?.dataset?.editId;
    if (editId){
      const { error } = await supabase.from('students').update({ first_name, last_name, code, gender, photo_url, class_id, display_no, dob }).eq('id', editId);
      if (error) return notify(error.message, false);
      notify('ការរក្សាទុកបានជោគជ័យ!');
    } else {
      const { data: uData, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(uErr.message, false);
      const user_id = uData?.user?.id;
      if (!user_id) return notify('Not signed in', false);
      const { error } = await supabase.from('students').insert({ first_name, last_name, code, gender, photo_url, class_id, display_no, dob, user_id });
      if (error) return notify(error.message, false);
      notify('ការរក្សាទុកបានជោគជ័យ!');
    }
    form.dataset.editId = '';
    if (btn) btn.textContent = 'Save';
    $('#student-first').value=''; $('#student-last').value=''; $('#student-code').value=''; if($('#student-gender')) $('#student-gender').value=''; if($('#student-photo')) $('#student-photo').value=''; const sd=document.getElementById('student-display'); if (sd) sd.value=''; const sDob=document.getElementById('student-dob'); if (sDob) sDob.value=''; const sAge=document.getElementById('student-age'); if (sAge) sAge.value='';
    await loadStudents();
  });
  document.body.addEventListener('click', async (e)=>{
    const edit = e.target.closest('.btn-edit-student');
    if (edit){
      const id = edit.dataset.id;
      const { data, error } = await supabase
        .from('students')
        .select('id, display_no, first_name, last_name, code, gender, photo_url, dob, class_id')
        .eq('id', id)
        .maybeSingle();
      if (error) return notify(error.message, false);
      if (data){
        const form = $('#form-student');
        form.dataset.editId = data.id;
        $('#student-first').value = data.first_name || '';
        $('#student-last').value = data.last_name || '';
        $('#student-code').value = data.code || '';
        if ($('#student-gender')) $('#student-gender').value = data.gender || '';
        if ($('#student-photo')) $('#student-photo').value = data.photo_url || '';
        if ($('#student-class')) $('#student-class').value = data.class_id || '';
        const sdisp=document.getElementById('student-display'); if (sdisp) sdisp.value = data.display_no ?? '';
        // Prefill DOB and age
        const dobInput = document.getElementById('student-dob');
        const ageInput = document.getElementById('student-age');
        if (dobInput){
          const formatDDMMYYYY = (iso)=>{ try{ const d=new Date(iso); const dd=pad2(d.getDate()); const mm=pad2(d.getMonth()+1); const yyyy=d.getFullYear(); return `${dd}/${mm}/${yyyy}`; }catch{ return ''; } };
          dobInput.value = data.dob ? formatDDMMYYYY(data.dob) : '';
          // Update age display based on class selection and DOB
          try { const evt = new Event('input'); dobInput.dispatchEvent(evt); } catch {}
        }
        if (ageInput){ try { const evt2 = new Event('change'); document.getElementById('student-class')?.dispatchEvent(evt2); } catch {} }
        const btn = form.querySelector('button[type="submit"]'); if (btn) btn.textContent = 'Update';
      }
      return;
    }
    const del = e.target.closest('.btn-del-student');
    if(del){
      const id = del.dataset.id;
      const { error } = await supabase.from('students').delete().eq('id', id);
      if (error) return notify(error.message, false); await loadStudents(); return;
    }
  });

  // Subjects
  $('#form-subject').addEventListener('submit', async (e)=>{
    e.preventDefault();
    const form = $('#form-subject');
    const btn = form.querySelector('button[type="submit"]');
    const name = $('#subject-name').value.trim();
    const maxInput = $('#subject-max');
    const max_score_raw = maxInput ? maxInput.value : '';
    const max_score = max_score_raw !== '' ? Number(max_score_raw) : null;
    const display_no_raw = document.getElementById('subject-display')?.value?.trim();
    const display_no = display_no_raw ? Number(display_no_raw) : null;
    if(!name) return notify('Name required', false);
    if(max_score !== null && (Number.isNaN(max_score) || max_score <= 0)) return notify('Max score must be a positive number', false);
    if (display_no !== null && (!Number.isInteger(display_no) || display_no <= 0)) return notify('ID must be a positive integer', false);
    const editId = form?.dataset?.editId;
    if (display_no !== null){
      const { data: uD2, error: uE2 } = await supabase.auth.getUser(); if (uE2) return notify(uE2.message, false); const u2 = uD2?.user?.id; if (!u2) return notify('Not signed in', false);
      let qSub = supabase.from('subjects').select('id').eq('user_id', u2).eq('display_no', display_no);
      if (editId) qSub = qSub.neq('id', editId);
      const { data: exists, error: e1 } = await qSub;
      if (e1) return notify(e1.message, false);
      if ((exists || []).length) return notify('ID already exists', false);
    }
    if (editId){
      const { error } = await supabase.from('subjects').update({ name, max_score, display_no }).eq('id', editId);
      if (error) return notify(error.message, false);
      notify('ការរក្សាទុកបានជោគជ័យ!');
    } else {
      const { data: uData, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(uErr.message, false);
      const user_id = uData?.user?.id;
      if (!user_id) return notify('Not signed in', false);
      const { error } = await supabase.from('subjects').insert({ name, max_score, display_no, user_id });
      if (error) return notify(error.message, false);
      notify('ការរក្សាទុកបានជោគជ័យ!');
    }
    form.dataset.editId = '';
    if (btn) btn.textContent = 'Save';
    $('#subject-name').value='';
    if (maxInput) maxInput.value='';
    const sd=document.getElementById('subject-display'); if (sd) sd.value='';
    const valSpan = document.getElementById('subject-value'); if (valSpan) valSpan.textContent = '0';
    await loadSubjects();
  });
  document.body.addEventListener('click', async (e)=>{
    const edit = e.target.closest('.btn-edit-subject');
    if (edit){
      const id = edit.dataset.id;
      const { data, error } = await supabase
        .from('subjects')
        .select('id, display_no, name, max_score')
        .eq('id', id)
        .maybeSingle();
      if (error) return notify(error.message, false);
      if (data){
        const form = $('#form-subject');
        form.dataset.editId = data.id;
        $('#subject-name').value = data.name || '';
        if ($('#subject-max')) $('#subject-max').value = data.max_score ?? '';
        const sdisp=document.getElementById('subject-display'); if (sdisp) sdisp.value = data.display_no ?? '';
        const btn = form.querySelector('button[type="submit"]'); if (btn) btn.textContent = 'Update';
        const valSpan = document.getElementById('subject-value'); if (valSpan) valSpan.textContent = data.max_score ? (Number(data.max_score)/50) : 0;
      }
      return;
    }
    const del = e.target.closest('.btn-del-subject');
    if(del){
      const id = del.dataset.id;
      const { error } = await supabase.from('subjects').delete().eq('id', id);
      if (error) return notify(error.message, false); await loadSubjects(); return;
    }
  });

  // Exams
  $('#form-exam').addEventListener('submit', async (e)=>{
    e.preventDefault();
    const form = $('#form-exam');
    const btn = form.querySelector('button[type="submit"]');
    const name = $('#exam-name').value.trim();
    const monthInput = $('#exam-month').value;
    const group_name = $('#exam-group') ? ($('#exam-group').value || null) : null;
    const class_id = $('#exam-class').value;
    const display_no_raw = document.getElementById('exam-display')?.value?.trim();
    const display_no = display_no_raw ? Number(display_no_raw) : null;
    if(!name || !monthInput || !class_id) return notify('All fields required', false);
    if (display_no !== null && (!Number.isInteger(display_no) || display_no <= 0)) return notify('ID must be a positive integer', false);
    if (display_no !== null){
      const editId = form?.dataset?.editId;
      const { data: uD3, error: uE3 } = await supabase.auth.getUser(); if (uE3) return notify(uE3.message, false); const u3 = uD3?.user?.id; if (!u3) return notify('Not signed in', false);
      let qE = supabase.from('exams').select('id').eq('user_id', u3).eq('display_no', display_no);
      if (editId) qE = qE.neq('id', editId);
      const { data: exists, error: e1 } = await qE;
      if (e1) return notify(e1.message, false);
      if ((exists || []).length) return notify('ID already exists', false);
    }
    const month = monthInput + '-01';
    const editId = form?.dataset?.editId;
    if (editId){
      const { error } = await supabase.from('exams').update({ name, month, group_name, class_id, display_no }).eq('id', editId);
      if (error) return notify(error.message, false);
      notify('ការរក្សាទុកបានជោគជ័យ!');
    } else {
      const { data: uData, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(uErr.message, false);
      const user_id = uData?.user?.id;
      if (!user_id) return notify('Not signed in', false);
      const { error } = await supabase.from('exams').insert({ name, month, group_name, class_id, display_no, user_id });
      if (error) return notify(error.message, false);
      notify('ការរក្សាទុកបានជោគជ័យ!');
    }
    form.dataset.editId = '';
    if (btn) btn.textContent = 'Save';
    $('#exam-name').value=''; $('#exam-month').value=''; if($('#exam-group')) $('#exam-group').value=''; const ed=document.getElementById('exam-display'); if (ed) ed.value='';
    await loadExams();
    await loadExamsForClass(class_id, ['#score-exam', '#report-exam']);
  });
  document.body.addEventListener('click', async (e)=>{
    const edit = e.target.closest('.btn-edit-exam');
    if (edit){
      const id = edit.dataset.id;
      const { data, error } = await supabase
        .from('exams')
        .select('id, display_no, name, month, group_name, class_id')
        .eq('id', id)
        .maybeSingle();
      if (error) return notify(error.message, false);
      if (data){
        const form = $('#form-exam');
        form.dataset.editId = data.id;
        $('#exam-name').value = data.name || '';
        if ($('#exam-group')) $('#exam-group').value = data.group_name || '';
        if ($('#exam-class')) $('#exam-class').value = data.class_id || '';
        const d = new Date(data.month);
        const mval = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
        $('#exam-month').value = mval;
        const ed=document.getElementById('exam-display'); if (ed) ed.value = data.display_no ?? '';
        const btn = form.querySelector('button[type="submit"]'); if (btn) btn.textContent = 'Update';
      }
      return;
    }
    const del = e.target.closest('.btn-del-exam');
    if(del){
      const id = del.dataset.id;
      const { error } = await supabase.from('exams').delete().eq('id', id);
      if (error) return notify(error.message, false); await loadExams(); return;
    }
  });

  // When class filter changes, update exam selects
  const scoreClassSel = document.querySelector('#score-class');
  const reportClassSel = document.querySelector('#report-class');
  scoreClassSel?.addEventListener('change', async ()=>{
    const classId = scoreClassSel.value;
    await loadExamsForClass(classId, ['#score-exam']);
    await loadScoresGrid();
    updateActionStates();
  });
  document.querySelector('#score-exam')?.addEventListener('change', async ()=>{
    await loadScoresGrid();
    document.querySelectorAll('#table-scores tbody .score-input').forEach(colorScoreInput);
    updateActionStates();
  });
  reportClassSel?.addEventListener('change', async ()=>{
    const classId = reportClassSel.value;
    $('#table-report tbody').innerHTML = '';
    await loadExamsForClass(classId, ['#report-exam']);
    updateActionStates();
  });

  // Header: update current tab name when switching tabs
  function updateHeaderTabFromButton(btn){
    const label = btn?.querySelector('span')?.textContent?.trim() || '';
    const el = document.getElementById('header-tab'); if (el) el.textContent = label;
  }
  document.querySelectorAll('.tab-btn').forEach(b=>{
    b.addEventListener('click', ()=> updateHeaderTabFromButton(b));
  });
  // Initialize once from the active button (if any)
  const activeBtn = document.querySelector('.tab-btn.active');
  if (activeBtn) updateHeaderTabFromButton(activeBtn);

  // Students: auto-calc age
  function parseDobDDMMYYYY(str){
    if (!str) return null;
    const m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!m) return null;
    const d = parseInt(m[1],10), mo = parseInt(m[2],10), y = parseInt(m[3],10);
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    const dt = new Date(y, mo-1, d);
    if (dt.getFullYear() !== y || (dt.getMonth()+1) !== mo || dt.getDate() !== d) return null;
    return dt;
  }
  function selectedStudentClassYear(){
    const sel = document.querySelector('#student-class');
    const optEl = sel?.selectedOptions?.[0];
    const y = optEl?.dataset?.year ? parseInt(optEl.dataset.year,10) : null;
    return Number.isFinite(y) ? y : null;
  }
  function updateStudentAge(){
    const dobStr = document.getElementById('student-dob')?.value?.trim();
    const ageEl = document.getElementById('student-age');
    if (!ageEl) return;
    const dob = parseDobDDMMYYYY(dobStr);
    const classYear = selectedStudentClassYear();
    if (!dob || !classYear){ ageEl.value = ''; return; }
    const ref = new Date(classYear, 9, 1); // 01/10/[year] (month index 9)
    let age = ref.getFullYear() - dob.getFullYear();
    const beforeBirthday = (ref.getMonth() < dob.getMonth()) || (ref.getMonth() === dob.getMonth() && ref.getDate() < dob.getDate());
    if (beforeBirthday) age -= 1;
    ageEl.value = age >= 0 && Number.isFinite(age) ? String(age) : '';
  }
  document.getElementById('student-dob')?.addEventListener('input', updateStudentAge);
  document.getElementById('student-class')?.addEventListener('change', updateStudentAge);

  // Scores
  document.querySelector('#btn-save-all-scores')?.addEventListener('click', async ()=>{
    // Auto-fill empty score fields to 0 before saving
    const inputs = document.querySelectorAll('#table-scores tbody .score-input');
    inputs.forEach(inp => {
      const v = (inp.value ?? '').toString().trim();
      if (v === '') inp.value = '0';
      colorScoreInput(inp);
    });
    updateActionStates();
    await saveAllScores();
  });
  document.querySelector('[data-tab="scores"]')?.addEventListener('click', async ()=>{
    const inputs = document.querySelectorAll('#table-scores tbody .score-input');
    // If grid not rendered but selections exist, load it
    const scoreClassSel = document.querySelector('#score-class');
    const scoreExamSel = document.querySelector('#score-exam');
    if ((!inputs || inputs.length === 0) && scoreClassSel?.value && scoreExamSel?.value && !scoreExamSel?.disabled){
      await loadScoresGrid();
    }
    document.querySelectorAll('#table-scores tbody .score-input').forEach(colorScoreInput);
  });
  document.body.addEventListener('change', (e)=>{
    const inp = e.target.closest('.score-input');
    if (!inp) return;
    try { colorScoreInput(inp); } catch {}
  });

  // Scores: Import from Excel (for selected class & exam)
  const importScoresBtn = document.getElementById('btn-import-scores');
  const importScoresFile = document.getElementById('file-import-scores');
  importScoresBtn?.addEventListener('click', ()=> importScoresFile?.click());
  importScoresFile?.addEventListener('change', async ()=>{
    const file = importScoresFile.files?.[0];
    if (!file) return;
    try{
      const classId = document.getElementById('score-class')?.value;
      const examId = document.getElementById('score-exam')?.value;
      if (!classId) return notify('Select a class first', false);
      if (!examId) return notify('Select an exam first', false);

      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (!rows.length) return notify('Empty spreadsheet', false);

      // Load subjects for current user and build maps (case/space-insensitive)
      const { data: uDataS, error: uErrS } = await supabase.auth.getUser();
      if (uErrS) return notify(uErrS.message, false);
      const user_idS = uDataS?.user?.id;
      if (!user_idS) return notify('Not signed in', false);
      const { data: subjects, error: subjErr } = await supabase.from('subjects').select('id, name, max_score').eq('user_id', user_idS).order('name');
      if (subjErr) return notify(subjErr.message, false);
      const norm = (s)=> String(s ?? '').trim().toLowerCase();
      const subjectByNorm = new Map();
      (subjects||[]).forEach(s=> subjectByNorm.set(norm(s.name), s));

      // Load students in class
      const { data: students, error: stuErr } = await supabase
        .from('students')
        .select('id, display_no, code, first_name, last_name')
        .eq('user_id', user_idS)
        .eq('class_id', classId);
      if (stuErr) return notify(stuErr.message, false);
      const studentById = new Map();
      const studentByCode = new Map();
      const studentByDisp = new Map();
      const studentByName = new Map(); // normalized "first last" -> array of students
      const normName = (f,l)=> `${String(f||'').trim().toLowerCase()} ${String(l||'').trim().toLowerCase()}`.trim();
      (students||[]).forEach(st => {
        studentById.set(String(st.id), st);
        const code = String(st.code ?? '').trim().toLowerCase();
        if (code) studentByCode.set(code, st);
        if (st.display_no != null) studentByDisp.set(Number(st.display_no), st);
        const n = normName(st.first_name, st.last_name);
        if (n){ const arr = studentByName.get(n) || []; arr.push(st); studentByName.set(n, arr); }
      });

      // Expect the exported template: columns: 'Student ID', 'Student Name', then each subject name
      const headers = Object.keys(rows[0] || {});
      if (!headers.includes('Student ID')) return notify("Missing 'Student ID' column. Export first to get the correct template.", false);
      // Map normalized header -> original header
      const headerNormToOrig = new Map();
      headers.forEach(h => headerNormToOrig.set(norm(h), h));
      // Build list of subject columns present in sheet that match existing subjects
      const matchedSubjects = [];
      subjectByNorm.forEach((sub, key)=>{
        if (headerNormToOrig.has(key)){
          matchedSubjects.push({ sub, key, orig: headerNormToOrig.get(key) });
        }
      });
      if (!matchedSubjects.length){
        const otherCols = headers.filter(h=> h !== 'Student ID' && h !== 'Student Name');
        return notify(`No subject columns matched. Sheet columns: ${otherCols.join(', ')}`, false);
      }

      const idHdr = headerNormToOrig.get('student id') || 'Student ID';
      const nameHdr = headerNormToOrig.get('student name') || 'Student Name';
      const codeHdr = headerNormToOrig.get('student code') || headerNormToOrig.get('code');
      const dispHdr = headerNormToOrig.get('display no') || headerNormToOrig.get('id');

      const upserts = [];
      let matchedRowCount = 0;
      const unmatched = [];
      for (const r of rows){
        let sidRaw = String(r[idHdr] ?? '').trim();
        let st = sidRaw ? studentById.get(sidRaw) : undefined;
        if (!st && codeHdr){
          const codeVal = String(r[codeHdr] ?? '').trim().toLowerCase();
          if (codeVal) st = studentByCode.get(codeVal);
        }
        if (!st && dispHdr){
          const dVal = r[dispHdr];
          const dn = (dVal !== '' && dVal != null) ? Number(dVal) : NaN;
          if (!Number.isNaN(dn)) st = studentByDisp.get(dn);
        }
        if (!st && nameHdr){
          const full = String(r[nameHdr] ?? '').trim().toLowerCase();
          const arr = full ? (studentByName.get(full) || []) : [];
          if (arr.length === 1) st = arr[0]; // only if unique name in class
        }
        if (!st){ unmatched.push(sidRaw || String(r[codeHdr]||'') || String(r[dispHdr]||'') || String(r[nameHdr]||'')); continue; }
        const student_id = String(st.id);
        let rowHadAny = false;
        for (const { sub, orig } of matchedSubjects){
          const raw = r[orig];
          if (raw === '' || raw == null){
            upserts.push({ student_id, exam_id: examId, subject_id: sub.id, score: null });
            rowHadAny = true;
            continue;
          }
          const num = Number(raw);
          if (Number.isNaN(num)) return notify(`Invalid score for subject '${sub.name}'`, false);
          const maxAllowed = sub.max_score != null ? Number(sub.max_score) : 100;
          if (num < 0 || num > maxAllowed) return notify(`Score for '${sub.name}' must be between 0 and ${maxAllowed}`, false);
          upserts.push({ student_id, exam_id: examId, subject_id: sub.id, score: num });
          rowHadAny = true;
        }
        if (rowHadAny) matchedRowCount += 1;
      }
      if (!upserts.length){
        const presentStudents = rows.map(r=> String(r['Student ID']||'').trim()).filter(x=> x);
        const matchedStudents = presentStudents.filter(id => studentById.has(id));
        return notify(`No valid scores to import. Matched students: ${matchedStudents.length}/${presentStudents.length}. Matched subjects: ${matchedSubjects.length}.`, false);
      }

      const { data: uData, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(uErr.message, false);
      const user_id = uData?.user?.id;
      if (!user_id) return notify('Not signed in', false);
      const withUser = upserts.map(x => ({ ...x, user_id }));
      const { error } = await supabase.from('scores').upsert(withUser, { onConflict: 'student_id,exam_id,subject_id' });
      if (error) return notify(error.message, false);
      notify(`Imported scores for ${rows.length} students`);
      await loadScoresGrid();
      updateActionStates();
    } catch(e){
      notify(`Import failed: ${String(e)})`, false);
    } finally {
      importScoresFile.value = '';
    }
  });

  // Scores: Export to Excel (matrix Student x Subject for selected class & exam)
  document.getElementById('btn-export-scores')?.addEventListener('click', async ()=>{
    try{
      const classId = document.getElementById('score-class')?.value;
      const examId = document.getElementById('score-exam')?.value;
      if (!classId) return notify('Select a class first', false);
      if (!examId) return notify('Select an exam first', false);

      const { data: uData0, error: uErr0 } = await supabase.auth.getUser();
      if (uErr0) return notify(uErr0.message, false);
      const user_id0 = uData0?.user?.id;
      if (!user_id0) return notify('Not signed in', false);

      const [{ data: subjects }, { data: students }, { data: existing } ] = await Promise.all([
        supabase.from('subjects').select('id, name').eq('user_id', user_id0).order('name'),
        supabase.from('students').select('id, first_name, last_name').eq('user_id', user_id0).eq('class_id', classId).order('last_name'),
        supabase.from('scores').select('student_id, subject_id, score').eq('user_id', user_id0).eq('exam_id', examId)
      ]);

      const header = ['Student ID','Student Name','Student Code','Display No', ...(subjects||[]).map(s=>s.name)];
      const rows = (students||[]).map(st =>{
        const row = { 'Student ID': st.id, 'Student Name': `${st.first_name} ${st.last_name}`, 'Student Code': st.code || '', 'Display No': st.display_no ?? '' };
        (subjects||[]).forEach(sub => { row[sub.name] = ''; });
        return row;
      });
      const scoreMap = new Map();
      (existing||[]).forEach(r => scoreMap.set(`${r.student_id}-${r.subject_id}`, r.score));
      rows.forEach(r => {
        (subjects||[]).forEach(sub => {
          const key = `${r['Student ID']}-${sub.id}`;
          const v = scoreMap.get(key);
          if (v != null) r[sub.name] = v;
        });
      });

      const ws = XLSX.utils.json_to_sheet(rows, { header });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Scores');
      XLSX.writeFile(wb, 'scores.xlsx');
      notify('Exported scores.xlsx');
    } catch(e){
      notify(`Export failed: ${String(e)}`, false);
    }
  });

  // Scores: Print matrix
  document.getElementById('btn-print-scores')?.addEventListener('click', async ()=>{
    try{
      const classId = document.getElementById('score-class')?.value;
      const examId = document.getElementById('score-exam')?.value;
      if (!classId) return notify('Select a class first', false);
      if (!examId) return notify('Select an exam first', false);

      const { data: uData0, error: uErr0 } = await supabase.auth.getUser();
      if (uErr0) return notify(uErr0.message, false);
      const user_id0 = uData0?.user?.id;
      if (!user_id0) return notify('Not signed in', false);

      const [{ data: settings }, { data: subjects }, { data: students }, { data: existing } ] = await Promise.all([
        supabase.from('settings').select('school_name').eq('user_id', user_id0).maybeSingle(),
        supabase.from('subjects').select('id, name').eq('user_id', user_id0).order('name'),
        supabase.from('students').select('id, first_name, last_name').eq('user_id', user_id0).eq('class_id', classId).order('last_name'),
        supabase.from('scores').select('student_id, subject_id, score').eq('user_id', user_id0).eq('exam_id', examId)
      ]);
      const schoolName = (settings?.school_name || document.getElementById('header-school')?.textContent || '').trim() || 'ឈ្មោះសាលា';

      const scoreMap = new Map();
      (existing||[]).forEach(r => scoreMap.set(`${r.student_id}-${r.subject_id}`, r.score));

      let html = `<!doctype html><html><head><meta charset="utf-8"><title>Print Scores</title>
        <style>
          body { font-family: 'Khmer OS Siemreap', 'Noto Sans Khmer', Arial, sans-serif; }
          .page { padding: 24px; }
          .center { text-align: center; }
          .kh-header { font-family: 'Khmer OS Muol Light', 'Noto Serif Khmer', Georgia, serif; line-height: 1.4; }
          table { width: 100%; border-collapse: collapse; font-size: 12px; }
          th, td { border: 1px solid #000; padding: 4px 6px; }
          th { background: #f3f4f6; text-align: left; }
        </style>
      </head><body>`;
      html += `<div class="page">
        <div class="center kh-header">ព្រះរាជាណាចក្រកម្ពុជា</div>
        <div class="center kh-header">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
        <div class="center kh-header" style="margin:8px 0;">${schoolName}</div>
        <div class="center kh-header" style="margin-bottom:8px;">តារាងពិន្ទុ</div>
        ${renderTable()}
      </div>`;

      function renderTable(){
        const headCells = ['Student', ...(subjects||[]).map(s=>s.name)];
        let head = '<tr>' + headCells.map(h=>`<th>${h}</th>`).join('') + '</tr>';
        let body = '';
        for (const st of (students||[])){
          let tds = `<td>${st.first_name} ${st.last_name}</td>`;
          for (const sub of (subjects||[])){
            const key = `${st.id}-${sub.id}`;
            const v = scoreMap.get(key);
            tds += `<td>${v ?? ''}</td>`;
          }
          body += `<tr>${tds}</tr>`;
        }
        return `<table><thead>${head}</thead><tbody>${body}</tbody></table>`;
      }

      const w = window.open('', '_blank');
      if (!w) return notify('Popup blocked. Allow popups to print.', false);
      w.document.open(); w.document.write(html); w.document.close();
      w.focus();
      setTimeout(()=>{ try { w.print(); } catch {}; }, 300);
    } catch(e){
      notify(`Print failed: ${String(e)}`, false);
    }
  });

  // Scores: Delete all for selected exam
  document.getElementById('btn-delete-all-scores')?.addEventListener('click', async ()=>{
    try{
      const examId = document.getElementById('score-exam')?.value;
      if (!examId) return notify('Select an exam first', false);
      const proceed = confirm('This will permanently delete ALL scores for the selected exam. Continue?');
      if (!proceed) return;
      const { data: uData, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(uErr.message, false);
      const user_id = uData?.user?.id;
      if (!user_id) return notify('Not signed in', false);
      const { error } = await supabase.from('scores').delete().eq('user_id', user_id).eq('exam_id', examId);
      if (error) return notify(error.message, false);
      notify('All scores deleted for the selected exam');
      await loadScoresGrid();
      updateActionStates();
    } catch(e){
      notify(`Delete failed: ${String(e)}`, false);
    }
  });

  // Import students from Excel
  const importBtn = document.getElementById('btn-import-students');
  const importFile = document.getElementById('file-import-students');
  importBtn?.addEventListener('click', ()=> importFile?.click());
  importFile?.addEventListener('change', async ()=>{
    const file = importFile.files?.[0];
    if (!file) return;
    try{
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (!rows.length) return notify('Empty spreadsheet', false);
      const required = ['ID','Code','First Name','Last Name','Gender','DOB','Class','Photo Url'];
      const hdr = Object.keys(rows[0] || {});
      const missing = required.filter(h=> !hdr.includes(h));
      if (missing.length){ return notify(`Missing headers: ${missing.join(', ')}`, false); }

      // Load classes for mapping
      const { data: uDataC, error: uErrC } = await supabase.auth.getUser();
      if (uErrC) return notify(uErrC.message, false);
      const user_idC = uDataC?.user?.id;
      if (!user_idC) return notify('Not signed in', false);
      const { data: classes, error: cErr } = await supabase.from('classes').select('id, display_no, name, year').eq('user_id', user_idC);
      if (cErr) return notify(cErr.message, false);
      const classMap = new Map();
      (classes||[]).forEach(c=>{
        classMap.set(String(c.display_no ?? '').toLowerCase(), c.id);
        classMap.set(`${c.name}`.toLowerCase(), c.id);
        classMap.set(`${c.name} (${c.year})`.toLowerCase(), c.id);
      });

      // Helpers for DOB parsing
      const excelSerialToDate = (num)=>{
        // Excel serial date to JS Date (epoch 1899-12-30)
        const base = new Date(Date.UTC(1899, 11, 30));
        return new Date(base.getTime() + Math.round(Number(num))*86400000);
      };
      const toISODate = (val)=>{
        if (!val && val !== 0) return null;
        let d = null;
        if (val instanceof Date && !isNaN(val)) d = val;
        else if (typeof val === 'number') d = excelSerialToDate(val);
        else if (typeof val === 'string') d = parseDobDDMMYYYY(val);
        if (!d || isNaN(d)) return null;
        return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
      };

      // Build upsert payloads
      const toUpsert = [];
      for (const r of rows){
        const display_no = r['ID'] !== '' ? Number(r['ID']) : null;
        const code = String(r['Code']||'').trim() || null;
        const first_name = String(r['First Name']||'').trim();
        const last_name = String(r['Last Name']||'').trim();
        const gender = String(r['Gender']||'').trim() || null;
        const dob = toISODate(r['DOB']);
        const photo_url = String(r['Photo Url']||'').trim() || null;
        const classText = String(r['Class']||'').trim().toLowerCase();
        const class_id = classMap.get(classText) ?? null;
        if (!first_name || !last_name) continue;
        toUpsert.push({ display_no, code, first_name, last_name, gender, photo_url, class_id, dob });
      }
      if (!toUpsert.length) return notify('No valid rows to import', false);

      // Upsert with user_id for RLS
      const { data: uData, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(uErr.message, false);
      const user_id = uData?.user?.id;
      if (!user_id) return notify('Not signed in', false);
      const withUser = toUpsert.map(r => ({ ...r, user_id }));
      const { error } = await supabase
        .from('students')
        .upsert(withUser);
      if (error) return notify(error.message, false);
      notify(`Imported ${toUpsert.length} students`);
      await loadStudents();
    } catch(e){
      return notify(`Import failed: ${String(e)}`, false);
    } finally {
      importFile.value = '';
    }
  });

  // Export students to Excel
  document.getElementById('btn-export-students')?.addEventListener('click', async ()=>{
    try{
      // Fetch students with related class for year computation
      const { data: uDataS, error: uErrS } = await supabase.auth.getUser();
      if (uErrS) return notify(uErrS.message, false);
      const user_idS = uDataS?.user?.id;
      if (!user_idS) return notify('Not signed in', false);
      const { data, error } = await supabase
        .from('students')
        .select('display_no, code, first_name, last_name, gender, photo_url, dob, classes(name, year)')
        .eq('user_id', user_idS)
        .order('display_no', { ascending: true })
        .order('id', { ascending: true });
      if (error) return notify(error.message, false);
      const fmtDOB = (iso)=>{ if(!iso) return ''; const d = new Date(iso); if(isNaN(d)) return ''; return `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}/${d.getFullYear()}`; };
      const rows = (data||[]).map(s=>({
        'ID': s.display_no ?? '',
        'Code': s.code ?? '',
        'First Name': s.first_name ?? '',
        'Last Name': s.last_name ?? '',
        'Gender': s.gender ?? '',
        'DOB': fmtDOB(s.dob),
        'Class': s.classes ? `${s.classes.name} (${s.classes.year})` : '',
        'Photo Url': s.photo_url ?? ''
      }));
      // Build worksheet/workbook
      const ws = XLSX.utils.json_to_sheet(rows, { header: ['ID','Code','First Name','Last Name','Gender','DOB','Class','Photo Url'] });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Students');
      // Download file
      XLSX.writeFile(wb, 'students.xlsx');
      notify('Exported students.xlsx');
    } catch(e){
      notify(`Export failed: ${String(e)}`, false);
    }
  });

  // Subjects: Import from Excel
  const importSubjectsBtn = document.getElementById('btn-import-subjects');
  const importSubjectsFile = document.getElementById('file-import-subjects');
  importSubjectsBtn?.addEventListener('click', ()=> importSubjectsFile?.click());
  importSubjectsFile?.addEventListener('change', async ()=>{
    const file = importSubjectsFile.files?.[0];
    if (!file) return;
    try{
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (!rows.length) return notify('Empty spreadsheet', false);
      const required = ['ID','Name','Max'];
      const hdr = Object.keys(rows[0] || {});
      const missing = required.filter(h=> !hdr.includes(h));
      if (missing.length){ return notify(`Missing headers: ${missing.join(', ')}`, false); }
      const toUpsert = [];
      for (const r of rows){
        const display_no = r['ID'] !== '' ? Number(r['ID']) : null;
        const name = String(r['Name']||'').trim();
        const max_raw = r['Max'];
        const max_score = (max_raw !== '' && max_raw != null) ? Number(max_raw) : null;
        if (!name) continue;
        if (display_no !== null && (!Number.isInteger(display_no) || display_no <= 0)) return notify('ID must be a positive integer', false);
        if (max_score !== null && (Number.isNaN(max_score) || max_score <= 0)) return notify('Max must be a positive number', false);
        toUpsert.push({ display_no, name, max_score });
      }
      if (!toUpsert.length) return notify('No valid rows to import', false);
      // Include user_id to satisfy RLS on subjects
      const { data: uData, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(uErr.message, false);
      const user_id = uData?.user?.id;
      if (!user_id) return notify('Not signed in', false);
      const withUser = toUpsert.map(r => ({ ...r, user_id }));
      const { error } = await supabase.from('subjects').upsert(withUser);
      if (error) return notify(error.message, false);
      notify(`Imported ${toUpsert.length} subjects`);
      await loadSubjects();
    } catch(e){
      notify(`Import failed: ${String(e)}`, false);
    } finally {
      importSubjectsFile.value = '';
    }
  });

  // Subjects: Export to Excel
  document.getElementById('btn-export-subjects')?.addEventListener('click', async ()=>{
    try{
      const { data, error } = await supabase
        .from('subjects')
        .select('display_no, name, max_score')
        .order('display_no', { ascending: true })
        .order('id', { ascending: true });
      if (error) return notify(error.message, false);
      const rows = (data||[]).map(s=>({
        'ID': s.display_no ?? '',
        'Name': s.name ?? '',
        'Max': s.max_score ?? '',
        'Value': s.max_score ? (Number(s.max_score)/50) : ''
      }));
      const ws = XLSX.utils.json_to_sheet(rows, { header: ['ID','Name','Max','Value'] });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Subjects');
      XLSX.writeFile(wb, 'subjects.xlsx');
      notify('Exported subjects.xlsx');
    } catch(e){
      notify(`Export failed: ${String(e)}`, false);
    }
  });

  // Subjects: Print list
  document.getElementById('btn-print-subjects')?.addEventListener('click', async ()=>{
    try{
      const { data: uDataS, error: uErrS } = await supabase.auth.getUser();
      if (uErrS) return notify(uErrS.message, false);
      const user_idS = uDataS?.user?.id;
      if (!user_idS) return notify('Not signed in', false);
      const [{ data: settings }, { data: subjects, error: sErr }] = await Promise.all([
        supabase.from('settings').select('school_name').eq('user_id', user_idS).maybeSingle(),
        supabase.from('subjects').select('display_no, name, max_score').order('display_no', { ascending: true }).order('id', { ascending: true })
      ]);
      if (sErr) return notify(sErr.message, false);
      const schoolName = (settings?.school_name || document.getElementById('header-school')?.textContent || '').trim() || 'ឈ្មោះសាលា';
      const rows = (subjects||[]).map(s=>({
        id: (s.display_no ?? ''),
        name: s.name ?? '',
        max: s.max_score ?? '',
        value: s.max_score ? (Number(s.max_score)/50).toFixed(2) : ''
      }));
      let html = `<!doctype html><html><head><meta charset="utf-8"><title>Print Subjects</title>
        <style>
          body { font-family: 'Khmer OS Siemreap', 'Noto Sans Khmer', Arial, sans-serif; }
          .page { padding: 24px; }
          .center { text-align: center; }
          .kh-header { font-family: 'Khmer OS Muol Light', 'Noto Serif Khmer', Georgia, serif; line-height: 1.4; }
          table { width: 100%; border-collapse: collapse; font-size: 12px; }
          th, td { border: 1px solid #000; padding: 4px 6px; }
          th { background: #f3f4f6; text-align: left; }
        </style>
      </head><body>`;
      html += `<div class="page">
        <div class="center kh-header">ព្រះរាជាណាចក្រកម្ពុជា</div>
        <div class="center kh-header">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
        <div class="center kh-header" style="margin:8px 0;">${schoolName}</div>
        <div class="center kh-header" style="margin-bottom:8px;">បញ្ជីមុខវិជ្ជា</div>
        ${renderTable(rows)}
      </div>`;
      function renderTable(arr){
        if (!arr || !arr.length) return '';
        const head = `<tr><th>ID</th><th>Name</th><th>Max</th><th>Value</th></tr>`;
        let body = '';
        for (let i=0;i<arr.length;i++){
          const r = arr[i];
          body += `<tr><td>${r?.id ?? ''}</td><td>${r?.name ?? ''}</td><td>${r?.max ?? ''}</td><td>${r?.value ?? ''}</td></tr>`;
        }
        return `<table><thead>${head}</thead><tbody>${body}</tbody></table>`;
      }
      const w = window.open('', '_blank');
      if (!w) return notify('Popup blocked. Allow popups to print.', false);
      w.document.open(); w.document.write(html); w.document.close();
      w.focus();
      setTimeout(()=>{ try { w.print(); } catch {}; }, 300);
    } catch(e){
      notify(`Print failed: ${String(e)}`, false);
    }
  });

  // Subjects: Delete all (with related scores)
  document.getElementById('btn-delete-all-subjects')?.addEventListener('click', async ()=>{
    try{
      const proceed = confirm('This will permanently delete ALL subjects and their scores. Continue?');
      if (!proceed) return;
      const { data: subjects, error: sErr } = await supabase.from('subjects').select('id');
      if (sErr) return notify(sErr.message, false);
      const ids = (subjects||[]).map(s=>s.id);
      if (!ids.length) return notify('No subjects to delete');
      // delete related scores first
      for (let i = 0; i < ids.length; i += 1000) {
        const group = ids.slice(i, i + 1000);
        const { error: delScoresErr } = await supabase.from('scores').delete().in('subject_id', group);
        if (delScoresErr) return notify(delScoresErr.message, false);
      }
      for (let i = 0; i < ids.length; i += 1000) {
        const group = ids.slice(i, i + 1000);
        const { error: delSubErr } = await supabase.from('subjects').delete().in('id', group);
        if (delSubErr) return notify(delSubErr.message, false);
      }
      notify('All subjects deleted');
      await loadSubjects();
    } catch(e){
      notify(`Delete failed: ${String(e)}`, false);
    }
  });

  // Exams: Import from Excel
  const importExamsBtn = document.getElementById('btn-import-exams');
  const importExamsFile = document.getElementById('file-import-exams');
  importExamsBtn?.addEventListener('click', ()=> importExamsFile?.click());
  importExamsFile?.addEventListener('change', async ()=>{
    const file = importExamsFile.files?.[0];
    if (!file) return;
    try{
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (!rows.length) return notify('Empty spreadsheet', false);
      const required = ['ID','Name','Month','Group','Class'];
      const hdr = Object.keys(rows[0] || {});
      const missing = required.filter(h=> !hdr.includes(h));
      if (missing.length){ return notify(`Missing headers: ${missing.join(', ')}`, false); }

      // Load classes for mapping
      const { data: uDataC, error: uErrC } = await supabase.auth.getUser();
      if (uErrC) return notify(uErrC.message, false);
      const user_idC = uDataC?.user?.id;
      if (!user_idC) return notify('Not signed in', false);
      const { data: classes, error: cErr } = await supabase.from('classes').select('id, display_no, name, year').eq('user_id', user_idC);
      if (cErr) return notify(cErr.message, false);
      const classMap = new Map();
      (classes||[]).forEach(c=>{
        classMap.set(String(c.display_no ?? '').toLowerCase(), c.id);
        classMap.set(`${c.name}`.toLowerCase(), c.id);
        classMap.set(`${c.name} (${c.year})`.toLowerCase(), c.id);
      });

      const toUpsert = [];
      const toISOFirstOfMonth = (val)=>{
        if (!val && val !== 0) return null;
        let d = null;
        if (val instanceof Date && !isNaN(val)) d = val;
        else if (typeof val === 'number') { const base = new Date(Date.UTC(1899, 11, 30)); d = new Date(base.getTime() + Math.round(Number(val))*86400000); }
        else if (typeof val === 'string') {
          const s = val.trim();
          if (/^\d{4}-\d{2}$/.test(s)) { const [Y,M] = s.split('-'); d = new Date(Number(Y), Number(M)-1, 1); }
          else {
            const tryD = new Date(s);
            if (!isNaN(tryD)) d = new Date(tryD.getFullYear(), tryD.getMonth(), 1);
          }
        }
        if (!d || isNaN(d)) return null;
        return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-01`;
      };

      for (const r of rows){
        const display_no = r['ID'] !== '' ? Number(r['ID']) : null;
        const name = String(r['Name']||'').trim();
        const group_name = String(r['Group']||'').trim() || null;
        const classText = String(r['Class']||'').trim().toLowerCase();
        const class_id = classMap.get(classText) ?? null;
        const month = toISOFirstOfMonth(r['Month']);
        if (!name || !month || !class_id) continue;
        if (display_no !== null && (!Number.isInteger(display_no) || display_no <= 0)) return notify('ID must be a positive integer', false);
        toUpsert.push({ display_no, name, month, group_name, class_id });
      }
      if (!toUpsert.length) return notify('No valid rows to import', false);

      // Include user_id to satisfy RLS
      const { data: uData, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(uErr.message, false);
      const user_id = uData?.user?.id;
      if (!user_id) return notify('Not signed in', false);
      const withUser = toUpsert.map(r => ({ ...r, user_id }));
      const { error } = await supabase.from('exams').upsert(withUser);
      if (error) return notify(error.message, false);
      notify(`Imported ${toUpsert.length} exams`);
      await loadExams();
    } catch(e){
      notify(`Import failed: ${String(e)}`, false);
    } finally {
      importExamsFile.value = '';
    }
  });

  // Exams: Export to Excel
  document.getElementById('btn-export-exams')?.addEventListener('click', async ()=>{
    try{
      const { data: uData, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(uErr.message, false);
      const user_id = uData?.user?.id;
      if (!user_id) return notify('Not signed in', false);
      const { data, error } = await supabase
        .from('exams')
        .select('display_no, name, month, group_name, classes(name, year)')
        .eq('user_id', user_id)
        .order('display_no', { ascending: true })
        .order('id', { ascending: true });
      if (error) return notify(error.message, false);
      const fmtMonth = (iso)=>{ if(!iso) return ''; const d = new Date(iso); if(isNaN(d)) return ''; return `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}/${d.getFullYear()}`; };
      const rows = (data||[]).map(ex=>({
        'ID': ex.display_no ?? '',
        'Name': ex.name ?? '',
        'Group': ex.group_name ?? '',
        'Month': fmtMonth(ex.month),
        'Class': ex.classes ? `${ex.classes.name} (${ex.classes.year})` : ''
      }));
      const ws = XLSX.utils.json_to_sheet(rows, { header: ['ID','Name','Group','Month','Class'] });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Exams');
      XLSX.writeFile(wb, 'exams.xlsx');
      notify('Exported exams.xlsx');
    } catch(e){
      notify(`Export failed: ${String(e)}`, false);
    }
  });

  // Exams: Print list
  document.getElementById('btn-print-exams')?.addEventListener('click', async ()=>{
    try{
      const { data: uDataS, error: uErrS } = await supabase.auth.getUser();
      if (uErrS) return notify(uErrS.message, false);
      const user_idS = uDataS?.user?.id;
      if (!user_idS) return notify('Not signed in', false);
      const [{ data: settings }, { data: exams, error: exErr }] = await Promise.all([
        supabase.from('settings').select('school_name').eq('user_id', user_idS).maybeSingle(),
        supabase.from('exams').select('display_no, name, month, group_name, classes(name, year)').order('display_no', { ascending: true }).order('id', { ascending: true })
      ]);
      if (exErr) return notify(exErr.message, false);
      const schoolName = (settings?.school_name || document.getElementById('header-school')?.textContent || '').trim() || 'ឈ្មោះសាលា';
      const fmtMonth = (iso)=>{ if(!iso) return ''; const d = new Date(iso); if(isNaN(d)) return ''; return `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}/${d.getFullYear()}`; };
      const rows = (exams||[]).map(ex=>({
        id: (ex.display_no ?? ''),
        name: ex.name ?? '',
        group: ex.group_name ?? '',
        month: fmtMonth(ex.month),
        class: ex.classes ? `${ex.classes.name} (${ex.classes.year})` : ''
      }));
      let html = `<!doctype html><html><head><meta charset="utf-8"><title>Print Exams</title>
        <style>
          body { font-family: 'Khmer OS Siemreap', 'Noto Sans Khmer', Arial, sans-serif; }
          .page { padding: 24px; }
          .center { text-align: center; }
          .kh-header { font-family: 'Khmer OS Muol Light', 'Noto Serif Khmer', Georgia, serif; line-height: 1.4; }
          table { width: 100%; border-collapse: collapse; font-size: 12px; }
          th, td { border: 1px solid #000; padding: 4px 6px; }
          th { background: #f3f4f6; text-align: left; }
        </style>
      </head><body>`;
      html += `<div class="page">
        <div class="center kh-header">ព្រះរាជាណាចក្រកម្ពុជា</div>
        <div class="center kh-header">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
        <div class="center kh-header" style="margin:8px 0;">${schoolName}</div>
        <div class="center kh-header" style="margin-bottom:8px;">បញ្ជីប្រឡង</div>
        ${renderTable(rows)}
      </div>`;
      function renderTable(arr){
        if (!arr || !arr.length) return '';
        const head = `<tr><th>ID</th><th>Name</th><th>Group</th><th>Month</th><th>Class</th></tr>`;
        let body = '';
        for (let i=0;i<arr.length;i++){
          const r = arr[i];
          body += `<tr><td>${r?.id ?? ''}</td><td>${r?.name ?? ''}</td><td>${r?.group ?? ''}</td><td>${r?.month ?? ''}</td><td>${r?.class ?? ''}</td></tr>`;
        }
        return `<table><thead>${head}</thead><tbody>${body}</tbody></table>`;
      }
      const w = window.open('', '_blank');
      if (!w) return notify('Popup blocked. Allow popups to print.', false);
      w.document.open(); w.document.write(html); w.document.close();
      w.focus();
      setTimeout(()=>{ try { w.print(); } catch {}; }, 300);
    } catch(e){
      notify(`Print failed: ${String(e)}`, false);
    }
  });

  // Exams: Delete all (with related scores)
  document.getElementById('btn-delete-all-exams')?.addEventListener('click', async ()=>{
    try{
      const proceed = confirm('This will permanently delete ALL exams and their scores. Continue?');
      if (!proceed) return;
      const { data: exams, error: eErr } = await supabase.from('exams').select('id');
      if (eErr) return notify(eErr.message, false);
      const ids = (exams||[]).map(s=>s.id);
      if (!ids.length) return notify('No exams to delete');
      // delete related scores first
      for (let i = 0; i < ids.length; i += 1000) {
        const group = ids.slice(i, i + 1000);
        const { error: delScoresErr } = await supabase.from('scores').delete().in('exam_id', group);
        if (delScoresErr) return notify(delScoresErr.message, false);
      }
      for (let i = 0; i < ids.length; i += 1000) {
        const group = ids.slice(i, i + 1000);
        const { error: delExErr } = await supabase.from('exams').delete().in('id', group);
        if (delExErr) return notify(delExErr.message, false);
      }
      notify('All exams deleted');
      await loadExams();
    } catch(e){
      notify(`Delete failed: ${String(e)}`, false);
    }
  });

  // Delete all students (with related scores)
  document.getElementById('btn-delete-all-students')?.addEventListener('click', async ()=>{
    try{
      const proceed = confirm('This will permanently delete ALL students and their scores. Continue?');
      if (!proceed) return;
      const { data: uData, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(uErr.message, false);
      const user_id = uData?.user?.id;
      if (!user_id) return notify('Not signed in', false);
      // Load all student IDs
      const { data: students, error: sErr } = await supabase.from('students').select('id').eq('user_id', user_id);
      if (sErr) return notify(sErr.message, false);
      const ids = (students||[]).map(s=>s.id);
      if (!ids.length) return notify('No students to delete');
      // Delete related scores first to avoid FK conflicts
      // Supabase/PostgREST limits: delete in chunks to be safe
      for (let i = 0; i < ids.length; i += 1000) {
        const group = ids.slice(i, i + 1000);
        const { error: delScoresErr } = await supabase.from('scores').delete().in('student_id', group).eq('user_id', user_id);
        if (delScoresErr) return notify(delScoresErr.message, false);
      }

      // Now delete students in chunks
      for (let i = 0; i < ids.length; i += 1000) {
        const group = ids.slice(i, i + 1000);
        const { error: delStuErr } = await supabase.from('students').delete().in('id', group).eq('user_id', user_id);
        if (delStuErr) return notify(delStuErr.message, false);
      }
      notify('All students deleted');
      await loadStudents();
    } catch(e){
      notify(`Print failed: ${String(e)}`, false);
    }
  });

  // Live input: allow free typing (including commas/partials); no bounds enforcement here
  document.body.addEventListener('input', (e)=>{
    const inp = e.target.closest('.score-input');
    if (!inp) return;
    // Do not block typing; just clear any previous validity message
    inp.setCustomValidity('');
  });
  // On blur: clean commas, validate bounds, warn+clear if invalid; else round to 2 decimals
  document.body.addEventListener('blur', (e)=>{
    const inp = e.target.closest('.score-input');
    if (!inp) return;
    const maxAttr = parseFloat(inp.getAttribute('max'));
    const minAttr = parseFloat(inp.getAttribute('min'));
    const max = Number.isFinite(maxAttr) ? maxAttr : Infinity;
    const min = Number.isFinite(minAttr) ? minAttr : 0;
    const raw = inp.value;
    if (raw === '') { inp.setCustomValidity(''); return; }
    const cleaned = raw.replace(/,/g, '');
    let num = parseFloat(cleaned);
    if (Number.isNaN(num)) { inp.value = ''; inp.setCustomValidity(''); return; }
    if (num > max) {
      notify(`Max score for this subject is ${max}.`, false);
      inp.value = '';
      inp.setCustomValidity(`Must be <= ${max}`);
      try { inp.reportValidity && inp.reportValidity(); } catch {}
      return;
    }
    if (num < min) {
      notify(`Minimum score is ${min}.`, false);
      inp.value = '';
      inp.setCustomValidity(`Must be >= ${min}`);
      try { inp.reportValidity && inp.reportValidity(); } catch {}
      return;
    }
    // valid: normalize to two decimals
    inp.value = String(Math.round(num * 100) / 100);
    inp.setCustomValidity('');
  }, true);

  // Report
  $('#btn-run-report')?.addEventListener('click', runReport);
  bindReportAuto();

  // Settings
  // Refresh settings when opening the tab (ensures latest values/preview)
  document.querySelector('[data-tab="settings"]')?.addEventListener('click', async ()=>{
    await loadSettings();
  });
  const formSettings = document.getElementById('form-settings');
  formSettings?.addEventListener('submit', async (e)=>{
    e.preventDefault();
    const { data: uData, error: uErr } = await supabase.auth.getUser();
    if (uErr) return notify(uErr.message, false);
    const user_id = uData?.user?.id;
    if (!user_id) return notify('Not signed in', false);
    const payload = {
      user_id,
      school_name: document.getElementById('settings-school')?.value?.trim() || null,
      teacher_name: document.getElementById('settings-teacher')?.value?.trim() || null,
      logo_url: document.getElementById('settings-logo')?.value?.trim() || null,
      teacher_photo_url: document.getElementById('settings-teacher-photo')?.value?.trim() || null,
      class_qr_url: document.getElementById('settings-qr')?.value?.trim() || null,
      teacher_Sign_Url: document.getElementById('settings-sign')?.value?.trim() || null,
    };
    const { error } = await supabase.from('settings').upsert(payload, { onConflict: 'user_id' });
    if (error) return notify(error.message, false);
    notify('Settings saved');
    await loadSettings();
  });
  const logoInput = document.getElementById('settings-logo');
  const teacherPhotoInput = document.getElementById('settings-teacher-photo');
  const qrInput = document.getElementById('settings-qr');
  const signInput = document.getElementById('settings-sign');
  const schoolInput = document.getElementById('settings-school');
  const teacherInput = document.getElementById('settings-teacher');
  const logoImg = document.getElementById('preview-logo');
  const teacherImg = document.getElementById('preview-teacher');
  const qrImg = document.getElementById('preview-qr');
  const signImg = document.getElementById('preview-sign');
  const previewSchool = document.getElementById('preview-school');
  const previewTeacherName = document.getElementById('preview-teacher-name');
  const headerSchool = document.getElementById('header-school');
  logoInput && logoInput.addEventListener('input', ()=>{ if (logoImg) logoImg.src = logoInput.value; });
  teacherPhotoInput && teacherPhotoInput.addEventListener('input', ()=>{ if (teacherImg) teacherImg.src = teacherPhotoInput.value; });
  qrInput && qrInput.addEventListener('input', ()=>{ if (qrImg) qrImg.src = qrInput.value; });
  signInput && signInput.addEventListener('input', ()=>{ if (signImg) signImg.src = signInput.value; });
  schoolInput && schoolInput.addEventListener('input', ()=>{
    if (previewSchool) previewSchool.textContent = schoolInput.value;
    if (headerSchool) headerSchool.textContent = schoolInput.value;
  });
  teacherInput && teacherInput.addEventListener('input', ()=>{ if (previewTeacherName) previewTeacherName.textContent = teacherInput.value; });

  // Change password (Settings)
  const formPwd = document.getElementById('form-password');
  const btnPwd = document.getElementById('btn-change-password');
  const inpOld = document.getElementById('settings-old-pass');
  const inpNew = document.getElementById('settings-new-pass');
  const inpNew2 = document.getElementById('settings-new-pass2');

  let __pwdVerified = false;
  const btnResetPwd = document.getElementById('btn-reset-password');
  function setPwdInputsEnabled(flag){
    if (inpNew) inpNew.disabled = !flag;
    if (inpNew2) inpNew2.disabled = !flag;
  }

  function updatePwdBtn(){
    const a = inpOld?.value || '';
    const b = inpNew?.value || '';
    const c = inpNew2?.value || '';
    const hasOld = __pwdVerified ? true : !!a; // if verified via Reset, don't require typing old password again
    const ok = hasOld && !!b && !!c && b === c && b.length >= 6;
    if (btnPwd){
      btnPwd.disabled = !ok;
      btnPwd.classList.toggle('opacity-50', !ok);
      btnPwd.classList.toggle('cursor-not-allowed', !ok);
    }
  }

  inpOld?.addEventListener('input', ()=>{ inpOld.setCustomValidity(''); updatePwdBtn(); });
  inpNew?.addEventListener('input', ()=>{ inpNew.setCustomValidity(''); inpNew2?.setCustomValidity(''); updatePwdBtn(); });
  inpNew2?.addEventListener('input', ()=>{ inpNew2.setCustomValidity(''); updatePwdBtn(); });

  // Initially keep new password inputs disabled until verification
  setPwdInputsEnabled(false);
  updatePwdBtn();

  // Reset flow: prompt for old password and verify
  btnResetPwd?.addEventListener('click', async ()=>{
    try{
      // Clear new fields
      if (inpNew) inpNew.value = '';
      if (inpNew2) inpNew2.value = '';
      updatePwdBtn();
      const { data: { user }, error: uErr } = await supabase.auth.getUser();
      if (uErr) return notify(`Get user failed: ${uErr.message}`, false);
      const email = user?.email;
      if (!email) return notify('No signed-in user', false);
      const entered = prompt('Enter your current password to continue:');
      if (entered == null) return; // cancelled
      if (!entered) return notify('Old password required', false);
      const { error: signErr } = await supabase.auth.signInWithPassword({ email, password: entered });
      if (signErr) return notify(`Old password is incorrect: ${signErr.message}`, false);
      __pwdVerified = true;
      setPwdInputsEnabled(true);
      updatePwdBtn();
      notify('Verified. You can now enter a new password.');
      inpNew?.focus();
    } catch(e){
      notify(`Reset verification error: ${String(e)}`, false);
    }
  });

  // Submit: verify (if needed) then update password
  formPwd?.addEventListener('submit', async (e)=>{
    e.preventDefault();
    const btn = btnPwd;
    const oldPass = inpOld?.value || '';
    const newPass = inpNew?.value || '';
    const newPass2 = inpNew2?.value || '';
    // Basic validations
    if (!newPass || !newPass2){
      if (!newPass){ inpNew?.setCustomValidity('Required'); try{ inpNew?.reportValidity(); }catch{} }
      if (!newPass2){ inpNew2?.setCustomValidity('Required'); try{ inpNew2?.reportValidity(); }catch{} }
      return notify('Fill new password fields', false);
    }
    if (newPass !== newPass2){
      inpNew2?.setCustomValidity('Passwords do not match'); try{ inpNew2?.reportValidity(); }catch{}
      return notify('New passwords do not match', false);
    }
    if (newPass.length < 6){
      inpNew?.setCustomValidity('Minimum 6 characters'); try{ inpNew?.reportValidity(); }catch{}
      return notify('New password must be at least 6 characters', false);
    }
    try{
      btn && (btn.disabled = true);
      if (!__pwdVerified){
        if (!oldPass){
          try{ inpOld?.reportValidity(); }catch{}
          return notify('Please click Reset and verify or enter your old password.', false);
        }
        const { data: { user }, error: uErr } = await supabase.auth.getUser();
        if (uErr) return notify(`Get user failed: ${uErr.message}`, false);
        const email = user?.email;
        if (!email) return notify('No signed-in user', false);
        const { error: signErr } = await supabase.auth.signInWithPassword({ email, password: oldPass });
        if (signErr){
          inpOld?.setCustomValidity('Old password is incorrect');
          try{ inpOld?.reportValidity(); }catch{}
          return notify(`Old password is incorrect: ${signErr.message}`, false);
        }
      }
      const { error: updErr } = await supabase.auth.updateUser({ password: newPass });
      if (updErr) return notify(`Update failed: ${updErr.message}`, false);
      notify('Password updated');
      if (inpOld) inpOld.value = '';
      if (inpNew) inpNew.value = '';
      if (inpNew2) inpNew2.value = '';
      __pwdVerified = false;
      setPwdInputsEnabled(false);
      updatePwdBtn();
    } catch(e){
      notify(`Change password error: ${String(e)}`, false);
    } finally {
      btn && (btn.disabled = false);
    }
  });

}

// Add this closing bracket to close the bindEvents() function


// Auth/UI gating
let __appInitialized = false;
function showAuth(){
  const auth = document.getElementById('auth-shell');
  const app = document.getElementById('app-shell');
  auth?.classList.remove('hidden');
  app?.classList.add('hidden');
}
function showApp(session){
  const auth = document.getElementById('auth-shell');
  const app = document.getElementById('app-shell');
  const emailEl = document.getElementById('user-email');
  if (emailEl) emailEl.textContent = session?.user?.email || '-';
  auth?.classList.add('hidden');
  app?.classList.remove('hidden');
}
function bindAuthUI(){
  // Sign in
  const form = document.getElementById('form-auth');
  form?.addEventListener('submit', async (e)=>{
    e.preventDefault();
    const email = document.getElementById('auth-email')?.value?.trim();
    const password = document.getElementById('auth-password')?.value;
    if (!email || !password) return notify('Email and password are required', false);
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) return notify(error.message, false);
    notify('Signed in');
  });
  // Sign up
  const btnUp = document.getElementById('btn-sign-up');
  btnUp?.addEventListener('click', async ()=>{
    const email = document.getElementById('auth-email')?.value?.trim();
    const password = document.getElementById('auth-password')?.value;
    if (!email || !password) return notify('Enter email and password to create account', false);
    const { error } = await supabase.auth.signUp({ email, password });
    if (error) return notify(error.message, false);
    notify('Account created. Check your email if confirmation is required.');
  });
  // Sign out
  const btnOut = document.getElementById('btn-sign-out');
  btnOut?.addEventListener('click', async ()=>{
    const { error } = await supabase.auth.signOut();
    if (error) return notify(error.message, false);
    notify('Signed out');
  });
}

async function loadSettings(){
  try{
    const { data: uData, error: uErr } = await supabase.auth.getUser();
    if (uErr) return notify(`Settings load failed: ${uErr.message}`, false);
    const user_id = uData?.user?.id;
    if (!user_id) return notify('Not signed in', false);
    let { data, error, status } = await supabase
      .from('settings')
      .select('user_id, school_name, teacher_name, logo_url, teacher_photo_url, class_qr_url, teacher_Sign_Url')
      .eq('user_id', user_id)
      .maybeSingle();
    if (error && error.code !== 'PGRST116') return notify(`Settings load failed: ${error.message}`, false);
    if (!data){
      const { error: insErr } = await supabase.from('settings').upsert({ user_id }, { onConflict: 'user_id' });
      if (insErr) return notify(`Settings init failed: ${insErr.message}`, false);
      ({ data } = await supabase
        .from('settings')
        .select('user_id, school_name, teacher_name, logo_url, teacher_photo_url, class_qr_url, teacher_Sign_Url')
        .eq('user_id', user_id)
        .maybeSingle());
    }
    const s = data || {};
    const setVal = (id, val)=>{ const el = document.getElementById(id); if (el) el.value = val || ''; };
    setVal('settings-school', s.school_name);
    setVal('settings-teacher', s.teacher_name);
    setVal('settings-logo', s.logo_url);
    setVal('settings-teacher-photo', s.teacher_photo_url);
    setVal('settings-qr', s.class_qr_url);
    setVal('settings-sign', s.teacher_Sign_Url);
    const logoImg = document.getElementById('preview-logo'); if (logoImg) logoImg.src = s.logo_url || '';
    const teacherImg = document.getElementById('preview-teacher'); if (teacherImg) teacherImg.src = s.teacher_photo_url || '';
    const qrImg = document.getElementById('preview-qr'); if (qrImg) qrImg.src = s.class_qr_url || '';
    const signImg = document.getElementById('preview-sign'); if (signImg) signImg.src = s.teacher_Sign_Url || '';
    const prevSchool = document.getElementById('preview-school'); if (prevSchool) prevSchool.textContent = s.school_name || '';
    const prevTeacherName = document.getElementById('preview-teacher-name'); if (prevTeacherName) prevTeacherName.textContent = s.teacher_name || '';
    const headerSchoolEl = document.getElementById('header-school'); if (headerSchoolEl) headerSchoolEl.textContent = s.school_name || '';
  } catch (e){
    notify(`Settings load exception: ${String(e)}`, false);
  }
}

async function initApp(){
  if (__appInitialized) return; // prevent double-binding
  __appInitialized = true;
  bindEvents();
  await Promise.all([
    loadClasses(),
    loadStudents(),
    loadSubjects(),
    loadExams(),
    loadSettings()
  ]);
  // If report selects have values, auto-run once
  await runReport();
  // If Scores selections are already set on refresh, render grid and color inputs
  try{
    const scoreClassSel = document.querySelector('#score-class');
    const scoreExamSel = document.querySelector('#score-exam');
    if (scoreClassSel?.value && scoreExamSel?.value && !scoreExamSel?.disabled){
      await loadScoresGrid();
      document.querySelectorAll('#table-scores tbody .score-input').forEach(colorScoreInput);
    }
  } catch {}
}

async function init(){
  bindAuthUI();
  const { data: { session } } = await supabase.auth.getSession();
  if (session){
    showApp(session);
    await initApp();
  } else {
    showAuth();
  }
  supabase.auth.onAuthStateChange(async (_event, sessionNew) => {
    if (sessionNew){
      showApp(sessionNew);
      await initApp();
    } else {
      showAuth();
    }
  });
}

init();
