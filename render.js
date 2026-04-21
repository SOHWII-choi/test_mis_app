// KT M&S 경영성과 대시보드 — 렌더링
// ─────────────────────────────────
// ═══════════════════════════════════════════
//  STATE & CONFIG
// ═══════════════════════════════════════════
const S={role:null,page:null,pi:{s:24,e:35},charts:{}};
Chart.defaults.color='#8b93b8';
Chart.defaults.font.family="'Noto Sans KR',sans-serif";
Chart.defaults.font.size=11;

// ── 차트 애니메이션 완전 차단 (배율/리사이즈 포함) ──
Chart.defaults.animation = false;
Chart.defaults.animations = false;
Chart.defaults.transitions = {};
try{
  Chart.defaults.plugins.tooltip.animation = false;
  // 모든 차트 타입별 강제 비활성화
  ['line','bar','doughnut','pie','radar','scatter','polarArea'].forEach(t=>{
    if(!Chart.overrides[t]) Chart.overrides[t]={};
    Chart.overrides[t].animation=false;
    Chart.overrides[t].animations=false;
    Chart.overrides[t].transitions={};
  });
}catch(e){}

// 라이트 테마 차트 색상
const C={
  primary:'rgba(91,110,245,1)',    primaryA:'rgba(91,110,245,0.12)',
  gold:'rgba(245,158,11,1)',       goldA:'rgba(245,158,11,0.12)',
  teal:'rgba(6,182,212,1)',        tealA:'rgba(6,182,212,0.12)',
  blue:'rgba(59,130,246,1)',       blueA:'rgba(59,130,246,0.12)',
  red:'rgba(239,68,68,1)',         redA:'rgba(239,68,68,0.12)',
  green:'rgba(16,185,129,1)',      greenA:'rgba(16,185,129,0.12)',
  purple:'rgba(139,92,246,1)',     purpleA:'rgba(139,92,246,0.12)',
  orange:'rgba(249,115,22,1)',     orangeA:'rgba(249,115,22,0.12)',
  pink:'rgba(236,72,153,1)',       pinkA:'rgba(236,72,153,0.12)',
  indigo:'rgba(99,102,241,1)',     indigoA:'rgba(99,102,241,0.12)',
};
const COLORS=[C.primary,C.teal,C.blue,C.purple,C.orange,C.red,C.green,C.pink];
const COLORS_A=[C.primaryA,C.tealA,C.blueA,C.purpleA,C.orangeA,C.redA,C.greenA,C.pinkA];
// 막대그래프용 불투명 색상 (0.72)
const COLORS_BAR=[
  'rgba(91,110,245,0.72)','rgba(6,182,212,0.72)','rgba(59,130,246,0.72)',
  'rgba(139,92,246,0.72)','rgba(249,115,22,0.72)','rgba(239,68,68,0.72)',
  'rgba(16,185,129,0.72)','rgba(236,72,153,0.72)'
];

// ── 차트 생성 — 떨림·애니메이션 완전 차단 ──
function mkC(id, cfg){
  if(S.charts[id]){ try{S.charts[id].destroy();}catch(e){} delete S.charts[id]; }
  const el = document.getElementById(id);
  if(!el) return;

  cfg.options = cfg.options || {};
  // animation 완전 차단
  cfg.options.animation = false;
  cfg.options.animations = false;
  cfg.options.transitions = {};
  // responsive:false → 배율 바꿔도 절대 재렌더 안 함
  cfg.options.responsive = false;
  cfg.options.maintainAspectRatio = false;
  // scale animation도 차단
  if(cfg.options.scales){
    Object.values(cfg.options.scales).forEach(s=>{if(s)s.animation=false;});
  }

  // 캔버스 크기 명시적 설정
  const parent = el.parentElement;
  if(parent){
    const isDonut = parent.classList.contains('cw-donut') || parent.classList.contains('donut-wrap');
    if(isDonut){
      // 도넛: 정확히 240×240, devicePixelRatio 고려
      const dpr = window.devicePixelRatio || 1;
      el.width = 240 * dpr; el.height = 240 * dpr;
      el.style.width = '240px'; el.style.height = '240px';
      cfg.options = cfg.options || {};
      cfg.options.responsive = false;
      cfg.options.maintainAspectRatio = false;
    } else {
      const w = parent.clientWidth || 400;
      const h = el.dataset.h ? parseInt(el.dataset.h) : (parent.classList.contains('cw-tall') ? 300 : 240);
      el.width = w; el.height = h;
      el.style.width = w+'px'; el.style.height = h+'px';
    }
  }

  S.charts[id] = new Chart(el, cfg);
  return S.charts[id];
}

function baseOpts(extra={}){
  const gridColor='rgba(228,231,240,0.8)';
  const tickColor='#8b93b8';
  return {
    animation: false,
    animations: false,
    responsive:true,maintainAspectRatio:true,
    plugins:{
      legend:{position:'bottom',labels:{boxWidth:10,padding:16,color:tickColor,font:{size:11},usePointStyle:true,pointStyleWidth:8}},
      tooltip:{backgroundColor:'#fff',titleColor:'#1a1f36',bodyColor:'#4a5380',borderColor:'#e4e7f0',borderWidth:1,padding:10,boxShadow:'0 4px 20px rgba(0,0,0,.1)',cornerRadius:10,displayColors:true}
    },
    scales:{
      x:{grid:{color:gridColor,drawTicks:false},ticks:{maxRotation:45,color:tickColor,font:{size:10}},border:{display:false}},
      y:{grid:{color:gridColor,drawTicks:false},ticks:{color:tickColor,font:{size:10}},border:{display:false}}
    },
    ...extra
  };
}

// 선 그래프 (매출 트렌드용)
function line(id,labels,datasets,extra={}){
  mkC(id,{type:'line',data:{labels,datasets:datasets.map((d,i)=>({
    pointRadius:3,pointHoverRadius:5,tension:.4,borderWidth:2.5,
    ...d,
    borderColor:d.color||COLORS[i],
    backgroundColor:d.fill?(d.fillColor||COLORS_A[i]):'transparent',
    pointBackgroundColor:d.color||COLORS[i],
    pointBorderColor:'#fff',
    pointBorderWidth:2,
  }))},options:baseOpts(extra)});
}

// 막대 그래프 (실적 비교용)
function bar(id,labels,datasets,stacked=false,extra={}){
  const sc=stacked?{
    x:{stacked:true,grid:{color:'rgba(228,231,240,.8)'},ticks:{maxRotation:45,color:'#8b93b8',font:{size:10}},border:{display:false}},
    y:{stacked:true,grid:{color:'rgba(228,231,240,.8)'},ticks:{color:'#8b93b8',font:{size:10}},border:{display:false}}
  }:undefined;
  mkC(id,{type:'bar',data:{labels,datasets:datasets.map((d,i)=>({
    borderWidth:0,borderRadius:4,borderSkipped:false,
    ...d,
    backgroundColor:d.bg||COLORS_BAR[i%COLORS_BAR.length],
    borderColor:'transparent',
    hoverBackgroundColor:COLORS[i%COLORS.length],
  }))},options:baseOpts(sc?{scales:sc}:{...extra})});
}
const MENUS={
  executive:[
    {id:'finance',icon:'💰',label:'재무',sec:'경영성과'},
    {id:'wireless',icon:'📱',label:'무선 가입자',sec:'가입자 현황'},
    {id:'wired',icon:'🌐',label:'유선 가입자',sec:'가입자 현황'},
    {id:'org',icon:'🏢',label:'조직별 실적',sec:'채널별 실적'},
    {id:'digital',icon:'💻',label:'디지털 채널',sec:'채널별 실적'},
    {id:'b2b',icon:'🏭',label:'B2B 채널',sec:'채널별 실적'},
    {id:'smb',icon:'🏪',label:'소상공인',sec:'채널별 실적'},
    {id:'platform',icon:'♻️',label:'유통플랫폼',sec:'채널별 실적'},
    {id:'quality',icon:'⭐',label:'TCSI / VOC',sec:'기타 실적'},
    {id:'infra',icon:'🗺️',label:'매장 인프라',sec:'기타 실적'},
    {id:'strategy',icon:'🚀',label:'전략상품',sec:'기타 실적'},
    {id:'hr',icon:'👥',label:'인력 현황',sec:'기타 실적'},
  ],
  admin:[
    {id:'admin-overview',icon:'🖥️',label:'시스템 현황',sec:'관리'},
    {id:'admin-upload',icon:'📂',label:'데이터 업로드',sec:'관리'},
    {id:'admin-metrics',icon:'📐',label:'지표 관리',sec:'관리'},
    {id:'admin-users',icon:'👤',label:'사용자 권한',sec:'관리'},
    {id:'admin-export',icon:'📤',label:'보고서 내보내기',sec:'보고'},
    {id:'admin-beta',icon:'🧪',label:'베타 테스트 · 배포',sec:'배포'},
  ]
};

// ═══════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════
const sl=(a,s,e)=>a.slice(s,e+1);
const sum=a=>a.reduce((t,v)=>t+(v||0),0);
const last=a=>{const f=a.filter(x=>x!=null);return f.length?f[f.length-1]:0};
const prev=a=>{const f=a.filter(x=>x!=null);return f.length>1?f[f.length-2]:f[0]||0};
const fmt=(v,d=0)=>Number(v).toLocaleString('ko-KR',{minimumFractionDigits:d,maximumFractionDigits:d});
const pct=(a,b)=>b?((a-b)/Math.abs(b)*100):0;

function gd(key){
  const src=RAW[key],{s,e}=S.pi,o={months:sl(src.months,s,e)};
  for(const k of Object.keys(src))if(k!=='months'&&!Array.isArray(src[k])&&typeof src[k]==='object'){o[k]={};for(const r of Object.keys(src[k]))o[k][r]=sl(src[k][r],s,e);}
  else if(k!=='months')o[k]=sl(src[k],s,e);
  return o;
}
function getPL(){const{s,e}=S.pi;return`${RAW.finance.months[s]} ~ ${RAW.finance.months[e]}`;}

function kpi(label,value,unit,diff,cls,sub='',onclickKey=''){
  const badge=diff!=null?`<div class="kpi-badge ${diff>=0?'kbd-up':'kbd-dn'}">${diff>=0?'▲':'▼'} ${Math.abs(diff).toFixed(1)}%</div>`:'';
  const hint = onclickKey ? `<div style="font-size:9px;color:var(--text3);margin-top:5px;opacity:.7;display:flex;align-items:center;gap:3px"><span>🔍</span><span>클릭하여 상세분석</span></div>` : '';
  const inner = `<div class="kpi-label">${label}</div><div class="kpi-value">${value}<span class="kpi-unit">${unit}</span></div>${badge}${sub?`<div class="kpi-sub">${sub}</div>`:''}${hint}`;
  const extra = onclickKey ? `data-kpi-key="${onclickKey}" style="cursor:pointer" title="${label} 상세분석 클릭"` : '';
  return `<div class="kpi-card kc-${cls}" ${extra}>${inner}</div>`;
}

// 미니 스파크라인 (KPI 카드 하단에 삽입)
function addSparklines(gridId, dataMap){
  requestAnimationFrame(()=>{
    const grid = document.getElementById(gridId);
    if(!grid) return;
    Object.entries(dataMap).forEach(([key, arr])=>{
      const card = grid.querySelector(`[data-kpi-key="${key}"]`);
      if(!card || !arr || arr.length < 3) return;
      const wrap = document.createElement('div');
      wrap.className='sparkline-wrap';
      const cvs = document.createElement('canvas');
      cvs.height=28; cvs.style.height='28px';
      wrap.appendChild(cvs); card.appendChild(wrap);
      const valid=arr.filter(v=>v!=null);
      const minV=Math.min(...valid),maxV=Math.max(...valid);
      const last3=arr.slice(-12).map(v=>v??0);
      const rising = last3[last3.length-1] > last3[0];
      new Chart(cvs,{
        type:'line',
        data:{labels:last3.map((_,i)=>i),datasets:[{data:last3,borderColor:rising?'rgba(16,185,129,0.7)':'rgba(239,68,68,0.7)',backgroundColor:'transparent',borderWidth:1.5,pointRadius:0,tension:.4}]},
        options:{animation:false,animations:false,responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{enabled:false}},scales:{x:{display:false},y:{display:false,min:minV*0.95,max:maxV*1.05}}}
      });
    });
  });
}
// ── KPI 카드 이벤트 위임 (버블링 방식으로 확실히 처리) ──
document.addEventListener('click', function(e){
  const card = e.target.closest('[data-kpi-key]');
  if(card){
    e.stopPropagation();
    openKpiPopup(card.dataset.kpiKey);
  }
});


// ═══════════════════════════════════════════
//  PERIOD BAR
// ═══════════════════════════════════════════
function buildPB(cid,fn){
  const el=document.getElementById(cid);if(!el)return;
  const mons=RAW.finance.months;
  el.innerHTML=`<div class="period-bar">
    <label>빠른선택:</label>
    ${['23','24','25'].map(y=>`<button class="pbtn ${getPL().startsWith(y)?'act':''}" onclick="selY('${y}','${cid}','${fn}')">${y}년</button>`).join('')}
    <button class="pbtn" onclick="selY('all','${cid}','${fn}')">전체</button>
    <span class="psep">|</span><label>시작:</label>
    <select class="period-select" id="${cid}-s" onchange="updP('${cid}','${fn}')">
      ${mons.map((m,i)=>`<option value="${i}"${i===S.pi.s?' selected':''}>${m}</option>`).join('')}
    </select>
    <label>종료:</label>
    <select class="period-select" id="${cid}-e" onchange="updP('${cid}','${fn}')">
      ${mons.map((m,i)=>`<option value="${i}"${i===S.pi.e?' selected':''}>${m}</option>`).join('')}
    </select>
    <span style="font-size:10px;color:var(--text3);margin-left:6px">📅 ${getPL()}</span>
  </div>`;
}
function selY(y,cid,fn){
  const m=RAW.finance.months;
  if(y==='all')S.pi={s:0,e:m.length-1};
  else{const s=m.findIndex(x=>x.startsWith(y+'.'));const e=m.reduce((a,x,i)=>x.startsWith(y+'.')?i:a,-1);if(s>=0)S.pi={s,e:e>=0?e:m.length-1};}
  buildPB(cid,fn);window[fn]();
}
function updP(cid,fn){
  const s=parseInt(document.getElementById(cid+'-s').value);
  const e=parseInt(document.getElementById(cid+'-e').value);
  S.pi={s:Math.min(s,e),e:Math.max(s,e)};
  buildPB(cid,fn);
  document.getElementById('topbar-period').textContent=getPL();
  window[fn]();
}

// ═══════════════════════════════════════════
//  PAGE RENDERS
// ═══════════════════════════════════════════

// ── 재무 ──
// ════════════════════════════════════════════════
//  AI 분석 (Claude API)
// ════════════════════════════════════════════════
const AI_MODEL = 'claude-sonnet-4-20250514';

// ── API 키 관리 ──
function getApiKey(){ try{ return localStorage.getItem('kts_api_key')||''; }catch(e){ return ''; } }
function saveApiKeyLocal(k){ try{ localStorage.setItem('kts_api_key',k); }catch(e){} }
function clearApiKey(){ try{ localStorage.removeItem('kts_api_key'); }catch(e){} }

function updateApiKeyBtn(){
  const btn=document.getElementById('apikey-btn');
  if(!btn) return;
  const k=getApiKey();
  if(k){
    btn.classList.add('set');
    btn.innerHTML='✓ API 연결됨';
    btn.title='AI API 키 설정됨 · 클릭하여 변경';
  } else {
    btn.classList.remove('set');
    btn.innerHTML='🔑 API 키';
    btn.title='AI 분석을 위한 API 키 설정';
  }
}

function openApiKey(){
  const k=getApiKey();
  document.getElementById('apikey-input').value=k?'sk-ant-••••••••••••••':'';
  document.getElementById('apikey-input').type='password';
  document.getElementById('apikey-eye-btn').textContent='👁';
  document.getElementById('apikey-status').textContent=k?'✓ 저장된 키가 있습니다. 변경하려면 새 키를 입력하세요.':'';
  document.getElementById('apikey-status').className='apikey-status '+(k?'ok':'');
  document.getElementById('apikey-modal').classList.add('open');
  setTimeout(()=>document.getElementById('apikey-input').focus(), 80);
}
function closeApiKey(){
  document.getElementById('apikey-modal').classList.remove('open');
}
function toggleApiKeyVisible(){
  const inp=document.getElementById('apikey-input');
  const btn=document.getElementById('apikey-eye-btn');
  inp.type=inp.type==='password'?'text':'password';
  btn.textContent=inp.type==='password'?'👁':'🙈';
}
function onApiKeyInput(){
  document.getElementById('apikey-status').textContent='';
  document.getElementById('apikey-status').className='apikey-status';
}

async function saveApiKey(){
  const raw=document.getElementById('apikey-input').value.trim();
  const statusEl=document.getElementById('apikey-status');
  const saveBtn=document.getElementById('apikey-save-btn');

  // 마스킹된 기존 값이면 → 기존 키 유지
  const key = (raw.startsWith('sk-ant-••') || raw==='') ? getApiKey() : raw;

  if(!key||!key.startsWith('sk-ant-')){
    statusEl.textContent='⚠ 올바른 키 형식이 아닙니다 (sk-ant-... 로 시작해야 합니다)';
    statusEl.className='apikey-status err';
    return;
  }

  statusEl.textContent='⏳ 연결 확인 중...';
  statusEl.className='apikey-status testing';
  saveBtn.disabled=true;

  // 실제 API 호출로 키 유효성 검증
  try{
    const resp=await fetch('https://api.anthropic.com/v1/messages',{
      method:'POST',
      headers:{
        'Content-Type':'application/json',
        'x-api-key': key,
        'anthropic-version':'2023-06-01',
        'anthropic-dangerous-direct-browser-access':'true'
      },
      body:JSON.stringify({
        model:AI_MODEL, max_tokens:10,
        messages:[{role:'user',content:'hi'}]
      })
    });
    if(resp.status===401){
      statusEl.textContent='❌ API 키가 유효하지 않습니다. 다시 확인해주세요.';
      statusEl.className='apikey-status err';
      saveBtn.disabled=false;
      return;
    }
    // 200 or other non-auth error = key is valid
    saveApiKeyLocal(key);
    statusEl.textContent='✅ API 키 연결 성공! AI 분석 기능이 활성화되었습니다.';
    statusEl.className='apikey-status ok';
    updateApiKeyBtn();
    setTimeout(closeApiKey, 1400);
    showToast('✅ API 키 저장 완료 · AI 분석 사용 가능');
  }catch(e){
    statusEl.textContent='⚠ 네트워크 오류. 인터넷 연결을 확인하세요.';
    statusEl.className='apikey-status err';
  }
  saveBtn.disabled=false;
}

// ── AI 호출 (키 포함) ──
async function callAI(prompt, systemPrompt){
  const key=getApiKey();
  if(!key){
    openApiKey();
    throw new Error('API 키를 먼저 설정해주세요 (우측 상단 🔑 버튼)');
  }
  const resp = await fetch('https://api.anthropic.com/v1/messages',{
    method:'POST',
    headers:{
      'Content-Type':'application/json',
      'x-api-key': key,
      'anthropic-version':'2023-06-01',
      'anthropic-dangerous-direct-browser-access':'true'
    },
    body: JSON.stringify({
      model: AI_MODEL,
      max_tokens: 1000,
      system: systemPrompt || '당신은 KT M&S의 경영성과 분석 전문가입니다. 데이터를 근거로 간결하고 통찰력 있는 분석을 제공합니다. 한국어로 답변하세요.',
      messages:[{role:'user',content:prompt}]
    })
  });
  const data = await resp.json();
  if(!resp.ok) throw new Error(data.error?.message || `API 오류 (${resp.status})`);
  return data.content?.[0]?.text || '';
}

function setAILoading(bodyId, btnId){
  document.getElementById(bodyId).innerHTML = `
    <div class="ai-skeleton">
      <div class="ai-skel-line" style="width:90%"></div>
      <div class="ai-skel-line" style="width:75%"></div>
      <div class="ai-skel-line" style="width:85%"></div>
      <div class="ai-skel-line short"></div>
    </div>`;
  const btn = document.getElementById(btnId);
  if(btn){ btn.disabled = true; btn.textContent = '⏳ 분석 중...'; }
}

function setAIResult(bodyId, btnId, text){
  document.getElementById(bodyId).textContent = text;
  const btn = document.getElementById(btnId);
  if(btn){ btn.disabled = false; btn.textContent = '↺ 재분석'; }
}

function setAIError(bodyId, btnId, msg){
  document.getElementById(bodyId).innerHTML = `<span style="color:var(--red);font-size:12px">⚠ ${msg}</span>`;
  const btn = document.getElementById(btnId);
  if(btn){ btn.disabled = false; btn.textContent = '✦ 다시 시도'; }
}

async function runFinanceAI(){
  const d = gd('finance');
  const period = getPL();
  setAILoading('fin-ai-body','fin-ai-btn');
  document.getElementById('fin-ai-subtitle').textContent = `분석 기간: ${period}`;

  const summary = `
[분석 기간] ${period}
[총매출] 누계 ${fmt(sum(d.총매출),1)}억 / 최근월 ${fmt(last(d.총매출),1)}억
[통신매출] 누계 ${fmt(sum(d.통신매출),1)}억 / 최근월 ${fmt(last(d.통신매출),1)}억
[상품매출] 누계 ${fmt(sum(d.상품매출),1)}억
[상품이익] 누계 ${fmt(sum(d.상품이익),1)}억
[영업이익] 누계 ${fmt(sum(d.영업이익),1)}억 / 최근월 ${fmt(last(d.영업이익),1)}억
[판관비] 누계 ${fmt(sum(d.판관비),1)}억
[인건비] 누계 ${fmt(sum(d.인건비),1)}억
[마케팅비] 누계 ${fmt(sum(d.마케팅비),1)}억
[월별 영업이익 최근 6개월] ${d.영업이익.slice(-6).map((v,i)=>`${d.months[d.months.length-6+i]}: ${fmt(v,1)}억`).join(' / ')}
[전체 기간 영업이익 추이] ${d.영업이익.map((v,i)=>`${d.months[i]}:${fmt(v,1)}`).join(', ')}`;

  try{
    const result = await callAI(
      `아래 KT M&S 재무 데이터를 분석해주세요:\n${summary}\n\n다음 항목을 포함해 3~5문단으로 분석하세요:\n1. 전반적인 재무 성과 평가\n2. 수익성 트렌드 및 주요 변곡점\n3. 비용 구조 특이사항\n4. 향후 1~3개월 전망 및 리스크`
    );
    setAIResult('fin-ai-body','fin-ai-btn', result);
  } catch(e){
    setAIError('fin-ai-body','fin-ai-btn', e.message);
  }
}

// ── 채널별 영업이익 추정 로직 ──
// 각 채널의 CAPA 비중으로 전체 영업이익/비용을 안분
let _finCurrentCh = 'all';

function getChColor(ch){
  return {all:C.gold, 소매:C.gold, 도매:C.teal, 디지털:C.blue, B2B:C.purple, 소상공인:C.orange}[ch]||C.gold;
}
function getChColorA(ch){
  return {all:C.goldA, 소매:C.goldA, 도매:C.tealA, 디지털:C.blueA, B2B:C.purpleA, 소상공인:C.orangeA}[ch]||C.goldA;
}

function calcChannelFinance(ch){
  const fin = gd('finance');
  const org = gd('org');
  const chKeys = ['소매','도매','디지털','B2B','소상공인'];

  // 각 월의 채널별 비중 계산
  const totalCapa = org.months.map((_,i)=>chKeys.reduce((s,k)=>s+(org[k]?.[i]||0),0));

  function apportionSeries(series, chKey){
    return series.map((v,i)=>{
      if(!totalCapa[i]) return 0;
      const ratio = chKey==='all' ? 1 : (org[chKey]?.[i]||0)/totalCapa[i];
      return parseFloat((v * ratio).toFixed(2));
    });
  }

  if(ch==='all'){
    return {
      months: fin.months,
      영업이익: fin.영업이익,
      판관비: fin.판관비,
      인건비: fin.인건비,
      마케팅비: fin.마케팅비,
      총매출: fin.총매출,
      통신매출: fin.통신매출,
    };
  }
  return {
    months: fin.months,
    영업이익: apportionSeries(fin.영업이익, ch),
    판관비: apportionSeries(fin.판관비, ch),
    인건비: apportionSeries(fin.인건비, ch),
    마케팅비: apportionSeries(fin.마케팅비, ch),
    총매출: apportionSeries(fin.총매출, ch),
    통신매출: apportionSeries(fin.통신매출, ch),
  };
}

function selectFinCh(ch, btn){
  _finCurrentCh = ch;
  document.querySelectorAll('#fin-ch-tabs .ch-tab').forEach(b=>b.classList.remove('active'));
  if(btn) btn.classList.add('active');
  const d = calcChannelFinance(ch);
  const chLabel = ch==='all'?'전체 합산':ch;
  const col = getChColor(ch);
  const colA = getChColorA(ch);

  // Channel KPI
  document.getElementById('fin-ch-kpi').innerHTML = [
    kpi('영업이익 (누계)', fmt(sum(d.영업이익),1), '억원', pct(last(d.영업이익),prev(d.영업이익)), sum(d.영업이익)>=0?'green':'red', `${chLabel} · 최근월 ${fmt(last(d.영업이익),1)}억`),
    kpi('총매출 기여 (누계)', fmt(sum(d.총매출),1), '억원', null, 'gold', `${chLabel}`),
    kpi('판관비 (누계)', fmt(sum(d.판관비),1), '억원', null, 'purple', `${chLabel}`),
    kpi('인건비 (누계)', fmt(sum(d.인건비),1), '억원', null, 'blue', `${chLabel}`),
    kpi('마케팅비 (누계)', fmt(sum(d.마케팅비),1), '억원', null, 'orange', `${chLabel}`),
    kpi('영업이익률 (추정)', fmt(sum(d.영업이익)/Math.max(sum(d.총매출),1)*100,1), '%', null, sum(d.영업이익)>=0?'teal':'red', `${chLabel}`),
  ].join('');

  // 채널 차트 제목
  document.getElementById('fin-ch-title').textContent = `${chLabel} 영업이익 누적 추이`;
  document.getElementById('fin-ch-sub').textContent = ch==='all'?'전체 합산 영업이익 (억원)':'CAPA 비중 기반 추정 영업이익 (억원)';
  document.getElementById('fin-ch-cost-sub').textContent = `${chLabel} 비용 구조 (인건비 · 마케팅비 · 판관비)`;

  // 영업이익 bar
  mkC('ch-fin-ch-profit',{type:'bar',data:{labels:d.months,datasets:[
    {label:'영업이익',data:d.영업이익,backgroundColor:d.영업이익.map(v=>v>=0?colA:C.redA),borderColor:d.영업이익.map(v=>v>=0?col:C.red),borderWidth:1.5},
    {label:'누계',type:'line',data:(()=>{let s=0;return d.영업이익.map(v=>{s+=v;return parseFloat(s.toFixed(2));})})(),borderColor:C.teal,backgroundColor:'transparent',borderWidth:2,pointRadius:0,tension:.3,yAxisID:'y2'},
  ]},options:baseOpts({scales:{x:{grid:{color:'rgba(42,53,85,0.5)'},ticks:{maxRotation:45,color:'#5c6e9a',font:{size:10}}},y:{grid:{color:'rgba(42,53,85,0.5)'},ticks:{color:'#5c6e9a'},title:{display:true,text:'월별(억원)',color:'#5c6e9a'}},y2:{position:'right',grid:{display:false},ticks:{color:'#5c6e9a'},title:{display:true,text:'누계(억원)',color:'#5c6e9a'}}}})});

  // 비용 stacked bar
  mkC('ch-fin-ch-cost',{type:'bar',data:{labels:d.months,datasets:[
    {label:'인건비',data:d.인건비,backgroundColor:C.blueA,borderColor:C.blue,borderWidth:1,stack:'s'},
    {label:'마케팅비',data:d.마케팅비,backgroundColor:C.orangeA,borderColor:C.orange,borderWidth:1,stack:'s'},
    {label:'판관비(기타)',data:d.판관비.map((v,i)=>Math.max(0,v-d.인건비[i]-d.마케팅비[i])),backgroundColor:C.purpleA,borderColor:C.purple,borderWidth:1,stack:'s'},
  ]},options:baseOpts({scales:{x:{stacked:true,grid:{color:'rgba(42,53,85,0.5)'},ticks:{maxRotation:45,color:'#5c6e9a',font:{size:10}}},y:{stacked:true,grid:{color:'rgba(42,53,85,0.5)'},ticks:{color:'#5c6e9a'}}}})});

  // 매출 기여 도넛 (채널이 all이면 채널별, 아니면 전체 vs 해당채널)
  const chKeys=['소매','도매','디지털','B2B','소상공인'];
  if(ch==='all'){
    const org=gd('org');
    const totals=chKeys.map(k=>sum(org[k]||[]));
    const grand=sum(totals)||1;
    const shareColors=COLORS_A.slice(0,chKeys.length).map(c=>c.replace('0.12','0.80'));
    makeDonut('ch-fin-ch-share', chKeys, totals.map(v=>parseFloat((v/grand*100).toFixed(1))), shareColors, '채널별', '비중(%)');
  } else {
    const org=gd('org');
    const chCapa=sum(org[ch]||[]);
    const otherCapa=chKeys.filter(k=>k!==ch).reduce((s,k)=>s+sum(org[k]||[]),0);
    const total=chCapa+otherCapa||1;
    const chPct=parseFloat((chCapa/total*100).toFixed(1));
    makeDonut('ch-fin-ch-share', [ch,'기타채널'], [chPct, parseFloat((otherCapa/total*100).toFixed(1))], [colA.replace('0.12','0.80'),'rgba(228,231,240,0.8)'], chPct+'%', ch+' 비중');
  }

  // AI 카드 표시
  const aiCard = document.getElementById('fin-ch-ai-card');
  aiCard.style.display = '';
  document.getElementById('fin-ch-ai-subtitle').textContent = `${chLabel} 채널 분석 준비 완료 · 기간: ${getPL()}`;
  document.getElementById('fin-ch-ai-body').innerHTML = `<div style="font-size:12px;color:var(--text3);text-align:center;padding:10px 0">▶ '분석 실행' 버튼을 누르면 ${chLabel} 채널을 AI가 분석합니다</div>`;
  document.getElementById('fin-ch-ai-btn').disabled = false;
  document.getElementById('fin-ch-ai-btn').textContent = '✦ 분석 실행';
}

async function runChannelAI(){
  const ch = _finCurrentCh;
  const d = calcChannelFinance(ch);
  const chLabel = ch==='all'?'전체 합산':ch;
  const period = getPL();
  setAILoading('fin-ch-ai-body','fin-ch-ai-btn');
  document.getElementById('fin-ch-ai-subtitle').textContent = `분석 중: ${chLabel} 채널 · ${period}`;

  const org = gd('org');
  const chKeys=['소매','도매','디지털','B2B','소상공인'];
  const totCapa = chKeys.reduce((s,k)=>s+sum(org[k]||[]),0)||1;
  const chCapa = ch==='all'? totCapa : sum(org[ch]||[]);
  const capaShare = (chCapa/totCapa*100).toFixed(1);

  const summary = `
[채널] ${chLabel}
[분석 기간] ${period}
[CAPA 비중] 전체 대비 ${capaShare}%
[영업이익 누계] ${fmt(sum(d.영업이익),1)}억원
[최근월 영업이익] ${fmt(last(d.영업이익),1)}억원
[총매출 기여 누계] ${fmt(sum(d.총매출),1)}억원
[판관비 누계] ${fmt(sum(d.판관비),1)}억원
[마케팅비 누계] ${fmt(sum(d.마케팅비),1)}억원
[최근 6개월 영업이익] ${d.영업이익.slice(-6).map((v,i)=>`${d.months[d.months.length-6+i]}: ${fmt(v,1)}억`).join(' / ')}`;

  try{
    const result = await callAI(
      `아래 KT M&S ${chLabel} 채널 재무 데이터를 분석해주세요:\n${summary}\n\n다음 항목을 포함해 2~4문단으로 분석하세요:\n1. ${chLabel} 채널의 수익성 평가\n2. 비용 효율성 및 개선 포인트\n3. 전체 채널 대비 포지셔닝\n4. 향후 전망 및 액션 포인트`,
      `당신은 KT M&S의 ${chLabel} 채널 전문 재무 분석가입니다. 채널 특성을 반영한 인사이트를 제공합니다.`
    );
    setAIResult('fin-ch-ai-body','fin-ch-ai-btn', result);
  } catch(e){
    setAIError('fin-ch-ai-body','fin-ch-ai-btn', e.message);
  }
}

// ── 재무 렌더 ──
// ── 전년동기대비 계산 ──
// months 형식: "23.1월", "24.1월", "25.12월"
function getYear(m){ return (m||'').match(/^'?(\d{2})\./)?.[1] || ''; }
function getMon(m){ return (m||'').replace(/^'?\d+\./,''); }

// 전체 기간 YoY 누계 비교
function yoySum(arr, months){
  if(!arr || !months || !arr.length) return null;
  const lastM = months[months.length-1];
  const lastY = getYear(lastM);
  const lastMon = getMon(lastM);
  // 작년 같은 월 인덱스
  const prevLastIdx = months.reduce((found, m, i) => {
    if(getYear(m) === String(parseInt(lastY)-1).padStart(2,'0') && getMon(m) === lastMon) return i;
    return found;
  }, -1);
  if(prevLastIdx < 0) return null;
  // 올해 시작 인덱스
  const curStartIdx = months.findIndex(m => getYear(m) === lastY);
  const prevStartIdx = months.findIndex(m => getYear(m) === String(parseInt(lastY)-1).padStart(2,'0'));
  if(curStartIdx < 0 || prevStartIdx < 0) return null;
  const curLen = months.length - curStartIdx;
  const curSum = arr.slice(curStartIdx).reduce((s,v)=>s+(v||0),0);
  const prevSum = arr.slice(prevStartIdx, prevStartIdx+curLen).reduce((s,v)=>s+(v||0),0);
  return {cur: curSum, prev: prevSum, pct: prevSum ? (curSum-prevSum)/Math.abs(prevSum)*100 : 0};
}


window.renderFinance=function(){
  const d=gd('finance');

  // 전년동기대비 계산 — RAW 원본 데이터로 (기간 필터 없이)
  const RF = RAW.finance;
  const yoyTot  = yoySum(RF.총매출,  RF.months);
  const yoyTong = yoySum(RF.통신매출, RF.months);
  const yoyOp   = yoySum(RF.영업이익, RF.months);
  const yoyMkt  = yoySum(RF.마케팅비, RF.months);
  const yoyOpex = yoySum(RF.판관비,   RF.months);

  function yoyBadge(yoyObj){
    if(!yoyObj) return null;
    return yoyObj.pct;
  }

  // 비교기준 표시
  const lastM = d.months[d.months.length-1];
  const lastMonthPart = lastM.replace(/^'?\d+\./,'');
  const curYearStr = lastM.match(/^'?(\d+)\./)?.[1];
  const prevYearStr = curYearStr ? String(parseInt(curYearStr)-1) : '';
  const compLabel = `전년동기대비 (${prevYearStr}년 1~${lastMonthPart} vs ${curYearStr}년 1~${lastMonthPart})`;

  // 유통플랫폼 데이터
  const dp = gd('platform');
  const RP = RAW.platform;
  const platRaw = (RP.시연폰_매각이익||[]).map((v,i)=>(v||0)+((RP.중고폰_매입금액||[])[i]||0));
  const yoyPlat = yoySum(platRaw, RF.months);

  document.getElementById('fin-kpi').innerHTML=`
    <div style="grid-column:1/-1;font-size:10px;color:var(--text3);background:rgba(91,110,245,.06);border:1px solid rgba(91,110,245,.15);border-radius:8px;padding:6px 12px;display:flex;align-items:center;gap:6px">
      <span style="font-size:12px">📊</span> <b style="color:var(--primary)">비교기준:</b> ${compLabel}
      <span style="margin-left:6px;color:var(--text3)">· 각 카드 클릭 시 상세 분석</span>
    </div>`+[
    kpi('총매출 (누계)',fmt(yoyTot?.cur??sum(RF.총매출),1),'억원',yoyBadge(yoyTot),'gold',`전년동기 ${fmt(yoyTot?.prev??0,1)}억`,'fin:총매출'),
    kpi('통신매출 (누계)',fmt(yoyTong?.cur??sum(RF.통신매출),1),'억원',yoyBadge(yoyTong),'blue',`전년동기 ${fmt(yoyTong?.prev??0,1)}억`,'fin:통신매출'),
    kpi('영업이익 (누계)',fmt(yoyOp?.cur??sum(RF.영업이익),1),'억원',yoyBadge(yoyOp),(yoyOp?.cur??sum(RF.영업이익))>=0?'green':'red',`전년동기 ${fmt(yoyOp?.prev??0,1)}억`,'fin:영업이익'),
    kpi('판관비 (누계)',fmt(yoyOpex?.cur??sum(RF.판관비),1),'억원',yoyBadge(yoyOpex),'purple',`인건비 ${fmt(sum(RF.인건비),1)}억`,'fin:판관비'),
    kpi('마케팅비 (누계)',fmt(yoyMkt?.cur??sum(RF.마케팅비),1),'억원',yoyBadge(yoyMkt),'orange',`전년동기 ${fmt(yoyMkt?.prev??0,1)}억`,'fin:마케팅비'),
    kpi('유통플랫폼 (누계)',fmt(yoyPlat?.cur??sum(platRaw),1),'억원',yoyBadge(yoyPlat),'teal',`전년동기 ${fmt(yoyPlat?.prev??0,1)}억`,'fin:유통플랫폼'),
  ].join('');

  // 공통 툴팁/포인트 설정 (융통성 높임)
  const pointCfg = {
    pointRadius: 4,
    pointHoverRadius: 8,
    pointHitRadius: 24,      // ← 클릭/호버 감지 반경 대폭 확대
    pointBorderWidth: 2,
    pointBorderColor: '#fff',
  };

  // 월 클릭 → 해당월 필터 자동 적용
  const onMonthClick = (evt, elements, chart) => {
    if(!elements.length) return;
    const idx = elements[0].index;
    S.pi = {s: idx, e: idx};
    buildPB('pb-finance','renderFinance');
    renderFinance();
    showToast(`📅 ${d.months[idx]} 기준으로 필터 변경`);
  };

  // ① 매출 추이 — 선 그래프
  mkC('ch-rev',{type:'line',data:{labels:d.months,datasets:[
    {label:'총매출',...pointCfg,data:d.총매출,borderColor:C.primary,backgroundColor:C.primaryA,borderWidth:2.5,tension:.4,fill:true,yAxisID:'y'},
    {label:'통신매출',...pointCfg,data:d.통신매출,borderColor:C.teal,backgroundColor:'transparent',borderWidth:2,tension:.4,borderDash:[5,3],yAxisID:'y'},
    {label:'영업이익',...pointCfg,pointRadius:4,data:d.영업이익,borderColor:C.green,backgroundColor:'transparent',borderWidth:2,tension:.4,yAxisID:'y2'},
  ]},options:baseOpts({
    onClick: onMonthClick,
    onHover:(e,el)=>{ e.native.target.style.cursor = el.length?'pointer':'default'; },
    scales:{
      x:{grid:{color:'rgba(228,231,240,.6)'},ticks:{maxRotation:45,color:'#8b93b8',font:{size:10}},border:{display:false}},
      y:{grid:{color:'rgba(228,231,240,.6)'},ticks:{color:'#8b93b8',font:{size:10}},title:{display:true,text:'매출 (억원)',color:'#8b93b8',font:{size:10}},border:{display:false}},
      y2:{position:'right',grid:{display:false},ticks:{color:'#10b981',font:{size:10}},title:{display:true,text:'이익 (억원)',color:'#10b981',font:{size:10}},border:{display:false}}
    }
  })});

  // ② 영업이익 막대
  mkC('ch-profit',{type:'bar',data:{labels:d.months,datasets:[{
    label:'영업이익',data:d.영업이익,
    backgroundColor:d.영업이익.map(v=>v>=0?'rgba(16,185,129,0.75)':'rgba(239,68,68,0.75)'),
    borderColor:d.영업이익.map(v=>v>=0?C.green:C.red),
    borderWidth:0,borderRadius:4,borderSkipped:false,
  }]},options:baseOpts({
    onClick: onMonthClick,
    onHover:(e,el)=>{ e.native.target.style.cursor = el.length?'pointer':'default'; },
    plugins:{legend:{display:false},tooltip:{backgroundColor:'#fff',titleColor:'#1a1f36',bodyColor:'#4a5380',borderColor:'#e4e7f0',borderWidth:1,cornerRadius:10}},
  })});

  // ③ 비용 구조 누적 막대
  mkC('ch-cost',{type:'bar',data:{labels:d.months,datasets:[
    {label:'인건비',data:d.인건비,backgroundColor:'rgba(91,110,245,0.72)',borderWidth:0,borderRadius:{topLeft:0,topRight:0,bottomLeft:4,bottomRight:4},stack:'s'},
    {label:'마케팅비',data:d.마케팅비,backgroundColor:'rgba(249,115,22,0.72)',borderWidth:0,stack:'s'},
    {label:'기타 판관비',data:d.판관비.map((v,i)=>Math.max(0,v-(d.인건비[i]||0)-(d.마케팅비[i]||0))),backgroundColor:'rgba(139,92,246,0.65)',borderWidth:0,borderRadius:{topLeft:4,topRight:4,bottomLeft:0,bottomRight:0},stack:'s'},
  ]},options:baseOpts({scales:{
    x:{stacked:true,grid:{color:'rgba(228,231,240,.6)'},ticks:{maxRotation:45,color:'#8b93b8',font:{size:10}},border:{display:false}},
    y:{stacked:true,grid:{color:'rgba(228,231,240,.6)'},ticks:{color:'#8b93b8',font:{size:10}},border:{display:false}}
  }})});

  selectFinCh('all', document.querySelector('#fin-ch-tabs .ch-tab'));
  document.getElementById('fin-ai-body').innerHTML = `<div style="font-size:12px;color:var(--text3);text-align:center;padding:10px 0">위 버튼을 클릭하면 선택 기간의 재무 데이터를 AI가 분석합니다</div>`;
  document.getElementById('fin-ai-subtitle').textContent = `분석 기간: ${getPL()}`;
  document.getElementById('fin-ai-btn').disabled = false;
  document.getElementById('fin-ai-btn').textContent = '✦ 분석 실행';
  // 스파크라인
  addSparklines('fin-kpi', {
    'fin:총매출': RF.총매출,
    'fin:통신매출': RF.통신매출,
    'fin:영업이익': RF.영업이익,
    'fin:판관비': RF.판관비,
    'fin:마케팅비': RF.마케팅비,
  });
};

// ── 무선 ──
window.renderWireless=function(){
  const d=gd('wireless'),od=gd('org');
  document.getElementById('wl-kpi').innerHTML=[
    kpi('유지 가입자',fmt(last(d.유지)),'건',pct(last(d.유지),d.유지[0]),'blue','일반후불 전체','wl:유지'),
    kpi('CAPA (최근월)',fmt(last(d.CAPA)),'건',pct(last(d.CAPA),prev(d.CAPA)),'gold','신규+기변','wl:CAPA'),
    kpi('해지 (최근월)',fmt(last(d.해지)),'건',pct(last(d.해지),prev(d.해지)),'red','직권해지 포함','wl:해지'),
    kpi('순증 (최근월)',fmt(last(d.순증)),'건',null,last(d.순증)>=0?'green':'red','유지 전월차','wl:순증'),
    kpi('기간 CAPA 합계',fmt(sum(d.CAPA)),'건',null,'teal',`월평균 ${fmt(Math.round(sum(d.CAPA)/d.months.length))}건`),
    kpi('기간 해지 합계',fmt(sum(d.해지)),'건',null,'orange',`월평균 ${fmt(Math.round(sum(d.해지)/d.months.length))}건`),
  ].join('');

  // 월 클릭 → 해당월 무선 현황 팝업
  const onWlMonthClick = (evt, elements) => {
    if(!elements.length) return;
    const idx = elements[0].index;
    openWlMonthModal(idx);
  };

  const pointCfg = { pointRadius:4, pointHoverRadius:8, pointHitRadius:24, pointBorderWidth:2, pointBorderColor:'#fff' };

  mkC('ch-wl-main',{type:'line',data:{labels:d.months,datasets:[
    {label:'유지',...pointCfg,data:d.유지,borderColor:C.blue,backgroundColor:C.blueA,borderWidth:2,tension:.3,fill:true,yAxisID:'y'},
    {label:'CAPA',...pointCfg,data:d.CAPA,borderColor:C.gold,backgroundColor:'transparent',borderWidth:2,tension:.3,yAxisID:'y2'},
    {label:'해지',...pointCfg,data:d.해지,borderColor:C.red,backgroundColor:'transparent',borderWidth:1.5,tension:.3,yAxisID:'y2'},
  ]},options:baseOpts({
    onClick: onWlMonthClick,
    onHover:(e,el)=>{ e.native.target.style.cursor=el.length?'pointer':'default'; },
    scales:{
      x:{grid:{color:'rgba(228,231,240,.6)'},ticks:{maxRotation:45,color:'#8b93b8',font:{size:10}},border:{display:false}},
      y:{grid:{color:'rgba(228,231,240,.6)'},ticks:{color:'#8b93b8',font:{size:10}},title:{display:true,text:'유지(건)',color:'#8b93b8',font:{size:10}},border:{display:false}},
      y2:{position:'right',grid:{display:false},ticks:{color:'#8b93b8',font:{size:10}},title:{display:true,text:'CAPA/해지(건)',color:'#8b93b8',font:{size:10}},border:{display:false}}
    }
  })});

  mkC('ch-wl-net',{type:'bar',data:{labels:d.months,datasets:[{
    label:'순증',data:d.순증,
    backgroundColor:d.순증.map(v=>v>=0?'rgba(16,185,129,0.75)':'rgba(239,68,68,0.75)'),
    borderColor:d.순증.map(v=>v>=0?C.green:C.red),
    borderWidth:0,borderRadius:4,borderSkipped:false,
  }]},options:baseOpts({
    onClick: onWlMonthClick,
    onHover:(e,el)=>{ e.native.target.style.cursor=el.length?'pointer':'default'; },
    plugins:{legend:{display:false}}
  })});

  bar('ch-wl-ch',od.months,[
    {label:'소매',data:od.소매},
    {label:'도매',data:od.도매,color:C.teal,bg:'rgba(6,182,212,0.72)'},
    {label:'디지털',data:od.디지털,color:C.blue,bg:'rgba(59,130,246,0.72)'},
    {label:'B2B',data:od.B2B,color:C.purple,bg:'rgba(139,92,246,0.72)'},
  ],true);

  // 스파크라인
  addSparklines('wl-kpi', {'wl:유지': RAW.wireless.유지, 'wl:CAPA': RAW.wireless.CAPA, 'wl:해지': RAW.wireless.해지, 'wl:순증': RAW.wireless.순증});};

// ── 무선 월별 현황 팝업 ──
function openWlMonthModal(idx){
  const d=gd('wireless'), od=gd('org');
  const month = d.months[idx];
  const prevIdx = idx > 0 ? idx-1 : 0;

  function diffBadge(cur, prv){
    if(!prv) return '';
    const p = ((cur-prv)/Math.abs(prv)*100).toFixed(1);
    const up = cur >= prv;
    return `<span style="font-size:10px;font-weight:700;color:${up?'var(--green)':'var(--red)'};">${up?'▲':'▼'}${Math.abs(p)}%</span>`;
  }
  function row(label, val, unit, cur, prv){
    return `<tr>
      <td style="padding:9px 12px;font-size:12px;color:var(--text2);border-bottom:1px solid var(--border)">${label}</td>
      <td style="padding:9px 12px;font-size:13px;font-weight:700;font-family:var(--mono);text-align:right;border-bottom:1px solid var(--border)">${fmt(val)}<span style="font-size:10px;color:var(--text3);margin-left:3px">${unit}</span></td>
      <td style="padding:9px 12px;text-align:center;border-bottom:1px solid var(--border)">${diffBadge(cur,prv)}</td>
    </tr>`;
  }

  const chLabels = ['소매','도매','디지털','B2B','소상공인'];
  const chRows = chLabels.map(ch => {
    const v = od[ch]?.[idx] ?? 0;
    const pv = od[ch]?.[prevIdx] ?? 0;
    return row(ch+' CAPA', v, '건', v, pv);
  }).join('');

  const html = `
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px">
      <div>
        <div style="font-size:18px;font-weight:800;color:var(--text)">${month} 무선 가입자 현황</div>
        <div style="font-size:11px;color:var(--text3);margin-top:2px">전월 대비 증감 포함 · 단위: 건</div>
      </div>
      <button onclick="document.getElementById('wl-month-modal').classList.remove('open')"
        style="background:var(--bg3);border:none;color:var(--text3);font-size:18px;cursor:pointer;border-radius:8px;width:32px;height:32px">✕</button>
    </div>

    <!-- 핵심 KPI -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:16px">
      ${[
        ['유지 가입자', d.유지?.[idx], d.유지?.[prevIdx], '건', 'var(--blue)'],
        ['CAPA (신규+기변)', d.CAPA?.[idx], d.CAPA?.[prevIdx], '건', 'var(--gold)'],
        ['해지', d.해지?.[idx], d.해지?.[prevIdx], '건', 'var(--red)'],
        ['순증', d.순증?.[idx], d.순증?.[prevIdx], '건', d.순증?.[idx]>=0?'var(--green)':'var(--red)'],
      ].map(([label, cur, prv, unit, color])=>`
        <div style="background:var(--bg3);border:1px solid var(--border);border-radius:12px;padding:14px 16px">
          <div style="font-size:10px;color:var(--text3);font-weight:700;letter-spacing:.5px;margin-bottom:6px;text-transform:uppercase">${label}</div>
          <div style="font-size:22px;font-weight:800;font-family:var(--mono);color:${color}">${fmt(cur??0)}<span style="font-size:11px;color:var(--text3);margin-left:4px">${unit}</span></div>
          <div style="font-size:10px;margin-top:4px">${diffBadge(cur??0, prv??0)} <span style="color:var(--text3)">전월 ${fmt(prv??0)}건</span></div>
        </div>
      `).join('')}
    </div>

    <!-- 채널별 CAPA -->
    <div style="font-size:12px;font-weight:700;color:var(--text);margin-bottom:8px">채널별 CAPA</div>
    <div style="background:#fff;border:1px solid var(--border);border-radius:10px;overflow:hidden">
      <table style="width:100%;border-collapse:collapse">
        <thead>
          <tr style="background:var(--bg3)">
            <th style="padding:8px 12px;font-size:10px;color:var(--text3);text-align:left;font-weight:700;letter-spacing:.5px">채널</th>
            <th style="padding:8px 12px;font-size:10px;color:var(--text3);text-align:right;font-weight:700">CAPA (건)</th>
            <th style="padding:8px 12px;font-size:10px;color:var(--text3);text-align:center;font-weight:700">전월비</th>
          </tr>
        </thead>
        <tbody>${chRows}</tbody>
      </table>
    </div>`;

  document.getElementById('wl-month-modal-body').innerHTML = html;
  document.getElementById('wl-month-modal').classList.add('open');
}

// ── 유선 ──
window.renderWired=function(){
  const d=gd('wired');
  document.getElementById('wd-kpi').innerHTML=[
    kpi('유지 (전체)',fmt(last(d.유지_전체)),'건',pct(last(d.유지_전체),d.유지_전체[0]),'blue','인터넷+TV','wd:유지전체'),
    kpi('신규 (최근월)',fmt(last(d.신규_전체)),'건',pct(last(d.신규_전체),prev(d.신규_전체)),'teal','유선 전체','wd:신규전체'),
    kpi('해지 (최근월)',fmt(last(d.해지_전체)),'건',pct(last(d.해지_전체),prev(d.해지_전체)),'red','유선 전체','wd:해지전체'),
    kpi('순증 (최근월)',fmt(last(d.순증_전체)),'건',null,last(d.순증_전체)>=0?'green':'red','유선 전체','wd:순증전체'),
    kpi('인터넷 유지',fmt(last(d.유지_인터넷)),'건',pct(last(d.유지_인터넷),d.유지_인터넷[0]),'gold','인터넷 전용','wd:유지인터넷'),
    kpi('인터넷 신규 (최근월)',fmt(last(d.신규_인터넷)),'건',null,'green','인터넷 전용','wd:신규인터넷'),
  ].join('');
  mkC('ch-wd-main',{type:'line',data:{labels:d.months,datasets:[{label:'유지(전체)',data:d.유지_전체,borderColor:C.blue,backgroundColor:C.blueA,borderWidth:2,pointRadius:0,tension:.3,fill:true,yAxisID:'y'},{label:'신규',data:d.신규_전체,borderColor:C.teal,backgroundColor:'transparent',borderWidth:2,pointRadius:2,tension:.3,yAxisID:'y2'},{label:'해지',data:d.해지_전체,borderColor:C.red,backgroundColor:'transparent',borderWidth:1.5,pointRadius:0,tension:.3,yAxisID:'y2'}]},options:baseOpts({scales:{x:{grid:{color:'rgba(42,53,85,0.5)'},ticks:{maxRotation:45,color:'#5c6e9a',font:{size:10}}},y:{grid:{color:'rgba(42,53,85,0.5)'},ticks:{color:'#5c6e9a'},title:{display:true,text:'유지(건)',color:'#5c6e9a'}},y2:{position:'right',grid:{display:false},ticks:{color:'#5c6e9a'},title:{display:true,text:'신규/해지(건)',color:'#5c6e9a'}}}})});
  mkC('ch-wd-net',{type:'bar',data:{labels:d.months,datasets:[{label:'순증(전체)',data:d.순증_전체,backgroundColor:d.순증_전체.map(v=>v>=0?C.greenA:C.redA),borderColor:d.순증_전체.map(v=>v>=0?C.green:C.red),borderWidth:1.5}]},options:baseOpts()});
  mkC('ch-wd-inet',{type:'line',data:{labels:d.months,datasets:[{label:'인터넷유지',data:d.유지_인터넷,borderColor:C.gold,backgroundColor:C.goldA,borderWidth:2,pointRadius:0,tension:.3,fill:true,yAxisID:'y'},{label:'신규',data:d.신규_인터넷,borderColor:C.teal,backgroundColor:'transparent',borderWidth:2,pointRadius:2,tension:.3,yAxisID:'y2'},{label:'해지',data:d.해지_인터넷,borderColor:C.red,backgroundColor:'transparent',borderWidth:1.5,pointRadius:0,tension:.3,yAxisID:'y2'}]},options:baseOpts({scales:{x:{grid:{color:'rgba(42,53,85,0.5)'},ticks:{maxRotation:45,color:'#5c6e9a',font:{size:10}}},y:{grid:{color:'rgba(42,53,85,0.5)'},ticks:{color:'#5c6e9a'},title:{display:true,text:'유지(건)',color:'#5c6e9a'}},y2:{position:'right',grid:{display:false},ticks:{color:'#5c6e9a'},title:{display:true,text:'신규/해지(건)',color:'#5c6e9a'}}}})});
};

// ── 조직별 ──

// ── 도넛 차트 헬퍼 (중앙 텍스트 포함, 정사각형 보장) ──
function makeDonut(canvasId, labels, data, colors, centerLabel='', centerSub=''){
  const el = document.getElementById(canvasId);
  if(!el) return;

  // 기존 차트 제거
  if(S.charts[canvasId]){ S.charts[canvasId].destroy(); delete S.charts[canvasId]; }

  // canvas 크기 고정
  el.width = 240; el.height = 240;
  el.style.width = '240px'; el.style.height = '240px';

  // 중앙 텍스트 플러그인
  const centerPlugin = {
    id: 'centerText_' + canvasId,
    afterDraw(chart){
      if(!centerLabel) return;
      const {ctx, chartArea:{top,bottom,left,right}} = chart;
      const cx = (left+right)/2, cy = (top+bottom)/2;
      ctx.save();
      ctx.textAlign='center'; ctx.textBaseline='middle';
      ctx.font = 'bold 15px var(--font, sans-serif)';
      ctx.fillStyle = '#1a1f36';
      ctx.fillText(centerLabel, cx, centerSub ? cy-9 : cy);
      if(centerSub){
        ctx.font = '11px var(--font, sans-serif)';
        ctx.fillStyle = '#8b93b8';
        ctx.fillText(centerSub, cx, cy+10);
      }
      ctx.restore();
    }
  };

  S.charts[canvasId] = new Chart(el, {
    type:'doughnut',
    data:{labels, datasets:[{data, backgroundColor:colors, borderColor:'#fff', borderWidth:2, hoverOffset:4}]},
    options:{
      animation:false, animations:false,
      responsive:false, maintainAspectRatio:false,
      cutout:'62%',
      plugins:{
        legend:{position:'bottom',labels:{boxWidth:8,padding:10,color:'#4a5380',font:{size:10},usePointStyle:true}},
        tooltip:{backgroundColor:'#fff',titleColor:'#1a1f36',bodyColor:'#4a5380',borderColor:'#e4e7f0',borderWidth:1,cornerRadius:10,
          callbacks:{label:ctx=>{
            const total=ctx.dataset.data.reduce((a,b)=>a+b,0);
            const pct=total?(ctx.raw/total*100).toFixed(1):0;
            return ` ${ctx.label}: ${typeof ctx.raw==='number'&&ctx.raw>100?fmt(ctx.raw):ctx.raw} (${pct}%)`;
          }}
        }
      }
    },
    plugins:[centerPlugin]
  });
}
window.renderOrg=function(){
  const d=gd('org'),chs=['소매','도매','디지털','B2B','소상공인'];
  const tots=chs.map(c=>sum(d[c]||[])),grand=sum(tots);
  // 채널 순위 정렬
  const orgRanked = chs.map((c,i)=>({c,i,v:tots[i]})).sort((a,b)=>b.v-a.v);
  const orgRankMap = {}; orgRanked.forEach((x,r)=>orgRankMap[x.c]=r+1);
  document.getElementById('org-kpi').innerHTML=chs.map((c,i)=>{
    const rank=orgRankMap[c];
    const rankBadge=rank===1?'🥇':rank===2?'🥈':rank===3?'🥉':'';
    return kpi(`${rankBadge}${c} CAPA`,fmt(tots[i]),'건',null,['gold','teal','blue','purple','orange'][i],`비중 ${grand?((tots[i]/grand*100).toFixed(1)):0}%`,`org:${c}`);
  }).join('');
  addSparklines('org-kpi', {'org:소매':RAW.org.소매,'org:도매':RAW.org.도매,'org:디지털':RAW.org.디지털,'org:B2B':RAW.org.B2B,'org:소상공인':RAW.org.소상공인});
  bar('ch-org-bar',d.months,chs.map((c,i)=>({label:c,data:d[c]||[],color:COLORS[i],bg:COLORS_A[i]})),true);
  const orgColors=['rgba(245,158,11,0.8)','rgba(6,182,212,0.8)','rgba(59,130,246,0.8)','rgba(139,92,246,0.8)','rgba(249,115,22,0.8)'];
  makeDonut('ch-org-pie', chs, tots, orgColors, fmt(grand), 'CAPA 합계');
  line('ch-org-trend',d.months,[{label:'소매',data:d.소매},{label:'도매',data:d.도매,color:C.teal},{label:'디지털',data:d.디지털,color:C.blue}]);
};

// ── 디지털 ──
window.renderDigital=function(){
  const d=gd('digital');
  document.getElementById('dig-kpi').innerHTML=[
    kpi('일반후불 (최근월)',fmt(last(d.일반후불_총계)),'건',pct(last(d.일반후불_총계),prev(d.일반후불_총계)),'gold','KT닷컴 전체','dig:일반후불'),
    kpi('KT닷컴 직접 (최근월)',fmt(last(d['일반후불_KT닷컴'])),'건',pct(last(d['일반후불_KT닷컴']),prev(d['일반후불_KT닷컴'])),'teal','','dig:KT닷컴'),
    kpi('운영후불 (최근월)',fmt(last(d.운영후불_총계)),'건',pct(last(d.운영후불_총계),prev(d.운영후불_총계)),'blue','','dig:운영후불'),
    kpi('유심단독 (최근월)',fmt(last(d.유심단독)),'건',null,'purple','','dig:유심단독'),
    kpi('유선순신규 (최근월)',fmt(last(d.유선순신규)),'건',pct(last(d.유선순신규),prev(d.유선순신규)),'green','','dig:유선순신규'),
    kpi('디지털채널 인력',fmt(last(d.인력_총계)),'명',pct(last(d.인력_총계),prev(d.인력_총계)),'orange',''),
  ].join('');
  bar('ch-dig-main',d.months,[{label:'KT닷컴 직접',data:d['일반후불_KT닷컴']},{label:'O2O',data:d['일반후불_O2O'],color:C.teal,bg:C.tealA},{label:'온라인유통',data:d['일반후불_온라인유통'],color:C.blue,bg:C.blueA}],true);
  line('ch-dig-type',d.months,[{label:'운영후불',data:d.운영후불_총계},{label:'유심단독',data:d.유심단독,color:C.purple}]);
  mkC('ch-dig-etc',{type:'line',data:{labels:d.months,datasets:[{label:'유선순신규',data:d.유선순신규,borderColor:C.teal,backgroundColor:'transparent',borderWidth:2,pointRadius:2,tension:.3,yAxisID:'y'},{label:'인력',data:d.인력_총계,borderColor:C.orange,backgroundColor:'transparent',borderWidth:1.5,pointRadius:0,tension:.3,yAxisID:'y2'}]},options:baseOpts({scales:{x:{grid:{color:'rgba(42,53,85,0.5)'},ticks:{maxRotation:45,color:'#5c6e9a',font:{size:10}}},y:{grid:{color:'rgba(42,53,85,0.5)'},ticks:{color:'#5c6e9a'},title:{display:true,text:'유선순신규(명)',color:'#5c6e9a'}},y2:{position:'right',grid:{display:false},ticks:{color:'#5c6e9a'},title:{display:true,text:'인력(명)',color:'#5c6e9a'}}}})});
};

// ── B2B ──
window.renderB2B=function(){
  const d=gd('b2b');
  document.getElementById('b2b-kpi').innerHTML=[
    kpi('일반후불 (최근월)',fmt(last(d.전체_일반후불)),'건',pct(last(d.전체_일반후불),prev(d.전체_일반후불)),'gold','B2B 전체','b2b:일반후불'),
    kpi('후불신규 (최근월)',fmt(last(d.전체_후불신규)),'건',pct(last(d.전체_후불신규),prev(d.전체_후불신규)),'teal','','b2b:후불신규'),
    kpi('무선순증 (최근월)',fmt(last(d.전체_무선순증)),'건',null,last(d.전체_무선순증)>=0?'green':'red','','b2b:무선순증'),
    kpi('무선유지 가입자',fmt(last(d.가입자_무선유지)),'건',pct(last(d.가입자_무선유지),d.가입자_무선유지[0]),'blue','B2B 전체','b2b:무선유지'),
    kpi('기업 무선유지',fmt(last(d.가입자_기업무선)),'건',pct(last(d.가입자_기업무선),d.가입자_기업무선[0]),'purple','기업채널','b2b:기업무선'),
    kpi('법인 무선유지',fmt(last(d.가입자_법인무선)),'건',pct(last(d.가입자_법인무선),d.가입자_법인무선[0]),'orange','법인채널','b2b:법인무선'),
  ].join('');
  bar('ch-b2b-main',d.months,[{label:'일반후불',data:d.전체_일반후불},{label:'운영후불',data:d.전체_운영후불,color:C.teal,bg:C.tealA},{label:'후불신규',data:d.전체_후불신규,color:C.blue,bg:C.blueA}],true);
  line('ch-b2b-seg',d.months,[{label:'기업 일반후불',data:d.기업_일반후불},{label:'법인 일반후불',data:d.법인_일반후불,color:C.teal}]);
  line('ch-b2b-sub',d.months,[{label:'B2B 전체',data:d.가입자_무선유지},{label:'기업',data:d.가입자_기업무선,color:C.teal},{label:'법인',data:d.가입자_법인무선,color:C.blue}]);
};

// ── 소상공인 ──
window.renderSMB=function(){
  const d=gd('smb');
  document.getElementById('smb-kpi').innerHTML=[
    kpi('일반후불 (최근월)',fmt(last(d.상품_일반후불)),'건',pct(last(d.상품_일반후불),prev(d.상품_일반후불)),'gold','','smb:일반후불'),
    kpi('운영후불 (최근월)',fmt(last(d.상품_운영후불)),'건',pct(last(d.상품_운영후불),prev(d.상품_운영후불)),'teal','','smb:운영후불'),
    kpi('인터넷순신규 (최근월)',fmt(last(d.인터넷순신규)),'건',pct(last(d.인터넷순신규),prev(d.인터넷순신규)),'blue','','smb:인터넷순신규'),
    kpi('기간 인터넷순신규 합계',fmt(sum(d.인터넷순신규)),'건',null,'green',''),
    kpi('전담 인력 (최근월)',fmt(last(d.인력_총원)),'명',pct(last(d.인력_총원),d.인력_총원[0]),'purple','소상공인 채널'),
    kpi('기간 일반후불 합계',fmt(sum(d.상품_일반후불)),'건',null,'orange',''),
  ].join('');
  bar('ch-smb-main',d.months,[{label:'일반후불',data:d.상품_일반후불},{label:'운영후불',data:d.상품_운영후불,color:C.teal,bg:C.tealA},{label:'인터넷순신규',data:d.인터넷순신규,color:C.blue,bg:C.blueA}],false);
  bar('ch-smb-hq',d.months,[{label:'강북',data:d.본부별_강북},{label:'강남',data:d.본부별_강남,color:C.teal,bg:C.tealA},{label:'강서',data:d.본부별_강서,color:C.blue,bg:C.blueA},{label:'동부',data:d.본부별_동부,color:C.purple,bg:C.purpleA},{label:'서부',data:d.본부별_서부,color:C.orange,bg:C.orangeA}],true);
  line('ch-smb-hr',d.months,[{label:'전담 인력',data:d.인력_총원,fill:true}]);
};

// ── 유통플랫폼 ──
window.renderPlatform=function(){
  const d=gd('platform');
  document.getElementById('plat-kpi').innerHTML=[
    kpi('중고폰 매입건수 (최근월)',fmt(last(d.중고폰_매입건수)),'건',pct(last(d.중고폰_매입건수),prev(d.중고폰_매입건수)),'gold',''),
    kpi('중고폰 매각건수 (최근월)',fmt(last(d.중고폰_매각건수)),'건',pct(last(d.중고폰_매각건수),prev(d.중고폰_매각건수)),'teal',''),
    kpi('중고폰 매입금액 (최근월)',fmt(last(d.중고폰_매입금액),1),'억원',null,'blue',''),
    kpi('기간 매입 합계',fmt(sum(d.중고폰_매입건수)),'건',null,'purple','중고폰'),
    kpi('시연폰 매각건수 (최근월)',fmt(last(d.시연폰_매각건수)),'건',null,'green',''),
    kpi('시연폰 매각이익 (최근월)',fmt(last(d.시연폰_매각이익),1),'억원',null,'orange',''),
  ].join('');
  line('ch-plat-main',d.months,[{label:'중고폰 매입',data:d.중고폰_매입건수,fill:true},{label:'중고폰 매각',data:d.중고폰_매각건수,color:C.teal}]);
  line('ch-plat-demo',d.months,[{label:'시연폰 매입',data:d.시연폰_매입건수,fill:true},{label:'시연폰 매각',data:d.시연폰_매각건수,color:C.teal}]);
  line('ch-plat-amt',d.months,[{label:'중고폰 매입금액(억)',data:d.중고폰_매입금액,fill:true},{label:'시연폰 매각이익(억)',data:d.시연폰_매각이익,color:C.teal}]);
};

// ── 품질 ──
window.renderQuality=function(){
  const tc=gd('tcsi'),vc=gd('voc');
  document.getElementById('qual-kpi').innerHTML=[
    kpi('TCSI 점수 (최근월)',fmt(last(tc.TCSI점수),1),'점',pct(last(tc.TCSI점수),tc.TCSI점수[0]),'gold',`KT점수 ${fmt(last(tc.KT점수),1)}`,'tc:TCSI점수'),
    kpi('KT 점수',fmt(last(tc.KT점수),1),'점',null,'teal',`대리점 ${fmt(last(tc.대리점점수),1)}`,'tc:KT점수'),
    kpi('VOC 도+소매 발생률',fmt(last(vc.도소매발생률),4),'%',pct(last(vc.도소매발생률),vc.도소매발생률[0]),'green','낮을수록 우수','voc:도소매발생률'),
    kpi('VOC 소매 발생률',fmt(last(vc.소매발생률),4),'%',null,'blue','','voc:소매발생률'),
    kpi('대외민원 건수 (최근월)',fmt(last(vc.대외민원_건수)),'건',pct(last(vc.대외민원_건수),prev(vc.대외민원_건수)),'red','','voc:대외민원'),
    kpi('대외민원 귀책률 (최근월)',fmt(last(vc.대외민원_귀책률),1),'%',null,'orange',''),
  ].join('');
  mkC('ch-tcsi',{type:'line',data:{labels:tc.months,datasets:[{label:'TCSI 전체',data:tc.TCSI점수,borderColor:C.gold,backgroundColor:C.goldA,borderWidth:2.5,pointRadius:2,tension:.3,fill:true},{label:'KT점수',data:tc.KT점수,borderColor:C.teal,backgroundColor:'transparent',borderWidth:2,pointRadius:0,tension:.3},{label:'대리점점수',data:tc.대리점점수,borderColor:C.blue,backgroundColor:'transparent',borderWidth:1.5,pointRadius:0,tension:.3}]},options:baseOpts({scales:{x:{grid:{color:'rgba(42,53,85,0.5)'},ticks:{maxRotation:45,color:'#5c6e9a',font:{size:10}}},y:{min:88,grid:{color:'rgba(42,53,85,0.5)'},ticks:{color:'#5c6e9a'}}}})});
  line('ch-voc',vc.months,[{label:'도+소매',data:vc.도소매발생률},{label:'소매',data:vc.소매발생률,color:C.teal},{label:'도매',data:vc.도매발생률,color:C.blue},{label:'디지털',data:vc.디지털발생률,color:C.purple}]);
  mkC('ch-민원',{type:'bar',data:{labels:vc.months,datasets:[{label:'대외민원 건수',data:vc.대외민원_건수,backgroundColor:C.redA,borderColor:C.red,borderWidth:1.5,yAxisID:'y'},{label:'귀책률(%)',data:vc.대외민원_귀책률,type:'line',borderColor:C.orange,backgroundColor:'transparent',borderWidth:2,pointRadius:0,tension:.3,yAxisID:'y2'}]},options:baseOpts({scales:{x:{grid:{color:'rgba(42,53,85,0.5)'},ticks:{maxRotation:45,color:'#5c6e9a',font:{size:10}}},y:{grid:{color:'rgba(42,53,85,0.5)'},ticks:{color:'#5c6e9a'},title:{display:true,text:'건수',color:'#5c6e9a'}},y2:{position:'right',grid:{display:false},ticks:{color:'#5c6e9a'},title:{display:true,text:'귀책률(%)',color:'#5c6e9a'}}}})});
  // Region tables
  const rEl=document.getElementById('region-tables');if(!rEl)return;
  const tcsiRegions=Object.keys(tc.regions||RAW.tcsi.regions||{});
  const vocRegions=Object.keys(vc.소매_지역||{});
  rEl.innerHTML=`
    <div class="region-table"><h3>🗺️ 소매 지역별 TCSI 점수 (최근월)</h3>
      <table class="data-table"><thead><tr><th>지역</th><th class="num">점수</th><th class="num">전월비</th></tr></thead><tbody>
      ${tcsiRegions.map(r=>{const v=last(tc.regions[r]),p=prev(tc.regions[r]),d=v-p;return`<tr><td>${r}</td><td class="num">${fmt(v,1)}</td><td class="num ${d>=0?'pos':'neg'}">${d>=0?'+':''}${fmt(d,2)}</td></tr>`}).join('')}
      </tbody></table>
    </div>
    <div class="region-table"><h3>📊 소매 지역별 VOC 발생률 (최근월)</h3>
      <table class="data-table"><thead><tr><th>지역</th><th class="num">발생률(%)</th><th class="num">전월비</th></tr></thead><tbody>
      ${vocRegions.map(r=>{const v=last(vc.소매_지역[r]),p=prev(vc.소매_지역[r]),d=v-p;return`<tr><td>${r}</td><td class="num">${fmt(v,4)}</td><td class="num ${d<=0?'pos':'neg'}">${d>=0?'+':''}${fmt(d,4)}</td></tr>`}).join('')}
      </tbody></table>
    </div>`;
};

// ── 인프라 ──
window.renderInfra=function(){
  const d=gd('infra');
  document.getElementById('infra-kpi').innerHTML=[
    kpi('소매 매장수',fmt(last(d.소매매장수_계)),'개',pct(last(d.소매매장수_계),d.소매매장수_계[0]),'gold','매장 Infra 계','inf:소매매장수'),
    kpi('출점 (기간합)',fmt(sum(d.출점_계)),'개',null,'green',''),
    kpi('퇴점 (기간합)',fmt(sum(d.퇴점_계)),'개',null,'red',''),
    kpi('도매 무선취급점',fmt(last(d.도매무선취급점)),'개',pct(last(d.도매무선취급점),d.도매무선취급점[0]),'blue','','inf:도매무선'),
    kpi('점당생산성 무선 (최근월)',fmt(last(d.점당생산성_무선),1),'건/점',null,'teal','','inf:점당생산성'),
    kpi('인당생산성 무선 (최근월)',fmt(last(d.인당생산성_무선),1),'건/인',null,'purple','','inf:인당생산성'),
  ].join('');
  mkC('ch-infra-main',{type:'bar',data:{labels:d.months,datasets:[{label:'매장수',data:d.소매매장수_계,backgroundColor:C.blueA,borderColor:C.blue,borderWidth:1.5,yAxisID:'y'},{label:'출점',data:d.출점_계,type:'line',borderColor:C.green,backgroundColor:'transparent',borderWidth:2,pointRadius:2,tension:.3,yAxisID:'y2'},{label:'퇴점',data:d.퇴점_계,type:'line',borderColor:C.red,backgroundColor:'transparent',borderWidth:1.5,pointRadius:2,tension:.3,yAxisID:'y2'}]},options:baseOpts({scales:{x:{grid:{color:'rgba(42,53,85,0.5)'},ticks:{maxRotation:45,color:'#5c6e9a',font:{size:10}}},y:{grid:{color:'rgba(42,53,85,0.5)'},ticks:{color:'#5c6e9a'},title:{display:true,text:'매장수(개)',color:'#5c6e9a'}},y2:{position:'right',grid:{display:false},ticks:{color:'#5c6e9a'},title:{display:true,text:'출/퇴점(개)',color:'#5c6e9a'}}}})});
  line('ch-infra-prod',d.months,[{label:'점당생산성(무선)',data:d.점당생산성_무선,fill:true},{label:'인당생산성(무선)',data:d.인당생산성_무선,color:C.teal}]);
  line('ch-infra-wholesale',d.months,[{label:'무선취급점',data:d.도매무선취급점,fill:true},{label:'유선취급점',data:d.도매유선취급점,color:C.teal}]);
  const hqs=['강북','강남','강서','동부','서부'];
  bar('ch-infra-hq',d.months,hqs.map((h,i)=>({label:h,data:d[`소매_${h}`]||[],color:COLORS[i],bg:COLORS_A[i]})),true);
  mkC('ch-infra-wnet',{type:'bar',data:{labels:d.months,datasets:[{label:'도매 순증(무선)',data:d.도매순증_무선,backgroundColor:d.도매순증_무선.map(v=>v>=0?C.greenA:C.redA),borderColor:d.도매순증_무선.map(v=>v>=0?C.green:C.red),borderWidth:1.5}]},options:baseOpts()});
  line('ch-infra-wline',d.months,[{label:'유선순신규 점당생산성',data:d.점당생산성_유선,fill:true}]);
};

// ── 전략상품 ──
window.renderStrategy=function(){
  const d=gd('strategy');
  document.getElementById('str-kpi').innerHTML=[
    kpi('하이오더 점포수 (최근월)',fmt(last(d.하이오더_점포수)),'개',pct(last(d.하이오더_점포수),d.하이오더_점포수[0]),'gold',''),
    kpi('하이오더 태블릿수 (최근월)',fmt(last(d.하이오더_태블릿수)),'대',null,'teal',''),
    kpi('소상공인 점포수 (최근월)',fmt(last(d.소상공인_점포수)),'개',null,'blue',''),
    kpi('GiGAeyes (최근월)',fmt(last(d.GiGAeyes_계)),'대',pct(last(d.GiGAeyes_계),d.GiGAeyes_계[0]),'purple','소상공인 포함'),
    kpi('AI전화 (최근월)',fmt(last(d.AI전화_계)),'대',null,'green',''),
    kpi('스마트로VAN (최근월)',fmt(last(d.스마트로VAN_계)),'대',null,'orange',''),
  ].join('');
  bar('ch-str-main',d.months,[{label:'하이오더 점포',data:d.하이오더_점포수},{label:'소상공인 점포',data:d.소상공인_점포수,color:C.teal,bg:C.tealA},{label:'매장 점포',data:d.매장_점포수,color:C.blue,bg:C.blueA}],true);
  line('ch-str-giga',d.months,[{label:'GiGAeyes 전체',data:d.GiGAeyes_계,fill:true},{label:'소상공인',data:d.GiGAeyes_소상공인,color:C.teal}]);
  line('ch-str-etc',d.months,[{label:'로봇',data:d.로봇_계},{label:'AI전화',data:d.AI전화_계,color:C.teal},{label:'스마트로VAN',data:d.스마트로VAN_계,color:C.blue}]);
};

// ── 인력 ──

// ════════════════════════════════════════════════
//  범용 KPI 팝업 — 모든 KPI 카드 클릭 시 호출
//  key 형식: 'fin:총매출' | 'org:소매' | 'wl:유지' 등
// ════════════════════════════════════════════════
function openKpiPopup(key){
  try{
    const parts=key.split(':'), ns=parts[0];
    const FIN=RAW.finance, ORG=RAW.org, PLAT=RAW.platform;
    const platArr=(PLAT.시연폰_매각이익||[]).map((v,i)=>(v||0)+((PLAT.중고폰_매입금액||[])[i]||0));

    const FIN=RAW.finance, ORG=RAW.org, PLAT=RAW.platform;
    const WL=RAW.wireless, WD=RAW.wired;
    const DIG=RAW.digital, B2B=RAW.b2b, SMB=RAW.smb;
    const TC=RAW.tcsi, VOC=RAW.voc, HR=RAW.hr, INF=RAW.infra;
    const platArr=(PLAT.시연폰_매각이익||[]).map((v,i)=>(v||0)+((PLAT.중고폰_매입금액||[])[i]||0));

    const MAP={
      // 재무
      'fin:총매출':    {arr:FIN.총매출,   months:FIN.months,label:'총매출',      unit:'억원',color:'#f59e0b',icon:'💰'},
      'fin:통신매출':  {arr:FIN.통신매출, months:FIN.months,label:'통신매출',    unit:'억원',color:'#3b82f6',icon:'📡'},
      'fin:영업이익':  {arr:FIN.영업이익, months:FIN.months,label:'영업이익',    unit:'억원',color:'#10b981',icon:'📈'},
      'fin:판관비':    {arr:FIN.판관비,   months:FIN.months,label:'판관비',      unit:'억원',color:'#8b5cf6',icon:'💼'},
      'fin:마케팅비':  {arr:FIN.마케팅비, months:FIN.months,label:'마케팅비',    unit:'억원',color:'#f97316',icon:'📢'},
      'fin:인건비':    {arr:FIN.인건비,   months:FIN.months,label:'인건비',      unit:'억원',color:'#ec4899',icon:'👥'},
      'fin:유통플랫폼':{arr:platArr,      months:FIN.months,label:'유통플랫폼매출',unit:'억원',color:'#06b6d4',icon:'♻️'},
      // 무선
      'wl:유지':       {arr:WL.유지,   months:WL.months,label:'무선 유지가입자',unit:'건',color:'#3b82f6',icon:'📱'},
      'wl:CAPA':       {arr:WL.CAPA,   months:WL.months,label:'무선 CAPA',     unit:'건',color:'#f59e0b',icon:'📶'},
      'wl:해지':       {arr:WL.해지,   months:WL.months,label:'무선 해지',     unit:'건',color:'#ef4444',icon:'❌'},
      'wl:순증':       {arr:WL.순증,   months:WL.months,label:'무선 순증',     unit:'건',color:'#10b981',icon:'📊'},
      // 유선
      'wd:유지전체':   {arr:WD.유지_전체,  months:WD.months,label:'유선 유지(전체)',unit:'건',color:'#3b82f6',icon:'🌐'},
      'wd:신규전체':   {arr:WD.신규_전체,  months:WD.months,label:'유선 신규',     unit:'건',color:'#10b981',icon:'➕'},
      'wd:해지전체':   {arr:WD.해지_전체,  months:WD.months,label:'유선 해지',     unit:'건',color:'#ef4444',icon:'➖'},
      'wd:순증전체':   {arr:WD.순증_전체,  months:WD.months,label:'유선 순증',     unit:'건',color:'#8b5cf6',icon:'📊'},
      'wd:유지인터넷': {arr:WD.유지_인터넷,months:WD.months,label:'인터넷 유지',   unit:'건',color:'#f59e0b',icon:'💻'},
      'wd:신규인터넷': {arr:WD.신규_인터넷,months:WD.months,label:'인터넷 신규',   unit:'건',color:'#06b6d4',icon:'🔌'},
      // 조직별
      'org:소매':      {arr:ORG.소매,    months:ORG.months,label:'소매 CAPA',   unit:'건',color:'#f59e0b',icon:'🏪'},
      'org:도매':      {arr:ORG.도매,    months:ORG.months,label:'도매 CAPA',   unit:'건',color:'#06b6d4',icon:'🏭'},
      'org:디지털':    {arr:ORG.디지털,  months:ORG.months,label:'디지털 CAPA', unit:'건',color:'#3b82f6',icon:'💻'},
      'org:B2B':       {arr:ORG.B2B,    months:ORG.months,label:'B2B CAPA',    unit:'건',color:'#8b5cf6',icon:'🤝'},
      'org:소상공인':  {arr:ORG.소상공인,months:ORG.months,label:'소상공인 CAPA',unit:'건',color:'#f97316',icon:'🛒'},
      // 디지털
      'dig:일반후불':  {arr:DIG.일반후불_총계,months:DIG.months,label:'디지털 일반후불',unit:'건',color:'#f59e0b',icon:'📱'},
      'dig:KT닷컴':    {arr:DIG['일반후불_KT닷컴'],months:DIG.months,label:'KT닷컴 직접',unit:'건',color:'#3b82f6',icon:'🌐'},
      'dig:운영후불':  {arr:DIG.운영후불_총계,months:DIG.months,label:'디지털 운영후불',unit:'건',color:'#06b6d4',icon:'⚙️'},
      'dig:유심단독':  {arr:DIG.유심단독,months:DIG.months,label:'유심단독',unit:'건',color:'#8b5cf6',icon:'💳'},
      'dig:유선순신규':{arr:DIG.유선순신규,months:DIG.months,label:'디지털 유선순신규',unit:'건',color:'#10b981',icon:'📡'},
      // B2B
      'b2b:일반후불':  {arr:B2B.전체_일반후불,months:B2B.months,label:'B2B 일반후불',unit:'건',color:'#f59e0b',icon:'🏢'},
      'b2b:후불신규':  {arr:B2B.전체_후불신규,months:B2B.months,label:'B2B 후불신규',unit:'건',color:'#06b6d4',icon:'➕'},
      'b2b:무선순증':  {arr:B2B.전체_무선순증,months:B2B.months,label:'B2B 무선순증',unit:'건',color:'#10b981',icon:'📊'},
      'b2b:무선유지':  {arr:B2B.가입자_무선유지,months:B2B.months,label:'B2B 무선유지',unit:'건',color:'#3b82f6',icon:'📱'},
      'b2b:기업무선':  {arr:B2B.가입자_기업무선,months:B2B.months,label:'기업 무선유지',unit:'건',color:'#8b5cf6',icon:'🤝'},
      'b2b:법인무선':  {arr:B2B.가입자_법인무선,months:B2B.months,label:'법인 무선유지',unit:'건',color:'#f97316',icon:'🏛️'},
      // 소상공인
      'smb:일반후불':  {arr:SMB.상품_일반후불,months:SMB.months,label:'소상공인 일반후불',unit:'건',color:'#f59e0b',icon:'🛒'},
      'smb:운영후불':  {arr:SMB.상품_운영후불,months:SMB.months,label:'소상공인 운영후불',unit:'건',color:'#06b6d4',icon:'⚙️'},
      'smb:인터넷순신규':{arr:SMB.인터넷순신규,months:SMB.months,label:'소상공인 인터넷순신규',unit:'건',color:'#10b981',icon:'🔌'},
      // TCSI
      'tc:TCSI점수':   {arr:TC.TCSI점수, months:TC.months,label:'TCSI 점수',  unit:'점',color:'#3b82f6',icon:'⭐'},
      'tc:KT점수':     {arr:TC.KT점수,   months:TC.months,label:'KT 점수',    unit:'점',color:'#10b981',icon:'📊'},
      'tc:대리점점수': {arr:TC.대리점점수,months:TC.months,label:'대리점 점수',unit:'점',color:'#f59e0b',icon:'🏪'},
      // VOC
      'voc:도소매발생률':{arr:VOC.도소매발생률,months:VOC.months,label:'도+소매 발생률',unit:'%',color:'#ef4444',icon:'📢'},
      'voc:소매발생률': {arr:VOC.소매발생률, months:VOC.months,label:'소매 발생률',    unit:'%',color:'#f97316',icon:'📢'},
      'voc:도매발생률': {arr:VOC.도매발생률, months:VOC.months,label:'도매 발생률',    unit:'%',color:'#8b5cf6',icon:'📢'},
      'voc:대외민원':   {arr:VOC.대외민원_건수,months:VOC.months,label:'대외민원 건수',unit:'건',color:'#ef4444',icon:'⚠️'},
      // 인프라
      'inf:소매매장수': {arr:INF.소매매장수_계,months:INF.months,label:'소매 매장수',    unit:'개',color:'#f59e0b',icon:'🏪'},
      'inf:도매무선':   {arr:INF.도매무선취급점,months:INF.months,label:'도매 무선취급점',unit:'개',color:'#3b82f6',icon:'🏭'},
      'inf:도매유선':   {arr:INF.도매유선취급점,months:INF.months,label:'도매 유선취급점',unit:'개',color:'#06b6d4',icon:'🏭'},
      'inf:점당생산성': {arr:INF.점당생산성_무선,months:INF.months,label:'무선 점당생산성',unit:'건',color:'#10b981',icon:'📊'},
      'inf:인당생산성': {arr:INF.인당생산성_무선,months:INF.months,label:'무선 인당생산성',unit:'건',color:'#8b5cf6',icon:'👤'},
      // 인력
      'hr:전사계':      {arr:HR.전사계,  months:HR.months,label:'전사 인원',  unit:'명',color:'#3b82f6',icon:'👥'},
      'hr:영업직':      {arr:HR.영업직,  months:HR.months,label:'영업직 인원',unit:'명',color:'#10b981',icon:'👤'},
      'hr:SC직':        {arr:HR.SC직,    months:HR.months,label:'SC직 인원',  unit:'명',color:'#f59e0b',icon:'👤'},
    };

    const info=MAP[key];;

    const info=MAP[key];
    if(!info||!info.arr||!info.arr.length){showToast('⚠ 데이터 없음');return;}

    const {arr,months,label,unit,color,icon}=info;
    const valid=arr.filter(v=>v!=null&&!isNaN(v));
    if(!valid.length){showToast('⚠ 유효 데이터 없음');return;}

    const maxV=Math.max(...valid),minV=Math.min(...valid),avgV=valid.reduce((s,v)=>s+v,0)/valid.length;
    const maxI=arr.findIndex(v=>v===maxV),minI=arr.findIndex(v=>v===minV);

    // QoQ: 이번 분기 평균 vs 직전 분기 평균
    const getQtr=i=>{const q=Math.floor(i/3);return arr.slice(q*3,(q+1)*3).filter(v=>v!=null);};
    const curQtr=Math.floor((arr.length-1)/3);
    const cqVals=arr.slice(curQtr*3).filter(v=>v!=null);
    const pqVals=curQtr>0?arr.slice((curQtr-1)*3,curQtr*3).filter(v=>v!=null):[];
    const cqAvg=cqVals.length?cqVals.reduce((s,v)=>s+v,0)/cqVals.length:0;
    const pqAvg=pqVals.length?pqVals.reduce((s,v)=>s+v,0)/pqVals.length:0;
    const qoqPct=pqAvg?((cqAvg-pqAvg)/Math.abs(pqAvg)*100):0;

    // YoY: 올해 연평균 vs 작년 연평균
    const getYearPart=m=>(m||'').match(/^['\s]*(\d{2,4})[.년]/)?.[1]||'';
    const yMap={};
    months.forEach((m,i)=>{
      const y=getYearPart(m); if(!y)return;
      if(!yMap[y]){yMap[y]={s:0,n:0,vals:[]};}
      if(arr[i]!=null){yMap[y].s+=arr[i];yMap[y].n++;yMap[y].vals.push(arr[i]);}
    });
    const yKeys=Object.keys(yMap).sort();
    const curY=yKeys[yKeys.length-1], prevY=yKeys[yKeys.length-2];
    const cyAvg=yMap[curY]&&yMap[curY].n?yMap[curY].s/yMap[curY].n:0;
    const pyAvg=yMap[prevY]&&yMap[prevY].n?yMap[prevY].s/yMap[prevY].n:0;
    const yoyPct=pyAvg?((cyAvg-pyAvg)/Math.abs(pyAvg)*100):0;

    // 트렌드 (최근 3개월 vs 이전 3개월)
    const r3=arr.slice(-3).filter(v=>v!=null),p3=arr.slice(-6,-3).filter(v=>v!=null);
    const rA=r3.length?r3.reduce((s,v)=>s+v,0)/r3.length:0;
    const pA=p3.length?p3.reduce((s,v)=>s+v,0)/p3.length:0;
    const tDir=!pA?'보합':rA>pA?'상승▲':rA<pA?'하락▼':'보합';
    const tPct=pA?Math.abs((rA-pA)/Math.abs(pA)*100).toFixed(1):0;
    const tC=tDir.includes('상승')?'#10b981':tDir.includes('하락')?'#ef4444':'#8b93b8';
    const f=(v,u)=>fmt(v??0,u==='억원'?1:0);
    const n12=Math.min(arr.length,12);

    const tRows=arr.slice(-n12).map((v,ti)=>{
      const ri=arr.length-n12+ti;
      const q=getQ(ri),y=getYoY(ri);
      const isMax=ri===maxI,isMin=ri===minI;
      const qS=q!=null?`<span style="color:${q>=0?'#10b981':'#ef4444'};font-weight:700">${q>=0?'▲':'▼'}${Math.abs(q).toFixed(1)}%</span>`:'<span style="color:#d1d5db">-</span>';
      const yS=y!=null?`<span style="color:${y>=0?'#10b981':'#ef4444'};font-weight:700">${y>=0?'▲':'▼'}${Math.abs(y).toFixed(1)}%</span>`:'<span style="color:#d1d5db">-</span>';
      return `<tr style="background:${isMax?'rgba(16,185,129,.06)':isMin?'rgba(239,68,68,.06)':''}">
        <td style="padding:8px 12px;font-size:11px;color:#4a5380;border-bottom:1px solid #f0f2f7;white-space:nowrap">${isMax?'🔺':isMin?'🔻':''} ${months[ri]}</td>
        <td style="padding:8px 12px;font-size:12px;font-weight:700;font-family:monospace;text-align:right;border-bottom:1px solid #f0f2f7;color:${(v??0)>=0?'#1a1f36':'#ef4444'}">${f(v,unit)}</td>
        <td style="padding:8px 12px;text-align:center;border-bottom:1px solid #f0f2f7;font-size:11px">${qS}</td>
        <td style="padding:8px 12px;text-align:center;border-bottom:1px solid #f0f2f7;font-size:11px">${yS}</td>
      </tr>`;
    }).join('');

    const yCards=Object.entries(yMap).map(([y,o])=>`
      <div style="background:#f8f9fc;border:1px solid #e4e7f0;border-radius:10px;padding:12px;text-align:center">
        <div style="font-size:10px;color:#8b93b8;font-weight:700;margin-bottom:4px">${y}년</div>
        <div style="font-size:16px;font-weight:800;font-family:monospace;color:${color}">${f(o.s,unit)}</div>
        <div style="font-size:10px;color:#8b93b8;margin-top:2px">${unit}·${o.n}개월</div>
      </div>`).join('');

    // 연간 카드 (평균 표시)
    const yCards=Object.entries(yMap).map(([y,o])=>`
      <div style="background:#f8f9fc;border:1px solid #e4e7f0;border-radius:10px;padding:12px;text-align:center">
        <div style="font-size:10px;color:#8b93b8;font-weight:700;margin-bottom:4px">${y}년</div>
        <div style="font-size:13px;font-weight:800;font-family:monospace;color:${color}">${f(o.s/o.n,unit)}</div>
        <div style="font-size:9px;color:#8b93b8;margin-top:1px">월평균 · ${o.n}개월</div>
        <div style="font-size:11px;font-weight:700;color:#4a5380;margin-top:3px">${f(o.s,unit)} 합계</div>
      </div>`).join('');

    // QoQ/YoY 헤더용 문구
    const qoqStr=pqAvg?`<span style="color:${qoqPct>=0?'#10b981':'#ef4444'};font-weight:700">${qoqPct>=0?'▲':'▼'}${Math.abs(qoqPct).toFixed(1)}%</span>`:'<span style="color:#d1d5db">-</span>';
    const yoyStr=pyAvg?`<span style="color:${yoyPct>=0?'#10b981':'#ef4444'};font-weight:700">${yoyPct>=0?'▲':'▼'}${Math.abs(yoyPct).toFixed(1)}%</span>`:'<span style="color:#d1d5db">-</span>';

    const html=`<div style="font-size:11px;color:#8b93b8;margin-bottom:14px">단위: ${unit} · ${arr.length}개월</div>

      <!-- 트렌드 요약 -->
      <div style="background:linear-gradient(135deg,rgba(91,110,245,.06),rgba(6,182,212,.04));border:1px solid rgba(91,110,245,.15);border-radius:12px;padding:14px 16px;margin-bottom:14px">
        <div style="font-size:10px;font-weight:700;color:#5b6ef5;margin-bottom:8px;letter-spacing:.5px;text-transform:uppercase">📊 트렌드 요약</div>
        <div style="font-size:13px;color:#4a5380;line-height:1.9">
          최근 3개월 평균 <b style="color:${color}">${f(rA,unit)}${unit}</b> · 직전 3개월 대비 <b style="color:${tC}">${tDir} ${tPct}%</b><br>
          기간 최고 <b style="color:#10b981">${f(maxV,unit)}${unit}</b>(${months[maxI]}) · 최저 <b style="color:#ef4444">${f(minV,unit)}${unit}</b>(${months[minI]})
        </div>
      </div>

      <!-- QoQ / YoY 핵심 비교 -->
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px">
        <div style="background:#fff;border:1.5px solid #e4e7f0;border-radius:12px;padding:14px">
          <div style="font-size:10px;color:#8b93b8;font-weight:700;margin-bottom:6px">QoQ · 분기평균 대비</div>
          <div style="font-size:22px;margin-bottom:4px">${qoqStr}</div>
          <div style="font-size:10px;color:#8b93b8">이번분기 평균 ${f(cqAvg,unit)} vs 전분기 ${f(pqAvg,unit)}</div>
        </div>
        <div style="background:#fff;border:1.5px solid #e4e7f0;border-radius:12px;padding:14px">
          <div style="font-size:10px;color:#8b93b8;font-weight:700;margin-bottom:6px">YoY · 연평균 대비</div>
          <div style="font-size:22px;margin-bottom:4px">${yoyStr}</div>
          <div style="font-size:10px;color:#8b93b8">${curY}년 평균 ${f(cyAvg,unit)} vs ${prevY||'전년'}년 ${f(pyAvg,unit)}</div>
        </div>
      </div>

      <!-- 최고/최저/평균 -->
      <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:14px">
        ${[['🔺 최고',f(maxV,unit)+unit,months[maxI],'#10b981'],['🔻 최저',f(minV,unit)+unit,months[minI],'#ef4444'],['— 평균',f(avgV,unit)+unit,'기간 평균','#6b7280']].map(([l,v,s,c])=>`<div style="background:#f8f9fc;border:1px solid #e4e7f0;border-radius:10px;padding:10px;text-align:center"><div style="font-size:10px;color:#8b93b8;font-weight:700;margin-bottom:4px">${l}</div><div style="font-size:14px;font-weight:800;font-family:monospace;color:${c}">${v}</div><div style="font-size:10px;color:#8b93b8;margin-top:2px">${s}</div></div>`).join('')}
      </div>

      <!-- 연도별 카드 -->
      <div style="font-size:12px;font-weight:700;color:#1a1f36;margin-bottom:8px">📅 연도별 현황</div>
      <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(100px,1fr));gap:8px;margin-bottom:14px">${yCards}</div>

      <!-- 월별 테이블 -->
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">
        <div style="font-size:12px;font-weight:700;color:#1a1f36">📋 최근 ${n12}개월 월별 추이</div>
        <div style="font-size:10px;color:#8b93b8">🔺최고 · 🔻최저</div>
      </div>
      <div style="border:1px solid #e4e7f0;border-radius:10px;overflow:hidden;max-height:300px;overflow-y:auto">
        <table style="width:100%;border-collapse:collapse">
          <thead style="background:#f8f9fc;position:sticky;top:0;z-index:1"><tr>
            <th style="padding:8px 12px;font-size:10px;color:#8b93b8;text-align:left;font-weight:700">월</th>
            <th style="padding:8px 12px;font-size:10px;color:#8b93b8;text-align:right;font-weight:700">실적(${unit})</th>
            <th style="padding:8px 12px;font-size:10px;color:#8b93b8;text-align:center;font-weight:700">전월비</th>
            <th style="padding:8px 12px;font-size:10px;color:#8b93b8;text-align:center;font-weight:700">전년동월비</th>
          </tr></thead>
          <tbody>${tRows}</tbody>
        </table>
      </div>`;

    // 패널을 document.body에 직접 붙임 (modal 레이어 없이)
    let panel=document.getElementById('kpi-side-panel');
    let overlay=document.getElementById('kpi-overlay');

    if(!panel){
      // CSS
      const st=document.createElement('style');
      st.textContent='#kpi-overlay{display:none;position:fixed;inset:0;z-index:8000;background:rgba(0,0,0,.3);backdrop-filter:blur(2px)}'
        +'#kpi-overlay.on{display:block}'
        +'#kpi-side-panel{position:fixed;top:0;right:-620px;width:min(580px,100vw);height:100vh;height:100dvh;background:#fff;z-index:8001;display:flex;flex-direction:column;box-shadow:-4px 0 30px rgba(0,0,0,.15);transition:right .3s ease}'
        +'#kpi-side-panel.on{right:0}';
      document.head.appendChild(st);

      overlay=document.createElement('div');
      overlay.id='kpi-overlay';
      overlay.onclick=closeKpiPanel;
      document.body.appendChild(overlay);

      panel=document.createElement('div');
      panel.id='kpi-side-panel';
      document.body.appendChild(panel);
    }

    // 패널 내용 구성
    panel.innerHTML='';

    // 헤더
    const hdr=document.createElement('div');
    hdr.style.cssText='display:flex;align-items:center;justify-content:space-between;padding:20px 24px;border-bottom:1px solid #e4e7f0;flex-shrink:0;background:#fff';
    const titleDiv=document.createElement('div');
    titleDiv.style.cssText='display:flex;align-items:center;gap:8px;font-size:17px;font-weight:800;color:#1a1f36';
    titleDiv.innerHTML=icon+' '+label+' 상세분석';
    const closeBtn=document.createElement('button');
    closeBtn.textContent='✕';
    closeBtn.style.cssText='background:#f0f2f7;border:none;color:#6b7280;font-size:18px;cursor:pointer;border-radius:10px;width:36px;height:36px;display:flex;align-items:center;justify-content:center';
    closeBtn.onclick=closeKpiPanel;
    hdr.appendChild(titleDiv);
    hdr.appendChild(closeBtn);
    panel.appendChild(hdr);

    // 스크롤 영역
    const scroll=document.createElement('div');
    scroll.style.cssText='flex:1;overflow-y:auto;padding:20px 24px 40px';
    scroll.innerHTML=html;
    panel.appendChild(scroll);

    // 열기
    overlay.classList.add('on');
    panel.classList.add('on');

  }catch(e){
    console.error('KPI popup error:',e,'key:',key);
    showToast('오류: '+e.message);
  }
}

function closeKpiPanel(){
  const p=document.getElementById('kpi-side-panel');
  const o=document.getElementById('kpi-overlay');
  if(p) p.classList.remove('on');
  if(o) o.classList.remove('on');
}

window.renderHR=function(){
  const d=gd('hr');

  // 퇴직률 계산 (퇴직/전사계)
  const lastTotal = last(d.전사계)||1;
  const monthlyQuit = last(d.퇴직계)||0;
  const quitRate = (monthlyQuit/lastTotal*100);
  const annualQuitEst = (sum(d.퇴직계)/d.전사계.reduce((s,v,i)=>s+(v||0),0)*12*100)||0;

  document.getElementById('hr-kpi').innerHTML=[
    kpi('전사계 (최근월)',fmt(last(d.전사계)),'명',pct(last(d.전사계),d.전사계[0]),'gold','기간 시작 대비','hr:전사계'),
    kpi('영업직',fmt(last(d.영업직)),'명',pct(last(d.영업직),d.영업직[0]),'blue','','hr:영업직'),
    kpi('SC직',fmt(last(d.SC직)),'명',pct(last(d.SC직),d.SC직[0]),'teal','','hr:SC직'),
    kpi('일반직',fmt(last(d.일반직)),'명',null,'purple',''),
    kpi('소매채널',fmt(last(d.소매채널)),'명',pct(last(d.소매채널),d.소매채널[0]),'green',''),
    kpi('퇴직률 (연환산)',fmt(annualQuitEst,1),'%',null,'red',`월 퇴직 평균 ${fmt(sum(d.퇴직계)/d.months.length,1)}명`),
  ].join('');

  // ① 인력 구성 추이 (전사계만 — 변화 없는 세부는 제거)
  line('ch-hr-main',d.months,[
    {label:'전사계',data:d.전사계,fill:true},
    {label:'영업직',data:d.영업직,color:C.blue},
    {label:'SC직',data:d.SC직,color:C.teal},
    {label:'일반직',data:d.일반직,color:C.purple}
  ]);

  // ② 퇴직 추이 — 막대 (유의미한 변화 보여줌)
  bar('ch-hr-ch',d.months,[{
    label:'월별 퇴직',
    data:d.퇴직계,
    color:C.red,
    bg:'rgba(239,68,68,0.65)'
  }]);

  // ③ 연마감 기준 인력 구성 도넛
  // 각 연도 12월(또는 마지막월) 기준
  const yearMap = {};
  d.months.forEach((m,i)=>{
    const y = m.match(/['\s]?(\d{2})\./)?.[1];
    if(y) yearMap[y] = i; // 마지막 인덱스로 덮어씌움 → 연마감
  });

  // 가장 최근 연도의 연마감 기준으로 도넛
  const years = Object.keys(yearMap).sort();
  const latestYear = years[years.length-1];
  const latestIdx = yearMap[latestYear];

  const jobTypes = [
    {label:'영업직', key:'영업직', color:'rgba(91,110,245,0.72)'},
    {label:'SC직',   key:'SC직',   color:'rgba(6,182,212,0.72)'},
    {label:'일반직', key:'일반직', color:'rgba(139,92,246,0.72)'},
  ];
  const vals = jobTypes.map(j=>d[j.key]?.[latestIdx]||0);

  // 연마감 도넛
  const hrTotal = vals.reduce((s,v)=>s+v,0);
  makeDonut('ch-hr-donut', jobTypes.map(j=>j.label), vals, jobTypes.map(j=>j.color), fmt(hrTotal)+'명', latestYear+'년');

  // ④ 연도별 비교 막대
  const hqs=['강북','강남','강서','동부','서부'];
  bar('ch-hr-hq',d.months,hqs.map((h,i)=>({label:h+' 본부',data:d[h+'본부']||[],color:COLORS[i],bg:COLORS_A[i]})),true);
};

// ── ADMIN DATA ──
const UPLOAD_H=[
  {id:1,file:'경영성과_재무.xlsx',cat:'재무',period:'23.1~25.12',rows:8,cols:36,size:'42KB',date:'2025-12-31 09:12',user:'최관리',ok:true,note:'23~25년 전체 재무지표',preview:[['월','총매출','통신매출','영업이익'],['25.12월','507.17','490.82','-21.86'],['25.11월','569.38','515.05','-8.25'],['25.10월','572.42','526.3','-18.31']]},
  {id:2,file:'경영성과_무선_가입자.xlsx',cat:'무선 가입자',period:'23.1~25.12',rows:7,cols:36,size:'38KB',date:'2025-12-31 09:14',user:'최관리',ok:true,note:'일반후불 채널별 무선 가입자',preview:[['월','유지','CAPA','해지','순증'],['25.12월','1,647,799','52,735','19,584','-24,490'],['25.11월','1,672,289','52,601','17,373','-21,327']]},
  {id:3,file:'경영성과_유선_가입자.xlsx',cat:'유선 가입자',period:'23.1~25.12',rows:7,cols:36,size:'36KB',date:'2025-12-31 09:15',user:'최관리',ok:true,note:'인터넷+TV 유선 가입자',preview:[['월','유지_전체','신규','해지','순증'],['25.12월','248,042','7,134','3,051','-6,438'],['25.11월','254,480','7,481','2,850','-6,165']]},
  {id:4,file:'경영성과_유무선_가입자_조직별_.xlsx',cat:'통합본(14시트)',period:'23.1~25.12',rows:120,cols:36,size:'290KB',date:'2025-12-31 09:20',user:'최관리',ok:true,note:'재무·무선·유선·채널·기타 전체 통합',preview:[['시트명','행수','비고'],['재무','8행','총매출~마케팅비'],['무선(채널별)','7행','유지/CAPA/해지/순증'],['유선(채널별)','7행','인터넷+TV'],['유무선(조직별)','5행','소매/도매/디지털/B2B/소상공인']]},
  {id:5,file:'경영성과_디지털_채널_.xlsx',cat:'디지털',period:'23.1~25.12',rows:8,cols:36,size:'34KB',date:'2025-12-31 09:22',user:'정데이터',ok:true,note:'KT닷컴 채널 실적',preview:[['월','일반후불_총계','운영후불','유선순신규','인력'],['25.12월','12,178','8,793','242','136'],['25.11월','11,138','8,605','270','135']]},
  {id:6,file:'경영성과_B2B_채널_.xlsx',cat:'B2B',period:'23.1~25.12',rows:16,cols:36,size:'45KB',date:'2025-12-31 09:23',user:'정데이터',ok:true,note:'기업/법인 B2B 실적',preview:[['월','전체_일반후불','전체_무선순증','가입자_무선유지'],['25.12월','3,470','69','207,314'],['25.11월','5,004','3,977','208,510']]},
  {id:7,file:'경영성과_소상공인_채널_.xlsx',cat:'소상공인',period:'23.1~25.12',rows:9,cols:36,size:'32KB',date:'2025-12-31 09:24',user:'정데이터',ok:true,note:'소상공인 채널 상품 실적',preview:[['월','일반후불','운영후불','인터넷순신규','인력'],['25.12월','103','74','381','68'],['25.11월','194','131','487','102']]},
  {id:8,file:'경영성과_유통플랫폼_채널_.xlsx',cat:'유통플랫폼',period:'24.1~25.12',rows:6,cols:24,size:'28KB',date:'2025-12-31 09:25',user:'정데이터',ok:true,note:'중고폰/시연폰 유통 실적',preview:[['월','중고폰_매입건수','중고폰_매각건수','시연폰_매각이익'],['25.12월','1,959','2,087','1.2억'],['25.11월','2,397','4,171','2.8억']]},
  {id:9,file:'경영성과_TCSI_기타_.xlsx',cat:'TCSI',period:'23.1~25.12',rows:13,cols:36,size:'40KB',date:'2025-12-31 09:26',user:'정데이터',ok:true,note:'TCSI·KT점수·지역별 소매점수',preview:[['월','TCSI점수','KT점수','대리점점수'],['25.12월','98.7','97.7','97.4'],['25.11월','98.6','97.7','97.3']]},
  {id:10,file:'경영성과_영업품질_기타_.xlsx',cat:'영업품질',period:'23.1~25.12',rows:55,cols:36,size:'48KB',date:'2025-12-31 09:27',user:'최관리',ok:true,note:'R-VOC 발생률·대외민원',preview:[['월','도+소매발생률','소매발생률','도매발생률','대외민원'],['25.12월','0.4832%','0.5054%','0.4635%','13건'],['25.11월','0.3911%','0.4638%','0.3251%','19건']]},
  {id:11,file:'경영성과_인력_기타_.xlsx',cat:'인력',period:'23.1~25.12',rows:115,cols:36,size:'62KB',date:'2025-12-31 09:28',user:'최관리',ok:true,note:'직급별·조직별·채널별 인력',preview:[['월','전사계','임원','일반직','영업직','SC직'],['25.12월','2,200','8','184','1,830','145'],['25.11월','2,193','8','189','1,815','146']]},
  {id:12,file:'경영성과_인프라_기타_.xlsx',cat:'인프라',period:'23.1~25.12',rows:340,cols:36,size:'95KB',date:'2025-12-31 09:29',user:'최관리',ok:true,note:'소매매장수·생산성·도매취급점',preview:[['월','소매매장수','출점','퇴점','도매무선취급점'],['25.12월','250','1','3','4,212'],['25.11월','252','1','1','4,174']]},
  {id:13,file:'경영성과_전략상품_기타_.xlsx',cat:'전략상품',period:'23.1~25.12',rows:29,cols:36,size:'38KB',date:'2025-12-31 09:30',user:'최관리',ok:true,note:'하이오더·GiGAeyes·AI전화 등',preview:[['월','하이오더_점포수','GiGAeyes_계','AI전화'],['25.12월','57','1,401','6'],['25.11월','64','1,839','12']]},
];

// METRICS with company priority data
let METRICS_L=[
  {id:1, name:'총매출',         cat:'재무',   unit:'억원', pos:'재무 KPI', desc:'수수료매출+상품매출 합산. 회사 최우선 핵심 지표',      src:'finance.총매출',          on:true},
  {id:2, name:'수수료매출(통신)',   cat:'재무',   unit:'억원', pos:'재무 KPI', desc:'통신 서비스 수수료 매출 (무선+유선)',                 src:'finance.통신매출',         on:true},
  {id:3, name:'수수료매출(유통플랫폼)',cat:'재무', unit:'억원', pos:'재무 KPI', desc:'유통플랫폼(중고폰/시연폰) 수수료 매출',              src:'platform.중고폰_매입금액',  on:true},
  {id:4, name:'영업이익',         cat:'재무',   unit:'억원', pos:'재무 KPI', desc:'총매출 - 총비용. 최우선 수익성 지표',                src:'finance.영업이익',         on:true},
  {id:5, name:'소매 채널 영업이익', cat:'채널',   unit:'억원', pos:'채널 KPI', desc:'소매 채널 기여 영업이익 (추정)',                    src:'org.소매',                on:true},
  {id:6, name:'도매 채널 영업이익', cat:'채널',   unit:'억원', pos:'채널 KPI', desc:'도매 채널 기여 영업이익 (추정)',                    src:'org.도매',                on:true},
  {id:7, name:'디지털 채널 영업이익',cat:'채널',  unit:'억원', pos:'채널 KPI', desc:'디지털(KT닷컴/O2O) 채널 기여 이익',               src:'org.디지털',              on:true},
  {id:8, name:'B2B 채널 영업이익', cat:'채널',   unit:'억원', pos:'채널 KPI', desc:'B2B(기업/법인) 채널 기여 이익',                   src:'org.B2B',                 on:true},
  {id:9, name:'유선 유지 가입자',   cat:'유선',   unit:'명',   pos:'유선 KPI', desc:'인터넷+TV 유지 가입자 전체',                        src:'wired.유지_전체',          on:true},
  {id:10,name:'TCSI 점수',        cat:'품질',   unit:'점',   pos:'품질 KPI', desc:'고객 서비스 만족도 (높을수록 우수)',                 src:'tcsi.TCSI점수',            on:true},
  {id:11,name:'VOC 도+소매 발생률', cat:'품질',   unit:'%',    pos:'품질 KPI', desc:'R-VOC 발생률 (낮을수록 우수)',                     src:'voc.도소매발생률',          on:true},
  {id:12,name:'소매 매장수',        cat:'인프라', unit:'개',   pos:'인프라 KPI',desc:'소매 매장 Infra 합계',                            src:'infra.소매매장수_계',       on:true},
  {id:13,name:'GiGAeyes',         cat:'전략상품',unit:'대',   pos:'전략 KPI', desc:'전략상품 GiGAeyes 신규 공급 누계',                 src:'strategy.GiGAeyes_계',    on:true},
  {id:14,name:'전사 인원',          cat:'인력',   unit:'명',   pos:'인력 KPI', desc:'KT M&S 전사 임직원 합계',                         src:'hr.전사계',                on:true},
];
let _editMetricIdx = -1;

// USERS with full access control
let USERS_FULL=[
  {idx:0,name:'김철수',id:'ceo',  pw:'kt2025',  email:'cs.kim@kt.com',  title:'대표이사',      role:'executive',pages:['all'],admin:[],  last:'2025-12-31',on:true},
  {idx:1,name:'이영희',id:'cfo',  pw:'kt2025',  email:'yh.lee@kt.com',  title:'재무본부장',    role:'executive',pages:['all'],admin:[],  last:'2025-12-30',on:true},
  {idx:2,name:'박민준',id:'cso',  pw:'kt2025',  email:'mj.park@kt.com', title:'영업총괄',      role:'executive',pages:['finance','wireless','wired','org'],admin:[],last:'2025-12-29',on:true},
  {idx:3,name:'정지현',id:'cmo',  pw:'kt2025',  email:'jh.jung@kt.com', title:'마케팅본부장',  role:'executive',pages:['all'],admin:[],  last:'2025-12-28',on:true},
  {idx:4,name:'최현우',id:'sales1',pw:'kt2025', email:'hw.choi@kt.com', title:'소매영업팀장',  role:'executive',pages:['wireless','wired','org','digital'],admin:[],last:'2025-12-27',on:true},
  {idx:5,name:'최관리',id:'admin',pw:'admin1234',email:'admin@kt.com',  title:'시스템관리자',  role:'admin',    pages:['all'],admin:['all'],last:'2025-12-31',on:true},
  {idx:6,name:'정데이터',id:'data',pw:'data1234',email:'data@kt.com',   title:'데이터팀',      role:'admin',    pages:['all'],admin:['upload','export'],last:'2025-12-28',on:true},
];
let _editUserIdx = -1;
const ALL_PAGES=['finance','wireless','wired','org','digital','b2b','smb','platform','quality','infra','strategy','hr'];
const PAGE_LABELS={finance:'재무',wireless:'무선',wired:'유선',org:'조직별',digital:'디지털',b2b:'B2B',smb:'소상공인',platform:'유통플랫폼',quality:'TCSI/VOC',infra:'인프라',strategy:'전략상품',hr:'인력'};
const ADMIN_PAGES=['overview','upload','metrics','users','export','beta'];
const ADMIN_LABELS={overview:'시스템현황',upload:'데이터업로드',metrics:'지표관리',users:'사용자관리',export:'보고서',beta:'베타배포'};
let _showPwIdx = new Set();
let _findTab = 'id';
let _deployHistory=[
  {dt:'2025-12-31 09:00',user:'최관리',cnt:3,ok:true},
  {dt:'2025-12-15 14:22',user:'최관리',cnt:7,ok:true},
  {dt:'2025-12-01 10:05',user:'최관리',cnt:2,ok:true},
];
const _pendingChanges=[
  {type:'add',  icon:'🟢', msg:'지표 추가: 수수료매출(유통플랫폼)'},
  {type:'mod',  icon:'🟡', msg:'지표 수정: 총매출 → 표시위치 변경'},
  {type:'del',  icon:'🔴', msg:'지표 삭제: 무선 순증 (비활성화)'},
  {type:'mod',  icon:'🟡', msg:'사용자 권한 변경: 최현우 열람범위 확대'},
  {type:'add',  icon:'🟢', msg:'파일 업로드: 경영성과_전략상품_기타_.xlsx'},
];
let _deployed = false;

// ── Admin Render ──
function renderAdminOverview(){
  document.getElementById('admin-kpi').innerHTML=[
    kpi('등록 지표',METRICS_L.filter(m=>m.on).length+'','개',null,'gold','전체 '+METRICS_L.length+'개'),
    kpi('등록 사용자',USERS_FULL.filter(u=>u.on).length+'','명',null,'teal','임원'+USERS_FULL.filter(u=>u.role==='executive'&&u.on).length+'/관리자'+USERS_FULL.filter(u=>u.role==='admin'&&u.on).length),
    kpi('업로드 파일',UPLOAD_H.length+'','개',null,'blue','최근 2025-12-31'),
    kpi('데이터 기간','36','개월',null,'green','23.1~25.12'),
  ].join('');
  document.getElementById('upload-log').innerHTML=UPLOAD_H.slice(0,5).map(u=>`<tr><td style="font-size:11px">${u.file}</td><td><span class="td-tag" style="background:rgba(74,158,255,.1);color:var(--blue)">${u.cat}</span></td><td style="color:var(--text3)">${u.date.split(' ')[0]}</td><td><span class="status st-on">정상</span></td></tr>`).join('');
  const cats=[...new Set(METRICS_L.map(m=>m.cat))];
  document.getElementById('metric-stats').innerHTML=cats.map(c=>{const cnt=METRICS_L.filter(m=>m.cat===c).length;const colors=['var(--gold)','var(--teal)','var(--blue)','var(--green)','var(--purple)','var(--orange)','var(--pink)','var(--red)'];const ci=['재무','채널','무선','유선','품질','인프라','전략상품','인력'].indexOf(c);const col=colors[ci]||colors[0];return`<div style="margin-bottom:10px"><div style="display:flex;justify-content:space-between;font-size:11px;margin-bottom:4px"><span style="color:var(--text2)">${c}</span><span style="color:var(--text3);font-family:var(--mono)">${cnt}개</span></div><div class="prog-bar"><div class="prog-fill" style="width:${Math.min(cnt*14,100)}%;background:${col}"></div></div></div>`;}).join('');
}
// ── RAW 데이터 소스 맵 (데이터 연결 피커용) ──
const SRC_META = {
  finance:  {label:'💰 재무',      icon:'💰'},
  wireless: {label:'📱 무선 가입자',icon:'📱'},
  wired:    {label:'🌐 유선 가입자',icon:'🌐'},
  org:      {label:'🏢 조직별 실적',icon:'🏢'},
  digital:  {label:'💻 디지털 채널',icon:'💻'},
  b2b:      {label:'🏭 B2B 채널',   icon:'🏭'},
  smb:      {label:'🏪 소상공인',   icon:'🏪'},
  platform: {label:'♻️ 유통플랫폼', icon:'♻️'},
  tcsi:     {label:'⭐ TCSI',       icon:'⭐'},
  voc:      {label:'📊 VOC 품질',   icon:'📊'},
  hr:       {label:'👥 인력',       icon:'👥'},
  infra:    {label:'🗺️ 인프라',     icon:'🗺️'},
  strategy: {label:'🚀 전략상품',   icon:'🚀'},
};

let _udCurrentIdx = -1; // 현재 열린 업로드 상세 인덱스
let _meSelectedSrc = ''; // 현재 선택된 데이터 소스키 (예: 'finance')
let _meSelectedField = ''; // 현재 선택된 필드 (예: '총매출')

// ── UPLOAD DETAIL (재다운로드 포함) ──
function openUploadDetail(i){
  _udCurrentIdx = i;
  const u = UPLOAD_H[i];
  document.getElementById('ud-title').textContent = u.file;
  document.getElementById('ud-sub').textContent = `${u.cat} · ${u.period} · ${u.size} · ${u.date}`;
  document.getElementById('ud-redownload-btn').textContent = `📥 "${u.file.replace('.xlsx','.csv')}" 다운로드`;

  const prevHTML = u.preview ? `
    <div class="card-wrap" style="margin-bottom:12px">
      <div class="chart-title" style="margin-bottom:10px">📊 데이터 미리보기 (최근 행)</div>
      <table class="data-table">
        <thead><tr>${u.preview[0].map(h=>`<th>${h}</th>`).join('')}</tr></thead>
        <tbody>${u.preview.slice(1).map(r=>`<tr>${r.map(c=>`<td style="font-family:var(--mono);font-size:11px">${c}</td>`).join('')}</tr>`).join('')}</tbody>
      </table>
    </div>` : '';

  document.getElementById('ud-body').innerHTML = `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px">
      ${[['파일명',u.file],['카테고리',u.cat],['데이터 기간',u.period],['파일 크기',u.size],
         ['업로드 일시',u.date],['담당자',u.user],['데이터 행수',u.rows+'행'],['데이터 열수',u.cols+'열'],
         ['상태',u.ok?'✅ 정상':'❌ 오류'],['비고',u.note]]
        .map(([k,v])=>`
        <div style="background:var(--bg3);border:1px solid var(--border);border-radius:7px;padding:10px 12px">
          <div style="font-size:10px;color:var(--text3);font-weight:600;margin-bottom:3px">${k}</div>
          <div style="font-size:12px;color:var(--text)">${v}</div>
        </div>`).join('')}
    </div>${prevHTML}`;

  openModal('modal-upload-detail');
}

function redownloadFile(){
  const i = _udCurrentIdx;
  if(i < 0) return;
  const u = UPLOAD_H[i];

  // 카테고리 → RAW 키 매핑
  const catMap = {
    '재무':'finance', '무선 가입자':'wireless', '유선 가입자':'wired',
    '통합본(14시트)':'finance', '디지털':'digital', 'B2B':'b2b',
    '소상공인':'smb', '유통플랫폼':'platform', 'TCSI':'tcsi',
    '영업품질':'voc', '인력':'hr', '인프라':'infra', '전략상품':'strategy',
  };
  const rawKey = catMap[u.cat];
  const d = rawKey ? RAW[rawKey] : null;

  let csv = '\uFEFF'; // BOM
  csv += `# ${u.file} — KT M&S 경영성과 데이터\n`;
  csv += `# 카테고리: ${u.cat} | 기간: ${u.period} | 생성: ${new Date().toLocaleString('ko-KR')}\n\n`;

  if(d && d.months){
    // 숫자 배열 필드만 추출
    const fields = Object.keys(d).filter(k => k !== 'months' && Array.isArray(d[k]));
    csv += ['월', ...fields].join(',') + '\n';
    d.months.forEach((m, idx) => {
      const row = [m, ...fields.map(f => d[f][idx] ?? '')];
      csv += row.join(',') + '\n';
    });
  } else if(u.preview){
    // fallback: 미리보기 데이터
    u.preview.forEach(row => { csv += row.join(',') + '\n'; });
  } else {
    csv += '데이터를 불러올 수 없습니다.\n';
  }

  const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = u.file.replace('.xlsx', '.csv');
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast(`📥 "${a.download}" 다운로드 완료`);
}

// ── METRIC DATA PICKER ──
function buildSrcPicker(currentSrc){
  // 좌측 소스 버튼 목록
  const list = document.getElementById('me-src-list');
  list.innerHTML = Object.entries(SRC_META).map(([key,meta])=>`
    <button class="src-btn ${currentSrc===key?'active':''}" onclick="selectMetricSrc('${key}')">
      <span>${meta.icon}</span><span>${meta.label.replace(/^[^ ]+ /,'')}</span>
    </button>`).join('');

  // 만약 currentSrc가 있으면 필드 목록도 초기 렌더
  if(currentSrc && RAW[currentSrc]){
    renderFieldList(currentSrc, _meSelectedField);
    document.getElementById('me-field-label').textContent = `${SRC_META[currentSrc].label} 필드 목록`;
  } else {
    document.getElementById('me-field-list').innerHTML = `
      <div style="grid-column:1/-1;font-size:11px;color:var(--text3);padding:20px;text-align:center">
        좌측에서 데이터 소스를 선택하세요
      </div>`;
  }
}

function selectMetricSrc(key){
  _meSelectedSrc = key;
  // 좌측 버튼 active 갱신
  document.querySelectorAll('#me-src-list .src-btn').forEach(b=>b.classList.remove('active'));
  event.currentTarget.classList.add('active');
  document.getElementById('me-field-label').textContent = `${SRC_META[key].label} 필드 목록`;
  renderFieldList(key, '');
}

function renderFieldList(key, selectedField){
  const d = RAW[key];
  if(!d){ document.getElementById('me-field-list').innerHTML='<div style="color:var(--text3);font-size:11px;padding:12px">데이터 없음</div>'; return; }

  const fields = Object.keys(d).filter(k => k !== 'months' && Array.isArray(d[k]));
  // 중첩 객체 (예: tcsi.regions)도 처리
  const objFields = Object.keys(d).filter(k => k !== 'months' && !Array.isArray(d[k]) && typeof d[k]==='object');

  let html = '';
  fields.forEach(f => {
    const path = `${key}.${f}`;
    const isSelected = path === document.getElementById('me-src').value;
    const lastVal = d[f] ? d[f][d[f].length-1] : '–';
    const dispVal = typeof lastVal === 'number' ? Number(lastVal).toLocaleString('ko-KR',{maximumFractionDigits:2}) : lastVal;
    html += `<button class="field-btn ${isSelected?'selected':''}" title="${path}&#10;최근값: ${dispVal}" onclick="selectMetricField('${key}','${f}')">
      <div style="font-size:10px;font-weight:600;color:inherit;margin-bottom:1px">${f}</div>
      <div style="font-size:9px;color:var(--text3)">${dispVal}</div>
    </button>`;
  });

  // nested object keys
  objFields.forEach(obj => {
    const subKeys = Object.keys(d[obj]);
    subKeys.forEach(sf => {
      const path = `${key}.${obj}.${sf}`;
      const arr = d[obj][sf];
      const lastVal = Array.isArray(arr) ? arr[arr.length-1] : '–';
      const dispVal = typeof lastVal==='number' ? Number(lastVal).toLocaleString('ko-KR',{maximumFractionDigits:2}) : lastVal;
      html += `<button class="field-btn" title="${path}&#10;최근값: ${dispVal}" onclick="selectMetricField('${key}','${obj}.${sf}')">
        <div style="font-size:10px;font-weight:600;color:inherit;margin-bottom:1px">${obj} › ${sf}</div>
        <div style="font-size:9px;color:var(--text3)">${dispVal}</div>
      </button>`;
    });
  });

  document.getElementById('me-field-list').innerHTML = html;
}

function selectMetricField(srcKey, fieldPath){
  const fullPath = `${srcKey}.${fieldPath}`;
  _meSelectedField = fieldPath;
  document.getElementById('me-src').value = fullPath;

  // UI 갱신
  document.getElementById('me-src-display').textContent = fullPath;
  document.getElementById('me-src-display').style.color = 'var(--teal)';

  // 미리보기 (최근 3개월)
  const d = RAW[srcKey];
  let arr;
  if(fieldPath.includes('.')){
    const parts = fieldPath.split('.');
    arr = d[parts[0]]?.[parts[1]];
  } else {
    arr = d[fieldPath];
  }
  if(Array.isArray(arr)){
    const months = d.months;
    const last3 = arr.slice(-3).map((v,i)=>`${months[months.length-3+i]}: ${typeof v==='number'?Number(v).toLocaleString('ko-KR',{maximumFractionDigits:2}):v}`);
    document.getElementById('me-src-preview').textContent = last3.join(' | ');
  } else {
    document.getElementById('me-src-preview').textContent = '–';
  }

  // 해당 버튼 selected 표시
  document.querySelectorAll('#me-field-list .field-btn').forEach(b=>b.classList.remove('selected'));
  event.currentTarget.classList.add('selected');
}

function clearMetricSrc(){
  _meSelectedSrc=''; _meSelectedField='';
  document.getElementById('me-src').value='';
  document.getElementById('me-src-display').textContent='선택 없음';
  document.getElementById('me-src-display').style.color='var(--text3)';
  document.getElementById('me-src-preview').textContent='–';
  document.querySelectorAll('#me-field-list .field-btn').forEach(b=>b.classList.remove('selected'));
  document.querySelectorAll('#me-src-list .src-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('me-field-list').innerHTML=`<div style="grid-column:1/-1;font-size:11px;color:var(--text3);padding:20px;text-align:center">좌측에서 데이터 소스를 선택하세요</div>`;
}
function renderAdminUpload(){
  resetUpload();
  // Supabase에서 이력 로드
  sb.from('kts_upload_log')
    .select('*')
    .order('created_at', {ascending: false})
    .limit(20)
    .then(({data, error}) => {
      const tbody = document.getElementById('upload-history');
      if(!tbody) return;
      if(error || !data || data.length === 0){
        // 로컬 UPLOAD_H fallback
        tbody.innerHTML = UPLOAD_H.map(u=>`
          <tr>
            <td style="font-size:11px"><b style="color:var(--text)">${u.file}</b></td>
            <td><span class="td-tag" style="background:rgba(74,158,255,.1);color:var(--blue)">${u.cat}</span></td>
            <td style="color:var(--text3);font-size:11px">${u.period}</td>
            <td style="color:var(--text3);font-size:11px;text-align:center">–</td>
            <td style="color:var(--text3);font-family:var(--mono);font-size:11px">${u.date}</td>
            <td style="color:var(--text2);font-size:11px">${u.user}</td>
            <td><span class="status ${u.ok?'st-on':'st-off'}">${u.ok?'정상':'오류'}</span></td>
          </tr>`).join('');
        return;
      }
      // Supabase 이력 표시
      tbody.innerHTML = data.map(u=>`
        <tr>
          <td style="font-size:11px"><b style="color:var(--text)">${u.filename}</b></td>
          <td><span class="td-tag" style="background:rgba(74,158,255,.1);color:var(--blue)">${CAT_SHEET_MAP[u.category]?.label || u.category}</span></td>
          <td style="color:var(--text3);font-size:11px">${u.period || '–'}</td>
          <td style="color:var(--text3);font-size:11px;text-align:center">${u.rows_count ? u.rows_count+'행' : '–'}</td>
          <td style="color:var(--text3);font-family:var(--mono);font-size:11px">${new Date(u.created_at).toLocaleString('ko-KR')}</td>
          <td style="color:var(--text2);font-size:11px">${u.uploaded_by || '–'}</td>
          <td><span class="status ${u.status==='success'?'st-on':'st-off'}">${u.status==='success'?'정상':'오류'}</span></td>
        </tr>`).join('');
    });
}

function renderAdminMetrics(){
  document.getElementById('metrics-tbody').innerHTML=METRICS_L.map((m,i)=>`
    <tr>
      <td><b style="color:var(--text)">${m.name}</b></td>
      <td><span class="td-tag" style="background:rgba(74,158,255,.1);color:var(--blue)">${m.cat}</span></td>
      <td style="font-family:var(--mono);color:var(--text3)">${m.unit}</td>
      <td style="color:var(--text3);font-size:11px">${m.pos}</td>
      <td style="color:var(--text3);font-size:11px;max-width:160px;white-space:normal">${m.desc||''}</td>
      <td>
        <label style="cursor:pointer;display:flex;align-items:center;gap:5px;font-size:11px">
          <input type="checkbox" ${m.on?'checked':''} onchange="toggleMetric(${i},this.checked)" style="accent-color:var(--gold)">
          <span class="status ${m.on?'st-on':'st-off'}">${m.on?'활성':'비활성'}</span>
        </label>
      </td>
      <td><button class="btn btn-ghost" style="padding:3px 10px;font-size:10px" onclick="openMetricEdit(${i})">편집</button></td>
    </tr>`).join('');
}
function toggleMetric(i, val){
  METRICS_L[i].on=val;
  const spans=document.querySelectorAll('#metrics-tbody tr');
  if(spans[i]){
    const st=spans[i].querySelector('.status');
    if(st){st.className='status '+(val?'st-on':'st-off');st.textContent=val?'활성':'비활성';}
  }
  showToast(`${METRICS_L[i].name} ${val?'활성화':'비활성화'}`);
}
function openMetricEdit(i){
  _editMetricIdx = i;
  const m = METRICS_L[i];
  document.getElementById('me-title').textContent = '지표 편집: ' + m.name;
  document.getElementById('me-name').value = m.name;
  document.getElementById('me-cat').value = m.cat;
  document.getElementById('me-unit').value = m.unit;
  document.getElementById('me-pos').value = m.pos;
  document.getElementById('me-desc').value = m.desc || '';
  document.getElementById('me-active').value = String(m.on);
  document.getElementById('me-agg').value = m.agg || 'last';
  document.getElementById('me-src').value = m.src || '';

  // 데이터 피커 초기화
  const srcParts = (m.src || '').split('.');
  _meSelectedSrc = srcParts[0] || '';
  _meSelectedField = srcParts.slice(1).join('.') || '';

  // 연결 표시 갱신
  if(m.src){
    document.getElementById('me-src-display').textContent = m.src;
    document.getElementById('me-src-display').style.color = 'var(--teal)';
    // 미리보기
    const d = RAW[srcParts[0]];
    const arr = srcParts.length === 2 ? d?.[srcParts[1]] : d?.[srcParts[1]]?.[srcParts[2]];
    if(Array.isArray(arr) && d?.months){
      const last3 = arr.slice(-3).map((v,i)=>`${d.months[d.months.length-3+i]}: ${typeof v==='number'?Number(v).toLocaleString('ko-KR',{maximumFractionDigits:2}):v}`);
      document.getElementById('me-src-preview').textContent = last3.join(' | ');
    }
  } else {
    document.getElementById('me-src-display').textContent = '선택 없음';
    document.getElementById('me-src-display').style.color = 'var(--text3)';
    document.getElementById('me-src-preview').textContent = '–';
  }

  buildSrcPicker(_meSelectedSrc);
  openModal('modal-metric-edit');
}

function openMetricAdd(){
  _editMetricIdx = -1;
  _meSelectedSrc = ''; _meSelectedField = '';
  document.getElementById('me-title').textContent = '새 지표 추가';
  ['me-name','me-unit','me-desc'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('me-active').value = 'true';
  document.getElementById('me-agg').value = 'last';
  document.getElementById('me-src').value = '';
  document.getElementById('me-src-display').textContent = '선택 없음';
  document.getElementById('me-src-display').style.color = 'var(--text3)';
  document.getElementById('me-src-preview').textContent = '–';
  buildSrcPicker('');
  openModal('modal-metric-edit');
}

function saveMetric(){
  const m = {
    id: _editMetricIdx >= 0 ? METRICS_L[_editMetricIdx].id : Date.now(),
    name: document.getElementById('me-name').value.trim(),
    cat:  document.getElementById('me-cat').value,
    unit: document.getElementById('me-unit').value.trim(),
    pos:  document.getElementById('me-pos').value,
    desc: document.getElementById('me-desc').value.trim(),
    on:   document.getElementById('me-active').value === 'true',
    agg:  document.getElementById('me-agg').value,
    src:  document.getElementById('me-src').value.trim(),
  };
  if(!m.name){ showToast('❗ 지표명을 입력하세요'); return; }
  if(_editMetricIdx >= 0) METRICS_L[_editMetricIdx] = m;
  else METRICS_L.push(m);
  closeModal('modal-metric-edit');
  renderAdminMetrics();
  showToast(`✅ 지표 "${m.name}" ${_editMetricIdx>=0?'수정':'추가'} 완료 ${m.src?'· 데이터 연결: '+m.src:''}`);
}
function deleteMetric(){
  if(_editMetricIdx<0)return;
  const name=METRICS_L[_editMetricIdx].name;
  METRICS_L.splice(_editMetricIdx,1);
  closeModal('modal-metric-edit');
  renderAdminMetrics();
  showToast(`🗑️ 지표 "${name}" 삭제 완료`);
}

function renderAdminUsers(){
  document.getElementById('users-tbody').innerHTML=USERS_FULL.map((u,i)=>{
    const pagesText=u.pages.includes('all')?'전체':u.pages.map(p=>PAGE_LABELS[p]||p).join(', ');
    const adminText=u.admin.includes('all')?'전체 관리':u.admin.length?u.admin.map(p=>ADMIN_LABELS[p]||p).join(', '):'–';
    const pwShown=_showPwIdx.has(i);
    return`<tr>
      <td><b style="color:var(--text)">${u.name}</b></td>
      <td style="font-family:var(--mono);font-size:11px;color:var(--text3)">${u.id}</td>
      <td style="font-size:11px;color:var(--text3)">${u.email}</td>
      <td style="font-size:12px">${u.title}</td>
      <td><span class="td-tag" style="background:${u.role==='admin'?'rgba(56,217,192,.1)':'rgba(245,200,66,.1)'};color:${u.role==='admin'?'var(--teal)':'var(--gold)'}">${u.role==='admin'?'관리자':'임원'}</span></td>
      <td><div class="pw-cell"><span>${pwShown?u.pw:'••••••••'}</span><button class="pw-toggle" onclick="togglePw(${i})">${pwShown?'숨김':'표시'}</button></div></td>
      <td style="font-size:11px;color:var(--text3);max-width:120px;white-space:normal;line-height:1.5">${pagesText}</td>
      <td style="font-size:11px;color:var(--text3)">${adminText}</td>
      <td style="font-size:11px;font-family:var(--mono);color:var(--text3)">${u.last}</td>
      <td><span class="status ${u.on?'st-on':'st-off'}">${u.on?'활성':'비활성'}</span></td>
      <td><button class="btn btn-ghost" style="padding:3px 10px;font-size:10px" onclick="openUserEdit(${i})">편집</button></td>
    </tr>`;
  }).join('');
}
function togglePw(i){
  if(_showPwIdx.has(i))_showPwIdx.delete(i);
  else _showPwIdx.add(i);
  renderAdminUsers();
}
function openUserEdit(i){
  _editUserIdx=i;
  const u=USERS_FULL[i];
  document.getElementById('ue-title').textContent='사용자 편집: '+u.name;
  document.getElementById('ue-name').value=u.name;
  document.getElementById('ue-id').value=u.id;
  document.getElementById('ue-email').value=u.email;
  document.getElementById('ue-title-input')?document.getElementById('ue-title-input').value=u.title:null;
  // find correct input
  const titleInput=document.getElementById('ue-title');if(titleInput)titleInput.value=u.title;
  document.getElementById('ue-pw').value='';
  document.getElementById('ue-role').value=u.role;
  document.getElementById('ue-status').value=String(u.on);
  document.getElementById('ue-note').value=u.note||'';
  // pages checkboxes
  const pg=document.getElementById('ue-pages');
  pg.innerHTML=ALL_PAGES.map(p=>`<label class="perm-item"><input type="checkbox" value="${p}" ${(u.pages.includes('all')||u.pages.includes(p))?'checked':''}>${PAGE_LABELS[p]}</label>`).join('');
  // admin checkboxes
  const ag=document.getElementById('ue-admin');
  ag.innerHTML=ADMIN_PAGES.map(p=>`<label class="perm-item"><input type="checkbox" value="${p}" ${(u.admin.includes('all')||u.admin.includes(p))?'checked':''}>${ADMIN_LABELS[p]}</label>`).join('');
  openModal('modal-user-edit');
}
function openUserAdd(){
  _editUserIdx=-1;
  document.getElementById('ue-title').textContent='새 사용자 추가';
  ['ue-name','ue-id','ue-email','ue-pw','ue-note'].forEach(id=>{const el=document.getElementById(id);if(el)el.value='';});
  const titleEl=document.getElementById('ue-title');if(titleEl&&titleEl.tagName==='INPUT')titleEl.value='';
  document.getElementById('ue-role').value='executive';
  document.getElementById('ue-status').value='true';
  document.getElementById('ue-pages').innerHTML=ALL_PAGES.map(p=>`<label class="perm-item"><input type="checkbox" value="${p}">${PAGE_LABELS[p]}</label>`).join('');
  document.getElementById('ue-admin').innerHTML=ADMIN_PAGES.map(p=>`<label class="perm-item"><input type="checkbox" value="${p}">${ADMIN_LABELS[p]}</label>`).join('');
  openModal('modal-user-edit');
}
function saveUser(){
  const name=document.getElementById('ue-name').value.trim();
  const uid=document.getElementById('ue-id').value.trim();
  const email=document.getElementById('ue-email').value.trim();
  const title=document.getElementById('ue-title').value?document.getElementById('ue-title').value.trim():'';
  const pw=document.getElementById('ue-pw').value;
  const role=document.getElementById('ue-role').value;
  const on=document.getElementById('ue-status').value==='true';
  if(!name||!uid){showToast('❗ 이름과 아이디는 필수입니다');return;}
  const pages=[...document.querySelectorAll('#ue-pages input:checked')].map(c=>c.value);
  const admin=[...document.querySelectorAll('#ue-admin input:checked')].map(c=>c.value);
  const u={name,id:uid,email,title,role,on,pages:pages.length===ALL_PAGES.length?['all']:pages,admin:admin.length===ADMIN_PAGES.length?['all']:admin,last:new Date().toISOString().slice(0,10),note:document.getElementById('ue-note').value};
  if(_editUserIdx>=0){
    u.pw=pw||USERS_FULL[_editUserIdx].pw;
    u.idx=_editUserIdx;
    USERS_FULL[_editUserIdx]=u;
    // also update USERS_DB for auth
    const dbIdx=USERS_DB.findIndex(d=>d.id===USERS_FULL[_editUserIdx].id);
    if(dbIdx>=0&&pw)USERS_DB[dbIdx].pw=pw;
  } else {
    if(!pw){showToast('❗ 신규 사용자는 비밀번호가 필요합니다');return;}
    u.pw=pw;u.idx=USERS_FULL.length;
    USERS_FULL.push(u);
    USERS_DB.push({id:uid,pw,role,name,title,pages:'all'});
  }
  closeModal('modal-user-edit');
  renderAdminUsers();
  showToast(`✅ ${name} 사용자 정보 저장 완료`);
}
function deleteUser(){
  if(_editUserIdx<0)return;
  const name=USERS_FULL[_editUserIdx].name;
  USERS_FULL.splice(_editUserIdx,1);
  closeModal('modal-user-edit');
  renderAdminUsers();
  showToast(`🗑️ ${name} 계정이 삭제되었습니다`);
}

// ID/PW FIND
function openFindModal(){openModal('modal-find');_findTab='id';document.getElementById('find-id-tab').style.display='';document.getElementById('find-pw-tab').style.display='none';}
function switchFindTab(tab, btn){
  _findTab=tab;
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('find-id-tab').style.display=tab==='id'?'':'none';
  document.getElementById('find-pw-tab').style.display=tab==='pw'?'':'none';
  document.getElementById('find-id-result').textContent='';
  document.getElementById('find-pw-result').textContent='';
}
function doFind(){
  if(_findTab==='id'){
    const name=document.getElementById('find-name').value.trim();
    const email=document.getElementById('find-email-id').value.trim();
    if(!name||!email){showToast('❗ 이름과 이메일을 입력하세요');return;}
    const u=USERS_FULL.find(u=>u.name===name&&u.email===email);
    const el=document.getElementById('find-id-result');
    if(u){el.innerHTML=`<div style="background:rgba(52,211,138,.08);border:1px solid rgba(52,211,138,.3);border-radius:7px;padding:10px 14px;color:var(--green)">✅ 아이디: <b style="font-family:var(--mono);font-size:14px">${u.id}</b></div>`;}
    else{el.innerHTML=`<div style="color:var(--red);font-size:12px;padding:8px">❌ 일치하는 계정을 찾을 수 없습니다</div>`;}
  } else {
    const uid=document.getElementById('find-uid').value.trim();
    const email=document.getElementById('find-email-pw').value.trim();
    if(!uid||!email){showToast('❗ 아이디와 이메일을 입력하세요');return;}
    const u=USERS_FULL.find(u=>u.id===uid&&u.email===email);
    const el=document.getElementById('find-pw-result');
    if(u){
      const tmp='TMP'+Math.random().toString(36).slice(2,8).toUpperCase();
      el.innerHTML=`<div style="background:rgba(245,200,66,.08);border:1px solid rgba(245,200,66,.3);border-radius:7px;padding:10px 14px;color:var(--gold)">✅ 임시 비밀번호: <b style="font-family:var(--mono);font-size:14px">${tmp}</b><div style="font-size:10px;color:var(--text3);margin-top:4px">로그인 후 반드시 비밀번호를 변경하세요</div></div>`;
      u.pw=tmp;const db=USERS_DB.find(d=>d.id===uid);if(db)db.pw=tmp;
    }else{el.innerHTML=`<div style="color:var(--red);font-size:12px;padding:8px">❌ 아이디 또는 이메일이 올바르지 않습니다</div>`;}
  }
}

// BETA / DEPLOY
function renderAdminBeta(){
  // Changes list
  document.getElementById('beta-changes').innerHTML=_pendingChanges.map(c=>`
    <div class="diff-row ${c.type==='add'?'diff-add':c.type==='del'?'diff-del-row':'diff-mod'}">
      <span style="font-size:14px">${c.icon}</span>
      <span style="font-size:12px;color:var(--text2)">${c.msg}</span>
    </div>`).join('');
  // Deploy history
  document.getElementById('deploy-history').innerHTML=_deployHistory.map(h=>`<tr>
    <td style="font-family:var(--mono);font-size:11px">${h.dt}</td>
    <td style="font-size:11px">${h.user}</td>
    <td style="font-family:var(--mono);font-size:11px;text-align:center">${h.cnt}</td>
    <td><span class="status ${h.ok?'st-on':'st-off'}">${h.ok?'성공':'실패'}</span></td>
  </tr>`).join('');
  // Preview buttons
  document.getElementById('preview-btns').innerHTML=MENUS.executive.map(m=>`<button class="btn btn-ghost" style="font-size:11px;padding:5px 12px;border-color:var(--border)" onclick="previewPage('${m.id}')">${m.icon} ${m.label}</button>`).join('');
  // Status
  const badge=document.getElementById('beta-status-badge');
  const deployBtn=document.getElementById('deploy-btn');
  if(_deployed){
    badge.textContent='● LIVE — 배포 완료';badge.style.background='rgba(52,211,138,.15)';badge.style.color='var(--green)';badge.style.borderColor='rgba(52,211,138,.3)';
    deployBtn.disabled=true;deployBtn.textContent='✅ 배포 완료';
    document.getElementById('beta-title').textContent='현재 버전이 배포되었습니다. 모든 임원이 최신 데이터를 볼 수 있습니다.';
  }
}
function previewPage(id){
  showToast(`🖥️ "${MENUS.executive.find(m=>m.id===id)?.label}" 미리보기 — 임원 화면과 동일합니다`);
  // Switch to that page in a "preview" context
  const pe=document.getElementById('page-'+id);
  if(pe){document.querySelector('.beta-preview').innerHTML=`<div style="padding:10px;font-size:12px;color:var(--green)">✅ "${MENUS.executive.find(m=>m.id===id)?.label}" 페이지 — 배포 후 임원이 보는 화면과 동일<br><span style="font-size:10px;color:var(--text3)">실제 데이터가 반영된 최신 상태입니다</span></div>`;}
}
function doDeploy(){
  if(_deployed)return;
  const btn=document.getElementById('deploy-btn');
  btn.disabled=true;btn.textContent='⏳ 배포 중...';
  setTimeout(()=>{
    _deployed=true;
    _deployHistory.unshift({dt:new Date().toLocaleString('ko-KR').replace(/\./g,'-').replace(/\s/g,' ').slice(0,16),user:S.user?.name||'관리자',cnt:_pendingChanges.length,ok:true});
    renderAdminBeta();
    showToast('🚀 배포 완료! 임원 화면에 변경사항이 반영되었습니다');
    try{localStorage.setItem('kts_deployed','true');}catch(e){}
  },1800);
}

// ════════════════════════════════════════════════
//  🔮 전망 모달 시스템
// ════════════════════════════════════════════════
let _ftab = 'short';
let _forecastCache = {}; // tab → generated text cache

// 전망 모달 열기
function openForecast(){
  document.getElementById('forecast-modal').classList.add('open');
  document.body.style.overflow='hidden';
  renderFtab(_ftab);
}
function closeForecast(){
  document.getElementById('forecast-modal').classList.remove('open');
  document.body.style.overflow='';
}
function switchFtab(tab, btn){
  _ftab=tab;
  document.querySelectorAll('.ftab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  renderFtab(tab);
}

// 탭별 렌더
function renderFtab(tab){
  const body=document.getElementById('forecast-body');
  if(tab==='short') renderFtabShort(body);
  else if(tab==='mid') renderFtabMid(body);
  else if(tab==='risk') renderFtabRisk(body);
  else if(tab==='scenario') renderFtabScenario(body);
}

// ── 공통 데이터 요약 생성 ──
function buildForecastContext(){
  const fin=RAW.finance, wl=RAW.wireless, wd=RAW.wired;
  const last3fin = fin.months.slice(-3).map((m,i)=>({
    m, 총매출:fin.총매출[fin.총매출.length-3+i],
    영업이익:fin.영업이익[fin.영업이익.length-3+i],
    마케팅비:fin.마케팅비[fin.마케팅비.length-3+i]
  }));
  const lastN=n=>arr=>arr.slice(-n);
  return `
[최근 3개월 재무]
${last3fin.map(r=>`  ${r.m}: 총매출 ${r.총매출}억, 영업이익 ${r.영업이익}억, 마케팅비 ${r.마케팅비}억`).join('\n')}

[영업이익 6개월 추이] ${fin.영업이익.slice(-6).map((v,i)=>`${fin.months[fin.months.length-6+i]}:${v}억`).join(', ')}
[총매출 6개월 추이] ${fin.총매출.slice(-6).map((v,i)=>`${fin.months[fin.총매출.length-6+i]}:${v}억`).join(', ')}
[무선 CAPA 최근 6개월] ${wl.CAPA.slice(-6).map((v,i)=>`${wl.months[wl.CAPA.length-6+i]}:${v}명`).join(', ')}
[무선 순증 최근 6개월] ${wl.순증.slice(-6).map((v,i)=>`${v}`).join(', ')}
[유선 유지 최근월] ${wd.유지_전체[wd.유지_전체.length-1]}명 (전월 ${wd.유지_전체[wd.유지_전체.length-2]}명)
[판관비 최근 3개월] ${fin.판관비.slice(-3).map(v=>v+'억').join(', ')}
[마케팅비 최근 3개월] ${fin.마케팅비.slice(-3).map(v=>v+'억').join(', ')}`;
}

// ── 단기 전망 탭 ──
function renderFtabShort(body){
  // 데이터 기반 정적 예측값 계산
  const fin=RAW.finance;
  const last3avg=arr=>arr.slice(-3).reduce((a,b)=>a+b,0)/3;
  const trend=arr=>{const l=arr.slice(-6);const f=l.slice(0,3).reduce((a,b)=>a+b,0)/3;const s=l.slice(3).reduce((a,b)=>a+b,0)/3;return s-f;};
  const nextRev=(last3avg(fin.총매출)+trend(fin.총매출)).toFixed(1);
  const nextProfit=(last3avg(fin.영업이익)+trend(fin.영업이익)*0.5).toFixed(1);
  const nextCost=(last3avg(fin.판관비)+trend(fin.판관비)*0.3).toFixed(1);
  const revTrend=trend(fin.총매출);
  const profTrend=trend(fin.영업이익);

  body.innerHTML=`
    <div class="forecast-section">
      <div class="forecast-section-title">📊 핵심 지표 예측 (다음 1개월)</div>
      <div class="forecast-kpi-row">
        <div class="fkpi">
          <div class="fkpi-label">예상 총매출</div>
          <div class="fkpi-value" style="color:var(--gold)">${nextRev}억</div>
          <div class="fkpi-trend ${revTrend>=0?'up':'dn'}">${revTrend>=0?'▲':'▼'} ${Math.abs(revTrend).toFixed(1)}억 추세</div>
        </div>
        <div class="fkpi">
          <div class="fkpi-label">예상 영업이익</div>
          <div class="fkpi-value" style="color:${Number(nextProfit)>=0?'var(--green)':'var(--red)'}">${nextProfit}억</div>
          <div class="fkpi-trend ${profTrend>=0?'up':'dn'}">${profTrend>=0?'▲':'▼'} ${Math.abs(profTrend).toFixed(1)}억 추세</div>
        </div>
        <div class="fkpi">
          <div class="fkpi-label">예상 판관비</div>
          <div class="fkpi-value" style="color:var(--purple)">${nextCost}억</div>
          <div class="fkpi-trend neutral">3개월 평균 기반</div>
        </div>
      </div>
    </div>
    <div class="forecast-section">
      <div class="forecast-section-title">📈 향후 3개월 추세</div>
      <div class="forecast-chart-wrap">
        <canvas id="fc-short-chart"></canvas>
      </div>
    </div>
    <div class="forecast-section">
      <div class="forecast-section-title">✦ AI 단기 전망 분석</div>
      <div class="forecast-insight" id="fc-short-ai">
        <div style="font-size:11px;color:var(--text3);text-align:center;padding:8px">
          아래 버튼으로 AI 분석을 실행하세요
        </div>
      </div>
      <button class="forecast-run-btn" id="fc-short-btn" onclick="runForecastAI('short')">
        ✦ AI 단기 전망 분석 실행
      </button>
    </div>`;

  // 예측 차트 (최근 6개월 + 예측 3개월)
  const fin2=RAW.finance;
  const recentMonths=fin2.months.slice(-6);
  const recentRev=fin2.총매출.slice(-6);
  const recentProfit=fin2.영업이익.slice(-6);
  const predMonths=['예측+1','예측+2','예측+3'];
  const t=trend(fin2.총매출), tp=trend(fin2.영업이익);
  const lastRev=fin2.총매출[fin2.총매출.length-1];
  const lastProf=fin2.영업이익[fin2.영업이익.length-1];
  const predRev=[lastRev+t*0.6,lastRev+t*1.1,lastRev+t*1.5];
  const predProf=[lastProf+tp*0.5,lastProf+tp*0.9,lastProf+tp*1.2];
  const allMonths=[...recentMonths,...predMonths];
  const allRev=[...recentRev,...predRev];
  const allProfit=[...recentProfit,...predProf];
  setTimeout(()=>{
    mkC('fc-short-chart',{type:'line',data:{labels:allMonths,datasets:[
      {label:'총매출',data:allRev,borderColor:C.gold,backgroundColor:'transparent',borderWidth:2,pointRadius:allMonths.map((_,i)=>i>=6?5:2),pointStyle:allMonths.map((_,i)=>i>=6?'triangle':'circle'),tension:.3,borderDash:allRev.map((_,i)=>i>=5?[4,4]:[]),segment:{borderDash:ctx=>ctx.p0DataIndex>=5?[5,5]:undefined}},
      {label:'영업이익',data:allProfit,borderColor:C.green,backgroundColor:'transparent',borderWidth:2,pointRadius:allMonths.map((_,i)=>i>=6?5:2),tension:.3,yAxisID:'y2',segment:{borderDash:ctx=>ctx.p0DataIndex>=5?[5,5]:undefined}},
    ]},options:baseOpts({scales:{x:{grid:{color:'rgba(42,53,85,0.5)'},ticks:{maxRotation:45,color:'#5c6e9a',font:{size:10}},afterFit:function(axis){axis.paddingRight=10;}},y:{grid:{color:'rgba(42,53,85,.5)'},ticks:{color:'#5c6e9a'}},y2:{position:'right',grid:{display:false},ticks:{color:'#5c6e9a'}}}})});
    // 예측 구간 배경 표시용 annotation 없이, 차트 위에 텍스트 표시
  },80);

  if(_forecastCache['short']){
    document.getElementById('fc-short-ai').textContent=_forecastCache['short'];
  }
}

// ── 중기 전망 탭 ──
function renderFtabMid(body){
  body.innerHTML=`
    <div class="forecast-section">
      <div class="forecast-section-title">📊 중기 예측 (향후 6개월 누적)</div>
      <div class="forecast-chart-wrap">
        <canvas id="fc-mid-chart"></canvas>
      </div>
    </div>
    <div class="forecast-section">
      <div class="forecast-section-title">✦ AI 중기 전망 분석</div>
      <div class="forecast-insight" id="fc-mid-ai">
        <div style="font-size:11px;color:var(--text3);text-align:center;padding:8px">아래 버튼으로 AI 분석을 실행하세요</div>
      </div>
      <button class="forecast-run-btn" id="fc-mid-btn" onclick="runForecastAI('mid')">
        ✦ AI 중기 전망 분석 실행
      </button>
    </div>`;

  const fin=RAW.finance;
  const avg6=arr=>arr.slice(-6).reduce((a,b)=>a+b,0)/6;
  const trend6=arr=>{const l=arr.slice(-6);const f=l.slice(0,3).reduce((a,b)=>a+b,0)/3;const s=l.slice(3).reduce((a,b)=>a+b,0)/3;return(s-f)/3;};
  const pred=(arr,n)=>{const last=arr[arr.length-1];const t=trend6(arr);return Array.from({length:n},(_,i)=>parseFloat((last+t*(i+1)).toFixed(1)));};
  const recentMonths=fin.months.slice(-6);
  const predMonths=['예측M+1','M+2','M+3','M+4','M+5','M+6'];
  const allMon=[...recentMonths,...predMonths];
  setTimeout(()=>{
    mkC('fc-mid-chart',{type:'line',data:{labels:allMon,datasets:[
      {label:'총매출',data:[...fin.총매출.slice(-6),...pred(fin.총매출,6)],borderColor:C.gold,backgroundColor:C.goldA,borderWidth:2,fill:false,tension:.35,pointRadius:allMon.map((_,i)=>i>=6?4:2),segment:{borderDash:ctx=>ctx.p0DataIndex>=5?[5,5]:undefined}},
      {label:'영업이익',data:[...fin.영업이익.slice(-6),...pred(fin.영업이익,6)],borderColor:C.green,backgroundColor:'transparent',borderWidth:2,tension:.35,pointRadius:allMon.map((_,i)=>i>=6?4:2),yAxisID:'y2',segment:{borderDash:ctx=>ctx.p0DataIndex>=5?[5,5]:undefined}},
      {label:'판관비',data:[...fin.판관비.slice(-6),...pred(fin.판관비,6)],borderColor:C.purple,backgroundColor:'transparent',borderWidth:1.5,tension:.35,pointRadius:0,segment:{borderDash:ctx=>ctx.p0DataIndex>=5?[4,4]:undefined}},
    ]},options:baseOpts({scales:{x:{grid:{color:'rgba(42,53,85,.5)'},ticks:{maxRotation:45,color:'#5c6e9a',font:{size:9}}},y:{grid:{color:'rgba(42,53,85,.5)'},ticks:{color:'#5c6e9a'}},y2:{position:'right',grid:{display:false},ticks:{color:'#5c6e9a'}}}})});
  },80);

  if(_forecastCache['mid']){
    document.getElementById('fc-mid-ai').textContent=_forecastCache['mid'];
  }
}

// ── 리스크 탭 ──
function renderFtabRisk(body){
  const risks=[
    {level:'high',icon:'🔴',title:'영업이익 적자 지속',desc:'최근 3개월 중 2개월 영업이익 적자 기록. 고정비 증가 추세와 마케팅비 급등 시 수익성 위협 가중.'},
    {level:'high',icon:'🔴',title:'무선 순증 마이너스 추세',desc:'지속적인 유지가입자 감소 (-24,490명/최근월). 해지율 증가 시 매출 기반 축소 가속화 우려.'},
    {level:'mid',icon:'🟡',title:'유선 가입자 이탈 가속',desc:'유선 유지가입자 248,042명으로 지속 하락 중. 인터넷 경쟁 심화 시 유선매출 추가 하락 가능성.'},
    {level:'mid',icon:'🟡',title:'마케팅비 변동성',desc:'마케팅비가 최근 6개월 15~88억 구간 급변동. 경쟁사 판촉 확대 시 추가 지출 압박 불가피.'},
    {level:'low',icon:'🟢',title:'디지털 채널 성장 기회',desc:'KT닷컴 CAPA가 최고 30,235명 기록 후 조정 중. 디지털 전환 가속화 시 비용 효율 개선 여지.'},
    {level:'low',icon:'🟢',title:'GiGAeyes 성장 모멘텀',desc:'GiGAeyes 최근 2,459대 기록으로 전략상품 확장 중. B2B·소상공인 채널 성장 견인 가능.'},
  ];
  body.innerHTML=`
    <div class="forecast-section">
      <div class="forecast-section-title">⚠ 리스크 레이더</div>
      ${risks.map(r=>`
        <div class="risk-item risk-${r.level}">
          <span class="risk-badge">${r.level==='high'?'HIGH':r.level==='mid'?'MID':'LOW'}</span>
          <div><b style="font-size:12px">${r.icon} ${r.title}</b><br>${r.desc}</div>
        </div>`).join('')}
    </div>
    <div class="forecast-section">
      <div class="forecast-section-title">✦ AI 리스크 심층 분석</div>
      <div class="forecast-insight" id="fc-risk-ai">
        <div style="font-size:11px;color:var(--text3);text-align:center;padding:8px">아래 버튼으로 AI 분석을 실행하세요</div>
      </div>
      <button class="forecast-run-btn" id="fc-risk-btn" onclick="runForecastAI('risk')">
        ✦ AI 리스크 분석 실행
      </button>
    </div>`;
  if(_forecastCache['risk']){
    document.getElementById('fc-risk-ai').textContent=_forecastCache['risk'];
  }
}

// ── 시나리오 탭 ──
function renderFtabScenario(body){
  const fin=RAW.finance;
  const avg3rev=fin.총매출.slice(-3).reduce((a,b)=>a+b,0)/3;
  const avg3prof=fin.영업이익.slice(-3).reduce((a,b)=>a+b,0)/3;
  body.innerHTML=`
    <div class="forecast-section">
      <div class="forecast-section-title">🎯 3개월 시나리오 분석</div>
      <div class="scenario-card" style="border-color:rgba(52,211,138,.3);background:rgba(52,211,138,.04)">
        <div class="scenario-label" style="color:var(--green)">🟢 낙관 시나리오 (25% 확률)</div>
        <div class="forecast-kpi-row" style="margin-bottom:8px">
          <div class="fkpi"><div class="fkpi-label">총매출</div><div class="fkpi-value" style="color:var(--green);font-size:14px">${(avg3rev*1.15).toFixed(0)}억</div></div>
          <div class="fkpi"><div class="fkpi-label">영업이익</div><div class="fkpi-value" style="color:var(--green);font-size:14px">${(Math.max(avg3prof,0)*1.3+20).toFixed(0)}억</div></div>
          <div class="fkpi"><div class="fkpi-label">이익률</div><div class="fkpi-value" style="color:var(--green);font-size:14px">${((Math.max(avg3prof,0)*1.3+20)/(avg3rev*1.15)*100).toFixed(1)}%</div></div>
        </div>
        <div class="scenario-desc">마케팅 투자 효율화로 CAPA 반등, 디지털 채널 성장 가속. 무선 순증 플러스 전환 시 매출 상승 견인.</div>
      </div>
      <div class="scenario-card" style="border-color:rgba(245,200,66,.3);background:rgba(245,200,66,.04)">
        <div class="scenario-label" style="color:var(--gold)">🟡 기준 시나리오 (50% 확률)</div>
        <div class="forecast-kpi-row" style="margin-bottom:8px">
          <div class="fkpi"><div class="fkpi-label">총매출</div><div class="fkpi-value" style="color:var(--gold);font-size:14px">${(avg3rev*1.0).toFixed(0)}억</div></div>
          <div class="fkpi"><div class="fkpi-label">영업이익</div><div class="fkpi-value" style="color:var(--gold);font-size:14px">${(avg3prof*1.0+5).toFixed(0)}억</div></div>
          <div class="fkpi"><div class="fkpi-label">이익률</div><div class="fkpi-value" style="color:var(--gold);font-size:14px">${((avg3prof*1.0+5)/(avg3rev*1.0)*100).toFixed(1)}%</div></div>
        </div>
        <div class="scenario-desc">현 추세 유지. 비용 구조 개선 소폭 진행. 무선 순증 마이너스 지속되나 감소폭 완화.</div>
      </div>
      <div class="scenario-card" style="border-color:rgba(255,94,106,.3);background:rgba(255,94,106,.04)">
        <div class="scenario-label" style="color:var(--red)">🔴 비관 시나리오 (25% 확률)</div>
        <div class="forecast-kpi-row" style="margin-bottom:8px">
          <div class="fkpi"><div class="fkpi-label">총매출</div><div class="fkpi-value" style="color:var(--red);font-size:14px">${(avg3rev*0.88).toFixed(0)}억</div></div>
          <div class="fkpi"><div class="fkpi-label">영업이익</div><div class="fkpi-value" style="color:var(--red);font-size:14px">${(avg3prof-25).toFixed(0)}억</div></div>
          <div class="fkpi"><div class="fkpi-label">이익률</div><div class="fkpi-value" style="color:var(--red);font-size:14px">${((avg3prof-25)/(avg3rev*0.88)*100).toFixed(1)}%</div></div>
        </div>
        <div class="scenario-desc">경쟁 심화로 마케팅비 급증, 무선 해지 증가. 유선 가입자 이탈 가속 및 판관비 상승 압박 가중.</div>
      </div>
    </div>
    <div class="forecast-section">
      <div class="forecast-section-title">✦ AI 시나리오 심층 분석</div>
      <div class="forecast-insight" id="fc-scenario-ai">
        <div style="font-size:11px;color:var(--text3);text-align:center;padding:8px">아래 버튼으로 AI 분석을 실행하세요</div>
      </div>
      <button class="forecast-run-btn" id="fc-scenario-btn" onclick="runForecastAI('scenario')">
        ✦ AI 시나리오 분석 실행
      </button>
    </div>`;
  if(_forecastCache['scenario']){
    document.getElementById('fc-scenario-ai').textContent=_forecastCache['scenario'];
  }
}

// ── AI 전망 분석 실행 ──
async function runForecastAI(tab){
  const bodyId=`fc-${tab}-ai`, btnId=`fc-${tab}-btn`;
  setAILoading(bodyId, btnId);
  const ctx=buildForecastContext();
  const prompts={
    short:`아래 KT M&S 데이터를 기반으로 향후 1~3개월 단기 전망을 분석하세요:\n${ctx}\n\n포함 항목:\n1. 다음 달 총매출/영업이익 예측 및 근거\n2. 주목해야 할 변수 2~3가지\n3. 권고 액션 1~2가지\n(3문단 이내로 간결하게)`,
    mid:`아래 KT M&S 데이터를 기반으로 향후 6개월 중기 전망을 분석하세요:\n${ctx}\n\n포함 항목:\n1. 6개월 매출/이익 방향성 전망\n2. 채널별 성장/위축 예상\n3. 비용 구조 개선 포인트\n(3~4문단)`,
    risk:`아래 KT M&S 데이터를 기반으로 리스크 요인을 심층 분석하세요:\n${ctx}\n\n포함 항목:\n1. 가장 시급한 리스크 2가지와 대응 방안\n2. 선제적 관리가 필요한 선행 지표\n3. 기회 요인 1~2가지\n(3문단)`,
    scenario:`아래 KT M&S 데이터를 기반으로 시나리오별 전략적 시사점을 분석하세요:\n${ctx}\n\n포함 항목:\n1. 낙관/기준/비관 시나리오 각각의 핵심 가정\n2. 기준 시나리오 달성을 위한 최우선 과제\n3. 비관 시나리오 발생 시 즉각 조치 사항\n(3~4문단)`,
  };
  const sysPrompts={
    short:'당신은 KT M&S의 단기 재무 전략 분석가입니다. 데이터 기반 예측과 실행 가능한 인사이트를 제공합니다.',
    mid:'당신은 KT M&S의 중기 경영전략 분석가입니다. 채널별 성과와 비용 구조를 종합적으로 분석합니다.',
    risk:'당신은 KT M&S의 리스크 관리 전문가입니다. 선제적 위험 식별과 대응 방안을 제시합니다.',
    scenario:'당신은 KT M&S의 경영 시나리오 플래닝 전문가입니다. 복수 시나리오의 전략적 의미를 도출합니다.',
  };
  try{
    const result=await callAI(prompts[tab], sysPrompts[tab]);
    _forecastCache[tab]=result;
    setAIResult(bodyId, btnId, result);
  }catch(e){
    setAIError(bodyId, btnId, e.message);
  }
}

// 전망 버튼 + API 키 버튼 표시/숨김 (로그인 후)
function showForecastBtn(role){
  const btn = document.getElementById('forecast-btn');
  const kbtn = document.getElementById('apikey-btn');
  // 전망 버튼 → 임원만
  if(btn) btn.style.display = role==='executive' ? '' : 'none';
  // API 키 버튼 → 관리자(admin)만
  if(kbtn){
    const user = S.user;
    const isAdmin = role==='admin' && user && user.id==='admin';
    kbtn.style.display = isAdmin ? '' : 'none';
    if(isAdmin) updateApiKeyBtn();
  }
}
function toggleSidebar(){
  const sb=document.getElementById('sidebar');
  const ov=document.getElementById('sidebar-overlay');
  const open=sb.classList.toggle('open');
  ov.classList.toggle('open',open);
}
function closeSidebar(){
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('sidebar-overlay').classList.remove('open');
}

function buildSidebar(role){
  const menus=MENUS[role];
  const sb=document.getElementById('sidebar');
  const bn=document.getElementById('bottom-nav');
  let html='',sec='';
  menus.forEach(m=>{
    if(m.sec!==sec){html+=`<div class="sidebar-label">${m.sec}</div>`;sec=m.sec;}
    html+=`<div class="nav-item" id="nav-${m.id}" onclick="nav('${m.id}')">${isMobile()?'':''}<span class="ni">${m.icon}</span>${m.label}</div>`;
  });
  sb.innerHTML=html;
  // Bottom nav: show first 5 most-used items
  const top5=menus.slice(0,Math.min(5,menus.length));
  bn.innerHTML=top5.map(m=>`<button class="bnav-item" id="bnav-${m.id}" onclick="nav('${m.id}')"><span class="bni">${m.icon}</span>${m.label.slice(0,4)}</button>`).join('');
}

function isMobile(){return window.innerWidth<=640;}

function nav(id){
  document.querySelectorAll('.nav-item').forEach(e=>e.classList.remove('active'));
  document.querySelectorAll('.bnav-item').forEach(e=>e.classList.remove('active'));
  document.querySelectorAll('.page').forEach(e=>e.classList.remove('active'));
  const ne=document.getElementById('nav-'+id);
  const bne=document.getElementById('bnav-'+id);
  const pe=document.getElementById('page-'+id);
  if(ne)ne.classList.add('active');
  if(bne)bne.classList.add('active');
  if(pe)pe.classList.add('active');
  S.page=id;
  if(RENDER_MAP[id])RENDER_MAP[id]();
  document.getElementById('topbar-period').textContent=getPL();
  // close sidebar on mobile after navigation
  if(isMobile())closeSidebar();
}
