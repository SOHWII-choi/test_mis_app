// KT M&S 경영성과 대시보드 — 앱 로직
// ─────────────────────────────────
const RENDER_MAP={
  finance:()=>{buildPB('pb-finance','renderFinance');renderFinance();},
  wireless:()=>{buildPB('pb-wireless','renderWireless');renderWireless();},
  wired:()=>{buildPB('pb-wired','renderWired');renderWired();},
  org:()=>{buildPB('pb-org','renderOrg');renderOrg();},
  digital:()=>{buildPB('pb-digital','renderDigital');renderDigital();},
  b2b:()=>{buildPB('pb-b2b','renderB2B');renderB2B();},
  smb:()=>{buildPB('pb-smb','renderSMB');renderSMB();},
  platform:()=>{buildPB('pb-platform','renderPlatform');renderPlatform();},
  quality:()=>{buildPB('pb-quality','renderQuality');renderQuality();},
  infra:()=>{buildPB('pb-infra','renderInfra');renderInfra();},
  strategy:()=>{buildPB('pb-strategy','renderStrategy');renderStrategy();},
  hr:()=>{buildPB('pb-hr','renderHR');renderHR();},
  'admin-overview':renderAdminOverview,
  'admin-upload':renderAdminUpload,
  'admin-metrics':renderAdminMetrics,
  'admin-users':renderAdminUsers,
  'admin-export':()=>{},
  'admin-beta':renderAdminBeta,
};

// nav() defined above with responsive logic
function _noop(){}

// ═══════════════════════════════════════════
//  SUPABASE 초기화
// ═══════════════════════════════════════════
const SB_URL = 'https://cypwrpkiltglbogixpva.supabase.co';
const SB_KEY = 'sb_publishable_B6EgEqqgfdoloD_OgLVGJg_rKShV2D3';
const sb = supabase.createClient(SB_URL, SB_KEY);

// ── DB 상태 표시 ──
function setDbStatus(msg){ const el=document.getElementById('db-load-text'); if(el)el.textContent=msg; }
function hideDbLoading(){
  const el=document.getElementById('db-loading');
  if(!el)return;
  el.classList.add('fade-out');
  setTimeout(()=>el.style.display='none',420);
}

// ══════════════════════════════════════════
//  USER DB & AUTH (로컬 fallback 포함)
// ══════════════════════════════════════════
const USERS_DB_LOCAL = [
  {user_id:'ceo',  pw:'kt2025',    role:'executive', name:'김철수',   title:'대표이사',     pages:'all', active:true},
  {user_id:'cfo',  pw:'kt2025',    role:'executive', name:'이영희',   title:'재무본부장',   pages:'all', active:true},
  {user_id:'cso',  pw:'kt2025',    role:'executive', name:'박민준',   title:'영업총괄',     pages:'all', active:true},
  {user_id:'cmo',  pw:'kt2025',    role:'executive', name:'정지현',   title:'마케팅본부장', pages:'all', active:true},
  {user_id:'sales1',pw:'kt2025',   role:'executive', name:'최현우',   title:'소매영업팀장', pages:'all', active:true},
  {user_id:'admin',pw:'admin1234', role:'admin',     name:'최관리',   title:'시스템관리자', pages:'all', active:true},
  {user_id:'data', pw:'data1234',  role:'admin',     name:'정데이터', title:'데이터팀',     pages:'all', active:true},
];

// Supabase에서 유저 조회, 실패 시 로컬 fallback
async function findUserDB(id, pw){
  try{
    const {data,error}=await sb.from('kts_users').select('*').eq('user_id',id).eq('pw',pw).eq('active',true).single();
    if(!error && data) return {id:data.user_id, pw:data.pw, role:data.role, name:data.name, title:data.title||'', pages:data.pages||'all'};
  }catch(e){}
  // 로컬 fallback
  const u=USERS_DB_LOCAL.find(u=>u.user_id===id&&u.pw===pw&&u.active);
  return u?{id:u.user_id,pw:u.pw,role:u.role,name:u.name,title:u.title,pages:u.pages}:null;
}
// 동기 버전 (자동로그인용)
function findUser(id,pw){return USERS_DB_LOCAL.find(u=>u.user_id===id&&u.pw===pw&&u.active)?{...USERS_DB_LOCAL.find(u=>u.user_id===id&&u.pw===pw),id:USERS_DB_LOCAL.find(u=>u.user_id===id&&u.pw===pw).user_id}:null;}

async function doLogin(){
  const id=document.getElementById('li-id').value.trim();
  const pw=document.getElementById('li-pw').value;
  const err=document.getElementById('li-err');
  if(!id||!pw){err.textContent='아이디와 비밀번호를 입력하세요.';return;}
  const btn=document.querySelector('.li-btn');
  btn.textContent='확인 중...';btn.disabled=true;
  const u=await findUserDB(id,pw);
  btn.textContent='로그인';btn.disabled=false;
  if(!u){err.textContent='아이디 또는 비밀번호가 올바르지 않습니다.';shakeInput('li-pw');return;}
  if(u.role==='admin'){err.textContent='관리자는 우측 하단 "관리자 로그인"을 이용하세요.';return;}
  err.textContent='';
  if(document.getElementById('li-remember').checked) saveSession(id,pw,'exec');
  // 마지막 로그인 기록 (비동기)
  sb.from('kts_users').update({last_login:new Date().toISOString()}).eq('user_id',id).then(()=>{});
  await loginUser(u);
}

async function doAdminLogin(){
  const id=document.getElementById('adm-id').value.trim();
  const pw=document.getElementById('adm-pw').value;
  const err=document.getElementById('adm-err');
  if(!id||!pw){err.textContent='아이디와 비밀번호를 입력하세요.';return;}
  const btn=document.querySelector('[onclick="doAdminLogin()"]');
  if(btn){btn.textContent='확인 중...';btn.disabled=true;}
  const u=await findUserDB(id,pw);
  if(btn){btn.textContent='관리자 접속';btn.disabled=false;}
  if(!u||u.role!=='admin'){err.textContent='관리자 계정 정보가 올바르지 않습니다.';shakeInput('adm-pw');return;}
  err.textContent='';
  if(document.getElementById('adm-remember').checked) saveSession(id,pw,'admin');
  sb.from('kts_users').update({last_login:new Date().toISOString()}).eq('user_id',id).then(()=>{});
  closeAdminModal();
  await loginUser(u);
}

async function loginUser(u){
  S.role=u.role; S.user=u;
  await loadDataFromDB();
  document.getElementById('login-page').style.display='none';
  document.getElementById('app').style.display='flex';
  const chip=document.getElementById('role-chip');
  chip.textContent=`${u.name} ${u.title}`;
  chip.className='role-chip '+(u.role==='admin'?'admin':'executive');
  showForecastBtn(u.role);
  buildSidebar(u.role);
  nav(MENUS[u.role][0].id);
  // PWA 설치 배너 (한 번도 설치 안 한 경우만)
  setTimeout(showPwaBanner, 2000);
}

function showPwaBanner(){
  if(localStorage.getItem('pwa_dismissed')) return;
  if(window.matchMedia('(display-mode: standalone)').matches) return; // 이미 설치됨
  if(!window._pwaPrompt && !navigator.userAgent.includes('iPhone') && !navigator.userAgent.includes('iPad')) return;
  const existing=document.getElementById('pwa-banner');
  if(existing) return;
  const banner=document.createElement('div');
  banner.id='pwa-banner';
  const isIOS=navigator.userAgent.includes('iPhone')||navigator.userAgent.includes('iPad');
  banner.innerHTML=`
    <div class="pwa-banner-icon">📱</div>
    <div class="pwa-banner-text">
      <div class="pwa-banner-title">앱으로 설치하기</div>
      <div class="pwa-banner-sub">${isIOS?'Safari 하단 공유 버튼 → 홈 화면에 추가':'홈 화면에 추가하면 앱처럼 사용 가능'}</div>
    </div>
    <div class="pwa-banner-btns">
      ${isIOS?'':`<button class="pwa-install-btn" onclick="window._pwaPrompt&&window._pwaPrompt()">설치</button>`}
      <button class="pwa-dismiss-btn" onclick="localStorage.setItem('pwa_dismissed','1');this.closest('#pwa-banner').remove()">✕</button>
    </div>`;
  document.body.appendChild(banner);
}

// ══════════════════════════════════════════
//  DB에서 데이터 로드 (실패 시 RAW 유지)
// ══════════════════════════════════════════
async function loadDataFromDB(){
  try{
    setDbStatus('경영성과 데이터 로드 중...');
    const {data,error}=await sb.from('kts_data').select('category,data');
    if(error||!data||data.length===0){
      console.log('DB 데이터 없음 — 내장 데이터 사용');
      return;
    }
    let loaded=0;
    for(const row of data){
      if(RAW[row.category] && row.data){
        RAW[row.category]=row.data;
        loaded++;
      }
    }
    if(loaded>0){
      setDbStatus(`✅ DB에서 ${loaded}개 카테고리 로드 완료`);
      console.log(`Supabase DB에서 ${loaded}개 카테고리 데이터 로드`);
      showToast(`☁️ 최신 데이터 로드 완료 (${loaded}개 카테고리)`);
    }
  }catch(e){
    console.warn('DB 로드 실패, 내장 데이터 사용:', e.message);
  }
}

function logout(){
  Object.values(S.charts).forEach(c=>c.destroy());S.charts={};
  clearSession();
  document.getElementById('login-page').style.display='flex';
  document.getElementById('app').style.display='none';
  document.getElementById('li-id').value='';
  document.getElementById('li-pw').value='';
  document.getElementById('li-err').textContent='';
}

// Auto-login via localStorage
function saveSession(id,pw,type){try{localStorage.setItem('kts_session',JSON.stringify({id,pw,type,ts:Date.now()}));}catch(e){}}
function clearSession(){try{localStorage.removeItem('kts_session');}catch(e){}}
async function tryAutoLogin(){
  try{
    const raw=localStorage.getItem('kts_session');
    if(!raw)return false;
    const s=JSON.parse(raw);
    if(Date.now()-s.ts > 30*24*3600*1000){clearSession();return false;}
    setDbStatus('자동 로그인 확인 중...');
    const u=await findUserDB(s.id,s.pw);
    if(!u)return false;
    await loginUser(u);
    showToast(`👋 자동 로그인: ${u.name} (${u.title})`);
    return true;
  }catch(e){return false;}
}

// Admin modal
function openAdminModal(){document.getElementById('admin-modal').classList.add('open');setTimeout(()=>document.getElementById('adm-id').focus(),80);}
function closeAdminModal(){document.getElementById('admin-modal').classList.remove('open');document.getElementById('adm-err').textContent='';}
document.getElementById('admin-modal').addEventListener('click',function(e){if(e.target===this)closeAdminModal();});

function shakeInput(id){
  const el=document.getElementById(id);
  el.style.borderColor='var(--red)';
  el.animate([{transform:'translateX(-4px)'},{transform:'translateX(4px)'},{transform:'translateX(-3px)'},{transform:'translateX(3px)'},{transform:'translateX(0)'}],{duration:300});
  setTimeout(()=>{el.style.borderColor='';},1200);
}

// ── 앱 시작 ──
window.addEventListener('DOMContentLoaded', async ()=>{
  // ① Service Worker 등록 (PWA)
  if('serviceWorker' in navigator){
    navigator.serviceWorker.register('service-worker.js')
      .then(reg=>{
        console.log('[PWA] Service Worker 등록 완료:', reg.scope);
        // 새 버전 감지 시 업데이트 토스트
        reg.addEventListener('updatefound', ()=>{
          const newSW = reg.installing;
          newSW.addEventListener('statechange', ()=>{
            if(newSW.state==='installed' && navigator.serviceWorker.controller){
              showToast('🔄 새 버전이 있어요! 새로고침하면 업데이트됩니다');
            }
          });
        });
      })
      .catch(err=>console.warn('[PWA] SW 등록 실패:', err));
  }

  // ② 홈화면 추가 배너 (안드로이드 크롬)
  let deferredPrompt;
  window.addEventListener('beforeinstallprompt', e=>{
    e.preventDefault();
    deferredPrompt=e;
    // 로그인 후 배너 표시
    window._pwaPrompt=()=>{
      if(deferredPrompt){
        deferredPrompt.prompt();
        deferredPrompt.userChoice.then(choice=>{
          if(choice.outcome==='accepted') showToast('✅ 홈 화면에 설치되었습니다!');
          deferredPrompt=null;
          document.getElementById('pwa-banner')?.remove();
        });
      }
    };
  });

  // ③ Supabase 연결 확인
  try{
    setDbStatus('Supabase 연결 중...');
    await sb.from('kts_users').select('user_id').limit(1);
    setDbStatus('✅ 연결 성공');
  }catch(e){
    setDbStatus('⚠ 오프라인 모드 (내장 데이터 사용)');
  }
  await new Promise(r=>setTimeout(r,400));
  hideDbLoading();
  const loggedIn = await tryAutoLogin();
  if(!loggedIn){
    document.getElementById('login-page').style.display='flex';
  }
});
// MODAL UTILS
function openModal(id){document.getElementById(id).classList.add('open');}
function closeModal(id){document.getElementById(id).classList.remove('open');}
document.addEventListener('click',e=>{if(e.target.classList.contains('gmodal'))closeModal(e.target.id);});

// EXPORT – CSV download
function exportCSV(type){
  const start=document.getElementById('exp-start')?.value||'2025-01';
  const end=document.getElementById('exp-end')?.value||'2025-12';
  let csvParts=[];

  function buildSection(title, src, keys){
    const d=RAW[src];
    if(!d)return;
    csvParts.push(`\n# ${title} (${start} ~ ${end})`);
    csvParts.push(['월',...keys].join(','));
    d.months.forEach((m,i)=>{
      const row=[m,...keys.map(k=>d[k]?.[i]??'')];
      csvParts.push(row.join(','));
    });
  }

  if(type==='finance'||type==='all') buildSection('재무','finance',['총매출','통신매출','상품매출','영업이익','판관비','인건비','마케팅비']);
  if(type==='wireless'||type==='all') buildSection('무선 가입자','wireless',['유지','CAPA','해지','순증']);
  if(type==='all'){
    buildSection('유선 가입자','wired',['유지_전체','신규_전체','해지_전체','순증_전체','유지_인터넷','신규_인터넷']);
    buildSection('조직별 실적','org',['소매','도매','디지털','B2B','소상공인']);
    buildSection('디지털 채널','digital',['일반후불_총계','운영후불_총계','유심단독','유선순신규','인력_총계']);
    buildSection('B2B 채널','b2b',['전체_일반후불','전체_무선순증','가입자_무선유지','가입자_기업무선','가입자_법인무선']);
    buildSection('소상공인','smb',['상품_일반후불','상품_운영후불','인터넷순신규','인력_총원']);
    buildSection('유통플랫폼','platform',['중고폰_매입건수','중고폰_매각건수','중고폰_매입금액','시연폰_매각건수']);
    buildSection('TCSI','tcsi',['TCSI점수','KT점수','대리점점수']);
    buildSection('인력','hr',['전사계','임원','일반직','영업직','SC직','소매채널','도매채널']);
    buildSection('인프라','infra',['소매매장수_계','출점_계','퇴점_계','도매무선취급점','점당생산성_무선']);
    buildSection('전략상품','strategy',['하이오더_점포수','하이오더_태블릿수','GiGAeyes_계','AI전화_계']);
  }

  const csv='\uFEFF'+csvParts.join('\n'); // BOM for Excel Korean
  const blob=new Blob([csv],{type:'text/csv;charset=utf-8;'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;
  const fname=type==='all'?`KTS_경영성과_전체_${start}~${end}.csv`:`KTS_경영성과_${type}_${start}~${end}.csv`;
  a.download=fname;
  document.body.appendChild(a);a.click();document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast(`📥 "${fname}" 다운로드 시작`);
}

// EXPORT – PDF (print)
function exportPDF(){
  const type=document.getElementById('pdf-type')?.value||'summary';
  const fin=RAW.finance;
  const lastIdx=fin.months.length-1;
  const fmtN=(v,d=1)=>Number(v).toLocaleString('ko-KR',{minimumFractionDigits:d,maximumFractionDigits:d});

  const kpiRows=type==='summary'?[
    ['총매출',     fmtN(fin.총매출[lastIdx],1)+'억원',   fmtN(fin.총매출.reduce((a,b)=>a+b,0),1)+'억원'],
    ['통신매출',   fmtN(fin.통신매출[lastIdx],1)+'억원', fmtN(fin.통신매출.reduce((a,b)=>a+b,0),1)+'억원'],
    ['영업이익',   fmtN(fin.영업이익[lastIdx],1)+'억원', fmtN(fin.영업이익.reduce((a,b)=>a+b,0),1)+'억원'],
    ['무선 CAPA', Number(RAW.wireless.CAPA[lastIdx]).toLocaleString()+'명', ''],
    ['유선 유지',  Number(RAW.wired.유지_전체[lastIdx]).toLocaleString()+'명',''],
    ['TCSI 점수',  fmtN(RAW.tcsi.TCSI점수[lastIdx],1)+'점', ''],
    ['전사 인원',  Number(RAW.hr.전사계[lastIdx]).toLocaleString()+'명', ''],
  ]:[
    ['총매출',fmtN(fin.총매출[lastIdx],1)+'억원',''],
    ['통신매출',fmtN(fin.통신매출[lastIdx],1)+'억원',''],
    ['영업이익',fmtN(fin.영업이익[lastIdx],1)+'억원',''],
    ['판관비',fmtN(fin.판관비[lastIdx],1)+'억원',''],
    ['무선 CAPA',Number(RAW.wireless.CAPA[lastIdx]).toLocaleString()+'명',''],
    ['무선 순증',Number(RAW.wireless.순증[lastIdx]).toLocaleString()+'명',''],
    ['유선 유지',Number(RAW.wired.유지_전체[lastIdx]).toLocaleString()+'명',''],
    ['TCSI점수',fmtN(RAW.tcsi.TCSI점수[lastIdx],1)+'점',''],
    ['VOC발생률',fmtN(RAW.voc.도소매발생률[lastIdx],4)+'%',''],
    ['소매 매장수',RAW.infra.소매매장수_계[lastIdx]+'개',''],
  ];

  const w=window.open('','_blank','width=860,height=1100');
  w.document.write(`<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">
  <title>KT M&amp;S 경영성과 보고서</title>
  <style>
    *{margin:0;padding:0;box-sizing:border-box}
    body{font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif;font-size:12px;color:#1a1a2e;background:#fff;padding:32px 36px}
    h1{font-size:22px;font-weight:700;color:#0c1120;margin-bottom:4px}
    .sub{font-size:11px;color:#666;margin-bottom:28px}
    .section{margin-bottom:24px}
    h2{font-size:14px;font-weight:700;border-bottom:2px solid #f5c842;padding-bottom:5px;margin-bottom:12px;color:#0c1120}
    table{width:100%;border-collapse:collapse;font-size:11px}
    th{background:#0c1120;color:#f5c842;padding:8px 12px;text-align:left;font-size:10px;letter-spacing:.5px}
    td{padding:8px 12px;border-bottom:1px solid #eee;color:#333}
    tr:nth-child(even) td{background:#f9f9fc}
    .footer{margin-top:40px;font-size:10px;color:#999;border-top:1px solid #eee;padding-top:10px;display:flex;justify-content:space-between}
    @media print{body{padding:18px}}
  </style></head><body>
  <h1>KT M&amp;S 경영성과 보고서</h1>
  <div class="sub">기준월: ${fin.months[lastIdx]} &nbsp;|&nbsp; 생성일시: ${new Date().toLocaleString('ko-KR')} &nbsp;|&nbsp; 유형: ${type==='summary'?'경영진 요약':'상세 분석'}</div>
  <div class="section">
    <h2>📊 핵심 KPI 현황</h2>
    <table><thead><tr><th>지표</th><th>최근월 실적</th><th>기간 합계</th></tr></thead><tbody>
    ${kpiRows.map(r=>`<tr><td><b>${r[0]}</b></td><td>${r[1]}</td><td style="color:#666">${r[2]}</td></tr>`).join('')}
    </tbody></table>
  </div>
  <div class="section">
    <h2>💰 최근 재무 추이 (최근 6개월)</h2>
    <table><thead><tr><th>월</th><th>총매출</th><th>통신매출</th><th>영업이익</th><th>판관비</th></tr></thead><tbody>
    ${fin.months.slice(-6).map((m,i)=>{const ri=fin.months.length-6+i;return`<tr><td>${m}</td><td>${fmtN(fin.총매출[ri],1)}억</td><td>${fmtN(fin.통신매출[ri],1)}억</td><td style="color:${fin.영업이익[ri]>=0?'#22c55e':'#ef4444'}">${fmtN(fin.영업이익[ri],1)}억</td><td>${fmtN(fin.판관비[ri],1)}억</td></tr>`;}).join('')}
    </tbody></table>
  </div>
  <div class="footer"><span>KT M&amp;S 경영성과 대시보드 · 대외비</span><span>Page 1</span></div>
  </body></html>`);
  w.document.close();
  setTimeout(()=>w.print(),500);
  showToast('🖨️ PDF 인쇄 창이 열렸습니다');
}

// ════════════════════════════════════════════════
//  엑셀 업로드 → Supabase DB 저장 파이프라인
// ════════════════════════════════════════════════

let _uploadParsedData = null; // 파싱된 데이터 임시 저장
let _uploadFileName = '';

// 카테고리 → 시트명 매핑
const CAT_SHEET_MAP = {
  finance:  { sheet: 0, label: '재무' },
  wireless: { sheet: 0, label: '무선 가입자' },
  wired:    { sheet: 0, label: '유선 가입자' },
  org:      { sheet: 0, label: '조직별 실적' },
  digital:  { sheet: 0, label: '디지털 채널' },
  b2b:      { sheet: 0, label: 'B2B 채널' },
  smb:      { sheet: 0, label: '소상공인' },
  platform: { sheet: 0, label: '유통플랫폼' },
  tcsi:     { sheet: 0, label: 'TCSI' },
  voc:      { sheet: 0, label: '영업품질(VOC)' },
  hr:       { sheet: 0, label: '인력' },
  infra:    { sheet: 0, label: '인프라' },
  strategy: { sheet: 0, label: '전략상품' },
};

function handleUpload(inp){
  const f = inp.files[0];
  if(!f) return;

  // 분류 선택 확인
  const cat = document.getElementById('upload-category').value;
  if(!cat){
    showToast('⚠ 데이터 분류를 먼저 선택해주세요!');
    // 파일 선택 초기화
    inp.value = '';
    return;
  }

  _uploadFileName = f.name;

  // 업로드존 로딩 표시
  document.getElementById('upload-zone').innerHTML = `
    <div style="font-size:28px;margin-bottom:8px">⏳</div>
    <div style="font-size:13px;font-weight:600;color:var(--text2)">${f.name} 파싱 중...</div>
    <div style="font-size:11px;color:var(--text3);margin-top:4px">${(f.size/1024).toFixed(0)}KB</div>`;

  // FileReader로 엑셀 읽기
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      parseExcel(e.target.result, f.name, cat);
    } catch(err) {
      showUploadError('파일 파싱 실패: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(f);
}

function parseExcel(buffer, fileName, cat){
  // SheetJS로 파싱
  const wb = XLSX.read(buffer, { type: 'array' });

  // 첫 번째 시트 사용
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];

  // JSON으로 변환 (헤더 포함)
  const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  if(!rawRows || rawRows.length < 2){
    showUploadError('데이터가 없거나 형식이 맞지 않습니다');
    return;
  }

  // 월 컬럼 찾기 (헤더 행에서)
  // 구조: [대분류, 중분류, 소분류, 월1, 월2, ...]
  const headerRow = rawRows[0];
  const monthCols = [];
  const monthIdxs = [];

  headerRow.forEach((h, i) => {
    if(h && String(h).includes('월')){
      monthCols.push(String(h).replace(/'/g,'').trim());
      monthIdxs.push(i);
    }
  });

  if(monthCols.length === 0){
    showUploadError('월 데이터 컬럼을 찾을 수 없습니다. 파일 형식을 확인해주세요.');
    return;
  }

  // 데이터 행 파싱
  const dataRows = [];
  for(let i = 1; i < rawRows.length; i++){
    const row = rawRows[i];
    if(!row || row.every(v => v === null)) continue; // 빈 행 스킵

    const rowObj = {
      _idx: i,
      _labels: [],
      values: {}
    };

    // 레이블 컬럼 (월 이전 컬럼들)
    for(let j = 0; j < monthIdxs[0]; j++){
      if(row[j] !== null && row[j] !== undefined){
        rowObj._labels.push(String(row[j]).replace(/\n/g,' ').trim());
      }
    }

    // 월별 값
    monthIdxs.forEach((mIdx, mi) => {
      const val = row[mIdx];
      rowObj.values[monthCols[mi]] = (val !== null && val !== undefined && val !== '') 
        ? parseFloat(val) || val 
        : null;
    });

    dataRows.push(rowObj);
  }

  // DB 저장용 구조로 변환
  _uploadParsedData = {
    months: monthCols,
    rows: dataRows,
    sheetName: sheetName,
    totalRows: dataRows.length,
    category: cat,
  };

  // 미리보기 표시
  showUploadPreview(monthCols, dataRows, fileName, cat);
}

function showUploadPreview(months, rows, fileName, cat){
  // 업로드존 성공 표시
  document.getElementById('upload-zone').innerHTML = `
    <div style="font-size:28px;margin-bottom:8px">✅</div>
    <div style="font-size:13px;font-weight:600;color:var(--text2)">${fileName}</div>
    <div style="font-size:11px;color:var(--text3);margin-top:4px">
      ${rows.length}행 · ${months.length}개월 · 파싱 완료
    </div>`;

  // Step 2 표시
  const step2 = document.getElementById('upload-step2');
  step2.style.display = '';

  document.getElementById('upload-parse-status').textContent =
    `시트 첫번째 · ${rows.length}개 행 · ${months.length}개월 (${months[0]} ~ ${months[months.length-1]})`;

  // 미리보기 테이블 (최근 3개월, 최대 8행)
  const previewMonths = months.slice(-3);
  const previewRows = rows.slice(0, 8);

  document.getElementById('upload-preview').innerHTML = `
    <div style="font-size:11px;color:var(--text3);margin-bottom:8px">
      📋 미리보기 (최대 8행 · 최근 3개월) — 전체 ${rows.length}행 ${months.length}개월
    </div>
    <div class="table-scroll">
      <table class="data-table">
        <thead>
          <tr>
            <th>항목</th>
            ${previewMonths.map(m=>`<th class="num">${m}</th>`).join('')}
          </tr>
        </thead>
        <tbody>
          ${previewRows.map(row=>`
            <tr>
              <td style="font-size:11px;color:var(--text2)">${row._labels.join(' > ') || '–'}</td>
              ${previewMonths.map(m=>`
                <td class="num" style="font-size:11px">
                  ${row.values[m] !== null && row.values[m] !== undefined 
                    ? Number(row.values[m]).toLocaleString('ko-KR', {maximumFractionDigits:2})
                    : '–'}
                </td>`).join('')}
            </tr>`).join('')}
          ${rows.length > 8 ? `<tr><td colspan="${previewMonths.length+1}" style="text-align:center;color:var(--text3);font-size:11px">... 외 ${rows.length-8}행</td></tr>` : ''}
        </tbody>
      </table>
    </div>
    <div style="margin-top:12px;padding:10px 14px;background:rgba(74,158,255,.06);border:1px solid rgba(74,158,255,.2);border-radius:8px;font-size:11px;color:var(--text2)">
      💡 위 데이터가 맞으면 아래 <b>Supabase DB에 저장</b> 버튼을 눌러주세요.<br>
      기존 <b>${CAT_SHEET_MAP[cat]?.label || cat}</b> 데이터는 새 데이터로 교체됩니다.
    </div>`;
}

async function saveToSupabase(){
  if(!_uploadParsedData){
    showToast('⚠ 업로드할 데이터가 없습니다');
    return;
  }

  const cat = document.getElementById('upload-category').value;
  const period = document.getElementById('upload-period').value;
  const note = document.getElementById('upload-note').value;
  const btn = document.getElementById('upload-save-btn');

  btn.disabled = true;
  btn.textContent = '⏳ DB 저장 중...';

  try {
    // 1. kts_data에 upsert
    const { error: dataErr } = await sb.from('kts_data').upsert({
      category: cat,
      data: _uploadParsedData,
      updated_by: S.user?.name || '관리자',
      updated_at: new Date().toISOString()
    });

    if(dataErr) throw new Error(dataErr.message);

    // 2. 업로드 이력 기록
    await sb.from('kts_upload_log').insert({
      filename: _uploadFileName,
      category: cat,
      period: period || `${_uploadParsedData.months[0]} ~ ${_uploadParsedData.months[_uploadParsedData.months.length-1]}`,
      rows_count: _uploadParsedData.totalRows,
      uploaded_by: S.user?.name || '관리자',
      status: 'success',
      note: note || null
    });

    // 3. 메모리의 RAW 데이터도 업데이트 (즉시 반영)
    RAW[cat] = _uploadParsedData;

    // 4. 성공 표시
    document.getElementById('upload-result').innerHTML = `
      <div style="background:rgba(52,211,138,.07);border:1px solid rgba(52,211,138,.3);border-radius:10px;padding:16px 18px">
        <div style="font-size:14px;font-weight:700;color:var(--green);margin-bottom:6px">
          ✅ Supabase DB 저장 완료!
        </div>
        <div style="font-size:12px;color:var(--text2);line-height:1.7">
          📁 파일: <b>${_uploadFileName}</b><br>
          🗂 카테고리: <b>${CAT_SHEET_MAP[cat]?.label || cat}</b><br>
          📊 저장된 행: <b>${_uploadParsedData.totalRows}행 · ${_uploadParsedData.months.length}개월</b><br>
          🕐 저장 시각: <b>${new Date().toLocaleString('ko-KR')}</b>
        </div>
        <div style="margin-top:10px;font-size:11px;color:var(--text3)">
          임원이 대시보드를 새로고침하면 최신 데이터가 바로 반영됩니다 🎉
        </div>
      </div>`;

    showToast(`✅ "${_uploadFileName}" DB 저장 완료!`);
    btn.textContent = '✅ 저장 완료';

    // 이력 새로고침
    setTimeout(()=>renderAdminUpload(), 500);

  } catch(err) {
    document.getElementById('upload-result').innerHTML = `
      <div style="background:rgba(255,94,106,.07);border:1px solid rgba(255,94,106,.3);border-radius:10px;padding:14px 16px;color:var(--red);font-size:12px">
        ❌ 저장 실패: ${err.message}<br>
        <span style="color:var(--text3);font-size:11px;margin-top:4px;display:block">
          Supabase 연결을 확인하거나 다시 시도해주세요
        </span>
      </div>`;
    btn.disabled = false;
    btn.textContent = '☁️ Supabase DB에 저장';
    showToast('❌ 저장 실패: ' + err.message);
  }
}

function resetUpload(){
  _uploadParsedData = null;
  _uploadFileName = '';
  document.getElementById('upload-step2').style.display = 'none';
  document.getElementById('upload-result').innerHTML = '';
  document.getElementById('file-input').value = '';
  document.getElementById('upload-zone').innerHTML = `
    <div style="font-size:36px;margin-bottom:10px">📂</div>
    <div style="font-size:14px;font-weight:700;color:var(--text2)">클릭하거나 파일을 드래그하여 업로드</div>
    <div style="font-size:11px;color:var(--text3);margin-top:6px">xlsx, xls 지원 · 최대 50MB</div>
    <input type="file" id="file-input" accept=".xlsx,.xls" style="display:none" onchange="handleUpload(this)">`;
  // 이벤트 재등록
  const uz2 = document.getElementById('upload-zone');
  uz2.onclick = ()=>document.getElementById('file-input').click();
}

function showUploadError(msg){
  document.getElementById('upload-zone').innerHTML = `
    <div style="font-size:28px;margin-bottom:8px">❌</div>
    <div style="font-size:13px;font-weight:600;color:var(--red)">${msg}</div>
    <div style="font-size:11px;color:var(--text3);margin-top:6px">클릭하여 다시 선택</div>
    <input type="file" id="file-input" accept=".xlsx,.xls" style="display:none" onchange="handleUpload(this)">`;
  document.getElementById('upload-zone').onclick = ()=>document.getElementById('file-input').click();
  showToast('❌ ' + msg);
}
let tTimer;
function showToast(msg){
  const t=document.getElementById('toast');t.textContent=msg;t.classList.add('show');
  clearTimeout(tTimer);tTimer=setTimeout(()=>t.classList.remove('show'),3200);
}
const uz=document.getElementById('upload-zone');
if(uz){uz.addEventListener('dragover',e=>{e.preventDefault();uz.style.borderColor='var(--gold)';});uz.addEventListener('dragleave',()=>{uz.style.borderColor='';});uz.addEventListener('drop',e=>{e.preventDefault();uz.style.borderColor='';const f=e.dataTransfer.files[0];if(f){document.getElementById('file-input').files=e.dataTransfer.files;handleUpload(document.getElementById('file-input'));}});}
