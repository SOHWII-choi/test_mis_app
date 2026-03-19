// KT M&S 대시보드 Service Worker
const CACHE_NAME = 'kts-dashboard-v1';
const CACHE_FILES = [
  '/',
  '/index.html',
  'https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap',
  'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js',
  'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/dist/umd/supabase.js'
];

// 설치 — 핵심 파일 캐시
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      console.log('[SW] 캐시 설치 중...');
      return cache.addAll(CACHE_FILES).catch(err => {
        console.warn('[SW] 일부 캐시 실패 (무시):', err);
      });
    })
  );
  self.skipWaiting();
});

// 활성화 — 오래된 캐시 삭제
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => {
          console.log('[SW] 오래된 캐시 삭제:', k);
          return caches.delete(k);
        })
      )
    )
  );
  self.clients.claim();
});

// 네트워크 요청 — Cache First 전략
// 캐시에 있으면 캐시 사용, 없으면 네트워크
self.addEventListener('fetch', e => {
  // Supabase API 요청은 항상 네트워크 사용 (최신 데이터)
  if (e.request.url.includes('supabase.co') ||
      e.request.url.includes('anthropic.com')) {
    return; // 그냥 네트워크로
  }

  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(response => {
        // 성공한 GET 요청만 캐시에 저장
        if (e.request.method === 'GET' && response.status === 200) {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(e.request, copy));
        }
        return response;
      }).catch(() => {
        // 완전 오프라인 → 캐시된 index.html 반환
        if (e.request.destination === 'document') {
          return caches.match('/index.html');
        }
      });
    })
  );
});

// 업데이트 알림 (새 버전 배포 시)
self.addEventListener('message', e => {
  if (e.data === 'SKIP_WAITING') self.skipWaiting();
});
