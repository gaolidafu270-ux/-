// 通知を受け取るサービスワーカー(アプリを閉じていても動く)
importScripts('https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.12.2/firebase-messaging-compat.js');
importScripts('config.js');

firebase.initializeApp(self.APP_CONFIG.firebase);
const messaging = firebase.messaging();

// 画面を見ていない(前面にない)ときに届いた通知を表示する。見ているときは、アプリ側が知らせるので出さない
messaging.onBackgroundMessage(function(payload) {
  const d = (payload && payload.data) || {};
  return self.registration.showNotification(d.title || self.APP_CONFIG.appName || 'お知らせ', {
    body: d.body || '',
    tag: d.tag || 'chat',
    renotify: true,
    icon: 'icon-192.png',
    badge: 'icon-192.png',
    data: { url: './' }
  });
});

// 通知をタップしたら、アプリを前に出す(閉じていれば開く)
self.addEventListener('notificationclick', function(event) {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function(list) {
      for (const c of list) { if ('focus' in c) return c.focus(); }
      return clients.openWindow((event.notification.data && event.notification.data.url) || './');
    })
  );
});
