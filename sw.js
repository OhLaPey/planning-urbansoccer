var CACHE_KEY = 'planning-urban7d-latest';
var CHECK_INTERVAL = 30 * 60 * 1000; // 30 min

self.addEventListener('install', function(e) { self.skipWaiting(); });
self.addEventListener('activate', function(e) { e.waitUntil(self.clients.claim()); });

// Messages from page
self.addEventListener('message', function(e) {
  if (e.data && e.data.type === 'CHECK_UPDATES') {
    checkForUpdates();
  }
  if (e.data && e.data.type === 'INIT') {
    // Store current state without notifying (first load)
    self._lastKnown = e.data.data;
  }
});

function checkForUpdates() {
  fetch('latest.json?_t=' + Date.now())
    .then(function(r) { return r.json(); })
    .then(function(data) {
      // Compare with stored version
      var stored = null;
      try { stored = JSON.parse(self._lastKnown || 'null'); } catch(e) {}

      if (stored && data.generatedAt !== stored.generatedAt) {
        var newWeeks = data.weeks.filter(function(w) {
          return stored.weeks.indexOf(w) === -1;
        });
        if (newWeeks.length > 0) {
          showNotification(
            'Nouveau planning disponible !',
            'Semaine ' + newWeeks.join(', S') + ' ajoutée sur Planning Urban 7D'
          );
        } else {
          showNotification(
            'Planning mis à jour',
            'Le planning S' + data.latestWeek + ' a été mis à jour'
          );
        }
      }
      self._lastKnown = JSON.stringify(data);
    })
    .catch(function() {});
}

function showNotification(title, body) {
  self.registration.showNotification(title, {
    body: body,
    icon: 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text y=".9em" font-size="90">⚽</text></svg>',
    badge: 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text y=".9em" font-size="90">⚽</text></svg>',
    tag: 'planning-update',
    renotify: true,
    data: { url: './' }
  });
}

self.addEventListener('notificationclick', function(e) {
  e.notification.close();
  e.waitUntil(
    self.clients.matchAll({ type: 'window' }).then(function(clients) {
      for (var i = 0; i < clients.length; i++) {
        if (clients[i].url.indexOf('planning') !== -1 && 'focus' in clients[i]) {
          return clients[i].focus();
        }
      }
      if (self.clients.openWindow) {
        return self.clients.openWindow(e.notification.data.url || './');
      }
    })
  );
});
