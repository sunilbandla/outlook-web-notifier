/* eslint-env browser, serviceworker, es6 */

'use strict';
self.addEventListener('push', function(event) {
    console.log(`[Service Worker] Push received with this data: "${event.data.text()}"`);
  
    const title = 'Outlook web notification';
    const options = {
      body: event.data.text(),
      icon: 'images/icon.png',
      badge: 'images/badge.png'
    };
  
    event.waitUntil(self.registration.showNotification(title, options));
  });

  self.addEventListener('notificationclick', function(event) {
    event.notification.close();
  });