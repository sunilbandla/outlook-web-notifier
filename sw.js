/* eslint-env browser, serviceworker, es8 */

'use strict';
self.addEventListener('push', async (event) => {
  console.debug(
    `[Service Worker] Push received with this data: "${event.data.text()}"`
  );

  if (event.data) {
    let data = event.data.json();
    if (data['@odata.type'].indexOf('Message') !== -1) {
      let mailId = data.id;
      const swListener = new BroadcastChannel('swListener');
      swListener.postMessage({
        method: 'getMailInfo',
        id: mailId
      });
    }
  }
});

self.addEventListener('notificationclick', (event) => {
  event.notification.close();
});

self.addEventListener('message', async (event) => {
  console.debug('SW received message: ', event.data);

  let message = 'You\'ve got mail.';
  let subject = 'Outlook web notification';

  if (event.data.method.indexOf('getMailInfo') !== -1) {
    let mailInfo = event.data.mailInfo;
    if (mailInfo) {
      subject = mailInfo.subject;
      message = mailInfo.bodyPreview;
    }
  }
  const title = subject;
  const options = {
    body: message,
    icon: 'images/icon.png',
    badge: 'images/badge.png'
  };

  event.waitUntil(self.registration.showNotification(title, options));
});
