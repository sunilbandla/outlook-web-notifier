/* global signIn, resetLoginCount, renewOrRecreateGraphSubscription, GRAPH_SUBSCRIPTION_KEY,
  removeGraphSubscription, userAgentApplication, getMailInfo, hideError */
/* eslint-env browser, es8 */

'use strict';

const applicationServerPublicKey = 'BHo2OFe7AOLXvMTfGdXqfuBKkA50qHUrsycyzuqGhaM5l3HUeT2n_1hugnJyWr6dWQEE7jT40eV16WgYf8-omRE';
const PUSH_SUBSCRIPTION_KEY = 'pushSubscription';
const pushButton = document.querySelector('.js-push-btn');

let isSubscribed = false;
let swRegistration = null;

function urlB64ToUint8Array(base64String) {
  const padding = '='.repeat((4 - base64String.length % 4) % 4);
  const base64 = (base64String + padding)
    .replace(/\-/g, '+')
    .replace(/_/g, '/');

  const rawData = window.atob(base64);
  const outputArray = new Uint8Array(rawData.length);

  for (let i = 0; i < rawData.length; ++i) {
    outputArray[i] = rawData.charCodeAt(i);
  }
  return outputArray;
}

(function() {
if ('serviceWorker' in navigator && 'PushManager' in window) {
  if (
    window.parent !== window &&
    window.opener
  ) {
    console.debug('In auth frame');
    return;
  }
  console.debug('Service Worker and Push is supported');

  navigator.serviceWorker.register('sw.js')
    .then(function(swReg) {
      console.debug('Service Worker is registered', swReg);

      swRegistration = swReg;
      initializeUI();
      registerMessageHandler();
    })
    .catch(function(error) {
      console.error('Service Worker registration error.', error);
    });
} else {
  console.warn('Push messaging is not supported');
  pushButton.textContent = 'Push Not Supported';
}
})();

function enableNotificationsClickHandler() {
  pushButton.disabled = true;
  if (isSubscribed) {
    unsubscribeUser();
  } else {
    subscribeUser();
  }
}

function registerMessageHandler() {
  const swListener = new BroadcastChannel('swListener');
  swListener.onmessage = async (event) => {
    console.debug('Main thread received message', event, event.data);
    if (event && event.data && event.data.method === 'getMailInfo') {
      let mailInfo = await getMailInfo(event.data.id);
      let data = Object.assign({}, event.data);
      data.mailInfo = mailInfo;
      sendMessageToServiceWorker(data);
    }
  };
}

function sendMessageToServiceWorker(msg) {
  console.debug('sendMessageToServiceWorker', msg);
  navigator.serviceWorker.controller.postMessage(msg);
}

function initializeUI() {
  pushButton.addEventListener('click', enableNotificationsClickHandler);

  // Set the initial subscription value
  swRegistration.pushManager.getSubscription()
    .then(subscriptionSuccess)
    .catch(() => {
      console.debug('Could not obtain push subscription. Reload or open app in a new tab.');
      subscribeUser();
    });
}

function updateBtn() {
  if (Notification.permission === 'denied') {
    pushButton.textContent = 'Push Messaging Blocked.';
    pushButton.disabled = true;
    updatePushSubscriptionInStorage(null);
    return;
  }

  if (isSubscribed) {
    pushButton.textContent = 'Disable notifications';
  } else {
    pushButton.textContent = 'Enable notifications';
  }

  pushButton.disabled = false;
}

function subscribeUser() {
  console.debug('subscribeUser');
  const applicationServerKey = urlB64ToUint8Array(applicationServerPublicKey);
  document.getElementById('Spinner').classList.remove('is-invisible');
  swRegistration.pushManager.subscribe({
      userVisibleOnly: true,
      applicationServerKey: applicationServerKey
    })
    .then((subscription) => {
      signIn();
      subscriptionSuccess(subscription);
    }, (err) => {
      console.debug('Failed to subscribe the user: ', err);
      updateBtn();
    })
    .catch(function(err) {
      console.debug('Failed to subscribe the user: ', err);
      updateBtn();
    });
}

function subscriptionSuccess(subscription) {
  console.debug('subscriptionSuccess', subscription);

  isSubscribed = !(subscription === null);

  if (isSubscribed) {
    console.debug('User IS subscribed.');
  } else {
    console.debug('User is NOT subscribed.');
  }

  updateBtn();
  // TODO check if subscription is the same for already subscribed user
  // Currently assuming it is same until unsubscribed
  updatePushSubscriptionInStorage(subscription);
}

function updatePushSubscriptionInStorage(subscription) {
  console.debug('updatePushSubscriptionInStorage');

  if (subscription) {
    localStorage.setItem(PUSH_SUBSCRIPTION_KEY, JSON.stringify(subscription));
    renewOrRecreateGraphSubscription();
  } else {
    localStorage.removeItem(PUSH_SUBSCRIPTION_KEY);
  }
}

function unsubscribeUser() {
  document.getElementById('Spinner').classList.remove('is-invisible');
  hideError();
  swRegistration.pushManager.getSubscription()
    .then(function(subscription) {
      if (subscription) {
        return subscription.unsubscribe();
      }
    })
    .catch(function(error) {
      console.debug('Error unsubscribing', error);
    })
    .then(function() {
      updatePushSubscriptionInStorage(null);
      removeGraphSubscription();

      console.debug('User is unsubscribed.');
      isSubscribed = false;

      updateBtn();
      resetLoginCount();
      document.getElementById('SubscriptionMessage').classList.add('is-invisible');
      document.getElementById('Spinner').classList.add('is-invisible');
      // TODO
      // signOut();
    });
}
