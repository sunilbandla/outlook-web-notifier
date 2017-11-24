/* global signIn, renewOrRecreateGraphSubscription, GRAPH_SUBSCRIPTION_KEY,
  removeGraphSubscription, userAgentApplication */
/* eslint-env browser, es7 */

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

if ('serviceWorker' in navigator && 'PushManager' in window) {
  console.log('Service Worker and Push is supported');

  navigator.serviceWorker.register('sw.js')
    .then(function(swReg) {
      console.log('Service Worker is registered', swReg);

      swRegistration = swReg;
      initializeUI();
    })
    .catch(function(error) {
      console.error('Service Worker Error', error);
    });
} else {
  console.warn('Push messaging is not supported');
  pushButton.textContent = 'Push Not Supported';
}

function initializeUI() {
  pushButton.addEventListener('click', function() {
    pushButton.disabled = true;
    if (isSubscribed) {
      unsubscribeUser();
    } else {
      subscribeUser();
    }
  });

  // Set the initial subscription value
  swRegistration.pushManager.getSubscription()
    .then(subscriptionSuccess)
    .catch(() => {
      console.log('User is NOT subscribed.');
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
  const applicationServerKey = urlB64ToUint8Array(applicationServerPublicKey);
  swRegistration.pushManager.subscribe({
      userVisibleOnly: true,
      applicationServerKey: applicationServerKey
    })
    .then(subscriptionSuccess, (err) => {
      console.log('Failed to subscribe the user: ', err);
      updateBtn();
    })
    .catch(function(err) {
      console.log('Failed to subscribe the user: ', err);
      updateBtn();
    });
}

function subscriptionSuccess(subscription) {
  isSubscribed = !(subscription === null);
  // TODO check if subscription is the same for already subscribed user
  // Currently assuming it is same until unsubscribed
  updatePushSubscriptionInStorage(subscription);

  if (isSubscribed) {
    console.log('User IS subscribed.');
    signIn();
  } else {
    console.log('User is NOT subscribed.');
    subscribeUser();
  }

  updateBtn();
}

function updatePushSubscriptionInStorage(subscription) {
  const subscriptionJson = document.querySelector('.js-subscription-json');
  const subscriptionDetails =
    document.querySelector('.js-subscription-details');

  if (subscription) {
    subscriptionJson.textContent = JSON.stringify(subscription);
    subscriptionDetails.classList.remove('is-invisible');
    localStorage.setItem(PUSH_SUBSCRIPTION_KEY, JSON.stringify(subscription));
    renewOrRecreateGraphSubscription();
  } else {
    localStorage.removeItem(PUSH_SUBSCRIPTION_KEY);
    subscriptionDetails.classList.add('is-invisible');
  }
}

function restart() {
  localStorage.removeItem(PUSH_SUBSCRIPTION_KEY);
  localStorage.removeItem(GRAPH_SUBSCRIPTION_KEY);
  removeGraphSubscription();
}

function unsubscribeUser() {
  swRegistration.pushManager.getSubscription()
    .then(function(subscription) {
      if (subscription) {
        return subscription.unsubscribe();
      }
    })
    .catch(function(error) {
      console.log('Error unsubscribing', error);
    })
    .then(function() {
      updatePushSubscriptionInStorage(null);
      removeGraphSubscription();

      console.log('User is unsubscribed.');
      isSubscribed = false;

      updateBtn();
      // TODO
      // signOut();
    });
}