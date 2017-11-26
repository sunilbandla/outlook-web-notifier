/* global getToken, PUSH_SUBSCRIPTION_KEY */
/* eslint-env browser, es7 */

'use strict';

const GRAPH_ROOT_URL = 'https://graph.microsoft.com/beta';
const GRAPH_URL = `${GRAPH_ROOT_URL}/subscriptions/`;
const NOTIFIER_API_URL = 'https://outlookwebnotifier.azurewebsites.net/api/subscribe';
const graphAPIScopes = [
  'https://graph.microsoft.com/user.read',
  'https://graph.microsoft.com/mail.read'
];

let inboxSubscriptionRequest = {
  changeType: 'created',
  notificationUrl: 'https://outlookwebnotifier.azurewebsites.net/api/subscribe',
  resource: '/me/mailfolders(\'inbox\')/messages',
  expirationDateTime: '',
  clientState: ''
};
const GRAPH_SUBSCRIPTION_KEY = 'graphSubscription';
const SUBSCRIPTION_EXP_HOURS = 70;
let graphRenewalReminder;
let graphSubscriptionRequestCount = 1;

async function sendSubscriptionRequestToGraph() {
  console.log('sendSubscriptionRequestToGraph');
  let graphSub = localStorage.getItem(GRAPH_SUBSCRIPTION_KEY);
  if (graphSub) {
    renewOrRecreateGraphSubscription();
  }

  let myHeaders = getDefaultHeaders();
  let token = await getAuthToken();
  if (!token) {
    return;
  }
  myHeaders.append('Authorization', `Bearer ${token}`);

  let data = Object.assign({}, inboxSubscriptionRequest);
  data.clientState = Date.now().toString();
  data.expirationDateTime = getNewExpirationTime();

  let request = {
    method: 'POST',
    headers: myHeaders,
    body: JSON.stringify(data)
  };

  graphSubscriptionRequestCount++;
  fetch(GRAPH_URL, request)
    .then(async (response) => {
      if (response.ok) {
        return response.json();
      } else if (response.status === 400) {
        let json = await response.json();
        if (json.error.message === 'Subscription validation request timed out.' &&
          graphSubscriptionRequestCount <= 2) {
          sendSubscriptionRequestToGraph();
          return;
        }
      }
      throw new Error('Could not subscribe to notifications. Reload to try again.');
    })
    .then((subscription) => {
      if (!subscription) {
        return;
      }
      graphSubscriptionRequestCount = 0;
      console.log(subscription);
      localStorage.setItem(
        GRAPH_SUBSCRIPTION_KEY,
        JSON.stringify(subscription)
      );
      let pushSub = localStorage.getItem(PUSH_SUBSCRIPTION_KEY);
      if (!pushSub) {
        showErrorMessage('Push subscription expired. Clear cache and reload.');
        return;
      }
      pushSub = JSON.parse(pushSub);
      subscription[PUSH_SUBSCRIPTION_KEY] = pushSub;
      console.log(pushSub);
      sendSubscriptionInfoToNotifierService(subscription, pushSub);
      setupGraphSubReminder();
    })
    .catch((error) => {
      console.warn(error);
    });
}

function renewOrRecreateGraphSubscription() {
  console.log('renewOrRecreateGraphSubscription');
  let graphSub = localStorage.getItem(GRAPH_SUBSCRIPTION_KEY);
  if (!graphSub) {
    sendSubscriptionRequestToGraph();
    return;
  }
  graphSub = JSON.parse(graphSub);
  let subTime = new Date(graphSub.expirationDateTime);
  if (subTime > Date.now()) {
    renewGraphSubscription();
  } else {
    removeGraphSubscription();
    sendSubscriptionRequestToGraph();
  }
}

async function removeGraphSubscription() {
  console.log('removeGraphSubscription');
  let graphSub = localStorage.getItem(GRAPH_SUBSCRIPTION_KEY);
  if (!graphSub) {
    return;
  }
  graphSub = JSON.parse(graphSub);
  if (!graphSub || !graphSub.id) {
    return;
  }
  let token = await getAuthToken();
  if (!token) {
    return;
  }
  let myHeaders = getDefaultHeaders();
  myHeaders.append('Authorization', `Bearer ${token}`);
  let request = {
    method: 'DELETE',
    headers: myHeaders
  };

  fetch(`${NOTIFIER_API_URL}/${graphSub.id}`, request)
    .then((response) => {
      if (response.ok) {
        console.log('Subscriptions sent to notifier server - response status =', response.status);
        return localStorage.removeItem(GRAPH_SUBSCRIPTION_KEY);
      }
      throw new Error('Could not remove Graph subscription from notifier server');
    })
    .catch((error) => {
      console.warn(error);
    });
}

async function renewGraphSubscription() {
  console.log('renewGraphSubscription');
  let graphSub = localStorage.getItem(GRAPH_SUBSCRIPTION_KEY);
  graphSub = JSON.parse(graphSub);

  let token = await getAuthToken();
  if (!token) {
    return;
  }
  let myHeaders = getDefaultHeaders();
  myHeaders.append('Authorization', `Bearer ${token}`);

  let data = {};
  data.expirationDateTime = getNewExpirationTime();

  let request = {
    method: 'PATCH',
    headers: myHeaders,
    body: JSON.stringify(data)
  };
  let url = `${GRAPH_URL}${graphSub.id}`;

  fetch(url, request)
    .then((response) => {
      if (response.ok) {
        return response.json();
      }
      throw new Error('Could not renew Graph subscription server');
    })
    .then((subscription) => {
      console.log(subscription);
      graphSub.expirationDateTime = subscription.expirationDateTime;
      localStorage.setItem(
        GRAPH_SUBSCRIPTION_KEY,
        JSON.stringify(graphSub)
      );
      console.log(graphSub);
      sendSubscriptionInfoToNotifierService(graphSub, null);
      setupGraphSubReminder();
    })
    .catch((error) => {
      console.warn(error);
    });
}

async function getAuthToken() {
  console.log('getAuthToken');
  let token = await getToken();
  if (!token) {
    showErrorMessage('Authentication error. Reload and/or try again later.');
    return;
  }
  return token;
}

function showErrorMessage(text) {
  let elem = document.getElementById('ErrorMessage');
  if (elem) {
    elem.classList.remove('is-invisible');
    elem.innerText = text;
  }
}

async function sendSubscriptionInfoToNotifierService(graphSub, pushSub) {
  console.log('sendSubscriptionInfoToNotifierService');
  if (!pushSub) {
    pushSub = localStorage.getItem(PUSH_SUBSCRIPTION_KEY);
    pushSub = JSON.parse(pushSub);
  }
  let data = {};
  data.pushSubscription = pushSub;
  data.id = graphSub.id;
  data.clientState = graphSub.clientState;
  data.changeType = graphSub.changeType;
  data.resource = graphSub.resource;
  data.expirationDateTime = graphSub.expirationDateTime;

  let token = await getAuthToken();
  if (!token) {
    return;
  }
  let myHeaders = getDefaultHeaders();
  myHeaders.append('Authorization', `Bearer ${token}`);

  let request = {
    method: 'POST',
    body: JSON.stringify(data),
    headers: myHeaders
  };

  fetch(NOTIFIER_API_URL, request)
    .then((response) => {
      if (response.ok) {
        subscriptionComplete('SubscriptionMessage');
        return console.log('Subscriptions sent to server - response status =', response.status);
      }
      throw new Error('Subscriptions could not be saved on server.');
    })
    .catch((error) => {
      console.warn(error);
    });
}

async function getMailInfo(id) {
  console.log('getMailInfo');
  let token = await getAuthToken();
  if (!token) {
    return;
  }
  let myHeaders = getDefaultHeaders();
  myHeaders.append('Authorization', `Bearer ${token}`);

  let request = {
    method: 'GET',
    headers: myHeaders
  };
  let url = `${GRAPH_ROOT_URL}/me/messages/${id}`;

  return fetch(url, request)
    .then((response) => {
      if (response.ok) {
        return response.json();
      }
      throw new Error('Could not fetch mail info from Graph.');
    })
    .then((mail) => {
      return mail;
    })
    .catch((error) => {
      console.warn(error);
    });
}

function setupGraphSubReminder() {
  console.log('setupGraphSubReminder');
  if (graphRenewalReminder) {
    clearTimeout(graphRenewalReminder);
  }
  let delay = 2 * 24 * 60 * 60 * 1000;
  graphRenewalReminder = setTimeout(() => {
    renewGraphSubscription();
  }, delay);
}

function getDefaultHeaders() {
  let myHeaders = new Headers();
  myHeaders.append('Content-Type', 'application/json');
  return myHeaders;
}

function getNewExpirationTime() {
  let expirationDateTime = new Date();
  expirationDateTime.setHours(expirationDateTime.getHours() + SUBSCRIPTION_EXP_HOURS);
  expirationDateTime = expirationDateTime.toISOString();
  return expirationDateTime;
}

function subscriptionComplete(id) {
  document.getElementById(id).classList.toggle('is-invisible');
  document.getElementById('Spinner').classList.toggle('is-invisible');
}
