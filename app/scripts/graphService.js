/* global signIn, getToken, PUSH_SUBSCRIPTION_KEY */
/* eslint-env browser, es7 */

'use strict';

const GRAPH_URL = 'https://graph.microsoft.com/beta/subscriptions/';
const NOTIFIER_API_URL = 'https://outlookwebnotifier.azurewebsites.net/api/subscribe';
// Graph API scope used to obtain the access token to read user profile
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
const SUBSCRIPTION_EXP_DAYS = 3;

async function sendSubscriptionRequestToGraph() {
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
  data.clientState = Date.now();
  let expirationDateTime = new Date();
  expirationDateTime.setDate(expirationDateTime.getDate() + SUBSCRIPTION_EXP_DAYS);
  data.expirationDateTime = expirationDateTime.toUTCString();

  let request = {
    method: 'POST',
    headers: myHeaders,
    body: JSON.stringify(data)
  };

  fetch(GRAPH_URL, request)
    .then((response) => {
      if (response.ok) {
        return response.json();
      }
      throw new Error('Could not subscribe to notifications. Reload to try again.');
    })
    .then((subscription) => {
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
    })
    .catch((error) => {
      console.warn(error);
    });
}

function renewOrRecreateGraphSubscription() {
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

function removeGraphSubscription() {
  let graphSub = localStorage.getItem(GRAPH_SUBSCRIPTION_KEY);
  if (!graphSub) {
    return;
  }
  let request = {
    method: 'DELETE',
    headers: getDefaultHeaders()
  };

  fetch(`${NOTIFIER_API_URL}/${graphSub.id}`, request)
    .then((response) => {
      if (response.ok) {
        console.log('Subscriptions sent to server - response status =', response.status);
        return localStorage.removeItem(GRAPH_SUBSCRIPTION_KEY);
      }
      throw new Error('Could not remove Graph subscription from server');
    })
    .catch((error) => {
      console.warn(error);
    });
}

async function renewGraphSubscription() {
  let graphSub = localStorage.getItem(GRAPH_SUBSCRIPTION_KEY);
  graphSub = JSON.parse(graphSub);

  let token = await getAuthToken();
  if (!token) {
    return;
  }
  let myHeaders = getDefaultHeaders();
  myHeaders.append('Authorization', `Bearer ${token}`);

  let request = {
    method: 'PATCH',
    headers: myHeaders
  };
  let url = `${GRAPH_URL}/${graphSub.id}`;

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
    })
    .catch((error) => {
      console.warn(error);
    });
}

async function getAuthToken() {
  let token = await getToken();
  if (!token) {
    showErrorMessage('Authentication error. Reload and/or try again later.');
    // signIn();
    return;
  }
  return token;
}

function showErrorMessage(text) {
  let elem = document.getElementById('ErrorMessage');
  if (elem) {
    elem.style.visibility = 'visible';
    elem.innerText = text;
  }
}

function sendSubscriptionInfoToNotifierService(graphSub, pushSub) {
  let data = {};
  data.pushSubscription = pushSub;
  data.id = graphSub.id;
  data.clientState = graphSub.clientState;
  data.changeType = graphSub.changeType;
  data.resource = graphSub.resource;
  data.expirationDateTime = graphSub.expirationDateTime;

  let request = {
    method: 'POST',
    body: JSON.stringify(data),
    headers: getDefaultHeaders()
  };

  fetch(NOTIFIER_API_URL, request)
    .then((response) => {
      if (response.ok) {
        return console.log('Subscriptions sent to server - response status =', response.status);
      }
      throw new Error('Subscriptions could not be saved on server.');
    })
    .catch((error) => {
      console.warn(error);
    });
}

function getDefaultHeaders() {
  let myHeaders = new Headers();
  myHeaders.append('Content-Type', 'application/json');
  return myHeaders;
}