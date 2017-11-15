/* global PUSH_SUBSCRIPTION_KEY */
/* eslint-env browser, es6 */

'use strict';

const GRAPH_URL = 'https://graph.microsoft.com/beta/subscriptions/';
let inboxSubscriptionRequest = {
  changeType: 'created',
  notificationUrl: 'https://outlookwebnotifier.azurewebsites.net/api/subscribe',
  resource: '/me/mailfolders(\'inbox\')/messages',
  expirationDateTime: '',
  clientState: ''
};
const GRAPH_SUBSCRIPTION_KEY = 'graphSubscription';
const SUBSCRIPTION_EXP_DAYS = 1;

function sendSubscriptionRequestToGraph(token) {
  let graphSub = localStorage.getItem(GRAPH_SUBSCRIPTION_KEY);
  if (graphSub) {
    renewOrRecreateGraphSubscription(token);
  }

  let myHeaders = new Headers();
  myHeaders.append('Authorization', `Bearer ${token}`);

  let data = Object.assign({}, inboxSubscriptionRequest);
  data.clientState = Date.now();
  let expirationDateTime = new Date();
  expirationDateTime.setDate(expirationDateTime.getDate() + SUBSCRIPTION_EXP_DAYS);
  data.expirationDateTime = expirationDateTime.toUTCString();

  let myInit = {
    method: 'POST',
    headers: myHeaders,
    body: JSON.stringify(data)
  };

  fetch(GRAPH_URL, myInit)
    .then((response) => {
      return response.json();
    })
    .then((subscription) => {
      console.log(subscription);
      localStorage.setItem(
        GRAPH_SUBSCRIPTION_KEY,
        JSON.stringify(subscription)
      );
      let pushSub = localStorage.getItem(PUSH_SUBSCRIPTION_KEY);
      if (!pushSub) {
        // TODO handle fatal state
      }
      pushSub = JSON.parse(pushSub);
      subscription[PUSH_SUBSCRIPTION_KEY] = pushSub;
      console.log(pushSub);
      // TODO send sub to notifier server
    })
    .catch((error) => {
      console.warn(error);
    });
}

function renewOrRecreateGraphSubscription() {
  let graphSub = localStorage.getItem(GRAPH_SUBSCRIPTION_KEY);
  graphSub = JSON.parse(graphSub);
  let subTime = new Date(graphSub.expirationDateTime);
  // TODO get auth token
  if (subTime > Date.now()) {
    renewSubscription();
  } else {
    removeSubscription();
    sendSubscriptionRequestToGraph();
  }
}

function removeSubscription() {
  localStorage.removeItem(GRAPH_SUBSCRIPTION_KEY);
  // TODO remove existing graph subscription from notifier server
}

function renewSubscription(token) {
  let graphSub = localStorage.getItem(GRAPH_SUBSCRIPTION_KEY);
  graphSub = JSON.parse(graphSub);

  let myHeaders = new Headers();
  myHeaders.append('Authorization', `Bearer ${token}`);

  let myInit = {
    method: 'PATCH',
    headers: myHeaders
  };
  let url = `${GRAPH_URL}/${graphSub.id}`;

  fetch(url, myInit)
    .then((response) => {
      return response.json();
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
