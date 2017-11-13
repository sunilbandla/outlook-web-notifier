/* eslint-env browser, es6 */

'use strict';

const GRAPH_URL = "https://graph.microsoft.com/beta/subscriptions/";
let inboxSubscriptionRequest = {
  "changeType": "created",
  "notificationUrl": "https://outlookwebnotifier.azurewebsites.net/api/subscribe",
  "resource": "/me/mailfolders('inbox')/messages",
  "expirationDateTime": "",
  "clientState": ""
};
const GRAPH_SUBSCRIPTION_KEY = "graphSubscription";

function sendSubscriptionRequestToGraph() {
  var myHeaders = new Headers();

  var myInit = {
    method: 'POST',
    headers: myHeaders
  };

  fetch(GRAPH_URL, myInit)
    .then(function (response) {
      response.json()
        .then(function (subscription) {
          localStorage.setItem(GRAPH_SUBSCRIPTION_KEY, JSON.stringify(subscription));
          // TODO send it to server
        });
    })
    .catch(function (error) {
      console.log(error);
    });
}