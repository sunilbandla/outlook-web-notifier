/* eslint-env browser, es6 */

'use strict';

let ADAL = new AuthenticationContext({
  instance: 'https://login.microsoftonline.com/',
  tenant: 'common', // COMMON OR YOUR TENANT ID

  clientId: 'e8c8e28c-ab9f-4c7a-904a-e64443f372b5',
  redirectUri: 'http://127.0.0.1:8887/',

  callback: userSignedIn,
  popUp: true
});

function signIn() {
  ADAL.login();
}

function userSignedIn(err, token) {
  console.log('userSignedIn called');
  if (!err) {
    console.log('token: ' + token);
    showWelcomeMessage();
  }
  else {
    console.error('error: ' + err);
  }
}

function showWelcomeMessage() {
  let user = ADAL.getCachedUser();
  let divWelcome = document.getElementById('WelcomeMessage');
  divWelcome.innerHTML = 'Welcome ' + user.profile.name;
}
