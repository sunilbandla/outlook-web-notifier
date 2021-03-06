/* global sendSubscriptionRequestToGraph */
/* eslint-env browser, es7 */

'use strict';

let LOGIN_COUNT_KEY = 'loginCount';
let PROD_REDIRECT_URI = 'https://sunilbandla.github.io/outlook-web-notifier';
let redirectUri = location.origin;
if (location.hostname !== '127.0.0.1') {
  redirectUri = PROD_REDIRECT_URI;
}
let msalconfig = {
  clientID: 'fac25078-e844-4720-a161-a6331ffd8119',
  redirectUri
};

// Initialize application
let userAgentApplication = new Msal.UserAgentApplication(
  msalconfig.clientID,
  null,
  loginCallback,
  {
    redirectUri: msalconfig.redirectUri
  }
);

if (userAgentApplication.redirectUri) {
  userAgentApplication.redirectUri = msalconfig.redirectUri;
}

window.onload = function() {
  // If page is refreshed, continue to display user info
  if (
    !userAgentApplication.isCallback(window.location.hash) &&
    window.parent === window &&
    !window.opener
  ) {
    let user = userAgentApplication.getUser();
    if (user) {
      console.debug('onload');
      getToken();
    }
  }
};

/*
 * Call the Microsoft Graph API and display the results on the page. Sign the user in if necessary
 */
function signIn() {
  console.debug('signIn');
  let user = userAgentApplication.getUser();
  if (!user) {
    // If user is not signed in, then prompt user to sign in via loginRedirect.
    // This will redirect user to the Azure Active Directory v2 Endpoint
    trackLoginCount();
    return userAgentApplication
      .loginRedirect(graphAPIScopes);
    // The call to loginRedirect above frontloads the consent to query Graph API during the sign-in.
    // If you want to use dynamic consent, just remove the graphAPIScopes from loginRedirect call.
    // As such, user will be prompted to give consent when requested access to a resource that
    // he/she hasn't consented before. In the case of this application -
    // the first time the Graph API call to obtain user's profile is executed.
  } else {
    // In order to call the Graph API, an access token needs to be acquired.
    // Try to acquire the token used to query Graph API silently first:
    return getToken();
  }
}

function getToken() {
  console.debug('getToken');
  return userAgentApplication
    .acquireTokenSilent(graphAPIScopes)
    .then(loginSuccess, (error) => {
      console.debug('acquireTokenSilent failed', error);
      if (error) {
        let count = Number.parseInt(localStorage.getItem(LOGIN_COUNT_KEY) || 0);
        if (count === 0) {
          signIn();
        }
        return;
      }
    });
}

/**
 * Callback method from sign-in: if no errors, call loginSuccess() to show results.
 * @param {string} errorDesc - If error occur, the error message
 * @param {object} token - The token received from login
 * @param {object} error - The error string
 * @param {string} tokenType - The token type: For loginRedirect, tokenType = "id_token".
 *  For acquireTokenRedirect, tokenType:"access_token".
 */
function loginCallback(errorDesc, token, error, tokenType) {
  console.debug('loginCallback');
  if (errorDesc) {
    console.error('error: ' + errorDesc);
    showError(window.msal.authority, error, errorDesc);
  } else {
    loginSuccess(token);
  }
}

function loginSuccess(token) {
  console.debug('loginSuccess');
  resetLoginCount();
  hideError();
  showWelcomeMessage();
  return token;
}

function showWelcomeMessage() {
  let user = userAgentApplication.getUser();
  if (!user) {
    return;
  }
  let divWelcome = document.getElementById('WelcomeMessage');
  divWelcome.innerHTML = 'Hello ' + user.name;
}

function hideError() {
  let elem = document.getElementById('ErrorMessage');
  if (elem) {
    elem.classList.add('is-invisible');
  }
}

/**
 * Show an error message in the page
 * @param {string} endpoint - the endpoint used for the error message
 * @param {string} error - Error string
 * @param {string} errorDesc - Error description
 */
function showError(endpoint, error, errorDesc) {
  let formattedError = JSON.stringify(error, null, 4);
  if (formattedError.length < 3) {
    formattedError = error;
  }
  document.getElementById('ErrorMessage').classList.remove('is-invisible');
  document.getElementById('ErrorMessage').innerHTML =
    'An error has occurred:<br/>Endpoint: ' +
    endpoint +
    '<br/>Error: ' +
    formattedError +
    '<br/>' +
    errorDesc;
  console.error(error);
}

/**
 * Sign-out the user
 */
function signOut() {
  userAgentApplication.logout();
}

function trackLoginCount() {
  let count = localStorage.getItem(LOGIN_COUNT_KEY);
  if (count === null || count === undefined) {
    count = 0;
  }
  localStorage.setItem(LOGIN_COUNT_KEY, ++count);
}

function resetLoginCount() {
  localStorage.setItem(LOGIN_COUNT_KEY, 0);
}
