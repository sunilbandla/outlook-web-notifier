/* global sendSubscriptionRequestToGraph */
/* eslint-env browser, es6 */

'use strict';

let msalconfig = {
  clientID: 'fac25078-e844-4720-a161-a6331ffd8119',
  redirectUri: location.origin
};

// Graph API scope used to obtain the access token to read user profile
let graphAPIScopes = ['https://graph.microsoft.com/user.read'];

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
      getToken();
    }
  }
};

/**
 * Call the Microsoft Graph API and display the results on the page. Sign the user in if necessary
 */
function signIn() {
  let user = userAgentApplication.getUser();
  if (!user) {
    // If user is not signed in, then prompt user to sign in via loginRedirect.
    // This will redirect user to the Azure Active Directory v2 Endpoint
    userAgentApplication
      .loginPopup(graphAPIScopes)
      .then(loginSuccess, function(error) {
        if (error) {
          showError(window.msal.authority, error);
        }
      });
    // The call to loginRedirect above frontloads the consent to query Graph API during the sign-in.
    // If you want to use dynamic consent, just remove the graphAPIScopes from loginRedirect call.
    // As such, user will be prompted to give consent when requested access to a resource that
    // he/she hasn't consented before. In the case of this application -
    // the first time the Graph API call to obtain user's profile is executed.
  } else {
    // TODO Show Sign-Out button
    // In order to call the Graph API, an access token needs to be acquired.
    // Try to acquire the token used to query Graph API silently first:
    getToken();
  }
}

function getToken() {
  userAgentApplication
    .acquireTokenSilent(graphAPIScopes)
    .then(loginSuccess, function(error) {
      if (error) {
        userAgentApplication.acquireTokenRedirect(graphAPIScopes);
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
  if (errorDesc) {
    console.error('error: ' + errorDesc);
    showError(window.msal.authority, error, errorDesc);
  } else {
    loginSuccess(token);
  }
}

function loginSuccess(token) {
  console.log('token: ' + token);
  hideError();
  showWelcomeMessage();
  // TODO
  // sendSubscriptionRequestToGraph(token);
}

function showWelcomeMessage() {
  let user = userAgentApplication.getUser();
  if (!user) {
    return;
  }
  let divWelcome = document.getElementById('WelcomeMessage');
  divWelcome.innerHTML = 'Welcome ' + user.name;
}

function hideError() {
  let elem = document.getElementById('ErrorMessage');
  if (elem) {
    elem.style.visibility = 'hidden';
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
