import * as Msal from 'msal';

export const applicationConfig = {
  clientID: '<your_app_client_id>',
  graphScopes: ['user.read', 'files.read.all']
};

function loggerCallback(level: Msal.LogLevel, message: string, containsPii: boolean) {
  console.log(message);
}

const logger = new Msal.Logger(
  loggerCallback,
  { level: Msal.LogLevel.Verbose, correlationId: '11657' });
// level and correlationId are optional parameters.

// Logger has other optional parameters like piiLoggingEnabled which can be assigned as shown aabove.
// Please refer to the docs to see the full list and their default values.

export function getUserAgentApplication(authCallback: any): Msal.UserAgentApplication {
  const userAgentApplication = new Msal.UserAgentApplication(
    applicationConfig.clientID,
    '',
    authCallback,
    { logger: logger, cacheLocation: 'localStorage' }); // logger and cacheLocation are optional parameters.
  // userAgentApplication has other optional parameters like redirectUri which can be assigned as shown above.
  // Please refer to the docs to see the full list and their default values.
  return userAgentApplication;
}