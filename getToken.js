
var adal = require('adal-node');
const path = require('path');

// Read botFilePath and botFileSecret from .env file.
const ENV_FILE = path.join(__dirname, '.env');
require('dotenv').config({ path: ENV_FILE });

var botId = process.env.BotId;
var botPassword = process.env.BotPassword;
var tenantId = process.env.TenantId;

module.exports = function getToken(context) {

  return new Promise((resolve, reject) => {

    // Get the ADAL client
    const authContext = new adal.AuthenticationContext("https://login.microsoftonline.com/" + tenantId);
    authContext.acquireTokenWithClientCredentials(
      "https://graph.microsoft.com/", botId, botPassword,
      (err, tokenRes) => {
        if (err) {
          reject(err);
        }
        else {
          resolve(tokenRes.accessToken);
        }
      });
  });
}