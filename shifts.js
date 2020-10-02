
const path = require('path');

// Read the retail team id and the actAs user from .env file.
// We need the actAs user - which is any user in the tenant for the shifts call with Client credentials.
const ENV_FILE = path.join(__dirname, '.env');
require('dotenv').config({ path: ENV_FILE });

const retailTeamId = process.env.RetailTeamId;
const actsAs = process.env.ActsAs;

// We need UTC dates for the shifts API
var d = new Date();
var todaysDate = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
var todayUTC = new Date(todaysDate).toISOString();
var tomorrowsDate = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()+1);
var tomorrowsUTC = new Date(tomorrowsDate).toISOString();

module.exports = function getShifts(context) {

        const response = await fetch("https://graph.microsoft.com/v1.0/teams/" + retailTeamId + "/schedule/shifts?$filter=sharedShift/startDateTime ge " + todayUTC + " and sharedShift/endDateTime le " + tomorrowsUTC,
        {
            method: 'GET',
            headers: {
                "accept": "application/json",
                "authorization": "bearer " + clientAuthToken,
                "MS-APP-ACTS-AS": actsAs,
            }
        });

        if (response.ok) {
            return response.json();
            const shifts = await response.json();
            let botOutput = `There are ${shifts.value.length} shifts today:`;
            for (const s of shifts.value) {
                botOutput += `<br />${s.id} - ${s.userId} `;
            }
            await context.sendActivity(botOutput);
        } else {
            await context.sendActivity(`Error ${response.status}: ${response.statusText}`);
        }
} 