// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    TurnContext,
    MessageFactory,
    TeamsActivityHandler,
    CardFactory,
    ActionTypes
} = require('botbuilder');

class BotActivityHandler extends TeamsActivityHandler {
    constructor() {
        super();
        /* Conversation Bot */
        /*  Teams bots are Microsoft Bot Framework bots.
            If a bot receives a message activity, the turn handler sees that incoming activity
            and sends it to the onMessage activity handler.
            Learn more: https://aka.ms/teams-bot-basics.

            NOTE:   Ensure the bot endpoint that services incoming conversational bot queries is
                    registered with Bot Framework.
                    Learn more: https://aka.ms/teams-register-bot. 
        */
        // Registers an activity event handler for the message event, emitted for every incoming message activity.
        this.onMessage(async (context, next) => {
            TurnContext.removeRecipientMention(context.activity);
            switch (context.activity.text.trim()) {
            case 'Hello':
                await this.mentionActivityAsync(context);
                break;
            case 'Get Token':
                await this.getAppToken(context);
                break;
            case 'Get Shifts':
                await this.getShifts(context);
                break;
            default:
                // By default for unknown activity sent by user show
                // a card with the available actions.
                const value = { count: 0 };
                const card = CardFactory.heroCard(
                    'Lets talk...',
                    null,
                    [{
                        type: ActionTypes.MessageBack,
                        title: 'Say Hello',
                        value: value,
                        text: 'Hello'
                    }]);
                await context.sendActivity({ attachments: [card] });
                break;
            }
            await next();
        });
        /* Conversation Bot */
    }

    /* Conversation Bot */
    /**
     * Say hello and @ mention the current user.
     */
    async mentionActivityAsync(context) {
        const TextEncoder = require('html-entities').XmlEntities;

        const mention = {
            mentioned: context.activity.from,
            text: `<at>${ new TextEncoder().encode(context.activity.from.name) }</at>`,
            type: 'mention'
        };

        const replyActivity = MessageFactory.text(`Hi ${ mention.text }`);
        replyActivity.entities = [mention];
        
        await context.sendActivity(replyActivity);
    }

    // Get the app token to call the graph.
    async getAppToken(context) {
        // import client auth and get the app token for the bot
        var clientAuthToken = "";
        const getToken = require('./getToken');

        await getToken(context)
        .then ((accessToken) => {
            console.log(`\n Got access token of ${accessToken.length} characters`);
            clientAuthToken = accessToken;
        })  
        .catch((error) => {
            console.error(`ERROR getting token: ${error}`);
            resolve();
        });        
       
        await context.sendActivity(clientAuthToken);
    }

    async getShifts(context) {
        // import client auth and get the app token for the bot
        var clientAuthToken = "";
        const getToken = require('./getToken');

        await getToken(context)
        .then ((accessToken) => {
            console.log(`\n Got access token of ${accessToken.length} characters`);
            clientAuthToken = accessToken;
        })  
        .catch((error) => {
            console.error(`ERROR getting token: ${error}`);
            resolve();
        }); 

        // Read the retail team id and the actAs user from .env file.
        // We need the actAs user - which is any user in the tenant for the shifts call with Client credentials.
        const path = require('path');
        const ENV_FILE = path.join(__dirname, '.env');
        require('dotenv').config({ path: ENV_FILE });
        const retailTeamId = process.env.RetailTeamId;
        const actsAs = process.env.ActsAs;
        // We need utc dates for shifts graph call
        var d = new Date();
        var todaysDate = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
        var tomorrowDate = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()+1);
        var todayUTC = new Date(todaysDate).toISOString();
        var tomorrowUTC = new Date(tomorrowDate).toISOString();
        // now call the graph
        const response = await fetch("https://graph.microsoft.com/v1.0/teams/" + retailTeamId + "/schedule/shifts?$filter=sharedShift/startDateTime ge " + todayUTC + " and sharedShift/endDateTime le " + tomorrowUTC,
        {
            method: 'GET',
            headers: {
                "accept": "application/json",
                "authorization": "bearer " + clientAuthToken,
                "MS-APP-ACTS-AS": actsAs,
            }
        });

        if (response.ok) {
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


}

module.exports.BotActivityHandler = BotActivityHandler;

