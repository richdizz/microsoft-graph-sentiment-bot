**I am fully aware I left sensitive information in the .env file. This was done on purpose so the project could be easly cloned and immediately run. There is nothing in there you can do anything melicous with...I think :)**

**If the bot does not hit the Search intent as you expect, my quota might be up for the month on the LUIS subscription key...you can replace it with your own subscription key as the LUIS model is published publically. OR, you can publish your own LUIS model using the LUIS model export in the root of the project**

# Microsoft Graph Sentiment Analysis Bot
This repository contains a hands on lab (HOL) for building a Bot Framework bot that connects to the [Microsoft Graph](https://developer.microsoft.com/en-us/graph/) for mail messages and performs sentiment analysis on those messages using the [Microsoft Text Analytics cognitive service](https://docs.microsoft.com/en-us/azure/cognitive-services/text-analytics/quick-start). The completed solution is available in this repo in case you have issues with the HOL.

**This HOL uses a quick start template not designed for production. It exists primarily to get you started quickly with learning and prototyping bots that connect to the Microsoft Graph**

Please keep the above in mind before posting issues and PRs on this repo.

## Prerequisites

The following is required to complete this module:

- Visual Studio Code
- NodeJS
- [Bot Framework Emulator](https://docs.microsoft.com/en-us/bot-framework/debug-bots-emulator)
- Office 365 or Outlook.com account
- Microsoft Cognitive Services Key for the Text Analytics Service (explained in Task 0)

## Task 0 - Signing up for the Text Analytics API
This HOL will leverage the Text Analytics service to perform sentiment analysis on email messages. This requires a key that can be acquired using the steps below.
1. Navigate to Cognitive Services in the Azure portal and ensure Text Analytics is selected as the 'API type'.
2. Select a plan. You may select the free tier for 5,000 transactions/month. As is a free tier, there are no charges for using the service. You need to login to your Azure subscription to sign up for the service.
3. Complete the other fields and create your account.
4. After you sign up for Text Analytics, find your API Key. Copy the primary key, as you need it when using the API services.

## Task 1 - Clone the quick start template from GitHub
In this task, you will clone a quick start repository into your own project and test it in the [Bot Framework emulator](https://docs.microsoft.com/en-us/bot-framework/debug-bots-emulator).
1. You will start by cloning the Quickstart template into a new project folder:

```
git clone https://github.com/richdizz/microsoft-graph-bot-quickstart-luis-node
SearchBot
```

2. Next, discard the git references:

```
rm -rf .git  # OS/X (bash)
rd .git /S/Q # windows
```

3. Install packages and start the application.

```
npm install
npm start
```

4. Launch the Bot Framework Emulator (available [HERE](https://docs.botframework.com/en-us/tools/bot-framework-emulator/)).

5. Enter the messaging endpoint for the bot project (likely http://localhost:3980/api/messages but could be a different port) and the click the "Connect" button (leave the Microsoft App ID and Microsoft App Password blank as these are for published bots).

6. Type anything to the bot and follow the prompts and validate the bot functions.

## Task 2 - Update the environment variables
In this task, you will update the environment variables for the bot with your text analytics subscription key and new LUIS model endpoint.

1. Open the .env file in the root of the repository
2. Update the **LUIS_MODEL_URL** property with the value below. Alternatively, you could take the **SearchBotLUIS.json** file in the project root and publish to your own LUIS model at [https://www.luis.ai/](https://www.luis.ai/):

```
https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/640d431b-70bf-4abc-8a19-4901c04289c0?subscription-key=25670274c6ad4180a3e4be0e8f4d91d5
```
3. Add a new environment variable to the .env file with a key of **COG_SUB_KEY** and the value of the key generated in Task 0. Example:

```
COG_SUB_KEY=9416c8887b8d4c69af33ffc00bfd6655
```

## Task 3 - Code the new solution
In this task, you will make make modifications to existing code and add new code to make the solution search mail messages and perform sentiment analysis on them.
1. Locate the **searchFilesDialog.ts** file (located under **src/server/dialogs**) and rename it to a more generic **searchDialog.ts** (feel free to delete the searchFilesDialog.js and searchFilesDialog.js.map as the TypeScript compiler will generate new files with the correct name).
2. Open the **server.ts** file (located under **src/server**) and fix the searchFilesDialog import to the new name and change it's use at the bottom of the run function.

```
...
import searchDialog from './dialogs/searchDialog';

export class Server {
    run() {
        ...

        // create dialogs for specific intents
        let dlg = new searchDialog(authHelper);
        bot.dialog(dlg.id, dlg.waterfall).triggerAction({ matches: dlg.name });
    }
}
```
3. Create a **sentimentHelper.ts** file under **src/server/helpers** and add the following code for calling the Text Analytics service for sentiment analysis.

```
require('dotenv').config();

import * as https from 'https';
import { HttpHelper } from './httpHelper';

export class SentimentHelper {
    // get sentiment score for all strings in the array
    static getSentimentScore(text: Array<string>) {
        return new Promise((resolve, reject) => {
            // set header information, including cognitive services subscription key
            let header = {
                'Ocp-Apim-Subscription-Key': process.env.COG_SUB_KEY,
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            }

            // prepare the payload for the text analytics service
            let payload = {documents: []};
            for (var i = 0; i < text.length; i++)
                payload.documents.push({
                    "language": "en",
                    "id": i,
                    "text": text[i]
                });
            
            // use the HttpHelper to post the payload and get back sentiment
            HttpHelper.postJson(header, 'westus.api.cognitive.microsoft.com', '/text/analytics/v2.0/sentiment', payload).then((data: any) => {
                // aggregate the total sentiment score
                let totalScore = 0.0;
                for (var i = 0; i < data.documents.length; i++) {
                    totalScore += data.documents[i].score;
                }

                // resolve the total score divided by the number of results
                resolve(totalScore / data.documents.length);
            }, (err) => {
                reject(err);
            });         
        });
    }
}
```
4. Open the **searchDialog.ts** (located under **src/server/dialogs**) and add reference to sentimentHelper at the top.
```
import { SentimentHelper } from '../helpers/sentimentHelper';
```
5. Modify the searchDialog for the new LUIS model but changing the id/name to **Search**, adding new properties for **term** and **source** then set them in the first waterfall step as seen below.
```
export class searchDialog implements IDialog {
    term, source;
    constructor(private authHelper: AuthHelper) {
        this.id = 'Search';
        this.name = 'Search';
        this.waterfall = [].concat(
            (session, args, next) => {
                // Read the LUIS detail and then move to auth
                this.term = builder.EntityRecognizer.findEntity(args.intent.entities, 'Term');
                this.source = builder.EntityRecognizer.findEntity(args.intent.entities, 'Source');
                next();                
            },
```
6. Update the Microsoft Search query to search mail messages as seen below.
```
let endpoint = `/v1.0/me/messages?$search="${encodeURIComponent(this.term.entity)}"&$select=bodyPreview&$top=5`;
```
7. Finally, handle the data returned by the graph by building a string array of bodyPreviews and send that to the **sentimentHelper** class.
```
HttpHelper.getJson(headers, 'graph.microsoft.com', endpoint).then(function(data: any) {
    // build string array with bodyPreview of each message
    let text = [];
    for (var i = 0; i < data.value.length; i++) {
        text.push(data.value[i].bodyPreview);
    }

    // use the sentiment helper to get sentiment score for the messages
    SentimentHelper.getSentimentScore(text).then((score) => {
        session.endConversation(`Sentiment Score: ${score}`);
    }, (err) => {
        session.endConversation(`Error occurred: ${err}`);
    });
}).catch(function(err) {
    // something went wrong
    session.endConversation(`Error occurred: ${err}`);
});
```
7. The completed **sentimentHelper** file should look as follows.
```
import { IDialog } from './idialog';
import * as builder from 'botbuilder';
import * as restify from 'restify';
import { AuthHelper } from '../helpers/authHelper';
import { HttpHelper } from '../helpers/httpHelper';
import { SentimentHelper } from '../helpers/sentimentHelper';

export class searchDialog implements IDialog {
    constructor(private authHelper: AuthHelper) {
        this.id = 'Search';
        this.name = 'Search';
        this.waterfall = [].concat(
            (session, args, next) => {
                // Read the LUIS detail and then move to auth
                this.term = builder.EntityRecognizer.findEntity(args.intent.entities, 'Term');
                this.source = builder.EntityRecognizer.findEntity(args.intent.entities, 'Source');
                next();                
            },
            authHelper.getAccessToken(),
            (session, results, next) => {
                if (results.response != null) {
                    // make a call to the Microsoft Graph to search messages
                    let headers = {
                        Accept: 'application/json',
                        Authorization: 'Bearer ' + results.response
                    };
                    let endpoint = `/v1.0/me/messages?$search="${encodeURIComponent(this.term.entity)}"&$select=bodyPreview&$top=5`;
                    HttpHelper.getJson(headers, 'graph.microsoft.com', endpoint).then(function(data: any) {
                        // build string array with bodyPreview of each message
                        let text = [];
                        for (var i = 0; i < data.value.length; i++) {
                            text.push(data.value[i].bodyPreview);
                        }

                        // use the sentiment helper to get sentiment score for the messages
                        SentimentHelper.getSentimentScore(text).then((score) => {
                            session.endConversation(`Sentiment Score: ${score}`);
                        }, (err) => {
                            session.endConversation(`Error occurred: ${err}`);
                        });
                    }).catch(function(err) {
                        // something went wrong
                        session.endConversation(`Error occurred: ${err}`);
                    });
                }
                else {
                    session.endConversation('Sorry, I did not understand');
                }
            }
        );
    }
    
    id; name; waterfall; term; source;
}

export default searchDialog;
```

## Task 4 - Test the solution
1. Start the application.

```
npm start
```

2. Launch the Bot Framework Emulator (available [HERE](https://docs.botframework.com/en-us/tools/bot-framework-emulator/)).

3. Enter the messaging endpoint for the bot project (likely http://localhost:3980/api/messages but could be a different port) and the click the "Connect" button (leave the Microsoft App ID and Microsoft App Password blank as these are for published bots).

4. Type a query such as ***Search Office 365 for Graph***.

5. If you have troubles, the completed solution is in this repo.