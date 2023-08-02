// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// Import required packages
//import { config } from 'dotenv';
import * as path from 'path';
import * as restify from 'restify';

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication,
    ConfigurationServiceClientCredentialFactory,
    MemoryStorage,
    TurnContext
} from 'botbuilder';

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
    {},
    new ConfigurationServiceClientCredentialFactory({
        MicrosoftAppId: process.env.BOT_ID,
        MicrosoftAppPassword: process.env.BOT_PASSWORD,
        MicrosoftAppType: 'MultiTenant'
    })
);

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about how bots work.
const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: any) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    console.error(`\n [onTurnError] unhandled error: ${error}`);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${error}`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    // Send a message to the user
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity(`${error}`);
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

function respond(req, res, next) {
    res.send('hello ' + req.params.name);
    next();
  }

server.get('/hello/:name', respond);
server.head('/hello/:name', respond);

server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${server.name} listening to ${server.url}`);
    console.log('\nGet Bot Framework Emulator: https://aka.ms/botframework-emulator');
    console.log('\nTo test your bot in Teams, sideload the app manifest.json within Teams Apps.');
});

import {
    AI,
    Application,
    ConversationHistory,
    DefaultPromptManager,
    DefaultTurnState,
    AzureOpenAIPlanner,
    OpenAIModerator,
    OpenAIPlanner
} from '@microsoft/teams-ai';

// eslint-disable-next-line @typescript-eslint/no-empty-interface
interface ConversationState {}
type ApplicationTurnState = DefaultTurnState<ConversationState>;

// Iterate over all available environment variables
for (const key in process.env) {
    if (Object.prototype.hasOwnProperty.call(process.env, key)) {
      const value = process.env[key];
      console.log(`${key}: ${value}`);
    }
  }

if (!process.env.OPENAI_API_KEY) {
    throw new Error('Missing environment variables - please check that OpenAIKey is set.');
}
// Create Azure Open AI components
const planner = new AzureOpenAIPlanner({
    apiKey: process.env.OPENAI_API_KEY,
    defaultModel: 'gpt-35-deployment',
    logRequests: true,
    endpoint: process.env.OPENAI_ENDPOINT,
    apiVersion: null//'2023-05-15'
  });

 const moderator = new OpenAIModerator({
     apiKey: process.env.OPENAI_API_KEY || '',
     moderate: 'both',
     endpoint: process.env.OPENAI_ENDPOINT,
 });
// Create AI components
// const planner = new OpenAIPlanner({
//     apiKey: process.env.OPENAI_API_KEY || '',
//     defaultModel: 'text-davinci-003',
//     logRequests: true
// });
// const moderator = new OpenAIModerator({
//     apiKey: process.env.OPENAI_API_KEY || '',
//     moderate: 'both'
// });
const promptManager = new DefaultPromptManager(path.join(__dirname, '../src/prompts'));

// Define storage and application
const storage = new MemoryStorage();
const app = new Application<ApplicationTurnState>({
    storage,
    ai: {
        planner: planner,
        moderator: null, //moderator, // 
        promptManager: promptManager,
        prompt: 'chat',
        history: {
            assistantHistoryType: 'text'
        }
    }
});

app.ai.action(
    AI.FlaggedInputActionName,
    async (context: TurnContext, state: ApplicationTurnState, data: Record<string, any>) => {
        await context.sendActivity(`I'm sorry your message was flagged: ${JSON.stringify(data)}`);
        return false;
    }
);

app.ai.action(AI.FlaggedOutputActionName, async (context: TurnContext, state: ApplicationTurnState, data: any) => {
    await context.sendActivity(`I'm not allowed to talk about such things.`);
    return false;
});

app.message('/history', async (context: TurnContext, state: ApplicationTurnState) => {
    console.log('sending history message');
    const history = ConversationHistory.toString(state, 2000, '\n\n');
    await context.sendActivity(history);
});

// Listen for incoming server requests.
server.post('/api/messages', (req, res, next) => {
    console.log('sending api/messages post');
    // Route received a request to adapter for processing
    return adapter.process(req, res as any, async (context) => {
        // Dispatch to application for routing
        await app.run(context);
    }).then(()=>{
        return next()
    });
});
