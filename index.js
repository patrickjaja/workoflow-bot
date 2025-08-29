const path = require('path');

const dotenv = require('dotenv');
// Import required bot configuration.
const ENV_FILE = path.join(__dirname, '.env');
dotenv.config({ path: ENV_FILE });

const restify = require('restify');
const axios = require('axios');

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
const {
    CloudAdapter,
    ConfigurationServiceClientCredentialFactory,
    createBotFrameworkAuthenticationFromConfiguration
} = require('botbuilder');

// This bot's main dialog.
const { EchoBot } = require('./bot');

// Create HTTP server
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.listen(process.env.WORKOFLOW_PORT || 3978, () => {
    console.log(`\n${ server.name } listening to ${ server.url }`);
    console.log('\nGet Bot Framework Emulator: https://aka.ms/botframework-emulator');
    console.log('\nTo talk to your bot, open the emulator select "Open Bot"');
});

const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
    MicrosoftAppId: process.env.MicrosoftAppId,
    MicrosoftAppPassword: process.env.MicrosoftAppPassword,
    MicrosoftAppType: process.env.MicrosoftAppType,
    MicrosoftAppTenantId: process.env.MicrosoftAppTenantId
});

const botFrameworkAuthentication = createBotFrameworkAuthenticationFromConfiguration(null, credentialsFactory);

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
const onTurnErrorHandler = async (context, error) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    console.error(`\n [onTurnError] unhandled error: ${ error }`);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${ error }`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    // Send a message to the user
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create the main dialog.
const myBot = new EchoBot();

// Health check endpoint with orchestrator API connectivity check
server.get('/api/health', async (req, res) => {
    const healthStatus = {
        status: 'healthy',
        timestamp: new Date().toISOString(),
        service: 'workoflow-bot',
        port: process.env.WORKOFLOW_PORT || 3978,
        services: {
            bot: 'operational'
        }
    };

    // Check orchestrator API connectivity if configured
    const ORCHESTRATOR_API_URL = process.env.ORCHESTRATOR_API_URL;
    const ORCHESTRATOR_API_KEY = process.env.ORCHESTRATOR_API_KEY;
    
    if (ORCHESTRATOR_API_URL && ORCHESTRATOR_API_KEY) {
        try {
            const orchestratorHealthResponse = await axios.get(
                `${ORCHESTRATOR_API_URL}/health`,
                {
                    headers: {
                        'x-api-key': ORCHESTRATOR_API_KEY
                    },
                    timeout: 5000 // 5 second timeout for health check
                }
            );
            
            healthStatus.services.orchestrator_api = {
                status: 'connected',
                url: ORCHESTRATOR_API_URL,
                response: orchestratorHealthResponse.data
            };
        } catch (error) {
            healthStatus.services.orchestrator_api = {
                status: 'disconnected',
                url: ORCHESTRATOR_API_URL,
                error: error.message
            };
            healthStatus.status = 'degraded';
        }
    } else {
        healthStatus.services.orchestrator_api = {
            status: 'not_configured',
            message: 'Orchestrator API credentials not provided'
        };
    }

    // Check n8n webhook connectivity (optional)
    const N8N_WEBHOOK_URL = process.env.WORKOFLOW_N8N_WEBHOOK_URL;
    if (N8N_WEBHOOK_URL) {
        healthStatus.services.n8n_webhook = {
            status: 'configured',
            url: N8N_WEBHOOK_URL
        };
    }

    res.send(healthStatus);
});

// Listen for incoming requests.
server.post('/api/messages', async (req, res) => {
    // Route received a request to adapter for processing
    await adapter.process(req, res, (context) => myBot.run(context));
});

// Listen for Upgrade requests for Streaming.
server.on('upgrade', async (req, socket, head) => {
    // Create an adapter scoped to this WebSocket connection to allow storing session data.
    const streamingAdapter = new CloudAdapter(botFrameworkAuthentication);

    // Set onTurnError for the CloudAdapter created for each connection.
    streamingAdapter.onTurnError = onTurnErrorHandler;

    await streamingAdapter.process(req, socket, head, (context) => myBot.run(context));
});
