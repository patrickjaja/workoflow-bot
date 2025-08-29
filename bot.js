const { ActivityHandler, MessageFactory, CardFactory } = require('botbuilder');
const axios = require('axios');

// Orchestrator API Configuration
const ORCHESTRATOR_API_URL = process.env.ORCHESTRATOR_API_URL || 'http://localhost:8080';
const ORCHESTRATOR_API_KEY = process.env.ORCHESTRATOR_API_KEY;
const ORCHESTRATOR_ORG_ID = process.env.ORCHESTRATOR_ORG_ID || 'test-ms-org';

// Fallback n8n configuration
const N8N_WEBHOOK_URL = process.env.WORKOFLOW_N8N_WEBHOOK_URL || 'https://workflows.vcec.cloud/webhook/016d8b95-d5a5-4ac6-acb5-359a547f642f';
const N8N_BASIC_AUTH_USERNAME = process.env.N8N_BASIC_AUTH_USERNAME;
const N8N_BASIC_AUTH_PASSWORD = process.env.N8N_BASIC_AUTH_PASSWORD;

console.log('ORCHESTRATOR_API_URL:', ORCHESTRATOR_API_URL);
console.log('N8N_WEBHOOK_URL (fallback):', N8N_WEBHOOK_URL);

class EchoBot extends ActivityHandler {
    constructor() {
        super();
        
        // Store conversation state for OAuth re-prompt
        this.conversationState = new Map();

        this.onMessage(async (context, next) => {
            try {
                // Log activity for debugging
                console.log('=== FULL ACTIVITY OBJECT ===');
                console.log(JSON.stringify(context.activity, null, 2));
                console.log('=== END ACTIVITY ===');

                const conversationId = context.activity.conversation.id;
                const userMessage = context.activity.text;
                
                // Store the last message for this conversation (for OAuth re-prompt)
                this.conversationState.set(conversationId, {
                    lastMessage: userMessage,
                    timestamp: new Date().toISOString(),
                    activity: context.activity
                });

                // Send thinking indicator
                await context.sendActivity(MessageFactory.text('Thinking...  \n(Responses will be generated using AI and may contain mistakes.)', 'Thinking...  \n(Responses will be generated using AI and may contain mistakes.)'));

                // Try orchestrator API first
                let response = null;
                let useOrchestratorAPI = ORCHESTRATOR_API_KEY && ORCHESTRATOR_API_URL;
                
                if (useOrchestratorAPI) {
                    try {
                        response = await this.callOrchestratorAPI(context.activity);
                        
                        // Handle orchestrator API response
                        if (response) {
                            await this.handleOrchestratorResponse(context, response);
                        }
                    } catch (orchestratorError) {
                        console.error('Orchestrator API failed, falling back to n8n:', orchestratorError.message);
                        useOrchestratorAPI = false;
                    }
                }
                
                // Fallback to n8n webhook if orchestrator API failed or not configured
                if (!useOrchestratorAPI) {
                    try {
                        response = await this.callN8NWebhook(context.activity);
                        await this.handleN8NResponse(context, response);
                    } catch (n8nError) {
                        console.error('Both orchestrator API and n8n webhook failed:', n8nError.message);
                        await context.sendActivity(MessageFactory.text('There was an error communicating with the AI services. Please try again later.'));
                    }
                }

            } catch (error) {
                console.error('Error in message handler:', error);
                await context.sendActivity(MessageFactory.text('An unexpected error occurred. Please try again.'));
            }

            await next();
        });

        // Handle adaptive card actions (for OAuth callbacks)
        this.onMessageReaction(async (context, next) => {
            console.log('=== MESSAGE REACTION ===');
            console.log(JSON.stringify(context.activity, null, 2));
            await next();
        });

        // Add handler for all events to catch file-related activities
        this.onEvent(async (context, next) => {
            console.log('=== EVENT ACTIVITY ===');
            console.log('Event name:', context.activity.name);
            console.log('Event value:', context.activity.value);
            
            // File consent activities in Teams
            if (context.activity.name === 'fileConsent/invoke') {
                console.log('FILE CONSENT DETECTED!');
                console.log('File info:', context.activity.value);
            }

            await next();
        });

        // Handle unrecognized activity types
        this.onUnrecognizedActivityType(async (context, next) => {
            console.log('=== UNRECOGNIZED ACTIVITY TYPE ===');
            console.log('Type:', context.activity.type);
            console.log('Full activity:', JSON.stringify(context.activity, null, 2));
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const welcomeText = 'Hello and welcome! I am your AI Assistant powered by the Workoflow Orchestrator. How can I help you today?';
            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                }
            }
            await next();
        });
    }

    async callOrchestratorAPI(activity) {
        console.log('=== CALLING ORCHESTRATOR API ===');
        
        // Transform Teams activity to orchestrator API format
        const apiRequest = {
            message: activity.text || '',
            user_id: `teams-${activity.from.id}`,
            channel: 'teams',
            conversation_id: activity.conversation.id,
            organization_id: ORCHESTRATOR_ORG_ID,
            metadata: {
                teams_user_id: activity.from.id,
                teams_tenant_id: activity.channelData?.tenant?.id || '',
                teams_channel_id: activity.channelId || '',
                teams_conversation_type: activity.conversation.conversationType || '',
                teams_user_name: activity.from.name || '',
                teams_aad_object_id: activity.from.aadObjectId || ''
            }
        };

        // Add attachment info if present
        if (activity.attachments && activity.attachments.length > 0) {
            apiRequest.metadata.attachments = activity.attachments.map(att => ({
                contentType: att.contentType,
                name: att.name,
                contentUrl: att.contentUrl
            }));
        }

        console.log('API Request:', JSON.stringify(apiRequest, null, 2));

        const response = await axios.post(
            `${ORCHESTRATOR_API_URL}/api/chat/`,
            apiRequest,
            {
                headers: {
                    'x-api-key': ORCHESTRATOR_API_KEY,
                    'Content-Type': 'application/json'
                },
                timeout: 30000 // 30 second timeout
            }
        );

        console.log('Orchestrator API Response:', JSON.stringify(response.data, null, 2));
        return response.data;
    }

    async handleOrchestratorResponse(context, response) {
        // Check for attachments first (for adaptive cards)
        if (response.attachments && response.attachments.length > 0) {
            // Handle adaptive card or other attachments
            for (const attachment of response.attachments) {
                if (attachment.contentType === 'application/vnd.microsoft.card.adaptive') {
                    // Render adaptive card (for OAuth or other purposes)
                    const adaptiveCard = CardFactory.adaptiveCard(attachment.content);
                    await context.sendActivity(MessageFactory.attachment(adaptiveCard));
                } else {
                    // Handle other attachment types
                    await context.sendActivity(MessageFactory.attachment(attachment));
                }
            }
        } else if (response.type === 'message' && response.message) {
            // Simple text message response
            const message = response.message || response.content || 'No response content';
            await context.sendActivity(MessageFactory.text(message, message));
        } else {
            // Fallback for unexpected response format
            console.warn('Unexpected response format:', response);
            const fallbackMessage = response.message || response.content || 'No response content';
            await context.sendActivity(MessageFactory.text(fallbackMessage, fallbackMessage));
        }
    }

    async callN8NWebhook(activity) {
        console.log('=== CALLING N8N WEBHOOK (FALLBACK) ===');
        
        const config = {};
        if (N8N_BASIC_AUTH_USERNAME && N8N_BASIC_AUTH_PASSWORD) {
            config.auth = {
                username: N8N_BASIC_AUTH_USERNAME,
                password: N8N_BASIC_AUTH_PASSWORD
            };
        }

        // Detect file URLs and attachments
        let detectedFileUrls = [];
        if (activity.text) {
            const sharePointPattern = /(https:\/\/[^\s]*\.(sharepoint\.com|microsoft\.com|office\.com)[^\s]*)/gi;
            const teamsFilePattern = /(https:\/\/teams\.microsoft\.com[^\s]*)/gi;
            
            const sharePointUrls = activity.text.match(sharePointPattern) || [];
            const teamsUrls = activity.text.match(teamsFilePattern) || [];
            
            detectedFileUrls = [...sharePointUrls, ...teamsUrls];
            
            if (detectedFileUrls.length > 0) {
                console.log('=== DETECTED FILE URLS IN TEXT ===');
                console.log(detectedFileUrls);
            }
        }

        // Create enriched payload for n8n
        const enrichedPayload = {
            ...activity,
            _fileDetection: {
                hasNonHtmlAttachments: activity.attachments?.some(att =>
                    att.contentType !== 'text/html' && att.contentType !== 'text/plain'
                ) || false,
                detectedFileUrls: detectedFileUrls,
                attachmentTypes: activity.attachments?.map(att => att.contentType) || [],
                entityTypes: activity.entities?.map(ent => ent.type) || [],
                possibleFileAttachments: activity.attachments?.filter(att =>
                    att.contentType !== 'text/html' &&
                    att.contentType !== 'text/plain' &&
                    att.contentUrl
                ) || []
            }
        };

        console.log('=== SENDING TO N8N ===');
        console.log('File detection summary:', enrichedPayload._fileDetection);

        const response = await axios.post(N8N_WEBHOOK_URL, enrichedPayload, config);
        console.log('N8N Response:', response.data);
        return response.data;
    }

    async handleN8NResponse(context, response) {
        let n8nReplyText = 'Sorry, I could not get a response from the agent.';
        let attachmentUrl = null;

        // Handle the n8n response structure
        if (response && response.output && Array.isArray(response.output) && response.output.length > 0) {
            const outputItem = response.output[0];
            if (outputItem.output) {
                n8nReplyText = outputItem.output;
            }
            // Check for optional attachment
            if (outputItem.attachment && outputItem.attachment.url) {
                attachmentUrl = outputItem.attachment.url;
            }
        }

        // Send the response with or without attachment
        if (attachmentUrl) {
            const replyWithLink = `${n8nReplyText}\n\n📎 [Download attachment](${attachmentUrl})`;
            await context.sendActivity(MessageFactory.text(replyWithLink, replyWithLink));
        } else {
            await context.sendActivity(MessageFactory.text(n8nReplyText, n8nReplyText));
        }
    }

    // Helper method to get stored conversation state
    getConversationState(conversationId) {
        return this.conversationState.get(conversationId);
    }

    // Helper method to clear conversation state (after successful auth)
    clearConversationState(conversationId) {
        this.conversationState.delete(conversationId);
    }
}

module.exports.EchoBot = EchoBot;