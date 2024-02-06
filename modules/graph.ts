import { BatchRequestBody, BatchRequestContent, BatchRequestData, BatchRequestStep, BatchResponseContent, Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { ClientSecretCredential } from "@azure/identity";
import { ChangeNotification, Subscription } from "@microsoft/microsoft-graph-types";
import { InvocationContext } from "@azure/functions"

export class Graph {

    private context: InvocationContext;
    public client: Client;

    constructor(context: InvocationContext) {
        try {
            this.context = context;
            const credential = new ClientSecretCredential(
                process.env.MicrosoftAppTenantId as string,
                process.env.MicrosoftAppId as string,
                process.env.MicrosoftAppPassword as string
            );
            // Auth provider
            const authProvider = new TokenCredentialAuthenticationProvider(credential, {
                scopes: ['https://graph.microsoft.com/.default'],
            });
            const clientOptions: ClientOptions = {
                defaultVersion: "v1.0",
                debugLogging: false,
                authProvider
            };
            const client = Client.initWithMiddleware(clientOptions);
            this.client = client;
        } catch (error) {
            throw error;
        }
    }

    public async getSubscriptions(): Promise<Subscription[]> {
        try {
            const result = await this.client.api('/subscriptions')
                .get();
            return result.value;
        } catch (error) {
            throw error;
        }
    }

    async createSubscription(subscriptionLengthDays: number = 1, resource: string): Promise<Subscription> {
        const expiryDate = Date.now() + (subscriptionLengthDays * 24 * 60 * 60 * 1000); // X Days * 24 hrs * 60 mins * 60 secs * 1000ms
        const expiryDateString = new Date(expiryDate).toISOString();

        const responseBody: Subscription = {
            changeType: "created",
            notificationUrl: `${process.env.WEB_HOST}/api/subscriptionNotification`,
            resource: resource,
            expirationDateTime: expiryDateString
        }

        const response = await this.client.api("/subscriptions")
            .post(responseBody)
            .catch((error) => {
                this.context.error(error);
            });
        return response;
    };

    async batchGetChangeNotificationResources(changeNotifications: ChangeNotification[]): Promise<any> {
        try {
            const batchRequestData = changeNotifications.map((notification, index) => {
                this.context.log(`----------------------------------------------------------`);
                this.context.log(`Notification #${index + 1}`);
                this.context.log(`----------------------------------------------------------`);
                this.context.log(`Subscription Id    : ${notification.subscriptionId}`);
                this.context.log(`Expiration         : ${notification.subscriptionExpirationDateTime}`);
                this.context.log(`Change Type        : ${notification.changeType}`);
                this.context.log(`Client State       : ${notification.clientState}`);
                this.context.log(`Resource           : ${notification.resource}`);
                this.context.log(`----------------------------------------------------------`);

                const requestStep: BatchRequestData = {
                    id: index.toString(),
                    method: "GET",
                    url: notification.resource
                };

                return requestStep;
            });

            const batchRequestBody: BatchRequestBody = {
                requests: batchRequestData
            };
            const batchResponse = await this.client.api('/$batch')
                .post(batchRequestBody);
            const batchResponseContent = new BatchResponseContent(batchResponse);

            const results = await Promise.all(
                changeNotifications.map(async (notification, index) => {
                    const response = batchResponseContent.getResponseById(index.toString());
                    if (response.status === 200) {
                        return response.json();
                    } else {
                        return null;
                    }
                }),
            );

            return results;
        } catch (error) {
            throw error;
        }
    }

}