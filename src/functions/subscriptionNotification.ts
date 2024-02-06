import { app, HttpHandler, HttpRequest, HttpResponse, InvocationContext, output } from '@azure/functions';
import { ChangeNotification, ChangeNotificationCollection } from "@microsoft/microsoft-graph-types";
import * as df from 'durable-functions';
import { ActivityHandler, OrchestrationContext, OrchestrationHandler } from 'durable-functions';
import { Graph } from "../../modules/graph";

const orchestratorName = 'subscriptionNotificationOrchestrator';
const getResourceFromGraphActivityName = 'subscriptionNotification';

const cosmosOutput = output.cosmosDB({
    databaseName: "teamsCallLogger",
    containerName: "calls",
    connection: "COSMOS_DB_CONNECTION",
    createIfNotExists: true,
    partitionKey: "/id"
});

const subscriptionNotificationOrchestrator: OrchestrationHandler = function* (context: OrchestrationContext) {
    try {
        // Cast input to ChangeNotificationCollection
        const input = context.df.getInput() as ChangeNotificationCollection;
        // Get the list of notifications
        const notifications = input.value;
        context.log(`Processing ${notifications.length} notifications(s)`);
        const result = yield context.df.callActivity(getResourceFromGraphActivityName, notifications);
        return result;
    } catch (error) {
        context.log(error);
    }
};
df.app.orchestration('subscriptionNotificationOrchestrator', subscriptionNotificationOrchestrator);

const getResourceFromGraph: ActivityHandler = async (input: ChangeNotification[], context: InvocationContext) => {
    const graph = new Graph(context);
    const resources = await graph.batchGetChangeNotificationResources(input);
    context.extraOutputs.set(cosmosOutput, resources);
    return resources;
};
df.app.activity(getResourceFromGraphActivityName, { handler: getResourceFromGraph, extraOutputs: [cosmosOutput] });

const subscriptionNotificationHttpStart: HttpHandler = async (request: HttpRequest, context: InvocationContext): Promise<HttpResponse> => {

    // Validate a new subscription
    if (request.query.get("validationToken")) {
        context.log("Validating new subscription notification...");
        // Return a 200 OK response and provide the instance id in the response body
        const response = new HttpResponse({
            status: 200,
            body: request.query.get("validationToken")
        });
        return response;
    }
    else {
        const body = await request.json() as ChangeNotificationCollection;
        if (body && body.value && body.value.length > 0) {
            context.log('Received new notification(s) from Microsoft Graph subscription');
            const client = df.getClient(context);
            const instanceId = await client.startNew(orchestratorName, { input: body });
            context.log(`Started orchestration with ID: '${instanceId}'`);
            // Return a 200 OK response and provide the instance id in the response body
            const response = new HttpResponse({
                status: 200,
                jsonBody: {
                    instanceId
                }
            });
            return response;
        }
    }

};

app.http('subscriptionNotificationHttpStart', {
    route: 'subscriptionNotification',
    extraInputs: [df.input.durableClient()],
    handler: subscriptionNotificationHttpStart
});