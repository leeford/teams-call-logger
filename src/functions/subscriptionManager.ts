import { app, InvocationContext, Timer } from '@azure/functions';
import * as df from 'durable-functions';
import { ActivityHandler, OrchestrationContext, OrchestrationHandler } from 'durable-functions';
import { Graph } from "../../modules/graph";

const orchestratorName = 'subscriptionManagerOrchestrator';
const activityName = 'subscriptionManager';

const subscriptionManagerOrchestrator: OrchestrationHandler = function* (context: OrchestrationContext) {
    const outputs = [];
    outputs.push(yield context.df.callActivity(activityName));

    return outputs;
};
df.app.orchestration('subscriptionManagerOrchestrator', subscriptionManagerOrchestrator);

const subscriptionManager: ActivityHandler = async (input: any, context: InvocationContext) => {
    const resource = "communications/callRecords";
    const graph = new Graph(context);
    const subscriptions = await graph.getSubscriptions();
    // Check if there are any subscriptions
    // And at least one matches the resource
    const filteredSubscriptions = subscriptions.filter((subscription) => {
        return subscription.resource === resource;
    });
    if (filteredSubscriptions.length > 0) {
        context.log(`${filteredSubscriptions.length} subscriptions found for resource ${resource}`);
    } else {
        // Create new subscription
        context.log("Creating new subscription");
        await graph.createSubscription(2, resource);
    }
};
df.app.activity(activityName, { handler: subscriptionManager });

export async function subscriptionManagerTimer(timer: Timer, context: InvocationContext): Promise<void> {
    const client = df.getClient(context);
    const instanceId: string = await client.startNew(orchestratorName);
    context.log(`Started orchestration with ID = '${instanceId}'.`);
}

app.timer('subscriptionManagerTimer', {
    schedule: '0 */5 * * * *',
    extraInputs: [df.input.durableClient()],
    handler: subscriptionManagerTimer
});
