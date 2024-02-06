# Teams Call Logger

## Summary

A basic example of capturing Teams call records using Microsoft Graph subscriptions and storing them in a Cosmos DB database.

## Functions overview

This project contains a single Azure Function that is triggered by a Microsoft Graph subscription for call records. When a new call record is received, the function processes the record and stores it in a Cosmos DB database.

* `subscriptionManager` - Runs on a timer trigger and manages the Microsoft Graph subscription for call records. It checks for an existing subscription and creates one if it doesn't exist.

* `subscriptionNotification` - Receives notifications from the Microsoft Graph subscription for call records. It processes the call record and stores it in a Cosmos DB database.

## Prerequisites

* [Azure Functions Core Tools v4](https://github.com/Azure/azure-functions-core-tools) installed (latest version)
* [Node.js](https://nodejs.org) version 18 or higher installed:

    ```bash
    # Determine node version
    node --version
    ```

* [ngrok](https://ngrok.com/) installed. Although a free account will work with this sample, the tunnel subdomain will change each time you run ngrok, requiring a change to the Azure Bot messaging endpoint and the Teams app manifest. A paid account with a permanent subdomain is recommended.

## Version history

Version|Date|Author|Comments
-------|----|----|--------
1.0|Feb 6, 2024|Lee Ford|Initial release

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

### Start ngrok

Start ngrok listening on port 7071 (this is what Azure Functions uses by default):

```bash
ngrok http 7071
```

If you have a paid account, add the subdomain option to the command:

```bash
# Replace 12345678 with your ngrok subdomain
ngrok http 7071 -subdomain=12345678
```

> Take a note of the forwarding URL, as you will need it later.

### Create Azure resources

It is recommended to use [Azure Cloud Shell](https://shell.azure.com) for this step, but you can also use the `az` CLI on your local machine if you prefer.

Switch to the Azure subscription you want to use:

```bash
az account set --subscription <subscriptionId>
```

> Replace `<subscriptionId>` with the ID of the subscription you want to use

Create a resource group:

```bash
az group create --name <resourceGroupName> --location <location>
```

> Choose an Azure region and replace `<location>` with the region name (e.g. `uksouth`).

Create an app registration:

```bash
az ad app create --display-name "Teams Call Logger" --sign-in-audience "AzureADMyOrg"
```

> Take note of the `appId` value from the output as this will be required in later steps.

Add the `CallRecords.Read.All` application permission (and grant consent) to the app registration:

```bash
az ad app permission add --id <appId> --api 00000003-0000-0000-c000-000000000000 --api-permissions 45bbb07e-7321-4fd7-a8f6-3ff27e6a81c8=Role
ad ad app permission admin-consent --id <appId>
```

> Replace `<appId>` with the `appId` value from an earlier step.

Reset the app registration credentials, making note of the `password` and `tenant` values from the output as this will be required in later steps:

```bash
az ad app credential reset --id <appId>
```

> Replace `<appId>` with the `appId` value from an earlier step.

Create a Cosmos DB database:

```bash
az cosmosdb create --name <accountName> --resource-group <resourceGroupName> --default-consistency-level Eventual --locations regionName="<location>" failoverPriority=0 isZoneRedundant=False --capabilities EnableServerless
```

> * Replace `<accountName>` with the name of the Cosmos DB account you want to use. This MUST be lowercase
> * Replace `<resourceGroupName>` with the name of the resource group created earlier
> * Replace `<location>` with the location name (e.g. `uksouth`)
> * Take note of the `documentEndpoint` value from the output as this will be required in later steps

Finally, get the primary SQL connection string for the Cosmos DB account:

```bash
az cosmosdb keys list --type connection-strings --name <accountName> --resource-group <resourceGroupName>
```

> * Replace `<accountName>` with the name of the Cosmos DB account you used earlier
> * Replace `<resourceGroupName>` with the name of the resource group created earlier
> * Take note of the `connectionString` value from the "Primary" key output as this will be required in later steps

Create a storage account:

```bash
az storage account create --name <accountName> --resource-group <resourceGroupName> --location <location> --sku Standard_LRS
```

> * Replace `<accountName>` with the name of the storage account you want to use. This MUST be lowercase and alphanumeric only
> * Replace `<resourceGroupName>` with the name of the resource group created earlier
> * Replace `<location>` with the location name (e.g. `uksouth`)

Get the connection string for the storage account:

```bash
az storage account show-connection-string --name <accountName> --resource-group <resourceGroupName>
```

> * Replace `<accountName>` with the name of the storage account you used earlier
> * Replace `<resourceGroupName>` with the name of the resource group created earlier
> * Take note of the `connectionString` value from the output as this will be required in later steps

### Run locally

1. Clone this repository
2. Create and populate a `local.settings.json` file in the `source` folder with the following (with your own values):

    ```json
    {
    "IsEncrypted": false,
    "Values": {
        "AzureWebJobsStorage": "<storageConnectionString>",
        "FUNCTIONS_WORKER_RUNTIME": "node",
        "AzureWebJobsFeatureFlags": "EnableWorkerIndexing",
        "MicrosoftAppTenantId": "<tenantId>",
        "MicrosoftAppId": "<appId>",
        "MicrosoftAppPassword": "<appPassword>",
        "MicrosoftAppType": "SingleTenant",
        "WEB_HOST": "<ngrokForwardingUrl>",
        "COSMOS_DB_CONNECTION": "<cosmosDbConnectionString>",
        }
    }
    ```

    > * Replace `<storageConnectionString>` with the `connectionString` value from the storage account created earlier
    > * Replace `<tenantId>` with the `tenant` value from the app registration created earlier
    > * Replace `<appId>` with the `appId` value from the app registration created earlier
    > * Replace `<appPassword>` with the `password` value from the app registration created earlier
    > * Replace `<cosmosDbEndpoint>` with the `documentEndpoint` value from the Cosmos DB account created earlier
    > * Replace `<cosmosDbConnectionString>` with the `primaryMasterKey` value from the Cosmos DB account created earlier
    > * Replace `<ngrokForwardingUrl>` with the forwarding URL from ngrok

3. Run the following commands in the `source` folder:

    ```bash
    npm install
    npm start
    ```

4. Leave running and new Teams call records will be captured and stored in the Cosmos DB database
