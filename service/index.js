const uuidv4 = require('uuid/v4');
const azure = require('azure-storage');
const webpush = require('web-push');

const PartitionKey = 'subscriptions';
const connectionString = process.env["outlookwebnotifier_STORAGE"];
const vapidKeys = {
    publicKey: process.env["VAPID_PUBLIC_KEY"],
    privateKey: process.env["VAPID_PRIVATE_KEY"]
};
const options = {
    vapidDetails: {
        subject: 'mailto: sunilbandla@gmail.com',
        publicKey: vapidKeys.publicKey,
        privateKey: vapidKeys.privateKey
    },
    // 1 hour in seconds.
    TTL: 60 * 60
};

module.exports = function (context, req, storageTable) {
    context.log("Function called with token =", req.query.validationToken,
     "body = ", req.body);
    
    context.bindings.clientState = uuidv4();

    if (req.method === "DELETE" && context.bindingData.subscriptionId) {
        deleteSubscription(context, context.bindingData.subscriptionId);
    }
    else if (!req.query.validationToken && req.body && req.body.value &&
        req.body.value.length > 0) {
        // Read notification from Graph and send push message
        context.log("Received notification details ", req.body.value,
         context.bindings.inputStorageTable);
        var notifications = req.body.value;
        var entities = context.bindings.inputStorageTable;
        for (var i = 0; i < notifications.length; i++) {
            entities.filter(function (entity) {
                return (entity.RowKey === notifications[i].subscriptionId) &&
                    (entity.clientState === notifications[i].clientState);
            }).forEach(function (subscriptionEntity) {
                context.log("Sending push notification",
                 subscriptionEntity.RowKey);
                sendPushMessage(context, JSON.parse(subscriptionEntity.pushSubscription),
                 notifications[i].resourceData, subscriptionEntity.RowKey);
            });
        }
        context.res
            .status(202)
            .send("Received notifications. Push message(s) may have been sent.");
    } else if (!req.query.validationToken && req.body && req.body.id) {
        // Save subscription to storage
        saveSubscriptionToStorage(req, context);
    } else if (req.query.validationToken) {
        // Validate callback URL
        context.log("Received validation token " + req.query.validationToken);
        context.res
            .status(200)
            .send(req.query.validationToken);
    } else {

        context.log("Invalid request");
        context.res
            .status(400)
            .send();
    }
    
};

function saveSubscriptionToStorage(req, context) {
    context.log("Received subscription details ", req.body, context.bindings.storageTable);
    context.bindings.storageTable = [];
    var entity = {
        PartitionKey: PartitionKey,
        RowKey: req.body.id,
        clientState: req.body.clientState,
        changeType: req.body.changeType,
        resource: req.body.resource,
        expirationDateTime: req.body.expirationDateTime,
        pushSubscription: JSON.stringify(req.body.pushSubscription)
    };
    context.log("New entity = ", entity);
    insertOrUpdateEntity(entity, context);
}

function insertOrUpdateEntity(entity, context) {
    context.log("Inserting/merging subscription info -", entity.RowKey);
    let tableService = azure.createTableService(connectionString);

    tableService.insertOrMergeEntity('subscriptions', entity, (error, result, response) => {
        let res = {
            statusCode: error ? 400 : 200,
            body: null
        };
        context.log("Azure storage insertOrMerge response", error, response);
        context.done(null, res);
    });
}

function sendPushMessage(context, pushSubscription, data, subscriptionId) {

    context.log("Sending push message", subscriptionId);
    
    webpush.sendNotification(
        pushSubscription,
        JSON.stringify(data),
        options
    )
    .then(() => {
        context.log("Push message sent.");
    })
    .catch((err) => {
        if (err.statusCode === 410) {
            deleteSubscription(context, subscriptionId);
            context.log("Subscription expired and deleted.", subscriptionId);
        } else if (err.statusCode) {
            context.log("Push message could not be sent.",
             err.statusCode, err.message);
        } else {
            context.log("Push message could not be sent.", err.message);
        }
    });
}


function deleteSubscription(context, subscriptionId) {
    var id = subscriptionId;
    
    context.log("Deleting subscription", id);

    let tableService = azure.createTableService(connectionString);

    let item = {
        PartitionKey: PartitionKey,
        RowKey: id
    };

    tableService.deleteEntity('subscriptions', item, (error, response) => {
        let res = {
            statusCode: error ? 400 : 204,
            body: null
        };
        context.log("Azure storage delete response", error);
        context.done(null, res);
    });
}