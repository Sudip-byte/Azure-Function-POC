const BoxSDK = require('box-node-sdk');
const path = require("path");
const boxEnterpriseSdkProps = require('./configs/box_enterprise_sdk_properties.json');

var privateKey;
if (process.env.BOX_privateKey) {
    privateKey = process.env.BOX_privateKey.replace(/\\n/g, '\n')
}
var boxParam = {
    clientID: process.env.BOX_clientID || boxEnterpriseSdkProps["box"]["clientID"],
    clientSecret: process.env.BOX_clientSecret || boxEnterpriseSdkProps["box"]["clientSecret"],
    appAuth: {
        keyID: process.env.BOX_keyID || boxEnterpriseSdkProps["box"]["appAuth"]["keyID"],
        privateKey: privateKey || boxEnterpriseSdkProps["box"]["appAuth"]["privateKey"],
        passphrase: process.env.BOX_passphrase || boxEnterpriseSdkProps["box"]["appAuth"]["passphrase"],
        enterprise: process.env.BOX_enterprise || boxEnterpriseSdkProps["box"]["appAuth"]["enterprise"]
    }
}

var sdk = new BoxSDK(boxParam);
var serviceAccountClient = sdk.getAppAuthClient('enterprise', process.env.BOX_enterprise || boxEnterpriseSdkProps["box"]["appAuth"]["enterprise"]);

module.exports = async function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');

    context.log("*** INSIDE : BOX CLIENT API FOLDER RENAME SERVICE ***")
    try {
        var rootFolderId = req.body.rootFolderId;
        var folderName = req.body.folderName;

        var responseMessage = await createFolder(rootFolderId, folderName);

        context.res = {
            body: responseMessage
        };
    } catch (error) {
        console.log(" ** ERROR :: ** ",error);
        context.res = {
            body: error
        };

    }
}

const createFolder = (rootFolderId, folderName) => {
    return new Promise((resolve, reject) => {

        serviceAccountClient.folders.create(rootFolderId, folderName)
            .then(folderInfo => resolve(folderInfo))
            .catch(error => reject(error));

    })
}