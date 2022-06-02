const BoxSDK = require('box-node-sdk');
const path = require("path");
const boxEnterpriseSdkProps = require('./configs/box_enterprise_sdk_properties.json');
const fs = require('fs');

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

    context.log("*** INSIDE : BOX CLIENT API DOWNLOAD FILE SERVICE ***")
    try {

        var folderId = req.body.folderId;
        var filesToBeDownloaded = req.body.listOfFileNames;

        var itemsInFolder = await listItemsInFolder(folderId);
        var entries = itemsInFolder.entries;
        var fileContentArrayOfJson = [];

        for (var i = 0; i < entries.length; i++) {

            if (entries[i].type == 'file' && filesToBeDownloaded.includes(entries[i].name)) {
                console.log("** " + itemsInFolder.entries[i].id);
                fileContentArrayOfJson.push(await downloadFile(itemsInFolder.entries[i].id))
            }

        }

        context.res = {
            body: fileContentArrayOfJson
        };
    } catch (error) {
        console.log(" ** ERROR :: ** ",error);
        context.res = {
            body: error
        };

    }
}

const downloadFile = (fileId) => {

    var fileContentJson = {};

    return new Promise((resolve, reject) => {

        serviceAccountClient.files.getReadStream(fileId, null)
            .then(stream => {

                var finalData = '';
                var bufs = [];

                stream.on("data", (chunk) => {
                    bufs.push(chunk);
                });

                stream.on("end", () => {
                    finalData = Buffer.concat(bufs)

                    fileContentJson["$content-type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    fileContentJson["$content"] = new Buffer.from(finalData).toString('base64');

                    resolve(fileContentJson);
                })

            })
            .catch(error => reject(error));

    })
}

const listItemsInFolder = (folderId) => {
    return new Promise((resolve, reject) => {

        serviceAccountClient.folders.getItems(folderId, null)
            .then(items => resolve(items))
            .catch(error => reject(error));

    })
}