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

    context.log("*** INSIDE : BOX CLIENT API BOX FILE UPLOAD SERVICE ***")
    try {
        var uploadFolderId = req.body.uploadFolderId;
        var fileExtension = req.body.dataStreamBase64["$content-type"];
        var uploadFileName = req.body.uploadFileName + getFileExtension(fileExtension);
        var dataStreamBase64X = req.body.dataStreamBase64["$content"];

        var buffer = new Buffer(dataStreamBase64X, 'base64');

        var itemsInFolder = await listItemsInFolder(uploadFolderId);
        var entries = itemsInFolder.entries;

        var fileExisitsAlreadyFlag = 0;
        var existingFileId = "";

        for (var i = 0; i < entries.length; i++) {
            if (entries[i].type == 'file' && uploadFileName == entries[i].name) {
                context.log(" ** FILE WITH SAME NAME EXISTS ** ");
                fileExisitsAlreadyFlag = 1;
                existingFileId = entries[i].id;
                break;
            }

        }

        if (fileExisitsAlreadyFlag == 1) {
            var responseMessage = await uploadNewFileVersion(existingFileId, buffer);
            console.log(" ** FILE ID ** " + responseMessage.entries[0].id + " ** NAME ** : " + responseMessage.entries[0].name)

            context.res = {
                body: responseMessage
            };
        }
        else {
            var responseMessage = await uploadFile(uploadFolderId, uploadFileName, buffer);
            console.log("** FILE ID UPLOADED ** : " + responseMessage.entries[0].id + " ** NAME ** : " + responseMessage.entries[0].name)
            context.res = {
                body: responseMessage
            };

        }

    } catch (error) {
        console.log(" ** ERROR :: ** ",error);
        context.res = {
            body: error
        };

    }
}

/* GET THE FILE EXTENSION AS PER MIME-TYPE */
const getFileExtension = (contentType) => {

    var mimetypeFileExtensionMap = new Map();

    mimetypeFileExtensionMap.set("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx");
    mimetypeFileExtensionMap.set("application/vnd.ms-excel.sheet.binary.macroEnabled.12", ".xlsb");
    mimetypeFileExtensionMap.set("application/vnd.ms-excel", ".xls");
    mimetypeFileExtensionMap.set("application/vnd.ms-excel.sheet.macroEnabled.12", ".xlsm");

    return mimetypeFileExtensionMap.get(contentType);

}

/* UPLOAD THE FILE USING BOX CLIENT SERVICE CLIENT DETAILS */
const uploadFile = (uploadFolderId, uploadFileName, dataStream) => {
    return new Promise((resolve, reject) => {

        serviceAccountClient.files.uploadFile(uploadFolderId, uploadFileName, dataStream)
            .then(file => {
                resolve(file)
            })
            .catch(error => reject(error));

    })

}

const uploadNewFileVersion = (uploadFolderId, dataStream) => {
    return new Promise((resolve, reject) => {

        serviceAccountClient.files.uploadNewFileVersion(uploadFolderId, dataStream)
            .then(file => {
                resolve(file)
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