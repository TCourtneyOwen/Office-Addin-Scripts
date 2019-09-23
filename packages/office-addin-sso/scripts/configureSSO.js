const ssoAppName = process.argv[2];
const manifestPath = process.argv[3];
const childProcess = require("child_process");
const fs = require("fs");
const officeAddinManifest = require("office-addin-manifest");
const util = require("util");
const readFileAsync = util.promisify(fs.readFile);
const writeFileAsync = util.promisify(fs.writeFile);

async function configureSSO(manifestPath, ssoAppName) {
    const userJson = await logIntoAzure();
    if (userJson) {
        console.log("Login was successful!");
    }
    const applicationJson = await createNewApplicaiton(manifestPath, ssoAppName);
    updateProjectFiles(manifestPath, applicationJson, userJson);
}

async function logIntoAzure() {
    console.log("Opening browser for authentication to Azure. Enter valid Azure credentials");
    return await promiseExecuteCommand("az login --allow-no-subscriptions");
}
async function createNewApplicaiton(ssoAppName) {
    let azRestCommand = await readFileAsync('./scripts/azRestAppCreateCmd.txt', 'utf8');
    azRestCommand = azRestCommand.replace("{SSO-AppName}", ssoAppName);
    const applicationJson = await promiseExecuteCommand(azRestCommand, true /* setIdentifierUri */);
    if (applicationJson) {
        console.log("Application was successfully registered with Azure");
    }
    return applicationJson;
}

async function promiseExecuteCommand(cmd, setIdentifierUrl = false) {
    return new Promise(function (resolve, reject) {
        try {
            childProcess.exec(cmd, async function (err, stdout, stderr) {
                let json;
                const results = stdout;
                if (results !== "") {
                    json = JSON.parse(results);
                }
                if (setIdentifierUrl) {
                    await setIdentifierUri(json);
                }
                resolve(json);
            });
        } catch (err) {
            reject(err);
        }
    });
}

async function setIdentifierUri(applicationJson) {
    let azRestCommand = await readFileAsync('./scripts/azRestSetIdentifierUri.txt', 'utf8' );
    azRestCommand = azRestCommand.replace("<App_Object_ID>", applicationJson.id).replace("<App_Id>", applicationJson.appId);
    await promiseExecuteCommand(azRestCommand);
}

async function updateProjectFiles(manifestPath, applicationJson, userJson) {
    console.log("Updating manifest and source files with required values");
    try {
        // Update manifest with application guid and unique manifest id
        const manifestContent = await readFileAsync(manifestPath, 'utf8');
        const re = new RegExp("{application GUID here}", 'g');
        const updatedManifestContent = manifestContent.replace(re, applicationJson.appId);
        await writeFileAsync(manifestPath, updatedManifestContent);
        officeAddinManifest.modifyManifestFile(manifestPath, 'random');

        // Update source files
        // const serverSource = "./src/server.ts";
        // const serverSourceContent = await readFileAsync(serverSource, 'utf8');
        // const updatedServerSourceContent = serverSourceContent.replace("{client GUID}", applicationJson.appId).replace("{audience GUID}", applicationJson.appId).replace("{O365 tenant GUID}", userJson[0].tenantId);
        // await writeFileAsync(serverSource, updatedServerSourceContent);
        // console.log("Manifest and source files successfully updated!");
    } catch (err) {
        throw new Error(`updating project files failed with error: ${err}`);
    }
}