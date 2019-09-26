import * as childProcess from 'child_process';
import * as crytpo from 'crypto';
import * as defaults from './defaults';
import * as fs from 'fs';
import { modifyManifestFile } from 'office-addin-manifest';
// import * as server from './server_new';
import * as util from 'util';
const readFileAsync = util.promisify(fs.readFile);
const writeFileAsync = util.promisify(fs.writeFile);
let secret: string = undefined;

export async function configureSSOApplication(manifestPath: string, ssoAppName: string) {
    const userJson = await logIntoAzure();
    if (userJson) {
        console.log('Login was successful!');
    }
    const applicationJson = await createNewApplicaiton(ssoAppName);
    writeApplicationJsonData(applicationJson, userJson);
    updateProjectManifest(manifestPath, applicationJson);
}

export async function logIntoAzure() {
    console.log('Opening browser for authentication to Azure. Enter valid Azure credentials');
    return await promiseExecuteCommand('az login --allow-no-subscriptions');
}
async function createNewApplicaiton(ssoAppName: string): Promise<Object> {
    try {
        console.log('Registering new application in Azure');
        let azRestNewAppCommand = await readFileAsync(defaults.azRestpCreateCommandPath, 'utf8');
        const re = new RegExp('{SSO-AppName}', 'g');
        secret = crytpo.randomBytes(20).toString('hex');
        azRestNewAppCommand = azRestNewAppCommand.replace(re, ssoAppName).replace('{SSO-Secret}', secret);
        const applicationJson: Object = await promiseExecuteCommand(azRestNewAppCommand, true /* setIdentifierUri */);
        if (applicationJson) {
            console.log('Application was successfully registered with Azure');
        }
        return applicationJson;
    } catch (err) {
        console.log('createNewApplication failed with error:' + err);
        return undefined;
    }
}

export async function promiseExecuteCommand(cmd: string, setIdentifierUrl: boolean = false): Promise<Object> {
    return new Promise((resolve, reject) => {
        try {
            childProcess.exec(cmd, async (err, stdout, stderr) => {
                let json: Object;
                const results = stdout;
                if (results !== '') {
                    json = JSON.parse(results);
                }
                if (setIdentifierUrl) {
                    await setIdentifierUri(json);
                    await grantAdminContent(json);
                    await setSignInAudience(json);
                }
                resolve(json);
            });
        } catch (err) {
            reject(err);
        }
    });
}

async function setIdentifierUri(applicationJson: any) {
    try {
        console.log('Setting identifierUri');
        let azRestCommand = await readFileAsync(defaults.setIdentifierUriCommmandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson.id).replace('<App_Id>', applicationJson.appId);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        console.log('setIdentifierUri failed with error:' + err);
    }
}

async function grantAdminContent(applicationJson: any) {
    try {
        console.log('Granting admin consent');
        let azRestCommand = await readFileAsync(defaults.grantAdminConsentCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_ID>', applicationJson.appId);
        console.log(`grant admin consent command is ${azRestCommand}`);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        console.log('grantAdminConsent failed with error:' + err);
    }
}

async function setSignInAudience(applicationJson: any) {
    try {
        console.log('Setting signin audience');
        let azRestCommand = await readFileAsync(defaults.setSigninAudienceCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson.id);
        console.log(`signing audeince command is ${azRestCommand}`);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        console.log('setSignInAudience failed with error:' + err);
    }
}

async function updateProjectManifest(manifestPath: string, applicationJson: any) {
    console.log('Updating manifest and source files with required values');
    try {
        // Update manifest with application guid and unique manifest id
        const manifestContent: string = await readFileAsync(manifestPath, 'utf8');
        const re = new RegExp('{application GUID here}', 'g');
        const updatedManifestContent: string = manifestContent.replace(re, applicationJson.appId);
        await writeFileAsync(manifestPath, updatedManifestContent);
        await modifyManifestFile(manifestPath, 'random');

    } catch (err) {
        throw new Error(`updating manifest failed with error: ${err}`);
    }
}

export function writeApplicationJsonData(applicationJson: any, userJson: any): void {
    if (fs.existsSync(defaults.ssoDataJsonFilePath) && fs.readFileSync(defaults.ssoDataJsonFilePath, 'utf8') !== '' && fs.readFileSync(defaults.ssoDataJsonFilePath, 'utf8') !== 'undefined') {
        const ssoJsonData = readSsoJsonData();
        ssoJsonData.ssoApplicationInstances[applicationJson.displayName] = { applicationId: String, applicationSecret: String, tenantId: String };
        ssoJsonData.ssoApplicationInstances[applicationJson.displayName].applicationId = applicationJson.appId;
        ssoJsonData.ssoApplicationInstances[applicationJson.displayName].tenantId = userJson[0].tenantId;
        ssoJsonData.ssoApplicationInstances[applicationJson.displayName].applicationSecret = secret;
        fs.writeFileSync(defaults.ssoDataJsonFilePath, JSON.stringify((ssoJsonData), null, 2));
    } else {
        let ssoJsonData = {};
        ssoJsonData[applicationJson.displayName] = applicationJson.displayName;
        ssoJsonData = { ssoApplicationInstances: ssoJsonData };
        ssoJsonData = { ssoApplicationInstances: { [applicationJson.displayName]: { ['applicationId']: applicationJson.appId, ['tenantId']: userJson[0].tenantId, ['applicationSecret']: secret } } };
        fs.writeFileSync(defaults.ssoDataJsonFilePath, JSON.stringify((ssoJsonData), null, 2));
    }
}

// export function modifySsoJsonData(applicationJson: any, userJson: any): void {
//     try {
//         if (fs.existsSync(defaults.ssoDataJsonFilePath) && fs.readFileSync(defaults.ssoDataJsonFilePath, 'utf8') !== '' && fs.readFileSync(defaults.ssoDataJsonFilePath, 'utf8') !== 'undefined') {
//             if (ssoApplicationExists(applicationJson.displayName)) {
//                 modifySsoJsonData(applicationJson, userJson);
//             } else {
//                 const ssoJsonData = readSsoJsonData();
//                 ssoJsonData.ssoApplicationInstances[applicationJson.displayName] = { applicationId: String, clientSecret: String, tenantId: String };
//                 ssoJsonData.ssoApplicationInstances[applicationJson.displayName].applicationId = applicationJson.appId;
//                 ssoJsonData.ssoApplicationInstances[applicationJson.displayName].tenantId = userJson[0].tenantId;
//                 ssoJsonData.ssoApplicationInstances[applicationJson.displayName].clientSecret = 'sso-secret';
//                 fs.writeFileSync(defaults.ssoDataJsonFilePath, JSON.stringify((ssoJsonData), null, 2));
//             }
//         } else {
//             let ssoJsonData = {};
//             ssoJsonData[groupName] = value;
//             ssoJsonData = { ssoApplicationInstances: ssoJsonData };
//             ssoJsonData = { ssoApplicationInstances: { [groupName]: { [property]: value } } };
//             fs.writeFileSync(defaults.ssoDataJsonFilePath, JSON.stringify((ssoJsonData), null, 2));
//         }
//     } catch (err) {
//         throw new Error(err);
//     }
// }

// /**
//  * Reads data from the usage data json config file
//  * @returns Parsed object from json file if it exists
//  */
export function readSsoJsonData(): any {
    if (fs.existsSync(defaults.ssoDataJsonFilePath)) {
        const jsonData = fs.readFileSync(defaults.ssoDataJsonFilePath, 'utf8');
        return JSON.parse(jsonData.toString());
    }
}

export function ssoApplicationExists(ssoAppName: string): boolean {
    if (fs.existsSync(defaults.ssoDataJsonFilePath) && fs.readFileSync(defaults.ssoDataJsonFilePath, 'utf8') !== '' && fs.readFileSync(defaults.ssoDataJsonFilePath, 'utf8') !== 'undefined') {
        const jsonData = readSsoJsonData();
        return Object.getOwnPropertyNames(jsonData.ssoApplicationInstances).includes(ssoAppName);
    }
    return false;
}

// export async function startSsoServer(ssoApplicationName: string): Promise<boolean> {
//     return new Promise<boolean>(async (resolve, reject) => {
//         try {
//             if (ssoApplicationExists(ssoApplicationName)) {
//                 const ssoApplicationData = readSsoJsonData();
//                 const serverOptions: server.ISSOAuthServiceOptions = {
//                     applicationId: ssoApplicationData.ssoApplicationInstances[ssoApplicationName].applicationId,
//                     tenantId: ssoApplicationData.ssoApplicationInstances[ssoApplicationName].tenantId,
//                     applicationSecret: ssoApplicationData.ssoApplicationInstances[ssoApplicationName].applicationSecret
//                 };
//                 const ssoAuthService: server.SSOAuthService = new server.SSOAuthService(serverOptions);
//                 ssoAuthService.startServer();
//                 resolve(true);
//             }
//             resolve(false);

//         } catch (err) {
//             reject(false);
//         }
//     });
// }
