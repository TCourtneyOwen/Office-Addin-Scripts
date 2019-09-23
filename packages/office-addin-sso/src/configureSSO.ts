import * as childProcess from 'child_process';
import * as defaults from './defaults';
import * as fs from 'fs';
import { modifyManifestFile } from 'office-addin-manifest';
import * as util from 'util';
const readFileAsync = util.promisify(fs.readFile);
const writeFileAsync = util.promisify(fs.writeFile);

export async function configureSSO(manifestPath: string, ssoAppName: string) {
    const userJson = await logIntoAzure();
    if (userJson) {
        console.log('Login was successful!');
    }
    const applicationJson = await createNewApplicaiton(ssoAppName);
    updateProjectFiles(manifestPath, applicationJson, userJson);
}

export async function logIntoAzure() {
    console.log('Opening browser for authentication to Azure. Enter valid Azure credentials');
    return await promiseExecuteCommand('az login --allow-no-subscriptions');
}
async function createNewApplicaiton(ssoAppName: string): Promise<Object> {
    let azRestNewAppCommand = await readFileAsync('./scripts/azRestAppCreateCmd.txt', 'utf8');
    azRestNewAppCommand = azRestNewAppCommand.replace('{SSO-AppName}', ssoAppName);
    const applicationJson: Object = await promiseExecuteCommand(azRestNewAppCommand, true /* setIdentifierUri */);
    if (applicationJson) {
        console.log('Application was successfully registered with Azure');
    }
    return applicationJson;
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
                }
                resolve(json);
            });
        } catch (err) {
            reject(err);
        }
    });
}

async function setIdentifierUri(applicationJson: any) {
    let azRestCommand = await readFileAsync('./scripts/azRestSetIdentifierUri.txt', 'utf8');
    azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson.id).replace('<App_Id>', applicationJson.appId);
    await promiseExecuteCommand(azRestCommand);
}

async function updateProjectFiles(manifestPath: string, applicationJson: any, userJson: any) {
    console.log('Updating manifest and source files with required values');
    try {
        // Update manifest with application guid and unique manifest id
        const manifestContent: string = await readFileAsync(manifestPath, 'utf8');
        const re = new RegExp('{application GUID here}', 'g');
        const updatedManifestContent: string = manifestContent.replace(re, applicationJson.appId);
        await writeFileAsync(manifestPath, updatedManifestContent);
        await modifyManifestFile(manifestPath, 'random');

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

export async function startSsoServer(applicationId: string): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
        try {
            const userJson = await logIntoAzure();
            if (userJson) {
                resolve(true);
            }
            resolve(false);

        } catch (err) {
            reject(false);
        }
    });
}

export function writeApplicationJsonData(applicationJson: any, userJson: any): void {
    if (fs.existsSync(defaults.ssoDataJsonFilePath) && fs.readFileSync(defaults.ssoDataJsonFilePath, 'utf8') !== '' && fs.readFileSync(defaults.usageDataJsonFilePath, 'utf8') !== 'undefined') {
            const ssoJsonData = readSsoJsonData();
            ssoJsonData.ssoInstances[applicationName] = { usageDataLevel: String };
            usageDataJsonData.usageDataInstances[groupName].usageDataLevel = level;
            fs.writeFileSync(defaults.usageDataJsonFilePath, JSON.stringify((usageDataJsonData), null, 2));
        }
    } else {
        let usageDataJsonData = {};
        usageDataJsonData[groupName] = level;
        usageDataJsonData = { usageDataInstances: usageDataJsonData };
        usageDataJsonData = { usageDataInstances: { [groupName]: { ['usageDataLevel']: level } } };
        fs.writeFileSync(defaults.usageDataJsonFilePath, JSON.stringify((usageDataJsonData), null, 2));
    }
}

/**
 * Allows developer to add or modify a specific property to the group
 * @param groupName Group name of property
 * @param property Property that will be created or modified
 * @param value Property's value that will be assigned
 */
export function modifySsoJsonData(applicationJson: any, userJson: any): void {
    try {
        if (fs.existsSync(defaults.ssoDataJsonFilePath) && fs.readFileSync(defaults.ssoDataJsonFilePath, 'utf8') !== '' && fs.readFileSync(defaults.ssoDataJsonFilePath, 'utf8') !== 'undefined') {
            if (ssoApplicationExists(applicationJson.displayName)) {
                modifySsoJsonData(applicationJson, userJson);
            } else {
                const ssoJsonData = readSsoJsonData();
                ssoJsonData.ssoApplicationInstances[applicationJson.displayName] = { applicationId: String, clientSecret: String, tenantId: String };
                ssoJsonData.ssoApplicationInstances[applicationJson.displayName].applicationId = applicationJson.appId;
                ssoJsonData.ssoApplicationInstances[applicationJson.displayName].tenantId = userJson[0].tenantId;
                ssoJsonData.ssoApplicationInstances[applicationJson.displayName].clientSecret = 'sso-secret';
                fs.writeFileSync(defaults.ssoDataJsonFilePath, JSON.stringify((ssoJsonData), null, 2));
            }
        } else {
            let usageDataJsonData = {};
            usageDataJsonData[groupName] = value;
            usageDataJsonData = { usageDataInstances: usageDataJsonData };
            usageDataJsonData = { usageDataInstances: { [groupName]: { [property]: value } } };
            fs.writeFileSync(defaults.usageDataJsonFilePath, JSON.stringify((usageDataJsonData), null, 2));
        }
    } catch (err) {
        throw new Error(err);
    }
}
/**
 * Reads data from the usage data json config file
 * @returns Parsed object from json file if it exists
 */
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
