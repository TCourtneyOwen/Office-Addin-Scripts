import * as childProcess from 'child_process';
import * as defaults from './defaults';
import * as fs from 'fs';
import * as passwordGenerator from "password-generator";
import { modifyManifestFile } from 'office-addin-manifest';
import { writeApplicationJsonData } from './ssoDataSetttings';

export async function configureSSOApplication(manifestPath: string, ssoAppName: string) {
    console.log(`local path is ${__dirname}`);
    const userJson = await logIntoAzure();
    if (userJson) {
        console.log('Login was successful!');
    }
    const secret = passwordGenerator(32, true);
    const applicationJson: any = await createNewApplication(ssoAppName, secret);
    writeApplicationJsonData(applicationJson, userJson, secret);
    updateProjectManifest(manifestPath, applicationJson.appId);
}

async function grantAdminContent(applicationJson: any) {
    try {
        console.log('Granting admin consent');
        let azRestCommand = fs.readFileSync(defaults.grantAdminConsentCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_ID>', applicationJson.appId);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        throw new Error(`Unable to set grant admin consent for ${applicationJson.displayName}. \n${err}`);
    }
}

export async function logIntoAzure() {
    console.log('Opening browser for authentication to Azure. Enter valid Azure credentials');
    return await promiseExecuteCommand('az login --allow-no-subscriptions');
}
async function createNewApplication(ssoAppName: string, secret: string): Promise<Object> {
    try {
        console.log('Registering new application in Azure');
        let azRestNewAppCommand = await fs.readFileSync(defaults.azRestpCreateCommandPath, 'utf8');
        const re = new RegExp('{SSO-AppName}', 'g');
        secret = passwordGenerator(32, false);
        azRestNewAppCommand = azRestNewAppCommand.replace(re, ssoAppName).replace('{SSO-Secret}', secret);
        const applicationJson: Object = await promiseExecuteCommand(azRestNewAppCommand, true /* configureSSO */);
        if (applicationJson) {
            console.log('Application was successfully registered with Azure');
        }
        return applicationJson;
    } catch (err) {
        throw new Error(`Unable to register new application ${ssoAppName}. \n${err}`);
    }
}

export async function promiseExecuteCommand(cmd: string, configureSSO: boolean = false): Promise<Object> {
    return new Promise((resolve, reject) => {
        try {
            childProcess.exec(cmd, async (err, stdout, stderr) => {
                let json: Object;
                const results = stdout;
                if (results !== '') {
                    json = JSON.parse(results);
                }
                if (configureSSO) {
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
        let azRestCommand = await fs.readFileSync(defaults.setIdentifierUriCommmandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson.id).replace('<App_Id>', applicationJson.appId);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        throw new Error(`Unable to set identifierUri for ${applicationJson.displayName}. \n${err}`);
    }
}

async function setSignInAudience(applicationJson: any) {
    try {
        console.log('Setting signin audience');
        let azRestCommand = await fs.readFileSync(defaults.setSigninAudienceCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson.id);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        throw new Error(`Unable to set signInAudience for ${applicationJson.displayName}. \n${err}`);
    }
}

async function updateProjectManifest(manifestPath: string, applicationId: string) {
    console.log('Updating manifest with application ID');
    try {
        // Update manifest with application guid and unique manifest id
        const manifestContent: string = await fs.readFileSync(manifestPath, 'utf8');
        const re = new RegExp('{application GUID here}', 'g');
        const updatedManifestContent: string = manifestContent.replace(re, applicationId);
        await fs.writeFileSync(manifestPath, updatedManifestContent);
        await modifyManifestFile(manifestPath, 'random');

    } catch (err) {
        throw new Error(`Unable to update ${manifestPath}. \n${err}`);
    }
}
