import * as childProcess from 'child_process';
import * as defaults from './defaults';
import * as fs from 'fs';
import * as passwordGenerator from "generate-password";
import { modifyManifestFile } from 'office-addin-manifest';
import { addSecretToCredentialStore, writeApplicationJsonData } from './ssoDataSettings';

export async function configureSSOApplication(manifestPath: string, ssoAppName: string) {
    // Check to see if Azure CLI is installed.  If it isn't installed then install it
    const cliInstalled = await azureCliInstalled();
    if (!cliInstalled) {
        console.log("Azure CLI is not installed.  Installing now before proceeding");
        await installAzureCli();
    }

    const userJson = await logIntoAzure();
    if (userJson) {
        console.log('Login was successful!');
        const secret = passwordGenerator.generate({ length: 32, numbers: true, uppercase: true, strict: true });
        const applicationJson: any = await createNewApplication(ssoAppName, secret);
        writeApplicationJsonData(applicationJson, userJson);
        addSecretToCredentialStore(ssoAppName, secret);
        updateProjectManifest(manifestPath, applicationJson.appId);
        console.log("Outputting Azure application info:\n");
        console.log(applicationJson);
    }
    else {
        throw new Error(`Login to Azure did not succeed.`);
    }
}

async function createNewApplication(ssoAppName: string, secret: string): Promise<Object> {
    try {
        console.log('Registering new application in Azure');
        let azRestNewAppCommand = await fs.readFileSync(defaults.azRestpCreateCommandPath, 'utf8');
        const re = new RegExp('{SSO-AppName}', 'g');
        azRestNewAppCommand = azRestNewAppCommand.replace(re, ssoAppName).replace('{SSO-Secret}', secret);
        const applicationJson: Object = await promiseExecuteCommand(azRestNewAppCommand, true /* returnJson */, true /* configureSSO */);
        if (applicationJson) {
            console.log('Application was successfully registered with Azure');
        }
        return applicationJson;
    } catch (err) {
        throw new Error(`Unable to register new application ${ssoAppName}. \n${err}`);
    }
}

async function applicationReady(applicationJson: any): Promise<boolean> {
    try {
        let azRestCommand = await fs.readFileSync(defaults.getApplicationInfoCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_ID>', applicationJson.appId);
        const appJson: any = await promiseExecuteCommand(azRestCommand);
        return appJson !== "";
    } catch (err) {
        throw new Error(`Unable to get application info for ${applicationJson.displayName}. \n${err}`);
    }
}

async function grantAdminContent(applicationJson: any) {
    try {
        console.log('Granting admin consent');
        // Check to see if the application is available before granting admin consent
        let appReady: boolean = false;
        while (appReady === false) {
            appReady = await applicationReady(applicationJson);
        }
        let azRestCommand = fs.readFileSync(defaults.grantAdminConsentCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_ID>', applicationJson.appId);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        throw new Error(`Unable to set grant admin consent for ${applicationJson.displayName}. \n${err}`);
    }
}

export async function azureCliInstalled(): Promise<boolean> {
    try {
        switch (process.platform) {
            case "win32":
                const appsInstalledWindowsCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.getInstalledAppsPath}"`;
                const appsWindows = await promiseExecuteCommand(appsInstalledWindowsCommand);
                return appsWindows.filter(app => app.DisplayName === 'Microsoft Azure CLI').length > 0
            case "darwin":
                const appsInstalledMacCommand = 'brew list';
                const appsMac: string = await promiseExecuteCommand(appsInstalledMacCommand, false /* returnJson */);
                return appsMac.includes('azure-cli');;;
            default:
                throw new Error(`Platform not supported: ${process.platform}`);
        }
    } catch (err) {
        throw new Error(`Unable to install Azure CLI. \n${err}`);
    }
}

export async function installAzureCli() {
    try {
        console.log("Downloading and installing Azure CLI - this could take a minute or so");
        switch (process.platform) {
            case "win32":
                const windowsCliInstallCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.azCliInstallCommandPath}"`;
                await promiseExecuteCommand(windowsCliInstallCommand, false /* returnJson */);
                break;
            case "darwin": // macOS
                const macCliInstallCommand = 'brew update && brew install azure-cli';
                await promiseExecuteCommand(macCliInstallCommand, false /* returnJson */);
                break;
            default:
                throw new Error(`Platform not supported: ${process.platform}`);
        }
    } catch (err) {
        throw new Error(`Unable to install Azure CLI. \n${err}`);
    }
}

export async function logIntoAzure() {
    console.log('Opening browser for authentication to Azure. Enter valid Azure credentials');
    return await promiseExecuteCommand('az login --allow-no-subscriptions');
}


export async function promiseExecuteCommand(cmd: string, returnJson: boolean = true, configureSSO: boolean = false): Promise<any> {
    return new Promise((resolve, reject) => {
        try {
            childProcess.exec(cmd, async (err, stdout, stderr) => {
                let results = stdout;
                if (results !== '' && returnJson) {
                    results = JSON.parse(results);
                }
                if (configureSSO) {
                    await setIdentifierUri(results);
                    await setSignInAudience(results);
                    await grantAdminContent(results);
                }
                resolve(results);
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
