import * as childProcess from 'child_process';
import * as defaults from './defaults';
import * as fs from 'fs';
import { addSecretToCredentialStore, writeApplicationData } from './ssoDataSettings';
import { ManifestInfo, readManifestFile } from 'office-addin-manifest';
require('dotenv').config();

export async function configureSSOApplication(manifestPath: string) {
    // Check to see if Azure CLI is installed.  If it isn't installed then install it
    const cliInstalled = await azureCliInstalled();
    if(!cliInstalled) {
        console.log("Azure CLI is not installed.  Installing now before proceeding");
        await installAzureCli();
        if (process.platform === "win32") {
            console.log('Please close your command shell, reopen and run configure-sso again.  This is neccessary to register the path to the Azure CLI');
        }
        return;
    }

    const userJson: Object = await logIntoAzure();
    if (userJson) {
        console.log('Login was successful!');
        const manifestInfo: ManifestInfo = await readManifestFile(manifestPath);

        // Register application
        const applicationJson: any = await createNewApplication(manifestInfo.displayName);

        // Write application data to manifest and .ENV file
        writeApplicationData(applicationJson.appId, manifestPath);

        // Log out of Azure
        await logoutAzure();

        // Output application definition to console
        console.log(applicationJson);
    }
    else {
        throw new Error(`Login to Azure did not succeed.`);
    }
}

async function createNewApplication(ssoAppName: string): Promise<Object> {
    try {
        console.log('Registering new application in Azure');
        let azRestCommand = await fs.readFileSync(defaults.azRestApppCreateCommandPath, 'utf8');
        const reName = new RegExp('{SSO-AppName}', 'g');
        const rePort = new RegExp('{PORT}', 'g');
        azRestCommand = azRestCommand.replace(reName, ssoAppName).replace(rePort, process.env.PORT);

        const applicationJson: Object = await promiseExecuteCommand(azRestCommand, true /* returnJson */);

        if (applicationJson) {
            console.log('Application was successfully registered with Azure');
            // Set application IdentifierUri
            await setIdentifierUri(applicationJson);

            // Set application sign-in audience
            await setSignInAudience(applicationJson);

            // Grant admin consent for application
            await grantAdminContent(applicationJson);

            // Create an application secret and add to the credential store
            const secret: string = await setApplicationSecret(applicationJson);
            console.log(`App secret is ${secret}`);
            addSecretToCredentialStore(ssoAppName, secret);

            return applicationJson;
        } else {
            console.log("Failed to register application");
            return undefined;
        }

    } catch (err) {
        throw new Error(`Unable to register new application ${ssoAppName}. \n${err}`);
    }
}

async function applicationReady(applicationJson: any): Promise<boolean> {
    try {
        const azRestCommand: string = `az ad app show --id ${applicationJson.appId}`
        const appJson: Object = await promiseExecuteCommand(azRestCommand, true /* returnJson */, true /* expectError */);
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
        let counter: number = 0;
        while (appReady === false && counter <= 20) {
            appReady = await applicationReady(applicationJson);
            counter++;
        }
        const azRestCommand: string = `az ad app permission admin-consent --id ${applicationJson.appId}`;
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        throw new Error(`Unable to set grant admin consent for ${applicationJson.displayName}. \n${err}`);
    }
}

export async function azureCliInstalled(): Promise<boolean> {
    try {
        switch (process.platform) {
            case "win32":
                const appsInstalledWindowsCommand: string = `powershell -ExecutionPolicy Bypass -File "${defaults.getInstalledAppsPath}"`;
                const appsWindows: any = await promiseExecuteCommand(appsInstalledWindowsCommand);
                return appsWindows.filter(app => app.DisplayName === 'Microsoft Azure CLI').length > 0
            case "darwin":
                const appsInstalledMacCommand = 'brew list';
                const appsMac: Object | string = await promiseExecuteCommand(appsInstalledMacCommand, false /* returnJson */);
                return appsMac.hasOwnProperty('azure-cli');
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

export async function logIntoAzure(): Promise<Object> {
    console.log('Opening browser for authentication to Azure. Enter valid Azure credentials');
    return await promiseExecuteCommand('az login --allow-no-subscriptions');
}

async function logoutAzure(): Promise<Object> {
    console.log('Logging out of Azure now');
    return await promiseExecuteCommand('az logout');
}

export async function promiseExecuteCommand(cmd: string, returnJson: boolean = true, expectError: boolean = false): Promise<Object | string> {
    return new Promise((resolve, reject) => {
        try {
            childProcess.exec(cmd, async (err, stdout, stderr) => {
                if (err && !expectError) {
                    console.log(stderr);
                    reject(stderr);
                }
                
                let results = stdout;
                if (results !== '' && returnJson) {
                    results = JSON.parse(results);
                }
                resolve(results);
            });
        } catch (err) {
            reject(err);
        }
    });
}

async function setApplicationSecret(applicationJson: any): Promise<string> {
    try {
        console.log('Setting identifierUri');
        let azRestCommand: string = await fs.readFileSync(defaults.azAddSecretCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson.id);
        const secretJson: any = await promiseExecuteCommand(azRestCommand);
        return secretJson.secretText;
    } catch (err) {
        throw new Error(`Unable to set identifierUri for ${applicationJson.displayName}. \n${err}`);
    }
}

async function setIdentifierUri(applicationJson: any) {
    try {
        console.log('Setting identifierUri');
        let azRestCommand: string = await fs.readFileSync(defaults.setIdentifierUriCommmandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson.id).replace('<App_Id>', applicationJson.appId).replace('{PORT}', process.env.PORT);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        throw new Error(`Unable to set identifierUri for ${applicationJson.displayName}. \n${err}`);
    }
}

async function setSignInAudience(applicationJson: any) {
    try {
        console.log('Setting signin audience');
        let azRestCommand: string = await fs.readFileSync(defaults.setSigninAudienceCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson.id);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        throw new Error(`Unable to set signInAudience for ${applicationJson.displayName}. \n${err}`);
    }
}
