import * as defaults from './defaults';
import { execSync } from "child_process";
import * as fs from 'fs';
import * as os from 'os';

export function addSecretToCredentialStore(ssoAppName: string, secret: string): void {
    try {
        switch (process.platform) {
            case "win32":
                console.log(`Adding application secret for ${ssoAppName} to Windows Credential Store`);
                const addSecretToWindowsStoreCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.addSecretCommandPath}" "${ssoAppName}" "${os.userInfo().username}" "${secret}"`;
                execSync(addSecretToWindowsStoreCommand, { stdio: "pipe" });
                break;
            case "darwin":
                console.log(`Adding application secret for ${ssoAppName} to Mac OS Keychain`);
                const addSecretToMacStoreCommand = `sudo security add-generic-password -a ${os.userInfo().username} -s ${ssoAppName} -w ${secret}`;
                execSync(addSecretToMacStoreCommand, { stdio: "pipe" });
                break;
            default:
                throw new Error(`Platform not supported: ${process.platform}`);
        }
    } catch (err) {
        throw new Error(`Unable to add secret for ${ssoAppName} to Windows Credential Store. \n${err}`);
    }
}

export function getSecretFromCredentialStore(ssoAppName: string): string {
    try {
        switch (process.platform) {
            case "win32":
                console.log(`Getting application secret for ${ssoAppName} from Windows Credential Store`);
                const getSecretFromWindowsStoreCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.getSecretCommandPath}" "${ssoAppName}" "${os.userInfo().username}"`;
                return execSync(getSecretFromWindowsStoreCommand, { stdio: "pipe" }).toString();
            case "darwin":
                console.log(`Getting application secret for ${ssoAppName} from Mac OS Keychain`);
                const getSecretFromMacStoreCommand = `sudo security find-generic-password -a ${os.userInfo().username} -s ${ssoAppName} -w`;
                return execSync(getSecretFromMacStoreCommand, { stdio: "pipe" }).toString();;
            default:
                throw new Error(`Platform not supported: ${process.platform}`);
        }

    } catch (err) {
        throw new Error(`Unable to retrieve secret for ${ssoAppName} to Windows Credential Store. \n${err}`);
    }
}

export function writeApplicationJsonData(applicationJson: any, userJson: any) {
    try {
        if (fs.existsSync(defaults.ssoDataJsonFilePath) && fs.readFileSync(defaults.ssoDataJsonFilePath, 'utf8') !== '' && fs.readFileSync(defaults.ssoDataJsonFilePath, 'utf8') !== 'undefined') {
            if (ssoApplicationExists(applicationJson.displayName)) {
                modifySsoJsonData(applicationJson, userJson);
            } else {
                const ssoJsonData = readSsoJsonData();
                ssoJsonData.ssoApplicationInstances[applicationJson.displayName] = { applicationId: String, applicationSecret: String, tenantId: String };
                ssoJsonData.ssoApplicationInstances[applicationJson.displayName].applicationId = applicationJson.appId;
                ssoJsonData.ssoApplicationInstances[applicationJson.displayName].tenantId = userJson[0].tenantId;
                fs.writeFileSync(defaults.ssoDataJsonFilePath, JSON.stringify((ssoJsonData), null, 2));
            }
        } else {
            let ssoJsonData = {};
            ssoJsonData[applicationJson.displayName] = applicationJson.displayName;
            ssoJsonData = { ssoApplicationInstances: ssoJsonData };
            ssoJsonData = { ssoApplicationInstances: { [applicationJson.displayName]: { ['applicationId']: applicationJson.appId, ['tenantId']: userJson[0].tenantId } } };
            fs.writeFileSync(defaults.ssoDataJsonFilePath, JSON.stringify((ssoJsonData), null, 2));
        }
    } catch (err) {
        throw new Error(`Unable to write SSO application data to ${defaults.ssoDataJsonFilePath}. \n${err}`);
    }
}

async function modifySsoJsonData(applicationJson: any, userJson: any): Promise<void> {
    try {
        const ssoJsonData = readSsoJsonData();
        if (ssoJsonData && ssoApplicationExists(applicationJson.displayName)) {
            ssoJsonData.ssoApplicationInstances[applicationJson.displayName].applicationId = applicationJson.appId;
            ssoJsonData.ssoApplicationInstances[applicationJson.displayName].tenantId = userJson[0].tenantId;
            fs.writeFileSync(defaults.ssoDataJsonFilePath, JSON.stringify((ssoJsonData), null, 2));
        } else {
            throw new Error(`SSO application ${applicationJson.displayName} doesn't exist in settings file`);
        }
    } catch (err) {
        throw new Error(`Unable to modify ${defaults.ssoDataJsonFilePath}. \n${err}`);
    }
}

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
