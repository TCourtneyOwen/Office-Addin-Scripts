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

export function writeApplicationData(applicationId, tenantId) {
    let isTypecript: boolean = false;
    try {
        // Update .ENV file
        if (fs.existsSync(defaults.ssoDataFilePath)) {
            const appData = fs.readFileSync(defaults.ssoDataFilePath, 'utf8');
            const updatedAppData = appData.replace('CLIENT_ID=', `CLIENT_ID=${applicationId}`).replace('TENANT_ID=', `TENANT_ID=${tenantId}`);
            fs.writeFileSync(defaults.ssoDataFilePath, updatedAppData);
        } else {
            throw new Error(`${defaults.ssoDataFilePath} does not exist`)
        }
    } catch (err) {
        throw new Error(`Unable to write SSO application data to ${defaults.ssoDataFilePath}. \n${err}`);
    }

    try {
        // Update fallbackAuthDialog.js

        if (fs.existsSync(defaults.fallbackAuthDialogTypescriptFilePath)) {
            isTypecript = true;
            const srcFile = fs.readFileSync(defaults.fallbackAuthDialogTypescriptFilePath, 'utf8');
            const updatedSrcFile = srcFile.replace('{application GUID here}', applicationId);
            fs.writeFileSync(defaults.fallbackAuthDialogTypescriptFilePath, updatedSrcFile);
        } else if (fs.existsSync(defaults.fallbackAuthDialogJavascriptFilePath)) {
            const srcFile = fs.readFileSync(defaults.fallbackAuthDialogJavascriptFilePath, 'utf8');
            const updatedSrcFile = srcFile.replace('{application GUID here}', applicationId);
            fs.writeFileSync(defaults.fallbackAuthDialogJavascriptFilePath, updatedSrcFile);
        }
         else {
            throw new Error(`${isTypecript ? defaults.fallbackAuthDialogTypescriptFilePath : defaults.fallbackAuthDialogJavascriptFilePath} does not exist`)
        }
    } catch (err) {
        throw new Error(`Unable to write SSO application data to ${isTypecript ? defaults.fallbackAuthDialogTypescriptFilePath : defaults.fallbackAuthDialogJavascriptFilePath}. \n${err}`);
    }
}
