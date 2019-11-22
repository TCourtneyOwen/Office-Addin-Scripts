// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides the writing and retrieval of application data
*/

import * as defaults from './defaults';
import { execSync } from "child_process";
import * as fs from 'fs';
import * as os from 'os';
import { modifyManifestFile } from 'office-addin-manifest';

export function addSecretToCredentialStore(ssoAppName: string, secret: string, isTest: boolean = false): void {
    try {
        switch (process.platform) {
            case "win32":
                console.log(`Adding application secret for ${ssoAppName} to Windows Credential Store`);
                const addSecretToWindowsStoreCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.addSecretCommandPath}" "${ssoAppName}" "${os.userInfo().username}" "${secret}"`;
                execSync(addSecretToWindowsStoreCommand, { stdio: "pipe" });
                break;
            case "darwin":
                console.log(`Adding application secret for ${ssoAppName} to Mac OS Keychain. You will need to provide an admin password to update the Keychain`);
                // Check first to see if the secret already exists i the keychain. If it does, delete it and recreate it
                const existingSecret = getSecretFromCredentialStore(ssoAppName);
                if (existingSecret !== '') {
                    const updatSecretInMacStoreCommand = `sudo security add-generic-password -a ${os.userInfo().username} -U -s "${ssoAppName}" -w "${secret}"`;
                    execSync(updatSecretInMacStoreCommand, { stdio: "pipe" });
                } else {
                    const addSecretToMacStoreCommand = `sudo security add-generic-password -a ${os.userInfo().username} -s "${ssoAppName}" -w "${secret}"`;
                    execSync(addSecretToMacStoreCommand, { stdio: "pipe" });
                }
                break;
            default:
                throw new Error(`Platform not supported: ${process.platform}`);
        }
    } catch (err) {
        throw new Error(`Unable to add secret for ${ssoAppName} to Windows Credential Store. \n${err}`);
    }
}

export function getSecretFromCredentialStore(ssoAppName: string, isTest: boolean = false): string {
    try {
        switch (process.platform) {
            case "win32":
                console.log(`Getting application secret for ${ssoAppName} from Windows Credential Store`);
                const getSecretFromWindowsStoreCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.getSecretCommandPath}" "${ssoAppName}" "${os.userInfo().username}"`;
                return execSync(getSecretFromWindowsStoreCommand, { stdio: "pipe" }).toString();
            case "darwin":
                console.log(`Getting application secret for ${ssoAppName} from Mac OS Keychain`);
                const getSecretFromMacStoreCommand = `${isTest ? "" : "sudo"} security find-generic-password -a ${os.userInfo().username} -s ${ssoAppName} -w`;
                return execSync(getSecretFromMacStoreCommand, { stdio: "pipe" }).toString();;
            default:
                throw new Error(`Platform not supported: ${process.platform}`);
        }

    } catch (err) {
        return '';
    }
}

function updateEnvFile(applicationId: string, envFilePath: string = defaults.envDataFilePath,): void {
    try {
        // Update .ENV file
        if (fs.existsSync(envFilePath)) {
            const appData = fs.readFileSync(envFilePath, 'utf8');
            const updatedAppData = appData.replace('CLIENT_ID=', `CLIENT_ID=${applicationId}`);
            fs.writeFileSync(envFilePath, updatedAppData);
        } else {
            throw new Error(`${envFilePath} does not exist`)
        }
    } catch (err) {
        throw new Error(`Unable to write SSO application data to ${envFilePath}. \n${err}`);
    }
}

function updateFallBackAuthDialogFile(applicationId: string, fallbackAuthDialogPath = defaults.fallbackAuthDialogTypescriptFilePath): void {
    let isTypecript: boolean = false;
    try {
        // Update fallbackAuthDialog.js
        if (fs.existsSync(fallbackAuthDialogPath)) {
            isTypecript = true;
            const srcFile = fs.readFileSync(fallbackAuthDialogPath, 'utf8');
            const updatedSrcFile = srcFile.replace('{application GUID here}', applicationId);
            fs.writeFileSync(fallbackAuthDialogPath, updatedSrcFile);
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

async function updateProjectManifest(applicationId: string, manifestPath: string = defaults.manifestFilePath): Promise<void> {
    console.log('Updating manifest with application ID');
    try {
        if (fs.existsSync(manifestPath)) {
            // Update manifest with application guid and unique manifest id
            const manifestContent: string = await fs.readFileSync(manifestPath, 'utf8');
            const re: RegExp = new RegExp('{application GUID here}', 'g');
            const updatedManifestContent: string = manifestContent.replace(re, applicationId);
            await fs.writeFileSync(manifestPath, updatedManifestContent);
            await modifyManifestFile(manifestPath, 'random');
        } else {
            throw new Error(`${manifestPath} does not exist`)
        }
    } catch (err) {
        throw new Error(`Unable to update ${manifestPath}. \n${err}`);
    }
}

export async function writeApplicationData(applicationId: string, manifestPath: string = defaults.manifestFilePath, envFilePath: string = defaults.envDataFilePath, fallbackAuthDialogPath = defaults.fallbackAuthDialogTypescriptFilePath) {
    updateEnvFile(applicationId, envFilePath);
    updateFallBackAuthDialogFile(applicationId, fallbackAuthDialogPath);
    await updateProjectManifest(applicationId, manifestPath)
}
