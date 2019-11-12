// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as path from 'path';

export const azCliInstallCommandPath: string = path.resolve(`${__dirname}/scripts/azCliInstallCmd.ps1`);
export const azAddSecretCommandPath = path.resolve(`${__dirname}/scripts/azAddSecretCmd.txt`);
export const azRestApppCreateCommandPath: string = path.resolve(`${__dirname}/scripts/azRestAppCreateCmd.txt`);
export const envDataFilePath = path.resolve(`${process.cwd()}/.ENV`);
export const fallbackAuthDialogTypescriptFilePath = path.resolve(`${process.cwd()}/src/taskpane/fallbackAuthTaskpane.ts`);
export const fallbackAuthDialogJavascriptFilePath = path.resolve(`${process.cwd()}/src/taskpane/fallbackAuthTaskpane.js`);
export const getInstalledAppsPath: string = path.resolve(`${__dirname}/scripts/getInstalledApps.ps1`);
export const manifestFilePath = path.resolve(`${process.cwd()}/manifest.xml`);
export const setIdentifierUriCommmandPath: string = path.resolve(`${__dirname}/scripts/azRestSetIdentifierUri.txt`);
export const setSigninAudienceCommandPath: string = path.resolve(`${__dirname}/scripts/azSetSignInAudienceCmd.txt`);
export const addSecretCommandPath: string = path.resolve(`${__dirname}/scripts/addAppSecret.ps1`);
export const getSecretCommandPath: string = path.resolve(`${__dirname}/scripts/getAppSecret.ps1`);

