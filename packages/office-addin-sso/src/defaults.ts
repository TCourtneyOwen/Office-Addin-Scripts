import * as path from 'path';

export const azCliInstallCommandPath: string = path.resolve(`${__dirname}/scripts/azCliInstallCmd.ps1`);
export const azRestpCreateCommandPath: string = path.resolve(`${__dirname}/scripts/azRestAppCreateCmd.txt`);
export const fallbackAuthDialogFilePath = path.resolve(`${process.cwd()}/src/taskpane/fallbackAuthDialog.ts`);
export const getApplicationInfoCommandPath: string = path.resolve(`${__dirname}/scripts/azGetApplicationInfoCmd.txt`);
export const getInstalledAppsPath: string = path.resolve(`${__dirname}/scripts/getInstalledApps.ps1`);
export const grantAdminConsentCommandPath = path.resolve(`${__dirname}/scripts/azGrantAdminConsentCmd.txt`);
export const setIdentifierUriCommmandPath: string = path.resolve(`${__dirname}/scripts/azRestSetIdentifierUri.txt`);
export const setSigninAudienceCommandPath: string = path.resolve(`${__dirname}/scripts/azSetSignInAudienceCmd.txt`);
export const ssoDataFilePath = path.resolve(`${process.cwd()}/.ENV`);
export const addSecretCommandPath: string = path.resolve(`${__dirname}/scripts/addAppSecret.ps1`);
export const getSecretCommandPath: string = path.resolve(`${__dirname}/scripts/getAppSecret.ps1`);

