import * as os from 'os';
import * as path from 'path';

export const ssoDataJsonFilePath: string = path.join(os.homedir(), '/office-addin-sso-data.json');
export const azRestpCreateCommandPath: string = `${__dirname}/scripts/azRestAppCreateCmd.txt`;
export const grantAdminConsentCommandPath = `${__dirname}/scripts/azGrantAdminConsentCmd.txt`;
export const setIdentifierUriCommmandPath: string = `${__dirname}/scripts/azRestSetIdentifierUri.txt`;
export const setSigninAudienceCommandPath: string = `${ __dirname }/scripts/azSetSignInAudienceCmd.txt`;
