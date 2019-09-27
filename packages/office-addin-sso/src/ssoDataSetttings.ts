import * as defaults from './defaults';
import * as fs from 'fs';

export function writeApplicationJsonData(applicationJson: any, userJson: any, secret: string) {
    try {
        if (fs.existsSync(defaults.ssoDataJsonFilePath) && fs.readFileSync(defaults.ssoDataJsonFilePath, 'utf8') !== '' && fs.readFileSync(defaults.ssoDataJsonFilePath, 'utf8') !== 'undefined') {
            if (ssoApplicationExists(applicationJson.displayName)) {
                modifySsoJsonData(applicationJson, userJson, secret);
            } else {
                const ssoJsonData = readSsoJsonData();
                ssoJsonData.ssoApplicationInstances[applicationJson.displayName] = { applicationId: String, applicationSecret: String, tenantId: String };
                ssoJsonData.ssoApplicationInstances[applicationJson.displayName].applicationId = applicationJson.appId;
                ssoJsonData.ssoApplicationInstances[applicationJson.displayName].tenantId = userJson[0].tenantId;
                ssoJsonData.ssoApplicationInstances[applicationJson.displayName].applicationSecret = secret;
                fs.writeFileSync(defaults.ssoDataJsonFilePath, JSON.stringify((ssoJsonData), null, 2));
            }
        } else {
            let ssoJsonData = {};
            ssoJsonData[applicationJson.displayName] = applicationJson.displayName;
            ssoJsonData = { ssoApplicationInstances: ssoJsonData };
            ssoJsonData = { ssoApplicationInstances: { [applicationJson.displayName]: { ['applicationId']: applicationJson.appId, ['tenantId']: userJson[0].tenantId, ['applicationSecret']: secret } } };
            fs.writeFileSync(defaults.ssoDataJsonFilePath, JSON.stringify((ssoJsonData), null, 2));
        }
    } catch (err) {
        throw new Error(`Unable to write SSO application data to ${defaults.ssoDataJsonFilePath}. \n${err}`);
    }
}

async function modifySsoJsonData(applicationJson: any, userJson: any, secret: string): Promise<void> {
    try {
        const ssoJsonData = readSsoJsonData();
        if (ssoJsonData && ssoApplicationExists(applicationJson.displayName)) {
            ssoJsonData.ssoApplicationInstances[applicationJson.displayName].applicationId = applicationJson.appId;
            ssoJsonData.ssoApplicationInstances[applicationJson.displayName].tenantId = userJson[0].tenantId;
            ssoJsonData.ssoApplicationInstances[applicationJson.displayName].applicationSecret = secret;
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
