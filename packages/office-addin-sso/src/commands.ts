import * as sso from './configureSSO';

export function configureSSO(manifestPath: string, ssoAppName: string) {
    sso.configureSSO(manifestPath, ssoAppName);
}

export async function startServer(manifestPath: string) {
    sso.startSsoServer(manifestPath);
}
