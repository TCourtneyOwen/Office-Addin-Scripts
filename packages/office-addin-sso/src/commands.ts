import { configureSSOApplication } from './configureSSO';
import { startSsoServer } from './server';

export async function configureSSO(manifestPath: string, ssoAppName: string) {
    await configureSSOApplication(manifestPath, ssoAppName);
}

export async function startServer(ssoApplicationName: string) {
    await startSsoServer(ssoApplicationName);
}
