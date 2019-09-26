import { configureSSOApplication } from './configure';
import { startSsoService } from './server';

export async function configureSSO(manifestPath: string, ssoAppName: string) {
    await configureSSOApplication(manifestPath, ssoAppName);
}

export async function startSSOService(ssoApplicationName: string) {
    await startSsoService(ssoApplicationName);
}
