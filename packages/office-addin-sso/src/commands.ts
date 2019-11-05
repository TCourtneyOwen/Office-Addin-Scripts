import { configureSSOApplication } from './configure';
import { SSOService } from './server';

export async function configureSSO(manifestPath: string) {
    await configureSSOApplication(manifestPath);
}

export async function startSSOService(manifestPath: string) {
    const sso = new SSOService(manifestPath);
    sso.startSsoService();
}
