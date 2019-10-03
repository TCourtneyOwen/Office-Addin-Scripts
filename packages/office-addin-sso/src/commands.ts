import { configureSSOApplication } from './configure';
// import { ISSOptions, SSOService } from './server';

export async function configureSSO(manifestPath: string, ssoAppName: string) {
    await configureSSOApplication(manifestPath, ssoAppName);
}

export async function startSSOService(ssoApplicationName: string) {
    // const ssoOptions: ISSOptions = {
    //     applicationName: ssoApplicationName,
    //     applicationApiScopes: { scp: 'access_as_user'},
    //     graphApi: '/me/drive/root/children',
    //     graphApiScopes: ['Files.Read.All'],
    //     queryParam: '?$select=name&$top=3'
    // };
    // const sso = new SSOService(ssoOptions);
    // sso.startSsoService();
}
