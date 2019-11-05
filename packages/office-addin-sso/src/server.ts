// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/*
    This file provides the provides server startup, authorization context creation, and the Web APIs of the add-in.
*/
import * as https from 'https';
import * as devCerts from 'office-addin-dev-certs';
import * as manifest from 'office-addin-manifest';
require('dotenv').config();
import { App } from './app';
import { getSecretFromCredentialStore } from './ssoDataSettings'

export class SSOService {
    private app: App;
    manifestPath: string;
    private port: number | string
    constructor(manifestPath: string) {
        this.port = process.env.PORT || '3000';
        this.app = new App(this.port);
        this.app.initialize();
        this.manifestPath = manifestPath;
    }

    public async startSsoService(): Promise<boolean> {
        return new Promise<boolean>(async (resolve, reject) => {
            try {
                this.getSecret();
                this.startServer(this.app.appInstance, this.port);
                resolve(true);
            } catch {
                reject(false);
            }
        });
    }

    private async getSecret() {
        const manifestInfo = await manifest.readManifestFile(this.manifestPath);
        const appSecret = getSecretFromCredentialStore(manifestInfo.displayName);
        process.env.secret = appSecret;
    }

    private async startServer(app, port) {
        const options = await devCerts.getHttpsServerOptions();
        https.createServer(options, app).listen(port, () => console.log(`Server running on ${port}`));
    }
}
