// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/*
    This file provides the provides server startup, authorization context creation, and the Web APIs of the add-in.
*/

import * as https from 'https';
import * as express from 'express';
import * as bodyParser from 'body-parser';
import * as cors from 'cors';
import * as devCerts from 'office-addin-dev-certs';
import * as morgan from 'morgan';
import { AuthModule } from './auth';
import { MSGraphHelper } from './msgraph-helper';
import { UnauthorizedError } from './errors';
import { getSecretFromCredentialStore, readSsoJsonData } from './ssoDataSettings';

export interface ISsoOptions {
    applicationName: string;
    multiTenant: boolean;
    applicationApiScopeName: string;
    graphApi: string;
    graphApiScopes: [string];
    queryParam: string;
}

export class SSOService {
    private app: express.Express;
    private auth: AuthModule;
    private port: number;
    private ssoOptions: ISsoOptions;

    constructor(ssoOptions: ISsoOptions, port: number) {
        this.app = express();
        this.port = port;
        this.ssoOptions = ssoOptions;

        const ssoJsonData = readSsoJsonData();
        const appSecret = getSecretFromCredentialStore(this.ssoOptions.applicationName);

        this.auth = new AuthModule(
            ssoJsonData.ssoApplicationInstances[this.ssoOptions.applicationName].applicationId,
            appSecret,
            'common',
            'https://login.microsoftonline.com',
            this.ssoOptions.multiTenant ? 'v2.0/.well-known/openid-configuration' : '.well-known/openid-configuration',
            '/oauth2/v2.0/token',
            ssoJsonData.ssoApplicationInstances[this.ssoOptions.applicationName].applicationId,
            this.ssoOptions.applicationApiScopeName,
            `https://login.microsoftonline.com/${ssoJsonData.ssoApplicationInstances[this.ssoOptions.applicationName].tenantId}/v2.0`,
        );
        this.auth.initialize();
    }

    public async startSsoService(): Promise<boolean> {
        return new Promise<boolean>(async (resolve, reject) => {
            /* Set the environment to development if not set */
            const env = process.env.NODE_ENV || 'development';

            /* A promisified express handler to catch errors easily */
            const handler =
                (callback: (req: express.Request, res: express.Response, next?: express.NextFunction) => Promise<any>) =>
                    (req, res, next) => callback(req, res, next)
                        .catch(error => {
                            /* If the headers are already sent then resort to the built in error handler */
                            if (res.headersSent) {
                                return next(error);
                            }

                            /**
                             * If running development environment we send the error details back.
                             * Else we send the right code and message.
                             */
                            if (env === 'development') {
                                return res.status(error.code || 500).json({ error });
                            }
                            else {
                                return res.status(error.code || 500).send(error.message);
                            }
                        });

            /* Create the express app and add the required middleware */
            this.app.use(bodyParser.json());
            this.app.use(bodyParser.urlencoded({ extended: true }));
            this.app.use(cors());
            this.app.use(morgan('dev'));
            this.app.use(express.static('dist'));
            /* Turn off caching when debugging */
            this.app.use((req, res, next) => {
                res.header('Cache-Control', 'private, no-cache, no-store, must-revalidate');
                res.header('Expires', '-1');
                res.header('Pragma', 'no-cache');
                next();
            });

            this.startServer(this.app, this.port);

            /**
             * HTTP GET: /api/values
             * When passed a JWT token in the header, it extracts it and
             * and exchanges for a token that has permissions to graph.
             */
            this.app.get('/api/values', handler(async (req, res) => {

                // 1. We don't pass a resource parameter because the token endpoint is Azure AD V2.
                // 2. Always ask for the minimal permissions that the application needs.
                const graphToken = await this.auth.getGraphToken(req, this.ssoOptions.graphApiScopes);
                const graphData = await MSGraphHelper.getGraphData(graphToken, this.ssoOptions.graphApi, this.ssoOptions.queryParam);
                // If Microsoft Graph returns an error, such as invalid or expired token,
                // relay it to the client.
                if (graphData.code) {
                    if (graphData.code === 401) {
                        throw new UnauthorizedError('Microsoft Graph error', graphData);
                    }
                }
                return res.json(graphData);
            }));

            /**
             * HTTP GET: /index.html
             * Loads the add-in home page.
             */
            this.app.get('/index.html', handler(async (req, res) => {
                return res.sendfile('index.html');
            }));
        });
    }

    private async startServer(app, port: number) {
        const options = await devCerts.getHttpsServerOptions();
        https.createServer(options, app).listen(port, () => console.log(`Server running on ${port}`));
    }

    public async getGraphToken(accessToken) {
        return await this.auth.getGraphToken(accessToken, this.ssoOptions.graphApiScopes);
    }
}
