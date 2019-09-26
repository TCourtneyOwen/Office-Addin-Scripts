// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/*
    This file provides the provides server startup, authorization context creation, and the Web APIs of the add-in.
*/

import * as https from 'https';
import * as express from 'express';
import * as bodyParser from 'body-parser';
import * as cors from 'cors';
import * as morgan from 'morgan';
import { AuthModule } from './auth';
import { MSGraphHelper } from './msgraph-helper';
import { UnauthorizedError } from './errors';
import * as devCerts from 'office-addin-dev-certs';

/* Set the environment to development if not set */
const env = process.env.NODE_ENV || 'development';

/* A promisified express handler to catch errors easily */
const handler = (callback: (req: express.Request, res: express.Response, next?: express.NextFunction) => Promise<any>) =>
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


export interface ISSOAuthServiceOptions {
    applicationId: string,
    tenantId: string,
    applicationSecret: string
}

export class SSOAuthService {
    private app: express.Express;
    private authModule: AuthModule;

    constructor(ssoAuthOptions: ISSOAuthServiceOptions) {
        this.authModule = new AuthModule(
            ssoAuthOptions.applicationId,
            ssoAuthOptions.applicationSecret,
            'common',
            'https://login.microsoftonline.com',
            '.well-known/openid-configuration',
            '/oauth2/v2.0/token',
            ssoAuthOptions.applicationId,
            ['access_as_user'],
            `https://login.microsoftonline.com/${ssoAuthOptions.tenantId}`
        );
        this.app = express();
    }

    public async startServer() {
        this.app = express();
        this.app.use(bodyParser.json());
        this.app.use(bodyParser.urlencoded({ extended: true }));
        this.app.use(cors());
        this.app.use(morgan('dev'));
        this.app.use(express.static('public'));
        /* Turn off caching when debugging */
        this.app.use((req, res, next) => {
            res.header('Cache-Control', 'private, no-cache, no-store, must-revalidate');
            res.header('Expires', '-1');
            res.header('Pragma', 'no-cache');
            next();
        });
        this.app.get('/api/values', handler(async (req, res) => {
            /**
             * Only initializes the auth the first time
             * and uses the downloaded keys information subsequently.
             */
            await this.authModule.initialize();
            const { jwt } = this.authModule.verifyJWT(req, { scp: 'access_as_user' });

            // 1. We don't pass a resource parameter because the token endpoint is Azure AD V2.
            // 2. Always ask for the minimal permissions that the application needs.
            const graphToken = await this.authModule.acquireTokenOnBehalfOf(jwt, ['files.read.all']);

            // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
            // and only the top 3 folder or file names.
            // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
            // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
            // sanitized so that it cannot be used in a Response header injection attack.
            const graphData = await MSGraphHelper.getGraphData(graphToken, '/me/drive/root/children', '?$select=name&$top=3');

            // If Microsoft Graph returns an error, such as invalid or expired token,
            // relay it to the client.
            if (graphData.code) {
                if (graphData.code === 401) {
                    throw new UnauthorizedError('Microsoft Graph error', graphData);
                }
            }

            // Graph data includes OData metadata and eTags that we don't need.
            // Send only what is actually needed to the client: the item names.
            const itemNames: string[] = [];
            const oneDriveItems: string[] = graphData['value'];
            for (let item of oneDriveItems) {
                itemNames.push(item['name']);
            }
            return res.json(itemNames);
        }));

        /**
         * HTTP GET: /index.html
         * Loads the add-in home page.
         */
        this.app.get('/index.html', handler(async (req, res) => {
            return res.sendfile('index.html');
        }));
        /**
     * If running on development env, then use the locally available certificates.
     */
        if (env === 'development') {
            const options = await devCerts.getHttpsServerOptions();
            https.createServer(options, this.app).listen(3000, () => console.log('Server running on 3000'));
        }
        else {
            /**
             * We don't use https as we are assuming the production environment would be on Azure.
             * Here IIS_NODE will handle https requests and pass them along to the node http module
             */
            this.app.listen(process.env.port || 1337, () => console.log(`Server listening on port ${process.env.port}`));
        }
    }
}


// export async function startSsoServer(ssoApplicationName: string): Promise<boolean> {
//     return new Promise<boolean>(async (resolve, reject) => {
//         try {
//             if (ssoApplicationExists(ssoApplicationName)) {
//                 const ssoApplicationData = readSsoJsonData();
//                 const serverOptions: server.ISSOAuthServiceOptions = {
//                     applicationId: ssoApplicationData.ssoApplicationInstances[ssoApplicationName].applicationId,
//                     tenantId: ssoApplicationData.ssoApplicationInstances[ssoApplicationName].tenantId,
//                     applicationSecret: ssoApplicationData.ssoApplicationInstances[ssoApplicationName].applicationSecret
//                 };
//                 const ssoAuthService: server.SSOAuthService = new server.SSOAuthService(serverOptions);
//                 ssoAuthService.startServer();
//                 resolve(true);
//             }
//             resolve(false);

//         } catch (err) {
//             reject(false);
//         }
//     });
// }
