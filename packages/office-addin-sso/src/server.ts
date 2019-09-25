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
import { readSsoJsonData } from './configureSSO';

export async function startSsoServer(ssoApplicationName: string): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
        /* Set the environment to development if not set */
        const env = process.env.NODE_ENV || 'development';
        const ssoJsonData = readSsoJsonData();
        if (ssoJsonData === undefined) {
            reject(new UnauthorizedError('No application data available for specified application'));
        }

        /* Instantiate AuthModule to assist with JWT parsing and verification, and token acquisition. */
        const auth = new AuthModule(
    /* These values are required for our application to exchange the token and get access to the resource data */
    /* client_id */ ssoJsonData.ssoApplicationInstances['ssotest'].applicationId,
    /* client_secret */ ssoJsonData.ssoApplicationInstances['ssotest'].applicationSecret,

    /* This information tells our server where to download the signing keys to validate the JWT that we received,
     * and where to get tokens. This is not configured for multi tenant; i.e., it is assumed that the source of the JWT and our application live
     * on the same tenant.
     */
    /* tenant */ 'common',
    /* stsDomain */ 'https://login.microsoftonline.com',
    /* discoveryURLsegment */ '.well-known/openid-configuration',
    /* tokenURLsegment */ '/oauth2/v2.0/token',

    /* Token is validated against the following values: */
    // Audience is the same as the client ID because, relative to the Office host, the add-in is the "resource".
    /* audience */ ssoJsonData.ssoApplicationInstances['ssotest'].applicationId,
    /* scopes */['access_as_user'],
    /* issuer */ `https://login.microsoftonline.com/${ssoJsonData.ssoApplicationInstances['ssotest'].tenantId}/v2.0`,
        );

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
        const app = express();
        app.use(bodyParser.json());
        app.use(bodyParser.urlencoded({ extended: true }));
        app.use(cors());
        app.use(morgan('dev'));
        app.use(express.static('public'));
        /* Turn off caching when debugging */
        app.use((req, res, next) => {
            res.header('Cache-Control', 'private, no-cache, no-store, must-revalidate');
            res.header('Expires', '-1');
            res.header('Pragma', 'no-cache');
            next();
        });

        startServer(app, env);

        /**
         * HTTP GET: /api/values
         * When passed a JWT token in the header, it extracts it and
         * and exchanges for a token that has permissions to graph.
         */
        app.get('/api/values', handler(async (req, res) => {
            /**
             * Only initializes the auth the first time
             * and uses the downloaded keys information subsequently.
             */
            await auth.initialize();
            const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' });

            // 1. We don't pass a resource parameter because the token endpoint is Azure AD V2.
            // 2. Always ask for the minimal permissions that the application needs.
            const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);

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
        app.get('/index.html', handler(async (req, res) => {
            return res.sendfile('index.html');
        }));

        async function startServer(app, env: string) {
            /**
         * If running on development env, then use the locally available certificates.
         */
            if (env === 'development') {
                const options = await devCerts.getHttpsServerOptions();
                https.createServer(options, app).listen(3000, () => console.log('Server running on 3000'));
            }
            else {
                /**
                 * We don't use https as we are assuming the production environment would be on Azure.
                 * Here IIS_NODE will handle https requests and pass them along to the node http module
                 */
                app.listen(process.env.port || 1337, () => console.log(`Server listening on port ${process.env.port}`));
            }
        }
    });
}
