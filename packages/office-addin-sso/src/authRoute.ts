/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file defines the routes within the authRoute router.
 */

import * as express from 'express';
import * as fetch from 'node-fetch';
const form = require('form-urlencoded').default;

export class AuthRouter {
    public router: express.Router;
    constructor() {
        this.router = express.Router();

        this.router.get('/', async function (req, res, next) {
            const authorization = req.get('Authorization');
            if (authorization == null) {
                let error = new Error('No Authorization header was found.');
                next(error);
            }
            else {
                const [/* schema */, jwt] = authorization.split(' ');
                const formParams = {
                    client_id: process.env.CLIENT_ID,
                    client_secret: process.env.secret,
                    grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                    assertion: jwt,
                    requested_token_use: 'on_behalf_of',
                    scope: ['User.Read'].join(' ')
                };

                const stsDomain = 'https://login.microsoftonline.com';
                const tenant = 'common';
                const tokenURLSegment = 'oauth2/v2.0/token';

                try {
                    const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
                        method: 'POST',
                        body: form(formParams),
                        headers: {
                            'Accept': 'application/json',
                            'Content-Type': 'application/x-www-form-urlencoded'
                        }
                    });
                    const json = await tokenResponse.json();

                    res.send(json);
                }
                catch (error) {
                    res.status(500).send(error);
                }
            }
        });
    }
}

