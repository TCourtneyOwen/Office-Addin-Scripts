/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file is the main Node.js server file that defines the express middleware.
 */

if (process.env.NODE_ENV !== 'production') {
    require('dotenv').config();
}
import * as createError from 'http-errors'
import * as express from 'express';
import * as path from 'path';
import * as cookieParser from 'cookie-parser';
import * as logger from 'morgan';
import { AuthRouter } from './authRoute';
import * as indexRouter from './indexRoute';
import { MSGraphHelper } from './msgraph-helper';

export class App {
    appInstance: express.Express;
    port: number | string;
    constructor(port: number | string) {
        this.appInstance = express();
        this.port = port;
    }

    async initialize() {

        this.appInstance.set('port', this.port);

        // // view engine setup
        this.appInstance.set('views', path.join(__dirname, 'views'));
        this.appInstance.set('view engine', 'pug');
        
        this.appInstance.use(logger('dev'));
        this.appInstance.use(express.json());
        this.appInstance.use(express.urlencoded({ extended: false }));
        this.appInstance.use(cookieParser());

        /* Turn off caching when developing */
        if (process.env.NODE_ENV !== 'production') {
            this.appInstance.use(express.static(path.join(process.cwd(), 'dist'),
                { etag: false }));

            this.appInstance.use(function (req, res, next) {
                res.header('Cache-Control', 'private, no-cache, no-store, must-revalidate');
                res.header('Expires', '-1');
                res.header('Pragma', 'no-cache');
                next()
            });
        } else {
            // In production mode, let static files be cached.
            this.appInstance.use(express.static(path.join(process.cwd(), 'dist')));
        }

        const authRouter = new AuthRouter();

        this.appInstance.use('/', indexRouter.router);
        this.appInstance.use('/auth', authRouter.router);

        this.appInstance.get('/taskpane.html', (async (req, res) => {
            return res.sendfile('taskpane.html');
        }));

        this.appInstance.get('/fallbackauthdialog.html', (async (req, res) => {
            return res.sendfile('fallbackauthdialog.html');
        }));

        this.appInstance.get('/getuserdata', async function (req, res, next) {
            const graphToken = req.get('access_token');

            // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
            // and only the top 10 folder or file names.
            // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
            // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
            // sanitized so that it cannot be used in a Response header injection attack. 
            const graphData = await MSGraphHelper.getGraphData(graphToken, "/me", "");

            // If Microsoft Graph returns an error, such as invalid or expired token,
            // there will be a code property in the returned object set to a HTTP status (e.g. 401).
            // Relay it to the client. It will caught in the fail callback of `makeGraphApiCall`.
            if (graphData.code) {
                next(createError(graphData.code, "Microsoft Graph error " + JSON.stringify(graphData)));
            }
            else {
                // MS Graph data includes OData metadata and eTags that we don't need.
                // Send only what is actually needed to the client: the item names.
                const userProfileInfo = [];
                userProfileInfo.push(graphData["displayName"]);
                userProfileInfo.push(graphData["jobTitle"]);
                userProfileInfo.push(graphData["mail"]);
                userProfileInfo.push(graphData["mobilePhone"]);
                userProfileInfo.push(graphData["officeLocation"]);

                res.send(userProfileInfo);
            }
        });


        // Catch 404 and forward to error handler
        this.appInstance.use(function (req, res, next) {
            next(createError(404));
        });

        // error handler
        this.appInstance.use(function (err, req, res, next) {
            // set locals, only providing error in development
            res.locals.message = err.message;
            res.locals.error = req.app.get('env') === 'development' ? err : {};

            // render the error page
            res.status(err.status || 500);
            res.render('error');
        });
    }
}