// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as path from 'path';
import * as server from './../src/server';
import * as ssoData from './../src/ssoDataSettings';
import * as testHelper from "../../office-addin-test-helpers";
const appId: string = '584c9885-baa7-44ef-95b6-6df1064a2e25';
const ssoAppName: string = 'Office-Addin-Taskpane-SSO-Test';
const secret: string = '9lcUGBHc8F0s/8FINhwLmTUuhn@KBp=_';

describe("Unit Tests", function () {
    describe("addSecretToCredentialStore()/getSecretFromCredentialStore()", function () {
        it("Add secret and retrive secret from credential store", function () {
            ssoData.addSecretToCredentialStore(ssoAppName, secret);
            const retrievedSecret: string = ssoData.getSecretFromCredentialStore(ssoAppName).trim();
            assert.strictEqual(secret, retrievedSecret);
        });        
    });
    describe("writeApplicationData()", function () {
        const copyEnvFile: string = path.resolve(`${__dirname}/copy-test-env`);
        const copyFallbackAuthTaskpaneFile: string = path.resolve(`${__dirname}/copy-test-fallbackauthtaskpane`);
        const copyManifestFile: string = path.resolve(`${__dirname}/copy-test-manifest.xml`);
        before("Create copies of original files so we can edit them and then delete the copies at the end of test", function () {
            fs.copyFileSync(path.resolve(`${__dirname}/test-env`), copyEnvFile);
            fs.copyFileSync(path.resolve(`${__dirname}/test-fallbackauthtaskpane`), copyFallbackAuthTaskpaneFile);
            fs.copyFileSync(path.resolve(`${__dirname}/test-manifest.xml`), copyManifestFile);
        });
        it("Write application id to env, manifest and fallbackauthtaskpane files", async function () {
            await ssoData.writeApplicationData(appId, copyManifestFile, copyEnvFile, copyFallbackAuthTaskpaneFile);

            // Read from updated env copy and ensure it contains the appId
            const envFile = fs.readFileSync(copyEnvFile, 'utf8');
            assert.equal(envFile.includes(appId), true);         

            // Read from updated manifest copy and ensure it contains the appId
            const manifestFile = fs.readFileSync(copyEnvFile, 'utf8');
            assert.equal(manifestFile.includes(appId), true);

            // Read from updated fallbackauthtaskpane copy and ensure it contains the appId
            const fallbackAuthTaskpaneFile = fs.readFileSync(copyEnvFile, 'utf8');
            assert.equal(fallbackAuthTaskpaneFile.includes(appId), true);
        });
        after("Delete copies of test files", function () {
            fs.unlinkSync(copyEnvFile);
            fs.unlinkSync(copyFallbackAuthTaskpaneFile);
            fs.unlinkSync(copyManifestFile);
        });
    });
    describe("SSO Service", function () {
        const sso = new server.SSOService(`${__dirname}/test-manifest.xml`);
        it("Start SSO Service and ping server to verify it's started", async function () {
            // Start SSO service
            const serverStarted = await sso.startSsoService(true /* mochaTest */);
            assert.equal(serverStarted, true);
            const serverState = sso.getTestServerState();
            assert.equal(serverState, true);

            // Ping server to validate it's listening as expected
            const testServerResponse = await testHelper.pingTestServer(3000);
            assert.equal(testServerResponse != undefined, true);
            assert.equal(testServerResponse["status"], 200);
            assert.equal(testServerResponse["platform"], process.platform);
        });
        it("Stop SSO Service and verify it's stopped", async function () {
            // Stop SSO service
            const serverStopped = await sso.stopServer();
            assert.equal(serverStopped, true);
            const serverState = sso.getTestServerState();
            assert.equal(serverState, false);
        });
    });
});

