// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { configureSSOApplication } from './configure';
import { SSOService } from './server';

export async function configureSSO(manifestPath: string) {
    await configureSSOApplication(manifestPath);
}

export async function startSSOService(manifestPath: string) {
    const sso = new SSOService(manifestPath);
    sso.startSsoService();
}
