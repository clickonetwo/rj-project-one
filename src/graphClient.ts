// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import 'isomorphic-fetch';
import * as azure from '@azure/identity';
import * as authProviders from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import * as graph from '@microsoft/microsoft-graph-client';

import {ClientData} from "./settings";

export function initializeGraphClient(clientData: ClientData) {
    const clientSecretCredential = new azure.ClientSecretCredential(
        clientData.tenantId,
        clientData.clientId,
        clientData.clientSecret
    );
    const authProvider = new authProviders.TokenCredentialAuthenticationProvider(
        clientSecretCredential, {
            scopes: ['https://graph.microsoft.com/.default']
        });
    return graph.Client.initWithMiddleware({
        authProvider: authProvider
    });
}
