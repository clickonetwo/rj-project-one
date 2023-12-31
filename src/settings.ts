// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import {Client} from '@microsoft/microsoft-graph-client';

export interface ClientData {
    clientId: string,
    clientSecret: string,
    tenantId: string,
    authSecret: string,
    groupId: string,
    driveId: string,
    horseId: string,
    horseName: string,
    firstRow: string,
    client?: Client,
}

let clientData: ClientData | undefined;

export function getClientData(): ClientData {
    if (clientData) {
        return clientData;
    }
    const environmentData: ClientData = {
        clientId: process.env?.MS_CLIENT_ID || '',
        tenantId: process.env?.MS_TENANT_ID || '',
        clientSecret: process.env?.MS_CLIENT_SECRET || '',
        authSecret: process.env?.MS_AUTH_SECRET || '',
        groupId: process.env?.MS_GROUP_ID || '',
        driveId: process.env?.MS_DRIVE_ID || '',
        horseId: process.env?.MS_HORSE_ID || '',
        horseName: process.env?.MS_HORSE_NAME || '',
        firstRow: process.env?.MS_FIRST_CASE_ROW || '9',
    }
    if (!environmentData.clientId || !environmentData.tenantId ||
        !environmentData.clientSecret || !environmentData.authSecret) {
        throw Error("No authentication data found in environment");
    }
    if (!environmentData.groupId && !environmentData.driveId) {
        throw Error("No drive data found in environment");
    }
    clientData = environmentData;
    return clientData;
}
