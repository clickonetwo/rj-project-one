// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import {ClientData, getClientData} from './settings';
import {discoverHorseId} from './discovery';
import {updateCase} from "./case";
import {initializeGraphClient} from "./graphClient";
import {tokenFromContent, validateToken} from "./auth";
import {prepareCase, QueryString} from "./routes";

export async function testDiscovery(clientData: ClientData) {
    if (!clientData.horseName) {
        throw Error(`No horse name found in environment`);
    }
    const horseId = await discoverHorseId(clientData, clientData.horseName);
    if (!horseId) {
        throw Error(`Failed to find a spreadsheet named '${clientData.horseName}.xlsx'`)
    }
    console.log(`The spreadsheet ${clientData.horseName}.xlsx has id '${horseId}'`);
    if (clientData.horseId !== horseId) {
        console.log(`Updating client data: be sure to update the environment settings`)
        clientData.horseId = horseId;
    }
}

export async function testUpdate(clientData: ClientData) {
    const testCases: QueryString[] = [
        {
            id: '34',
            client: 'Jane D.',
            clinic: 'PSV Only',
            contact: 'Sara',
        },
        {
            id: '35',
            pledgeDate: '2023-07-15',
            appointmentDate: '2023-07-22',
            pledgeAmount: 700,
            client: 'John D.',
            clinic: 'Safe House',
            contact: 'Laura',
        },
        {
            id: Math.floor(Math.random() * 200) + 100,
            pledgeDate: '2023-07-16',
            appointmentDate: '2023-07-21',
            pledgeAmount: 1200,
            client: 'Jill D.',
            clinic: "Equal Justice",
            contact: "Sanni",
        },
    ]
    for (const caseInfo of testCases) {
        const caseData = prepareCase(caseInfo);
        if (typeof caseData === 'string') {
            throw Error(caseData);
        }
        await updateCase(clientData, caseData);
    }
}

async function testAuth() {
    function makeContent(date: Date) {
        const lastMinute = new Date(date);
        lastMinute.setSeconds(0);
        lastMinute.setMilliseconds(0);
        return lastMinute.toISOString();
    }
    const testSecret = 'a pretty stupid secret that is reasonably long';
    const now = new Date();
    const seconds = now.getUTCSeconds();
    const ms = now.getUTCMilliseconds();
    const token = tokenFromContent(testSecret, makeContent(now));
    if (!validateToken(testSecret, token)) {
        throw Error('Immediate token failed to validate')
    }
    const delay = ((60 - seconds) * 1000) - ms
    console.log(`Waiting ${delay/1000} seconds during token test...`);
    await new Promise(resolve => setTimeout(resolve, delay));
    if (!validateToken(testSecret, token)) {
        throw Error('Delayed token failed to validate');
    }
}

async function test(...what: string[]) {
    if (what.length == 0) {
        what = ['auth', 'discover', 'update'];
    } else {
        what = what.map((s) => s.toLowerCase());
    }
    if (what.includes('auth')) {
        await testAuth();
    }
    if (what.includes('discover') || what.includes('update')) {
        const clientData = getClientData();
        clientData.client = initializeGraphClient(clientData);
        if (what.includes('discover')) {
            await testDiscovery(clientData);
        }
        if (what.includes('update')) {
            await testUpdate(clientData);
        }
    }
}

test(...process.argv.slice(2))
    .then(() => console.log("Tests completed with no errors"));
