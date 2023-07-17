// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import {Client} from '@microsoft/microsoft-graph-client';

import {initializeGraphClient} from './client';
import {discoverHorseId} from './discovery';
import {CaseData, updateCase} from "./case";

async function main() {
    let client = initializeGraphClient();
    // await findHorse(client, "FY2024");
    await updateTestCases(client);
}

async function findHorse(client: Client, horseName: string) {
    let horseId = await discoverHorseId(client, horseName);
    if (!horseId) {
        throw Error(`Failed to find a spreadsheet named ${horseName}.xlsx`)
    }
    console.log(`The spreadsheet ${horseName}.xlsx has id '${horseId}'`);
}

async function updateTestCases(client: Client) {
    let testCases: CaseData[] = [
        {id: 27710, pledgeDate: new Date()},
        {
            id: 35,
            pledgeDate: new Date(),
            appointmentDate: new Date(),
            pledgeAmount: 700,
            client: "Dan B.",
            clinic: "Test Clinic",
            contact: "Sumeyye",
        },
        {
            id: 37,
            pledgeDate: new Date(),
            appointmentDate: new Date(),
            pledgeAmount: 1200,
            client: "Leanne B.",
            clinic: "Test Clinic",
            contact: "Chi Chi",
        },
    ]
    for (let caseInfo of testCases) {
        let rowInfo = await updateCase(client, caseInfo);
        if (rowInfo.isNew) {
            console.log(`New case ${caseInfo.id} added to spreadsheet on row ${rowInfo.row}.`)
        } else {
            console.log(`Existing case ${caseInfo.id} updated in spreadsheet on row ${rowInfo.row}.`)
        }
    }
}

main().then(() => {});
