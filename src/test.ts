// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import {ClientData, getClientData} from './settings';
import {discoverHorseId} from './discovery';
import {CaseData, updateCase} from "./case";
import {initializeGraphClient} from "./graphClient";

export async function updateHorseId(clientData: ClientData, horseName: string) {
    if (clientData.horseId) {
        return;
    }
    const horseId = await discoverHorseId(clientData, horseName);
    if (!horseId) {
        throw Error(`Failed to find a spreadsheet named ${horseName}.xlsx`)
    }
    clientData.horseId = horseId;
    console.log(`The spreadsheet ${horseName}.xlsx has id '${horseId}'`);
}

export async function updateTestCases(clientData: ClientData) {
    const testCases: CaseData[] = [
        {id: 27710, pledgeDate: '7/15/2023'},
        {
            id: 35,
            pledgeDate: '7/15/2023',
            appointmentDate: '7/22/2023',
            pledgeAmount: 700,
            client: "Dan B.",
            clinic: "Test Clinic",
            contact: "Chi Chi",
        },
        {
            id: 37,
            pledgeDate: '7/16/2023',
            appointmentDate: '7/21/2023',
            pledgeAmount: 1200,
            client: "Leanne B.",
            clinic: "Test Clinic",
            contact: "Chi Chi",
        },
    ]
    for (const caseInfo of testCases) {
        const rowInfo = await updateCase(clientData, caseInfo);
        if (rowInfo.isNew) {
            console.log(`New case ${caseInfo.id} added to spreadsheet on row ${rowInfo.row}.`)
        } else {
            console.log(`Existing case ${caseInfo.id} updated in spreadsheet on row ${rowInfo.row}.`)
        }
    }
}

function testAuth() {

}

async function test() {
    testAuth();
    const clientData = getClientData();
    clientData.client = initializeGraphClient(clientData);
    await updateHorseId(clientData, "FY2024 Development");
    await updateTestCases(clientData);
}

test()
    .then(() => console.log("Tests completed with no errors"));
