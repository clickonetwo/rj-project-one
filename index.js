"use strict";
// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.
Object.defineProperty(exports, "__esModule", { value: true });
const client_1 = require("./client");
const discovery_1 = require("./discovery");
const case_1 = require("./case");
async function main() {
    let client = (0, client_1.initializeGraphClient)();
    // await findHorse(client, "FY2024");
    await updateTestCases(client);
}
async function findHorse(client, horseName) {
    let horseId = await (0, discovery_1.discoverHorseId)(client, horseName);
    if (!horseId) {
        throw Error(`Failed to find a spreadsheet named ${horseName}.xlsx`);
    }
    console.log(`The spreadsheet ${horseName}.xlsx has id '${horseId}'`);
}
async function updateTestCases(client) {
    let testCases = [
        { id: 27710, pledgeDate: new Date() },
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
    ];
    for (let caseInfo of testCases) {
        let rowInfo = await (0, case_1.updateCase)(client, caseInfo);
        if (rowInfo.isNew) {
            console.log(`New case ${caseInfo.id} added to spreadsheet on row ${rowInfo.row}.`);
        }
        else {
            console.log(`Existing case ${caseInfo.id} updated in spreadsheet on row ${rowInfo.row}.`);
        }
    }
}
main().then(() => { });
//# sourceMappingURL=index.js.map