"use strict";
// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.updateCase = void 0;
require("isomorphic-fetch");
const md5_1 = __importDefault(require("crypto-js/md5"));
const settings_1 = require("./local/settings");
async function updateCase(client, caseData) {
    let horsePath = `/drives/${settings_1.devData.driveId}/items/${settings_1.devData.horseId}`;
    let sessionId = await openSession(client, horsePath);
    let rowInfo = await findCase(client, horsePath, sessionId, caseData.id);
    await writeCase(client, horsePath, sessionId, caseData, rowInfo);
    await closeSession(client, horsePath, sessionId);
    return rowInfo;
}
exports.updateCase = updateCase;
async function findCase(client, horsePath, sessionId, caseId) {
    try {
        // First, get the filled range
        let range = await client.api(`${horsePath}/workbook/worksheets/Cases/usedRange(valuesOnly=true)`)
            .header('workbook-session-id', sessionId)
            .select(['address', 'columnIndex', 'columnCount', 'rowIndex', 'rowCount'])
            .get();
        // console.log(range)
        // Next, search for the case in the first column of that range
        //
        // Note: Excel row numbers start at 1, but the returned rowIndex starts at 0
        // Since the end of the range *includes* the starting row, we don't add 1 there.
        let opBody = {
            lookupValue: caseId,
            lookupArray: {
                address: `Cases!A${range.rowIndex + 1}:A${range.rowIndex + range.rowCount}`
            },
            matchType: 0,
        };
        let found = await client.api(`${horsePath}/workbook/functions/match`)
            .header('workbook-session-id', sessionId)
            .post(opBody);
        // console.log(found)
        if (found.error === null) {
            console.log(`Found existing ${caseId} in cell A${found.value + 1}`);
            return { row: found.value + 1, isNew: false };
        }
        else {
            let newRow = range.rowIndex + range.rowCount + 1;
            console.log(`Inserting new case ${caseId} into cell A${newRow}`);
            return { row: newRow, isNew: true };
        }
    }
    catch (err) {
        throw Error(`Can't find case ${caseId}: ${err}`);
    }
}
async function writeCase(client, horsePath, sessionId, caseData, rowInfo) {
    function excelDate(date) {
        return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
    }
    try {
        let rangeValues = [
            caseData.id,
            (caseData === null || caseData === void 0 ? void 0 : caseData.client) ? caseData.client : null,
            (caseData === null || caseData === void 0 ? void 0 : caseData.pledgeDate) ? excelDate(caseData.pledgeDate) : null,
            (caseData === null || caseData === void 0 ? void 0 : caseData.appointmentDate) ? excelDate(caseData.appointmentDate) : null,
            (caseData === null || caseData === void 0 ? void 0 : caseData.clinic) ? caseData.clinic : null,
            (caseData === null || caseData === void 0 ? void 0 : caseData.pledgeAmount) ? caseData.pledgeAmount : null,
            (caseData === null || caseData === void 0 ? void 0 : caseData.invoiceStatus) ? caseData.invoiceStatus : null,
            (caseData === null || caseData === void 0 ? void 0 : caseData.contact) ? caseData.contact : null,
        ];
        let rangeAddress = `A${rowInfo.row}:H${rowInfo.row}`;
        let opPath = `/workbook/worksheets/Cases/range(address='${rangeAddress}')`;
        let update = await client.api(`${horsePath}/${opPath}`)
            .header('workbook-session-id', sessionId)
            .select(['address', 'values'])
            .patch({ values: [rangeValues] });
        // console.log(update)
    }
    catch (err) {
        throw Error(`Can't update case ${caseData.id}: ${err}`);
    }
}
async function openSession(client, horsePath) {
    try {
        let result = await client.api(`${horsePath}/workbook/createSession`)
            .post({ persistChanges: true });
        if (result.id !== undefined) {
            console.log(`Created workbook session ID ${(0, md5_1.default)(result.id)}`);
            // console.log(result)
            return result.id;
        }
    }
    catch (err) {
        throw Error(`Failed to open session: ${err}`);
    }
    throw Error(`No session ID was returned`);
}
async function closeSession(client, horsePath, sessionId) {
    try {
        let result = await client.api(`${horsePath}/workbook/closeSession`)
            .header('workbook-session-id', sessionId)
            .post({});
        console.log(`Closed workbook session ID ${(0, md5_1.default)(sessionId)}`);
        // console.log(result)
    }
    catch (err) {
        throw Error(`Failed to close session: ${err}`);
    }
}
//# sourceMappingURL=case.js.map