// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import 'isomorphic-fetch';
import md5 from 'crypto-js/md5';

import {Client} from '@microsoft/microsoft-graph-client';

import {devData as clientData} from './local/settings';

export interface CaseData {
    id: number,
    client?: string,
    pledgeDate?: Date,
    appointmentDate?: Date,
    clinic?: string,
    pledgeAmount?: number,
    invoiceStatus?: string,
    contact?: string,
}

export interface RowInfo {
    row: number,
    isNew: boolean,
}

export async function updateCase(client: Client, caseData: CaseData) {
    let horsePath = `/drives/${clientData.driveId}/items/${clientData.horseId}`
    let sessionId = await openSession(client, horsePath);
    let rowInfo = await findCase(client, horsePath, sessionId, caseData.id);
    await writeCase(client, horsePath, sessionId, caseData, rowInfo)
    await closeSession(client, horsePath, sessionId);
    return rowInfo;
}

async function findCase(client: Client, horsePath: string, sessionId: string, caseId: number): Promise<RowInfo> {
    try {
        // First, get the filled range
        let range = await client.api(`${horsePath}/workbook/worksheets/Cases/usedRange(valuesOnly=true)`)
            .header('workbook-session-id', sessionId)
            .select(['address', 'columnIndex', 'columnCount', 'rowIndex', 'rowCount'])
            .get()
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
        }
        let found = await client.api(`${horsePath}/workbook/functions/match`)
            .header('workbook-session-id', sessionId)
            .post(opBody);
        // console.log(found)
        if (found.error === null) {
            console.log(`Found existing ${caseId} in cell A${found.value + 1}`)
            return {row: found.value + 1, isNew: false}
        } else {
            let newRow = range.rowIndex + range.rowCount + 1
            console.log(`Inserting new case ${caseId} into cell A${newRow}`)
            return {row: newRow, isNew: true}
        }
    } catch (err) {
        throw Error(`Can't find case ${caseId}: ${err}`)
    }
}

async function writeCase(client: Client, horsePath: string, sessionId: string, caseData: CaseData, rowInfo: RowInfo) {
    function excelDate(date: Date) {
        return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`
    }
    try {
        let rangeValues: Array<number | string | null> = [
            caseData.id,
            caseData?.client ? caseData.client : null,
            caseData?.pledgeDate ? excelDate(caseData.pledgeDate) : null,
            caseData?.appointmentDate ? excelDate(caseData.appointmentDate) : null,
            caseData?.clinic ? caseData.clinic : null,
            caseData?.pledgeAmount ? caseData.pledgeAmount : null,
            caseData?.invoiceStatus ? caseData.invoiceStatus : null,
            caseData?.contact ? caseData.contact : null,
        ]
        let rangeAddress = `A${rowInfo.row}:H${rowInfo.row}`
        let opPath = `/workbook/worksheets/Cases/range(address='${rangeAddress}')`;
        let update = await client.api(`${horsePath}/${opPath}`)
            .header('workbook-session-id', sessionId)
            .select(['address', 'values'])
            .patch({values: [rangeValues]})
        // console.log(update)
    } catch (err) {
        throw Error(`Can't update case ${caseData.id}: ${err}`);
    }
}

async function openSession(client: Client, horsePath: string): Promise<string> {
    try {
        let result = await client.api(`${horsePath}/workbook/createSession`)
            .post({persistChanges: true});
        if (result.id !== undefined) {
            console.log(`Created workbook session ID ${md5(result.id)}`)
            // console.log(result)
            return result.id
        }
    } catch (err) {
        throw Error(`Failed to open session: ${err}`)
    }
    throw Error(`No session ID was returned`);
}

async function closeSession(client: Client, horsePath: string, sessionId: string) {
    try {
        let result = await client.api(`${horsePath}/workbook/closeSession`)
            .header('workbook-session-id', sessionId)
            .post({});
        console.log(`Closed workbook session ID ${md5(sessionId)}`)
        // console.log(result)
    } catch (err) {
        throw Error(`Failed to close session: ${err}`);
    }
}
