// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import 'isomorphic-fetch';
import md5 from 'crypto-js/md5';

import {Client} from '@microsoft/microsoft-graph-client';

import {ClientData} from "./settings";

export interface CaseData {
    id: number,
    client?: string,
    pledgeDate?: string,
    appointmentDate?: string,
    clinic?: string,
    pledgeAmount?: number,
    invoiceStatus?: string,
    contact?: string,
    [index: string]: string | number | undefined,
}

export interface RowInfo {
    row: number,
    isNew: boolean,
}

export function rowUrl(clientData: ClientData, rowData: RowInfo) {
    const SITE_URL= `https://arcse.sharepoint.com/:x:/r/sites/healthline`
    const NAME = encodeURIComponent(clientData.horseName);
    const PATH = `/Shared%20Documents/Spreadsheets/${NAME}.xlsx`
    const QUERY_PREFIX = '?web=1&nav='
    const row = rowData.row;
    const navParam = Buffer.from(`12_A${row}:H${row}_{00000000-0001-0000-0000-000000000000}`, 'utf8');
    const navParamEncoded = navParam.toString('base64url');
    return SITE_URL + PATH + QUERY_PREFIX + navParamEncoded;
}

export async function updateCase(clientData: ClientData, caseData: CaseData) {
    const horsePath = `/drives/${clientData.driveId}/items/${clientData.horseId}`
    const sessionId = await openSession(clientData.client!, horsePath);
    const rowInfo = await findCase(clientData.client!, horsePath, sessionId, caseData.id);
    await writeCase(clientData.client!, horsePath, sessionId, caseData, rowInfo)
    await closeSession(clientData.client!, horsePath, sessionId);
    return rowInfo;
}

async function findCase(client: Client, horsePath: string, sessionId: string, caseId: number): Promise<RowInfo> {
    try {
        // First, get the filled range
        const range = await client.api(`${horsePath}/workbook/worksheets/Cases/usedRange(valuesOnly=true)`)
            .header('workbook-session-id', sessionId)
            .select(['address', 'columnIndex', 'columnCount', 'rowIndex', 'rowCount'])
            .get()
        // console.log(range)
        // Next, search for the case in the first column of that range
        //
        // Note: Excel row numbers start at 1, but the returned rowIndex starts at 0
        // Since the end of the range *includes* the starting row, we don't add 1 there.
        const opBody = {
            lookupValue: caseId,
            lookupArray: {
                address: `Cases!A${range.rowIndex + 1}:A${range.rowIndex + range.rowCount}`
            },
            matchType: 0,
        }
        const found = await client.api(`${horsePath}/workbook/functions/match`)
            .header('workbook-session-id', sessionId)
            .post(opBody);
        // console.log(found)
        if (found.error === null) {
            console.log(`Found existing ${caseId} in cell A${found.value + 1}`)
            return {row: found.value + 1, isNew: false}
        } else {
            const newRow = range.rowIndex + range.rowCount + 1
            console.log(`Inserting new case ${caseId} into cell A${newRow}`)
            return {row: newRow, isNew: true}
        }
    } catch (err) {
        throw Error(`Can't find case ${caseId}: ${err}`)
    }
}

async function writeCase(client: Client, horsePath: string, sessionId: string, caseData: CaseData, rowInfo: RowInfo) {
    try {
        const rangeValues: Array<number | string | null> = [
            caseData.id,
            caseData?.client ? caseData.client : null,
            caseData?.pledgeDate ? caseData.pledgeDate : null,
            caseData?.appointmentDate ? caseData.appointmentDate : null,
            caseData?.clinic ? caseData.clinic : null,
            caseData?.pledgeAmount ? caseData.pledgeAmount : null,
            caseData?.invoiceStatus ? caseData.invoiceStatus : null,
            caseData?.contact ? caseData.contact : null,
        ]
        const rangeAddress = `A${rowInfo.row}:H${rowInfo.row}`
        const opPath = `/workbook/worksheets/Cases/range(address='${rangeAddress}')`;
        const _update = await client.api(`${horsePath}/${opPath}`)
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
        const result = await client.api(`${horsePath}/workbook/createSession`)
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
        const _result = await client.api(`${horsePath}/workbook/closeSession`)
            .header('workbook-session-id', sessionId)
            .post({});
        console.log(`Closed workbook session ID ${md5(sessionId)}`)
        // console.log(result)
    } catch (err) {
        throw Error(`Failed to close session: ${err}`);
    }
}
