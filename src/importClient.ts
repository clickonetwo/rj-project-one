// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import {readFileSync} from 'fs';
import {parse} from 'csv-parse/sync'

import {prepareCase, QueryString} from './routes';
import {CaseData, updateMultipleCases} from './case';
import {getClientData} from "./settings";
import {initializeGraphClient} from "./graphClient";

type CaseExport = { [index: string]: string };

function readCaseExports(path: string) {
    return parse(
        readFileSync(path),
        {
            bom: true,
            cast: true,
            columns: true,
        }) as unknown as CaseExport[];
}

const COLUMN_MAP = {
    id: 'Case Number',
    client: 'Contact Name in Excel',
    pledgeDate: 'Invoice Sent Date',
    appointmentDate: 'Appointment Date',
    clinic: 'Clinic Name',
    reason: 'Case Reason',
    pledgeAmount: 'Total Pledge',
    invoiceStatus: 'Invoice Status',
    contact: 'Case Owner in Excel',
};

function mapExportsToQueries(records: CaseExport[]): QueryString[] {
    const result: QueryString[] = [];
    for (const record of records) {
        result.push({
            id: record[COLUMN_MAP.id],
            client: record[COLUMN_MAP.client],
            pledgeDate: record[COLUMN_MAP.pledgeDate],
            appointmentDate: record[COLUMN_MAP.appointmentDate],
            clinic: record[COLUMN_MAP.clinic] || record[COLUMN_MAP.reason],
            pledgeAmount: record[COLUMN_MAP.pledgeAmount],
            invoiceStatus: record[COLUMN_MAP.invoiceStatus],
            contact: record[COLUMN_MAP.contact],
        });
    }
    return result;
}

function mapQueriesToCases(queries: QueryString[]): CaseData[] {
    const result: CaseData[] = [];
    for (const query of queries) {
        const caseData = prepareCase(query);
        if (typeof caseData === 'string') {
            throw Error(`Illegal input - ${caseData}:\n${JSON.stringify(query, null, 2)}`)
        }
        result.push(caseData);
    }
    return result;
}

async function importCases(path: string) {
    const exports = readCaseExports(path);
    console.log(`Read ${exports.length} exports from Salesforce`);
    const queries = mapExportsToQueries(exports);
    console.log(`Converted all exports to queries`)
    const cases = mapQueriesToCases(queries);
    console.log(`Converted all queries to cases`)
    const clientData = getClientData();
    clientData.client = initializeGraphClient(clientData);
    console.log(`Updating Excel from cases`);
    await updateMultipleCases(clientData, cases);
    console.log(`Updated all ${exports.length} cases`)
}

const path = process.argv[2];
if (path) {
    importCases(path)
        .then(() => {
            console.log('Import complete');
        });
} else {
    throw Error(`You must specify the path`)
}