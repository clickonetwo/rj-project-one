// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import {getClientData} from "./settings";
import {initializeGraphClient} from "./graphClient";
import {CaseData, rowUrl, updateCase} from "./case";
import express from 'express';

export type QueryString = { [index: string]: number | string | unknown }

export async function statusHandler(_req: express.Request, res: express.Response) {
    res.status(200).send({status: 'success'});
}

export function getUpdateHandler(req: express.Request, res: express.Response) {
    const caseData = prepareCase(req.query as unknown as QueryString);
    if (typeof caseData === 'string') {
        res.status(400).send({status: 'error', reason: caseData});
    } else {
        performUpdate(res, caseData);
    }
}

export function postUpdateHandler(req: express.Request, res: express.Response) {
    const caseData = prepareCase(req.body);
    if (typeof caseData === 'string') {
        res.status(400).send({status: 'error', reason: caseData});
    } else {
        performUpdate(res, caseData);
    }
}

export function prepareCase(submission: QueryString): CaseData | string {
    const caseData: CaseData = {id: 0};
    if (!('id' in submission)) {
        return `No id field found in submitted object ${JSON.stringify(submission)}`;
    }
    let caseNumber = submission.id;
    if (typeof caseNumber === 'string' && /^[0-9]+$/.test(caseNumber)) {
        caseNumber = parseInt(caseNumber);
    }
    if (typeof caseNumber !== 'number' || caseNumber <= 0) {
        return `id (${JSON.stringify(caseNumber)} must be a positive integer`;
    }
    caseData.id = caseNumber;
    let pledgeAmount = submission?.pledgeAmount;
    if (pledgeAmount) {
        if (typeof pledgeAmount === 'string' && /^[0-9]+$/.test(pledgeAmount)) {
            pledgeAmount = parseInt(pledgeAmount);
        }
        if (typeof pledgeAmount !== 'number' || pledgeAmount < 0) {
            return `pledgeAmount (${JSON.stringify(pledgeAmount)} must be a non-negative integer`;
        }
        caseData.pledgeAmount = pledgeAmount;
    }
    const pledgeDate = submission?.pledgeDate;
    if (pledgeDate) {
        const pattern1 = /20[0-9][0-9]-[0-9][0-9]-[0-9][0-9]/;
        const pattern2 = /[0-9]?[0-9]\/[0-9]?[0-9]\/20[0-9][0-9]/;
        if (typeof pledgeDate !== 'string') {
            return `pledgeDate (${JSON.stringify(pledgeDate)}) must be a string`;
        }
        if (!pattern1.test(pledgeDate) && !pattern2.test(pledgeDate)) {
            return `pledgeDate (${JSON.stringify(pledgeDate)}) must be a date in Excel form (m/d/y or y-m-d)`;
        }
        caseData.pledgeDate = pledgeDate;
    }
    const appointmentDate = submission?.appointmentDate;
    if (appointmentDate) {
        const pattern1 = /20[0-9][0-9]-[0-9][0-9]-[0-9][0-9]/;
        const pattern2 = /[0-9]?[0-9]\/[0-9]?[0-9]\/20[0-9][0-9]/;
        if (typeof appointmentDate !== 'string') {
            return `appointmentDate (${JSON.stringify(appointmentDate)}) must be a string`;
        }
        if (!pattern1.test(appointmentDate) && !pattern2.test(appointmentDate)) {
            return `appointmentDate (${JSON.stringify(appointmentDate)}) must be a date in Excel form (m/d/y or y-m-d)`;
        }
        caseData.appointmentDate = appointmentDate;
    }
    for (const propName of ['client', 'clinic', 'invoiceStatus', 'contact']) {
        const value = submission[propName];
        if (value && typeof value === 'string') {
            caseData[propName] = value;
        }
    }
    return caseData;
}

function performUpdate(res: express.Response, caseData: CaseData) {
    try {
        const clientData = getClientData();
        clientData.client = initializeGraphClient(clientData);
        updateCase(clientData, caseData).then((rowData) => {
            const result = rowData.isNew ?
                `Inserted case ${caseData.id} at row ${rowData.row}` :
                `Updated case ${caseData.id} at row ${rowData.row}`;
            console.log(result);
            res.status(200).send({status: 'success', result, url: rowUrl(clientData, rowData)});
        });
    } catch (err) {
        res.status(500).send({status: 'error', reason: err});
    }
}
