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

export async function statusHandler(req: express.Request, res: express.Response) {
    res.status(200).send({status: 'success'});
}

export function getUpdateHandler(req: express.Request, res: express.Response) {
    const caseData = prepareCase(req.query);
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

function prepareCase(submission: object): CaseData | string {
    if ('id' in submission) {
        return submission as CaseData;
    } else {
        return `No id field found in submitted object ${submission}`;
    }
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
