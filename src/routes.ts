// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import {getClientData} from "./settings";
import {initializeGraphClient} from "./graphClient";
import {CaseData, updateCase} from "./case";
import express from "express";

export async function statusHandler(req: express.Request, res: express.Response) {
    res.status(200).send({ status: 'success' });
}

export async function updateCaseHandler(req: express.Request, res: express.Response) {
    const caseData: CaseData = req.body
    if (caseData?.id) {
        try {
            const rowData = await update(caseData);
            const result = rowData.isNew ?
                `Inserted case ${caseData.id} at row ${rowData.row}` :
                `Updated case ${caseData.id} at row ${rowData.row}`;
            res.status(200).send({status: 'success', result: result});
        }
        catch (err) {
            res.status(500).send({status: 'error', reason: err});
        }
    } else {
        res.status(400).send({status: 'error', reason: 'Update request must specify the id field'});
    }
}

async function update(caseData: CaseData) {
    const clientData = getClientData();
    clientData.client = initializeGraphClient(clientData);
    return await updateCase(clientData, caseData);
}

