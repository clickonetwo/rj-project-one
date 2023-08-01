// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import express from 'express';

import {getClientData} from "./settings";

export function tokenFromContent(secret: string, content: string) {
    // uses the Java string hash function, cribbed from here: https://stackoverflow.com/a/8831937/558006
    function hashCode(str: string) {
        let hash = 0;
        for (let i = 0, len = str.length; i < len; i++) {
            const chr = str.charCodeAt(i);
            hash = (hash << 5) - hash + chr;
            hash |= 0; // Convert to 32bit integer
        }
        return hash;
    }
    return hashCode(`${content}|${secret}`).toString();
}

export function validateToken(secret: string, token: string) {
    const now = new Date().toISOString()
    const dateStrings = getDateStrings();
    const correctToken = tokenFromContent(secret, dateStrings.last);
    if (token === correctToken) {
        return true;
    }
    const priorToken = tokenFromContent(secret, dateStrings.prior);
    if (token !== priorToken) {
        console.log(`Validation failure for request at ${now}`)
        console.log(`Received token '${token}' from request`)
        console.log(`Last minute is ${dateStrings.last}`)
        console.log(`Last token is '${correctToken}'`)
        console.log(`Prior minute is ${dateStrings.prior}`)
        console.log(`Prior token is '${priorToken}'`);
        return false;
    }
    return true;
}

function getDateStrings() {
    const now = new Date();
    const lastMinute = new Date(now);
    lastMinute.setSeconds(0);
    lastMinute.setMilliseconds(0);
    const priorMinute = new Date(lastMinute.valueOf() - 60*1000);
    return { prior: priorMinute.toISOString(), last: lastMinute.toISOString() };
}

export function authMiddleware(req: express.Request, res: express.Response, next: express.NextFunction) {
    const token = req.get("X-Salesforce-Token")
    if (!token) {
        res.status(403).send({status: 'error', error: 'Salesforce token required but not provided'})
        return;
    }
    const secret = getClientData().authSecret;
    if (!validateToken(secret, token)) {
        res.status(403).send({status: 'error', error: 'Salesforce token is invalid'});
        return;
    }
    next();
}
