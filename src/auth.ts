// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import * as OTPAuth from 'otpauth';
import express from 'express';

import {getClientData} from "./settings";

export function newSecret(): string {
    const secret = new OTPAuth.Secret()
    return secret.base32
}

function fromSecret(secret: string): OTPAuth.TOTP {
    return new OTPAuth.TOTP({
        issuer: "rj-project-1",
        digits: 8,
        period: 30,
        secret,
    });
}

export function tokenFromSecret(secret: string): string {
    const totp = fromSecret(secret);
    return totp.generate();
}

export function validateTokenAgainstSecret(secret: string, token: string, window = 1) {
    const totp = fromSecret(secret);
    const delta = totp.validate({token, window})
    return delta !== null
}

const is_production = process.env.NODE_ENV === 'production';

export function authMiddleware(req: express.Request, res: express.Response, next: express.NextFunction) {
    const authValue = req.get("Authorization")
    if (!authValue) {
        res.status(401).send({status: 'error', error: 'Authorization required but not provided'})
        return;
    }
    const clientData = getClientData();
    if (validateTokenAgainstSecret(clientData.totpSecret, authValue, is_production ? 1 : 1000) === null) {
        res.status(403).send({status: 'error', error: 'Authorization failed'});
    }
    next();
}