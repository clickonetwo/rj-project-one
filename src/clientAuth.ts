// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import {TOTP} from 'otpauth';

function fromSecret(secret: string): TOTP {
    return new TOTP({
        issuer: "rj-project-1",
        digits: 10,
        period: 30,
        secret,
    });
}

export default function tokenFromSecret(secret: string): string {
    const totp = fromSecret(secret);
    return totp.generate();
}
