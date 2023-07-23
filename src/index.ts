// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import express from 'express';
import {statusHandler, updateCaseHandler} from "./routes"

const PORT = process.env.PORT || 5001;

express()
    .use(express.json())
    .use(express.urlencoded({extended: true}))
    .get('/status', statusHandler)
    .post('/update', updateCaseHandler)
    .listen(PORT, () => console.log(`Listening on ${ PORT }`))
