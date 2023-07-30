// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import express from 'express';
import cors from 'cors';

import {statusHandler, postUpdateHandler, getUpdateHandler} from "./routes"
import {authMiddleware} from "./serverAuth";

const PORT = process.env.PORT || 5001;

express()
    .use(express.json())
    .use(express.urlencoded({extended: true}))
    .use(cors())
    .use(authMiddleware)
    .get('/status', statusHandler)
    .get('/update', getUpdateHandler)
    .post('/update', postUpdateHandler)
    .listen(PORT, () => console.log(`Listening on ${ PORT }`))
