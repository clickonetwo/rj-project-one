"use strict";
// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.
Object.defineProperty(exports, "__esModule", { value: true });
exports.discoverHorseId = void 0;
require("isomorphic-fetch");
const settings_1 = require("./local/settings");
async function discoverHorseId(client, horseName) {
    let horseId = '';
    try {
        let driveId = settings_1.devData.driveId;
        let listing = await client.api(`/drives/${driveId}/root/children`)
            .select(['createdDateTime', 'lastModifiedDateTime', 'name', 'id'])
            .get();
        // console.log(listing)
        for (let child of listing.value) {
            console.log(`Processing root child with name ${child.name} and id ${child.id}`);
            if (child.name == 'Spreadsheets') {
                console.log(`Fetching spreadsheets...`);
                listing = await client.api(`/drives/${driveId}/items/${child.id}/children`)
                    .select(['createdDateTime', 'lastModifiedDateTime', 'name', 'id'])
                    .get();
                // console.log(listing)
                for (let sheet of listing.value) {
                    console.log(`Processing spreadsheet with name ${sheet.name} and id ${sheet.id}`);
                    if (sheet.name == `${horseName}.xlsx`) {
                        horseId = sheet.id;
                    }
                }
            }
        }
    }
    catch (err) {
        console.log(`Error discovering horseId: ${err}`);
    }
    return horseId;
}
exports.discoverHorseId = discoverHorseId;
//# sourceMappingURL=discovery.js.map