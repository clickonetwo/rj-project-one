// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import 'isomorphic-fetch';

import {Client} from '@microsoft/microsoft-graph-client';

import {devData as clientData} from './local/settings';

interface driveItem {
    id: string,
    name: string,
}

interface driveListing {
    value: driveItem[],
}

export async function discoverHorseId(client: Client, horseName: string) {
    let horseId = '';
    try {
        let driveId = clientData.driveId;
        let listing: driveListing = await client.api(`/drives/${driveId}/root/children`)
            .select(['createdDateTime', 'lastModifiedDateTime', 'name', 'id'])
            .get();
        // console.log(listing)
        for (let child of listing.value) {
            console.log(`Processing root child with name ${child.name} and id ${child.id}`);
            if (child.name == 'Spreadsheets') {
                console.log(`Fetching spreadsheets...`)
                listing = await client.api(`/drives/${driveId}/items/${child.id}/children`)
                    .select(['createdDateTime', 'lastModifiedDateTime', 'name', 'id'])
                    .get()
                // console.log(listing)
                for (let sheet of listing.value) {
                    console.log(`Processing spreadsheet with name ${sheet.name} and id ${sheet.id}`)
                    if (sheet.name == `${horseName}.xlsx`) {
                        horseId = sheet.id
                    }
                }
            }
        }
    } catch (err) {
        console.log(`Error discovering horseId: ${err}`)
    }
    return horseId
}
