// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import {ClientData} from './settings';

interface driveItem {
    id: string,
    name: string,
}

interface driveListing {
    value: driveItem[],
}

export async function discoverHorseId(clientData: ClientData, horseName: string) {
    if (clientData.horseId) {
        return clientData.horseId;
    }
    if (!clientData.groupId && !clientData.driveId) {
        throw Error("Can't discover horseId without either a groupId or a driveId")
    }
    let horseId: string = '';
    try {
        const client = clientData.client!;
        let driveId = clientData.driveId;
        if (!driveId) {
            const groupId = clientData.groupId;
            const drives: driveListing = await client.api(`/groups/${groupId}/sites/root/drives`)
                .select(['createdDateTime', 'lastModifiedDateTime', 'name', 'id'])
                .get();
            driveId = drives.value[0].id;
            console.log(`Drive ID for group is '${driveId}'`)
        }
        let listing: driveListing = await client.api(`/drives/${driveId}/root/children`)
            .select(['createdDateTime', 'lastModifiedDateTime', 'name', 'id'])
            .get();
        // console.log(listing)
        for (const child of listing.value) {
            console.log(`Processing root child with name ${child.name} and id ${child.id}`);
            if (child.name == 'Spreadsheets') {
                console.log(`Fetching spreadsheets...`)
                listing = await client.api(`/drives/${driveId}/items/${child.id}/children`)
                    .select(['createdDateTime', 'lastModifiedDateTime', 'name', 'id'])
                    .get()
                // console.log(listing)
                for (const sheet of listing.value) {
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
    if (!horseId) {
        throw Error("Can't find id for horse named 'horseName'")
    }
    return horseId
}
