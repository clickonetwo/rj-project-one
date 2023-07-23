// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

let fs = require('fs');

if (process.argv[2] === "dev") {
    fs.copyFileSync('local/dev.env', '.env');
    console.log("Installed ClickOneTwo environment.");
} else {
    fs.copyFileSync('local/arc.env', '.env')
    console.log("Installed ARC environment.");
}
