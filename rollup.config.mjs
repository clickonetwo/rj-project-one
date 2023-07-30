// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.

import terser from "@rollup/plugin-terser";
import commonjs from '@rollup/plugin-commonjs';
import nodeResolve from '@rollup/plugin-node-resolve';

export default {
    input: 'dist/clientAuth.js',
    output: [{
        name: 'tokenFromSecret',
        file: 'local/clientAuth.bundle.js',
        format: 'umd'
    }, {
        name: 'tokenFromSecret',
        file: 'local/clientAuth.bundle.min.js',
        format: 'umd',
        plugins: [terser()]
    }
    ],

    plugins: [
        commonjs(),
        nodeResolve({
            preferBuiltins: false,
        })
    ]
}