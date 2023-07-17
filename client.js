"use strict";
// Copyright 2023 Daniel C. Brotsky. All rights reserved.
// Licensed under the GNU Affero General Public License v3.
// See the LICENSE file for details.
//
// Portions of this code may be excerpted under MIT license
// from SDK samples provided by Microsoft.
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.initializeGraphClient = void 0;
const settings_1 = require("./local/settings");
require("isomorphic-fetch");
const azure = __importStar(require("@azure/identity"));
const authProviders = __importStar(require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials"));
const graph = __importStar(require("@microsoft/microsoft-graph-client"));
function initializeGraphClient() {
    let clientSecretCredential = new azure.ClientSecretCredential(settings_1.devData.tenantId, settings_1.devData.clientId, settings_1.devData.clientSecret);
    const authProvider = new authProviders.TokenCredentialAuthenticationProvider(clientSecretCredential, {
        scopes: ['https://graph.microsoft.com/.default']
    });
    return graph.Client.initWithMiddleware({
        authProvider: authProvider
    });
}
exports.initializeGraphClient = initializeGraphClient;
//# sourceMappingURL=client.js.map