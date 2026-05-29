import * as fs from 'fs';
import { Uri, window, workspace } from 'vscode';

export async function createConfigFile() {
    const URIs: Uri[] = await workspace.findFiles('**/.object-control.json', null, 1);
    if (URIs.length === 0) {
        if (workspace.workspaceFolders !== undefined) {
            const current_path = workspace.workspaceFolders[0].uri.fsPath;
            fs.writeFile(`${current_path}//.vscode//.object-control.json`, `{
    "tenantId": "<tenantId>",
    "environmentName": "<environmentName>",
    "clientId": "<clientId>",
    "APIPublisher": "<APIPublisher>",
    "APIGroup": "<APIGroup>",
    "APIVersion": "v1.0",
    "entitySetName": {
        "allObjects": "allObjs",
        "reservedObjects": "reservedObjects"
    },
    "range": {
        "from": 50000,
        "to": 59999
    },
    "expiresIn": 1
}`,
                { encoding: 'utf-8' }, (err: NodeJS.ErrnoException | null) => {
                    if (err !== null) {
                        throw err;
                    }
                });
        }
    } else {
        window.showErrorMessage(`File already exists.`);
    }
}