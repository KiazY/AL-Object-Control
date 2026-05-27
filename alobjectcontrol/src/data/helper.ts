import * as fs from 'fs';
import { Uri, window, workspace } from 'vscode';

export async function createConfigFile() {
    const fileName: string = '.object-control.json';
    const URIs: Uri[] = await workspace.findFiles('**/.object-control.json', null, 1);
    if (URIs.length === 0) {
        if (workspace.workspaceFolders !== undefined) {
            const current_path = workspace.workspaceFolders[0].uri.fsPath;
            fs.writeFile(`${current_path}//.vscode//.object-control.json`, `{
    "tenantId": "<tenantId>",
    "environmentName": "TESTPT",
    "clientId": "<clientId>",
    "APIPublisher": "<APIPublisher>",
    "APIGroup": "<APIGroup>",
    "APIVersion": "v1.0",
    "EntitySetName_AllObjects": "allObjs",
    "EntitySetName_ReservedObjects": "reservedObjects",
    "range": {
        "from": 50000,
        "to": 59999
    }
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