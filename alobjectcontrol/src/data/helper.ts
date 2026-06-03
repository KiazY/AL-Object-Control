import * as fs from 'fs';
import { ExtensionContext, Uri, window, workspace, Disposable } from 'vscode';
import { getConfigurationFile } from './get-data';
import { reserveId } from './post-data';

export async function createConfigFile() {
    const URIs: Uri[] = await workspace.findFiles('**/.object-control.json', null, 1);
    if (URIs.length === 0) {
        if (workspace.workspaceFolders !== undefined) {
            const current_path = workspace.workspaceFolders[0].uri.fsPath;
            fs.writeFile(`${current_path}//.vscode//.object-control.json`, `{
    "tenantId": "<tenantId>",
    "clientId": "<clientId>",
    "sharepointBaseUrl": <sharepointBaseUrl>,
    "sharepointSiteName": <sharepointSiteName>,
    "sharepointListName": <sharepointListName>,
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

function findObjectFromContent(fileContent: string): { objectType: string; objectId: number } | undefined {
    const lines = fileContent.split(/\r?\n/);
    const declarationRegex = /\b(table|tableextension|page|pageextension|report|reportextension|codeunit|xmlport|query|enum|enumextension)\s+(\d+)\b/i;

    for (const line of lines) {
        const match = line.match(declarationRegex);
        if (match !== null) {
            return {
                objectType: match[1].toLowerCase(),
                objectId: Number(match[2])
            };
        }
    }

    return undefined;
}

async function deleteSharePointItemByTitle(context: ExtensionContext, objectType: string, objectId: number) {
    const token = await context.secrets.get('tokenSp');
    if (!token) {
        return;
    }

    const config = await getConfigurationFile();
    const headers = new Headers();
    headers.append('Authorization', `Bearer ${token}`);
    headers.append('Accept', 'application/json;odata=nometadata');
    headers.append('Content-Type', 'application/json;odata=nometadata');
    headers.append('IF-MATCH', '*');

    const objectTitle = `${objectType} ${objectId}`;
    const escapedTitle = objectTitle.replace(/'/g, "''");
    const listBaseUrl = `${config.sharepoint_baseUrl}/sites/${config.sharepoint_siteName}/_api/web/lists/getbytitle('${config.sharepoint_listName}')`;
    const queryUrl = `${listBaseUrl}/items?$select=Id,ObjectName&$filter=ObjectName eq '${escapedTitle}'`;

    const queryResponse = await fetch(queryUrl, { headers });
    if (!queryResponse.ok) {
        console.log(`Couldn't find item in Sharepoint: ${queryResponse.status} ${queryResponse.statusText}`);
        return;
    }

    const queryData = await queryResponse.json();
    const itemId = queryData.value?.[0]?.Id;
    if (!itemId) {
        return;
    }

    const deleteUrl = `${listBaseUrl}/items(${itemId})`;
    const deleteResponse = await fetch(deleteUrl, {
        method: 'DELETE',
        headers
    });

    if (!deleteResponse.ok) {
        console.log(`Couldn't delete '${objectTitle}' from SharePoint: ${deleteResponse.status} ${deleteResponse.statusText}`);
        return;
    }

    console.log(`Item '${objectTitle}' removed from Sharepoint list.`);
}

export function releaseIdListener(context: ExtensionContext): Disposable {
    return workspace.onWillDeleteFiles((event) => {
        for (const file of event.files) {
            void (async () => {
                try {
                    const fileContent = await fs.promises.readFile(file.fsPath, 'utf8');
                    const deletedObject = findObjectFromContent(fileContent);
                    if (!deletedObject) {
                        return;
                    }

                    await deleteSharePointItemByTitle(context, deletedObject.objectType, deletedObject.objectId);
                } catch (error) {
                    console.log(`Couldn't process delete file '${file.fsPath}'.`);
                    console.error(error);
                }
            })();
        }
    });
}

export function saveFileListener(context: ExtensionContext): Disposable {
    return workspace.onWillSaveTextDocument((e) => {
        if (e.document.languageId !== 'al') {
            return;
        }

        void (async () => {
            const savedObject = findObjectFromContent(e.document.getText());
            if (!savedObject) {
                return;
            }

            try {
                const previousContent = await fs.promises.readFile(e.document.uri.fsPath, 'utf8');
                const previousObject = findObjectFromContent(previousContent);

                if (previousObject && (previousObject.objectId !== savedObject.objectId || previousObject.objectType !== savedObject.objectType)) {
                    await deleteSharePointItemByTitle(context, previousObject.objectType, previousObject.objectId);
                }
            } catch (error) {
                // If there is no prior saved file version, keep going and reserve current object.
                console.log(`Previous version could not be read. '${e.document.uri.fsPath}'.`);
                console.error(error);
            }

            await reserveId(context, savedObject.objectId, savedObject.objectType);
        })();
    });
}