import { ExtensionContext, window } from "vscode";
import { getAppJsonFile, getConfigurationFile } from "./get-data";

async function objectAlreadyReserved(url: string, headers: Headers, objectName: string): Promise<boolean> {
    const escapedObjectName = objectName.replace(/'/g, "''");
    const queryUrl = `${url}?$select=Id,ObjectName&$filter=ObjectName eq '${escapedObjectName}'`;
    const response = await fetch(queryUrl, { headers });

    if (!response.ok) {
        throw new Error(`${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    const items = data.value ?? data.d?.results ?? [];
    return items.length > 0;
}

async function getListItemEntityTypeFullName(listBaseUrl: string, headers: Headers): Promise<string> {
    const response = await fetch(`${listBaseUrl}?$select=ListItemEntityTypeFullName`, { headers });
    if (!response.ok) {
        throw new Error(`${response.status} ${response.statusText}`);
    }
    const data = await response.json();
    return data.ListItemEntityTypeFullName ?? data.d?.ListItemEntityTypeFullName;
}

async function createReserveItem(listBaseUrl: string, url: string, headers: Headers, idToReserve: number, objectType: string) {
    const objectName = `${objectType} ${idToReserve}`;
    const app_json = await getAppJsonFile();
    if (await objectAlreadyReserved(url, headers, objectName)) {
        window.showInformationMessage('ID already reserved.');
        return;
    }

    const entityType = await getListItemEntityTypeFullName(listBaseUrl, headers);
    const response = await fetch(url,
        {
            headers: headers,
            method: 'POST',
            body: JSON.stringify({
                __metadata:
                {
                    type: entityType
                },
                ObjectID: idToReserve, ObjectType: objectType, Title: app_json.appName, ObjectName: objectName
            })
        }
    );

    if (!response.ok) {
        window.showInformationMessage(`ID was not reserved. ${response.status}: ${response.statusText}`);
        throw new Error(`${response.status} ${response.statusText}`);
    }

    window.showInformationMessage('ID reserved successfully');
}

export async function reserveId(context: ExtensionContext, idToReserve: number, objectType: string) {
    const config_file = await getConfigurationFile();
    const headers = new Headers();
    headers.append('Authorization', `Bearer ${await context.secrets?.get('tokenSp')}`);
    headers.append('Accept', 'application/json;odata=verbose');
    headers.append('Content-Type', 'application/json;odata=verbose');
    const listBaseUrl = `${config_file.sharepoint_siteUrl}/_api/web/lists/getbytitle('${config_file.sharepoint_listName}')`;
    const url = `${listBaseUrl}/items`;
    await fetch(url,
        {
            headers: headers
        }).then(async (resolve) => {
            if (resolve.ok) {
                await createReserveItem(listBaseUrl, url, headers, idToReserve, objectType);
            } else if (resolve.status === 404) {
                const choice = await window.showWarningMessage(
                    `The SharePoint list '${config_file.sharepoint_listName}' doesn't exist yet. Do you want to create it?`,
                    'Yes', 'No'
                );
                if (choice === 'Yes') {
                    await createSharepointList(headers);
                    await createReserveItem(listBaseUrl, url, headers, idToReserve, objectType);
                } else {
                    window.showInformationMessage('Object ID was not reserved. The SharePoint list was not created.');
                }
            }
        }, (reject) => {
            console.error(reject);
            window.showErrorMessage('Object ID was not reserved due to an error. Check the console for more information.');
            throw reject;
        });
}

async function createSharepointList(headers: Headers) {
    try {
        const config_file = await getConfigurationFile();
        const create_list_json = JSON.stringify(
            {
                __metadata: {
                    type: "SP.List"
                },
                BaseTemplate: 100,
                Title: config_file.sharepoint_listName,
                Description: "Object Control List created by AL Object Control VSCode Extension",
                OnQuickLaunch: true
            });
        const url = `${config_file.sharepoint_siteUrl}/_api/web/lists`;
        await fetch(url,
            { headers: headers, body: create_list_json, method: 'POST' }
        ).then(async (resolve) => {
            if (resolve.ok) {
                console.log(`Sharepoint list created with name ${config_file.sharepoint_listName}`);
                await createSharepointField(headers, 'ObjectType', 2, 'SP.FieldText', false, false);
                await createSharepointField(headers, 'ObjectID', 1, 'SP.FieldNumber', false, false);
                await createSharepointField(headers, 'ObjectName', 2, 'SP.FieldText', true, true);
                await addSharepointFieldToView(headers, 'Created');
            } else { console.log(resolve.status + ' ' + resolve.statusText); console.log(resolve.text); }
        }, (reject) => {
            console.error(reject);
        });
    } catch (error) {
        throw error;
    }
}

async function createSharepointField(headers: Headers, fieldName: string, fieldType: number, fieldMetadata: string, enforceUniqueValues: boolean, index: boolean) {
    // Create Sharepoint list fields
    const config_file = await getConfigurationFile();
    const url = `${config_file.sharepoint_siteUrl}/_api/web/lists/GetByTitle('${config_file.sharepoint_listName}')/fields`;
    const fields_json = JSON.stringify({
        __metadata: {
            type: `${fieldMetadata}`
        },
        Title: fieldName,
        FieldTypeKind: fieldType,
        EnforceUniqueValues: enforceUniqueValues,
        Indexed: index,
        StaticName: fieldName
    });
    await fetch(url, { headers: headers, method: 'POST', body: fields_json }).then(async (resolve) => {
        if (resolve.ok) {
            await addSharepointFieldToView(headers, fieldName);
            console.log(`${fieldName} field created.`);
        }
    }, (reject) => {
        console.error(reject);
    });
}

async function addSharepointFieldToView(headers: Headers, fieldInternalName: string) {
    const config = await getConfigurationFile();
    const url = `${config.sharepoint_siteUrl}/_api/web/lists/GetByTitle('${config.sharepoint_listName}')/DefaultView/ViewFields/AddViewField('${fieldInternalName}')`;
    const response = await fetch(url, {
        method: 'POST',
        headers
    });

    if (!response.ok) {
        throw new Error(await response.text());
    }

    console.log(`Field ${fieldInternalName} added to default view.`);
}