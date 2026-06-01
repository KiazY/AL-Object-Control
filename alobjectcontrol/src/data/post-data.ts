import { ExtensionContext, window } from "vscode";
import { getConfigurationFile } from "./get-data";

export async function reserveId(context: ExtensionContext, idToReserve: number, objectType: string) {
    const config_file = await getConfigurationFile();
    const headers = new Headers();
    headers.append('Authorization', `Bearer ${await context.secrets?.get('tokenSp')}`);
    headers.append('Accept', 'application/json;odata=verbose');
    headers.append('Content-Type', 'application/json;odata=verbose');
    const url = `${config_file.sharepoint_baseUrl}/sites/${config_file.sharepoint_siteName}/_api/web/lists/getbytitle('${config_file.sharepoint_listName}')/items`;
    await fetch(url,
        {
            headers: headers
        }).then(async (resolve) => {
            if (resolve.ok) {
                // Add item to sharepoint list
                await fetch(url,
                    {
                        headers: headers,
                        method: 'POST',
                        body: JSON.stringify({
                            "__metadata":
                            {
                                "type": "SP.Data.Object_x0020_ControlListItem"
                            },
                            "Object_x0020_ID": idToReserve, "Object_x0020_Type": objectType, "Title": `${objectType} ${idToReserve}`
                        })
                    }
                ).then(async (resolve) => {
                    if (resolve.ok) {
                        window.showInformationMessage('ID reserved successfully');
                    } else {
                        window.showInformationMessage(`ID was not reserved. ${resolve.status}: ${resolve.statusText} `);
                    }
                }, (reject) => {
                    console.error(reject);
                    window.showErrorMessage('Object ID was not reserved due to an error. Check the console for more information.');
                });
            } else if (resolve.status === 404) {
                createSharepointList(headers).then(async () => {
                    // Add item to sharepoint list
                    await fetch(url,
                        {
                            headers: headers,
                            method: 'POST',
                            body: JSON.stringify({
                                "__metadata":
                                {
                                    "type": "SP.Data.Object_x0020_ControlListItem"
                                },
                                "Object_x0020_ID": idToReserve, "Object_x0020_Type": objectType, "Title": `${objectType} ${idToReserve}`
                            })
                        }
                    ).then(async (resolve) => {
                        if (resolve.ok) {
                            window.showInformationMessage('ID reserved successfully');
                        } else {
                            window.showInformationMessage(`ID was not reserved. ${resolve.status}: ${resolve.statusText} `);
                        }
                    }, (reject) => {
                        console.error(reject);
                        window.showErrorMessage('Object ID was not reserved due to an error. Check the console for more information.');
                    });
                });
            }
        });
}

export async function clearExpiredReservedObjects(context: ExtensionContext) {
    try {
        const config_file = await getConfigurationFile();
        if (config_file.reserved_expiresIn === 0 || config_file.reserved_expiresIn === undefined) {
            return;
        }

        const headers = new Headers();
        headers.append('Authorization', `Bearer ${await context.secrets?.get('token')}`);
        headers.append('Accept', 'application/json');
        headers.append('Content-Type', 'application/json');
        const company_id = await context.secrets.get('companyId');
        const url = `https://api.businesscentral.dynamics.com/v2.0/${config_file.tenantId}/${config_file.environmentName}/api/${config_file.apiPublisher}/${config_file.apiGroup}/${config_file.apiVersion}/companies(${company_id})/${config_file.entitySetName_ReservedObjects}`;
        await fetch(url,
            { headers: headers }
        ).then(async (resolve) => {
            const data = await resolve?.json();
            const current_datetime = new Date();
            const reservedObjects = data.value ?? [];
            await Promise.all(reservedObjects.map(async (value: { odata_etag: string, objectID: number, objectType: string, systemCreatedAt: string }) => {
                const url_with_query = `https://api.businesscentral.dynamics.com/v2.0/${config_file.tenantId}/${config_file.environmentName}/api/${config_file.apiPublisher}/${config_file.apiGroup}/${config_file.apiVersion}/companies(${company_id})/${config_file.entitySetName_ReservedObjects}(objectType='${value.objectType}',objectID=${value.objectID})`;
                const system_created_at = new Date(value.systemCreatedAt);
                const duration = (current_datetime.getTime() - system_created_at.getTime()) / (1000 * 60 * 60);
                if (duration >= config_file.reserved_expiresIn) {
                    await fetch(url_with_query, { headers: headers, method: 'DELETE' }).then((resolve) => {
                        if (resolve.ok) {
                            console.log(`Object ${value.objectType} ${value.objectID} cleared successfully!`);
                        } else {
                            console.error(`${resolve.status} ${resolve.statusText}`);
                        }
                    }, (reject) => {
                        console.error(reject);
                    });
                }
            }));
        }, (reject) => {
            console.log(reject);
        });
    } catch (error) {
        console.error(error);
    }
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
                Description: "Object Control List created by AL Object Control VSCode Extension"
            });
        const url = `${config_file.sharepoint_baseUrl}/sites/${config_file.sharepoint_siteName}/_api/web/lists`;
        await fetch(url,
            { headers: headers, body: create_list_json, method: 'POST' }
        ).then((resolve) => {
            if (resolve.ok) {
                console.log(`Sharepoint list created with name ${config_file.sharepoint_listName}`);
                createSharepointField(headers, 'Object Type', 2, 'SP.FieldText');
                createSharepointField(headers, 'Object ID', 1, 'SP.FieldNumber');
                addSharepointFieldToView(headers, 'Created');
            } else { console.log(resolve.status + ' ' + resolve.statusText); console.log(resolve.text); }
        }, (reject) => {
            console.error(reject);
        });
    } catch (error) {
        throw error;
    }
}

async function createSharepointField(headers: Headers, fieldName: string, fieldType: number, fieldMetadata: string) {
    // Create Sharepoint list fields
    const config_file = await getConfigurationFile();
    const url = `${config_file.sharepoint_baseUrl}/sites/${config_file.sharepoint_siteName}/_api/web/lists/GetByTitle('${config_file.sharepoint_listName}')/fields`;
    const fields_json = JSON.stringify({
        __metadata: {
            type: `${fieldMetadata}`
        },
        Title: fieldName,
        FieldTypeKind: fieldType
    });
    await fetch(url, { headers: headers, method: 'POST', body: fields_json }).then(async (resolve) => {
        if (resolve.ok) {
            addSharepointFieldToView(headers, fieldName);
            console.log(`${fieldName} field created.`);
        }
    }, (reject) => {
        console.error(reject);
    });
}

async function addSharepointFieldToView(headers: Headers, fieldInternalName: string) {
    const config = await getConfigurationFile();
    const url = `${config.sharepoint_baseUrl}/sites/${config.sharepoint_siteName}/_api/web/lists/GetByTitle('${config.sharepoint_listName}')/DefaultView/ViewFields/AddViewField('${fieldInternalName}')`;
    const response = await fetch(url, {
        method: 'POST',
        headers
    });

    if (!response.ok) {
        throw new Error(await response.text());
    }

    console.log(`Field ${fieldInternalName} added to default view.`);
}