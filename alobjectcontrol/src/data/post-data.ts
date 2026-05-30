import { ExtensionContext, window } from "vscode";
import { getConfigurationFile } from "./get-data";

export async function reserveId(context: ExtensionContext, idToReserve: number, objectType: string) {
    const config_file = await getConfigurationFile();
    const headers = new Headers();
    headers.append('Authorization', `Bearer ${await context.secrets?.get('tokenSp')}`);
    headers.append('Accept', 'application/json');
    headers.append('Content-Type', 'application/json');
    const url = `${config_file.sharepoint_baseUrl}/sites/${config_file.sharepoint_siteName}/_api/web/lists/getbytitle('${config_file.sharepoint_listName}')/items`;
    await fetch(url,
        {
            headers: headers,
            method: 'POST',
            body: JSON.stringify({ "ObjectID": idToReserve, "ObjectType": objectType })
        }
    ).then(async (resolve) => {
        if (resolve.ok) {
            window.showInformationMessage('ID reserved successfully');
        } else {
            const json_response = await resolve.json();
            window.showInformationMessage(`ID was not reserved. ${resolve.status}: ${json_response.error.message} `);
        }
    }, (reject) => {
        console.error(reject);
        window.showErrorMessage('Object ID was not reserved due to an error. Check the console for more information.');
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