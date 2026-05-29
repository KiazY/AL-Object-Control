import { Uri, workspace, env, window, ExtensionContext } from 'vscode';
import * as fs from 'fs';
import { AuthenticationResult, PublicClientApplication } from "@azure/msal-node";

var tenant_id: string;
var client_id: string;
var environment_name: string;
var api_publisher: string;
var api_group: string;
var api_version: string;
var entity_setname_all_objs: string;
var entity_setname_reserved_objs: string;
var rangeFrom: number;
var rangeTo: number;

var auth_token: AuthenticationResult | undefined;

export async function getLastRealObjNo(context: ExtensionContext, objectType: string) {

    const config_file = await getConfigurationFile();
    tenant_id = config_file.tenantId;
    environment_name = config_file.environmentName;
    client_id = config_file.clientId;
    api_publisher = config_file.apiPublisher;
    api_group = config_file.apiGroup;
    api_version = config_file.apiVersion;
    entity_setname_all_objs = config_file.entitySetName_AllObjects;
    entity_setname_reserved_objs = config_file.entitySetName_ReservedObjects;
    rangeFrom = config_file.rangeFrom;
    rangeTo = config_file.rangeTo;

    const current_date_time = new Date();
    if (!auth_token?.accessToken || !auth_token.expiresOn || current_date_time >= auth_token.expiresOn) {
        const choice = await window.showInformationMessage('Do you wish to authenticate to get the next ID?', 'Yes');
        if (choice === 'Yes') {
            auth_token = await getAuthenticationToken();
            await context.secrets.store('token', auth_token?.accessToken);
        }
    }
    if (auth_token === undefined) {
        window.showWarningMessage('Authentication is required to retrieve the next ID.');
        return rangeFrom - 1;
    }
    const headers = new Headers();
    headers.append("Authorization", `Bearer ${auth_token?.accessToken}`);


    const company_id = await getCompanyId(headers);

    await context.secrets.store('companyId', company_id.toString());
    return (
        Promise.all(
            [
                getAllObjsSet(headers, company_id, objectType),
                getReservedObjsSet(headers, company_id, objectType)
            ]
        ).then(async (values) => {
            const last_obj_set = await values[0];
            const last_reserved_obj_set = await values[1];
            const merged_objs_set = new Set([...last_obj_set, ...last_reserved_obj_set]);
            for (let i = rangeFrom; i <= rangeTo; i++) {
                if (!merged_objs_set.has(i)) {
                    return i;
                }
            }
            throw new Error('No more available IDs');
        }, (reject) => {
            console.log(reject);
            return rangeFrom;
        })
    );
}

async function getCompanyId(headers: Headers) {
    const companies_url: string = `https://api.businesscentral.dynamics.com/v2.0/${tenant_id}/${environment_name}/api/v2.0/companies?$top=1`;
    const companies_response = await fetch(companies_url, {
        headers: headers
    });
    if (!companies_response.ok) {
        throw new Error(`Unable to retrieve companies. ${companies_response.status} ${companies_response.statusText}`);
    }
    const companies_data = await companies_response.json();
    if (!companies_data?.value || companies_data.value.length === 0) {
        throw new Error('No companies found in the current Business Central environment.');
    }
    const company_id: number = companies_data.value[0].id;
    return company_id;
}

async function getAllObjsSet(headers: Headers, company_id: number, objectType: string) {
    try {
        const all_obj_url = `https://api.businesscentral.dynamics.com/v2.0/${tenant_id}/${environment_name}/api/${api_publisher}/${api_group}/${api_version}/companies(${company_id})/${entity_setname_all_objs}?$filter=objectType eq '${objectType}' and objectID ge ${rangeFrom} and objectID lt ${rangeTo}&$orderby=objectID asc`;
        const all_objs_response = await fetch(all_obj_url, {
            headers: headers
        });
        const all_objs_data = await all_objs_response.json();
        if (all_objs_data.value?.length !== 0) {
            const all_objs_data_array: Array<{ odata_etag: string, objectType: string, objectID: number }> = [...all_objs_data.value].sort();
            const existingIds = new Set(all_objs_data_array.map((item) => { return item.objectID; }).sort());
            return (existingIds);
        }
        else {
            return (new Set<number>());
        }
    } catch (error) {
        throw error;
    }
}

async function getReservedObjsSet(headers: Headers, company_id: number, objectType: string) {
    try {
        const reserved_objects_url: string = `https://api.businesscentral.dynamics.com/v2.0/${tenant_id}/${environment_name}/api/${api_publisher}/${api_group}/${api_version}/companies(${company_id})/${entity_setname_reserved_objs}?$filter=objectType eq '${objectType}' and objectID ge ${rangeFrom} and objectID lt ${rangeTo}&$orderby=objectID asc`;
        const reserved_objects = await fetch(reserved_objects_url, {
            headers: headers
        });
        const reserved_objects_data = await reserved_objects.json();
        if (reserved_objects_data.value?.length !== 0) {
            const reserved_objects_data_array: Array<{ odata_etag: string, objectType: string, objectID: number, systemCreatedAt: string }> = [...reserved_objects_data.value];
            const existingIds = new Set(reserved_objects_data_array.map((item) => { return item.objectID; }).sort());
            return (existingIds);
        } else {
            return (new Set<number>());
        }
    } catch (error) {
        throw error;
    }
}

export async function getConfigurationFile(): Promise<{ environmentName: string, tenantId: string, clientId: string, apiPublisher: string, apiGroup: string, apiVersion: string, entitySetName_AllObjects: string, entitySetName_ReservedObjects: string, rangeFrom: number, rangeTo: number }> {
    const URIs: Uri[] = await workspace.findFiles('**/.object-control.json', null, 1);
    if (URIs.length > 0) {
        const file_content = fs.readFileSync(URIs[0].fsPath, 'utf8');
        const file_content_asJson = JSON?.parse(file_content);
        return ({
            environmentName: file_content_asJson.environmentName,
            tenantId: file_content_asJson.tenantId,
            clientId: file_content_asJson.clientId,
            apiPublisher: file_content_asJson.APIPublisher,
            apiGroup: file_content_asJson.APIGroup,
            apiVersion: file_content_asJson.APIVersion,
            entitySetName_AllObjects: file_content_asJson.EntitySetName_AllObjects,
            entitySetName_ReservedObjects: file_content_asJson.EntitySetName_ReservedObjects,
            rangeFrom: file_content_asJson.range.from,
            rangeTo: file_content_asJson.range.to
        });
    } else {
        window.showErrorMessage(".object-control.json doesn't exist in the current workspace.");
        throw new Error(".object-control.json doesn't exist in the current workspace.");
    }
}

async function getAuthenticationToken(): Promise<AuthenticationResult> {
    const pca = new PublicClientApplication({
        auth: {
            clientId: client_id,
            authority: `https://login.microsoftonline.com/${tenant_id}`
        }
    });

    return (await pca.acquireTokenInteractive({
        openBrowser: async (url: string) => {
            env.openExternal(Uri.parse(url));
        }, scopes: ['https://api.businesscentral.dynamics.com/.default']
    }));

}

export function clearCache(context: ExtensionContext) {
    auth_token = undefined;
    context.secrets.delete('companyId');
    context.secrets.delete('token');
}