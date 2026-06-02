import { Uri, workspace, env, window, ExtensionContext } from 'vscode';
import * as fs from 'fs';
import { AuthenticationResult, PublicClientApplication } from "@azure/msal-node";

var tenant_id: string;
var client_id: string;
var rangeFrom: number;
var rangeTo: number;
var sharepoint_baseUrl: string;
var sharepoint_siteName: string;
var sharepoint_listName: string;

var auth_token_sharepoint: AuthenticationResult | undefined;

export async function getLastRealObjNo(context: ExtensionContext, objectType: string) {

    const config_file = await getConfigurationFile();
    tenant_id = config_file.tenantId;
    client_id = config_file.clientId;
    rangeFrom = config_file.rangeFrom;
    rangeTo = config_file.rangeTo;
    sharepoint_baseUrl = config_file.sharepoint_baseUrl;
    sharepoint_siteName = config_file.sharepoint_siteName;
    sharepoint_listName = config_file.sharepoint_listName;

    const current_date_time = new Date();
    if (!auth_token_sharepoint?.accessToken || !auth_token_sharepoint.expiresOn || current_date_time >= auth_token_sharepoint.expiresOn) {
        const choice = await window.showInformationMessage('Do you wish to authenticate to get the next ID?', 'Yes');
        if (choice === 'Yes') {
            auth_token_sharepoint = await getAuthenticationToken_Sharepoint();
            await context.secrets.store('tokenSp', auth_token_sharepoint?.accessToken);
        }
    }
    if (auth_token_sharepoint === undefined) {
        window.showWarningMessage('Authentication is required to retrieve the next ID.');
        return rangeFrom;
    }
    return (
        getReservedObjsSet(objectType)
    ).then(async (values) => {
        const reserved_obj_set = values;
        for (let i = rangeFrom; i <= rangeTo; i++) {
            if (!reserved_obj_set.has(i)) {
                return i;
            }
        }
        throw new Error('No more available IDs');
    }, (reject) => {
        console.log(reject);
        return rangeFrom;
    });
}
async function getReservedObjsSet(objectType: string): Promise<Set<number>> {
    try {
        const headers = new Headers();
        headers.append("Authorization", `Bearer ${auth_token_sharepoint?.accessToken}`);
        headers.append("Accept", 'application/json');
        headers.append("Content-Type", 'application/json;odata=nometadata');

        // const reserved_objects_url = `${sharepoint_baseUrl}/sites/${sharepoint_siteName}/_api/web/lists/getbytitle('${sharepoint_listName}')/items?$filter=Object_x0020_Type eq '${objectType}'`; // https://m365b784709.sharepoint.com/sites/NW-B2000eBike/_api/web/lists/getbytitle('Object Control')/items?$filter=ObjectType eq '${objectType}'`
        const reserved_objects_url = `${sharepoint_baseUrl}/sites/${sharepoint_siteName}/_api/web/lists/getbytitle('${sharepoint_listName}')/items`;
        const reserved_objects = await fetch(reserved_objects_url, {
            headers: headers,
        });
        if (reserved_objects.ok) {
            const reserved_objects_data = await reserved_objects.json();
            if (reserved_objects_data.value?.length !== 0) {
                const reserved_objects_data_array = [...reserved_objects_data.value];
                const existingIds: Set<number> = new Set(reserved_objects_data_array.map((item) => { return item.ObjectID ?? item.ObjectID0 ?? item.Object_x0020_ID; }).sort());
                return (existingIds);
            } else {
                return (new Set<number>());
            }
        } else {
            console.log(`${reserved_objects.status}: ${reserved_objects.statusText}`);
            return (new Set<number>());
        }
    } catch (error) {
        throw error;
    }
}

export async function getConfigurationFile(): Promise<{ tenantId: string, clientId: string, rangeFrom: number, rangeTo: number, sharepoint_baseUrl: string, sharepoint_siteName: string; sharepoint_listName: string }> {
    const URIs: Uri[] = await workspace.findFiles('**/.object-control.json', null, 1);
    if (URIs.length > 0) {
        const file_content = fs.readFileSync(URIs[0].fsPath, 'utf8');
        const file_content_asJson = JSON?.parse(file_content);
        return ({
            tenantId: file_content_asJson.tenantId,
            clientId: file_content_asJson.clientId,
            rangeFrom: file_content_asJson.range.from,
            rangeTo: file_content_asJson.range.to,
            sharepoint_baseUrl: file_content_asJson.sharepointBaseUrl,
            sharepoint_siteName: file_content_asJson.sharepointSiteName,
            sharepoint_listName: file_content_asJson.sharepointListName
        });
    } else {
        window.showErrorMessage(".object-control.json doesn't exist in the current workspace.");
        throw new Error(".object-control.json doesn't exist in the current workspace.");
    }
}

async function getAuthenticationToken_Sharepoint(): Promise<AuthenticationResult> {
    const config_file = await getConfigurationFile();
    const pca = new PublicClientApplication({
        auth: {
            clientId: client_id,
            authority: `https://login.microsoftonline.com/${tenant_id}`
        }
    });

    return (await pca.acquireTokenInteractive({
        openBrowser: async (url: string) => {
            env.openExternal(Uri.parse(url));
        }, scopes: [`${config_file.sharepoint_baseUrl}/.default`]
    }));
}

export function clearCache(context: ExtensionContext) {
    auth_token_sharepoint = undefined;
    context.secrets.delete('companyId');
    context.secrets.delete('tokenSp');
}