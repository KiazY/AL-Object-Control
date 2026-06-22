import { Uri, workspace, env, window, ExtensionContext } from 'vscode';
import * as fs from 'fs';
import { AccountInfo, AuthenticationResult, PublicClientApplication } from "@azure/msal-node";

var auth_token_sharepoint: AuthenticationResult | undefined;
var msal_account: AccountInfo | undefined;

export async function ensureAuthenticated(context: ExtensionContext, prompt: string): Promise<boolean> {
    const current_date_time = new Date();
    if (!auth_token_sharepoint?.accessToken || !auth_token_sharepoint.expiresOn || current_date_time >= auth_token_sharepoint.expiresOn) {
        if (msal_account) {
            auth_token_sharepoint = await getAuthenticationToken_Sharepoint_Silent();
        }
        if (!auth_token_sharepoint) {
            const choice = await window.showInformationMessage(prompt, 'Yes');
            if (choice === 'Yes') {
                auth_token_sharepoint = await getAuthenticationToken_Sharepoint();
            }
        }
        if (auth_token_sharepoint) {
            await context.secrets.store('tokenSp', auth_token_sharepoint.accessToken);
        }
    }
    return auth_token_sharepoint !== undefined;
}

export async function getLastRealObjNo(context: ExtensionContext, objectType: string) {

    const config_file = await getConfigurationFile();

    if (!(await ensureAuthenticated(context, 'Do you wish to authenticate to get the next ID?'))) {
        window.showWarningMessage('Authentication is required to retrieve the next ID.');
        return config_file.rangeFrom;
    }
    return (
        getReservedObjsSet(objectType)
    ).then(async (values) => {
        const reserved_obj_set = values;
        for (let i = config_file.rangeFrom; i <= config_file.rangeTo; i++) {
            if (!reserved_obj_set.has(i)) {
                return i;
            }
        }
        throw new Error('No more available IDs');
    }, (reject) => {
        console.log(reject);
        return config_file.rangeFrom;
    });
}
async function getReservedObjsSet(objectType: string): Promise<Set<number>> {
    try {
        const config_file = await getConfigurationFile();
        const headers = new Headers();
        headers.append("Authorization", `Bearer ${auth_token_sharepoint?.accessToken}`);
        headers.append("Accept", 'application/json');
        headers.append("Content-Type", 'application/json;odata=nometadata');

        const reserved_objects_url = `${config_file.sharepoint_siteUrl}/_api/web/lists/getbytitle('${config_file.sharepoint_listName}')/items?$filter=ObjectType eq '${objectType}'`;
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

export async function getConfigurationFile(): Promise<{ tenantId: string, clientId: string, rangeFrom: number, rangeTo: number, sharepoint_siteUrl: string, sharepoint_listName: string }> {
    const URIs: Uri[] = await workspace.findFiles('**/.object-control.json', null, 1);
    if (URIs.length > 0) {
        const file_content = fs.readFileSync(URIs[0].fsPath, 'utf8');
        const file_content_asJson = JSON?.parse(file_content);
        return ({
            tenantId: file_content_asJson.tenantId,
            clientId: file_content_asJson.clientId,
            rangeFrom: file_content_asJson.range.from,
            rangeTo: file_content_asJson.range.to,
            sharepoint_siteUrl: file_content_asJson.sharepointSiteUrl?.replace(/\/$/, ''),
            sharepoint_listName: file_content_asJson.sharepointListName
        });
    } else {
        window.showErrorMessage(".object-control.json doesn't exist in the current workspace.");
        throw new Error(".object-control.json doesn't exist in the current workspace.");
    }
}

var msal_pca: PublicClientApplication | undefined;

async function getMsalApp(): Promise<PublicClientApplication> {
    if (!msal_pca) {
        const config_file = await getConfigurationFile();
        msal_pca = new PublicClientApplication({
            auth: {
                clientId: config_file.clientId,
                authority: `https://login.microsoftonline.com/${config_file.tenantId}`
            }
        });
    }
    return msal_pca;
}

async function getAuthenticationToken_Sharepoint(): Promise<AuthenticationResult> {
    const config_file = await getConfigurationFile();
    const pca = await getMsalApp();

    const result = await pca.acquireTokenInteractive({
        openBrowser: async (url: string) => {
            env.openExternal(Uri.parse(url));
        }, scopes: [`${new URL(config_file.sharepoint_siteUrl).origin}/.default`]
    });
    msal_account = result.account ?? undefined;
    return result;
}

async function getAuthenticationToken_Sharepoint_Silent(): Promise<AuthenticationResult | undefined> {
    if (!msal_account) {
        return undefined;
    }
    try {
        const config_file = await getConfigurationFile();
        const pca = await getMsalApp();
        const result = await pca.acquireTokenSilent({
            account: msal_account,
            scopes: [`${new URL(config_file.sharepoint_siteUrl).origin}/.default`]
        });
        msal_account = result.account ?? msal_account;
        return result;
    } catch (error) {
        console.log('Silent re-authentication failed, falling back to interactive.', error);
        return undefined;
    }
}

export function isAuthenticated(): boolean {
    const now = new Date();
    return !!(auth_token_sharepoint?.accessToken) &&
        (!auth_token_sharepoint.expiresOn || now < auth_token_sharepoint.expiresOn);
}

export async function isObjectReserved(objectType: string, objectId: number): Promise<boolean> {
    if (!auth_token_sharepoint?.accessToken) {
        return false;
    }

    try {
        const config = await getConfigurationFile();
        const headers = new Headers();
        headers.append('Authorization', `Bearer ${auth_token_sharepoint.accessToken}`);
        headers.append('Accept', 'application/json;odata=nometadata');

        const objectName = `${objectType} ${objectId}`;
        const escapedName = objectName.replace(/'/g, "''");
        const listUrl = `${config.sharepoint_siteUrl}/_api/web/lists/getbytitle('${config.sharepoint_listName}')/items`;
        const queryUrl = `${listUrl}?$select=Id,ObjectName&$filter=ObjectName eq '${escapedName}'`;

        const response = await fetch(queryUrl, { headers });
        if (!response.ok) {
            return false;
        }

        const data = await response.json();
        const items = data.value ?? data.d?.results ?? [];
        return items.length > 0;
    } catch {
        return false;
    }
}

export function clearCache(context: ExtensionContext) {
    auth_token_sharepoint = undefined;
    msal_account = undefined;
    msal_pca = undefined;
    context.secrets.delete('tokenSp');
}

export async function getAppJsonFile(): Promise<{ appName: string }> {
    const URIs: Uri[] = await workspace.findFiles('**/app.json', null, 1);
    if (URIs.length > 0) {
        const file_content = fs.readFileSync(URIs[0].fsPath, 'utf8');
        const file_content_asJson = JSON?.parse(file_content);
        return ({
            appName: file_content_asJson.name
        });
    } else {
        window.showErrorMessage(".object-control.json doesn't exist in the current workspace.");
        throw new Error(".object-control.json doesn't exist in the current workspace.");
    }
}