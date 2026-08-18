import { SYNC_FILE_NAME, SYNC_SCOPE } from "./sync-domain.mjs";

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
let msalClientPromise;

function configuration() {
  const clientId = String(import.meta.env.VITE_MS_CLIENT_ID || "").trim();
  const tenantId = String(import.meta.env.VITE_MS_TENANT_ID || "common").trim();
  const redirectUri = String(import.meta.env.VITE_MS_REDIRECT_URI || window.location.origin).trim();
  return { clientId, tenantId, redirectUri };
}

export function isOneDriveConfigured() {
  return Boolean(configuration().clientId);
}

async function getMsalClient() {
  if (!isOneDriveConfigured()) throw new Error("Microsoft 연결 설정이 아직 완료되지 않았어요.");
  if (!msalClientPromise) {
    msalClientPromise = (async () => {
      const { BrowserCacheLocation, PublicClientApplication } = await import("@azure/msal-browser");
      const { clientId, tenantId, redirectUri } = configuration();
      const client = new PublicClientApplication({
        auth: {
          clientId,
          authority: `https://login.microsoftonline.com/${tenantId}`,
          redirectUri,
          postLogoutRedirectUri: redirectUri,
        },
        cache: { cacheLocation: BrowserCacheLocation.LocalStorage },
      });
      await client.initialize();
      await client.handleRedirectPromise();
      return client;
    })();
  }
  return msalClientPromise;
}

export async function getConnectedAccount() {
  if (!isOneDriveConfigured()) return null;
  const client = await getMsalClient();
  return client.getActiveAccount() || client.getAllAccounts()[0] || null;
}

export async function connectOneDrive() {
  const client = await getMsalClient();
  const result = await client.loginPopup({ scopes: [SYNC_SCOPE], prompt: "select_account" });
  client.setActiveAccount(result.account);
  return result.account;
}

export async function disconnectOneDrive() {
  const client = await getMsalClient();
  const account = client.getActiveAccount() || client.getAllAccounts()[0];
  if (account) await client.logoutPopup({ account, mainWindowRedirectUri: window.location.href });
}

async function accessToken() {
  const client = await getMsalClient();
  const account = client.getActiveAccount() || client.getAllAccounts()[0];
  if (!account) throw new Error("Microsoft 계정을 먼저 연결해주세요.");
  client.setActiveAccount(account);
  try {
    return (await client.acquireTokenSilent({ account, scopes: [SYNC_SCOPE] })).accessToken;
  } catch {
    return (await client.acquireTokenPopup({ account, scopes: [SYNC_SCOPE] })).accessToken;
  }
}

async function graphFetch(path, options = {}) {
  const token = await accessToken();
  const response = await fetch(`${GRAPH_ROOT}${path}`, {
    ...options,
    headers: { Authorization: `Bearer ${token}`, ...options.headers },
  });
  if (options.allowNotFound && response.status === 404) return null;
  if (!response.ok) {
    let detail = "";
    try { detail = (await response.json())?.error?.message || ""; } catch { /* no-op */ }
    throw new Error(detail || `OneDrive 요청에 실패했어요. (${response.status})`);
  }
  return response;
}

export async function getCloudSyncFile() {
  const response = await graphFetch(`/me/drive/special/approot:/${SYNC_FILE_NAME}?$select=id,name,eTag,lastModifiedDateTime,size,@microsoft.graph.downloadUrl`, { allowNotFound: true });
  if (!response) return null;
  const metadata = await response.json();
  const downloadUrl = metadata["@microsoft.graph.downloadUrl"];
  if (!downloadUrl) throw new Error("OneDrive 동기화 파일을 내려받을 수 없어요.");
  const contentResponse = await fetch(downloadUrl);
  if (!contentResponse.ok) throw new Error("OneDrive 동기화 파일을 읽지 못했어요.");
  return { metadata, payload: await contentResponse.json() };
}

export async function uploadCloudSyncFile(payload) {
  const folderResponse = await graphFetch("/me/drive/special/approot?$select=id");
  const folder = await folderResponse.json();
  const response = await graphFetch(`/me/drive/items/${encodeURIComponent(folder.id)}:/${SYNC_FILE_NAME}:/content`, {
    method: "PUT",
    headers: { "Content-Type": "application/json; charset=utf-8" },
    body: JSON.stringify(payload),
  });
  return response.json();
}
