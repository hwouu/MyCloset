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
      const redirectResult = await client.handleRedirectPromise();
      if (redirectResult?.account) client.setActiveAccount(redirectResult.account);
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
  await client.loginRedirect({ scopes: [SYNC_SCOPE], prompt: "select_account" });
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
  if (options.allowMissingAppRoot && response.status === 403) return null;
  if (response.status === 412) throw new Error("다른 기기에서 먼저 저장했어요. 최신 데이터를 확인해주세요.");
  if (response.status === 429) throw new Error("OneDrive 요청이 잠시 많아요. 잠시 후 자동으로 다시 시도할게요.");
  if (!response.ok) {
    let detail = "";
    try { detail = (await response.json())?.error?.message || ""; } catch { /* no-op */ }
    throw new Error(detail || `OneDrive 요청에 실패했어요. (${response.status})`);
  }
  return response;
}

export async function getCloudSyncFile() {
  // A first-time app root can return 403 for a child lookup before the first
  // write creates the folder. Treat that first lookup like a missing file; an
  // actual permission problem is still surfaced by the following PUT request.
  const response = await graphFetch(`/me/drive/special/approot:/${SYNC_FILE_NAME}?$select=id,name,eTag,lastModifiedDateTime,size`, {
    allowNotFound: true,
    allowMissingAppRoot: true,
  });
  if (!response) return null;
  const metadata = await response.json();
  const contentResponse = await graphFetch(`/me/drive/items/${encodeURIComponent(metadata.id)}/content`);
  return { metadata, payload: await contentResponse.json() };
}

export async function uploadCloudSyncFile(payload, { ifMatch = "" } = {}) {
  const response = await graphFetch(`/me/drive/special/approot:/${SYNC_FILE_NAME}:/content`, {
    method: "PUT",
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      ...(ifMatch ? { "If-Match": ifMatch } : {}),
    },
    body: JSON.stringify(payload),
  });
  return response.json();
}
