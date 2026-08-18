export const SYNC_FILE_NAME = "mycloset-sync.json";
export const SYNC_SCOPE = "Files.ReadWrite.AppFolder";

export function syncComparablePayload(payload) {
  return {
    format: payload.format,
    version: payload.version,
    data: payload.data,
    preferences: payload.preferences ?? {},
  };
}

export async function hashBackupPayload(payload) {
  const bytes = new TextEncoder().encode(JSON.stringify(syncComparablePayload(payload)));
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

export function decideSyncAction({ localHash, cloudHash = "", lastHash = "", cloudExists = true }) {
  if (!cloudExists) return "upload";
  if (localHash === cloudHash) return "synced";
  if (!lastHash) return "conflict";
  const localChanged = localHash !== lastHash;
  const cloudChanged = cloudHash !== lastHash;
  if (localChanged && !cloudChanged) return "upload";
  if (!localChanged && cloudChanged) return "download";
  return "conflict";
}

export function withSyncMetadata(payload, { deviceId, deviceName, updatedAt = new Date().toISOString(), revision = crypto.randomUUID() }) {
  return {
    ...payload,
    sync: { revision, updatedAt, deviceId, deviceName },
  };
}

export function createDeviceName(userAgent = "") {
  const ua = userAgent.toLowerCase();
  const device = /ipad/.test(ua) ? "iPad" : /iphone/.test(ua) ? "iPhone" : /android/.test(ua) ? "Android" : /mac/.test(ua) ? "Mac" : /windows/.test(ua) ? "Windows PC" : "웹 브라우저";
  const browser = /edg\//.test(ua) ? "Edge" : /chrome\//.test(ua) ? "Chrome" : /safari\//.test(ua) ? "Safari" : /firefox\//.test(ua) ? "Firefox" : "Browser";
  return `${device} · ${browser}`;
}
