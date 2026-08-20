export const SYNC_FILE_NAME = "mycloset-sync.json";
export const SYNC_SCOPE = "Files.ReadWrite.AppFolder";
export const SYNC_SCHEMA_VERSION = 2;

const ENTITY_GROUPS = ["items", "categories", "categoryOrder", "outfits", "lookbooks", "wishlist", "inspirations"];
const PRESERVE_ON_DELETE_GROUPS = new Set(["items", "categories", "lookbooks", "wishlist", "inspirations"]);

function stableValue(value) {
  if (Array.isArray(value)) return value.map(stableValue);
  if (!value || typeof value !== "object") return value;
  return Object.fromEntries(Object.keys(value).sort().map((key) => [key, stableValue(value[key])]));
}

async function hashValue(value) {
  const bytes = new TextEncoder().encode(JSON.stringify(stableValue(value)));
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

function entityRecords(payload) {
  const data = payload?.data ?? {};
  const byId = (records) => Object.fromEntries((Array.isArray(records) ? records : []).filter((record) => record?.id).map((record) => [record.id, record]));
  const categories = Array.isArray(data.categories) ? data.categories : [];
  const outfitDates = new Set([...Object.keys(data.outfits ?? {}), ...Object.keys(data.outfitNotes ?? {})]);
  const outfits = {};
  outfitDates.forEach((date) => {
    const itemIds = Array.isArray(data.outfits?.[date]) ? data.outfits[date] : [];
    const note = typeof data.outfitNotes?.[date] === "string" ? data.outfitNotes[date] : "";
    if (itemIds.length || note) outfits[date] = { itemIds, note };
  });
  return {
    items: byId(data.items),
    categories: Object.fromEntries(categories.map((name) => [name, name])),
    categoryOrder: { order: categories },
    outfits,
    lookbooks: byId(data.lookbooks),
    wishlist: byId(data.wishlist),
    inspirations: byId(data.inspirations),
  };
}

function validManifest(manifest) {
  return manifest && typeof manifest === "object" && !Array.isArray(manifest) ? manifest : {};
}

function timestampValue(value) {
  const parsed = Date.parse(value || "");
  return Number.isFinite(parsed) ? parsed : 0;
}

function newerMeta(left = {}, right = {}) {
  const timeDifference = timestampValue(left.updatedAt) - timestampValue(right.updatedAt);
  if (timeDifference) return timeDifference > 0 ? left : right;
  const revisionDifference = String(left.revision || "").localeCompare(String(right.revision || ""));
  if (revisionDifference) return revisionDifference > 0 ? left : right;
  return String(left.deviceId || "").localeCompare(String(right.deviceId || "")) >= 0 ? left : right;
}

function isDeleted(meta) {
  return Boolean(meta?.deletedAt);
}

function chooseEntity(group, localRecord, localMeta, cloudRecord, cloudMeta) {
  if (!localMeta) return { record: cloudRecord, meta: cloudMeta };
  if (!cloudMeta) return { record: localRecord, meta: localMeta };
  if (isDeleted(localMeta) !== isDeleted(cloudMeta) && PRESERVE_ON_DELETE_GROUPS.has(group)) {
    const deleted = isDeleted(localMeta) ? localMeta : cloudMeta;
    const live = isDeleted(localMeta) ? cloudMeta : localMeta;
    const liveRecord = isDeleted(localMeta) ? cloudRecord : localRecord;
    const deletionStillTargetsSameVersion = Boolean(deleted.previousHash) && deleted.previousHash === live.hash;
    if (!deletionStillTargetsSameVersion || timestampValue(live.updatedAt) > timestampValue(deleted.deletedAt)) {
      return { record: liveRecord, meta: live };
    }
    return { record: undefined, meta: deleted };
  }
  const winner = newerMeta(localMeta, cloudMeta);
  if (winner === localMeta) return { record: isDeleted(localMeta) ? undefined : localRecord, meta: localMeta };
  return { record: isDeleted(cloudMeta) ? undefined : cloudRecord, meta: cloudMeta };
}

function rebuildData(records) {
  const activeItems = Object.values(records.items);
  const itemIds = new Set(activeItems.map((item) => item.id));
  const activeCategories = new Set(Object.keys(records.categories));
  activeItems.forEach((item) => { if (item.category) activeCategories.add(item.category); });
  const preferredOrder = Array.isArray(records.categoryOrder.order) ? records.categoryOrder.order : [];
  const categories = [...preferredOrder.filter((name) => activeCategories.has(name)), ...[...activeCategories].filter((name) => !preferredOrder.includes(name))];
  const outfits = {};
  const outfitNotes = {};
  Object.entries(records.outfits).forEach(([date, outfit]) => {
    const ids = (outfit.itemIds ?? []).filter((id) => itemIds.has(id));
    if (ids.length) outfits[date] = ids;
    if (outfit.note) outfitNotes[date] = outfit.note;
  });
  return {
    items: activeItems,
    categories,
    outfits,
    outfitNotes,
    lookbooks: Object.values(records.lookbooks).map((lookbook) => ({ ...lookbook, itemIds: lookbook.itemIds.filter((id) => itemIds.has(id)) })),
    wishlist: Object.values(records.wishlist),
    inspirations: Object.values(records.inspirations),
  };
}

export function syncComparablePayload(payload) {
  return {
    format: payload.format,
    version: payload.version,
    data: stableValue(payload.data),
  };
}

export async function hashBackupPayload(payload) {
  return hashValue(syncComparablePayload(payload));
}

async function hashSyncDocument(payload) {
  return hashValue({
    ...syncComparablePayload(payload),
    entities: payload?.sync?.entities ?? {},
  });
}

export async function prepareSyncPayload(payload, {
  previousManifest = payload?.sync?.entities,
  deviceId = "unknown-device",
  deviceName = "웹 브라우저",
  now = new Date().toISOString(),
  revision = crypto.randomUUID(),
} = {}) {
  const records = entityRecords(payload);
  const previous = validManifest(previousManifest);
  const entities = {};
  for (const group of ENTITY_GROUPS) {
    const groupRecords = records[group];
    const previousGroup = validManifest(previous[group]);
    const nextGroup = {};
    for (const [id, record] of Object.entries(groupRecords)) {
      const hash = await hashValue(record);
      const old = previousGroup[id];
      nextGroup[id] = old?.hash === hash && !old.deletedAt
        ? old
        : { hash, updatedAt: now, deviceId, revision };
    }
    for (const [id, old] of Object.entries(previousGroup)) {
      if (id in groupRecords) continue;
      nextGroup[id] = old.deletedAt
        ? old
        : { hash: "", previousHash: old.hash || old.previousHash || "", updatedAt: now, deletedAt: now, deviceId, revision };
    }
    entities[group] = nextGroup;
  }
  return {
    ...payload,
    sync: {
      schemaVersion: SYNC_SCHEMA_VERSION,
      revision,
      updatedAt: now,
      deviceId,
      deviceName,
      entities,
    },
  };
}

export async function normalizeSyncPayload(payload, { deviceId = "legacy-device", deviceName = "이전 버전", now } = {}) {
  if (payload?.sync?.schemaVersion === SYNC_SCHEMA_VERSION && payload.sync.entities) return payload;
  return prepareSyncPayload(payload, {
    previousManifest: {},
    deviceId: payload?.sync?.deviceId || deviceId,
    deviceName: payload?.sync?.deviceName || deviceName,
    now: payload?.sync?.updatedAt || payload?.exportedAt || now || new Date(0).toISOString(),
    revision: payload?.sync?.revision || "legacy",
  });
}

export async function mergeSyncPayloads(localPayload, cloudPayload, {
  deviceId = "unknown-device",
  deviceName = "웹 브라우저",
  now = new Date().toISOString(),
  revision = crypto.randomUUID(),
} = {}) {
  const local = await normalizeSyncPayload(localPayload, { deviceId, deviceName, now });
  const cloud = await normalizeSyncPayload(cloudPayload);
  const localRecords = entityRecords(local);
  const cloudRecords = entityRecords(cloud);
  const records = {};
  const entities = {};
  for (const group of ENTITY_GROUPS) {
    records[group] = {};
    entities[group] = {};
    const ids = new Set([
      ...Object.keys(localRecords[group]),
      ...Object.keys(cloudRecords[group]),
      ...Object.keys(local.sync.entities[group] ?? {}),
      ...Object.keys(cloud.sync.entities[group] ?? {}),
    ]);
    ids.forEach((id) => {
      const chosen = chooseEntity(
        group,
        localRecords[group][id],
        local.sync.entities[group]?.[id],
        cloudRecords[group][id],
        cloud.sync.entities[group]?.[id],
      );
      if (chosen.record !== undefined) records[group][id] = chosen.record;
      if (chosen.meta) entities[group][id] = chosen.meta;
    });
  }
  const data = rebuildData(records);
  const merged = {
    ...local,
    exportedAt: now,
    data,
    sync: { schemaVersion: SYNC_SCHEMA_VERSION, revision, updatedAt: now, deviceId, deviceName, entities },
  };
  return {
    payload: merged,
    localChanged: await hashBackupPayload(merged) !== await hashBackupPayload(local),
    cloudChanged: await hashSyncDocument(merged) !== await hashSyncDocument(cloud),
  };
}

export function createDeviceName(userAgent = "") {
  const ua = userAgent.toLowerCase();
  const device = /ipad/.test(ua) ? "iPad" : /iphone/.test(ua) ? "iPhone" : /android/.test(ua) ? "Android" : /mac/.test(ua) ? "Mac" : /windows/.test(ua) ? "Windows PC" : "웹 브라우저";
  const browser = /edg\//.test(ua) ? "Edge" : /chrome\//.test(ua) ? "Chrome" : /safari\//.test(ua) ? "Safari" : /firefox\//.test(ua) ? "Firefox" : "Browser";
  return `${device} · ${browser}`;
}
