import assert from "node:assert/strict";
import test from "node:test";
import {
  createDeviceName, hashBackupPayload, mergeSyncPayloads, prepareSyncPayload, syncComparablePayload,
} from "../src/sync-domain.mjs";

const payload = { format: "mycloset-backup", version: 1, exportedAt: "old", data: { items: [] }, preferences: { theme: "dark" } };

test("sync comparison ignores export and cloud metadata", async () => {
  const a = { ...payload, exportedAt: "a", sync: { revision: "one" } };
  const b = { ...payload, exportedAt: "b", sync: { revision: "two" } };
  assert.deepEqual(syncComparablePayload(a), syncComparablePayload(b));
  assert.equal(await hashBackupPayload(a), await hashBackupPayload(b));
});

test("sync comparison ignores device-specific screen preferences", () => {
  const a = { ...payload, preferences: { theme: "dark", currentView: "closet" } };
  const b = { ...payload, preferences: { theme: "light", currentView: "calendar" } };
  assert.deepEqual(syncComparablePayload(a), syncComparablePayload(b));
});

test("cloud metadata records the originating device", async () => {
  const result = await prepareSyncPayload(payload, { deviceId: "device-1", deviceName: "Mac · Chrome", now: "2026-01-01T00:00:00.000Z", revision: "rev" });
  assert.equal(result.sync.deviceId, "device-1");
  assert.equal(result.sync.deviceName, "Mac · Chrome");
  assert.equal(result.sync.revision, "rev");
  assert.equal(createDeviceName("Mozilla Macintosh Chrome/120"), "Mac · Chrome");
});

const fullPayload = (data = {}) => ({
  format: "mycloset-backup",
  version: 1,
  exportedAt: "2026-01-01T00:00:00.000Z",
  data: {
    items: [], categories: ["상의"], outfits: {}, outfitNotes: {}, lookbooks: [], wishlist: [], inspirations: [],
    ...data,
  },
  preferences: {},
});

test("sync merge keeps wardrobe additions made on different devices", async () => {
  const base = await prepareSyncPayload(fullPayload(), { deviceId: "base", now: "2026-01-01T00:00:00.000Z", revision: "base" });
  const local = await prepareSyncPayload(fullPayload({ items: [{ id: "local", name: "로컬 옷", category: "상의" }] }), {
    previousManifest: base.sync.entities, deviceId: "pc", now: "2026-01-02T00:00:00.000Z", revision: "pc-1",
  });
  const cloud = await prepareSyncPayload(fullPayload({ items: [{ id: "cloud", name: "클라우드 옷", category: "상의" }] }), {
    previousManifest: base.sync.entities, deviceId: "tablet", now: "2026-01-02T00:00:01.000Z", revision: "tablet-1",
  });
  const merged = await mergeSyncPayloads(local, cloud, { deviceId: "pc", now: "2026-01-03T00:00:00.000Z", revision: "merge" });
  assert.deepEqual(merged.payload.data.items.map((item) => item.id).sort(), ["cloud", "local"]);
  assert.equal(merged.localChanged, true);
  assert.equal(merged.cloudChanged, true);
});

test("wardrobe deletion removes an unchanged item but preserves a concurrently edited item", async () => {
  const original = { id: "coat", name: "코트", category: "상의", color: "검정" };
  const basePayload = fullPayload({ items: [original] });
  const base = await prepareSyncPayload(basePayload, { deviceId: "base", now: "2026-01-01T00:00:00.000Z", revision: "base" });
  const deleted = await prepareSyncPayload(fullPayload(), {
    previousManifest: base.sync.entities, deviceId: "pc", now: "2026-01-02T00:00:00.000Z", revision: "delete",
  });
  const unchangedMerge = await mergeSyncPayloads(deleted, base, { deviceId: "pc", now: "2026-01-03T00:00:00.000Z", revision: "merge-1" });
  assert.equal(unchangedMerge.payload.data.items.length, 0);

  const edited = await prepareSyncPayload(fullPayload({ items: [{ ...original, color: "네이비" }] }), {
    previousManifest: base.sync.entities, deviceId: "tablet", now: "2026-01-02T00:00:01.000Z", revision: "edit",
  });
  const editedMerge = await mergeSyncPayloads(deleted, edited, { deviceId: "pc", now: "2026-01-03T00:00:00.000Z", revision: "merge-2" });
  assert.equal(editedMerge.payload.data.items[0].color, "네이비");
});

test("outfit merge applies the latest change per date", async () => {
  const item = { id: "shirt", name: "셔츠", category: "상의" };
  const basePayload = fullPayload({ items: [item], outfits: { "2026-01-10": ["shirt"] } });
  const base = await prepareSyncPayload(basePayload, { deviceId: "base", now: "2026-01-01T00:00:00.000Z", revision: "base" });
  const local = await prepareSyncPayload(fullPayload({ items: [item], outfits: { "2026-01-10": ["shirt"] }, outfitNotes: { "2026-01-10": "PC 메모" } }), {
    previousManifest: base.sync.entities, deviceId: "pc", now: "2026-01-02T00:00:00.000Z", revision: "pc",
  });
  const cloud = await prepareSyncPayload(fullPayload({ items: [item], outfits: { "2026-01-10": ["shirt"] }, outfitNotes: { "2026-01-10": "태블릿 최신 메모" } }), {
    previousManifest: base.sync.entities, deviceId: "tablet", now: "2026-01-02T00:00:02.000Z", revision: "tablet",
  });
  const merged = await mergeSyncPayloads(local, cloud, { deviceId: "pc", now: "2026-01-03T00:00:00.000Z", revision: "merge" });
  assert.equal(merged.payload.data.outfitNotes["2026-01-10"], "태블릿 최신 메모");
});

test("outfit merge preserves changes made on different dates", async () => {
  const item = { id: "shirt", name: "셔츠", category: "상의" };
  const base = await prepareSyncPayload(fullPayload({ items: [item] }), {
    deviceId: "base", now: "2026-01-01T00:00:00.000Z", revision: "base",
  });
  const local = await prepareSyncPayload(fullPayload({ items: [item], outfits: { "2026-01-10": ["shirt"] } }), {
    previousManifest: base.sync.entities, deviceId: "pc", now: "2026-01-02T00:00:00.000Z", revision: "pc",
  });
  const cloud = await prepareSyncPayload(fullPayload({ items: [item], outfits: { "2026-01-11": ["shirt"] } }), {
    previousManifest: base.sync.entities, deviceId: "tablet", now: "2026-01-02T00:00:01.000Z", revision: "tablet",
  });
  const merged = await mergeSyncPayloads(local, cloud, { deviceId: "pc", now: "2026-01-03T00:00:00.000Z", revision: "merge" });
  assert.deepEqual(Object.keys(merged.payload.data.outfits).sort(), ["2026-01-10", "2026-01-11"]);
});

test("outfit deletion wins when it is the latest change for that date", async () => {
  const item = { id: "shirt", name: "셔츠", category: "상의" };
  const basePayload = fullPayload({ items: [item], outfits: { "2026-01-10": ["shirt"] }, outfitNotes: { "2026-01-10": "기존 메모" } });
  const base = await prepareSyncPayload(basePayload, {
    deviceId: "base", now: "2026-01-01T00:00:00.000Z", revision: "base",
  });
  const deleted = await prepareSyncPayload(fullPayload({ items: [item] }), {
    previousManifest: base.sync.entities, deviceId: "pc", now: "2026-01-03T00:00:00.000Z", revision: "delete",
  });
  const staleCloud = await prepareSyncPayload(basePayload, {
    previousManifest: base.sync.entities, deviceId: "tablet", now: "2026-01-02T00:00:00.000Z", revision: "stale",
  });
  const merged = await mergeSyncPayloads(deleted, staleCloud, { deviceId: "pc", now: "2026-01-04T00:00:00.000Z", revision: "merge" });
  assert.equal(merged.payload.data.outfits["2026-01-10"], undefined);
  assert.equal(merged.payload.data.outfitNotes["2026-01-10"], undefined);
});
