import assert from "node:assert/strict";
import test from "node:test";
import { createDeviceName, decideSyncAction, hashBackupPayload, syncComparablePayload, withSyncMetadata } from "../src/sync-domain.mjs";

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

test("sync action resolves one-sided changes and detects conflicts", () => {
  assert.equal(decideSyncAction({ localHash: "a", cloudExists: false }), "upload");
  assert.equal(decideSyncAction({ localHash: "a", cloudHash: "a", lastHash: "old" }), "synced");
  assert.equal(decideSyncAction({ localHash: "local", cloudHash: "base", lastHash: "base" }), "upload");
  assert.equal(decideSyncAction({ localHash: "base", cloudHash: "cloud", lastHash: "base" }), "download");
  assert.equal(decideSyncAction({ localHash: "local", cloudHash: "cloud", lastHash: "base" }), "conflict");
  assert.equal(decideSyncAction({ localHash: "local", cloudHash: "cloud" }), "conflict");
});

test("cloud metadata records the originating device", () => {
  const result = withSyncMetadata(payload, { deviceId: "device-1", deviceName: "Mac · Chrome", updatedAt: "now", revision: "rev" });
  assert.deepEqual(result.sync, { deviceId: "device-1", deviceName: "Mac · Chrome", updatedAt: "now", revision: "rev" });
  assert.equal(createDeviceName("Mozilla Macintosh Chrome/120"), "Mac · Chrome");
});
