import assert from "node:assert/strict";
import test from "node:test";
import {
  filterItemsByCategory, moveItemByOffset, removeItemReferences, reorderItemIds, validItemIds,
} from "../src/domain.mjs";

test("wardrobe, outfit, and lookbook data stay consistent through a typical lifecycle", () => {
  const items = [
    { id: "shirt", category: "상의" },
    { id: "pants", category: "하의" },
    { id: "shoes", category: "신발" },
  ];
  const itemMap = Object.fromEntries(items.map((item) => [item.id, item]));

  assert.deepEqual(filterItemsByCategory(items, "상의").map((item) => item.id), ["shirt"]);

  const savedOutfit = ["shirt", "pants", "shoes"];
  const dragged = reorderItemIds(savedOutfit, "shoes", "shirt");
  assert.deepEqual(dragged, ["shoes", "shirt", "pants"]);
  assert.deepEqual(moveItemByOffset(dragged, "shirt", 1), ["shoes", "pants", "shirt"]);

  const lookbookIds = validItemIds(["shirt", "pants", "removed"], itemMap);
  assert.deepEqual(lookbookIds, ["shirt", "pants"]);

  const cascaded = removeItemReferences(
    { "2026-08-04": dragged },
    [{ id: "daily", itemIds: lookbookIds }],
    "pants",
  );
  assert.deepEqual(cascaded.outfits["2026-08-04"], ["shoes", "shirt"]);
  assert.deepEqual(cascaded.lookbooks[0].itemIds, ["shirt"]);
});
