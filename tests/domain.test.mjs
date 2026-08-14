import assert from "node:assert/strict";
import test from "node:test";
import {
  buildMonth, clampDetailWidth, createBackupPayload, createWardrobeResetState, filterItemsByCategory, formatFileTimestamp, getDetailWidthLimits, iso,
  matchesSearchQuery, moveItemByOffset, nextLookbookName, parseWardrobeExcelRows, removeItemReferences, reorderItemIds, searchLookbooks,
  searchWardrobeItems, searchWishlistItems, validItemIds, validateBackupPayload,
  wardrobeItemsToExcelRows, WARDROBE_EXCEL_HEADERS,
} from "../src/domain.mjs";

test("iso formats a local calendar date without timezone drift", () => {
  assert.equal(iso(new Date(2026, 7, 4)), "2026-08-04");
});

test("backup timestamps include a filesystem-safe local time", () => {
  assert.equal(formatFileTimestamp(new Date(2026, 7, 14, 9, 5, 7)), "2026-08-14_09-05-07");
});

test("buildMonth creates a Sunday-first six-week grid", () => {
  const days = buildMonth(new Date(2026, 7, 1));
  assert.equal(days.length, 42);
  assert.equal(days[0].key, "2026-07-26");
  assert.equal(days[41].key, "2026-09-05");
  assert.equal(days.filter((day) => day.current).length, 31);
  assert.equal(days[0].date.getDay(), 0);
});

test("detail rail limits preserve the main workspace at desktop sizes", () => {
  assert.deepEqual(getDetailWidthLimits(1976, true), { minWidth: 360, maxWidth: 1120 });
  assert.deepEqual(getDetailWidthLimits(901, false), { minWidth: 293, maxWidth: 293 });
  assert.equal(clampDetailWidth(2000, 1976, true), 1120);
  assert.equal(clampDetailWidth(100, 1976, true), 360);
});

test("category filtering preserves order and supports the all category", () => {
  const items = [{ id: "a", category: "상의" }, { id: "b", category: "하의" }, { id: "c", category: "상의" }];
  assert.equal(filterItemsByCategory(items, "전체"), items);
  assert.deepEqual(filterItemsByCategory(items, "상의").map((item) => item.id), ["a", "c"]);
});

test("wardrobe search matches normalized multi-word item fields", () => {
  const items = [
    { id: "a", name: "배럴 팬츠", category: "하의", brand: "UNIQLO", color: "01 OFF WHITE" },
    { id: "b", name: "옥스포드 셔츠", category: "상의", brand: "COS", color: "블루" },
  ];
  assert.deepEqual(searchWardrobeItems(items, "  uniqlo   하의  ").map((item) => item.id), ["a"]);
  assert.deepEqual(searchWardrobeItems(items, "off white").map((item) => item.id), ["a"]);
  assert.equal(matchesSearchQuery(["ＵＮＩＱＬＯ"], "uniqlo"), true);
});

test("wishlist search matches product details and status", () => {
  const items = [
    { id: "a", name: "후드 집업", category: "상의", brand: "ADIDAS", color: "블랙", status: "관심" },
    { id: "b", name: "몬트리올", category: "신발", brand: "ADIDAS", color: "카본", status: "구매 예정" },
  ];
  assert.deepEqual(searchWishlistItems(items, "후드 블랙").map((item) => item.id), ["a"]);
  assert.deepEqual(searchWishlistItems(items, "구매 예정").map((item) => item.id), ["b"]);
});

test("lookbook search includes its memo and contained wardrobe item fields", () => {
  const lookbooks = [
    { id: "office", name: "룩북 1", memo: "비 오는 날 출근", itemIds: ["pants"] },
    { id: "weekend", name: "주말 산책", memo: "", itemIds: ["shirt"] },
  ];
  const itemMap = {
    pants: { name: "배럴팬츠", brand: "UNIQLO", category: "하의", color: "OFF WHITE" },
    shirt: { name: "리넨 셔츠", brand: "COS", category: "상의", color: "베이지" },
  };
  assert.deepEqual(searchLookbooks(lookbooks, itemMap, "비 오는").map((entry) => entry.id), ["office"]);
  assert.deepEqual(searchLookbooks(lookbooks, itemMap, "uniqlo 하의").map((entry) => entry.id), ["office"]);
});

test("lookbook names use the first available default number", () => {
  assert.equal(nextLookbookName([]), "룩북 1");
  assert.equal(nextLookbookName([{ name: "룩북 1" }, { name: "여름 휴가" }, { name: "룩북 3" }]), "룩북 2");
});

test("outfit reordering is immutable and ignores invalid targets", () => {
  const ids = ["a", "b", "c"];
  assert.deepEqual(reorderItemIds(ids, "a", "c"), ["b", "c", "a"]);
  assert.deepEqual(ids, ["a", "b", "c"]);
  assert.equal(reorderItemIds(ids, "missing", "c"), ids);
  assert.equal(reorderItemIds(ids, "a", "a"), ids);
});

test("keyboard outfit movement clamps at both ends", () => {
  const ids = ["a", "b", "c"];
  assert.deepEqual(moveItemByOffset(ids, "b", -1), ["b", "a", "c"]);
  assert.deepEqual(moveItemByOffset(ids, "b", 1), ["a", "c", "b"]);
  assert.equal(moveItemByOffset(ids, "a", -1), ids);
  assert.equal(moveItemByOffset(ids, "missing", 1), ids);
});

test("deleting an item removes outfit and lookbook references", () => {
  const result = removeItemReferences(
    { "2026-08-04": ["a", "b"], "2026-08-05": ["b"] },
    [{ id: "one", itemIds: ["a", "b"] }, { id: "two", itemIds: ["b"] }],
    "b",
  );
  assert.deepEqual(result.outfits, { "2026-08-04": ["a"], "2026-08-05": [] });
  assert.deepEqual(result.lookbooks, [{ id: "one", itemIds: ["a"] }]);
});

test("lookbook application keeps only existing wardrobe items", () => {
  assert.deepEqual(validItemIds(["a", "missing", "b"], { a: {}, b: {} }), ["a", "b"]);
});

test("wardrobe reset clears every dependent collection and restores default categories", () => {
  const defaults = ["상의", "하의"];
  assert.deepEqual(createWardrobeResetState(defaults), {
    items: [],
    outfits: {},
    outfitNotes: {},
    lookbooks: [],
    categories: ["상의", "하의"],
  });
  assert.notEqual(createWardrobeResetState(defaults).categories, defaults);
});

test("wardrobe Excel rows preserve URL-first registration fields", () => {
  const rows = wardrobeItemsToExcelRows([{ name: "데님 팬츠", category: "하의", brand: "무신사 스탠다드", color: "다크 블루", url: "https://shop.example/item", image: "https://img.example/item.jpg" }]);
  assert.deepEqual(WARDROBE_EXCEL_HEADERS, ["이름*", "분류*", "브랜드", "색상", "상품 URL", "이미지 URL"]);
  assert.deepEqual(rows[0], ["데님 팬츠", "하의", "무신사 스탠다드", "다크 블루", "https://shop.example/item", "https://img.example/item.jpg"]);
  assert.deepEqual(parseWardrobeExcelRows([Object.fromEntries(WARDROBE_EXCEL_HEADERS.map((header, index) => [header, rows[0][index]]))], 123), [{
    id: "excel-123-0", name: "데님 팬츠", category: "하의", brand: "무신사 스탠다드", color: "다크 블루", url: "https://shop.example/item", image: "https://img.example/item.jpg",
  }]);
});

test("wardrobe Excel export omits embedded upload data and importer ignores example rows", () => {
  assert.equal(wardrobeItemsToExcelRows([{ name: "셔츠", category: "상의", image: "data:image/png;base64,abc" }])[0][5], "");
  assert.deepEqual(parseWardrobeExcelRows([{ "이름*": "예시) 셔츠", "분류*": "상의" }, { "이름*": "", "분류*": "하의" }], 123), []);
});

test("full backup payload round-trips every collection and screen preference", () => {
  const data = {
    items: [{ id: "shirt", image: "data:image/png;base64,abc" }],
    categories: ["상의"],
    outfits: { "2026-08-14": ["shirt"] },
    outfitNotes: { "2026-08-14": "회의" },
    lookbooks: [{ id: "daily", itemIds: ["shirt"] }],
    wishlist: [{ id: "wish" }],
    inspirations: [{ id: "pin", image: "data:image/png;base64,look", caption: "여름 코디" }],
  };
  const preferences = {
    theme: "dark", sidebarCollapsed: true, detailWidth: 520,
    wardrobeView: "table", wardrobeSort: "brand", lookbookSort: "items", wishlistSort: "name",
  };
  const payload = createBackupPayload(data, preferences, "2026-08-14T00:00:00.000Z");
  assert.equal(payload.format, "mycloset-backup");
  assert.equal(payload.version, 1);
  assert.deepEqual(validateBackupPayload(JSON.parse(JSON.stringify(payload))), { data, preferences });
});

test("older backups without inspiration pins remain compatible", () => {
  const payload = createBackupPayload({
    items: [], categories: ["상의"], outfits: {}, outfitNotes: {}, lookbooks: [], wishlist: [],
  });
  assert.equal("inspirations" in payload.data, false);
  assert.equal("inspirations" in validateBackupPayload(payload).data, false);
});

test("backup validation rejects unrelated and incomplete JSON files", () => {
  assert.throws(() => validateBackupPayload({}), /MyCloset/);
  assert.throws(() => validateBackupPayload({ format: "mycloset-backup", version: 1, data: { items: [] } }), /형식/);
  assert.throws(() => validateBackupPayload({
    format: "mycloset-backup", version: 1,
    data: { items: [null], categories: ["상의"], outfits: {}, outfitNotes: {}, lookbooks: [], wishlist: [] },
  }), /올바르지 않은 항목/);
});
