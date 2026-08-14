export const DETAIL_RAIL_MIN = 360;
export const DETAIL_RAIL_MAX = 1120;
export const DETAIL_RAIL_MAX_RATIO = 0.6;
export const MAIN_PANEL_MIN = 420;
export const BACKUP_FORMAT = "mycloset-backup";
export const BACKUP_VERSION = 1;
export const WARDROBE_EXCEL_HEADERS = ["이름*", "분류*", "브랜드", "색상", "상품 URL", "이미지 URL"];

export function iso(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

export function formatFileTimestamp(date = new Date()) {
  return `${iso(date)}_${String(date.getHours()).padStart(2, "0")}-${String(date.getMinutes()).padStart(2, "0")}-${String(date.getSeconds()).padStart(2, "0")}`;
}

export function buildMonth(cursor) {
  const start = new Date(cursor.getFullYear(), cursor.getMonth(), 1);
  const first = new Date(cursor.getFullYear(), cursor.getMonth(), 1 - start.getDay());
  return Array.from({ length: 42 }, (_, index) => {
    const date = new Date(first);
    date.setDate(first.getDate() + index);
    return { date, key: iso(date), current: date.getMonth() === cursor.getMonth() };
  });
}

export function getDetailWidthLimits(viewportWidth, sidebarCollapsed) {
  const sidebarWidth = sidebarCollapsed ? 72 : 188;
  const contentWidth = Math.max(0, viewportWidth - sidebarWidth);
  const minWidth = Math.min(DETAIL_RAIL_MIN, Math.max(280, contentWidth - MAIN_PANEL_MIN));
  return {
    minWidth,
    maxWidth: Math.max(minWidth, Math.min(DETAIL_RAIL_MAX, contentWidth * DETAIL_RAIL_MAX_RATIO, contentWidth - MAIN_PANEL_MIN)),
  };
}

export function clampDetailWidth(width, viewportWidth, sidebarCollapsed) {
  const { minWidth, maxWidth } = getDetailWidthLimits(viewportWidth, sidebarCollapsed);
  return Math.min(maxWidth, Math.max(minWidth, width));
}

export function filterItemsByCategory(items, category) {
  return category === "전체" ? items : items.filter((item) => item.category === category);
}

export function normalizeSearchText(value) {
  return String(value ?? "").normalize("NFKC").toLocaleLowerCase("ko-KR").replace(/\s+/g, " ").trim();
}

export function matchesSearchQuery(values, query) {
  const terms = normalizeSearchText(query).split(" ").filter(Boolean);
  if (!terms.length) return true;
  const searchableText = normalizeSearchText(values.filter(Boolean).join(" "));
  return terms.every((term) => searchableText.includes(term));
}

export function searchWardrobeItems(items, query) {
  return items.filter((item) => matchesSearchQuery([item.name, item.brand, item.category, item.color], query));
}

export function searchWishlistItems(items, query) {
  return items.filter((item) => matchesSearchQuery([item.name, item.brand, item.category, item.color, item.status], query));
}

export function searchLookbooks(lookbooks, itemMap, query) {
  return lookbooks.filter((lookbook) => {
    const itemValues = lookbook.itemIds.flatMap((id) => {
      const item = itemMap[id];
      return item ? [item.name, item.brand, item.category, item.color] : [];
    });
    return matchesSearchQuery([lookbook.name, lookbook.memo, ...itemValues], query);
  });
}

export function nextLookbookName(lookbooks) {
  const usedNumbers = new Set(lookbooks.flatMap((lookbook) => {
    const match = String(lookbook?.name || "").trim().match(/^룩북\s+(\d+)$/);
    return match ? [Number(match[1])] : [];
  }));
  let number = 1;
  while (usedNumbers.has(number)) number += 1;
  return `룩북 ${number}`;
}

export function reorderItemIds(ids, sourceId, targetId) {
  if (!sourceId || sourceId === targetId) return ids;
  const next = [...ids];
  const sourceIndex = next.indexOf(sourceId);
  const targetIndex = next.indexOf(targetId);
  if (sourceIndex < 0 || targetIndex < 0) return ids;
  next.splice(targetIndex, 0, next.splice(sourceIndex, 1)[0]);
  return next;
}

export function moveItemByOffset(ids, id, direction) {
  const sourceIndex = ids.indexOf(id);
  if (sourceIndex < 0) return ids;
  const targetIndex = Math.min(ids.length - 1, Math.max(0, sourceIndex + direction));
  if (sourceIndex === targetIndex) return ids;
  const next = [...ids];
  next.splice(targetIndex, 0, next.splice(sourceIndex, 1)[0]);
  return next;
}

export function removeItemReferences(outfits, lookbooks, itemId) {
  return {
    outfits: Object.fromEntries(Object.entries(outfits).map(([date, ids]) => [date, ids.filter((id) => id !== itemId)])),
    lookbooks: lookbooks
      .map((lookbook) => ({ ...lookbook, itemIds: lookbook.itemIds.filter((id) => id !== itemId) }))
      .filter((lookbook) => lookbook.itemIds.length),
  };
}

export function validItemIds(ids, itemMap) {
  return ids.filter((id) => Boolean(itemMap[id]));
}

export function createWardrobeResetState(defaultCategories) {
  return {
    items: [],
    outfits: {},
    outfitNotes: {},
    lookbooks: [],
    categories: [...defaultCategories],
  };
}

export function wardrobeItemsToExcelRows(items) {
  return items.map((item) => [
    String(item.name || "").trim(),
    String(item.category || "").trim(),
    String(item.brand || "").trim(),
    String(item.color || "").trim(),
    String(item.url || "").trim(),
    /^https?:\/\//i.test(String(item.image || "").trim()) ? String(item.image).trim() : "",
  ]);
}

export function parseWardrobeExcelRows(rows, idSeed = Date.now()) {
  return rows
    .filter((row) => {
      const name = String(row["이름*"] || row["이름"] || "").trim();
      const category = String(row["분류*"] || row["분류"] || "").trim();
      return name && category && !name.startsWith("예시");
    })
    .map((row, index) => ({
      id: `excel-${idSeed}-${index}`,
      name: String(row["이름*"] || row["이름"]).trim(),
      category: String(row["분류*"] || row["분류"]).trim(),
      brand: String(row["브랜드"] || "").trim(),
      color: String(row["색상"] || "").trim(),
      url: String(row["상품 URL"] || row["상품URL"] || row["URL"] || "").trim(),
      image: String(row["이미지 URL"] || row["이미지URL"] || "").trim(),
    }));
}

export function createBackupPayload(data, preferences = {}, exportedAt = new Date().toISOString()) {
  return {
    format: BACKUP_FORMAT,
    version: BACKUP_VERSION,
    exportedAt,
    data: {
      items: data.items,
      categories: data.categories,
      outfits: data.outfits,
      outfitNotes: data.outfitNotes,
      lookbooks: data.lookbooks,
      wishlist: data.wishlist,
      ...(Array.isArray(data.inspirations) ? { inspirations: data.inspirations } : {}),
    },
    preferences: {
      theme: preferences.theme,
      sidebarCollapsed: preferences.sidebarCollapsed,
      detailWidth: preferences.detailWidth,
      ...(preferences.wardrobeView ? { wardrobeView: preferences.wardrobeView } : {}),
      ...(preferences.wardrobeSort ? { wardrobeSort: preferences.wardrobeSort } : {}),
      ...(preferences.lookbookSort ? { lookbookSort: preferences.lookbookSort } : {}),
      ...(preferences.wishlistSort ? { wishlistSort: preferences.wishlistSort } : {}),
    },
  };
}

export function validateBackupPayload(payload) {
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) throw new Error("올바른 MyCloset 백업 파일이 아니에요.");
  if (payload.format !== BACKUP_FORMAT) throw new Error("MyCloset에서 만든 백업 파일만 가져올 수 있어요.");
  if (payload.version !== BACKUP_VERSION) throw new Error("지원하지 않는 백업 파일 버전이에요.");
  const data = payload.data;
  if (!data || typeof data !== "object" || Array.isArray(data)) throw new Error("백업 데이터가 비어 있어요.");
  const requiredArrays = ["items", "categories", "lookbooks", "wishlist"];
  const requiredObjects = ["outfits", "outfitNotes"];
  if (requiredArrays.some((key) => !Array.isArray(data[key])) || requiredObjects.some((key) => !data[key] || typeof data[key] !== "object" || Array.isArray(data[key]))) {
    throw new Error("백업 데이터의 형식이 올바르지 않아요.");
  }
  const isRecord = (value) => Boolean(value) && typeof value === "object" && !Array.isArray(value);
  const hasStringId = (value) => isRecord(value) && typeof value.id === "string" && value.id.length > 0;
  const collectionsAreValid = data.items.every(hasStringId)
    && data.lookbooks.every((lookbook) => hasStringId(lookbook) && Array.isArray(lookbook.itemIds) && lookbook.itemIds.every((id) => typeof id === "string"))
    && data.wishlist.every(hasStringId)
    && Object.values(data.outfits).every((ids) => Array.isArray(ids) && ids.every((id) => typeof id === "string"))
    && Object.values(data.outfitNotes).every((note) => typeof note === "string");
  const inspirationsAreValid = data.inspirations === undefined
    || (Array.isArray(data.inspirations) && data.inspirations.every((pin) => hasStringId(pin) && typeof pin.image === "string" && pin.image.length > 0));
  if (!collectionsAreValid || !inspirationsAreValid) throw new Error("백업 데이터에 올바르지 않은 항목이 있어요.");
  const categories = data.categories.filter((name) => typeof name === "string" && name.trim()).map((name) => name.trim());
  if (!categories.length) throw new Error("백업 파일에 사용할 수 있는 카테고리가 없어요.");
  return {
    data: {
      items: data.items,
      categories: [...new Set(categories)],
      outfits: data.outfits,
      outfitNotes: data.outfitNotes,
      lookbooks: data.lookbooks,
      wishlist: data.wishlist,
      ...(Array.isArray(data.inspirations) ? { inspirations: data.inspirations } : {}),
    },
    preferences: payload.preferences && typeof payload.preferences === "object" && !Array.isArray(payload.preferences) ? payload.preferences : {},
  };
}
