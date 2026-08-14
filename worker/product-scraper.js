const MAX_REDIRECTS = 4;
const MAX_HTML_LENGTH = 2_000_000;
const PRODUCT_FETCH_TIMEOUT_MS = 12_000;
const PRODUCT_FETCH_TIMEOUT_MESSAGE = "상품 페이지의 응답 시간이 초과됐어요. 직접 입력으로 등록해주세요.";

const CATEGORY_RULES = [
  ["원피스", /원피스|dress/i],
  ["아우터", /재킷|자켓|코트|점퍼|패딩|가디건|블레이저|후드집업|outer|jacket|coat|cardigan|blazer/i],
  ["하의", /바지|팬츠|데님|진|슬랙스|쇼츠|반바지|스커트|pants|trousers|denim|jeans|shorts|skirt/i],
  ["신발", /신발|스니커즈|로퍼|부츠|샌들|슈즈|운동화|shoes|sneakers|loafer|boots|sandals/i],
  ["가방", /가방|백팩|토트백|숄더백|크로스백|bag|backpack|tote/i],
  ["액세서리", /모자|캡|벨트|목걸이|반지|팔찌|시계|안경|선글라스|양말|hat|cap|belt|necklace|ring|watch|glasses|socks/i],
  ["상의", /셔츠|티셔츠|니트|스웨터|후드|맨투맨|블라우스|탑|shirt|t-shirt|tee|knit|sweater|hoodie|blouse|top/i],
];

function decodeHtml(value = "") {
  return value
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code)))
    .replace(/&#x([\da-f]+);/gi, (_, code) => String.fromCodePoint(Number.parseInt(code, 16)))
    .replace(/\s+/g, " ")
    .trim();
}

function parseAttributes(tag) {
  const attributes = {};
  for (const match of tag.matchAll(/([:\w-]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/g)) {
    attributes[match[1].toLowerCase()] = decodeHtml(match[2] ?? match[3] ?? match[4] ?? "");
  }
  return attributes;
}

function readMeta(html) {
  const meta = {};
  for (const match of html.matchAll(/<meta\b[^>]*>/gi)) {
    const attributes = parseAttributes(match[0]);
    const key = (attributes.property || attributes.name || "").toLowerCase();
    if (key && attributes.content && !meta[key]) meta[key] = attributes.content;
  }
  return meta;
}

function flattenJsonLd(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value.flatMap(flattenJsonLd);
  if (typeof value !== "object") return [];
  return [value, ...flattenJsonLd(value["@graph"])];
}

function readJsonLd(html) {
  const records = [];
  for (const match of html.matchAll(/<script\b[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi)) {
    try { records.push(...flattenJsonLd(JSON.parse(decodeHtml(match[1])))); }
    catch { /* Ignore malformed third-party structured data. */ }
  }
  return records.find((record) => {
    const type = record?.["@type"];
    return Array.isArray(type) ? type.includes("Product") : type === "Product";
  }) || {};
}

function firstImage(value) {
  if (Array.isArray(value)) return firstImage(value[0]);
  if (value && typeof value === "object") return value.url || value.contentUrl || "";
  return typeof value === "string" ? value : "";
}

function safePublicAssetUrl(value, baseUrl) {
  if (!value) return "";
  try {
    const url = new URL(String(value), baseUrl);
    return ["http:", "https:"].includes(url.protocol) ? url.toString() : "";
  } catch { return ""; }
}

function cleanBrand(value = "") {
  const brand = typeof value === "object" ? value.name : value;
  return decodeHtml(String(brand || "")).replace(/\s*\([^)]*\)\s*$/, "").trim();
}

function inferCategory(...values) {
  const haystack = values.filter(Boolean).join(" ");
  return CATEGORY_RULES.find(([, pattern]) => pattern.test(haystack))?.[0] || "상의";
}

function cleanGenericTitle(title = "", siteName = "") {
  let value = decodeHtml(title);
  if (siteName) value = value.replace(new RegExp(`\\s*[-|｜]\\s*${siteName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}.*$`, "i"), "");
  return value.replace(/\s*[-|｜]\s*(사이즈\s*&?\s*후기|공식.*|온라인.*).*$/i, "").trim();
}

function parseMusinsa(meta) {
  const description = meta["og:description"] || meta.description || "";
  const title = meta["og:title"] || "";
  const brandMatch = description.match(/브랜드\s*:\s*(.+?)\s+제품번호\s*:/i);
  const productMatch = description.match(/제품\s*:\s*(.+?)(?:\s+-\s+[\d,]+(?:원)?(?:\s|$)|$)/i);
  const categoryMatch = description.match(/제품분류\s*:\s*(.+?)\s+브랜드\s*:/i);
  let name = decodeHtml(productMatch?.[1] || "");
  let color = "";
  const colorMatch = name.match(/\s*\[([^\]]+)]\s*$/);
  if (colorMatch) {
    color = colorMatch[1].trim();
    name = name.slice(0, colorMatch.index).trim();
  }
  return {
    name: name || cleanGenericTitle(title, meta["og:site_name"]),
    brand: cleanBrand(brandMatch?.[1] || ""),
    categoryText: categoryMatch?.[1] || "",
    color,
  };
}

function readAssignedJson(html, variableName) {
  const pattern = new RegExp(`window\\.${variableName}\\s*=\\s*({[\\s\\S]*?})\\s*;?\\s*<\\/script>`, "i");
  const match = html.match(pattern);
  if (!match) return {};
  try { return JSON.parse(match[1]); }
  catch { return {}; }
}

function parseUniqlo(html, pageUrl, meta) {
  const state = readAssignedJson(html, "__PRELOADED_STATE__");
  const productEntry = Object.values(state?.entity?.pdpEntity || {})[0];
  const product = productEntry?.product || {};
  const selectedFromState = Object.values(state?.selection || {})[0];
  const selectedColorCode = new URL(pageUrl).searchParams.get("colorDisplayCode")
    || selectedFromState?.colorDisplayCode
    || product.representative?.color?.displayCode
    || "";
  const selectedColor = product.colors?.find((color) => String(color.displayCode) === String(selectedColorCode))
    || product.representative?.color;
  const fallbackName = cleanGenericTitle(meta["og:title"] || "", meta["og:site_name"]).replace(/^젠더리스\s+/i, "");

  return {
    name: decodeHtml(product.name || fallbackName),
    brand: "UNIQLO",
    categoryText: product.breadcrumbs?.class?.locale || product.breadcrumbs?.category?.locale || "",
    color: [selectedColor?.displayCode || selectedColorCode, selectedColor?.name].filter(Boolean).join(" "),
    image: product.images?.main?.[selectedColorCode]?.image || "",
  };
}

export function extractProductMetadata(html, pageUrl) {
  const meta = readMeta(html);
  const product = readJsonLd(html);
  const hostname = new URL(pageUrl).hostname.toLowerCase();
  const isMusinsa = hostname === "musinsa.com" || hostname.endsWith(".musinsa.com");
  const isUniqlo = hostname === "uniqlo.com" || hostname.endsWith(".uniqlo.com");
  const specialized = isMusinsa ? parseMusinsa(meta) : isUniqlo ? parseUniqlo(html, pageUrl, meta) : {};
  const genericName = cleanGenericTitle(product.name || meta["og:title"] || meta["twitter:title"] || "", meta["og:site_name"]);
  const name = specialized.name || genericName;
  const brand = specialized.brand || cleanBrand(product.brand || meta["product:brand"] || "");
  const color = specialized.color || decodeHtml(String(product.color || meta["product:color"] || ""));
  const categoryText = specialized.categoryText || product.category || meta["product:category"] || "";
  const image = safePublicAssetUrl(specialized.image || firstImage(product.image) || meta["og:image:secure_url"] || meta["og:image"] || meta["twitter:image"] || "", pageUrl);
  const canonicalUrl = isUniqlo ? pageUrl : (safePublicAssetUrl(meta["og:url"] || product.url || pageUrl, pageUrl) || pageUrl);

  return {
    name,
    category: inferCategory(categoryText, name, meta.description, meta["og:description"]),
    brand,
    color,
    url: canonicalUrl,
    image,
    source: isMusinsa ? "무신사" : isUniqlo ? "UNIQLO" : (meta["og:site_name"] || hostname.replace(/^www\./, "")),
  };
}

function isPrivateIpv4(hostname) {
  const parts = hostname.split(".").map(Number);
  if (parts.length !== 4 || parts.some((part) => !Number.isInteger(part) || part < 0 || part > 255)) return false;
  return parts[0] === 10
    || parts[0] === 127
    || (parts[0] === 169 && parts[1] === 254)
    || (parts[0] === 172 && parts[1] >= 16 && parts[1] <= 31)
    || (parts[0] === 192 && parts[1] === 168)
    || parts[0] === 0;
}

export function validateProductUrl(value) {
  let url;
  try { url = new URL(String(value || "").trim()); }
  catch { throw new Error("http 또는 https 상품 URL을 입력해주세요."); }
  if (!["http:", "https:"].includes(url.protocol)) throw new Error("http 또는 https 상품 URL만 사용할 수 있어요.");
  const hostname = url.hostname.toLowerCase().replace(/^\[|\]$/g, "");
  if (!hostname || hostname === "localhost" || hostname.endsWith(".localhost") || hostname.endsWith(".local") || hostname === "::1" || isPrivateIpv4(hostname)) {
    throw new Error("내부 네트워크 주소는 조회할 수 없어요.");
  }
  return url;
}

async function fetchWithSafeRedirects(initialUrl, fetchImpl, timeoutMs) {
  const controller = new AbortController();
  let timeoutId;
  const timeout = new Promise((_, reject) => {
    timeoutId = setTimeout(() => {
      controller.abort();
      reject(new Error(PRODUCT_FETCH_TIMEOUT_MESSAGE));
    }, timeoutMs);
  });
  const fetchRedirectChain = async () => {
    let current = validateProductUrl(initialUrl);
    for (let redirectCount = 0; redirectCount <= MAX_REDIRECTS; redirectCount += 1) {
      const response = await fetchImpl(current, {
        headers: {
          accept: "text/html,application/xhtml+xml",
          // Several commerce sites return a deliberately delayed response to crawler-style agents.
          "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36",
        },
        redirect: "manual",
        signal: controller.signal,
      });
      if ([301, 302, 303, 307, 308].includes(response.status)) {
        const location = response.headers.get("location");
        if (!location) throw new Error("상품 페이지의 이동 주소를 확인할 수 없어요.");
        current = validateProductUrl(new URL(location, current).toString());
        continue;
      }
      if (!response.ok) throw new Error(`상품 페이지를 불러오지 못했어요. (${response.status})`);
      const contentType = response.headers.get("content-type") || "";
      if (!contentType.includes("text/html") && !contentType.includes("application/xhtml+xml")) throw new Error("HTML 상품 페이지만 분석할 수 있어요.");
      const declaredLength = Number(response.headers.get("content-length") || 0);
      if (declaredLength > MAX_HTML_LENGTH) throw new Error("상품 페이지가 너무 커서 분석할 수 없어요.");
      const html = await response.text();
      if (html.length > MAX_HTML_LENGTH) throw new Error("상품 페이지가 너무 커서 분석할 수 없어요.");
      return { html, finalUrl: current.toString() };
    }
    throw new Error("상품 페이지가 너무 여러 번 이동했어요.");
  };
  try {
    return await Promise.race([fetchRedirectChain(), timeout]);
  } catch (error) {
    if (error?.name === "AbortError") throw new Error(PRODUCT_FETCH_TIMEOUT_MESSAGE);
    throw error;
  } finally {
    clearTimeout(timeoutId);
    controller.abort();
  }
}

export async function fetchProductMetadata(value, fetchImpl = fetch, timeoutMs = PRODUCT_FETCH_TIMEOUT_MS) {
  const { html, finalUrl } = await fetchWithSafeRedirects(value, fetchImpl, timeoutMs);
  const metadata = extractProductMetadata(html, finalUrl);
  if (!metadata.name && !metadata.image) throw new Error("자동으로 찾은 정보가 없어요. 직접 입력으로 등록해주세요.");
  return metadata;
}

export function jsonResponse(payload, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" },
  });
}

export async function handleProductMetadataRequest(request, fetchImpl = fetch) {
  if (request.method !== "POST") return jsonResponse({ error: "POST 요청만 지원합니다." }, 405);
  try {
    const body = await request.json();
    const product = await fetchProductMetadata(body?.url, fetchImpl);
    return jsonResponse({ product });
  } catch (error) {
    const message = error instanceof Error ? error.message : "상품 정보를 가져오지 못했어요.";
    const status = /입력|주소|URL|네트워크/.test(message) ? 400 : 422;
    return jsonResponse({ error: message }, status);
  }
}
