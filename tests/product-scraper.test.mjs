import assert from "node:assert/strict";
import test from "node:test";
import {
  extractProductMetadata,
  fetchProductMetadata,
  handleProductMetadataRequest,
  validateProductUrl,
} from "../worker/product-scraper.js";

const MUSINSA_HTML = `<!doctype html><html><head>
  <meta property="og:type" content="product">
  <meta property="og:site_name" content="무신사">
  <meta property="og:title" content="무신사 스탠다드(MUSINSA STANDARD) 빅 브러쉬 카펜터 버뮤다 데님 팬츠 [옐로우 다크 블루] - 사이즈 &amp; 후기 | 무신사">
  <meta property="og:image" content="https://image.msscdn.net/images/goods_img/20241202/4644763/4644763_17435766508836_500.jpg">
  <meta property="og:url" content="https://www.musinsa.com/products/4644763">
  <meta property="og:description" content="제품분류 :바지 &gt; 데님 팬츠 브랜드 : 무신사 스탠다드(MUSINSA STANDARD) 제품번호 : MMENH6Z04-YD 제품 : 빅 브러쉬 카펜터 버뮤다 데님 팬츠 [옐로우 다크 블루] - 35,990">
</head></html>`;

const UNIQLO_URL = "https://www.uniqlo.com/kr/ko/products/E484875-000/00?colorDisplayCode=01&sizeDisplayCode=004";
const UNIQLO_IMAGE = "https://image.uniqlo.com/UQ/ST3/kr/imagesgoods/484875/item/krgoods_01_484875_3x4.jpg";
const UNIQLO_HTML = `<!doctype html><html><head>
  <meta property="og:site_name" content="UNIQLO">
  <meta property="og:title" content="젠더리스 배럴팬츠 | UNIQLO KR">
  <meta property="og:url" content="https://www.uniqlo.com/kr/ko/products/E484875-000/00">
</head><body><script>window.__PRELOADED_STATE__ = ${JSON.stringify({
  selection: { "E484875-000": { colorDisplayCode: "01", sizeDisplayCode: "004" } },
  entity: {
    pdpEntity: {
      "E484875-000": {
        product: {
          name: "배럴팬츠",
          breadcrumbs: { class: { locale: "팬츠" } },
          colors: [
            { displayCode: "32", name: "BEIGE" },
            { displayCode: "01", name: "OFF WHITE" },
          ],
          representative: { color: { displayCode: "32", name: "BEIGE" } },
          images: { main: { "01": { image: UNIQLO_IMAGE } } },
        },
      },
    },
  },
})};</script></body></html>`;

test("extracts normalized wardrobe fields from Musinsa metadata", () => {
  assert.deepEqual(extractProductMetadata(MUSINSA_HTML, "https://www.musinsa.com/products/4644763"), {
    name: "빅 브러쉬 카펜터 버뮤다 데님 팬츠",
    category: "하의",
    brand: "무신사 스탠다드",
    color: "옐로우 다크 블루",
    url: "https://www.musinsa.com/products/4644763",
    image: "https://image.msscdn.net/images/goods_img/20241202/4644763/4644763_17435766508836_500.jpg",
    source: "무신사",
  });
});

test("extracts the selected Uniqlo color and preserves the product query", () => {
  assert.deepEqual(extractProductMetadata(UNIQLO_HTML, UNIQLO_URL), {
    name: "배럴팬츠",
    category: "하의",
    brand: "UNIQLO",
    color: "01 OFF WHITE",
    url: UNIQLO_URL,
    image: UNIQLO_IMAGE,
    source: "UNIQLO",
  });
});

test("uses Product JSON-LD and Open Graph as a generic fallback", () => {
  const html = `<script type="application/ld+json">${JSON.stringify({
    "@type": "Product",
    name: "Classic Linen Shirt",
    brand: { "@type": "Brand", name: "Sample Brand" },
    color: "Sky Blue",
    category: "Shirts",
    image: ["https://shop.example/images/shirt.jpg"],
    url: "https://shop.example/products/shirt",
  })}</script>`;
  const result = extractProductMetadata(html, "https://shop.example/products/shirt");
  assert.equal(result.name, "Classic Linen Shirt");
  assert.equal(result.brand, "Sample Brand");
  assert.equal(result.category, "상의");
  assert.equal(result.color, "Sky Blue");
});

test("rejects local and private network product URLs", () => {
  for (const url of ["http://localhost:5173", "http://127.0.0.1/item", "http://192.168.0.2/item", "file:///tmp/item.html"]) {
    assert.throws(() => validateProductUrl(url));
  }
});

test("follows a validated redirect and extracts the final product", async () => {
  const calls = [];
  const fetchImpl = async (url) => {
    calls.push(String(url));
    if (calls.length === 1) return new Response(null, { status: 302, headers: { location: "https://www.musinsa.com/products/4644763" } });
    return new Response(MUSINSA_HTML, { status: 200, headers: { "content-type": "text/html; charset=utf-8" } });
  };
  const product = await fetchProductMetadata("https://m.musinsa.com/products/4644763", fetchImpl);
  assert.equal(product.name, "빅 브러쉬 카펜터 버뮤다 데님 팬츠");
  assert.equal(calls.length, 2);
});

test("times out slow product pages and recommends direct input", async () => {
  await assert.rejects(
    fetchProductMetadata("https://slow-shop.example/product", () => new Promise(() => {}), 10),
    /직접 입력으로 등록해주세요/,
  );
});

test("API handler returns a clear validation error", async () => {
  const response = await handleProductMetadataRequest(new Request("https://example.test/api/product-metadata", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ url: "http://localhost/private" }),
  }));
  assert.equal(response.status, 400);
  assert.match((await response.json()).error, /내부 네트워크/);
});
