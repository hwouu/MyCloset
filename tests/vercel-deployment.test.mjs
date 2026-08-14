import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import productMetadataFunction from "../api/product-metadata.js";

test("Vercel serves the Vite client build and falls back to the SPA shell", async () => {
  const config = JSON.parse(await readFile(new URL("../vercel.json", import.meta.url), "utf8"));

  assert.equal(config.framework, "vite");
  assert.equal(config.buildCommand, "npm run build");
  assert.equal(config.outputDirectory, "dist/client");
  assert.deepEqual(config.rewrites, [{ source: "/(.*)", destination: "/index.html" }]);
});

test("Vercel product metadata function delegates to the shared API handler", async () => {
  const response = await productMetadataFunction.fetch(
    new Request("https://example.test/api/product-metadata", { method: "GET" }),
  );

  assert.equal(response.status, 405);
  assert.deepEqual(await response.json(), { error: "POST 요청만 지원합니다." });
});
