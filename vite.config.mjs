import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { handleProductMetadataRequest } from "./worker/product-scraper.js";

function productMetadataApi() {
  return {
    name: "mycloset-product-metadata-api",
    configureServer(server) {
      server.middlewares.use("/api/product-metadata", async (request, response) => {
        const chunks = [];
        for await (const chunk of request) chunks.push(chunk);
        const origin = `http://${request.headers.host || "127.0.0.1"}`;
        const webRequest = new Request(new URL(request.originalUrl || request.url || "/api/product-metadata", origin), {
          method: request.method,
          headers: request.headers,
          body: chunks.length ? Buffer.concat(chunks) : undefined,
        });
        const webResponse = await handleProductMetadataRequest(webRequest);
        response.statusCode = webResponse.status;
        webResponse.headers.forEach((value, key) => response.setHeader(key, value));
        response.end(Buffer.from(await webResponse.arrayBuffer()));
      });
    },
  };
}

export default defineConfig({
  build: {
    outDir: "dist/client",
  },
  optimizeDeps: {
    include: ["react", "react-dom/client"],
  },
  server: {
    host: "0.0.0.0",
    allowedHosts: ["terminal.local"],
    warmup: {
      clientFiles: ["./src/main.jsx"],
    },
  },
  plugins: [react(), productMetadataApi()],
});
