import { handleProductMetadataRequest } from "../worker/product-scraper.js";

export default {
  fetch(request) {
    return handleProductMetadataRequest(request);
  },
};
