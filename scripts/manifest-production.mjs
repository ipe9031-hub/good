import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
const input = process.argv[2];
if (!input)
  throw new Error(
    "Usage: node scripts/manifest-production.mjs https://your-domain.vercel.app",
  );
const url = new URL(input);
if (
  url.protocol !== "https:" ||
  url.username ||
  url.password ||
  url.pathname !== "/" ||
  url.search ||
  url.hash ||
  ["localhost", "127.0.0.1"].includes(url.hostname)
)
  throw new Error(
    "HTTPS production origin without a path or credentials is required.",
  );
mkdirSync("artifacts", { recursive: true });
const xml = readFileSync("manifest.xml", "utf8").replaceAll(
  "https://localhost:3000",
  url.origin.replaceAll("&", "&amp;"),
);
writeFileSync("artifacts/manifest.production.xml", xml);
console.log("Created artifacts/manifest.production.xml for", url.origin);
