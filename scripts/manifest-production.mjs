import { readFileSync, writeFileSync, mkdirSync, existsSync } from "node:fs";
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
const version=JSON.parse(readFileSync('package.json','utf8')).version;
const template=existsSync('manifest.development.xml')?'manifest.development.xml':'manifest.xml';
const xml = readFileSync(template, 'utf8')
  .replace(/<Id>[^<]+<\/Id>/, '<Id>7a4a06c7-71c9-444f-b16a-589377521b98</Id>')
  .replace('<DisplayName DefaultValue="PPT AI Studio"/>','<DisplayName DefaultValue="PPT AI Studio 온라인"/>')
  .replace('<bt:String id="Tab.Label" DefaultValue="PPT AI Studio"/>','<bt:String id="Tab.Label" DefaultValue="PPT AI Studio 온라인"/>')
  .replaceAll('https://localhost:3000/index.html', 'https://localhost:3000/index.html?addin=hosted&amp;v='+version)
  .replaceAll('https://localhost:3000/commands.html', 'https://localhost:3000/commands.html?addin=hosted&amp;v='+version)
  .replaceAll('https://localhost:3000', url.origin.replaceAll('&','&amp;'));
writeFileSync("artifacts/manifest.production.xml", xml);
console.log("Created artifacts/manifest.production.xml for", url.origin);
