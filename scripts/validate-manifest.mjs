import { readFileSync } from "node:fs";
import { parseStringPromise } from "xml2js";
import https from "node:https";
import { HttpsProxyAgent } from "https-proxy-agent";
const path = process.argv.find((s) => s.endsWith(".xml")) || "manifest.xml";
const xml = readFileSync(path, "utf8");
const { OfficeApp: app } = await parseStringPromise(xml);
function requireValue(condition, message) {
  if (!condition) throw new Error(message);
}
requireValue(
  app.$["xsi:type"] === "TaskPaneApp",
  "Must be a TaskPaneApp XML manifest",
);
requireValue(
  app.Hosts[0].Host[0].$.Name === "Presentation",
  "Host must be Presentation",
);
requireValue(
  app.Permissions[0] === "ReadWriteDocument",
  "Editing permission missing",
);
requireValue(
  app.Requirements[0].Sets[0].Set.some(
    (s) => s.$.Name === "PowerPointApi" && s.$.MinVersion === "1.5",
  ),
  "API requirement missing",
);
const overrides = app.VersionOverrides[0];
requireValue(
  overrides.Hosts[0].Host[0].$["xsi:type"] === "Presentation",
  "Override host mismatch",
);
const resources = overrides.Resources[0];
const ids = new Set();
for (const key of [
  "bt:Images",
  "bt:Urls",
  "bt:ShortStrings",
  "bt:LongStrings",
]) {
  for (const [kind, entries] of Object.entries(resources[key][0])) {
    if (kind === "$") continue;
    for (const entry of entries) {
      requireValue(!ids.has(entry.$.id), `Duplicate resource ${entry.$.id}`);
      ids.add(entry.$.id);
    }
  }
}
for (const match of xml.matchAll(/resid="([^"]+)"/g))
  requireValue(ids.has(match[1]), `Missing resource ${match[1]}`);
for (const match of xml.matchAll(/DefaultValue="(https?:[^"]+)"/g))
  requireValue(match[1].startsWith("https://"), "All URLs must use HTTPS");
requireValue(
  xml.includes('xsi:type="ShowTaskpane"'),
  "Ribbon taskpane action missing",
);
console.log(
  "Local XML and resource checks passed (not full Microsoft schema validation).",
);
if (!process.argv.includes("--online")) process.exit(0);
const endpoint =
  "https://validationgateway.omex.office.net/package/api/check?clientId=devx";
const proxy = process.env.HTTPS_PROXY || process.env.https_proxy;
const response = await new Promise((resolve, reject) => {
  const req = https.request(
    endpoint,
    {
      method: "POST",
      agent: proxy ? new HttpsProxyAgent(proxy) : undefined,
      headers: {
        "Content-Type": "application/xml",
        "Content-Length": Buffer.byteLength(xml),
      },
    },
    (res) => {
      let data = "";
      res.on("data", (chunk) => {
        data += chunk;
      });
      res.on("end", () => resolve({ status: res.statusCode, body: data }));
    },
  );
  req.setTimeout(30000, () =>
    req.destroy(new Error("Manifest service timeout")),
  );
  req.on("error", reject);
  req.end(xml);
});
if (response.status !== 200)
  throw new Error(
    `Microsoft validator unavailable: HTTP ${response.status}. Check network access to validationgateway.omex.office.net.`,
  );
const report = JSON.parse(response.body.trim());
console.log(JSON.stringify(report, null, 2));
if (String(report.status).toLowerCase() !== "accepted") process.exitCode = 1;
