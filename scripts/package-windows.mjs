import { execFileSync } from "node:child_process";
import { readFileSync, readdirSync, writeFileSync, mkdirSync } from "node:fs";
import { join } from "node:path";
import { createHash } from "node:crypto";
import { zipSync, strToU8 } from "fflate";
const tracked = execFileSync("git", ["ls-files", "-z"], { encoding: "utf8" })
  .split("\0")
  .filter(Boolean);
if (!tracked.includes("package-lock.json"))
  throw new Error("Stage source and lockfile before packaging.");
const files = {};
for (const file of tracked) {
  if (
    /^(?:\.git(?:\/|$)|node_modules\/|artifacts\/|downloads\/|\.env$)/.test(
      file,
    )
  )
    continue;
  files[`PPT-AI-Studio/${file}`] = new Uint8Array(readFileSync(file));
}
function addDir(path) {
  for (const item of readdirSync(path, { withFileTypes: true })) {
    const name = join(path, item.name);
    if (item.isDirectory()) addDir(name);
    else
      files[`PPT-AI-Studio/${name.replaceAll("\\", "/")}`] = new Uint8Array(
        readFileSync(name),
      );
  }
}
addDir("dist");
files["PPT-AI-Studio/BUILD-INFO.txt"] = strToU8(
  `PPT AI Studio 0.2.0\nBuilt: ${new Date().toISOString()}\nSource commit: ${execFileSync("git", ["rev-parse", "HEAD"], { encoding: "utf8" }).trim()}\nWindows PowerPoint host validation: NOT RUN in Linux cloud\nNode.js and npm installation required. Read docs/WINDOWS.md.\n`,
);
mkdirSync("artifacts", { recursive: true });
const data = zipSync(files, { level: 6 });
const name = "PPT-AI-Studio-Phase2-Windows.zip";
writeFileSync(`artifacts/${name}`, data);
const sha = createHash("sha256").update(data).digest("hex");
writeFileSync(`artifacts/${name}.sha256`, `${sha}  ${name}\n`);
console.log(name, data.length, "bytes; SHA256", sha);
