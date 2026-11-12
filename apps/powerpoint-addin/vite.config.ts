import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { verifyCertificates } from "office-addin-dev-certs";
import { homedir } from "node:os";
import { join } from "node:path";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
function developmentHttps() {
  if (process.env.PPT_DEV_CERT || process.env.PPT_DEV_KEY) {
    if (!process.env.PPT_DEV_CERT || !process.env.PPT_DEV_KEY)
      throw new Error("PPT_DEV_CERT와 PPT_DEV_KEY 파일 경로를 모두 지정해 주세요.");
    return { cert: readFileSync(process.env.PPT_DEV_CERT), key: readFileSync(process.env.PPT_DEV_KEY) };
  }
  try {
    if (!verifyCertificates()) throw new Error("인증서가 없거나 신뢰되지 않습니다.");
    const directory = join(homedir(), ".office-addin-dev-certs");
    return { ca: readFileSync(join(directory, "ca.crt")), cert: readFileSync(join(directory, "localhost.crt")), key: readFileSync(join(directory, "localhost.key")) };
  } catch {
    throw new Error("HTTPS 개발 인증서를 확인할 수 없습니다. 먼저 npm run certs를 실행하세요. E_ACCESSDENIED / 0x80070005이면 Windows-Trust-Certificate.cmd 또는 docs/WINDOWS.md의 인증서 복구 절차를 사용하세요. 회사 정책으로 등록이 차단된 경우 관리자 지원이 필요합니다. 서버 시작 시 인증서를 자동 생성하거나 신뢰 설정을 변경하지 않습니다.");
  }
}
export default defineConfig(async ({ command, isPreview }) => ({
  root: fileURLToPath(new URL(".", import.meta.url)),
  plugins: [react(), tailwindcss()],
  server: {
    host: "127.0.0.1",
    port: 3000,
    strictPort: true,
    https: command === "serve" && !isPreview ? developmentHttps() : undefined,
    fs: { allow: [fileURLToPath(new URL("../..", import.meta.url))] },
  },
  build: { outDir: "../../dist", emptyOutDir: true },
}));
