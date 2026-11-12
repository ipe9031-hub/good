import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { getHttpsServerOptions } from "office-addin-dev-certs";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
export default defineConfig(async ({ command, isPreview }) => ({
  root: fileURLToPath(new URL(".", import.meta.url)),
  plugins: [react(), tailwindcss()],
  server: {
    host: "127.0.0.1",
    port: 3000,
    strictPort: true,
    https:
      command === "serve" && !isPreview
        ? process.env.PPT_DEV_CERT && process.env.PPT_DEV_KEY
          ? {
              cert: readFileSync(process.env.PPT_DEV_CERT),
              key: readFileSync(process.env.PPT_DEV_KEY),
            }
          : await getHttpsServerOptions()
        : undefined,
    fs: { allow: [fileURLToPath(new URL("../..", import.meta.url))] },
  },
  build: { outDir: "../../dist", emptyOutDir: true },
}));
