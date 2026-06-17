import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";
import devCerts from "office-addin-dev-certs";

async function httpsOptions() {
  return devCerts.getHttpsServerOptions();
}

export default defineConfig(async ({ command }) => ({
  plugins: [react()],
  server:
    command === "serve"
      ? {
          https: await httpsOptions(),
          proxy: {
            "/api": {
              target: "http://127.0.0.1:3001",
              changeOrigin: true
            }
          }
        }
      : undefined,
  build: {
    outDir: "dist",
    sourcemap: true
  }
}));
