import { defineConfig } from "vite";
import devCerts from "office-addin-dev-certs";
import { readFile } from "node:fs/promises";

async function getHttpsOptions() {
    const keyPath = process.env.HTTPS_KEY_PATH;
    const certPath = process.env.HTTPS_CERT_PATH;

    if (keyPath && certPath) {
        return {
            key: await readFile(keyPath),
            cert: await readFile(certPath)
        };
    }

    return devCerts.getHttpsServerOptions();
}

export default defineConfig(async ({ command }) => {
    const https = command === "serve"
        ? await getHttpsOptions()
        : undefined;

    return {
        server: {
            host: "0.0.0.0",
            port: 3000,
            strictPort: true,
            https,
            proxy: {
                "/api": {
                    target: process.env.BACKEND_URL ?? "http://localhost:8000",
                    changeOrigin: true,
                    rewrite: (path: string) => path.replace(/^\/api/, "")
                }
            }
        },
        preview: {
            host: "0.0.0.0",
            port: 3000,
            strictPort: true
        }
    };
});
