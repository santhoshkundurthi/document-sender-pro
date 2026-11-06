import { defineConfig, loadEnv } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { mailtrapTestSendMiddleware } from "./server/mailtrap-test-send.js";

const serverOnlyEnvironmentVariables = [
  "MAILTRAP_API_TOKEN",
  "MAILTRAP_ALLOWED_RECIPIENTS",
  "MAILTRAP_ALLOWED_FROM_EMAILS",
];

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), "");
  for (const key of serverOnlyEnvironmentVariables) {
    if (env[key] && !process.env[key]) process.env[key] = env[key];
  }

  return {
    plugins: [
      react(),
      tailwindcss(),
      {
        name: "mailtrap-test-send-api",
        configureServer(server) {
          server.middlewares.use(mailtrapTestSendMiddleware);
        },
        configurePreviewServer(server) {
          server.middlewares.use(mailtrapTestSendMiddleware);
        },
      },
    ],
    server: {
      host: "0.0.0.0",
      port: 4173,
      strictPort: false,
    },
    preview: {
      host: "0.0.0.0",
      port: 4174,
      strictPort: false,
    },
  };
});