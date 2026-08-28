// @lovable.dev/vite-tanstack-config already includes the following — do NOT add them manually
// or the app will break with duplicate plugins:
//   - TanStack devtools (dev-only, first), tanstackStart, viteReact, tailwindcss, tsConfigPaths,
//     nitro (build-only using cloudflare as a default target), VITE_* env injection, @ path alias,
//     React/TanStack dedupe, error logger plugins, and sandbox detection (port/host/strictPort).
// You can pass additional config via defineConfig({ vite: { ... }, etc... }) if needed.
import { defineConfig } from "@lovable.dev/vite-tanstack-config";
import { VitePWA } from "vite-plugin-pwa";
import { nitro } from "nitro/vite";

// Nome del repository GitHub: serve per far funzionare i link/asset quando l'app
// è pubblicata su https://<utente>.github.io/<nome-repo>/.
// Se invece pubblichi su <utente>.github.io (repo "root"), imposta BASE_PATH su "/".
const BASE_PATH = "/Turni/";

export default defineConfig({
  tanstackStart: {
    // Redirect TanStack Start's bundled server entry to src/server.ts (our SSR error wrapper).
    // nitro/vite builds from this
    server: { entry: "server" },
  },
  vite: {
    base: BASE_PATH,
    plugins: [
      // GitHub Pages serve solo file statici (nessun runtime server), quindi sovrascriviamo
      // il preset Nitro di default (cloudflare) con "static": prerenderizza tutto in HTML/JS
      // statici in .output/public. L'app è interamente client-side, quindi è sicuro farlo.
      // "prerender.routes" indica esplicitamente quali pagine generare come HTML: senza questa
      // opzione Nitro esporta solo gli asset (JS/CSS/icone) ma NESSUN index.html, causando 404
      // su GitHub Pages. L'app ha un'unica rotta ("/"), quindi basta prerenderizzare quella.
      nitro({
        preset: "static",
        prerender: {
          routes: ["/"],
          crawlLinks: true,
        },
      }),
      VitePWA({
        registerType: "autoUpdate",
        injectRegister: null,
        filename: "sw.js",
        devOptions: { enabled: false },
        manifest: false,
        workbox: {
          globPatterns: ["**/*.{js,css,html,png,svg,webmanifest}"],
          navigateFallback: "/",
          navigateFallbackDenylist: [/^\/~oauth/, /^\/api\//],
          runtimeCaching: [
            {
              urlPattern: ({ request }: { request: Request }) => request.mode === "navigate",
              handler: "NetworkFirst",
              options: {
                cacheName: "html-navigations",
                networkTimeoutSeconds: 5,
              },
            },
            {
              urlPattern: ({ request, sameOrigin }: { request: Request; sameOrigin: boolean }) =>
                sameOrigin &&
                (request.destination === "script" ||
                  request.destination === "style" ||
                  request.destination === "font" ||
                  request.destination === "image"),
              handler: "CacheFirst",
              options: {
                cacheName: "static-assets",
                expiration: { maxEntries: 200, maxAgeSeconds: 60 * 60 * 24 * 30 },
              },
            },
          ],
        },
      }),
    ],
  },
});
