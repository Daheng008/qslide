import { defineConfig } from "vite";
import basicSsl from "@vitejs/plugin-basic-ssl";
import { resolve } from "path";

// Office 要求 https；basic-ssl 提供自签证书（sideload 时信任一次即可）
// base：本地 dev 用根路径 "/"；GitHub Pages 部署时由 PAGES_BASE 注入仓库子路径（如 /qslide/）。
const pagesBase = process.env.PAGES_BASE || "/";
export default defineConfig({
  root: "src",
  publicDir: "public",
  base: pagesBase,
  plugins: [
    basicSsl(),
    {
      // 根路径 / 自动跳转到 taskpane（避免浏览器打开 localhost:3000 时白屏）
      name: "root-redirect",
      configureServer(s) {
        s.middlewares.use((req, _res, next) => {
          if (req.url === "/" || req.url === "") {
            req.url = "/taskpane.html";
          }
          next();
        });
      },
    },
  ],
  server: { port: 3000, strictPort: true },
  build: {
    outDir: "../dist",
    emptyOutDir: true,
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "src/taskpane.html"),
        commands: resolve(__dirname, "src/commands.html"),
      },
    },
  },
});
