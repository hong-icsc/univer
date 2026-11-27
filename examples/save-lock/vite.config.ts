import type { Plugin } from 'vite';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';
// @ts-expect-error 純 JS 模組，沒有型別宣告
import { createSaveLockApi } from './api.mjs';

/** 把存檔與編輯鎖 API 掛在 Vite dev server 上，一個指令就能同時啟動前後端。 */
function saveLockApi(): Plugin {
    return {
        name: 'save-lock-api',
        configureServer(server) {
            const api = createSaveLockApi({ dataDir: fileURLToPath(new URL('./.data', import.meta.url)) });
            server.middlewares.use((req, res, next) => {
                void api(req, res, next);
            });
        },
    };
}

/** 與工作台一樣，忽略各套件 src/global.css，避免載入開發用的全域樣式。 */
function ignoreWorkspaceGlobalCss(): Plugin {
    return {
        name: 'ignore-workspace-global-css',
        enforce: 'pre',
        transform: {
            filter: { id: /\/packages\/.*\/src\/global\.css$/ },
            handler() {
                return '';
            },
        },
    };
}

export default defineConfig(({ command }) => ({
    root: fileURLToPath(new URL('.', import.meta.url)),
    // 與工作台相同：開發時先打包，避免瀏覽器逐一載入上千個原始碼模組
    experimental: {
        bundledDev: command === 'serve',
    },
    plugins: [saveLockApi(), ignoreWorkspaceGlobalCss()],
    server: {
        port: 5174,
        strictPort: true,
    },
}));
