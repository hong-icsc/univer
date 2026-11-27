// 單獨執行後端 API（不經過 Vite），方便接到正式的前端建置或用 curl 測試。
// 用法：node examples/save-lock/server.mjs  （預設 port 3030，可用 PORT 環境變數調整）

import { createServer } from 'node:http';
import process from 'node:process';
import { fileURLToPath } from 'node:url';
import { createSaveLockApi } from './api.mjs';

const port = Number(process.env.PORT) || 3030;
const dataDir = process.env.DATA_DIR || fileURLToPath(new URL('./.data', import.meta.url));
const api = createSaveLockApi({ dataDir });

createServer((req, res) => {
    void api(req, res);
}).listen(port, () => {
    process.stdout.write(`save-lock API listening on http://localhost:${port}  (data: ${dataDir})\n`);
});
