// 最小範例的後端：活頁簿 JSON 快照存檔 + 編輯鎖 + 版本號檢查。
// 只用 Node 內建模組，可以掛在 Vite dev server 上，也可以用 server.mjs 單獨執行。
// 資料存在 dataDir 底下，每份活頁簿一個 JSON 檔；鎖放在記憶體裡（重啟即釋放）。

import { Buffer } from 'node:buffer';
import { mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';

/** 鎖的有效時間。前端每 10 秒續約一次，關掉分頁 30 秒後鎖會自動失效。 */
export const LOCK_TTL_MS = 30_000;
const MAX_BODY_BYTES = 20 * 1024 * 1024;
const ID_PATTERN = /^[\w-]{1,64}$/;

class HttpError extends Error {
    constructor(status, message) {
        super(message);
        this.status = status;
    }
}

async function readJson(req) {
    let size = 0;
    const chunks = [];
    for await (const chunk of req) {
        size += chunk.length;
        if (size > MAX_BODY_BYTES) {
            throw new HttpError(413, 'payload_too_large');
        }
        chunks.push(chunk);
    }
    if (!chunks.length) {
        return {};
    }
    try {
        return JSON.parse(Buffer.concat(chunks).toString('utf8'));
    } catch {
        throw new HttpError(400, 'invalid_json');
    }
}

function send(res, status, payload) {
    res.statusCode = status;
    res.setHeader('content-type', 'application/json; charset=utf-8');
    res.setHeader('cache-control', 'no-store');
    res.end(JSON.stringify(payload));
}

/**
 * @param {{ dataDir: string, now?: () => number }} options
 * @returns {(req: import('node:http').IncomingMessage, res: import('node:http').ServerResponse, next?: () => void) => Promise<void>} 可直接當作 Node http 或 Connect 中介層使用的處理函式
 */
export function createSaveLockApi({ dataDir, now = Date.now }) {
    /** @type {Map<string, { clientId: string, userName: string, expiresAt: number }>} */
    const locks = new Map();

    function activeLock(id) {
        const lock = locks.get(id);
        if (lock && lock.expiresAt <= now()) {
            locks.delete(id);
            return undefined;
        }
        return lock;
    }

    function publicLock(lock) {
        return lock ? { userName: lock.userName, expiresAt: lock.expiresAt } : null;
    }

    async function readRecord(id) {
        try {
            return JSON.parse(await readFile(path.join(dataDir, `${id}.json`), 'utf8'));
        } catch (error) {
            if (error.code === 'ENOENT') {
                return { version: 0, snapshot: null, savedAt: null, savedBy: null };
            }
            throw error;
        }
    }

    async function writeRecord(id, record) {
        await mkdir(dataDir, { recursive: true });
        const file = path.join(dataDir, `${id}.json`);
        const tmp = `${file}.${process.pid}.tmp`;
        await writeFile(tmp, JSON.stringify(record), 'utf8');
        await rename(tmp, file); // 先寫暫存檔再改名，避免寫到一半當機留下壞檔
    }

    return async function handle(req, res, next) {
        const url = new URL(req.url ?? '/', 'http://localhost');
        const match = url.pathname.match(/^\/api\/workbooks\/([^/]+)(\/lock|\/unlock)?$/);
        if (!match) {
            if (next) next();
            else send(res, 404, { error: 'not_found' });
            return;
        }

        const [, id, action] = match;
        if (!ID_PATTERN.test(id)) {
            send(res, 400, { error: 'invalid_id' });
            return;
        }

        try {
            // GET /api/workbooks/:id → 讀取快照、版本號與目前誰在編輯
            if (!action && req.method === 'GET') {
                const record = await readRecord(id);
                send(res, 200, { ...record, lock: publicLock(activeLock(id)) });
                return;
            }

            const body = await readJson(req);
            const clientId = typeof body.clientId === 'string' ? body.clientId : '';
            if (!clientId) {
                send(res, 400, { error: 'missing_client_id' });
                return;
            }

            // POST /api/workbooks/:id/lock → 取得或續約編輯鎖
            if (action === '/lock' && req.method === 'POST') {
                const lock = activeLock(id);
                if (lock && lock.clientId !== clientId) {
                    send(res, 423, { error: 'locked', lock: publicLock(lock) });
                    return;
                }
                const granted = {
                    clientId,
                    userName: String(body.userName || '匿名使用者').slice(0, 50),
                    expiresAt: now() + LOCK_TTL_MS,
                };
                locks.set(id, granted);
                send(res, 200, { lock: publicLock(granted) });
                return;
            }

            // POST /api/workbooks/:id/unlock → 釋放編輯鎖（前端關閉分頁時用 sendBeacon 呼叫）
            if (action === '/unlock' && req.method === 'POST') {
                if (activeLock(id)?.clientId === clientId) {
                    locks.delete(id);
                }
                send(res, 200, { ok: true });
                return;
            }

            // PUT /api/workbooks/:id → 存檔：必須持有鎖，且 baseVersion 等於伺服器目前的版本
            if (!action && req.method === 'PUT') {
                const lock = activeLock(id);
                if (!lock || lock.clientId !== clientId) {
                    send(res, 423, { error: 'lock_required', lock: publicLock(lock) });
                    return;
                }
                if (!body.snapshot || typeof body.snapshot !== 'object') {
                    send(res, 400, { error: 'missing_snapshot' });
                    return;
                }
                const current = await readRecord(id);
                if (body.baseVersion !== current.version) {
                    send(res, 409, { error: 'version_conflict', version: current.version });
                    return;
                }
                const record = {
                    version: current.version + 1,
                    snapshot: body.snapshot,
                    savedAt: new Date(now()).toISOString(),
                    savedBy: lock.userName,
                };
                await writeRecord(id, record);
                lock.expiresAt = now() + LOCK_TTL_MS;
                send(res, 200, { version: record.version, savedAt: record.savedAt });
                return;
            }

            send(res, 405, { error: 'method_not_allowed' });
        } catch (error) {
            if (error instanceof HttpError) {
                send(res, error.status, { error: error.message });
                return;
            }
            console.error('[save-lock api]', error);
            send(res, 500, { error: 'internal_error' });
        }
    };
}
