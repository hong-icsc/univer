/**
 * 最小範例：開源版 Univer 試算表 + 存檔到後端 + 編輯鎖。
 *
 * 流程：
 * 1. 開啟時向後端讀取快照（沒有就建立空白活頁簿）。
 * 2. 嘗試取得編輯鎖：拿到就能編輯，並每 10 秒續約；拿不到就切成唯讀，顯示誰正在編輯。
 * 3. 存檔時把 workbook.save() 的 JSON 快照連同 baseVersion 送出，後端確認持有鎖且版本沒變才寫入。
 * 4. 唯讀模式每 5 秒檢查一次，鎖釋放或有新版本時提示重新載入。
 */
import type { IWorkbookData } from '@univerjs/core';
import { CommandType, LocaleType, mergeLocales } from '@univerjs/core';
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting';
import conditionalFormattingLocale from '@univerjs/preset-sheets-conditional-formatting/locales/zh-TW';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import coreLocale from '@univerjs/preset-sheets-core/locales/zh-TW';
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation';
import dataValidationLocale from '@univerjs/preset-sheets-data-validation/locales/zh-TW';
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter';
import filterLocale from '@univerjs/preset-sheets-filter/locales/zh-TW';
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace';
import findReplaceLocale from '@univerjs/preset-sheets-find-replace/locales/zh-TW';
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort';
import sortLocale from '@univerjs/preset-sheets-sort/locales/zh-TW';
import { createUniver } from '@univerjs/presets';
import '../src/global.css';
import '../src/sheets/global.css';

const HEARTBEAT_MS = 10_000;
const READONLY_POLL_MS = 5_000;

interface ILockInfo {
    userName: string;
    expiresAt: number;
}

interface IWorkbookRecord {
    version: number;
    snapshot: Partial<IWorkbookData> | null;
    savedAt: string | null;
    savedBy: string | null;
    lock: ILockInfo | null;
}

const params = new URLSearchParams(location.search);
const workbookId = /^[\w-]{1,64}$/.test(params.get('id') ?? '') ? params.get('id')! : 'demo';
const apiBase = `/api/workbooks/${workbookId}`;
const clientId = getClientId();

const elements = {
    userName: document.querySelector<HTMLInputElement>('#user-name')!,
    badge: document.querySelector<HTMLElement>('#mode-badge')!,
    save: document.querySelector<HTMLButtonElement>('#save-button')!,
    reload: document.querySelector<HTMLButtonElement>('#reload-button')!,
    status: document.querySelector<HTMLElement>('#save-status')!,
};

const state = {
    mode: 'readonly' as 'edit' | 'readonly' | 'error',
    version: 0,
    dirty: false,
    saving: false,
};

elements.userName.value = getUserName();
elements.userName.addEventListener('change', () => {
    const name = elements.userName.value.trim() || getUserName();
    elements.userName.value = name;
    safeStorage(() => localStorage.setItem('save-lock:user-name', name));
});
elements.reload.addEventListener('click', () => location.reload());

const { univerAPI } = createUniver({
    locale: LocaleType.ZH_TW,
    locales: {
        [LocaleType.ZH_TW]: mergeLocales(
            coreLocale,
            filterLocale,
            sortLocale,
            findReplaceLocale,
            conditionalFormattingLocale,
            dataValidationLocale
        ),
    },
    presets: [
        UniverSheetsCorePreset({ container: 'app' }),
        UniverSheetsFilterPreset(),
        UniverSheetsSortPreset(),
        UniverSheetsFindReplacePreset(),
        UniverSheetsConditionalFormattingPreset(),
        UniverSheetsDataValidationPreset(),
    ],
});

type FWorkbook = ReturnType<typeof univerAPI.createWorkbook>;

void start();

async function start() {
    let record: IWorkbookRecord;
    try {
        record = await requestJson<IWorkbookRecord>('GET');
    } catch {
        setMode('error', '無法連線到後端');
        return;
    }

    state.version = record.version;
    const workbook = univerAPI.createWorkbook(record.snapshot ?? createBlankWorkbook());
    watchDirty(workbook);
    showSavedInfo(record);

    const acquired = await tryLock();
    if (acquired) {
        enterEditMode(workbook);
    } else {
        enterReadonlyMode(workbook, record.lock);
    }
}

function enterEditMode(workbook: FWorkbook) {
    workbook.setEditable(true);
    setMode('edit', '編輯中（你持有編輯鎖）');
    updateSaveButton();

    elements.save.addEventListener('click', () => void save(workbook));
    window.addEventListener('keydown', (event) => {
        if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
            event.preventDefault();
            event.stopPropagation();
            void save(workbook);
        }
    }, true);

    const heartbeat = window.setInterval(async () => {
        if (!(await tryLock())) {
            // 例如電腦休眠超過 30 秒，鎖已過期並被別人拿走
            window.clearInterval(heartbeat);
            workbook.setEditable(false);
            setMode('readonly', '編輯鎖已失效，改為唯讀');
            setStatus(state.dirty ? '⚠️ 尚未存檔的修改無法再儲存，請先複製需要的內容後重新載入。' : '請重新載入以取得最新內容。');
            elements.reload.hidden = false;
            updateSaveButton();
        }
    }, HEARTBEAT_MS);

    window.addEventListener('beforeunload', (event) => {
        if (state.dirty) {
            event.preventDefault();
        }
    });
    window.addEventListener('pagehide', () => {
        navigator.sendBeacon(`${apiBase}/unlock`, new Blob([JSON.stringify({ clientId })], { type: 'application/json' }));
    });
}

function enterReadonlyMode(workbook: FWorkbook, lock: ILockInfo | null) {
    workbook.setEditable(false);
    setMode('readonly', lock ? `唯讀：${lock.userName} 正在編輯` : '唯讀');
    updateSaveButton();

    const poll = window.setInterval(async () => {
        try {
            const record = await requestJson<IWorkbookRecord>('GET');
            if (record.version !== state.version) {
                setStatus(`${record.savedBy ?? '其他人'} 已存了新版本（第 ${record.version} 版）。`);
                elements.reload.hidden = false;
            }
            if (!record.lock) {
                window.clearInterval(poll);
                setMode('readonly', '唯讀：目前沒有人在編輯');
                setStatus('可以重新載入以取得編輯權。');
                elements.reload.hidden = false;
            } else {
                setMode('readonly', `唯讀：${record.lock.userName} 正在編輯`);
            }
        } catch {
            // 暫時連不上就等下一輪
        }
    }, READONLY_POLL_MS);
}

async function save(workbook: FWorkbook) {
    if (state.mode !== 'edit' || state.saving) {
        return;
    }
    state.saving = true;
    updateSaveButton();
    setStatus('存檔中…');

    try {
        const response = await fetch(apiBase, {
            method: 'PUT',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ clientId, baseVersion: state.version, snapshot: workbook.save() }),
        });
        const result = await response.json();
        if (response.ok) {
            state.version = result.version;
            state.dirty = false;
            setStatus(`已存檔（第 ${result.version} 版，${formatTime(result.savedAt)}）`);
        } else if (response.status === 409) {
            setStatus(`⚠️ 伺服器上已有較新的第 ${result.version} 版，這次沒有存檔。請重新載入。`);
            elements.reload.hidden = false;
        } else if (response.status === 423) {
            setStatus('⚠️ 你已經沒有編輯鎖，這次沒有存檔。請重新載入。');
            elements.reload.hidden = false;
        } else {
            setStatus(`⚠️ 存檔失敗（${result.error ?? response.status}）`);
        }
    } catch {
        setStatus('⚠️ 存檔失敗：無法連線到後端，修改仍保留在畫面上，可稍後再存。');
    } finally {
        state.saving = false;
        updateSaveButton();
    }
}

/** 只要有會寫進快照的 Mutation 就標記為未存檔；公式計算等只在本機用的 Mutation 不算。 */
function watchDirty(workbook: FWorkbook) {
    workbook.onCommandExecuted((command) => {
        if (command.type !== CommandType.MUTATION || isLocalOnlyMutation(command.id)) {
            return;
        }
        if (!state.dirty) {
            state.dirty = true;
            if (state.mode === 'edit') {
                setStatus('有尚未存檔的修改');
            }
            updateSaveButton();
        }
    });
}

function isLocalOnlyMutation(id: string) {
    return id.startsWith('formula.mutation.') || id.includes('mark-dirty') || id === 'sheet.mutation.toggle-note-popup';
}

async function tryLock(): Promise<boolean> {
    try {
        const response = await fetch(`${apiBase}/lock`, {
            method: 'POST',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ clientId, userName: elements.userName.value }),
        });
        return response.ok;
    } catch {
        // 網路暫時中斷時先不放棄編輯權，等下次續約；後端過期後自然會失效
        return state.mode === 'edit';
    }
}

async function requestJson<T>(method: string): Promise<T> {
    const response = await fetch(apiBase, { method });
    if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
    }
    return response.json() as Promise<T>;
}

function setMode(mode: typeof state.mode, label: string) {
    state.mode = mode;
    elements.badge.dataset.mode = mode;
    elements.badge.textContent = label;
}

function setStatus(text: string) {
    elements.status.textContent = text;
}

function updateSaveButton() {
    elements.save.disabled = state.mode !== 'edit' || state.saving || !state.dirty;
}

function showSavedInfo(record: IWorkbookRecord) {
    if (record.version === 0) {
        setStatus('新的活頁簿，尚未存檔');
    } else {
        setStatus(`目前是第 ${record.version} 版，${record.savedBy ?? ''} 於 ${formatTime(record.savedAt)} 存檔`);
    }
}

function createBlankWorkbook(): Partial<IWorkbookData> {
    return {
        id: workbookId,
        name: workbookId,
        locale: LocaleType.ZH_TW,
        sheetOrder: ['sheet-1'],
        sheets: {
            'sheet-1': {
                id: 'sheet-1',
                name: '工作表1',
                rowCount: 1000,
                columnCount: 26,
                cellData: {
                    0: { 0: { v: '品項' }, 1: { v: '數量' }, 2: { v: '單價' }, 3: { v: '小計' } },
                    1: { 0: { v: '筆電' }, 1: { v: 2 }, 2: { v: 32000 }, 3: { f: '=B2*C2' } },
                    2: { 0: { v: '螢幕' }, 1: { v: 3 }, 2: { v: 6500 }, 3: { f: '=B3*C3' } },
                    3: { 0: { v: '合計' }, 3: { f: '=SUM(D2:D3)' } },
                },
            },
        },
    };
}

function getClientId() {
    // sessionStorage 以分頁為單位：同一分頁重新整理時沿用同一個 id，不會被自己的舊鎖擋住
    const key = 'save-lock:client-id';
    const existing = safeStorage(() => sessionStorage.getItem(key));
    if (existing) {
        return existing;
    }
    const id = crypto.randomUUID();
    safeStorage(() => sessionStorage.setItem(key, id));
    return id;
}

function getUserName() {
    return safeStorage(() => localStorage.getItem('save-lock:user-name')) || `使用者-${clientId.slice(0, 4)}`;
}

function safeStorage<T>(fn: () => T): T | null {
    try {
        return fn();
    } catch {
        return null;
    }
}

function formatTime(iso: string | null) {
    return iso ? new Date(iso).toLocaleString('zh-TW', { hour12: false }) : '';
}
