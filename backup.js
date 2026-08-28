// backup.js
// ─────────────────────────────────────────────────────────────────
// 設定／班表字典 備份與還原
//
// 目的：擴充功能以「載入未封裝項目」安裝時，chrome.storage.local 的資料
// 會綁定在當下的擴充功能 ID 上。若使用者更換擴充功能檔案的儲存位置（例如
// 每次更新都重新解壓縮到新資料夾）而 manifest.json 沒有固定的 "key"，
// Chrome 會配發全新的擴充功能 ID，導致 storage 資料「看似消失」。
// 換電腦、重灌系統時，storage 資料本來就不會跟著走。
//
// 本檔案提供「匯出備份檔 / 匯入備份檔」功能，讓使用者可以手動把設定與
// 班表字典打包成 JSON 檔下載，換位置/換電腦後再匯入還原。
// 與 manifest.json 的固定 "key" 屬於兩道互補的保險：
//   - 固定 key：只要 manifest.json 沒被覆蓋，ID 不會變，storage 資料自動延續，
//     使用者完全不用做任何操作。
//   - 本備份機制：即使 key 遺失／換了全新電腦，仍有備份檔可還原。
//
// 依賴：必須在 shared.js 之後載入（使用其 STORAGE_KEYS / createPopupWindow）。
// ─────────────────────────────────────────────────────────────────
(function () {
    // 涵蓋「設定」與「班表字典」全部相關 key，直接取自 shared.js 的 STORAGE_KEYS，
    // 避免字串寫死、日後跟 shared.js 改動的 key 名稱脫鉤。
    const BACKUP_KEYS = [
        STORAGE_KEYS.AUTO_MODE,
        STORAGE_KEYS.SHOW_WEB_PREVIEW,
        STORAGE_KEYS.SHOW_EXCEL_REPORT,
        STORAGE_KEYS.BLANK_FILL_MODE,
        STORAGE_KEYS.BLANK_FILL_CODE,
        STORAGE_KEYS.HR_SHIFTS,
        STORAGE_KEYS.SHIFT_DICT,
        STORAGE_KEYS.WW_MODE,
    ];
    const BACKUP_FILE_TAG = 'kmuh-shift-helper-backup';

    function setStatus(el, text, ok) {
        if (!el) return;
        el.style.color = ok ? '#27ae60' : '#e74c3c';
        el.textContent = text;
    }

    function exportBackup(statusEl) {
        chrome.storage.local.get(BACKUP_KEYS, (data) => {
            const payload = {
                _type: BACKUP_FILE_TAG,
                _version: 1,
                _exportedAt: new Date().toISOString(),
                data,
            };
            const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
            const url  = URL.createObjectURL(blob);
            const a    = document.createElement('a');
            a.href     = url;
            a.download = `kmuh-backup-${new Date().toISOString().slice(0, 10)}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            setStatus(statusEl, '✅ 備份檔已下載', true);
            createPopupWindow({
                message:  '✅ 備份檔已下載！\n請妥善保存此 JSON 檔，換電腦或重新安裝時可用來還原設定與班表字典。',
                btnColor: '#27ae60',
                width: 340, height: 200,
            });
        });
    }

    function importBackup(file, statusEl, onDone) {
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
            let parsed;
            try {
                parsed = JSON.parse(ev.target.result);
            } catch (err) {
                setStatus(statusEl, '❌ 檔案不是有效的 JSON', false);
                createPopupWindow({
                    title: '❌ 匯入失敗', message: '檔案不是有效的 JSON：\n' + err.message,
                    btnColor: '#e74c3c', width: 340, height: 200,
                });
                return;
            }
            // 相容兩種格式：本工具匯出的 { _type, data } 包裝格式，
            // 或直接是 { autoMode: ..., shiftDict: ... } 的純資料格式
            const payload = (parsed && parsed._type === BACKUP_FILE_TAG && parsed.data) ? parsed.data : parsed;
            const toRestore = {};
            BACKUP_KEYS.forEach((k) => {
                if (payload && payload[k] !== undefined) toRestore[k] = payload[k];
            });
            if (Object.keys(toRestore).length === 0) {
                setStatus(statusEl, '❌ 檔案內容無法辨識', false);
                createPopupWindow({
                    title: '❌ 匯入失敗', message: '檔案內容無法辨識，請確認是否為正確的備份檔。',
                    btnColor: '#e74c3c', width: 340, height: 200,
                });
                return;
            }
            chrome.storage.local.set(toRestore, () => {
                setStatus(statusEl, '✅ 已還原備份', true);
                createPopupWindow({
                    message:  '✅ 已還原備份！\n畫面即將重新整理。',
                    btnColor: '#27ae60',
                    width: 320, height: 180,
                });
                if (onDone) setTimeout(onDone, 600);
            });
        };
        reader.readAsText(file);
    }

    document.addEventListener('DOMContentLoaded', () => {
        const statusEl    = document.getElementById('backupStatus');
        const exportBtn   = document.getElementById('exportBackupBtn');
        const importBtn   = document.getElementById('importBackupBtn');
        const importInput = document.getElementById('importBackupFile');

        if (exportBtn) {
            exportBtn.onclick = () => exportBackup(statusEl);
        }
        if (importBtn && importInput) {
            importBtn.onclick = () => importInput.click();
            importInput.addEventListener('change', (e) => {
                const file = e.target.files[0];
                importBackup(file, statusEl, () => location.reload());
                e.target.value = '';
            });
        }
    });
})();
