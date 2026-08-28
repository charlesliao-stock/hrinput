document.addEventListener('DOMContentLoaded', async () => {
    if (typeof XLSX === 'undefined') {
        document.getElementById('status').textContent = '❌ xlsx 函式庫載入失敗，請確認 xlsx.full.min.js 存在';
        document.getElementById('step2Btn').disabled = true;
        return;
    }
    const statusDiv = document.getElementById('status'), excelFile = document.getElementById('excelFile');
    const googleSheetsUrl = document.getElementById('googleSheetsUrl');
    const googleSheetsReadBtn = document.getElementById('googleSheetsReadBtn');
    let currentWorkbook = null;
    let lastSelectedSheet = null;
    let lastNhDates = []; // 匯入前確認視窗中選定的國定假日日期（下個月「幾號」的整數陣列），供重新檢測時沿用

    // 本月班表記憶保存期限：30 分鐘。超過此時間視為過期，避免瀏覽器/分頁重開後
    // 使用者誤以為「已記憶」的仍是最新資料，實際上網頁班表可能早已被他人異動過。
    const MEMORY_TTL_MS = 30 * 60 * 1000;

    // 判斷 storage 中的 lastMonthData 目前狀態：'none' 尚未記憶／'fresh' 有效／'expired' 已逾時
    function getMemoryStatus(lastMonthData) {
        if (!lastMonthData?.yymm) return 'none';
        // 舊版資料沒有 savedAt 時間戳記，保守起見一律視為已逾時，強制使用者重新讀取一次
        if (!lastMonthData.savedAt) return 'expired';
        return (Date.now() - lastMonthData.savedAt < MEMORY_TTL_MS) ? 'fresh' : 'expired';
    }

    // ── 步驟 1 按鈕文字更新（開啟時 + 記憶成功後共用） ──────────────
    function updateStep1BtnLabel(yymm, expired = false) {
        const btn = document.getElementById('step1Btn');
        if (expired) {
            btn.textContent = '⚠️ 記憶已逾時，需重新讀取';
            btn.style.background = '#e67e22';
        } else if (yymm && yymm.length === 6) {
            const y = yymm.substring(0, 4);
            const m = yymm.substring(4, 6);
            btn.textContent = `✅ 已記憶 ${y}/${m}（點擊重新讀取）`;
            btn.style.background = '#27ae60';
        } else {
            btn.textContent = '💾 記憶本月班表並跳轉至次月';
            btn.style.background = '';
        }
    }

    // 頁面開啟時，若 storage 已有記憶資料則立即反映（並檢查是否已逾時）
    chrome.storage.local.get('lastMonthData', (d) => {
        const status = getMemoryStatus(d.lastMonthData);
        if (status === 'fresh')   updateStep1BtnLabel(d.lastMonthData.yymm);
        if (status === 'expired') updateStep1BtnLabel(null, true);
    });

    // --- 更新提醒邏輯 ---
    const updateAlert = document.getElementById('updateAlert');
    const updateVersion = document.getElementById('updateVersion');
    const downloadUpdateBtn = document.getElementById('downloadUpdateBtn');

    chrome.storage.local.get(['updateAvailable', 'latestVersion', 'downloadUrl'], (data) => {
        if (data.updateAvailable) {
            updateAlert.style.display = 'block';
            updateVersion.textContent = `最新版本：v${data.latestVersion}`;
            downloadUpdateBtn.onclick = () => {
                // 直接下載 ZIP 檔
                const downloadUrl = data.downloadUrl || 'https://github.com/charlesliao-stock/hrinput/archive/refs/heads/main.zip';
                chrome.tabs.create({ url: downloadUrl });
                
                // 提示使用者後續步驟
                statusDiv.innerHTML = "<b>📥 已開始下載更新檔！</b><br>請解壓縮後，到 Chrome 擴充功能頁面點擊「重新載入」即可完成更新。";
            };
        }
    });

    // 手動觸發檢查更新 (點擊標題時)
    document.querySelector('h2').onclick = () => {
        statusDiv.textContent = "⏳ 正在檢查更新...";
        chrome.runtime.sendMessage({ action: "manualCheckUpdate" }, (res) => {
            setTimeout(() => {
                chrome.storage.local.get(['updateAvailable'], (d) => {
                    statusDiv.textContent = d.updateAvailable ? "🚀 發現新版本！" : "✅ 目前已是最新版本";
                });
            }, 1000);
        });
    };

    document.getElementById('openQuickSettings').onclick = () => chrome.windows.create({ url: 'quick_settings.html', type: 'popup', width: 700, height: 300 });
    document.getElementById('openDictManager').onclick   = () => chrome.windows.create({ url: 'dict_manager.html',   type: 'popup', width: 780, height: 500 });

    // Excel 匯入結果已交由網頁端顯示後，關閉目前的擴充功能 popup，
    // 避免右上角選單遮住檢測報告與下一步操作；若瀏覽器不允許 script 關閉，
    // 以簡短提示取代整個選單內容作為 fallback。
    function hidePopupMenuAfterImport(message = 'Excel 已匯入，請查看頁面上的檢測報告。') {
        setTimeout(() => {
            try { window.close(); } catch (err) { /* 某些瀏覽器不允許關閉 popup */ }
            if (!window.closed) {
                document.body.replaceChildren();
                document.body.className = 'import-complete';
                const notice = document.createElement('div');
                notice.className = 'import-complete-note';
                notice.textContent = message;
                document.body.appendChild(notice);
            }
        }, 80);
    }

    // 只有需要選取國定假日時才放大整個 popup，讓左側訊息與右側日期區都有足夠操作空間。
    // 未進入國定假日流程時維持原本寬度與高度。
    function getNhPopupHeight(targetYymm) {
        const year = parseInt(String(targetYymm || '').slice(0, 4), 10);
        const month = parseInt(String(targetYymm || '').slice(4, 6), 10);
        if (!year || !month) return 760;
        const daysInMonth = new Date(year, month, 0).getDate();
        const firstDow = new Date(year, month - 1, 1).getDay();
        const calendarRows = Math.ceil((firstDow + daysInMonth) / 7);
        // 5／6 列月份均在約 600px 的固定 popup 高度內完整呈現；
        // 不再設定 720px 以上高度，避免 Chrome popup 外層出現垂直拉 bar。
        return Math.max(600, Math.min(640, 348 + calendarRows * 42));
    }

    function setPopupNhExpanded(expanded, targetYymm = '') {
        const root = document.documentElement;
        if (expanded) {
            root.style.setProperty('--popup-nh-height', `${getNhPopupHeight(targetYymm)}px`);
        } else {
            root.style.removeProperty('--popup-nh-height');
        }
        root.classList.toggle('popup-nh-expanded', expanded);
        if (document.body) document.body.classList.toggle('popup-nh-expanded', expanded);
    }

    // 匯入前確認視窗的尺寸由模式明確控制，不再由目前 popup viewport 自動壓縮。
    function setPreImportWindowMode(mode, targetYymm = '') {
        const root = document.documentElement;
        const compact = mode === 'compact';
        const expanded = mode === 'nh';
        root.classList.toggle('popup-preimport-compact', compact);
        if (document.body) document.body.classList.toggle('popup-preimport-compact', compact);
        setPopupNhExpanded(expanded, targetYymm);

    }

    function showAlertWindow(message) {
        const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
        <style>
            body { width:300px; height:150px; margin:0; display:flex; flex-direction:column;
                   align-items:center; justify-content:center; gap:16px; padding:0 20px;
                   box-sizing:border-box; font-family:"Microsoft JhengHei",sans-serif; background:#fff; overflow:hidden; }
            .msg { font-size:15px; color:#c0392b; font-weight:bold; text-align:center; }
            button { width:100%; padding:10px; background:#e74c3c; color:white; border:none;
                     border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer; }
            button:hover { background:#c0392b; }
        </style></head><body>
        <div class="msg">${message}</div>
        <button onclick="window.close()">確定</button>
        </body></html>`;
        chrome.windows.create({
            url: 'data:text/html;charset=utf-8,' + encodeURIComponent(html),
            type: 'popup', width: 320, height: 170, focused: true
        });
    }

    // ── 紅色閃動邊框的自訂確認對話框（取代原生 confirm，因原生對話框無法自訂樣式，
    //    外框也不夠明顯）。在 popup 自己的 DOM 內顯示，不會有失焦關閉的問題。
    function showConfirmDialog(message) {
        if (!document.getElementById('kmuhFlashStyle')) {
            const style = document.createElement('style');
            style.id = 'kmuhFlashStyle';
            style.textContent = `
                @keyframes kmuhFlashBorder {
                    0%, 100% { border-color:#e74c3c; box-shadow:0 0 10px 3px rgba(231,76,60,0.7); }
                    50%      { border-color:#ffb3ab; box-shadow:0 0 2px 0 rgba(231,76,60,0.15); }
                }
            `;
            document.head.appendChild(style);
        }
        return new Promise((resolve) => {
            const overlay = document.createElement('div');
            overlay.style.cssText = `
                position:fixed; inset:0; background:rgba(0,0,0,0.45);
                display:flex; align-items:center; justify-content:center; z-index:99999;
            `;
            const box = document.createElement('div');
            box.style.cssText = `
                background:#fff; border-radius:8px; padding:20px 22px; width:250px;
                box-sizing:border-box; border:3px solid #e74c3c;
                animation: kmuhFlashBorder 0.9s infinite;
                font-family:"Microsoft JhengHei",sans-serif;
            `;
            box.innerHTML = `
                <div style="white-space:pre-line; font-size:13px; color:#333; line-height:1.7; margin-bottom:16px;">${message}</div>
                <div style="display:flex; gap:8px; justify-content:flex-end;">
                    <button id="kmuhConfirmCancel" style="padding:6px 16px; border:none; border-radius:5px; background:#bdc3c7; color:#333; font-weight:bold; cursor:pointer;">取消</button>
                    <button id="kmuhConfirmOk" style="padding:6px 16px; border:none; border-radius:5px; background:#3498db; color:#fff; font-weight:bold; cursor:pointer;">確定</button>
                </div>
            `;
            overlay.appendChild(box);
            document.body.appendChild(overlay);
            const cleanup = (result) => { overlay.remove(); resolve(result); };
            box.querySelector('#kmuhConfirmOk').onclick     = () => cleanup(true);
            box.querySelector('#kmuhConfirmCancel').onclick = () => cleanup(false);
        });
    }

    async function sendMessage(msg) {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab) return { success: false, message: "❌ 找不到分頁" };

        const url = (tab.url || "").toLowerCase();
        if (!url.includes("kmuhdeptshiftedit.aspx")) {
            showAlertWindow("❌ 請先開啟 排班編輯畫面");
            return { success: false, message: "❌ 請先開啟 排班編輯畫面" };
        }

        try {
            return await chrome.tabs.sendMessage(tab.id, msg);
        } catch (e) {
            try {
                await chrome.scripting.executeScript({ target: { tabId: tab.id }, files: ['content.js'] });
                await new Promise(r => setTimeout(r, 500));
                return await chrome.tabs.sendMessage(tab.id, msg);
            } catch (e2) {
                return { success: false, message: "❌ 無法連線至頁面，請手動重整後再試" };
            }
        }
    }

    // --- 步驟 1：讀取舊月並記憶 ---
    document.getElementById('step1Btn').onclick = async () => {
        statusDiv.textContent = "⏳ 正在記憶本月班表...";
        const set = await chrome.storage.local.get(['showWebPreview', 'autoMode']);
        const baseMsg = {
            action: "readAndMemorize",
            showPreview: set.showWebPreview === true,  // 預設不顯示，需明確設為 true 才顯示
            autoMode: set.autoMode || false
        };
        let res = await sendMessage(baseMsg);

        // 網頁月份與系統當前月份不一致：在 popup 自己的視窗內跳出確認（而非在網頁分頁上跳
        // confirm，否則分頁搶走焦點會讓這個 popup 被瀏覽器直接關閉，之後按確定也沒反應、
        // 更不會自動跳轉到下個月），使用者確認後再帶 forceProceed 重新呼叫一次。
        if (res?.monthMismatch) {
            const proceed = await showConfirmDialog(
                `⚠️ 月份提醒\n\n網頁顯示月份：${res.pageYymm}\n系統當前月份：${res.sysYymm}\n\n兩者不一致，是否仍要繼續記憶？`
            );
            if (!proceed) {
                statusDiv.textContent = "❌ 已取消記憶";
                return;
            }
            statusDiv.textContent = "⏳ 正在記憶本月班表...";
            res = await sendMessage({ ...baseMsg, forceProceed: true });
        }

        if (res?.success) {
            updateStep1BtnLabel(res.yymm);
            let msg = `✅ 記憶完成 (${res.yymm})`;
            if (res.targetPeriod) {
                msg += `\n📅 檢測週期：【${res.targetPeriod.label}】${res.targetPeriod.start}～${res.targetPeriod.end}`;
            } else if (res.periods && res.periods.length === 0) {
                msg += `\n⚠️ 未偵測到四週變形週期，請確認頁面`;
            }

            if (res.nextUrl) {
                // 有下個月 URL：直接跳轉，不管全自動模式
                statusDiv.textContent = `${msg}\n⚡ 即將跳轉至次月...`;
                setTimeout(() => chrome.tabs.update({ url: res.nextUrl }), 800);
            } else {
                statusDiv.textContent = msg;
            }
        } else {
            statusDiv.textContent = res?.message || "❌ 記憶失敗，請確認頁面正確";
        }
    };

    // --- 步驟 2：選擇 Excel 或讀取 Google Sheets ---
    async function ensureFreshMonthMemory() {
        const d = await chrome.storage.local.get('lastMonthData');
        const status = getMemoryStatus(d.lastMonthData);
        if (status === 'fresh') return true;
        updateStep1BtnLabel(null, status === 'expired');
        statusDiv.textContent = status === 'expired'
            ? "⚠️ 本月班表記憶已逾時（超過30分鐘），請重新執行步驟1"
            : "❌ 尚未記憶本月班表，請先執行步驟1";
        return false;
    }

    async function prepareWorkbook(workbook, sourceLabel) {
        currentWorkbook = workbook;
        lastSelectedSheet = null;
        document.getElementById('sheetSelectBox').style.display = 'none';
        const sheetNames = currentWorkbook.SheetNames || [];
        if (sheetNames.length === 0) {
            statusDiv.textContent = `❌ ${sourceLabel}中沒有任何工作表`;
            return;
        }
        if (sheetNames.length === 1) {
            statusDiv.textContent = `偵測到唯一工作表「${sheetNames[0]}」，自動匯入中...`;
            lastSelectedSheet = sheetNames[0];
            await processExcelSheet(sheetNames[0]);
            return;
        }
        const sel = document.getElementById('sheetSelect');
        sel.innerHTML = sheetNames.map((name, i) =>
            `<option value="${String(name).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')}">${i + 1}. ${String(name).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</option>`
        ).join('');
        document.getElementById('sheetSelectBox').style.display = 'block';
        statusDiv.textContent = `📋 ${sourceLabel}偵測到 ${sheetNames.length} 個工作表，請選擇後按確認`;
    }

    document.getElementById('step2Btn').onclick = async () => {
        if (!(await ensureFreshMonthMemory())) return;
        excelFile.click();
    };

    excelFile.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;
        statusDiv.textContent = "⏳ 讀取 Excel 檔案中...";
        const reader = new FileReader();
        reader.onload = async (ev) => {
            try {
                await prepareWorkbook(XLSX.read(new Uint8Array(ev.target.result), { type: 'array' }), 'Excel ');
            } catch (err) {
                console.error('[Excel 讀取錯誤]', err);
                statusDiv.textContent = "❌ Excel 讀取失敗：" + err.message;
            }
        };
        reader.readAsArrayBuffer(file);
        e.target.value = "";
    });

    function parseGoogleSheetsUrl(rawValue) {
        const raw = String(rawValue || '').trim();
        if (!raw) throw new Error('請先貼上 Google Sheets 網址');
        let url;
        try { url = new URL(raw); } catch (err) { throw new Error('網址格式不正確，請貼上完整的 Google Sheets 網址'); }
        const host = url.hostname.toLowerCase();
        if (!['docs.google.com', 'spreadsheets.google.com'].includes(host)) {
            throw new Error('網址不是 Google Sheets 網址，請確認是 docs.google.com/spreadsheets/...');
        }
        const match = url.pathname.match(/\/spreadsheets\/(?:u\/\d+\/)?d\/([a-zA-Z0-9_-]+)/);
        if (!match) throw new Error('無法從網址辨識試算表 ID，請貼上試算表的分享網址');
        // 刻意忽略網址上的 gid：匯入時必須下載整本試算表，才能讓使用者選擇所有 Sheet。
        return { id: match[1] };
    }

    async function readGoogleSheetsWorkbook() {
        if (!(await ensureFreshMonthMemory())) return;
        const { id } = parseGoogleSheetsUrl(googleSheetsUrl?.value);
        // 不帶 gid，否則 Google 可能只回傳網址目前所在的單一工作表。
        const exportUrl = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(id)}/export?format=xlsx`;
        googleSheetsReadBtn.disabled = true;
        statusDiv.textContent = '⏳ 正在讀取 Google Sheets 整本試算表，請稍候...';
        try {
            // 公開／已開放連結存取的試算表應匿名讀取；帶 Cookie 會讓 Google 匯出端點的跨來源請求觸發 Failed to fetch。
            const response = await fetch(exportUrl, { redirect: 'follow', credentials: 'omit', cache: 'no-store' });
            if (!response.ok) throw new Error(`伺服器回應 ${response.status}`);
            const buffer = await response.arrayBuffer();
            if (buffer.byteLength < 100) throw new Error('回傳內容不是有效的試算表檔案');
            const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
            await prepareWorkbook(workbook, 'Google Sheets ');
        } catch (err) {
            console.error('[Google Sheets 讀取錯誤]', err);
            statusDiv.textContent = `❌ Google Sheets 讀取失敗：${err.message}。請確認試算表已開放連結存取（可檢視或可編輯），且網址正確。`;
        } finally {
            googleSheetsReadBtn.disabled = false;
        }
    }

    googleSheetsReadBtn?.addEventListener('click', readGoogleSheetsWorkbook);
    googleSheetsUrl?.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') readGoogleSheetsWorkbook();
    });

    document.getElementById('sheetConfirmBtn').onclick = async () => {
        const selectedSheet = document.getElementById('sheetSelect').value;
        if (!selectedSheet || !currentWorkbook) return;
        document.getElementById('sheetSelectBox').style.display = 'none';
        lastSelectedSheet = selectedSheet;
        await processExcelSheet(selectedSheet);
    };

    // ── 共用 helper：欄位 index → Excel 欄名（0→A, 1→B, ... 25→Z, 26→AA...） ──
    function colIdxToLetter(idx) {
        if (idx < 0) return "?";
        let n = idx + 1, s = "";
        while (n > 0) {
            const rem = (n - 1) % 26;
            s = String.fromCharCode(65 + rem) + s;
            n = Math.floor((n - 1) / 26);
        }
        return s;
    }

    // ── 共用 helper：職編有效性檢核（與 content.js 保持相同邏輯） ────
    // 純數字、若第一碼為 0 則去掉該碼，剩餘長度需為 6 或 7 碼。
    function isValidEmpId(id) {
        const s = String(id || "").trim();
        if (!/^\d+$/.test(s)) return false;
        const stripped = s[0] === '0' ? s.slice(1) : s;
        return stripped.length === 6 || stripped.length === 7;
    }

    // ── 共用 helper：依偵測到的欄位配置，統計預計匯入的員工數（去重） ──
    function countExpectedEmployees(data, layout) {
        const { empIdColIdx } = layout;
        if (empIdColIdx === -1 || empIdColIdx === undefined) return 0;
        const seen = new Set();
        for (const row of data) {
            if (!row) continue;
            const raw = String(row[empIdColIdx] || "").trim();
            if (isValidEmpId(raw)) {
                const stripped = raw[0] === '0' ? raw.slice(1) : raw;
                seen.add(stripped);
            }
        }
        return seen.size;
    }

    // ── 共用：把一批「職編/姓名」名單轉成顯示用的字串 ───────────────
    function formatEmpList(list) {
        return (list || []).map(w => `${w.empId}${w.name ? '（' + w.name + '）' : ''}`).join('、');
    }

    // ── 整合式「匯入前確認」視窗 ───────────────────────────────────
    // 將以下 5 類訊息整合成單一彈窗顯示：
    //   1. 職編/姓名/班表開始(1號) 欄位偵測結果（恆顯示）
    //   2. 日期行有重複（如有）→ 強制中止，僅顯示「確定」，不可繼續
    //   3. 人員本月有、下個月無（如有）→ 併入確認/取消
    //   4. 人員本月無、下個月有（如有）→ 併入確認/取消
    //   5. 班表天數不足（如有）→ 提供「系統補足天數後繼續」或「取消」
    //
    // 回傳值：Promise<'continue' | 'fill' | 'cancel'>
    //   - 'continue'：可直接送往 content.js 處理（無天數不足問題）
    //   - 'fill'    ：使用者選擇由系統補足缺少的天數後再送往 content.js 處理
    //   - 'cancel'  ：使用者取消，或日期行重複強制中止
    // ── 國定假日日曆 HTML 產生（下個月，全員統一日期） ─────────────────
    function buildNhCalendarHtml(targetYymm, nhRequired) {
        const year  = parseInt(targetYymm.slice(0, 4), 10);
        const month = parseInt(targetYymm.slice(4, 6), 10);
        const daysInMonth = new Date(year, month, 0).getDate();
        const firstDow    = new Date(year, month - 1, 1).getDay(); // 0=日
        const weekdayLabels = ['日', '一', '二', '三', '四', '五', '六'];

        const headerCells = weekdayLabels.map(w =>
            `<div style="text-align:center;font-size:11px;color:#888;font-weight:bold;">${w}</div>`).join('');

        let cells = '';
        for (let i = 0; i < firstDow; i++) cells += `<div></div>`;
        for (let d = 1; d <= daysInMonth; d++) {
            const dow = (firstDow + d - 1) % 7;
            const isWeekend = dow === 0 || dow === 6;
            cells += `
                <div class="_nhDayCell" data-day="${d}" data-weekend="${isWeekend}" style="
                    padding:6px 0; text-align:center; border-radius:5px; cursor:pointer;
                    font-size:12px; border:1px solid #ddd; user-select:none; background:#fff;
                    color:${isWeekend ? '#e74c3c' : '#2c3e50'};
                ">${d}</div>`;
        }

        return `
            <div class="_nhPanel" style="background:#fff8f0;border:1px solid #f0c060;border-radius:6px;
                        padding:10px 12px;margin-bottom:10px;font-size:12px;color:#7d4500;">
                <div style="font-weight:bold;margin-bottom:6px;">
                    📅 國定假日日期指定（${year}年${month}月，本月應排 <span style="color:#c0392b">${nhRequired}</span> 天）
                </div>
                <div style="font-size:11px;color:#888;margin-bottom:8px;line-height:1.6">
                    請勾選對應數量的確切日期，全體人員將統一套用：當天若為上班班別將加註 <b>N+</b>，
                    當天若為「代表放假」符號將轉為 <b>NH</b>。數量須與應排天數完全相符才能繼續匯入。
                </div>
                <div style="display:grid;grid-template-columns:repeat(7,1fr);gap:4px;margin-bottom:6px;">
                    ${headerCells}${cells}
                </div>
                <div id="_nhCountHint" style="font-size:12px;font-weight:bold;color:#c0392b;">已選 0／${nhRequired} 天</div>
            </div>`;
    }

    function showConsolidatedPreImportWindow(sheetName, report) {
        return new Promise((resolve) => {
            const { layout, empCount, structuralIssues, multiDateWarning, daysWarning, departedWarnings, noOldDataWarnings, nhRequired, targetYymm } = report;
            const blocking = !!multiDateWarning || (structuralIssues && structuralIssues.length > 0); // 項目2、欄位偵測失敗：強制中止
            const needsNh  = !blocking && nhRequired > 0 && !!targetYymm;
            const selectedNhDays = new Set();

            // 兩種確認狀態都使用明確固定尺寸：無國假為可閱讀的小視窗；
            // 有國假則展開成完整雙欄日曆視窗。
            setPreImportWindowMode(needsNh ? 'nh' : 'compact', targetYymm);

            const empLetter  = layout.empIdColIdx !== -1 ? colIdxToLetter(layout.empIdColIdx) : "未偵測到";
            const nameLetter = layout.nameColIdx  !== -1 ? colIdxToLetter(layout.nameColIdx)  : "未偵測到";
            const dayLetter  = layout.day1ColIdx  !== -1 ? colIdxToLetter(layout.day1ColIdx)  : "未偵測到";


            const overlay = document.createElement('div');
            overlay.style.cssText = [
                // 確認視窗本身採文件流中的固定尺寸，避免 position:fixed 只取得
                // 原始 popup viewport，導致無國假內容被擠成窄直欄或日曆被裁切。
                'position:absolute', 'top:0', 'left:0',
                needsNh ? 'width:800px' : 'width:520px',
                needsNh ? 'height:var(--popup-nh-height,760px)' : 'height:500px',
                'background:rgba(0,0,0,0.45)',
                'display:flex', 'align-items:center', 'justify-content:center',
                'z-index:9999', 'padding:10px', 'box-sizing:border-box',
            ].join(';');

            // ── box 改為「上：標題（固定）／中：警示訊息（可捲動）／下：按鈕（固定）」──
            // 國定假日流程使用 760～900px 的固定可視高度；一般流程仍依 viewport
            // 彈性縮放。這樣日曆欄不會因原 popup viewport 太窄而被壓縮。
            const box = document.createElement('div');
            box.style.cssText = [
                'background:#fff', 'border-radius:8px',
                needsNh ? 'width:780px' : 'width:500px',
                needsNh ? 'min-width:640px' : 'min-width:480px',
                needsNh ? 'max-width:780px' : 'max-width:500px',
                needsNh ? 'max-height:calc(var(--popup-nh-height, 760px) - 20px)' : 'max-height:480px',
                'display:flex', 'flex-direction:column',
                'box-shadow:0 4px 20px rgba(0,0,0,0.25)',
                'font-family:"Microsoft JhengHei",sans-serif',
                'overflow:hidden',
            ].filter(Boolean).join(';');

            let sections = '';
            let nhSectionHtml = '';

            // ── 1. 欄位偵測結果（恆顯示） ──────────────────────────
            sections += `
                <div style="background:#f1f8ff;border:1px solid #bcd9f7;border-radius:6px;
                            padding:10px 12px;margin-bottom:10px;font-size:12px;color:#2c3e50;line-height:1.9">
                    📋 已偵測：職編 = <b style="color:#2980b9">${empLetter}</b> 欄，
                    姓名 = <b style="color:#2980b9">${nameLetter}</b> 欄，
                    日期開始(1號) = <b style="color:#2980b9">${dayLetter}</b> 欄<br>
                    預計匯入員工數：<b style="color:#c0392b;font-size:14px">${empCount}</b> 人
                </div>`;

            // ── 欄位偵測失敗（如找不到職編欄、找不到1號欄）：強制中止 ──────
            if (structuralIssues && structuralIssues.length > 0) {
                const issuesHtml = structuralIssues.map(({ icon, title, detail }) => `
                    <div style="margin-bottom:${structuralIssues.length > 1 ? '8px' : '0'}">
                        <div style="font-weight:bold">${icon} ${title}</div>
                        <div style="margin-top:2px">${detail.replace(/\n/g, '<br>')}</div>
                    </div>`).join('');
                sections += `
                <div style="background:#fdecea;border:1px solid #f5b4ac;border-radius:6px;
                            padding:10px 12px;margin-bottom:10px;font-size:12px;color:#c0392b;line-height:1.8">
                    ⛔ <b>欄位偵測失敗</b><br>
                    ${issuesHtml}
                </div>`;
            }

            // ── 2. 日期行重複（強制中止） ──────────────────────────
            if (multiDateWarning) {
                const rowsText = multiDateWarning.rows.join('、');
                sections += `
                <div style="background:#fdecea;border:1px solid #f5b4ac;border-radius:6px;
                            padding:10px 12px;margin-bottom:10px;font-size:12px;color:#c0392b;line-height:1.8">
                    ⛔ <b>日期行有重複</b><br>
                    在第 <b>${rowsText}</b> 列，各偵測到一段疑似「1、2…」連續日期欄位，
                    共 <b>${multiDateWarning.count}</b> 段，系統無法判斷哪一段才是正確的日期。<br>
                    請修正 Excel 檔案（僅保留一段正確的日期列）後，重新匯入。
                </div>`;
            }

            // ── 3. 人員本月有、下個月無 ─────────────────────────────
            if (departedWarnings && departedWarnings.length > 0) {
                sections += `
                <div style="background:#fdf0e0;border:1px solid #f5c88a;border-radius:6px;
                            padding:10px 12px;margin-bottom:10px;font-size:12px;color:#7d4500;line-height:1.8">
                    ⚠️ <b>人員本月有、下個月無</b>（共 ${departedWarnings.length} 人，可能離職或調離單位）<br>
                    ${formatEmpList(departedWarnings)}
                </div>`;
            }

            // ── 4. 人員本月無、下個月有 ─────────────────────────────
            if (noOldDataWarnings && noOldDataWarnings.length > 0) {
                sections += `
                <div style="background:#fdf0e0;border:1px solid #f5c88a;border-radius:6px;
                            padding:10px 12px;margin-bottom:10px;font-size:12px;color:#7d4500;line-height:1.8">
                    ⚠️ <b>人員本月無、下個月有</b>（共 ${noOldDataWarnings.length} 人，可能新調入或找不到本月資料）<br>
                    ${formatEmpList(noOldDataWarnings)}
                </div>`;
            }

            // ── 5. 班表天數不足 ─────────────────────────────────────
            if (daysWarning) {
                const { tYear, tMonth, expectedDays, consecutiveDays } = daysWarning;
                const missing = expectedDays - consecutiveDays;
                sections += `
                <div style="background:#fffbf0;border:1px solid #f0c060;border-radius:6px;
                            padding:10px 12px;margin-bottom:10px;font-size:12px;color:#555;line-height:1.8">
                    ⚠️ <b>班表天數不足</b><br>
                    目標月份 <b style="color:#c0392b">${tYear} 年 ${tMonth} 月</b> 應有
                    <b style="color:#c0392b">${expectedDays} 天</b>，但日期列中只偵測到連續
                    <b style="color:#c0392b">${consecutiveDays} 天</b>（缺少 <b style="color:#c0392b">${missing} 天</b>）。<br>
                    選擇「系統補足天數後繼續」：將自動在日期列末端補上第
                    ${consecutiveDays + 1}～${expectedDays} 天，班表資料欄留空。
                </div>`;
            }

            // ── 6. 國定假日日期指定日曆（下個月，全員統一日期） ──────────────
            // 日期選擇區獨立放到右側，避免與左側警示訊息一起堆疊成過長頁面。
            if (needsNh) {
                nhSectionHtml = buildNhCalendarHtml(targetYymm, nhRequired);
            }

            let buttonsHtml = '';
            if (blocking) {
                buttonsHtml = `
                <div style="font-size:11px;color:#888;margin-bottom:10px;line-height:1.7">
                    此問題須修正 Excel 檔案後才能繼續匯入。
                </div>
                <button id="_preImportBtnAck" style="width:100%;padding:9px;
                    background:#e74c3c;color:#fff;border:none;border-radius:6px;
                    font-size:12px;font-weight:bold;cursor:pointer">
                    ✖ 確定，我將修正 Excel 後重新匯入
                </button>`;
            } else if (daysWarning) {
                buttonsHtml = `
                <div style="font-size:11px;color:#888;margin-bottom:10px;line-height:1.7">
                    請確認以上偵測結果與提醒事項；若確認無誤可補足天數後繼續。
                </div>
                <button id="_preImportBtnFill" ${needsNh ? 'disabled' : ''} style="width:100%;padding:9px;margin-bottom:7px;
                    background:#27ae60;color:#fff;border:none;border-radius:6px;
                    font-size:12px;font-weight:bold;cursor:pointer;${needsNh ? 'opacity:.5;cursor:not-allowed;' : ''}">
                    📅 系統補足天數欄位（班表資料空白）後繼續
                </button>
                <button id="_preImportBtnCancel" style="width:100%;padding:9px;
                    background:#bdc3c7;color:#2c3e50;border:none;border-radius:6px;
                    font-size:12px;font-weight:bold;cursor:pointer">
                    ✖ 取消，重新確認檔案
                </button>`;
            } else {
                buttonsHtml = `
                <div style="font-size:11px;color:#888;margin-bottom:10px;line-height:1.7">
                    請確認以上偵測結果與提醒事項是否符合預期。
                </div>
                <button id="_preImportBtnContinue" ${needsNh ? 'disabled' : ''} style="width:100%;padding:9px;margin-bottom:7px;
                    background:#27ae60;color:#fff;border:none;border-radius:6px;
                    font-size:12px;font-weight:bold;cursor:pointer;${needsNh ? 'opacity:.5;cursor:not-allowed;' : ''}">
                    ✅ 確認無誤，繼續匯入
                </button>
                <button id="_preImportBtnCancel" style="width:100%;padding:9px;
                    background:#bdc3c7;color:#2c3e50;border:none;border-radius:6px;
                    font-size:12px;font-weight:bold;cursor:pointer">
                    ✖ 取消，重新確認檔案
                </button>`;
            }

            box.innerHTML = `
                <style>
                    .preimport-content {
                        display:grid;
                        grid-template-columns:minmax(0,1fr) minmax(250px,320px);
                        gap:12px;
                        padding:0 18px;
                        overflow:hidden;
                        flex:1 1 auto;
                        min-height:0;
                        align-items:start;
                    }
                    .preimport-main { min-width:0; }
                    .preimport-nh {
                        min-width:0;
                        position:sticky;
                        top:0;
                        align-self:start;
                    }
                    .preimport-nh ._nhPanel { margin-bottom:0 !important; }
                    @media (max-width:620px) {
                        .preimport-content {
                            grid-template-columns:minmax(0,1fr) minmax(215px,42%);
                            gap:8px;
                            padding:0 10px;
                        }
                        .preimport-content ._nhPanel { padding:8px; }
                        .preimport-content ._nhPanel > div:first-child { font-size:11px !important; }
                        .preimport-content ._nhPanel > div:nth-child(2) { font-size:10px !important; }
                        .preimport-content ._nhDayCell { padding:5px 0 !important; font-size:11px !important; }
                    }
                    @media (max-width:420px) {
                        .preimport-content {
                            grid-template-columns:minmax(0,1fr) minmax(190px,58%);
                            gap:6px;
                            padding:0 8px;
                        }
                        .preimport-nh { position:sticky; top:0; }
                        .preimport-content ._nhDayCell { padding:4px 0 !important; font-size:10px !important; }
                    }
                    /* 無國定假日時只有左側訊息，必須恢復單欄，避免內容只佔一半寬度。 */
                    .preimport-content-compact {
                        grid-template-columns:minmax(0,1fr) !important;
                    }
                </style>
                <div style="padding:16px 18px 0;flex:0 0 auto;">
                    <div style="font-size:15px;font-weight:bold;color:${blocking ? '#c0392b' : '#2c3e50'};margin-bottom:4px">
                        ${blocking ? '⛔ 匯入前檢查未通過' : '📋 匯入前確認'}
                    </div>
                    <div style="font-size:11px;color:#999;margin-bottom:10px">工作表：${sheetName}</div>
                </div>
                <div class="preimport-content${needsNh ? '' : ' preimport-content-compact'}">
                    <div class="preimport-main">${sections}</div>
                    ${nhSectionHtml ? `<div class="preimport-nh">${nhSectionHtml}</div>` : ''}
                </div>
                <div style="padding:10px 18px 14px;flex:0 0 auto;border-top:1px solid #eee;margin-top:4px;">
                    ${buttonsHtml}
                </div>`;

            overlay.appendChild(box);
            document.body.appendChild(overlay);

            function cleanup(action) {
                document.body.removeChild(overlay);
                setPreImportWindowMode('none');
                resolve({ action, nhDates: Array.from(selectedNhDays).sort((a, b) => a - b) });
            }

            // ── 國定假日日曆：勾選/取消 + 數量檢核，數量須「完全等於」nhRequired 才放行 ──
            if (needsNh) {
                const countHint = box.querySelector('#_nhCountHint');
                const fillBtnEl     = box.querySelector('#_preImportBtnFill');
                const continueBtnEl = box.querySelector('#_preImportBtnContinue');
                const updateNhUi = () => {
                    const ok = selectedNhDays.size === nhRequired;
                    countHint.textContent = `已選 ${selectedNhDays.size}／${nhRequired} 天`;
                    countHint.style.color = ok ? '#27ae60' : '#c0392b';
                    [fillBtnEl, continueBtnEl].forEach(btn => {
                        if (!btn) return;
                        btn.disabled = !ok;
                        btn.style.opacity      = ok ? '1' : '.5';
                        btn.style.cursor       = ok ? 'pointer' : 'not-allowed';
                    });
                };
                box.querySelectorAll('._nhDayCell').forEach(cell => {
                    cell.onclick = () => {
                        const day = parseInt(cell.dataset.day, 10);
                        if (selectedNhDays.has(day)) {
                            selectedNhDays.delete(day);
                            cell.style.background = '#fff';
                            cell.style.borderColor = '#ddd';
                            cell.style.fontWeight = 'normal';
                            cell.style.color = cell.dataset.weekend === 'true' ? '#e74c3c' : '#2c3e50';
                        } else {
                            if (selectedNhDays.size >= nhRequired) return; // 已達應選天數，不可再多選
                            selectedNhDays.add(day);
                            cell.style.background = '#27ae60';
                            cell.style.borderColor = '#219150';
                            cell.style.color = '#fff';
                            cell.style.fontWeight = 'bold';
                        }
                        updateNhUi();
                    };
                });
            }

            const ackBtn = box.querySelector('#_preImportBtnAck');
            if (ackBtn) ackBtn.onclick = () => cleanup('cancel');

            const fillBtn = box.querySelector('#_preImportBtnFill');
            if (fillBtn) fillBtn.onclick = () => { if (!fillBtn.disabled) cleanup('fill'); };

            const continueBtn = box.querySelector('#_preImportBtnContinue');
            if (continueBtn) continueBtn.onclick = () => { if (!continueBtn.disabled) cleanup('continue'); };

            const cancelBtn = box.querySelector('#_preImportBtnCancel');
            if (cancelBtn) cancelBtn.onclick = () => cleanup('cancel');
        });
    }

    // ── 共用 helper：強制從 A1 開始讀取工作表為陣列 ──────────────────
    // SheetJS 在 A、B 欄完全無資料時，會將 !ref 使用範圍自動從 C 欄起算，
    // 導致 sheet_to_json 回傳的陣列索引 0 其實對應的是 C 欄而非 A 欄，
    // 造成欄位偵測整體往左偏移。這裡明確將 range 起點鎖定在 (row 0, col 0)，
    // 確保陣列索引永遠與真實 Excel 欄位字母（A, B, C...）一一對應。
    function readSheetAsArray(ws) {
        if (!ws || !ws['!ref']) return XLSX.utils.sheet_to_json(ws, { header: 1 });
        const range = XLSX.utils.decode_range(ws['!ref']);
        range.s.r = 0;
        range.s.c = 0;
        return XLSX.utils.sheet_to_json(ws, { header: 1, range });
    }

    // ── Excel 結構預檢（在送往 content.js 之前，於 popup 端先行驗證） ──
    // 回傳 { ok: true } 或 { ok: false, issues: [{icon, title, detail}] }
    function preValidateExcelSheet(data, targetYymm) {
        const issues = [];

        // ── 共用：parseCellDate（與 content.js 保持相同邏輯） ──────────
        function parseCellDate(val) {
            if (val === undefined || val === null) return null;
            if (typeof val === 'number' && val > 1000) {
                const d = new Date(Math.round((val - 25569) * 86400 * 1000));
                return { month: d.getUTCMonth() + 1, day: d.getUTCDate() };
            }
            const s = String(val).trim();
            if (!s) return null;
            const mDate = s.match(/(?:\d{4}[\/\-])?(\d{1,2})[\/\-](\d{1,2})$/);
            if (mDate) {
                const month = parseInt(mDate[1]), day = parseInt(mDate[2]);
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31) return { month, day };
            }
            if (/^\d{1,2}$/.test(s)) {
                const n = parseInt(s);
                if (n >= 1 && n <= 31) return { month: null, day: n };
            }
            return null;
        }

        // 職編有效性檢核已改用外層共用的 isValidEmpId()

        // ── 日期行判定：需連續 8 欄皆為遞增連續數字（1,2,3...8）才視為日期行起點 ──
        // （原本只檢查相鄰 2 欄為 1、2，容易誤判班表資料中恰好出現「1、2」的班別代碼）
        const DATE_RUN_LENGTH = 8;
        function isDateRowStart(row, ci) {
            for (let k = 0; k < DATE_RUN_LENGTH; k++) {
                const cd = parseCellDate(row[ci + k]);
                if (!cd || cd.day !== k + 1) return false;
            }
            return true;
        }

        const EMP_KEYWORDS  = ["職編", "員工編號", "工號", "員編", "職員編號"];
        const NAME_KEYWORDS = ["姓名", "員工姓名", "名字"];
        let empIdColIdx = -1;
        let nameColIdx  = -1;
        let day1ColIdx  = -1;
        let day1RowIdx  = -1;  // 記錄日期列的 row index，供後續計算連續天數
        const dateRowMatches = []; // 記錄所有疑似「日期列」(含 1、2 連續欄位) 的 row/col，用於偵測多段日期

        // ── 水平掃描：只看前 20 欄（A~T），不限列數 ────────────────────
        const SCAN_COL_LIMIT = 20;
        const ROW_SCAN_LIMIT = 20; // 多段日期偵測僅限表頭區域（前 20 列），避免誤判下方班表資料
        for (let ri = 0; ri < data.length; ri++) {
            const row = data[ri];
            if (!row) continue;
            const colLimit = Math.min(SCAN_COL_LIMIT, row.length);
            for (let ci = 0; ci < colLimit; ci++) {
                const val = String(row[ci] || "").trim();
                if (empIdColIdx === -1 && EMP_KEYWORDS.some(k => val.includes(k))) {
                    empIdColIdx = ci;
                }
                if (nameColIdx === -1 && NAME_KEYWORDS.some(k => val.includes(k))) {
                    nameColIdx = ci;
                }
                // 多段日期偵測：限制在表頭區域（前 ROW_SCAN_LIMIT 列），
                // 避免下方員工班表資料列中恰好出現連續數字班別代碼造成誤判。
                // 需連續 8 欄皆為遞增連續數字才視為日期行，避免短序列誤判。
                const isDateStart = isDateRowStart(row, ci);
                if (isDateStart && ri < ROW_SCAN_LIMIT) {
                    // 同一列若已記錄過，不重複記錄（取該列第一個符合的欄位）
                    if (!dateRowMatches.some(m => m.rowIdx === ri)) {
                        dateRowMatches.push({ rowIdx: ri, colIdx: ci });
                    }
                }
                if (day1ColIdx === -1 && isDateStart) {
                    day1ColIdx = ci;
                    day1RowIdx = ri;
                }
            }
            // 注意：不再因找到日期列就提早 break，需繼續掃描其餘列以偵測是否有多段日期；
            // 職編/姓名欄位一旦找到即不再更新，但仍持續掃描全部列。
        }

        // 若關鍵字沒找到，嘗試以數字型職編推斷欄位（同樣限制前 20 欄）
        if (empIdColIdx === -1) {
            const colHits = {};
            const fallbackLimit = Math.min(SCAN_COL_LIMIT, day1ColIdx !== -1 ? day1ColIdx : SCAN_COL_LIMIT);
            for (let ri = 0; ri < data.length; ri++) {
                const row = data[ri]; if (!row) continue;
                for (let ci = 0; ci < fallbackLimit; ci++) {
                    if (isValidEmpId(String(row[ci] || "").trim()))
                        colHits[ci] = (colHits[ci] || 0) + 1;
                }
            }
            let bestCol = -1, bestHits = 1;
            for (const [ci, hits] of Object.entries(colHits)) {
                if (hits > bestHits) { bestHits = hits; bestCol = parseInt(ci); }
            }
            if (bestCol !== -1) empIdColIdx = bestCol;
        }
        // 姓名欄關鍵字掃描不到時，改用內容特徵偵測：
        // 若某欄「多數」儲存格內容皆為 2 個(含)以上的純中文字，視為姓名欄
        // （掃描範圍限制在職編欄之後、1號日期欄之前的表頭資料區，避免誤判到班表班別欄）
        if (nameColIdx === -1) {
            const nameScanColLimit = day1ColIdx !== -1 ? day1ColIdx : SCAN_COL_LIMIT;
            const chineseNameRe = /^[\u4e00-\u9fa5]{2,}/;  // 不再要求整格都是中文，只要求開頭是姓名
            const colStats = {}; // ci -> { hit, total }
            for (let ri = 0; ri < data.length; ri++) {
                const row = data[ri]; if (!row) continue;
                for (let ci = 0; ci < Math.min(nameScanColLimit, row.length); ci++) {
                    if (ci === empIdColIdx) continue;
                    const val = String(row[ci] || "").trim();
                    if (!val) continue;
                    if (!colStats[ci]) colStats[ci] = { hit: 0, total: 0 };
                    colStats[ci].total++;
                    if (chineseNameRe.test(val)) colStats[ci].hit++;
                }
            }
            let bestCol = -1, bestHits = 0;
            for (const [ci, s] of Object.entries(colStats)) {
                // 至少 2 筆命中、且命中率達 70% 以上，才視為姓名欄，避免誤判其他中文欄位
                if (s.hit >= 2 && s.hit / s.total >= 0.7 && s.hit > bestHits) {
                    bestHits = s.hit;
                    bestCol  = parseInt(ci);
                }
            }
            if (bestCol !== -1) nameColIdx = bestCol;
        }

        // 仍找不到時，退回舊邏輯：假設姓名欄緊接在職編欄之後
        if (nameColIdx === -1 && empIdColIdx !== -1 && empIdColIdx + 1 < SCAN_COL_LIMIT) {
            nameColIdx = empIdColIdx + 1;
        }

        // 錯誤①：完全找不到職編欄
        if (empIdColIdx === -1) {
            issues.push({
                icon:   '🔍',
                title:  '找不到職編欄位',
                detail: '掃描前 20 列均未偵測到「職編」、「員工編號」等關鍵字，\n也未發現 6～7 位數字型職編資料。\n請確認此工作表是否為正確的班表頁籤。',
            });
        }

        // 錯誤②：找不到 1 號欄
        if (day1ColIdx === -1) {
            issues.push({
                icon:   '📅',
                title:  '找不到月初 1 號欄位',
                detail: '掃描前 20 列均未找到連續的「1、2」日期欄位。\n可能原因：\n• 日期欄以非數字格式呈現（如「1日」、「01日」）\n• 班表日期列超過第 20 列，請確認格式',
            });
        }

        // 警告④：偵測到多段日期列（例如第3列與第5列都各有一段「1、2…」連續日期）
        // 只負責提醒使用者，不判斷哪一段才是正確的，也不阻擋匯入。
        let multiDateWarning = null;
        if (dateRowMatches.length > 1) {
            multiDateWarning = {
                rows: dateRowMatches.map(m => m.rowIdx + 1), // 轉成 1-based 給使用者看的列號
                count: dateRowMatches.length,
            };
        }

        // 錯誤③：資料天數不足
        // ★ 正確做法：從日期列本身數出「連續遞增整數」的天數，
        //   不用 row.length，避免被日期列後方的統計欄（OFF/假日/大/小）
        //   或員工列後方的統計公式欄撐大而誤判。
        let daysWarning = null;
        if (day1ColIdx !== -1 && day1RowIdx !== -1 && targetYymm && targetYymm.length === 6) {
            const tYear  = parseInt(targetYymm.substring(0, 4));
            const tMonth = parseInt(targetYymm.substring(4, 6));
            const expectedDays = new Date(tYear, tMonth, 0).getDate();

            // 從日期列的 day1ColIdx 開始，數連續遞增的天數
            const dateRow = data[day1RowIdx] || [];
            let consecutiveDays = 0;
            for (let ci = day1ColIdx; ci < dateRow.length; ci++) {
                const cd = parseCellDate(dateRow[ci]);
                if (cd && cd.day === consecutiveDays + 1) {
                    consecutiveDays++;
                } else {
                    break;  // 遇到非連續日期（包含統計欄）就停止
                }
            }

            if (consecutiveDays > 0 && consecutiveDays < expectedDays) {
                // 天數不足改為警告，不列入阻擋性 issues
                daysWarning = { tYear, tMonth, expectedDays, consecutiveDays, day1ColIdx, day1RowIdx };
            }
        }

        // ── 統一回傳：即使有 daysWarning，也絕不能遺漏職編欄／1號欄等
        // 阻擋性 issues（例如找不到職編欄），否則會讓錯誤被靜默吞掉。 ──
        return {
            ok: issues.length === 0,
            issues: issues.length > 0 ? issues : null,
            layout: { empIdColIdx, nameColIdx, day1ColIdx, day1RowIdx },
            daysWarning,
            multiDateWarning,
        };
    }

    // ── 顯示結構預檢錯誤視窗 ──────────────────────────────────────────
    // ── 補足 excelData 中缺少的日期天數（班表資料留空） ─────────────────
    function fillMissingDays(data, warn) {
        const { expectedDays, consecutiveDays, day1ColIdx, day1RowIdx } = warn;
        const filled = data.map(row => row ? [...row] : []);
        for (let d = consecutiveDays + 1; d <= expectedDays; d++) {
            const colIdx = day1ColIdx + (d - 1);
            while (filled[day1RowIdx].length <= colIdx) filled[day1RowIdx].push(undefined);
            filled[day1RowIdx][colIdx] = d;
            // 其他資料列該欄保持 undefined（空白），不另行寫入
        }
        return filled;
    }

    async function processExcelSheet(sheetName) {
        if (!currentWorkbook) {
            statusDiv.textContent = "❌ 請先載入 Excel 檔案";
            return;
        }
        statusDiv.textContent = `⏳ 正在處理工作表 [${sheetName}]...`;
        const excelData = readSheetAsArray(currentWorkbook.Sheets[sheetName]);

        // ── 結構預檢：有問題直接阻擋，不送往 content.js ────────────────
        const storageForYymm = await chrome.storage.local.get('lastMonthData');
        const oldYymm = storageForYymm.lastMonthData?.yymm || "";
        const targetYymm = oldYymm
            ? (() => {
                let y = parseInt(oldYymm.slice(0, 4)), m = parseInt(oldYymm.slice(4, 6)) + 1;
                if (m > 12) { m = 1; y++; }
                return String(y) + String(m).padStart(2, '0');
              })()
            : "";
        const preCheck = preValidateExcelSheet(excelData, targetYymm);

        // ── 整合式匯入前確認：一次收集並顯示訊息 ────────────────────────
        //   1. 欄位偵測結果（恆顯示）
        //   • 欄位偵測失敗，如找不到職編欄／找不到1號欄（如有，強制中止）
        //   2. 日期行有重複（如有，強制中止）
        //   3. 人員本月有、下個月無（如有）
        //   4. 人員本月無、下個月有（如有）
        //   5. 班表天數不足（如有）
        const expectedEmpCount = countExpectedEmployees(excelData, preCheck.layout);

        // 欄位偵測失敗（職編欄或1號欄找不到）時，layout 不可靠，
        // 無法可靠解析 Excel，故不呼叫 content.js 做本月／下月人員名單比對。
        let memberRes = null;
        if (preCheck.ok) {
            statusDiv.textContent = `📋 [${sheetName}] 正在比對本月／下月人員名單...`;
            memberRes = await sendMessage({ action: "preflightWarnings", excelData, sheetName });
            if (!memberRes) {
                statusDiv.textContent = `❌ [${sheetName}] 無法連線至頁面，請重整後再試`;
                return;
            }
        }

        const consolidatedReport = {
            layout: preCheck.layout,
            empCount: expectedEmpCount,
            structuralIssues: preCheck.issues || null,
            multiDateWarning: preCheck.multiDateWarning || null,
            daysWarning: preCheck.daysWarning || null,
            departedWarnings: memberRes?.success ? (memberRes.departedWarnings || []) : [],
            noOldDataWarnings: memberRes?.success ? (memberRes.noOldDataWarnings || []) : [],
            nhRequired: memberRes?.success ? (memberRes.nhRequired || 0) : 0,
            targetYymm,
        };

        statusDiv.textContent = `📋 [${sheetName}] 請確認偵測結果...`;
        const decision = await showConsolidatedPreImportWindow(sheetName, consolidatedReport);
        if (decision.action === 'cancel') {
            statusDiv.textContent = (preCheck.issues || preCheck.multiDateWarning)
                ? `⛔ [${sheetName}] 已中止，請修正 Excel 後重新匯入`
                : `⛔ [${sheetName}] 已取消匯入，請重新檢查檔案`;
            return;
        }
        lastNhDates = decision.nhDates || []; // 供本次及後續「重新執行檢測」沿用同一批國定假日日期

        let finalExcelData = excelData;
        if (decision.action === 'fill' && preCheck.daysWarning) {
            finalExcelData = fillMissingDays(excelData, preCheck.daysWarning);
            statusDiv.textContent = `⏳ 已補足天數，正在處理工作表 [${sheetName}]...`;
        }

        const set = await chrome.storage.local.get(['showExcelReport', 'autoMode', 'blankFillMode', 'blankFillCode']);
        const res = await sendMessage({
            action: "autoProcessExcel",
            excelData: finalExcelData,
            sheetName,
            showReport:    set.showExcelReport !== false,
            blankFillMode: set.blankFillMode || 'keep',
            blankFillCode: set.blankFillCode || '',
            nhDates: lastNhDates,
        });
        if (res?.overtimeGaps && res.overtimeGaps.length > 0) {
            const labels = res.overtimeGaps.map(g => `${g.originExcel || '空白來源'}：${g.over}→${g.sys}`).join('、');
            await chrome.storage.local.set({ pendingOvertimeGaps: res.overtimeGaps });
            statusDiv.textContent = `⚠️ 匯入後發現 ${res.overtimeGaps.length} 組自定班別缺少自身 W+／N+ 對應：${labels}`;
            chrome.windows.create({ url: 'dict_manager.html', type: 'popup', width: 780, height: 500 });
            hidePopupMenuAfterImport('已開啟班別字典管理，請補齊各 Excel 代號自身的 W+／N+ 對應後重新匯入。');
        } else if (res?.success) {
            document.getElementById('step3Box').style.display = 'block';
            document.getElementById('step4Box').style.display = 'block';
            statusDiv.textContent = `✅ [${sheetName}] 通過檢測，可執行寫入`;
            // 註：「人員本月無、下個月有」等名單差異已於匯入前的整合式確認視窗中提示，
            // 此處不再重複彈出視窗。

            if (set.autoMode && confirm("✅ 檢測通過，是否立即寫入？")) {
                document.getElementById('step4Btn').click();
                return; // 寫入完成後由 step4 handler 關閉 popup
            }
            hidePopupMenuAfterImport('Excel 已匯入，請查看頁面上的檢測報告或進行下一步操作。');
        } else if (res?.unknownCodes && res.unknownCodes.length > 0) {
            await chrome.storage.local.set({ pendingUnknownCodes: res.unknownCodes });
            statusDiv.textContent = `⚠️ 發現 ${res.unknownCodes.length} 個未知班別：${res.unknownCodes.join('、')}，請在字典管理中補填後重新匯入。`;
            chrome.windows.create({ url: 'dict_manager.html', type: 'popup', width: 780, height: 500 });
            hidePopupMenuAfterImport('已開啟班別字典管理，請補齊對應後重新匯入。');
        } else {
            statusDiv.textContent = res?.message || `❌ [${sheetName}] 檢測未通過，請確認錯誤訊息`;
            // content.js 會在有檢測錯誤時於網頁端顯示報告，關閉右上角選單避免遮擋。
            if (res) hidePopupMenuAfterImport('Excel 已匯入，請查看頁面上的檢測報告。');
        }
    }

    document.getElementById('step3Btn').onclick = async () => {
        if (!currentWorkbook) { statusDiv.textContent = "❌ 請先載入 Excel 檔案"; return; }
        const sheetName = lastSelectedSheet || currentWorkbook.SheetNames[0];
        const excelData = readSheetAsArray(currentWorkbook.Sheets[sheetName]);
        statusDiv.textContent = "⏳ 重新執行檢測...";
        const set3 = await chrome.storage.local.get(['blankFillMode', 'blankFillCode']);
        const res = await sendMessage({
            action: "autoProcessExcel",
            excelData,
            sheetName,
            showReport:    true,
            blankFillMode: set3.blankFillMode || 'keep',
            blankFillCode: set3.blankFillCode || '',
            nhDates: lastNhDates,
        });
        if (res?.overtimeGaps && res.overtimeGaps.length > 0) {
            const labels = res.overtimeGaps.map(g => `${g.originExcel || '空白來源'}：${g.over}→${g.sys}`).join('、');
            await chrome.storage.local.set({ pendingOvertimeGaps: res.overtimeGaps });
            statusDiv.textContent = `⚠️ 仍缺少 ${res.overtimeGaps.length} 組自身 W+／N+ 對應：${labels}`;
            chrome.windows.create({ url: 'dict_manager.html', type: 'popup', width: 780, height: 500 });
            hidePopupMenuAfterImport('請先補齊各 Excel 代號自身的 W+／N+ 對應，再重新檢測。');
        } else {
            statusDiv.textContent = res?.success ? "✅ 檢測完成，請查看報告" : (res?.message || "❌ 檢測未通過");
        }
    };

    document.getElementById('step4Btn').onclick = async () => {
        if (!currentWorkbook) {
            statusDiv.textContent = "❌ 請先載入 Excel 檔案";
            return;
        }
        const sheetName = lastSelectedSheet || currentWorkbook.SheetNames[0];
        const excelData = readSheetAsArray(currentWorkbook.Sheets[sheetName]);
        statusDiv.textContent = "⏳ 寫入中，請稍候...";
        const res = await sendMessage({ action: "injectOnly", excelData });
        statusDiv.textContent = res?.message || (res?.success ? "✅ 寫入完成" : "❌ 寫入失敗，請重整頁面");
        if (res?.success) hidePopupMenuAfterImport('寫入完成。');
    };
});