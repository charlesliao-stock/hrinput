document.addEventListener('DOMContentLoaded', async () => {
    if (typeof XLSX === 'undefined') {
        document.getElementById('status').textContent = '❌ xlsx 函式庫載入失敗，請確認 xlsx.full.min.js 存在';
        document.getElementById('step2Btn').disabled = true;
        return;
    }
    const statusDiv = document.getElementById('status'), excelFile = document.getElementById('excelFile');
    let currentWorkbook = null;
    let lastSelectedSheet = null;

    // --- 更新提醒邏輯 ---
    const updateAlert = document.getElementById('updateAlert');
    const updateVersion = document.getElementById('updateVersion');
    const downloadUpdateBtn = document.getElementById('downloadUpdateBtn');

    chrome.storage.local.get(['updateAvailable', 'latestVersion', 'downloadUrl'], (data) => {
        if (data.updateAvailable) {
            updateAlert.style.display = 'block';
            updateVersion.textContent = `最新版本：v${data.latestVersion}`;
            downloadUpdateBtn.onclick = () => {
                chrome.tabs.create({ url: 'https://github.com/charlesliao-stock/hrinput' });
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

    document.getElementById('openQuickSettings').onclick = () => chrome.windows.create({ url: 'quick_settings.html', type: 'popup', width: 360, height: 400 });
    document.getElementById('openDictManager').onclick   = () => chrome.windows.create({ url: 'dict_manager.html',   type: 'popup', width: 780, height: 500 });

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
        const res = await sendMessage({
            action: "readAndMemorize",
            showPreview: set.showWebPreview !== false,
            autoMode: set.autoMode || false
        });

        if (res?.success) {
            let msg = `✅ 記憶完成 (${res.yymm})`;
            if (res.targetPeriod) {
                msg += `\n📅 檢測週期：【${res.targetPeriod.label}】${res.targetPeriod.start}～${res.targetPeriod.end}`;
            } else if (res.periods && res.periods.length === 0) {
                msg += `\n⚠️ 未偵測到四週變形週期，請確認頁面`;
            }
            statusDiv.textContent = msg;

            if (set.autoMode && res.nextUrl) {
                if (res.hasPreview) {
                    statusDiv.textContent = `${msg}\n⌛ 請查看網頁預覽，關閉後將自動跳轉...`;
                } else {
                    statusDiv.textContent = "⚡ 立即執行自動跳轉...";
                    setTimeout(() => chrome.tabs.update({ url: res.nextUrl }), 800);
                }
            }
        } else {
            statusDiv.textContent = res?.message || "❌ 記憶失敗，請確認頁面正確";
        }
    };

    // --- 步驟 2：選擇 Excel 檔案 ---
    document.getElementById('step2Btn').onclick = () => excelFile.click();

    excelFile.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;

        statusDiv.textContent = "⏳ 讀取 Excel 檔案中...";
        const reader = new FileReader();

        reader.onload = async (ev) => {
            try {
                currentWorkbook = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
                const sheetNames = currentWorkbook.SheetNames;
                document.getElementById('sheetSelectBox').style.display = 'none';

                if (sheetNames.length === 0) {
                    statusDiv.textContent = "❌ Excel 檔案中沒有任何工作表";
                    return;
                }

                if (sheetNames.length === 1) {
                    statusDiv.textContent = `偵測到唯一工作表「${sheetNames[0]}」，自動匯入中...`;
                    lastSelectedSheet = sheetNames[0];
                    await processExcelSheet(sheetNames[0]);
                } else {
                    const sel = document.getElementById('sheetSelect');
                    sel.innerHTML = sheetNames.map((name, i) =>
                        `<option value="${name}">${i + 1}. ${name}</option>`
                    ).join('');
                    document.getElementById('sheetSelectBox').style.display = 'block';
                    statusDiv.textContent = `📋 偵測到 ${sheetNames.length} 個工作表，請選擇後按確認`;
                }
            } catch (err) {
                console.error('[Excel 讀取錯誤]', err);
                statusDiv.textContent = "❌ Excel 讀取失敗：" + err.message;
            }
        };

        reader.readAsArrayBuffer(file);
        e.target.value = "";
    });

    document.getElementById('sheetConfirmBtn').onclick = async () => {
        const selectedSheet = document.getElementById('sheetSelect').value;
        if (!selectedSheet || !currentWorkbook) return;
        document.getElementById('sheetSelectBox').style.display = 'none';
        lastSelectedSheet = selectedSheet;
        await processExcelSheet(selectedSheet);
    };

    async function processExcelSheet(sheetName) {
        if (!currentWorkbook) {
            statusDiv.textContent = "❌ 請先載入 Excel 檔案";
            return;
        }
        statusDiv.textContent = `⏳ 正在處理工作表 [${sheetName}]...`;
        const excelData = XLSX.utils.sheet_to_json(currentWorkbook.Sheets[sheetName], { header: 1 });
        const set = await chrome.storage.local.get(['showExcelReport', 'autoMode', 'blankFillMode', 'blankFillCode']);
        const res = await sendMessage({
            action: "autoProcessExcel",
            excelData,
            sheetName,
            showReport:    set.showExcelReport !== false,
            blankFillMode: set.blankFillMode || 'keep',
            blankFillCode: set.blankFillCode || '',
        });
        if (res?.success) {
            document.getElementById('step3Box').style.display = 'block';
            document.getElementById('step4Box').style.display = 'block';
            statusDiv.textContent = `✅ [${sheetName}] 通過檢測，可執行寫入`;

            if (res.noOldDataWarnings && res.noOldDataWarnings.length > 0) {
                const list = res.noOldDataWarnings.map(w => `• ${w.empId} ${w.name}`).join('<br>');
                const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
                <style>
                    body { width:320px; margin:0; display:flex; flex-direction:column;
                           align-items:center; gap:14px; padding:24px 20px;
                           box-sizing:border-box; font-family:"Microsoft JhengHei",sans-serif; background:#fff; }
                    h4 { margin:0; color:#e67e22; font-size:14px; text-align:center; }
                    .list { width:100%; background:#fff8f0; border:1px solid #f0c080; border-radius:6px;
                            padding:8px 12px; font-size:13px; color:#2c3e50; line-height:1.8; }
                    .note { font-size:12px; color:#888; text-align:center; line-height:1.6; }
                    button { width:100%; padding:10px; background:#e67e22; color:white; border:none;
                             border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer; }
                    button:hover { background:#ca6f1e; }
                </style></head><body>
                <h4>⚠️ 以下人員無舊月資料</h4>
                <div class="list">${list}</div>
                <div class="note">無法檢測上個月的資料<br>但班表仍可正常寫入。</div>
                <button onclick="window.close()">確定</button>
                </body></html>`;
                chrome.windows.create({
                    url: 'data:text/html;charset=utf-8,' + encodeURIComponent(html),
                    type: 'popup', width: 340,
                    height: 220 + res.noOldDataWarnings.length * 26,
                    focused: true
                });
            }

            if (set.autoMode && confirm("✅ 檢測通過，是否立即寫入？")) {
                document.getElementById('step4Btn').click();
            }
        } else if (res?.unknownCodes && res.unknownCodes.length > 0) {
            await chrome.storage.local.set({ pendingUnknownCodes: res.unknownCodes });
            statusDiv.textContent = `⚠️ 發現 ${res.unknownCodes.length} 個未知班別：${res.unknownCodes.join('、')}，請在字典管理中補填後重新匯入。`;
            chrome.windows.create({ url: 'dict_manager.html', type: 'popup', width: 780, height: 500 });
        } else {
            statusDiv.textContent = res?.message || `❌ [${sheetName}] 檢測未通過，請確認班別字典`;
        }
    }

    document.getElementById('step3Btn').onclick = async () => {
        if (!currentWorkbook) { statusDiv.textContent = "❌ 請先載入 Excel 檔案"; return; }
        const sheetName = lastSelectedSheet || currentWorkbook.SheetNames[0];
        const excelData = XLSX.utils.sheet_to_json(currentWorkbook.Sheets[sheetName], { header: 1 });
        statusDiv.textContent = "⏳ 重新執行檢測...";
        const set3 = await chrome.storage.local.get(['blankFillMode', 'blankFillCode']);
        const res = await sendMessage({
            action: "autoProcessExcel",
            excelData,
            sheetName,
            showReport:    true,
            blankFillMode: set3.blankFillMode || 'keep',
            blankFillCode: set3.blankFillCode || '',
        });
        statusDiv.textContent = res?.success ? "✅ 檢測完成，請查看報告" : (res?.message || "❌ 檢測未通過");
    };

    document.getElementById('step4Btn').onclick = async () => {
        if (!currentWorkbook) {
            statusDiv.textContent = "❌ 請先載入 Excel 檔案";
            return;
        }
        const sheetName = lastSelectedSheet || currentWorkbook.SheetNames[0];
        const excelData = XLSX.utils.sheet_to_json(currentWorkbook.Sheets[sheetName], { header: 1 });
        statusDiv.textContent = "⏳ 寫入中，請稍候...";
        const res = await sendMessage({ action: "injectOnly", excelData });
        statusDiv.textContent = res?.message || (res?.success ? "✅ 寫入完成" : "❌ 寫入失敗，請重整頁面");
    };
});
