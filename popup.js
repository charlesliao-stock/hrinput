document.addEventListener('DOMContentLoaded', async () => {
    if (typeof XLSX === 'undefined') {
        document.getElementById('status').textContent = '❌ xlsx 函式庫載入失敗，請確認 xlsx.full.min.js 存在';
        document.getElementById('step2Btn').disabled = true;
        return;
    }
    const statusDiv = document.getElementById('status'), excelFile = document.getElementById('excelFile');
    let currentWorkbook = null;
    let lastSelectedSheet = null;

    // ── 步驟 1 按鈕文字更新（開啟時 + 記憶成功後共用） ──────────────
    function updateStep1BtnLabel(yymm) {
        const btn = document.getElementById('step1Btn');
        if (yymm && yymm.length === 6) {
            const y = yymm.substring(0, 4);
            const m = yymm.substring(4, 6);
            btn.textContent = `✅ 已記憶 ${y}/${m}（點擊重新讀取）`;
            btn.style.background = '#27ae60';
        } else {
            btn.textContent = '💾 記憶本月班表並跳轉至次月';
            btn.style.background = '';
        }
    }

    // 頁面開啟時，若 storage 已有記憶資料則立即反映
    chrome.storage.local.get('lastMonthData', (d) => {
        if (d.lastMonthData?.yymm) updateStep1BtnLabel(d.lastMonthData.yymm);
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
            showPreview: set.showWebPreview === true,  // 預設不顯示，需明確設為 true 才顯示
            autoMode: set.autoMode || false
        });

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

        const EMP_KEYWORDS = ["職編", "員工編號", "工號", "員編", "職員編號"];
        let empIdColIdx = -1;
        let day1ColIdx  = -1;
        let day1RowIdx  = -1;  // 記錄日期列的 row index，供後續計算連續天數

        // ── 掃描前 20 列，找職編欄與 1 號欄（無欄位上限） ────────────
        for (let ri = 0; ri < Math.min(20, data.length); ri++) {
            const row = data[ri];
            if (!row) continue;
            for (let ci = 0; ci < row.length; ci++) {
                const val = String(row[ci] || "").trim();
                if (empIdColIdx === -1 && EMP_KEYWORDS.some(k => val.includes(k))) {
                    empIdColIdx = ci;
                }
                if (day1ColIdx === -1) {
                    const cd  = parseCellDate(row[ci]);
                    const cd2 = parseCellDate(row[ci + 1]);
                    if (cd?.day === 1 && cd2?.day === 2) {
                        day1ColIdx = ci;
                        day1RowIdx = ri;
                    }
                }
            }
            if (empIdColIdx !== -1 && day1ColIdx !== -1) break;
        }

        // 若關鍵字沒找到，嘗試以數字型職編推斷欄位
        if (empIdColIdx === -1) {
            const colHits = {};
            for (let ri = 0; ri < data.length; ri++) {
                const row = data[ri]; if (!row) continue;
                const limit = day1ColIdx !== -1 ? day1ColIdx : row.length;
                for (let ci = 0; ci < limit; ci++) {
                    if (/^\d{6,7}$/.test(String(row[ci] || "").trim()))
                        colHits[ci] = (colHits[ci] || 0) + 1;
                }
            }
            let bestCol = -1, bestHits = 1;
            for (const [ci, hits] of Object.entries(colHits)) {
                if (hits > bestHits) { bestHits = hits; bestCol = parseInt(ci); }
            }
            if (bestCol !== -1) empIdColIdx = bestCol;
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

        // 錯誤③：資料天數不足
        // ★ 正確做法：從日期列本身數出「連續遞增整數」的天數，
        //   不用 row.length，避免被日期列後方的統計欄（OFF/假日/大/小）
        //   或員工列後方的統計公式欄撐大而誤判。
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
                return {
                    ok: true,
                    daysWarning: {
                        tYear, tMonth,
                        expectedDays,
                        consecutiveDays,
                        day1ColIdx,
                        day1RowIdx,
                    }
                };
            }
        }

        return issues.length === 0 ? { ok: true } : { ok: false, issues };
    }

    // ── 顯示結構預檢錯誤視窗 ──────────────────────────────────────────
    function showStructureErrorWindow(sheetName, issues) {
        const itemsHtml = issues.map(({ icon, title, detail }) => `
            <div class="issue-block">
                <div class="issue-title">${icon} ${title}</div>
                <div class="issue-detail">${detail.replace(/\n/g, '<br>')}</div>
            </div>`).join('');
        const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
        <style>
            body { width:380px; margin:0; padding:20px; box-sizing:border-box;
                   font-family:"Microsoft JhengHei",sans-serif; background:#fff; }
            h4 { margin:0 0 4px 0; color:#c0392b; font-size:15px; }
            .sub { font-size:12px; color:#888; margin-bottom:14px; }
            .issue-block { background:#fff5f5; border:1px solid #f5c6cb;
                           border-radius:6px; padding:10px 12px; margin-bottom:10px; }
            .issue-title { font-weight:bold; font-size:13px; color:#c0392b; margin-bottom:5px; }
            .issue-detail { font-size:12px; color:#555; line-height:1.7; }
            button { width:100%; padding:10px; margin-top:6px; background:#e74c3c; color:white;
                     border:none; border-radius:6px; font-size:14px;
                     font-weight:bold; cursor:pointer; }
            button:hover { background:#c0392b; }
        </style></head><body>
        <h4>❌ Excel 結構預檢失敗</h4>
        <div class="sub">工作表：${sheetName}　｜　共發現 ${issues.length} 個問題</div>
        ${itemsHtml}
        <button onclick="window.close()">確定，重新確認檔案</button>
        </body></html>`;
        const winHeight = Math.min(140 + issues.length * 140, 600);
        chrome.windows.create({
            url: 'data:text/html;charset=utf-8,' + encodeURIComponent(html),
            type: 'popup', width: 410, height: winHeight, focused: true,
        });
    }

    // ── 天數不足：在 popup 頁面內顯示 modal，回傳 Promise<'fill'|'cancel'> ──
    function showDaysWarningWindow(sheetName, warn) {
        return new Promise((resolve) => {
            const { tYear, tMonth, expectedDays, consecutiveDays } = warn;
            const missing = expectedDays - consecutiveDays;

            // 建立遮罩
            const overlay = document.createElement('div');
            overlay.style.cssText = [
                'position:fixed', 'inset:0', 'background:rgba(0,0,0,0.45)',
                'display:flex', 'align-items:center', 'justify-content:center',
                'z-index:9999',
            ].join(';');

            // 建立對話框
            const box = document.createElement('div');
            box.style.cssText = [
                'background:#fff', 'border-radius:8px', 'padding:18px 18px 14px',
                'width:320px', 'box-shadow:0 4px 20px rgba(0,0,0,0.25)',
                'font-family:"Microsoft JhengHei",sans-serif',
            ].join(';');

            box.innerHTML = `
                <div style="font-size:15px;font-weight:bold;color:#e67e22;margin-bottom:4px">⚠️ 班表天數不足</div>
                <div style="font-size:11px;color:#999;margin-bottom:12px">工作表：${sheetName}</div>
                <div style="background:#fffbf0;border:1px solid #f0c060;border-radius:6px;
                            padding:10px 12px;margin-bottom:12px;font-size:12px;color:#555;line-height:1.8">
                    目標月份 <b style="color:#c0392b">${tYear} 年 ${tMonth} 月</b> 應有
                    <b style="color:#c0392b">${expectedDays} 天</b>，<br>
                    但日期列中只偵測到連續 <b style="color:#c0392b">${consecutiveDays} 天</b>
                    （缺少 <b style="color:#c0392b">${missing} 天</b>）。
                </div>
                <div style="font-size:11px;color:#888;margin-bottom:14px;line-height:1.7">
                    選擇「補足後繼續」：系統將自動在日期列末端補上第
                    ${consecutiveDays + 1}～${expectedDays} 天，班表資料欄留空，繼續匯入。<br>
                    若班表月份有誤，請取消後重新確認檔案。
                </div>
                <button id="_warnBtnFill" style="width:100%;padding:9px;margin-bottom:7px;
                    background:#27ae60;color:#fff;border:none;border-radius:6px;
                    font-size:12px;font-weight:bold;cursor:pointer">
                    📅 系統補足天數欄位（班表資料空白）後繼續
                </button>
                <button id="_warnBtnCancel" style="width:100%;padding:9px;
                    background:#bdc3c7;color:#2c3e50;border:none;border-radius:6px;
                    font-size:12px;font-weight:bold;cursor:pointer">
                    ✖ 取消，重新確認檔案
                </button>`;

            overlay.appendChild(box);
            document.body.appendChild(overlay);

            function cleanup(choice) {
                document.body.removeChild(overlay);
                resolve(choice);
            }

            box.querySelector('#_warnBtnFill').onclick   = () => cleanup('fill');
            box.querySelector('#_warnBtnCancel').onclick = () => cleanup('cancel');
        });
    }

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
        const excelData = XLSX.utils.sheet_to_json(currentWorkbook.Sheets[sheetName], { header: 1 });

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
        if (!preCheck.ok) {
            statusDiv.textContent = `❌ [${sheetName}] 結構預檢未通過，請查看錯誤說明`;
            showStructureErrorWindow(sheetName, preCheck.issues);
            return;
        }

        // ── 天數不足：警告視窗，讓使用者決定是否補足後繼續 ──────────────
        let finalExcelData = excelData;
        if (preCheck.daysWarning) {
            statusDiv.textContent = `⚠️ [${sheetName}] 天數不足，請查看警告說明`;
            const choice = await showDaysWarningWindow(sheetName, preCheck.daysWarning);
            if (choice !== 'fill') {
                statusDiv.textContent = `⛔ [${sheetName}] 已取消匯入`;
                return;
            }
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
            statusDiv.textContent = res?.message || `❌ [${sheetName}] 檢測未通過，請確認錯誤訊息`;
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