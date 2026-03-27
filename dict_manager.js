document.addEventListener('DOMContentLoaded', async () => {
    const hrBody     = document.getElementById('hrBody');
    const customBody = document.getElementById('customBody');

    // 跳過接班檢測的班別（休假類 + 加班類），時間欄顯示為灰色不可輸入
    const SKIP_SHIFT_CODES = new Set(['FF', 'WW', 'NH', 'N+', 'W+']);

    // 摺疊面板控制
    const coll = document.getElementsByClassName("collapsible");
    for (let i = 0; i < coll.length; i++) {
        coll[i].addEventListener("click", function () {
            this.classList.toggle("active");
            const content = this.nextElementSibling;
            content.style.display = (content.style.display === "block") ? "none" : "block";
            updateWindowHeight();
        });
    }

    const data = await chrome.storage.local.get(['hrShifts', 'shiftDict', 'pendingUnknownCodes']);

    const hrShifts     = (data.hrShifts && data.hrShifts.length > 0) ? data.hrShifts : [];
    const customShifts = data.shiftDict || [];
    const pendingCodes = data.pendingUnknownCodes || [];

    hrShifts.forEach(item => addHrRow(item));
    customShifts.forEach(item => addCustomRow(item));

    // ✅ 若有待補填的未知班別，自動展開自定義區塊並預填
    if (pendingCodes.length > 0) {
        chrome.storage.local.remove('pendingUnknownCodes');

        const customCollapsible = document.getElementsByClassName("collapsible")[1];
        customCollapsible.classList.add("active");
        customCollapsible.nextElementSibling.style.display = "block";

        pendingCodes.forEach(code => addCustomRow({ excel: code, sys: '', over: '', am: '', pm: '', night: '' }));

        const banner = document.createElement('div');
        banner.className = 'error-banner';
        banner.id = 'unknown-banner';
        banner.innerHTML = `⚠️ 發現 <b>${pendingCodes.length}</b> 個未知班別（<b>${pendingCodes.join('、')}</b>）已自動加入下方，請填寫「系統」欄後儲存，再重新匯入 Excel。`;
        document.querySelector('.main-container').insertBefore(banner, document.querySelector('.scroll-area'));

        setTimeout(() => customCollapsible.scrollIntoView({ behavior: 'smooth', block: 'start' }), 200);
        updateWindowHeight();
    }

    document.getElementById('addHrRow').onclick    = () => addHrRow({ code: '', start: null, end: null });
    document.getElementById('addCustomRow').onclick = () => addCustomRow();

    // ── HR 班別列（代號 + 上班時間 + 下班時間 + 備註） ──────────────────
    function addHrRow(item = { code: '', start: null, end: null }) {
        // 相容舊格式：若 item 是字串（純代號），轉換為物件
        if (typeof item === 'string') {
            item = { code: item, start: null, end: null };
        }

        const code  = item.code  || '';
        const start = item.start || '';
        const end   = item.end   || '';

        const isSkip = SKIP_SHIFT_CODES.has(code);

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" class="hr-code" value="${code}" maxlength="5" placeholder="代號"></td>
            <td><input type="time" class="hr-start${isSkip ? ' skip-shift' : ''}" value="${isSkip ? '' : start}" ${isSkip ? 'disabled title="休假/加班類，不參與接班檢測"' : ''}></td>
            <td><input type="time" class="hr-end${isSkip ? ' skip-shift' : ''}"   value="${isSkip ? '' : end}"   ${isSkip ? 'disabled title="休假/加班類，不參與接班檢測"' : ''}></td>
            <td class="${isSkip ? 'skip-label' : ''}">${isSkip ? '跳過接班檢測' : ''}</td>
            <td><button class="del-btn">刪</button></td>
        `;

        // 代號變更時，動態切換時間欄的跳過狀態
        const codeInput  = tr.querySelector('.hr-code');
        const startInput = tr.querySelector('.hr-start');
        const endInput   = tr.querySelector('.hr-end');
        const noteCell   = tr.querySelector('td:nth-child(4)');

        codeInput.addEventListener('input', () => {
            const newCode  = codeInput.value.trim().toUpperCase();
            const newIsSkip = SKIP_SHIFT_CODES.has(newCode);
            startInput.disabled = newIsSkip;
            endInput.disabled   = newIsSkip;
            startInput.classList.toggle('skip-shift', newIsSkip);
            endInput.classList.toggle('skip-shift',   newIsSkip);
            if (newIsSkip) {
                startInput.value = '';
                endInput.value   = '';
                noteCell.textContent = '跳過接班檢測';
                noteCell.className   = 'skip-label';
            } else {
                noteCell.textContent = '';
                noteCell.className   = '';
            }
        });

        tr.querySelector('.del-btn').onclick = () => tr.remove();
        hrBody.appendChild(tr);
    }

    // ── 自定義班別列 ─────────────────────────────────────────────────────
    function addCustomRow(item = { excel: '', sys: '', over: '', am: '', pm: '', night: '' }) {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" value="${item.excel || ''}" placeholder="Excel代號"></td>
            <td><input type="text" class="sys-input" value="${item.sys || ''}" placeholder="系統代號 *"></td>
            <td><input type="text" value="${item.over  || ''}"></td>
            <td><input type="text" value="${item.am    || ''}"></td>
            <td><input type="text" value="${item.pm    || ''}"></td>
            <td><input type="text" value="${item.night || ''}"></td>
            <td><button class="del-btn">刪</button></td>
        `;

        // sys 欄即時標紅
        const sysInput = tr.querySelector('.sys-input');
        sysInput.addEventListener('input', () => {
            sysInput.classList.toggle('sys-empty', sysInput.value.trim() === '');
        });
        // 初始狀態：若 sys 已空則標紅（pendingCodes 預填的列）
        if (!item.sys || item.sys.trim() === '') {
            sysInput.classList.add('sys-empty');
        }

        tr.querySelector('.del-btn').onclick = () => tr.remove();
        customBody.appendChild(tr);
    }

// ── 儲存 ─────────────────────────────────────────────────────────────
    document.getElementById('saveAll').onclick = async () => {

        // 驗證：自定義班別的 sys 欄不可為空，且 W+/N+ 必須填寫逾時欄位
        const badRows = Array.from(customBody.querySelectorAll('tr')).filter(tr => {
            const inputs = tr.querySelectorAll('input');
            const excel  = inputs[0]?.value.trim();
            const sys    = inputs[1]?.value.trim().toUpperCase();
            const over   = inputs[2]?.value.trim();

            // 條件 1: 有 Excel 代號但沒填系統代號
            const isSysMissing = excel && !sys;
            
            // 條件 2: 系統代號是 W+ 或 N+，但逾時欄位沒填
            const isOverMissing = (sys === 'W+' || sys === 'N+') && !over;

            // 標註視覺紅框 (即時反映)
            inputs[1].classList.toggle('sys-empty', isSysMissing);
            inputs[2].classList.toggle('sys-empty', isOverMissing);

            return isSysMissing || isOverMissing;
        });

        if (badRows.length > 0) {
            // 展開自定義區塊
            const customCollapsible = document.getElementsByClassName("collapsible")[1];
            if (customCollapsible.nextElementSibling.style.display !== "block") {
                customCollapsible.classList.add("active");
                customCollapsible.nextElementSibling.style.display = "block";
                updateWindowHeight();
            }

            badRows[0].scrollIntoView({ behavior: 'smooth', block: 'center' });

            const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
            <style>
                body { width:320px; margin:0; display:flex; flex-direction:column;
                       align-items:center; gap:14px; padding:24px 20px;
                       box-sizing:border-box; font-family:"Microsoft JhengHei",sans-serif; background:#fff; }
                h4 { margin:0; color:#e74c3c; font-size:14px; text-align:center; }
                .note { font-size:13px; color:#2c3e50; text-align:center; line-height:1.6; }
                button { width:100%; padding:10px; background:#e74c3c; color:white; border:none;
                         border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer; }
                button:hover { background:#c0392b; }
            </style></head><body>
            <h4>❌ 儲存失敗</h4>
            <div class="note">欄位填寫不完整：<br>1. 系統代號不可空白<br>2. <b>W+ 或 N+ 班別必須填寫「逾時」欄位</b></div>
            <button onclick="window.close()">確定</button>
            </body></html>`;
            
            chrome.windows.create({
                url: 'data:text/html;charset=utf-8,' + encodeURIComponent(html),
                type: 'popup', width: 340, height: 240, focused: true
            });
            return; // 阻擋儲存
        }

        // 收集 HR 班別資料（先收集，用於後續比對）
        const newHr = Array.from(hrBody.querySelectorAll('tr')).map(tr => {
            const code  = tr.querySelector('.hr-code')?.value.trim()  || '';
            const start = tr.querySelector('.hr-start')?.value.trim() || null;
            const end   = tr.querySelector('.hr-end')?.value.trim()   || null;
            return {
                code,
                start: start || null,
                end:   end   || null,
            };
        }).filter(item => item.code);

        // 收集自定義班別資料
        const newCustom = Array.from(customBody.querySelectorAll('tr')).map(tr => {
            const ins = tr.querySelectorAll('input');
            return {
                excel: ins[0].value.trim(),
                sys:   ins[1].value.trim(),
                over:  ins[2].value.trim(),
                am:    ins[3].value.trim(),
                pm:    ins[4].value.trim(),
                night: ins[5].value.trim(),
            };
        }).filter(item => item.excel);

        // ── 檢查自定義 sys 代號是否已在 HR 清單中 ──────────────────
        const hrCodeSet = new Set(newHr.map(x => x.code));
        const missingSysCodes = [...new Set(
            newCustom.map(x => x.sys).filter(s => s && !hrCodeSet.has(s))
        )];

        if (missingSysCodes.length > 0) {
            // 自動在 HR 表格末端新增缺漏的代號（時間留空）
            missingSysCodes.forEach(code => addHrRow({ code, start: null, end: null }));

            // 展開 HR 區塊
            const hrCollapsible = document.getElementsByClassName("collapsible")[0];
            if (hrCollapsible.nextElementSibling.style.display !== "block") {
                hrCollapsible.classList.add("active");
                hrCollapsible.nextElementSibling.style.display = "block";
            }
            updateWindowHeight();

            // 捲動到 HR 區塊
            setTimeout(() => {
                hrCollapsible.scrollIntoView({ behavior: 'smooth', block: 'start' });
                // 再捲到最後一筆新增的列（讓使用者看到新增的代號）
                const lastHrRow = hrBody.querySelector('tr:last-child');
                if (lastHrRow) lastHrRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }, 150);

            // 提示橫幅（若已有舊橫幅則移除）
            const oldBanner = document.getElementById('missing-hr-banner');
            if (oldBanner) oldBanner.remove();
            const banner = document.createElement('div');
            banner.id = 'missing-hr-banner';
            banner.className = 'error-banner';
            banner.innerHTML = `⚠️ 以下自定義班別的「系統代號」在 HR 清單中尚未建立：<b>${missingSysCodes.join('、')}</b>。<br>已自動新增至上方 HR 清單，請填寫上下班時間後再儲存。`;
            document.querySelector('.main-container').insertBefore(banner, document.querySelector('.scroll-area'));

            return; // 阻擋儲存
        }

        await chrome.storage.local.set({ hrShifts: newHr, shiftDict: newCustom });

        window.close();

        const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
        <style>
            body { width:300px; margin:0; display:flex; flex-direction:column;
                   align-items:center; justify-content:center; gap:16px; padding:30px 20px;
                   box-sizing:border-box; font-family:"Microsoft JhengHei",sans-serif; background:#fff; }
            .msg { font-size:15px; color:#2c3e50; font-weight:bold; text-align:center; line-height:1.6; }
            button { width:100%; padding:10px; background:#27ae60; color:white; border:none;
                     border-radius:6px; font-size:14px; font-weight:bold; cursor:pointer; }
            button:hover { background:#219150; }
        </style></head><body>
        <div class="msg">✅ 班別字典已更新！<br>📂 請重新載入 Excel 檔案。</div>
        <button onclick="window.close()">確定</button>
        </body></html>`;
        chrome.windows.create({
            url: 'data:text/html;charset=utf-8,' + encodeURIComponent(html),
            type: 'popup', width: 320, height: 180, focused: true
        });
    };

    function updateWindowHeight() {
        const targetHeight = Math.min(document.body.scrollHeight + 60, window.screen.availHeight * 0.85);
        chrome.windows.getCurrent((win) => {
            chrome.windows.update(win.id, { height: Math.round(targetHeight) });
        });
    }
    setTimeout(updateWindowHeight, 200);
});
