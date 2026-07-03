// dict_manager.js — 使用 shared.js 的 STORAGE_KEYS / createPopupWindow
document.addEventListener('DOMContentLoaded', async () => {
    const hrBody     = document.getElementById('hrBody');
    const customBody = document.getElementById('customBody');

    const SKIP_SHIFT_CODES = new Set(['FF', 'WW', 'NH', 'N+', 'W+']);

    // ── 取得目前 HR 表格中「即時」的代號集合（大小寫不敏感），用來驗證「逾時」欄位 ──
    function getCurrentHrCodesUpper() {
        return new Set(
            Array.from(hrBody.querySelectorAll('.hr-code'))
                .map(inp => inp.value.trim().toUpperCase())
                .filter(Boolean)
        );
    }

    // 重新驗證單一列「逾時」欄位是否為 HR 清單中已存在的代號
    function revalidateOverField(sysInput, overInput) {
        if (overInput.disabled) { overInput.classList.remove('sys-empty'); return; }
        const sys      = sysInput.value.trim().toUpperCase();
        const allow    = (sys === 'N+' || sys === 'W+');
        const overVal  = overInput.value.trim().toUpperCase();
        const hrCodes  = getCurrentHrCodesUpper();
        const invalid  = allow && (overVal === '' || !hrCodes.has(overVal));
        overInput.classList.toggle('sys-empty', invalid);
    }

    // HR 表格任何代號變動（新增/修改/刪除）都要重新檢查所有自定義班別列的「逾時」欄位
    function revalidateAllOverFields() {
        customBody.querySelectorAll('tr').forEach(tr => {
            const sysInput  = tr.querySelector('.sys-input');
            const overInput = tr.querySelector('.over-input');
            if (sysInput && overInput) revalidateOverField(sysInput, overInput);
        });
    }
    // 事件委派：HR 代號輸入框是動態產生的，用委派監聽即可涵蓋所有列
    hrBody.addEventListener('input', revalidateAllOverFields);

    // 將時間字串的「分」校正為僅 00 或 30（0-14→00, 15-44→30, 45-59→進位到下一個小時的 00）
    function snapToHalfHour(timeStr) {
        if (!timeStr) return timeStr;
        const m = /^(\d{2}):(\d{2})/.exec(timeStr);
        if (!m) return timeStr;
        let hh = parseInt(m[1], 10);
        let mm = parseInt(m[2], 10);
        if (mm < 15)      mm = 0;
        else if (mm < 45) mm = 30;
        else { mm = 0; hh = (hh + 1) % 24; }
        return `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
    }

    // 為時間輸入框套用「分」校正：使用者輸入/選擇後（change 觸發），自動吸附到最近的 00 或 30 分
    function enforceHalfHourStep(input) {
        input.addEventListener('change', () => {
            if (!input.value) return;
            const snapped = snapToHalfHour(input.value);
            if (snapped !== input.value) input.value = snapped;
        });
    }

    // 即時將輸入內容強制轉為大寫（保留游標位置），用於各類「系統代號」欄位
    function forceUppercase(input) {
        input.addEventListener('input', () => {
            const pos   = input.selectionStart;
            const upper = input.value.toUpperCase();
            if (upper !== input.value) {
                input.value = upper;
                input.setSelectionRange(pos, pos);
            }
        });
    }

    // 摺疊面板控制
    Array.from(document.getElementsByClassName("collapsible")).forEach(btn => {
        btn.addEventListener("click", function () {
            this.classList.toggle("active");
            const content = this.nextElementSibling;
            content.style.display = (content.style.display === "block") ? "none" : "block";
            updateWindowHeight();
        });
    });

    const data = await chrome.storage.local.get([
        STORAGE_KEYS.HR_SHIFTS, STORAGE_KEYS.SHIFT_DICT, STORAGE_KEYS.PENDING_UNKNOWN,
    ]);

    const hrShifts     = data[STORAGE_KEYS.HR_SHIFTS]      || [];
    const customShifts = data[STORAGE_KEYS.SHIFT_DICT]     || [];
    const pendingCodes = data[STORAGE_KEYS.PENDING_UNKNOWN] || [];

    hrShifts.forEach(item => addHrRow(item));
    customShifts.forEach(item => addCustomRow(item));

    if (pendingCodes.length > 0) {
        chrome.storage.local.remove(STORAGE_KEYS.PENDING_UNKNOWN);
        const customCollapsible = document.getElementsByClassName("collapsible")[1];
        customCollapsible.classList.add("active");
        customCollapsible.nextElementSibling.style.display = "block";
        pendingCodes.forEach(code => addCustomRow({ excel: code, sys: '', over: '', am: '', pm: '', night: '' }));

        const banner = document.createElement('div');
        banner.className = 'error-banner';
        banner.id        = 'unknown-banner';
        banner.innerHTML = `⚠️ 發現 <b>${pendingCodes.length}</b> 個未知班別（<b>${pendingCodes.join('、')}</b>）已自動加入下方，請填寫「系統」欄後儲存，再重新匯入 Excel。`;
        document.querySelector('.main-container').insertBefore(banner, document.querySelector('.scroll-area'));
        setTimeout(() => customCollapsible.scrollIntoView({ behavior: 'smooth', block: 'start' }), 200);
        updateWindowHeight();
    }

    document.getElementById('addHrRow').onclick     = () => addHrRow({ code: '', start: null, end: null });
    document.getElementById('addCustomRow').onclick = () => addCustomRow();

    // ── HR 班別列 ──────────────────────────────────────────────────
    function addHrRow(item = { code: '', start: null, end: null }) {
        if (typeof item === 'string') item = { code: item, start: null, end: null };
        const code   = item.code  || '';
        const start  = snapToHalfHour(item.start || '');
        const end    = snapToHalfHour(item.end   || '');
        const isSkip = SKIP_SHIFT_CODES.has(code);

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" class="hr-code" value="${code.toUpperCase()}" maxlength="5" placeholder="代號"></td>
            <td><input type="time" step="1800" class="hr-start${isSkip ? ' skip-shift' : ''}" value="${isSkip ? '' : start}" ${isSkip ? 'disabled title="休假/加班類，不參與接班檢測"' : ''}></td>
            <td><input type="time" step="1800" class="hr-end${isSkip ? ' skip-shift' : ''}"   value="${isSkip ? '' : end}"   ${isSkip ? 'disabled title="休假/加班類，不參與接班檢測"' : ''}></td>
            <td class="${isSkip ? 'skip-label' : ''}">${isSkip ? '跳過接班檢測' : ''}</td>
            <td><button class="del-btn">刪</button></td>
        `;

        const codeInput  = tr.querySelector('.hr-code');
        const startInput = tr.querySelector('.hr-start');
        const endInput   = tr.querySelector('.hr-end');
        const noteCell   = tr.querySelector('td:nth-child(4)');
        forceUppercase(codeInput);

        // 即時反白：非跳過類代號（有填代號但不在 SKIP_SHIFT_CODES 中）若上下班時間空白，立即標紅
        function revalidateHrRow() {
            const curCode = codeInput.value.trim().toUpperCase();
            const isSkipNow = !curCode || SKIP_SHIFT_CODES.has(curCode);
            startInput.classList.toggle('sys-empty', !isSkipNow && startInput.value.trim() === '');
            endInput.classList.toggle('sys-empty',   !isSkipNow && endInput.value.trim()   === '');
        }

        codeInput.addEventListener('input', () => {
            const newCode   = codeInput.value.trim().toUpperCase();
            const newIsSkip = SKIP_SHIFT_CODES.has(newCode);
            startInput.disabled = newIsSkip;
            endInput.disabled   = newIsSkip;
            startInput.classList.toggle('skip-shift', newIsSkip);
            endInput.classList.toggle('skip-shift',   newIsSkip);
            if (newIsSkip) {
                startInput.value     = '';
                endInput.value       = '';
                noteCell.textContent = '跳過接班檢測';
                noteCell.className   = 'skip-label';
            } else {
                noteCell.textContent = '';
                noteCell.className   = '';
            }
            revalidateHrRow();
        });
        startInput.addEventListener('input', revalidateHrRow);
        endInput.addEventListener('input', revalidateHrRow);
        enforceHalfHourStep(startInput);
        enforceHalfHourStep(endInput);
        revalidateHrRow(); // 列建立當下（含從 storage 載入、或自動新增缺漏代號時）就先套用一次

        tr.querySelector('.del-btn').onclick = () => { tr.remove(); revalidateAllOverFields(); };
        hrBody.appendChild(tr);
    }

    // ── 自定義班別列 ───────────────────────────────────────────────
    function addCustomRow(item = { excel: '', sys: '', over: '', am: '', pm: '', night: '' }) {
        const isOverEnabled = (item.sys === 'N+' || item.sys === 'W+');
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" value="${(item.excel || '').toUpperCase()}" placeholder="Excel代號"></td>
            <td><input type="text" class="sys-input" value="${(item.sys || '').toUpperCase()}" placeholder="系統代號 *"></td>
            <td><input type="text" class="over-input" value="${isOverEnabled ? (item.over || '').toUpperCase() : ''}" ${isOverEnabled ? '' : 'disabled title="僅 N+ / W+ 班別需填寫逾時（且需為HR清單中已登記的代號）"'}></td>
            <td><input type="text" value="${item.am    || ''}"></td>
            <td><input type="text" value="${item.pm    || ''}"></td>
            <td><input type="text" value="${item.night || ''}"></td>
            <td><button class="del-btn">刪</button></td>
        `;
        const sysInput  = tr.querySelector('.sys-input');
        const overInput = tr.querySelector('.over-input');
        const excelInput = tr.querySelector('td:first-child input');
        forceUppercase(excelInput);
        forceUppercase(sysInput);
        forceUppercase(overInput);

        function updateOverState() {
            const sys = sysInput.value.trim().toUpperCase();
            const allow = (sys === 'N+' || sys === 'W+');
            overInput.disabled = !allow;
            overInput.classList.toggle('skip-shift', !allow);
            if (!allow) { overInput.value = ''; overInput.classList.remove('sys-empty'); }
            // 即時反白：系統代號為 N+ / W+ 時，逾時欄位必須是 HR 清單中已存在的代號，
            // 空白或查無此代號都立即標紅提醒，不用等儲存才檢查
            revalidateOverField(sysInput, overInput);
        }

        sysInput.addEventListener('input', () => {
            sysInput.classList.toggle('sys-empty', sysInput.value.trim() === '');
            updateOverState();
        });

        overInput.addEventListener('input', () => revalidateOverField(sysInput, overInput));

        if (!item.sys || item.sys.trim() === '') sysInput.classList.add('sys-empty');
        if (!isOverEnabled) {
            overInput.classList.add('skip-shift');
        } else {
            revalidateOverField(sysInput, overInput);
        }

        tr.querySelector('.del-btn').onclick = () => tr.remove();
        customBody.appendChild(tr);
    }

    // ── 儲存 ───────────────────────────────────────────────────────
    document.getElementById('saveAll').onclick = async () => {

        // 驗證：sys 不可為空，W+/N+ 必須填逾時
        const badRows = Array.from(customBody.querySelectorAll('tr')).filter(tr => {
            const inputs = tr.querySelectorAll('input');
            const excel  = inputs[0]?.value.trim();
            const sys    = inputs[1]?.value.trim().toUpperCase();
            const over   = inputs[2]?.value.trim();
            const isSysMissing  = excel && !sys;
            const isOverMissing = (sys === 'W+' || sys === 'N+') && !over;
            inputs[1].classList.toggle('sys-empty', isSysMissing);
            inputs[2].classList.toggle('sys-empty', isOverMissing);
            return isSysMissing || isOverMissing;
        });

        if (badRows.length > 0) {
            const customCollapsible = document.getElementsByClassName("collapsible")[1];
            if (customCollapsible.nextElementSibling.style.display !== "block") {
                customCollapsible.classList.add("active");
                customCollapsible.nextElementSibling.style.display = "block";
                updateWindowHeight();
            }
            badRows[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
            createPopupWindow({
                title:    '❌ 儲存失敗',
                message:  '欄位填寫不完整：\n1. 系統代號不可空白\n2. <b>W+ 或 N+ 班別必須填寫「逾時」欄位</b>（逾時代號需為 HR 清單中已登記的代號，若尚未建立，儲存後會自動新增，請補上上下班時間）',
                btnColor: '#e74c3c',
                width: 340, height: 240,
            });
            return;
        }

        // 收集 HR 班別
        const newHr = Array.from(hrBody.querySelectorAll('tr')).map(tr => ({
            code:  tr.querySelector('.hr-code')?.value.trim().toUpperCase() || '',
            start: tr.querySelector('.hr-start')?.value.trim() || null,
            end:   tr.querySelector('.hr-end')?.value.trim()   || null,
        })).map(item => ({ ...item, start: item.start || null, end: item.end || null }))
           .filter(item => item.code);

        // 收集自定義班別
        const newCustom = Array.from(customBody.querySelectorAll('tr')).map(tr => {
            const ins = tr.querySelectorAll('input');
            return {
                excel: ins[0].value.trim().toUpperCase(), sys: ins[1].value.trim().toUpperCase(),
                over:  ins[2].value.trim().toUpperCase(), am:    ins[3].value.trim(),
                pm:    ins[4].value.trim(), night: ins[5].value.trim(),
            };
        }).filter(item => item.excel);

        // 檢查自定義 sys 是否已在 HR 清單中（原有規則，大小寫比對方式不變）
        const hrCodeSet      = new Set(newHr.map(x => x.code));
        const missingSysCodes = [...new Set(newCustom.map(x => x.sys).filter(s => s && !hrCodeSet.has(s)))];

        // 檢查 N+/W+ 的「逾時」代號是否已在 HR 清單中（大小寫不敏感，比對規則同即時反白）
        const hrCodeSetUpper   = new Set(newHr.map(x => String(x.code || '').trim().toUpperCase()));
        const missingOverCodes = [...new Set(
            newCustom
                .filter(x => x.sys.trim().toUpperCase() === 'N+' || x.sys.trim().toUpperCase() === 'W+')
                .map(x => x.over)
                .filter(o => o && !hrCodeSetUpper.has(o.trim().toUpperCase()))
        )];

        const missingCodes = [...new Set([...missingSysCodes, ...missingOverCodes])];

        if (missingCodes.length > 0) {
            missingCodes.forEach(code => addHrRow({ code, start: null, end: null }));
            const hrCollapsible = document.getElementsByClassName("collapsible")[0];
            if (hrCollapsible.nextElementSibling.style.display !== "block") {
                hrCollapsible.classList.add("active");
                hrCollapsible.nextElementSibling.style.display = "block";
            }
            updateWindowHeight();
            setTimeout(() => {
                hrCollapsible.scrollIntoView({ behavior: 'smooth', block: 'start' });
                const lastHrRow = hrBody.querySelector('tr:last-child');
                if (lastHrRow) lastHrRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }, 150);

            const oldBanner = document.getElementById('missing-hr-banner');
            if (oldBanner) oldBanner.remove();
            const banner = document.createElement('div');
            banner.id        = 'missing-hr-banner';
            banner.className = 'error-banner';
            banner.innerHTML = `⚠️ 以下代號（來自「系統」或「逾時」欄位）在 HR 清單中尚未建立：<b>${missingCodes.join('、')}</b>。<br>已自動新增至上方 HR 清單，請填寫上下班時間後再儲存。`;
            document.querySelector('.main-container').insertBefore(banner, document.querySelector('.scroll-area'));
            return;
        }

        // 檢查 HR 內建班別（非休假/加班類）是否都已填寫完整的上下班時間，
        // 沒填時間的代號無法用來計算接班間隔，資料不完整不可存檔。
        const incompleteHrRows = Array.from(hrBody.querySelectorAll('tr')).filter(tr => {
            const codeInput  = tr.querySelector('.hr-code');
            const startInput = tr.querySelector('.hr-start');
            const endInput   = tr.querySelector('.hr-end');
            const code = codeInput?.value.trim().toUpperCase() || '';
            if (!code || SKIP_SHIFT_CODES.has(code)) return false; // 空白列或休假/加班類不需要時間
            const start = startInput?.value.trim() || '';
            const end   = endInput?.value.trim()   || '';
            startInput.classList.toggle('sys-empty', !start);
            endInput.classList.toggle('sys-empty',   !end);
            return !start || !end;
        });

        if (incompleteHrRows.length > 0) {
            const hrCollapsible = document.getElementsByClassName("collapsible")[0];
            if (hrCollapsible.nextElementSibling.style.display !== "block") {
                hrCollapsible.classList.add("active");
                hrCollapsible.nextElementSibling.style.display = "block";
                updateWindowHeight();
            }
            incompleteHrRows[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
            const codes = incompleteHrRows
                .map(tr => tr.querySelector('.hr-code')?.value.trim())
                .filter(Boolean);
            createPopupWindow({
                title:    '❌ 儲存失敗',
                message:  `以下 HR 班別尚未填寫完整的上下班時間：\n<b>${codes.join('、') || '（有代號欄未命名）'}</b>\n請填寫後再儲存。`,
                btnColor: '#e74c3c',
                width: 360, height: 220,
            });
            return;
        }

        await chrome.storage.local.set({
            [STORAGE_KEYS.HR_SHIFTS]:  newHr,
            [STORAGE_KEYS.SHIFT_DICT]: newCustom,
        });
        window.close();

        createPopupWindow({
            message:  '✅ 班別字典已更新！\n📂 請重新載入 Excel 檔案。',
            btnColor: '#27ae60',
            width: 320, height: 180,
        });
    };

    function updateWindowHeight() {
        const targetHeight = Math.min(document.body.scrollHeight + 60, window.screen.availHeight * 0.85);
        chrome.windows.getCurrent(win => chrome.windows.update(win.id, { height: Math.round(targetHeight) }));
    }
    setTimeout(updateWindowHeight, 200);
});