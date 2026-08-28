// dict_manager.js — 使用 shared.js 的 STORAGE_KEYS / createPopupWindow
document.addEventListener('DOMContentLoaded', async () => {
    const hrBody     = document.getElementById('hrBody');
    const customBody = document.getElementById('customBody');

    const SKIP_SHIFT_CODES = new Set(['FF', 'WW', 'NH', 'N+', 'W+']); // 不需時間的代號
    const DEFAULT_LEAVE_CODES = new Set(['FF', 'NH', 'SS', 'WW', 'VV']); // 預設非上班代號

    // ── 取得目前 HR 表格中「即時」的代號集合（大小寫不敏感） ──
    function getCurrentHrCodes() {
        return Array.from(hrBody.querySelectorAll('.hr-code'))
            .map(inp => inp.value.trim().toUpperCase())
            .filter(Boolean);
    }

    function getCurrentHrCodesUpper() {
        return new Set(getCurrentHrCodes());
    }

    // 將目前 HR 內建班別的系統代號同步到「逾時」下拉選單。
    // 若舊資料的逾時代號已不在 HR 清單中，暫時保留並標示，避免重新開啟頁面時遺失設定。
    function syncOverOptions(overSelect, preferredValue = overSelect.value) {
        if (!overSelect || overSelect.tagName !== 'SELECT') return;

        const selected = String(preferredValue || '').trim().toUpperCase();
        const hrCodes = [...new Set(getCurrentHrCodes())];
        overSelect.replaceChildren();

        const placeholder = document.createElement('option');
        placeholder.value = '';
        placeholder.textContent = '請選擇 HR 系統代號';
        overSelect.appendChild(placeholder);

        hrCodes.forEach(code => {
            const option = document.createElement('option');
            option.value = code;
            option.textContent = code;
            overSelect.appendChild(option);
        });

        // 相容既有資料：值不在目前 HR 清單時保留，並讓驗證提示使用者修正。
        if (selected && !hrCodes.includes(selected)) {
            const missingOption = document.createElement('option');
            missingOption.value = selected;
            missingOption.textContent = `${selected}（未在 HR 清單）`;
            missingOption.dataset.missing = 'true';
            overSelect.appendChild(missingOption);
        }

        overSelect.value = selected;
        if (overSelect.value !== selected) overSelect.value = '';
    }

    // 重新驗證單一列「逾時」欄位是否為 HR 清單中已存在的代號
    function revalidateOverField(sysInput, overInput) {
        syncOverOptions(overInput);
        if (overInput.disabled) { overInput.classList.remove('sys-empty'); return; }
        const sys      = sysInput.value.trim().toUpperCase();
        const allow    = (sys === 'N+' || sys === 'W+');
        const overVal  = overInput.value.trim().toUpperCase();
        const hrCodes  = getCurrentHrCodesUpper();
        const invalid  = allow && (overVal === '' || !hrCodes.has(overVal));
        overInput.classList.toggle('sys-empty', invalid);
    }

    // HR 表格任何代號變動（新增/修改/刪除）都要重新檢查所有自定義班別列的「逾時」欄位與重複性
    function revalidateAllOverFields() {
        customBody.querySelectorAll('tr').forEach(tr => {
            const sysInput  = tr.querySelector('.sys-input');
            const overInput = tr.querySelector('.over-input');
            if (sysInput && overInput) revalidateOverField(sysInput, overInput);
        });
        checkDuplicateHrCodes();
    }
    // 事件委派：HR 代號輸入框是動態產生的，用委派監聽即可涵蓋所有列
    hrBody.addEventListener('input', revalidateAllOverFields);

    // ── 重複代號即時檢查機制 ─────────────────────────────────────
    function checkDuplicateHrCodes() {
        const inputs = Array.from(hrBody.querySelectorAll('.hr-code'));
        const counts = {};
        inputs.forEach(inp => {
            const val = inp.value.trim().toUpperCase();
            if (val) counts[val] = (counts[val] || 0) + 1;
        });

        inputs.forEach(inp => {
            const val = inp.value.trim().toUpperCase();
            const isDup = val && counts[val] > 1;
            inp.classList.toggle('sys-empty', isDup);
            if (isDup) inp.title = '系統代號不可重複！';
            else inp.removeAttribute('title');
        });
    }

    function checkDuplicateCustomCodes() {
        const inputs = Array.from(customBody.querySelectorAll('.excel-input'));
        const counts = {};
        inputs.forEach(inp => {
            const val = inp.value.trim().toUpperCase();
            if (val) counts[val] = (counts[val] || 0) + 1;
        });

        inputs.forEach(inp => {
            const val = inp.value.trim().toUpperCase();
            const isDup = val && counts[val] > 1;
            inp.classList.toggle('sys-empty', isDup);
            if (isDup) inp.title = 'Excel 代號不可重複！';
            else inp.removeAttribute('title');
        });
    }

    customBody.addEventListener('input', checkDuplicateCustomCodes);

    // ── 時間計算與驗證輔助函式 (4-12小時、跨日判斷) ───────────────
    function validateShiftTime(startStr, endStr) {
        if (!startStr || !endStr) return { valid: true }; // 尚未填寫完整由必填機制處理
        
        const [sH, sM] = startStr.split(':').map(Number);
        const [eH, eM] = endStr.split(':').map(Number);
        
        let startMin = sH * 60 + sM;
        let endMin = eH * 60 + eM;
        let isOvernight = false;

        if (endMin < startMin) {
            endMin += 24 * 60; // 跨日 +24小時
            isOvernight = true;
        }

        const duration = (endMin - startMin) / 60;

        if (duration < 4 || duration > 12) {
            return {
                valid: false,
                msg: `工作時間必須介於 4 至 12 小時之間（目前：${duration} 小時）`
            };
        }

        return { valid: true, duration, isOvernight };
    }

    // 將時間字串的「分」校正為僅 00 或 30
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

    function normalizeHalfHourValue(input) {
        if (!input || !input.value) return false;
        const snapped = snapToHalfHour(input.value);
        if (snapped !== input.value) {
            input.value = snapped;
            return true;
        }
        return false;
    }

    // 保留原本單一時間欄位的外觀，但改用緊湊的自訂面板，避免瀏覽器原生
    // time picker 顯示 00～59 的分鐘清單。面板中的分鐘選項只有 00、30。
    function createHalfHourPicker(input, initialValue = '') {
        const parent = input.parentNode;
        const wrapper = document.createElement('span');
        wrapper.className = 'time-picker-wrap';

        input.type = 'text';
        input.readOnly = true;
        input.setAttribute('inputmode', 'none');
        input.setAttribute('aria-haspopup', 'dialog');
        input.value = initialValue ? snapToHalfHour(initialValue) : '';

        const panel = document.createElement('div');
        panel.className = 'time-picker-panel';
        panel.setAttribute('role', 'dialog');
        panel.setAttribute('aria-label', '選擇時間');

        const hourSelect = document.createElement('select');
        hourSelect.className = 'time-hour';
        hourSelect.setAttribute('aria-label', '小時');
        for (let hour = 0; hour < 24; hour++) {
            const option = document.createElement('option');
            option.value = String(hour).padStart(2, '0');
            option.textContent = option.value;
            hourSelect.appendChild(option);
        }

        const separator = document.createElement('span');
        separator.textContent = ':';
        separator.setAttribute('aria-hidden', 'true');

        const minuteSelect = document.createElement('select');
        minuteSelect.className = 'time-minute';
        minuteSelect.setAttribute('aria-label', '分鐘');
        ['00', '30'].forEach(minute => {
            const option = document.createElement('option');
            option.value = minute;
            option.textContent = minute;
            minuteSelect.appendChild(option);
        });

        panel.append(hourSelect, separator, minuteSelect);
        wrapper.appendChild(panel);

        function syncPickerValue() {
            const [hour = '', minute = ''] = String(input.value || '').split(':');
            hourSelect.value = /^\d{2}$/.test(hour) ? hour : '00';
            minuteSelect.value = minute === '30' ? '30' : '00';
        }

        function close() {
            panel.classList.remove('open');
        }

        function applyPickerValue() {
            input.value = `${hourSelect.value}:${minuteSelect.value}`;
            // 保持面板開啟，使用者可連續選擇小時與分鐘；點擊外部才關閉。
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
        }

        input.addEventListener('click', event => {
            event.stopPropagation();
            if (input.disabled) return;
            document.querySelectorAll('.time-picker-panel.open').forEach(other => {
                if (other !== panel) other.classList.remove('open');
            });
            syncPickerValue();
            panel.classList.add('open');
        });
        panel.addEventListener('click', event => event.stopPropagation());
        document.addEventListener('click', close);
        hourSelect.addEventListener('change', applyPickerValue);
        minuteSelect.addEventListener('change', applyPickerValue);

        if (parent) parent.replaceChild(wrapper, input);
        wrapper.insertBefore(input, panel);

        const setDisabled = disabled => {
            input.disabled = disabled;
            wrapper.classList.toggle('disabled', disabled);
            if (disabled) close();
        };
        setDisabled(input.disabled);

        return { input, setDisabled, close };
    }

    function enforceHalfHourStep(input) {
        const normalize = () => normalizeHalfHourValue(input);
        // 自訂選擇器與程式寫入都經過 input/change，值只會保留 00／30 分鐘。
        input.addEventListener('input', normalize);
        input.addEventListener('change', normalize);
    }

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
        STORAGE_KEYS.HR_SHIFTS, STORAGE_KEYS.SHIFT_DICT, STORAGE_KEYS.PENDING_UNKNOWN, STORAGE_KEYS.PENDING_OVERTIME_GAPS,
    ]);

    const hrShifts      = data[STORAGE_KEYS.HR_SHIFTS]      || [];
    const customShifts  = data[STORAGE_KEYS.SHIFT_DICT]     || [];
    const pendingCodes  = data[STORAGE_KEYS.PENDING_UNKNOWN] || [];
    const pendingGaps   = data[STORAGE_KEYS.PENDING_OVERTIME_GAPS] || [];

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

    if (pendingGaps.length > 0) {
        chrome.storage.local.remove(STORAGE_KEYS.PENDING_OVERTIME_GAPS);
        const customCollapsible = document.getElementsByClassName("collapsible")[1];
        customCollapsible.classList.add("active");
        customCollapsible.nextElementSibling.style.display = "block";
        pendingGaps.forEach(g => addCustomRow({
            excel: '', sys: g.sys, over: g.over, am: '', pm: '', night: '',
            overtimeFor: g.originExcel || '',
        }));

        const gapLabels = pendingGaps.map(g => `${g.originExcel ? `${g.originExcel}：` : ''}${g.over}→${g.sys}`).join('、');
        const banner = document.createElement('div');
        banner.className = 'error-banner';
        banner.id        = 'overtime-gap-banner';
        banner.innerHTML = `⚠️ 一鍵配置發現 <b>${pendingGaps.length}</b> 組缺少加班代號對應（<b>${gapLabels}</b>）已自動加入下方，請補上「Excel」欄的加班代號後儲存，再重新執行一鍵配置。`;
        document.querySelector('.main-container').insertBefore(banner, document.querySelector('.scroll-area'));
        setTimeout(() => customCollapsible.scrollIntoView({ behavior: 'smooth', block: 'start' }), 200);
        updateWindowHeight();
    }

    document.getElementById('addHrRow').onclick     = () => addHrRow({ code: '', start: null, end: null });
    document.getElementById('addCustomRow').onclick = () => addCustomRow();

    // ── HR 班別列 ──────────────────────────────────────────────────
    function addHrRow(item = { code: '', start: null, end: null, isLeave: undefined }) {
        if (typeof item === 'string') item = { code: item, start: null, end: null, isLeave: undefined };
        const code   = item.code  || '';
        const start  = snapToHalfHour(item.start || '');
        const end    = snapToHalfHour(item.end   || '');
        const isSkip = SKIP_SHIFT_CODES.has(code);
        
        let isLeave = item.isLeave;
        if (typeof isLeave === 'undefined') {
            isLeave = DEFAULT_LEAVE_CODES.has(code.toUpperCase());
        }

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="checkbox" class="hr-leave-cb" ${isLeave ? 'checked' : ''} title="勾選代表此為非上班代號（排班不會將其轉為加班）"></td>
            <td><input type="text" class="hr-code" value="${code.toUpperCase()}" maxlength="5" placeholder="代號"></td>
            <td><input type="text" class="hr-start${isSkip ? ' skip-shift' : ''}" value="${isSkip ? '' : start}" readonly ${isSkip ? 'disabled title="休假/加班類，不參與接班檢測"' : ''}></td>
            <td><input type="text" class="hr-end${isSkip ? ' skip-shift' : ''}"   value="${isSkip ? '' : end}"   readonly ${isSkip ? 'disabled title="休假/加班類，不參與接班檢測"' : ''}></td>
            <td class="${isSkip ? 'skip-label' : ''}">${isSkip ? '跳過時間' : ''}</td>
            <td><button class="del-btn">刪</button></td>
        `;

        const codeInput      = tr.querySelector('.hr-code');
        const startPicker    = createHalfHourPicker(tr.querySelector('.hr-start'), isSkip ? '' : start);
        const endPicker      = createHalfHourPicker(tr.querySelector('.hr-end'), isSkip ? '' : end);
        const startInput     = startPicker.input;
        const endInput       = endPicker.input;
        const noteCell       = tr.querySelector('td:nth-child(5)');
        forceUppercase(codeInput);

        function revalidateHrRow() {
            const curCode = codeInput.value.trim().toUpperCase();
            const isSkipNow = !curCode || SKIP_SHIFT_CODES.has(curCode);
            startInput.classList.toggle('sys-empty', !isSkipNow && startInput.value.trim() === '');
            endInput.classList.toggle('sys-empty',   !isSkipNow && endInput.value.trim()   === '');

            if (!isSkipNow && startInput.value && endInput.value) {
                const res = validateShiftTime(startInput.value, endInput.value);
                if (!res.valid) {
                    startInput.classList.add('sys-empty');
                    endInput.classList.add('sys-empty');
                    noteCell.textContent = '時間異常';
                    noteCell.className = 'skip-label';
                    noteCell.style.color = '#e74c3c';
                } else {
                    noteCell.textContent = res.isOvernight ? '跨日班別' : '';
                    noteCell.className = res.isOvernight ? 'skip-label' : '';
                    noteCell.style.color = res.isOvernight ? '#e67e22' : '';
                }
            }
            checkDuplicateHrCodes();
        }

        enforceHalfHourStep(startInput);
        enforceHalfHourStep(endInput);

        codeInput.addEventListener('input', () => {
            const newCode   = codeInput.value.trim().toUpperCase();
            const newIsSkip = SKIP_SHIFT_CODES.has(newCode);
            startPicker.setDisabled(newIsSkip);
            endPicker.setDisabled(newIsSkip);
            startInput.classList.toggle('skip-shift', newIsSkip);
            endInput.classList.toggle('skip-shift',   newIsSkip);
            if (newIsSkip) {
                startInput.value     = '';
                endInput.value       = '';
                noteCell.textContent = '跳過接班檢測';
                noteCell.className   = 'skip-label';
                noteCell.style.color = '#999';
            } else {
                noteCell.textContent = '';
                noteCell.className   = '';
            }
            revalidateHrRow();
        });
        startInput.addEventListener('input', revalidateHrRow);
        startInput.addEventListener('change', revalidateHrRow);
        endInput.addEventListener('input', revalidateHrRow);
        endInput.addEventListener('change', revalidateHrRow);
        revalidateHrRow();

        tr.querySelector('.del-btn').onclick = () => { tr.remove(); revalidateAllOverFields(); };
        hrBody.appendChild(tr);
        revalidateAllOverFields();
    }

    // ── 自定義班別列 ───────────────────────────────────────────────
    function addCustomRow(item = { excel: '', sys: '', over: '', am: '', pm: '', night: '', isLeave: false, overtimeFor: '' }) {
        const itemSys = String(item.sys || '').trim().toUpperCase();
        const isOverEnabled = (itemSys === 'N+' || itemSys === 'W+');
        const isLeave = !!item.isLeave;
        const tr = document.createElement('tr');
        const overtimeFor = String(item.overtimeFor || item.originExcel || '').trim().toUpperCase();
        if (overtimeFor) tr.dataset.overtimeFor = overtimeFor;
        tr.innerHTML = `
            <td><input type="checkbox" class="leave-cb" ${isLeave ? 'checked' : ''} title="勾選代表此 Excel 代號屬於「代表放假」的符號，其餘欄位可不填"></td>
            <td><input type="text" class="excel-input" value="${(item.excel || '').toUpperCase()}" placeholder="Excel代號"></td>
            <td><input type="text" class="sys-input" value="${itemSys}" placeholder="系統代號 *"></td>
            <td><select class="over-input" ${isOverEnabled ? '' : 'disabled title="僅 N+ / W+ 班別需選擇逾時（且需為 HR 清單中的系統代號）"'}></select></td>
            <td><input type="text" class="origin-input" value="${overtimeFor}" placeholder="原始 Excel 代號" ${isOverEnabled ? '' : 'disabled title="僅 N+ / W+ 班別需填寫對應原始 Excel 代號"'}></td>
            <td><input type="text" class="am-input"    value="${item.am    || ''}"></td>
            <td><input type="text" class="pm-input"    value="${item.pm    || ''}"></td>
            <td><input type="text" class="night-input" value="${item.night || ''}"></td>
            <td><button class="del-btn">刪</button></td>
        `;
        const leaveCb    = tr.querySelector('.leave-cb');
        const excelInput = tr.querySelector('.excel-input');
        const sysInput   = tr.querySelector('.sys-input');
        const overInput  = tr.querySelector('.over-input');
        const originInput = tr.querySelector('.origin-input');
        const amInput    = tr.querySelector('.am-input');
        const pmInput    = tr.querySelector('.pm-input');
        const nightInput = tr.querySelector('.night-input');
        forceUppercase(excelInput);
        forceUppercase(sysInput);
        forceUppercase(originInput);
        syncOverOptions(overInput, (isLeave || !isOverEnabled) ? '' : (item.over || ''));

        function updateOverState() {
            const sys = sysInput.value.trim().toUpperCase();
            const allow = (sys === 'N+' || sys === 'W+');
            syncOverOptions(overInput);
            overInput.disabled = !allow;
            originInput.disabled = !allow;
            overInput.classList.toggle('skip-shift', !allow);
            originInput.classList.toggle('skip-shift', !allow);
            if (!allow) {
                overInput.value = '';
                originInput.value = '';
                overInput.classList.remove('sys-empty');
                originInput.classList.remove('sys-empty');
            }
            revalidateOverField(sysInput, overInput);
        }

        function updateLeaveState() {
            const leave = leaveCb.checked;
            [sysInput, overInput, originInput, amInput, pmInput, nightInput].forEach(inp => {
                inp.disabled = leave;
                inp.classList.toggle('leave-disabled', leave);
            });
            if (leave) {
                sysInput.value = ''; overInput.value = ''; originInput.value = ''; amInput.value = ''; pmInput.value = ''; nightInput.value = '';
                sysInput.classList.remove('sys-empty');
                overInput.classList.remove('sys-empty', 'skip-shift');
            } else {
                sysInput.classList.toggle('sys-empty', sysInput.value.trim() === '');
                updateOverState();
            }
        }

        leaveCb.addEventListener('change', updateLeaveState);
        sysInput.addEventListener('input', () => {
            sysInput.classList.toggle('sys-empty', sysInput.value.trim() === '');
            updateOverState();
        });

        overInput.addEventListener('change', () => revalidateOverField(sysInput, overInput));
        originInput.addEventListener('input', () => {
            const allow = ['W+', 'N+'].includes(sysInput.value.trim().toUpperCase());
            originInput.classList.toggle('sys-empty', allow && originInput.value.trim() === '');
        });

        if (isLeave) {
            updateLeaveState();
        } else {
            if (!item.sys || item.sys.trim() === '') sysInput.classList.add('sys-empty');
            if (!isOverEnabled) {
                overInput.classList.add('skip-shift');
            } else {
                revalidateOverField(sysInput, overInput);
            }
        }

        tr.querySelector('.del-btn').onclick = () => { tr.remove(); checkDuplicateCustomCodes(); revalidateAllOverFields(); };
        customBody.appendChild(tr);
        checkDuplicateCustomCodes();
    }

    // ── 儲存 ───────────────────────────────────────────────────────
    document.getElementById('saveAll').onclick = async () => {

        // 儲存前再次校正，避免程式或瀏覽器直接寫入非 00／30 分鐘。
        hrBody.querySelectorAll('.hr-start, .hr-end').forEach(normalizeHalfHourValue);

        // 1. 檢查 HR 系統代號是否有重複
        const hrCodes = Array.from(hrBody.querySelectorAll('.hr-code')).map(i => i.value.trim().toUpperCase()).filter(Boolean);
        const dupHr = hrCodes.filter((c, idx) => hrCodes.indexOf(c) !== idx);
        if (dupHr.length > 0) {
            createPopupWindow({
                title:    '❌ 儲存失敗',
                message:  `HR 系統內建班別有重複代號：<b>${[...new Set(dupHr)].join('、')}</b>\n請修改重複代號後再儲存。`,
                btnColor: '#e74c3c',
                width: 360, height: 200,
            });
            return;
        }

        // 2. 檢查 Excel 代號是否有重複
        const customCodes = Array.from(customBody.querySelectorAll('.excel-input')).map(i => i.value.trim().toUpperCase()).filter(Boolean);
        const dupCustom = customCodes.filter((c, idx) => customCodes.indexOf(c) !== idx);
        if (dupCustom.length > 0) {
            createPopupWindow({
                title:    '❌ 儲存失敗',
                message:  `使用者自定義班別有重複 Excel 代號：<b>${[...new Set(dupCustom)].join('、')}</b>\n請修改重複代號後再儲存。`,
                btnColor: '#e74c3c',
                width: 360, height: 200,
            });
            return;
        }

        // 同一來源 Excel 代號對同一種加班類型只能有一筆，避免一對多而造成反查不確定。
        const overtimeSourceKeys = [];
        Array.from(customBody.querySelectorAll('tr')).forEach(tr => {
            const sys = tr.querySelector('.sys-input')?.value.trim().toUpperCase() || '';
            const source = tr.querySelector('.origin-input')?.value.trim().toUpperCase() || '';
            if (['W+', 'N+'].includes(sys) && source) overtimeSourceKeys.push(`${sys}::${source}`);
        });
        const duplicateOvertimeSources = overtimeSourceKeys.filter((key, idx) => overtimeSourceKeys.indexOf(key) !== idx);
        if (duplicateOvertimeSources.length > 0) {
            createPopupWindow({
                title: '❌ 儲存失敗',
                message: `同一原始 Excel 代號不可重複指定相同加班類型：<b>${[...new Set(duplicateOvertimeSources)].map(k => k.replace('::', ' → ')).join('、')}</b>`,
                btnColor: '#e74c3c',
                width: 380, height: 220,
            });
            return;
        }

        // 3. 驗證自定義班別必填欄位
        const badRows = Array.from(customBody.querySelectorAll('tr')).filter(tr => {
            const isLeave = tr.querySelector('.leave-cb')?.checked;
            const sysEl   = tr.querySelector('.sys-input');
            const overEl  = tr.querySelector('.over-input');
            if (isLeave) {
                sysEl.classList.remove('sys-empty');
                overEl.classList.remove('sys-empty');
                tr.querySelector('.origin-input')?.classList.remove('sys-empty');
                return false;
            }
            const excel  = tr.querySelector('.excel-input')?.value.trim();
            const sys    = sysEl?.value.trim().toUpperCase();
            const over   = overEl?.value.trim();
            const originEl = tr.querySelector('.origin-input');
            const origin = originEl?.value.trim();
            const isSysMissing  = excel && !sys;
            const isOverMissing = (sys === 'W+' || sys === 'N+') && !over;
            const isOriginMissing = (sys === 'W+' || sys === 'N+') && !origin;
            sysEl.classList.toggle('sys-empty', isSysMissing);
            overEl.classList.toggle('sys-empty', isOverMissing);
            originEl?.classList.toggle('sys-empty', isOriginMissing);
            return isSysMissing || isOverMissing || isOriginMissing;
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
                message:  '欄位填寫不完整：\n1. 系統代號不可空白\n2. <b>W+ 或 N+ 班別必須選擇「逾時」的 HR 系統代號</b>\n3. <b>W+ 或 N+ 班別必須填寫「對應 Excel 代號」</b>',
                btnColor: '#e74c3c',
                width: 340, height: 240,
            });
            return;
        }

        // 收集 HR 班別
        const newHr = Array.from(hrBody.querySelectorAll('tr')).map(tr => ({
            isLeave: tr.querySelector('.hr-leave-cb')?.checked || false,
            code:  tr.querySelector('.hr-code')?.value.trim().toUpperCase() || '',
            start: tr.querySelector('.hr-start')?.value.trim() || null,
            end:   tr.querySelector('.hr-end')?.value.trim()   || null,
        })).map(item => ({ ...item, start: item.start || null, end: item.end || null }))
           .filter(item => item.code);

        // 收集自定義班別
        const newCustom = Array.from(customBody.querySelectorAll('tr')).map(tr => {
            const isLeave = !!tr.querySelector('.leave-cb')?.checked;
            return {
                excel:   tr.querySelector('.excel-input')?.value.trim().toUpperCase() || '',
                sys:     tr.querySelector('.sys-input')?.value.trim().toUpperCase()   || '',
                over:    tr.querySelector('.over-input')?.value.trim().toUpperCase()  || '',
                // W+／N+ 專用：明確保存其對應的原始 Excel 班別，避免同 HR 代號誤配。
                // 保留相容性：若來源欄位未被保留，仍讀取自動缺漏列的 dataset。
                overtimeFor: tr.querySelector('.origin-input')?.value.trim().toUpperCase() || tr.dataset.overtimeFor || '',
                am:      tr.querySelector('.am-input')?.value.trim()    || '',
                pm:      tr.querySelector('.pm-input')?.value.trim()    || '',
                night:   tr.querySelector('.night-input')?.value.trim()   || '',
                isLeave,
            };
        }).filter(item => item.excel);

        // 檢查未登記至 HR 的代號
        const hrCodeSet      = new Set(newHr.map(x => x.code));
        const missingSysCodes = [...new Set(newCustom.map(x => x.sys).filter(s => s && !hrCodeSet.has(s)))];

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

        // 4. 檢查 HR 上下班時間完整性與 4-12 小時的時數規則
        let timeRangeError = '';
        const incompleteHrRows = Array.from(hrBody.querySelectorAll('tr')).filter(tr => {
            const codeInput  = tr.querySelector('.hr-code');
            const startInput = tr.querySelector('.hr-start');
            const endInput   = tr.querySelector('.hr-end');
            const code = codeInput?.value.trim().toUpperCase() || '';
            if (!code || SKIP_SHIFT_CODES.has(code)) return false;
            
            const start = startInput?.value.trim() || '';
            const end   = endInput?.value.trim()   || '';
            
            if (!start || !end) {
                startInput.classList.toggle('sys-empty', !start);
                endInput.classList.toggle('sys-empty',   !end);
                return true;
            }

            const timeCheck = validateShiftTime(start, end);
            if (!timeCheck.valid) {
                startInput.classList.add('sys-empty');
                endInput.classList.add('sys-empty');
                timeRangeError = `班別 <b>${code}</b> 的` + timeCheck.msg;
                return true;
            }

            return false;
        });

        if (incompleteHrRows.length > 0) {
            const hrCollapsible = document.getElementsByClassName("collapsible")[0];
            if (hrCollapsible.nextElementSibling.style.display !== "block") {
                hrCollapsible.classList.add("active");
                hrCollapsible.nextElementSibling.style.display = "block";
                updateWindowHeight();
            }
            incompleteHrRows[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
            
            const msg = timeRangeError || `以下 HR 班別尚未填寫完整的上下班時間：\n<b>${incompleteHrRows.map(tr => tr.querySelector('.hr-code')?.value.trim()).filter(Boolean).join('、')}</b>\n請填寫後再儲存。`;
            createPopupWindow({
                title:    '❌ 儲存失敗',
                message:  msg,
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