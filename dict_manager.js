// dict_manager.js — 使用 shared.js 的 STORAGE_KEYS / createPopupWindow
document.addEventListener('DOMContentLoaded', async () => {
    const hrBody     = document.getElementById('hrBody');
    const customBody = document.getElementById('customBody');

    const SKIP_SHIFT_CODES = new Set(['FF', 'WW', 'NH', 'N+', 'W+']);

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
        const start  = item.start || '';
        const end    = item.end   || '';
        const isSkip = SKIP_SHIFT_CODES.has(code);

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" class="hr-code" value="${code}" maxlength="5" placeholder="代號"></td>
            <td><input type="time" class="hr-start${isSkip ? ' skip-shift' : ''}" value="${isSkip ? '' : start}" ${isSkip ? 'disabled title="休假/加班類，不參與接班檢測"' : ''}></td>
            <td><input type="time" class="hr-end${isSkip ? ' skip-shift' : ''}"   value="${isSkip ? '' : end}"   ${isSkip ? 'disabled title="休假/加班類，不參與接班檢測"' : ''}></td>
            <td class="${isSkip ? 'skip-label' : ''}">${isSkip ? '跳過接班檢測' : ''}</td>
            <td><button class="del-btn">刪</button></td>
        `;

        const codeInput  = tr.querySelector('.hr-code');
        const startInput = tr.querySelector('.hr-start');
        const endInput   = tr.querySelector('.hr-end');
        const noteCell   = tr.querySelector('td:nth-child(4)');

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
        });

        tr.querySelector('.del-btn').onclick = () => tr.remove();
        hrBody.appendChild(tr);
    }

    // ── 自定義班別列 ───────────────────────────────────────────────
    function addCustomRow(item = { excel: '', sys: '', over: '', am: '', pm: '', night: '' }) {
        const isOverEnabled = (item.sys === 'N+' || item.sys === 'W+');
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" value="${item.excel || ''}" placeholder="Excel代號"></td>
            <td><input type="text" class="sys-input" value="${item.sys || ''}" placeholder="系統代號 *"></td>
            <td><input type="text" class="over-input" value="${isOverEnabled ? (item.over || '') : ''}" ${isOverEnabled ? '' : 'disabled title="僅 N+ / W+ 班別需填寫逾時"'}></td>
            <td><input type="text" value="${item.am    || ''}"></td>
            <td><input type="text" value="${item.pm    || ''}"></td>
            <td><input type="text" value="${item.night || ''}"></td>
            <td><button class="del-btn">刪</button></td>
        `;
        const sysInput  = tr.querySelector('.sys-input');
        const overInput = tr.querySelector('.over-input');

        function updateOverState() {
            const sys = sysInput.value.trim().toUpperCase();
            const allow = (sys === 'N+' || sys === 'W+');
            overInput.disabled = !allow;
            overInput.classList.toggle('skip-shift', !allow);
            if (!allow) overInput.value = '';
        }

        sysInput.addEventListener('input', () => {
            sysInput.classList.toggle('sys-empty', sysInput.value.trim() === '');
            updateOverState();
        });

        if (!item.sys || item.sys.trim() === '') sysInput.classList.add('sys-empty');
        if (!isOverEnabled) overInput.classList.add('skip-shift');

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
                message:  '欄位填寫不完整：\n1. 系統代號不可空白\n2. <b>W+ 或 N+ 班別必須填寫「逾時」欄位</b>',
                btnColor: '#e74c3c',
                width: 340, height: 240,
            });
            return;
        }

        // 收集 HR 班別
        const newHr = Array.from(hrBody.querySelectorAll('tr')).map(tr => ({
            code:  tr.querySelector('.hr-code')?.value.trim()  || '',
            start: tr.querySelector('.hr-start')?.value.trim() || null,
            end:   tr.querySelector('.hr-end')?.value.trim()   || null,
        })).map(item => ({ ...item, start: item.start || null, end: item.end || null }))
           .filter(item => item.code);

        // 收集自定義班別
        const newCustom = Array.from(customBody.querySelectorAll('tr')).map(tr => {
            const ins = tr.querySelectorAll('input');
            return {
                excel: ins[0].value.trim(), sys:   ins[1].value.trim(),
                over:  ins[2].value.trim(), am:    ins[3].value.trim(),
                pm:    ins[4].value.trim(), night: ins[5].value.trim(),
            };
        }).filter(item => item.excel);

        // 檢查自定義 sys 是否已在 HR 清單中
        const hrCodeSet       = new Set(newHr.map(x => x.code));
        const missingSysCodes = [...new Set(newCustom.map(x => x.sys).filter(s => s && !hrCodeSet.has(s)))];

        if (missingSysCodes.length > 0) {
            missingSysCodes.forEach(code => addHrRow({ code, start: null, end: null }));
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
            banner.innerHTML = `⚠️ 以下自定義班別的「系統代號」在 HR 清單中尚未建立：<b>${missingSysCodes.join('、')}</b>。<br>已自動新增至上方 HR 清單，請填寫上下班時間後再儲存。`;
            document.querySelector('.main-container').insertBefore(banner, document.querySelector('.scroll-area'));
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