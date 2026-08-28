// options.js — 使用 shared.js 的 getValidCodes / createPopupWindow
document.addEventListener('DOMContentLoaded', async () => {
    const data = await chrome.storage.local.get([
        STORAGE_KEYS.AUTO_MODE, STORAGE_KEYS.SHOW_WEB_PREVIEW, STORAGE_KEYS.SHOW_EXCEL_REPORT,
        STORAGE_KEYS.BLANK_FILL_MODE, STORAGE_KEYS.BLANK_FILL_CODE,
        STORAGE_KEYS.HR_SHIFTS, STORAGE_KEYS.SHIFT_DICT, STORAGE_KEYS.WW_MODE,
    ]);

    document.getElementById('autoMode').checked        = data[STORAGE_KEYS.AUTO_MODE]        || false;
    document.getElementById('showWebPreview').checked  = data[STORAGE_KEYS.SHOW_WEB_PREVIEW]  === true;
    document.getElementById('showExcelReport').checked = data[STORAGE_KEYS.SHOW_EXCEL_REPORT] !== false;

    const mode      = data[STORAGE_KEYS.BLANK_FILL_MODE] || 'keep';
    const codeInput = document.getElementById('blankFillCode');
    const codeHint  = document.getElementById('fillCodeHint');
    document.querySelector(`input[name="blankFillMode"][value="${mode}"]`).checked = true;
    codeInput.value    = data[STORAGE_KEYS.BLANK_FILL_CODE] || '';
    codeInput.disabled = (mode === 'keep');

    function validateCode() {
        const val     = codeInput.value.trim();
        const saveBtn = document.getElementById('saveSettings');
        if (!val) {
            codeInput.classList.remove('valid', 'invalid');
            codeHint.textContent = '';
            codeHint.className   = 'fill-code-hint';
            saveBtn.disabled = true;
            return;
        }
        const valid = getValidCodes(data[STORAGE_KEYS.HR_SHIFTS], data[STORAGE_KEYS.SHIFT_DICT]).has(val);
        codeInput.classList.toggle('valid',   valid);
        codeInput.classList.toggle('invalid', !valid);
        codeHint.textContent = valid ? '✔ 有效代號' : '✘ 非有效系統班別代號';
        codeHint.className   = `fill-code-hint ${valid ? 'ok' : 'err'}`;
        saveBtn.disabled = !valid;
    }

    document.querySelectorAll('input[name="blankFillMode"]').forEach(radio => {
        radio.addEventListener('change', () => {
            const isFill = radio.value === 'fill';
            codeInput.disabled = !isFill;
            if (isFill) {
                codeInput.focus();
                validateCode();
            } else {
                codeInput.classList.remove('valid', 'invalid');
                codeHint.textContent = '';
                codeHint.className   = 'fill-code-hint';
                document.getElementById('saveSettings').disabled = false;
            }
        });
    });
    codeInput.addEventListener('input', validateCode);
    if (mode === 'fill') validateCode();

    // ── WW/W+ 分配策略 ──────────────────────────────────────────
    const savedWwMode = data[STORAGE_KEYS.WW_MODE];
    const wwMode     = savedWwMode === 'B' ? 'B' : 'A';
    const wwModeHint = document.getElementById('wwModeHint');
    const checkedRadio = document.querySelector(`input[name="wwMode"][value="${wwMode}"]`);
    if (checkedRadio) checkedRadio.checked = true;
    document.querySelectorAll('input[name="wwMode"]').forEach(radio => {
        radio.addEventListener('change', () => { wwModeHint.textContent = ''; });
    });

    document.getElementById('openDictManager').onclick = () => {
        chrome.windows.create({ url: 'dict_manager.html', type: 'popup', width: 780, height: 500 });
    };

    document.getElementById('saveSettings').onclick = async () => {
        const selectedMode   = document.querySelector('input[name="blankFillMode"]:checked').value;
        const selectedWwMode = document.querySelector('input[name="wwMode"]:checked')?.value || '';
        await chrome.storage.local.set({
            [STORAGE_KEYS.AUTO_MODE]:        document.getElementById('autoMode').checked,
            [STORAGE_KEYS.SHOW_WEB_PREVIEW]: document.getElementById('showWebPreview').checked,
            [STORAGE_KEYS.SHOW_EXCEL_REPORT]:document.getElementById('showExcelReport').checked,
            [STORAGE_KEYS.BLANK_FILL_MODE]:  selectedMode,
            [STORAGE_KEYS.BLANK_FILL_CODE]:  selectedMode === 'fill' ? codeInput.value.trim() : '',
            [STORAGE_KEYS.WW_MODE]:          selectedWwMode,
        });
        alert("✅ 設定已儲存！");
    };
});
