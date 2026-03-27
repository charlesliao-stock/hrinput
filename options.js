document.addEventListener('DOMContentLoaded', async () => {
    const data = await chrome.storage.local.get([
        'autoMode', 'showWebPreview', 'showExcelReport',
        'blankFillMode', 'blankFillCode',
        'hrShifts', 'shiftDict'
    ]);

    // 初始化開關狀態
    document.getElementById('autoMode').checked        = data.autoMode || false;
    document.getElementById('showWebPreview').checked  = data.showWebPreview  !== false;
    document.getElementById('showExcelReport').checked = data.showExcelReport !== false;

    // 初始化寫入行為設定
    const mode = data.blankFillMode || 'keep';
    document.querySelector(`input[name="blankFillMode"][value="${mode}"]`).checked = true;
    const codeInput = document.getElementById('blankFillCode');
    const codeHint  = document.getElementById('fillCodeHint');
    codeInput.value = data.blankFillCode || '';
    codeInput.disabled = (mode === 'keep');

    // 有效班別清單（hrShifts 為系統代號；shiftDict 的 sys 欄也是系統代號）
    function getValidCodes() {
        const hr = data.hrShifts || [];
        const custom = (data.shiftDict || []).map(d => String(d.sys || '').trim()).filter(v => v);
        return new Set([...hr, ...custom]);
    }

    // 即時驗證
    function validateCode() {
        const val = codeInput.value.trim();
        const saveBtn = document.getElementById('saveSettings');
        if (!val) {
            codeInput.classList.remove('valid', 'invalid');
            codeHint.textContent = '';
            codeHint.className = 'fill-code-hint';
            saveBtn.disabled = true;
            return;
        }
        const valid = getValidCodes().has(val);
        codeInput.classList.toggle('valid',   valid);
        codeInput.classList.toggle('invalid', !valid);
        if (valid) {
            codeHint.textContent = '✔ 有效代號';
            codeHint.className = 'fill-code-hint ok';
            saveBtn.disabled = false;
        } else {
            codeHint.textContent = '✘ 非有效系統班別代號';
            codeHint.className = 'fill-code-hint err';
            saveBtn.disabled = true;
        }
    }

    // radio 切換
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
                codeHint.className = 'fill-code-hint';
                document.getElementById('saveSettings').disabled = false;
            }
        });
    });

    // 邊打邊驗
    codeInput.addEventListener('input', validateCode);

    // 初始狀態：若已選 fill 且有值，觸發一次驗證
    if (mode === 'fill') validateCode();

    // ✅ 開啟 dict_manager 視窗按鈕
    document.getElementById('openDictManager').onclick = () => {
        chrome.windows.create({ url: 'dict_manager.html', type: 'popup', width: 780, height: 500 });
    };

    // 儲存
    document.getElementById('saveSettings').onclick = async () => {
        const selectedMode = document.querySelector('input[name="blankFillMode"]:checked').value;
        await chrome.storage.local.set({
            autoMode:        document.getElementById('autoMode').checked,
            showWebPreview:  document.getElementById('showWebPreview').checked,
            showExcelReport: document.getElementById('showExcelReport').checked,
            blankFillMode:   selectedMode,
            blankFillCode:   selectedMode === 'fill' ? codeInput.value.trim() : '',
        });
        alert("✅ 設定已儲存！");
    };
});
