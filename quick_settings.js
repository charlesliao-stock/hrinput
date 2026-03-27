document.addEventListener('DOMContentLoaded', async () => {
  const data = await chrome.storage.local.get([
    'autoMode', 'showWebPreview', 'showExcelReport',
    'blankFillMode', 'blankFillCode',
    'hrShifts', 'shiftDict'
  ]);

  // 初始化勾選狀態
  document.getElementById('autoMode').checked = data.autoMode || false;
  document.getElementById('showWebPreview').checked = data.showWebPreview !== false;
  document.getElementById('showExcelReport').checked = data.showExcelReport !== false;

  // 初始化寫入行為
  const mode = data.blankFillMode || 'keep';
  document.querySelector(`input[name="blankFillMode"][value="${mode}"]`).checked = true;
  const codeInput = document.getElementById('blankFillCode');
  const fillHint = document.getElementById('fillHint');
  codeInput.value = data.blankFillCode || '';
  codeInput.disabled = (mode === 'keep');

  // ✅ 修正：補上 getValidCodes() 的右括號
  function getValidCodes() {
    const hr = data.hrShifts || [];
    const custom = (data.shiftDict || []).map(d => String(d.sys || '').trim()).filter(v => v);
    return new Set([...hr, ...custom]);
  } // ✅ 補上這個 }

  // 即時驗證
  function validateCode() {
    const val = codeInput.value.trim();
    const saveBtn = document.getElementById('saveBtn');
    if (!val) {
      codeInput.classList.remove('valid', 'invalid');
      fillHint.textContent = '';
      fillHint.className = 'fill-hint';
      saveBtn.disabled = true;
      return;
    }
    const valid = getValidCodes().has(val);
    codeInput.classList.toggle('valid', valid);
    codeInput.classList.toggle('invalid', !valid);
    if (valid) {
      fillHint.textContent = '✔ 有效代號';
      fillHint.className = 'fill-hint ok';
      saveBtn.disabled = false;
    } else {
      fillHint.textContent = '✘ 非有效系統班別';
      fillHint.className = 'fill-hint err';
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
        fillHint.textContent = '';
        fillHint.className = 'fill-hint';
        document.getElementById('saveBtn').disabled = false;
      }
    });
  });

  // 邊打邊驗
  codeInput.addEventListener('input', validateCode);

  // 初始狀態：若已選 fill 且有值，觸發一次驗證
  if (mode === 'fill') validateCode();

  // 自動調整視窗高度
  const updateHeight = () => {
    const h = document.body.scrollHeight + 30;
    chrome.windows.getCurrent(win => chrome.windows.update(win.id, { height: h }));
  };
  setTimeout(updateHeight, 100);

  // 儲存
  document.getElementById('saveBtn').onclick = async () => {
    const selectedMode = document.querySelector('input[name="blankFillMode"]:checked').value;
    await chrome.storage.local.set({
      autoMode: document.getElementById('autoMode').checked,
      showWebPreview: document.getElementById('showWebPreview').checked,
      showExcelReport: document.getElementById('showExcelReport').checked,
      blankFillMode: selectedMode,
      blankFillCode: selectedMode === 'fill' ? codeInput.value.trim() : '',
    });
    window.close();
  };
});
