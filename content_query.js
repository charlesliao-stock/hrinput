// content_query.js 修改後版本
(function () {
  function initDropdowns() {
    // 檢查關鍵元素是否存在，若不存在則稍後再試或退出
    const checkTarget = document.getElementById('ctl00_ContentPlaceHolder1_QUERY_syymm_TextBox');
    if (!checkTarget) return;

    console.log("🔧 [KMUH Helper] 正在替換查詢頁面下拉選單...");

    // --- 以下為您原本的函式定義 ---
    function getMonthOptions(listId) {
      const rows = document.querySelectorAll(`#${listId} tr`);
      const options = [];
      rows.forEach(row => {
        const val = row.getAttribute('value');
        const tds = row.querySelectorAll('td');
        if (val && tds[1]) {
          options.push({ value: val, text: tds[1].textContent.trim() });
        }
      });
      return options;
    }

    function buildSelect(options, currentValue, onChange, width) {
      const sel = document.createElement('select');
      sel.style.cssText = `font-size:12px;height:24px;width:${width || 84}px;vertical-align:middle;margin:0 2px;border:1px solid #3498db;border-radius:4px;`;
      options.forEach(opt => {
        const o = document.createElement('option');
        o.value = opt.value;
        o.textContent = opt.text;
        if (opt.value === currentValue) o.selected = true;
        sel.appendChild(o);
      });
      sel.addEventListener('change', () => onChange(sel.value));
      return sel;
    }

    function replaceMonthPicker(textBoxId, hiddenId, listId) {
      const textBox = document.getElementById(textBoxId);
      const hidden  = document.getElementById(hiddenId);
      if (!textBox || !hidden) return;
      const options = getMonthOptions(listId);
      if (!options.length) return;
      const sel = buildSelect(options, hidden.value, val => {
        hidden.value = val;
        textBox.value = val;
        textBox.setAttribute('itemvalue', val);
      }, 84);
      textBox.style.display = 'none';
      textBox.parentNode.insertBefore(sel, textBox);
    }

    function replaceOrgPicker(textBoxId, hiddenId, listId, width) {
      const textBox = document.getElementById(textBoxId);
      const hidden  = document.getElementById(hiddenId);
      if (!textBox || !hidden) return;
      const options = getMonthOptions(listId);
      if (!options.length) return;
      const sel = buildSelect(options, hidden.value, val => {
        hidden.value = val;
        const label = options.find(o => o.value === val)?.text || val;
        textBox.value = label;
        textBox.setAttribute('itemvalue', val);
      }, width || 200);
      textBox.style.display = 'none';
      textBox.parentNode.insertBefore(sel, textBox);
    }

    // --- 執行替換邏輯 ---
    replaceMonthPicker(
      'ctl00_ContentPlaceHolder1_QUERY_syymm_TextBox',
      'ctl00_ContentPlaceHolder1_QUERY_syymm_Value',
      'ctl00_ContentPlaceHolder1_QUERY_syymm_list'
    );

    replaceMonthPicker(
      'ctl00_ContentPlaceHolder1_QUERY_eyymm_TextBox',
      'ctl00_ContentPlaceHolder1_QUERY_eyymm_Value',
      'ctl00_ContentPlaceHolder1_QUERY_eyymm_list'
    );

    replaceOrgPicker(
      'ctl00_ContentPlaceHolder1_QUERY_orgid_InputSelect1_TextBox',
      'ctl00_ContentPlaceHolder1_QUERY_orgid_InputSelect1_Value',
      'ctl00_ContentPlaceHolder1_QUERY_orgid_InputSelect1_list',
      220
    );

    replaceOrgPicker(
      'ctl00_ContentPlaceHolder1_QUERY_orgid_InputSelect2_TextBox',
      'ctl00_ContentPlaceHolder1_QUERY_orgid_InputSelect2_Value',
      'ctl00_ContentPlaceHolder1_QUERY_orgid_InputSelect2_list',
      180
    );
  }

  // 確保在網頁載入完成後執行
  if (document.readyState === 'complete') {
    initDropdowns();
  } else {
    window.addEventListener('load', initDropdowns);
  }
})();