// shared.js
// 跨頁面共用常數與純函式，供 options.js / quick_settings.js / dict_manager.js 引入

// ── Storage 鍵名常數（統一定義，避免字串散落各處） ─────────────────────
const STORAGE_KEYS = {
    AUTO_MODE:        'autoMode',
    SHOW_WEB_PREVIEW: 'showWebPreview',
    SHOW_EXCEL_REPORT:'showExcelReport',
    HR_SHIFTS:        'hrShifts',
    SHIFT_DICT:       'shiftDict',
    BLANK_FILL_MODE:  'blankFillMode',
    BLANK_FILL_CODE:  'blankFillCode',
    LAST_MONTH_DATA:  'lastMonthData',
    PENDING_NEXT_URL: 'pendingNextUrl',
    PENDING_UNKNOWN:  'pendingUnknownCodes',
    UPDATE_AVAILABLE: 'updateAvailable',
    LATEST_VERSION:   'latestVersion',
    DOWNLOAD_URL:     'downloadUrl',
    CHANGELOG:        'changelog',
};

// ── 取得所有有效的「系統班別代號」集合 ──────────────────────────────────
// hrShifts: Array<{code, start, end}>
// shiftDict: Array<{excel, sys, ...}>
function getValidCodes(hrShifts, shiftDict) {
    const hrCodes = (hrShifts || []).map(x =>
        typeof x === 'string' ? x : (x.code || '')
    ).filter(Boolean);
    const customCodes = (shiftDict || [])
        .map(d => String(d.sys || '').trim())
        .filter(Boolean);
    return new Set([...hrCodes, ...customCodes]);
}

// ── 將 hrShifts 陣列轉為 {code: {start, end}} 查找表 ───────────────────
function buildHrTimeMap(hrShiftsRaw) {
    const map = {};
    (hrShiftsRaw || []).forEach(x => {
        if (typeof x === 'object' && x.code) {
            map[x.code] = { start: x.start || null, end: x.end || null };
        }
    });
    return map;
}

// ── 從 storage 衍生月份相關計算結果 ────────────────────────────────────
// 回傳 { oldYymm, targetYymm, targetMonth, oldMonthDays, newMonthDays }
function deriveMonthContext(lastMonthData) {
    const oldYymm      = lastMonthData?.yymm || "";
    const targetYymm   = oldYymm ? getNextYM(oldYymm) : "";
    const targetMonth  = targetYymm ? parseInt(targetYymm.substring(4, 6)) : -1;
    const oldMonthDays = lastMonthData?.monthDays || 0;
    const newMonthDays = targetYymm
        ? new Date(parseInt(targetYymm.substring(0, 4)), parseInt(targetYymm.substring(4, 6)), 0).getDate()
        : 31;
    return { oldYymm, targetYymm, targetMonth, oldMonthDays, newMonthDays };
}

// ── 通用 mini-popup 視窗 ─────────────────────────────────────────────
// opts: { title, message, btnColor, btnLabel, width, height }
function createPopupWindow(opts) {
    const {
        title    = '',
        message  = '',
        btnColor = '#3498db',
        btnLabel = '確定',
        width    = 320,
        height   = 200,
    } = opts;
    const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
    <style>
        body { width:${width}px; margin:0; display:flex; flex-direction:column;
               align-items:center; justify-content:center; gap:16px;
               padding:24px 20px; box-sizing:border-box;
               font-family:"Microsoft JhengHei",sans-serif; background:#fff; overflow:hidden; }
        .msg { font-size:14px; color:#2c3e50; font-weight:bold;
               text-align:center; line-height:1.7; white-space:pre-line; }
        button { width:100%; padding:10px; background:${btnColor}; color:white;
                 border:none; border-radius:6px; font-size:14px;
                 font-weight:bold; cursor:pointer; }
        button:hover { filter:brightness(0.9); }
    </style></head><body>
    <div class="msg">${title ? `<b style="font-size:15px;">${title}</b>\n` : ''}${message}</div>
    <button onclick="window.close()">${btnLabel}</button>
    </body></html>`;
    chrome.windows.create({
        url: 'data:text/html;charset=utf-8,' + encodeURIComponent(html),
        type: 'popup', width, height, focused: true,
    });
}
