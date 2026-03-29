// content.js
// 頁面驗證統一由 popup.js 的 sendMessage 負責，content.js 不重複處理
// HR 預設班別統一由 background.js 初始化至 storage，content.js 不再 hardcode

console.log("🚀 [KMUH Helper] 核心啟動 (MAIN World)");

// ── 修復網頁原始碼中的 IE 專屬 API (主環境 Polyfill) ──────────────
// 解決 KmuhdeptshiftEdit.aspx 第 855 行 Uncaught TypeError: window.attachEvent is not a function
// 由於 manifest.json 已設定 world: "MAIN"，此處代碼直接運行在網頁環境中
(function() {
    // 1. 修復 attachEvent
    if (!window.attachEvent && window.addEventListener) {
        window.attachEvent = function(event, handler) {
            const eventName = event.startsWith('on') ? event.substring(2) : event;
            window.addEventListener(eventName, handler);
        };
        console.log("🔧 [KMUH Helper] 已成功注入 window.attachEvent 相容性修補");
    }

    // 2. 修復 showModalDialog (將其轉向 alert)
    if (!window.showModalDialog) {
        window.showModalDialog = function(url, arg, feat) {
            console.log("🔧 [KMUH Helper] 攔截到 showModalDialog 呼叫");
            // 嘗試從參數中提取純文字訊息（去除 HTML 標籤）
            const msg = (typeof arg === 'string') ? arg.replace(/<[^>]+>/g, '') : "網頁發生錯誤，請檢查輸入內容。";
            alert("⚠️ 網頁錯誤訊息：\n\n" + msg);
        };
        console.log("🔧 [KMUH Helper] 已成功注入 window.showModalDialog 相容性修補");
    }
})();


function formatEmpId(id) {
    if (!id) return "";
    const s = String(id).trim();
    if (!/^\d+$/.test(s)) return "";
    return s.padStart(7, '0');
}

function getNextYM(yymm) {
    if (!yymm || yymm.length !== 6) return "";
    let y = parseInt(yymm.substring(0, 4)), m = parseInt(yymm.substring(4, 6)) + 1;
    if (m > 12) { m = 1; y++; }
    return String(y) + String(m).padStart(2, '0');
}

function parseCyclePeriods() {
    const periods = [];
    const re = /【(\d+)】\s*(\d{1,2}\/\d{1,2})\s*[~～]\s*(\d{1,2}\/\d{1,2})/g;
    let m;
    while ((m = re.exec(document.body.innerText)) !== null) {
        periods.push({ label: m[1], start: m[2], end: m[3] });
    }
    return periods;
}

// 解析頁面上的雙週 FF 班檢查週別，例如《1》02/16~03/01
function parseFFPeriods() {
    const periods = [];
    const re = /《(\d+)》\s*(\d{1,2}\/\d{1,2})\s*[~～]\s*(\d{1,2}\/\d{1,2})/g;
    let m;
    while ((m = re.exec(document.body.innerText)) !== null) {
        periods.push({ label: m[1], start: m[2], end: m[3] });
    }
    return periods;
}

// mm/dd 轉為 Date 物件，年份以 refYymm 為基準推算（處理跨年）
function mmddToDate(mmdd, refYymm) {
    const [mm, dd] = mmdd.split('/').map(Number);
    const refYear  = parseInt(refYymm.substring(0, 4));
    const refMonth = parseInt(refYymm.substring(4, 6));
    const year = (mm < refMonth - 6) ? refYear + 1 : refYear;
    return new Date(year, mm - 1, dd);
}

// Date 物件轉為 mm/dd 字串
function dateToMmdd(d) {
    return `${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;
}

// mm/dd 轉為全局索引（index 0 = 讀取月第1天，oldMonthDays 以後為匯入目標月）
function mmddToGlobalIdx(mmdd, oldYymm, oldMonthDays) {
    const base   = mmddToDate(`${oldYymm.substring(4, 6)}/01`, oldYymm);
    const target = mmddToDate(mmdd, oldYymm);
    return Math.round((target - base) / 86400000);
}

// 從讀取月頁面的最後一個週期向後延伸，產生涵蓋匯入目標月的完整檢查區間清單
// periodDays: FF週別=14天, 四週變形=28天
// 篩選條件：開始或結束月份包含 targetMonth
function buildCheckRanges(lastPeriod, targetMonth, periodDays, oldYymm, oldMonthDays) {
    if (!lastPeriod) return [];

    const ranges = [];
    let startDate = mmddToDate(lastPeriod.start, oldYymm);
    let endDate   = mmddToDate(lastPeriod.end,   oldYymm);
    // 若 end < start（例如跨年），end 往後推一年
    if (endDate < startDate) endDate.setFullYear(endDate.getFullYear() + 1);

    while (true) {
        const startMonth = startDate.getMonth() + 1;
        const endMonth   = endDate.getMonth() + 1;

        // 停止條件：開始月份已超過 targetMonth
        if (startMonth > targetMonth) break;

        // 包含 targetMonth 才納入檢查
        if (startMonth === targetMonth || endMonth === targetMonth) {
            const mmddStart = dateToMmdd(startDate);
            const mmddEnd   = dateToMmdd(endDate);
            ranges.push({
                start:    mmddStart,
                end:      mmddEnd,
                startIdx: mmddToGlobalIdx(mmddStart, oldYymm, oldMonthDays),
                endIdx:   mmddToGlobalIdx(mmddEnd,   oldYymm, oldMonthDays),
            });
        }

        // 向後延伸一個週期
        const nextStart = new Date(endDate);
        nextStart.setDate(nextStart.getDate() + 1);
        const nextEnd = new Date(nextStart);
        nextEnd.setDate(nextEnd.getDate() + periodDays - 1);
        startDate = nextStart;
        endDate   = nextEnd;
    }

    return ranges;
}

// ─────────────────────────────────────────────────────────────────
// 訊息監聽入口
// ─────────────────────────────────────────────────────────────────
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {

    // ── 步驟 1：記憶本月班表 ──────────────────────────────────────
    if (request.action === "readAndMemorize") {
        const data = captureWebSchedule();
        const now = new Date();
        const sysYymm = String(now.getFullYear()) + String(now.getMonth() + 1).padStart(2, '0');

        if (data.yymm && data.yymm !== sysYymm) {
            const proceed = confirm(
                `⚠️ 月份提醒\n\n網頁顯示月份：${data.yymm}\n系統當前月份：${sysYymm}\n\n兩者不一致，是否仍要繼續記憶？`
            );
            if (!proceed) return sendResponse({ success: false, message: "使用者取消" });
        }

        const periods   = parseCyclePeriods();
        const ffPeriods = parseFFPeriods();
        data.cyclePeriods = periods;
        data.ffPeriods    = ffPeriods;

        const nextUrl = window.location.href.replace(/yymm=\d{6}/, `yymm=${getNextYM(data.yymm)}`);
        const toSave = { lastMonthData: data };
        if (request.autoMode && request.showPreview) {
            toSave["pendingNextUrl"] = nextUrl;
        } else {
            chrome.storage.local.remove('pendingNextUrl');
        }

        chrome.storage.local.set(toSave, () => {
            if (request.showPreview) {
                const hint = request.autoMode
                    ? "記憶完成。關閉此視窗後將自動跳轉至下個月。"
                    : "記憶完成。";
                showModal(`步驟 1：${data.yymm} 預覽報告`, data, hint);
            }
            sendResponse({
                success: true,
                yymm: data.yymm,
                nextUrl,
                hasPreview: request.showPreview,
                periods,
                ffPeriods,
            });
        });
        return true;
    }

    // ── 步驟 2：匯入 Excel 並驗證 ────────────────────────────────
    if (request.action === "autoProcessExcel") {
        handleExcelProcess(request).then(res => sendResponse(res));
        return true;
    }

    // ── 步驟 4：寫入班表 ─────────────────────────────────────────
    if (request.action === "injectOnly") {
        executeInjectionFlow(request.excelData).then(res => sendResponse(res));
        return true;
    }
});

// ─────────────────────────────────────────────────────────────────
// 步驟 2：匯入 Excel 並驗證
// ─────────────────────────────────────────────────────────────────
async function handleExcelProcess(req) {
    const storage = await chrome.storage.local.get(['shiftDict', 'hrShifts', 'lastMonthData']);
    const oldYymm     = storage.lastMonthData?.yymm || "";
    const targetYymm  = oldYymm ? getNextYM(oldYymm) : "";
    const targetMonth = targetYymm ? parseInt(targetYymm.substring(4, 6)) : -1;
    const excelMap    = parseExcel(req.excelData, targetYymm);
    const customDict  = storage.shiftDict || [];
    const hrShiftsRaw = storage.hrShifts  || [];
    const lastData    = storage.lastMonthData;

    const hrShiftsList = hrShiftsRaw.map(x => typeof x === 'string' ? x : x.code);
    const hrTimeMap    = {};
    hrShiftsRaw.forEach(x => {
        if (typeof x === 'object' && x.code) {
            hrTimeMap[x.code] = { start: x.start || null, end: x.end || null };
        }
    });

    const unknownCodes = new Set();
    for (let id in excelMap) {
        excelMap[id].shifts.forEach(code => {
            const cStr = String(code || "").trim();
            if (!cStr) return;
            if (!hrShiftsList.includes(cStr) && !customDict.some(d => String(d.excel).trim() === cStr)) {
                unknownCodes.add(cStr);
            }
        });
    }
    if (unknownCodes.size > 0) {
        return { success: false, unknownCodes: Array.from(unknownCodes) };
    }

    const dataWithId   = Object.entries(excelMap).map(([id, v]) => ({ empId: id, ...v }));
    const oldMonthDays = lastData?.monthDays || 0;
    const newMonthDays = targetYymm
        ? new Date(parseInt(targetYymm.substring(0, 4)), parseInt(targetYymm.substring(4, 6)), 0).getDate()
        : 31;

    const lastCycle = (lastData?.cyclePeriods || []).at(-1) || null;
    const lastFF    = (lastData?.ffPeriods    || []).at(-1) || null;

    const cycleRanges = buildCheckRanges(lastCycle, targetMonth, 28, oldYymm, oldMonthDays);
    const ffRanges    = buildCheckRanges(lastFF,    targetMonth, 14, oldYymm, oldMonthDays);

    const allRanges = [...cycleRanges, ...ffRanges];
    const biStart = allRanges.length > 0 ? Math.min(...allRanges.map(r => r.startIdx)) : oldMonthDays;
    const biEnd   = allRanges.length > 0 ? Math.max(...allRanges.map(r => r.endIdx))   : oldMonthDays + 27;

    const cycleLabel = cycleRanges.map((r, i) => `【${i + 1}】${r.start}～${r.end}`).join('、') || '未知';
    const ffLabel    = ffRanges.map((r, i)    => `《${i + 1}》${r.start}～${r.end}`).join('、') || '未知';
    const infoText   = `四週變形：${cycleLabel}　／　FF雙週：${ffLabel}`;

    const check = runDetailedCheck(lastData, excelMap, customDict, hrTimeMap, cycleRanges, ffRanges, oldMonthDays, newMonthDays, oldYymm, targetYymm);
    if (req.showReport || check.errors.length > 0) {
        showModal("Excel 班表預覽與檢測報告", {
            headers:      getHeaders(),
            data:         dataWithId,
            errors:       check.errors,
            monthDays:    oldMonthDays,
            biStart,
            biEnd,
            cycleRanges,
            ffRanges,
            blankFillMode: req.blankFillMode || 'keep',
            blankFillCode: req.blankFillCode || '',
        }, infoText);
    }
    return {
        success:           check.errors.length === 0,
        noOldDataWarnings: check.noOldDataWarnings,
    };
}

function timeToMinutes(t) {
    if (!t) return null;
    const [h, m] = t.split(':').map(Number);
    return h * 60 + (m || 0);
}

function getShiftTime(code, hrTimeMap, dict) {
    if (hrTimeMap[code]) {
        const { start, end } = hrTimeMap[code];
        if (!start && !end) return null;
        const startMin = timeToMinutes(start);
        let   endMin   = timeToMinutes(end);
        if (endMin !== null && startMin !== null && endMin <= startMin) {
            endMin += 1440;
        }
        return { startMin, endMin };
    }
    const entry = dict.find(x => String(x.sys || '').trim() === code);
    if (entry) return null;
    return null;
}

function giToDateStr(gi, oldYymm, targetYymm, oldMonthDays) {
    if (!oldYymm) return `第${gi + 1}天`;
    let year, month, day;
    if (gi < oldMonthDays) {
        year  = parseInt(oldYymm.substring(0, 4));
        month = parseInt(oldYymm.substring(4, 6));
        day   = gi + 1;
    } else {
        year  = parseInt(targetYymm.substring(0, 4));
        month = parseInt(targetYymm.substring(4, 6));
        day   = gi - oldMonthDays + 1;
    }
    return `${month}月${day}日`;
}

function runDetailedCheck(old, exc, dict, hrTimeMap, cycleRanges, ffRanges, oldMonthDays, newMonthDays, oldYymm, targetYymm) {
    if (!old?.data && !exc) return { errors: [], noOldDataWarnings: [] };
    const err = [];
    const noOldDataWarnings = [];
    const toDate = (gi) => giToDateStr(gi, oldYymm, targetYymm, oldMonthDays);

    for (let id in exc) {
        const oStf       = old?.data?.find(p => formatEmpId(p.empId) === formatEmpId(id));
        const hasOldData = !!oStf;
        const validStart = hasOldData ? 0 : oldMonthDays;
        const validEnd   = oldMonthDays + newMonthDays - 1;

        if (!hasOldData) {
            noOldDataWarnings.push({ empId: id, name: exc[id].name || '' });
        }

        const oldShifts = hasOldData ? oStf.shifts : Array(oldMonthDays).fill('');
        const combined  = [...oldShifts, ...exc[id].shifts].map(s => {
            const d = dict.find(x => String(x.excel).trim() === String(s).trim());
            return d ? d.sys : s;
        });

        const isRangeValid = (r) => r.startIdx >= validStart && r.endIdx <= validEnd;

        ffRanges.forEach((r, i) => {
            if (!isRangeValid(r)) return;
            const count = combined.slice(r.startIdx, r.endIdx + 1).filter(s => s === 'FF').length;
            if (count !== 2) {
                err.push({
                    empId:    id,
                    startIdx: r.startIdx,
                    endIdx:   r.endIdx,
                    type:     `FF_${i + 1}`,
                    msg:      `FF雙週《${i + 1}》${r.start}～${r.end} FF=${count}（應2）`,
                });
            }
        });

        const ffIndices = [];
        for (let gi = validStart; gi <= validEnd; gi++) {
            if (combined[gi] === 'FF') ffIndices.push(gi);
        }
        for (let fi = 0; fi < ffIndices.length - 1; fi++) {
            const gap = ffIndices[fi + 1] - ffIndices[fi] - 1;
            if (gap > 12) {
                const d1 = toDate(ffIndices[fi]);
                const d2 = toDate(ffIndices[fi + 1]);
                err.push({
                    empId:    id,
                    startIdx: ffIndices[fi],
                    endIdx:   ffIndices[fi + 1],
                    type:     'FF_GAP',
                    msg:      `FF間隔過長：${d1}(FF) 與 ${d2}(FF) 之間間隔 ${gap} 天（最多12天）`,
                });
            }
        }

        cycleRanges.forEach((r, i) => {
            if (!isRangeValid(r)) return;
            const count = combined.slice(r.startIdx, r.endIdx + 1)
                .filter(s => s === 'WW' || s === 'W+').length;
            if (count !== 4) {
                err.push({
                    empId:    id,
                    startIdx: r.startIdx,
                    endIdx:   r.endIdx,
                    type:     `WW_${i + 1}`,
                    msg:      `四週變形【${i + 1}】${r.start}～${r.end} WW=${count}（應4）`,
                });
            }
        });

        let prevCode    = null;
        let prevEndMin  = null;
        let prevGi      = -1;

        for (let gi = Math.max(0, validStart - 1); gi <= validEnd; gi++) {
            const code = combined[gi] || '';
            if (!code) continue;

            const timeInfo = getShiftTime(code, hrTimeMap, dict);
            if (!timeInfo) {
                prevCode   = null;
                prevEndMin = null;
                prevGi     = -1;
                continue;
            }

            const { startMin, endMin } = timeInfo;

            if (prevEndMin !== null) {
                const daysBetween = gi - prevGi - 1;
                const prevEndAbs = prevGi * 1440 + prevEndMin;
                const nextDayOffset = (daysBetween + 1) * 1440;
                const nextStartAbs = prevGi * 1440 + nextDayOffset + startMin;
                const gap = nextStartAbs - prevEndAbs;

                const MIN_GAP = 660;
                if (gap < MIN_GAP) {
                    const gapH   = Math.floor(Math.max(gap, 0) / 60);
                    const gapM   = Math.max(gap, 0) % 60;
                    const gapStr = gap <= 0 ? '0分（班別重疊）' : (gapM > 0 ? `${gapH}小時${gapM}分` : `${gapH}小時`);
                    const d1 = toDate(prevGi);
                    const d2 = toDate(gi);
                    err.push({
                        empId:    id,
                        startIdx: prevGi,
                        endIdx:   gi,
                        type:     'REST_SHORT',
                        msg:      `接班間距不足：${d1}(${prevCode}) 與 ${d2}(${code}) 間距僅 ${gapStr}（未達11小時）`,
                    });
                }
            }
            prevCode   = code;
            prevEndMin = endMin;
            prevGi     = gi;
        }
    }
    return { errors: err, noOldDataWarnings };
}

// ─────────────────────────────────────────────────────────────────
// UI：Modal 報告視窗
// ─────────────────────────────────────────────────────────────────
let modalState = {
    dataset: null,
    info: '',
    storage: null,
    hrTimeMap: {},
    oldYymm: '',
    targetYymm: '',
    oldMonthDays: 0,
    newMonthDays: 0,
    cycleRanges: [],
    ffRanges: []
};

async function showModal(title, dataset, info) {
    const oldModal = document.getElementById('kmuh-modal'); if (oldModal) oldModal.remove();
    const oldStyle = document.getElementById('kmuh-modal-style'); if (oldStyle) oldStyle.remove();

    const storage = await chrome.storage.local.get(['shiftDict', 'hrShifts', 'lastMonthData']);
    const hrShiftsRaw = storage.hrShifts || [];
    const hrTimeMap = {};
    hrShiftsRaw.forEach(x => {
        if (typeof x === 'object' && x.code) {
            hrTimeMap[x.code] = { start: x.start || null, end: x.end || null };
        }
    });

    const oldYymm = storage.lastMonthData?.yymm || "";
    const targetYymm = oldYymm ? getNextYM(oldYymm) : "";
    const targetMonth = targetYymm ? parseInt(targetYymm.substring(4, 6)) : -1;
    const oldMonthDays = storage.lastMonthData?.monthDays || 0;
    const newMonthDays = targetYymm
        ? new Date(parseInt(targetYymm.substring(0, 4)), parseInt(targetYymm.substring(4, 6)), 0).getDate()
        : 31;

    const lastCycle = (storage.lastMonthData?.cyclePeriods || []).at(-1) || null;
    const lastFF = (storage.lastMonthData?.ffPeriods || []).at(-1) || null;
    const cycleRanges = buildCheckRanges(lastCycle, targetMonth, 28, oldYymm, oldMonthDays);
    const ffRanges = buildCheckRanges(lastFF, targetMonth, 14, oldYymm, oldMonthDays);

    modalState = {
        dataset,
        info,
        storage,
        hrTimeMap,
        oldYymm,
        targetYymm,
        oldMonthDays,
        newMonthDays,
        cycleRanges,
        ffRanges
    };

    renderModalContent(title);
}

function renderModalContent(title) {
    const { dataset, info, oldMonthDays, cycleRanges, ffRanges } = modalState;
    const h = dataset.headers;
    const mDays = oldMonthDays;
    const total = dataset.data.length;
    const errorIds = new Set(dataset.errors?.map(e => formatEmpId(e.empId)));
    const errCount = errorIds.size;

    const CYCLE_COLORS = ['#dbeafe', '#bfdbfe', '#93c5fd'];
    const FF_COLORS = ['#ede9fe', '#ddd6fe', '#c4b5fd'];

    const cycleCss = cycleRanges.map((_, i) =>
        `.hd-cy-${i} { background:${CYCLE_COLORS[i % CYCLE_COLORS.length]} !important; }`
    ).join('\n');
    const ffCss = ffRanges.map((_, i) =>
        `.hd-ff-${i} { background:${FF_COLORS[i % FF_COLORS.length]} !important; }`
    ).join('\n');

    const colCls = (gi) => {
        for (let i = 0; i < ffRanges.length; i++) {
            if (gi >= ffRanges[i].startIdx && gi <= ffRanges[i].endIdx) return `hd-ff-${i}`;
        }
        for (let i = 0; i < cycleRanges.length; i++) {
            if (gi >= cycleRanges[i].startIdx && gi <= cycleRanges[i].endIdx) return `hd-cy-${i}`;
        }
        return "";
    };

    const legendItems = [
        ...cycleRanges.map((r, i) =>
            `<span style="display:inline-flex;align-items:center;gap:3px;margin-right:8px;">
              <span style="display:inline-block;width:12px;height:12px;background:${CYCLE_COLORS[i % CYCLE_COLORS.length]};border:1px solid #aaa;border-radius:2px;"></span>
              四週【${i + 1}】${r.start}～${r.end}
            </span>`),
        ...ffRanges.map((r, i) =>
            `<span style="display:inline-flex;align-items:center;gap:3px;margin-right:8px;">
              <span style="display:inline-block;width:12px;height:12px;background:${FF_COLORS[i % FF_COLORS.length]};border:1px solid #aaa;border-radius:2px;"></span>
              FF《${i + 1}》${r.start}～${r.end}
            </span>`),
    ].join('');

    const errLegend = [
        { color: '#e74c3c', bg: '#fff2f2', label: '四週變形/FF數量錯誤' },
        { color: '#e67e22', bg: '#fff8f0', label: 'FF間隔超過12天' },
        { color: '#8e44ad', bg: '#fdf2ff', label: '接班間距不足11小時' },
    ].map(x =>
        `<span style="display:inline-flex;align-items:center;gap:3px;margin-right:10px;">
          <span style="display:inline-block;width:24px;height:14px;background:${x.bg};border:2px solid ${x.color};border-radius:2px;"></span>
          ${x.label}
        </span>`
    ).join('');

    const oldStyle = document.getElementById('kmuh-modal-style'); if (oldStyle) oldStyle.remove();
    const style = document.createElement('style');
    style.id = 'kmuh-modal-style';
    style.innerHTML = `
        #kmuh-modal { position:fixed; top:2%; left:2%; width:96%; height:94%; background:#fdfdfe; z-index:10000; padding:25px; box-shadow:0 15px 60px rgba(0,0,0,0.4); overflow:auto; border-radius:15px; font-family:sans-serif; }
        .summary-row { display:flex; gap:15px; margin-bottom:15px; }
        .card { flex:1; padding:15px; border-radius:10px; color:white; display:flex; flex-direction:column; align-items:center; }
        .card-blue { background:#3498db; } .card-green { background:#2ecc71; } .card-red { background:#e74c3c; }
        .card-val { font-size:2em; font-weight:bold; margin-top:5px; }
        .table-container { overflow-x:auto; border:1px solid #dfe6e9; border-radius:8px; }
        .report-table { width:100%; border-collapse:separate; border-spacing:0; background:white; }
        .report-table th, .report-table td { border:1px solid #ecf0f1; padding:8px; text-align:center; font-size:13px; min-width:32px; }
        .sticky-col { position:sticky; left:0; background:#f8f9fa !important; z-index:5; font-weight:bold; border-right:2px solid #bdc3c7 !important; min-width:70px; }
        .sticky-name { position:sticky; left:71px; background:#f8f9fa !important; z-index:5; font-weight:bold; border-right:2px solid #bdc3c7 !important; min-width:60px; }
        .cell-err { background:#fff2f2 !important; border:2px solid #ff7675 !important; }
        .tooltip { position:relative; cursor:help; }
        #kmuh-tip { position:fixed; background:#2d3436; color:white; padding:8px 14px; border-radius:6px; font-size:12px; z-index:99999; pointer-events:none; display:none; box-shadow:0 4px 12px rgba(0,0,0,0.4); }
        .editable-cell:focus { outline: 2px solid #3498db; background: #fff !important; }
        ${cycleCss}
        ${ffCss}
    `;
    document.head.appendChild(style);

    const thW = h.weekdays.map((w, i) =>
        `<th class="${colCls(mDays + i)}" style="color:${w === '日' || w === '六' ? '#e74c3c' : 'inherit'}">${w}</th>`
    ).join('');

    const thD = h.dates.map((d, i) =>
        `<th class="${colCls(mDays + i)}">${d}</th>`
    ).join('');

    const rows = dataset.data.map((p, pIdx) => {
        const pErrs = dataset.errors?.filter(e => formatEmpId(p.empId) === formatEmpId(e.empId)) || [];
        const isFill = dataset.blankFillMode === 'fill' && dataset.blankFillCode;

        const ERR_COLOR_MAP = {
            WW: { border: '#e74c3c', bg: '#fff2f2' },
            FF: { border: '#e74c3c', bg: '#fff2f2' },
            GAP: { border: '#e67e22', bg: '#fff8f0' },
            REST: { border: '#8e44ad', bg: '#fdf2ff' },
        };

        function getErrColor(type) {
            if (!type) return ERR_COLOR_MAP.WW;
            if (type === 'FF_GAP') return ERR_COLOR_MAP.GAP;
            if (type === 'REST_SHORT') return ERR_COLOR_MAP.REST;
            if (type.startsWith('FF_')) return ERR_COLOR_MAP.FF;
            return ERR_COLOR_MAP.WW;
        }

        const cells = p.shifts.map((s, i) => {
            const gi = mDays + i;
            const isBlank = !s;
            const displayVal = isBlank && isFill
                ? `<span style="color:#e67e22;font-size:11px;">→${dataset.blankFillCode}</span>`
                : (s || '');

            const cellErrs = pErrs.filter(e => gi >= e.startIdx && gi <= e.endIdx);

            let borderStyle = '';
            let bgStyle = '';
            let tipText = '';

            if (cellErrs.length > 0) {
                const bigErr = cellErrs.reduce((a, b) =>
                    (b.endIdx - b.startIdx) > (a.endIdx - a.startIdx) ? b : a
                );
                const { border, bg } = getErrColor(bigErr.type);
                const isFirst = gi === bigErr.startIdx;
                const isLast = gi === bigErr.endIdx;
                borderStyle = `border-top:2px solid ${border} !important; border-bottom:2px solid ${border} !important;`
                    + (isFirst ? `border-left:2px solid ${border} !important;` : 'border-left:none !important;')
                    + (isLast ? `border-right:2px solid ${border} !important;` : 'border-right:none !important;');
                bgStyle = `background:${bg} !important;`;
                tipText = pErrs.map(e => e.msg).join('\n');
            } else if (isBlank && isFill) {
                tipText = `將填入 ${dataset.blankFillCode}`;
            }

            const wkBg = h.weekdays[i] === '日' || h.weekdays[i] === '六' ? '#fef9f9' : 'white';
            const cellBg = cellErrs.length > 0 ? '' : `background:${wkBg};`;
            const tipAttr = tipText ? `data-kmuh-tip="${tipText.replace(/"/g, '&quot;')}"` : '';
            const cls = (tipText ? 'tooltip ' : '') + 'editable-cell';

            return `<td class="${cls}" ${tipAttr} contenteditable="true" data-p-idx="${pIdx}" data-s-idx="${i}" style="${cellBg}${bgStyle}${borderStyle}">${displayVal}</td>`;
        }).join('');
        return `<tr><td class="sticky-col">${p.empId || ''}</td><td class="sticky-name">${p.name || ''}</td>${cells}</tr>`;
    }).join('');

    let m = document.getElementById('kmuh-modal');
    if (!m) {
        m = document.createElement('div');
        m.id = 'kmuh-modal';
        document.body.appendChild(m);
    }

    m.innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <h2 style="margin:0;">📊 ${title}</h2>
            <div style="display:flex; gap:10px;">
                <button id="saveM" style="padding:10px 35px; background:#2ecc71; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold; font-size:14px;">💾 寫入班表</button>
                <button id="closeM" style="padding:10px 35px; background:#3498db; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold; font-size:14px;">✖ 關閉</button>
            </div>
        </div>
        ${info ? `<div style="margin-bottom:8px; padding:8px 12px; background:#eaf4fb; border-radius:6px; font-size:13px; color:#2c3e50;">ℹ️ ${info}</div>` : ''}
        <div style="margin-bottom:8px; padding:8px 12px; background:#fff3cd; border-radius:6px; font-size:13px; color:#856404; border:1px solid #ffeeba;">💡 提示：您可以直接點擊表格中的班別進行修改，系統會自動重新驗證。</div>
        ${legendItems ? `<div style="margin-bottom:6px; padding:6px 12px; background:#f8f9fa; border-radius:6px; font-size:12px; color:#555; display:flex; flex-wrap:wrap; gap:4px; align-items:center;"><b style="margin-right:6px;">檢查區間：</b>${legendItems}</div>` : ''}
        <div style="margin-bottom:12px; padding:6px 12px; background:#f8f9fa; border-radius:6px; font-size:12px; color:#555; display:flex; flex-wrap:wrap; gap:4px; align-items:center;"><b style="margin-right:6px;">錯誤類型：</b>${errLegend}</div>
        <div class="summary-row">
            <div class="card card-blue"><span>檢測總人數</span><div class="card-val">${total}</div></div>
            <div class="card card-green"><span>通過檢核</span><div class="card-val">${total - errCount}</div></div>
            <div class="card card-red"><span>違反規範</span><div class="card-val">${errCount}</div></div>
        </div>
        <div class="table-container">
            <table class="report-table">
                <thead>
                    <tr style="background:#f1f2f6;">
                        <th rowspan="2" class="sticky-col">職編</th>
                        <th rowspan="2" class="sticky-name">姓名</th>
                        ${thW}
                    </tr>
                    <tr style="background:#f1f2f6;">${thD}</tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>
        </div>`;

    setupModalEvents(m, title);
}

function setupModalEvents(m, title) {
    const tip = document.getElementById('kmuh-tip') || document.createElement('div');
    if (!tip.id) {
        tip.id = 'kmuh-tip';
        document.body.appendChild(tip);
    }

    const showTip = e => {
        const td = e.target.closest('[data-kmuh-tip]');
        if (!td) return;
        const lines = td.getAttribute('data-kmuh-tip').split('\n');
        tip.innerHTML = lines.map(l =>
            `<div style="white-space:nowrap; line-height:1.8;">${l}</div>`
        ).join('');
        tip.style.display = 'block';
    };
    const moveTip = e => {
        if (tip.style.display === 'none') return;
        const x = e.clientX + 14;
        const y = e.clientY - tip.offsetHeight - 10;
        const maxX = window.innerWidth - tip.offsetWidth - 10;
        tip.style.left = Math.min(x, maxX) + 'px';
        tip.style.top = Math.max(y, 10) + 'px';
    };
    const hideTip = () => { tip.style.display = 'none'; };

    m.addEventListener('mouseover', showTip);
    m.addEventListener('mousemove', moveTip);
    m.addEventListener('mouseleave', hideTip);
    m.addEventListener('mouseout', e => {
        if (!e.target.closest('[data-kmuh-tip]')) hideTip();
    });

    m.querySelectorAll('.editable-cell').forEach(cell => {
        cell.addEventListener('blur', e => {
            const pIdx = parseInt(e.target.dataset.pIdx);
            const sIdx = parseInt(e.target.dataset.sIdx);
            const newVal = e.target.innerText.trim().toUpperCase();
            
            if (modalState.dataset.data[pIdx].shifts[sIdx] !== newVal) {
                modalState.dataset.data[pIdx].shifts[sIdx] = newVal;
                revalidateAndRefresh(title);
            }
        });
        cell.addEventListener('keydown', e => {
            if (e.key === 'Enter') {
                e.preventDefault();
                e.target.blur();
            }
        });
    });

    document.getElementById('closeM').onclick = () => {
        m.remove();
        tip.remove();
        const style = document.getElementById('kmuh-modal-style'); if (style) style.remove();
        chrome.runtime.sendMessage({ action: "modalClosed" });
    };

    document.getElementById('saveM').onclick = async () => {
        const proceed = confirm("確定要將目前修改後的班表寫入網頁嗎？");
        if (!proceed) return;
        
        const excelMap = {};
        modalState.dataset.data.forEach(p => {
            excelMap[p.empId] = { name: p.name, shifts: p.shifts };
        });
        
        const res = await executeInjectionFlowFromMap(excelMap);
        if (res.success) {
            alert("班表寫入完成！");
            document.getElementById('closeM').click();
        } else {
            alert("寫入失敗：" + (res.message || "未知錯誤"));
        }
    };
}

function revalidateAndRefresh(title) {
    const { dataset, storage, hrTimeMap, cycleRanges, ffRanges, oldMonthDays, newMonthDays, oldYymm, targetYymm } = modalState;
    const excelMap = {};
    dataset.data.forEach(p => {
        excelMap[p.empId] = { name: p.name, shifts: p.shifts };
    });
    const customDict = storage.shiftDict || [];
    const lastData = storage.lastMonthData;
    const check = runDetailedCheck(lastData, excelMap, customDict, hrTimeMap, cycleRanges, ffRanges, oldMonthDays, newMonthDays, oldYymm, targetYymm);
    modalState.dataset.errors = check.errors;
    renderModalContent(title);
}

async function executeInjectionFlowFromMap(excelMap) {
    const storage = await chrome.storage.local.get(['shiftDict', 'blankFillMode', 'blankFillCode']);
    const customDict = storage.shiftDict || [];
    const isFill = (storage.blankFillMode || 'keep') === 'fill' && storage.blankFillCode;
    const fillCode = storage.blankFillCode || '';

    const webMap = {};
    document.querySelectorAll("input[id^='Hidden_empno_']").forEach(f => {
        const empId = formatEmpId(f.value.split('-')[0]);
        if (empId) webMap[empId] = f.id.split('_').pop();
    });

    for (let id in excelMap) {
        const sfx = webMap[formatEmpId(id)];
        if (!sfx) continue;
        excelMap[id].shifts.forEach((code, i) => {
            const el = document.getElementById(`Field_day${String(i + 1).padStart(2, '0')}_${sfx}`);
            if (el) {
                // 1. 處理空白填補與不覆蓋邏輯
                let finalCode = code;
                
                // 如果 Excel/編輯器是空白，且未開啟填補模式，則跳過此格（不覆蓋網頁原本內容）
                if (!finalCode && !isFill) {
                    return; // 跳過，維持網頁原本班別
                }
                
                // 如果 Excel/編輯器是空白，但開啟了填補模式，則使用填補代號
                if (!finalCode && isFill) {
                    finalCode = fillCode;
                }
                
                // 2. 透過班別字典轉換 (Excel 代號 -> 系統代號)
                const dictEntry = customDict.find(x => String(x.excel).trim() === String(finalCode).trim());
                if (dictEntry && dictEntry.sys) {
                    finalCode = dictEntry.sys;
                }

                // 3. 執行寫入
                if (el.value !== finalCode) {
                    el.value = finalCode;
                    el.style.backgroundColor = "#fff3cd"; // 標記已修改
                }
            }
        });
    }
    return { success: true };
}

// ─────────────────────────────────────────────────────────────────
// 網頁班表擷取
// ─────────────────────────────────────────────────────────────────
function captureWebSchedule() {
    const h = getHeaders();
    const d = h.dates.filter(x => x !== "").length;
    const yymm = document.getElementById("ctl00_ContentPlaceHolder1_FIELD_yymm")?.value || "";
    const res = [];
    document.querySelectorAll("input[id^='Hidden_empno_']").forEach(f => {
        const sfx = f.id.split('_').pop();
        const parts = f.value.split('-');
        const empId = formatEmpId(parts[0]?.trim());
        const name  = parts[1]?.trim() || "";
        const shifts = [];
        for (let i = 1; i <= d; i++) {
            const el = document.getElementById(`Field_day${String(i).padStart(2, '0')}_${sfx}`);
            shifts.push(el ? el.value : "");
        }
        res.push({ empId, name, shifts });
    });
    return { headers: h, data: res, monthDays: d, yymm };
}

function getHeaders() {
    const w = Array(31).fill(""), d = Array(31).fill("");
    const td = Array.from(document.querySelectorAll("td")).find(t => t.innerText.trim() === "01");
    if (td) {
        const r = td.parentElement, wr = r.previousElementSibling, idx = Array.from(r.children).indexOf(td);
        for (let i = 0; i < 31; i++) {
            const dt = r.children[idx + i];
            if (dt && /^\d+$/.test(dt.innerText.trim())) {
                d[i] = dt.innerText.trim();
                if (wr?.children[idx + i]) w[i] = wr.children[idx + i].innerText.trim();
            }
        }
    }
    return { weekdays: w, dates: d };
}

// ─────────────────────────────────────────────────────────────────
// Excel 解析
// ─────────────────────────────────────────────────────────────────
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

function detectExcelLayout(data, targetYymm) {
    const targetMonth = parseInt(targetYymm.substring(4, 6));
    const targetYear  = parseInt(targetYymm.substring(0, 4));
    const monthDays   = new Date(targetYear, targetMonth, 0).getDate();
    let empIdColIdx = -1, nameColIdx = -1, day1ColIdx = -1;

    const EMP_KEYWORDS  = ["職編", "員工編號", "工號", "員編", "職員編號"];
    const NAME_KEYWORDS = ["姓名", "員工姓名", "名字"];

    for (let ri = 0; ri < Math.min(10, data.length); ri++) {
        const row = data[ri];
        if (!row) continue;
        for (let ci = 0; ci < row.length; ci++) {
            const val = String(row[ci] || "").trim();
            if (empIdColIdx === -1 && EMP_KEYWORDS.some(k => val.includes(k))) empIdColIdx = ci;
            if (nameColIdx  === -1 && NAME_KEYWORDS.some(k => val.includes(k))) nameColIdx  = ci;
            if (day1ColIdx  === -1) {
                const cd  = parseCellDate(row[ci]);
                const cd2 = parseCellDate(row[ci + 1]);
                if (cd?.day === 1 && cd2?.day === 2) day1ColIdx = ci;
            }
        }
        if (empIdColIdx !== -1 && nameColIdx !== -1 && day1ColIdx !== -1) break;
    }

    if (empIdColIdx === -1) {
        const colHits = {};
        for (let ri = 0; ri < data.length; ri++) {
            const row = data[ri];
            if (!row) continue;
            for (let ci = 0; ci < (day1ColIdx !== -1 ? day1ColIdx : row.length); ci++) {
                const val = String(row[ci] || "").trim();
                if (/^\d{6,7}$/.test(val)) {
                    colHits[ci] = (colHits[ci] || 0) + 1;
                }
            }
        }
        let bestCol = -1, bestHits = 1;
        for (const [ci, hits] of Object.entries(colHits)) {
            if (hits > bestHits) { bestHits = hits; bestCol = parseInt(ci); }
        }
        if (bestCol !== -1) {
            empIdColIdx = bestCol;
        }
    }

    if (nameColIdx === -1 && empIdColIdx !== -1) nameColIdx = empIdColIdx + 1;
    return {
        empIdColIdx: empIdColIdx !== -1 ? empIdColIdx : 1,
        nameColIdx:  nameColIdx  !== -1 ? nameColIdx  : 2,
        day1ColIdx:  day1ColIdx  !== -1 ? day1ColIdx  : 3,
        monthDays
    };
}

function parseExcel(data, targetYymm) {
    const { empIdColIdx, nameColIdx, day1ColIdx, monthDays } = detectExcelLayout(data, targetYymm);
    const m = {};
    data.forEach(r => {
        const rawId = String(r[empIdColIdx] || "").trim();
        if (!/^\d{6,7}$/.test(rawId)) return;
        const empId  = formatEmpId(rawId);
        const name   = String(r[nameColIdx] || "").trim();
        const shifts = [];
        for (let i = 0; i < monthDays; i++) {
            const val = r[day1ColIdx + i];
            shifts.push(val !== undefined && val !== null ? String(val).trim() : "");
        }
        m[empId] = { name, shifts };
    });
    return m;
}

async function executeInjectionFlow(excelData) {
    const storage = await chrome.storage.local.get(['lastMonthData', 'shiftDict', 'blankFillMode', 'blankFillCode']);
    const oldYymm = storage.lastMonthData?.yymm || "";
    const excelMap = parseExcel(excelData, oldYymm ? getNextYM(oldYymm) : "");
    return executeInjectionFlowFromMap(excelMap);
}
