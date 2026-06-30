// content.js
// 以 world: "ISOLATED"（預設）注入，可正常使用 chrome.* API。
// IE 相容 polyfill 已移至 content_main.js（world: "MAIN"）處理。
// 頁面驗證統一由 popup.js 的 sendMessage 負責，content.js 不重複處理。
// HR 預設班別統一由 background.js 初始化至 storage，content.js 不再 hardcode。

console.log("🚀 [KMUH Helper] 核心啟動 (ISOLATED World)");

// ─────────────────────────────────────────────────────────────────
// 純函式 / 工具
// ─────────────────────────────────────────────────────────────────
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

// ── 合併後的週期解析（原 parseCyclePeriods / parseFFPeriods） ─────
// 傳入不同括號即可區分四週變形（【】）與 FF 雙週（《》）
function parsePeriods(bracketOpen, bracketClose) {
    const esc = (c) => c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(
        `${esc(bracketOpen)}(\\d+)${esc(bracketClose)}\\s*(\\d{1,2}\\/\\d{1,2})\\s*[~～]\\s*(\\d{1,2}\\/\\d{1,2})`,
        'g'
    );
    const periods = [];
    let m;
    while ((m = re.exec(document.body.innerText)) !== null) {
        periods.push({ label: m[1], start: m[2], end: m[3] });
    }
    return periods;
}
const parseCyclePeriods = () => parsePeriods('【', '】');
const parseFFPeriods    = () => parsePeriods('《', '》');

// ── hrShifts 陣列 → {code: {start, end}} 查找表 ──────────────────
function buildHrTimeMap(hrShiftsRaw) {
    const map = {};
    (hrShiftsRaw || []).forEach(x => {
        if (typeof x === 'object' && x.code) {
            map[x.code] = { start: x.start || null, end: x.end || null };
        }
    });
    return map;
}

// ── lastMonthData → 月份衍生值（避免多處重複計算） ────────────────
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

function mmddToDate(mmdd, refYymm) {
    const [mm, dd] = mmdd.split('/').map(Number);
    const refYear  = parseInt(refYymm.substring(0, 4));
    const refMonth = parseInt(refYymm.substring(4, 6));
    const year = (mm < refMonth - 6) ? refYear + 1 : refYear;
    return new Date(year, mm - 1, dd);
}

function dateToMmdd(d) {
    return `${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;
}

function mmddToGlobalIdx(mmdd, oldYymm, oldMonthDays) {
    const base   = mmddToDate(`${oldYymm.substring(4, 6)}/01`, oldYymm);
    const target = mmddToDate(mmdd, oldYymm);
    return Math.round((target - base) / 86400000);
}

function buildCheckRanges(lastPeriod, targetMonth, periodDays, oldYymm, oldMonthDays) {
    if (!lastPeriod) return [];
    const ranges = [];
    let startDate = mmddToDate(lastPeriod.start, oldYymm);
    let endDate   = mmddToDate(lastPeriod.end,   oldYymm);
    if (endDate < startDate) endDate.setFullYear(endDate.getFullYear() + 1);
    while (true) {
        const startMonth = startDate.getMonth() + 1;
        const endMonth   = endDate.getMonth() + 1;
        if (startMonth > targetMonth) break;
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

    if (request.action === "readAndMemorize") {
        const data = captureWebSchedule();
        const now  = new Date();
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
        const toSave  = { lastMonthData: data };
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
            sendResponse({ success: true, yymm: data.yymm, nextUrl, hasPreview: request.showPreview, periods, ffPeriods });
        });
        return true;
    }

    if (request.action === "autoProcessExcel") {
        handleExcelProcess(request).then(res => sendResponse(res));
        return true;
    }

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
    const { oldYymm, targetYymm, targetMonth, oldMonthDays, newMonthDays } =
        deriveMonthContext(storage.lastMonthData);

    const excelMap = parseExcel(req.excelData, targetYymm);
    if (excelMap.error) return { success: false, message: excelMap.message };

    const customDict  = storage.shiftDict  || [];
    const hrShiftsRaw = storage.hrShifts   || [];
    const lastData    = storage.lastMonthData;
    const hrShiftsList = hrShiftsRaw.map(x => typeof x === 'string' ? x : x.code);
    const hrTimeMap    = buildHrTimeMap(hrShiftsRaw);

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
    if (unknownCodes.size > 0) return { success: false, unknownCodes: Array.from(unknownCodes) };

    const dataWithId  = Object.entries(excelMap).map(([id, v]) => ({ empId: id, ...v }));
    const lastCycle   = (lastData?.cyclePeriods || []).at(-1) || null;
    const lastFF      = (lastData?.ffPeriods    || []).at(-1) || null;
    const cycleRanges = buildCheckRanges(lastCycle, targetMonth, 28, oldYymm, oldMonthDays);
    const ffRanges    = buildCheckRanges(lastFF,    targetMonth, 14, oldYymm, oldMonthDays);

    const allRanges = [...cycleRanges, ...ffRanges];
    const biStart   = allRanges.length > 0 ? Math.min(...allRanges.map(r => r.startIdx)) : oldMonthDays;
    const biEnd     = allRanges.length > 0 ? Math.max(...allRanges.map(r => r.endIdx))   : oldMonthDays + 27;

    const cycleLabel = cycleRanges.map((r, i) => `【${i + 1}】${r.start}～${r.end}`).join('、') || '未知';
    const ffLabel    = ffRanges.map((r, i)    => `《${i + 1}》${r.start}～${r.end}`).join('、') || '未知';
    const nhRequired = parseInt(
        document.getElementById('ctl00_ContentPlaceHolder1_lbncount')?.textContent?.trim() || '0'
    , 10) || 0;
    const nhLabel  = nhRequired > 0 ? `　／　NH/N+ 應排：${nhRequired} 天` : '';
    const infoText = `四週變形：${cycleLabel}　／　FF雙週：${ffLabel}${nhLabel}`;

    const check = runDetailedCheck(lastData, excelMap, customDict, hrTimeMap, cycleRanges, ffRanges, oldMonthDays, newMonthDays, oldYymm, targetYymm, nhRequired);
    if (req.showReport || check.errors.length > 0) {
        showModal("Excel 班表預覽與檢測報告", {
            headers: getHeaders(), data: dataWithId, errors: check.errors,
            monthDays: oldMonthDays, biStart, biEnd, cycleRanges, ffRanges, nhRequired,
            blankFillMode: req.blankFillMode || 'keep',
            blankFillCode: req.blankFillCode || '',
        }, infoText);
    }
    return { success: check.errors.length === 0, noOldDataWarnings: check.noOldDataWarnings };
}

// ─────────────────────────────────────────────────────────────────
// 檢測工具
// ─────────────────────────────────────────────────────────────────
function timeToMinutes(t) {
    if (!t) return null;
    const [h, m] = t.split(':').map(Number);
    return h * 60 + (m || 0);
}

// 修正原版：移除無意義的 dict.find（不論找沒找到都 return null），並移除多餘的 dict 參數
function getShiftTime(code, hrTimeMap) {
    const entry = hrTimeMap[code];
    if (!entry) return null;
    const { start, end } = entry;
    if (!start && !end) return null;
    const startMin = timeToMinutes(start);
    let   endMin   = timeToMinutes(end);
    if (endMin !== null && startMin !== null && endMin <= startMin) endMin += 1440;
    return { startMin, endMin };
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

function runDetailedCheck(old, exc, dict, hrTimeMap, cycleRanges, ffRanges, oldMonthDays, newMonthDays, oldYymm, targetYymm, nhRequired = 0) {
    if (!old?.data && !exc) return { errors: [], noOldDataWarnings: [] };
    const err = [], noOldDataWarnings = [];
    const toDate = (gi) => giToDateStr(gi, oldYymm, targetYymm, oldMonthDays);

    for (let id in exc) {
        const oStf       = old?.data?.find(p => formatEmpId(p.empId) === formatEmpId(id));
        const hasOldData = !!oStf;
        const validStart = hasOldData ? 0 : oldMonthDays;
        const validEnd   = oldMonthDays + newMonthDays - 1;

        if (!hasOldData) noOldDataWarnings.push({ empId: id, name: exc[id].name || '' });

        const oldShifts      = hasOldData ? oStf.shifts : Array(oldMonthDays).fill('');
        const rawExcelShifts = exc[id].shifts;
        const combined = [...oldShifts, ...rawExcelShifts].map(s => {
            const d = dict.find(x => String(x.excel).trim() === String(s).trim());
            return d ? d.sys : s;
        });

        // W+ / N+ 建議更換提醒
        for (let i = 0; i < rawExcelShifts.length; i++) {
            const rawCode = String(rawExcelShifts[i] || "").trim();
            const gi = oldMonthDays + i;
            if (rawCode === 'W+' || rawCode === 'N+') {
                err.push({ empId: id, startIdx: gi, endIdx: gi, type: 'REPLACE_REQUIRED',
                    msg: `⚠️ 建議更換：${toDate(gi)} Excel 原始代號為 ${rawCode}，請更換為正確的加班代號（如加班小時別）。` });
            }
        }

        const isRangeValid = (r) => r.startIdx >= validStart && r.endIdx <= validEnd;

        // FF 雙週檢查
        ffRanges.forEach((r, i) => {
            if (!isRangeValid(r)) return;
            const count = combined.slice(r.startIdx, r.endIdx + 1).filter(s => s === 'FF').length;
            if (count !== 2) err.push({ empId: id, startIdx: r.startIdx, endIdx: r.endIdx, type: `FF_${i + 1}`,
                msg: `FF雙週《${i + 1}》${r.start}～${r.end} FF=${count}（應2）` });
        });

        // FF 間隔檢查 (不可超過 12 天)
        const ffIndices = [];
        for (let gi = validStart; gi <= validEnd; gi++) {
            if (combined[gi] === 'FF') ffIndices.push(gi);
        }
        for (let fi = 0; fi < ffIndices.length - 1; fi++) {
            const gap = ffIndices[fi + 1] - ffIndices[fi] - 1;
            if (gap > 12) err.push({ empId: id, startIdx: ffIndices[fi], endIdx: ffIndices[fi + 1], type: 'FF_GAP',
                msg: `FF間隔過長：${toDate(ffIndices[fi])}(FF) 與 ${toDate(ffIndices[fi + 1])}(FF) 之間間隔 ${gap} 天（最多12天）` });
        }

        // 四週變形檢查 (WW+W+ 應為 4 天)
        cycleRanges.forEach((r, i) => {
            if (!isRangeValid(r)) return;
            const count = combined.slice(r.startIdx, r.endIdx + 1).filter(s => s === 'WW' || s === 'W+').length;
            if (count !== 4) err.push({ empId: id, startIdx: r.startIdx, endIdx: r.endIdx, type: `WW_${i + 1}`,
                msg: `四週變形【${i + 1}】${r.start}～${r.end} WW=${count}（應4）` });
        });

        // NH / N+ 整月天數檢查
        if (nhRequired > 0) {
            const nhCount = combined.slice(oldMonthDays, oldMonthDays + newMonthDays)
                .filter(s => s === 'NH' || s === 'N+').length;
            if (nhCount !== nhRequired) err.push({ empId: id, startIdx: oldMonthDays, endIdx: oldMonthDays + newMonthDays - 1, type: 'NH_COUNT',
                msg: `NH/N+ 天數不符：實際排 ${nhCount} 天（本月應排 ${nhRequired} 天）` });
        }

        // 接班間隔檢查 (應達 11 小時)
        let prevCode = null, prevEndMin = null, prevGi = -1;
        for (let gi = Math.max(0, validStart - 1); gi <= validEnd; gi++) {
            const code = combined[gi] || '';
            if (!code) continue;
            const timeInfo = getShiftTime(code, hrTimeMap);
            if (!timeInfo) { prevCode = null; prevEndMin = null; prevGi = -1; continue; }
            const { startMin, endMin } = timeInfo;
            if (prevEndMin !== null) {
                const daysBetween  = gi - prevGi - 1;
                const gap = (prevGi * 1440 + (daysBetween + 1) * 1440 + startMin) - (prevGi * 1440 + prevEndMin);
                if (gap < 660) {
                    const gapH   = Math.floor(Math.max(gap, 0) / 60);
                    const gapM   = Math.max(gap, 0) % 60;
                    const gapStr = gap <= 0 ? '0分（班別重疊）' : (gapM > 0 ? `${gapH}小時${gapM}分` : `${gapH}小時`);
                    err.push({ empId: id, startIdx: prevGi, endIdx: gi, type: 'REST_SHORT',
                        msg: `接班間距不足：${toDate(prevGi)}(${prevCode}) 與 ${toDate(gi)}(${code}) 間距僅 ${gapStr}（未達11小時）` });
                }
            }
            prevCode = code; prevEndMin = endMin; prevGi = gi;
        }
    }
    return { errors: err, noOldDataWarnings };
}

// ─────────────────────────────────────────────────────────────────
// UI：Modal 報告視窗
// ─────────────────────────────────────────────────────────────────
let modalState = {
    dataset: null, info: '', storage: null, hrTimeMap: {},
    oldYymm: '', targetYymm: '', oldMonthDays: 0, newMonthDays: 0,
    cycleRanges: [], ffRanges: [], nhRequired: 0,
};

async function showModal(title, dataset, info) {
    const oldModal = document.getElementById('kmuh-modal'); if (oldModal) oldModal.remove();
    const oldStyle = document.getElementById('kmuh-modal-style'); if (oldStyle) oldStyle.remove();

    const storage   = await chrome.storage.local.get(['shiftDict', 'hrShifts', 'lastMonthData']);
    const hrTimeMap = buildHrTimeMap(storage.hrShifts);
    const { oldYymm, targetYymm, targetMonth, oldMonthDays, newMonthDays } =
        deriveMonthContext(storage.lastMonthData);

    const lastCycle   = (storage.lastMonthData?.cyclePeriods || []).at(-1) || null;
    const lastFF      = (storage.lastMonthData?.ffPeriods    || []).at(-1) || null;
    const cycleRanges = buildCheckRanges(lastCycle, targetMonth, 28, oldYymm, oldMonthDays);
    const ffRanges    = buildCheckRanges(lastFF,    targetMonth, 14, oldYymm, oldMonthDays);

    modalState = { dataset, info, storage, hrTimeMap, oldYymm, targetYymm, oldMonthDays, newMonthDays, cycleRanges, ffRanges, nhRequired: dataset.nhRequired || 0 };
    renderModalContent(title);
}

// ── 錯誤顏色對應表（移出迴圈，僅定義一次） ────────────────────────
const ERR_COLOR_MAP = {
    WW:               { border: '#e74c3c', bg: '#fff2f2' },
    FF:               { border: '#e74c3c', bg: '#fff2f2' },
    GAP:              { border: '#e67e22', bg: '#fff8f0' },
    REST:             { border: '#8e44ad', bg: '#fdf2ff' },
    REPLACE_REQUIRED: { border: '#f39c12', bg: '#fef5e7' },
    NH:               { border: '#0f6e56', bg: '#e1f5ee' },
};

function getErrColor(type) {
    if (!type)                       return ERR_COLOR_MAP.WW;
    if (type === 'REPLACE_REQUIRED') return ERR_COLOR_MAP.REPLACE_REQUIRED;
    if (type === 'FF_GAP')           return ERR_COLOR_MAP.GAP;
    if (type === 'REST_SHORT')       return ERR_COLOR_MAP.REST;
    if (type === 'NH_COUNT')         return ERR_COLOR_MAP.NH;
    if (type.startsWith('FF_'))      return ERR_COLOR_MAP.FF;
    return ERR_COLOR_MAP.WW;
}

function renderModalContent(title) {
    const { dataset, info, oldMonthDays, cycleRanges, ffRanges } = modalState;
    const h        = dataset.headers;
    const mDays    = oldMonthDays;
    const total    = dataset.data.length;
    const errorIds = new Set(dataset.errors?.map(e => formatEmpId(e.empId)));
    const errCount = errorIds.size;

    const CYCLE_COLORS = ['#dbeafe', '#bfdbfe', '#93c5fd'];
    const FF_COLORS    = ['#ede9fe', '#ddd6fe', '#c4b5fd'];

    const cycleCss = cycleRanges.map((_, i) =>
        `.hd-cy-${i} { background:${CYCLE_COLORS[i % CYCLE_COLORS.length]} !important; }`).join('\n');
    const ffCss = ffRanges.map((_, i) =>
        `.hd-ff-${i} { background:${FF_COLORS[i % FF_COLORS.length]} !important; }`).join('\n');

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
            `<span style="display:inline-flex;align-items:center;gap:3px;margin-right:8px;"><span style="display:inline-block;width:12px;height:12px;background:${CYCLE_COLORS[i % CYCLE_COLORS.length]};border:1px solid #aaa;border-radius:2px;"></span>四週【${i + 1}】${r.start}～${r.end}</span>`),
        ...ffRanges.map((r, i) =>
            `<span style="display:inline-flex;align-items:center;gap:3px;margin-right:8px;"><span style="display:inline-block;width:12px;height:12px;background:${FF_COLORS[i % FF_COLORS.length]};border:1px solid #aaa;border-radius:2px;"></span>FF《${i + 1}》${r.start}～${r.end}</span>`),
    ].join('');

    const errLegend = [
        { color: '#e74c3c', bg: '#fff2f2', label: '四週變形/FF數量錯誤' },
        { color: '#e67e22', bg: '#fff8f0', label: 'FF間隔超過12天' },
        { color: '#8e44ad', bg: '#fdf2ff', label: '接班間距不足11小時' },
        { color: '#f39c12', bg: '#fef5e7', label: '建議更換 W+/N+' },
        { color: '#0f6e56', bg: '#e1f5ee', label: 'NH/N+ 天數不符' },
    ].map(x =>
        `<span style="display:inline-flex;align-items:center;gap:3px;margin-right:10px;"><span style="display:inline-block;width:24px;height:14px;background:${x.bg};border:2px solid ${x.color};border-radius:2px;"></span>${x.label}</span>`
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
        .sticky-col  { position:sticky; left:0;    background:#f8f9fa !important; z-index:5; font-weight:bold; border-right:2px solid #bdc3c7 !important; min-width:70px; }
        .sticky-name { position:sticky; left:71px; background:#f8f9fa !important; z-index:5; font-weight:bold; border-right:2px solid #bdc3c7 !important; min-width:60px; }
        .cell-err { background:#fff2f2 !important; border:2px solid #ff7675 !important; }
        .tooltip { position:relative; cursor:help; }
        #kmuh-tip { position:fixed; background:#2d3436; color:white; padding:8px 14px; border-radius:6px; font-size:12px; z-index:99999; pointer-events:none; display:none; box-shadow:0 4px 12px rgba(0,0,0,0.4); }
        .editable-cell:focus { outline: 2px solid #3498db; background: #fff !important; }
        ${cycleCss} ${ffCss}
    `;
    document.head.appendChild(style);

    const thW = h.weekdays.map((w, i) =>
        `<th class="${colCls(mDays + i)}" style="color:${w === '日' || w === '六' ? '#e74c3c' : 'inherit'}">${w}</th>`).join('');
    const thD = h.dates.map((d, i) =>
        `<th class="${colCls(mDays + i)}">${d}</th>`).join('');

    const isFill = dataset.blankFillMode === 'fill' && dataset.blankFillCode;

    const rows = dataset.data.map((p, pIdx) => {
        const pErrs = dataset.errors?.filter(e => formatEmpId(p.empId) === formatEmpId(e.empId)) || [];
        const cells = p.shifts.map((s, i) => {
            const gi         = mDays + i;
            const isBlank    = !s;
            const displayVal = isBlank && isFill
                ? `<span style="color:#e67e22;font-size:11px;">→${dataset.blankFillCode}</span>`
                : (s || '');
            const cellErrs   = pErrs.filter(e => gi >= e.startIdx && gi <= e.endIdx);
            let borderStyle  = '', bgStyle = '', tipText = '';
            if (cellErrs.length > 0) {
                const bigErr  = cellErrs.reduce((a, b) => (b.endIdx - b.startIdx) > (a.endIdx - a.startIdx) ? b : a);
                const { border, bg } = getErrColor(bigErr.type);
                const isFirst = gi === bigErr.startIdx, isLast = gi === bigErr.endIdx;
                borderStyle = `border-top:2px solid ${border} !important; border-bottom:2px solid ${border} !important;`
                    + (isFirst ? `border-left:2px solid ${border} !important;`  : 'border-left:none !important;')
                    + (isLast  ? `border-right:2px solid ${border} !important;` : 'border-right:none !important;');
                bgStyle  = `background:${bg} !important;`;
                tipText  = pErrs.map(e => e.msg).join('\n');
            } else if (isBlank && isFill) {
                tipText = `將填入 ${dataset.blankFillCode}`;
            }
            const wkBg   = h.weekdays[i] === '日' || h.weekdays[i] === '六' ? '#fef9f9' : 'white';
            const cellBg = cellErrs.length > 0 ? '' : `background:${wkBg};`;
            const tipAttr = tipText ? `data-kmuh-tip="${tipText.replace(/"/g, '&quot;')}"` : '';
            const cls     = (tipText ? 'tooltip ' : '') + 'editable-cell';
            return `<td class="${cls}" ${tipAttr} contenteditable="true" data-p-idx="${pIdx}" data-s-idx="${i}" style="${cellBg}${bgStyle}${borderStyle}">${displayVal}</td>`;
        }).join('');
        return `<tr><td class="sticky-col">${p.empId || ''}</td><td class="sticky-name">${p.name || ''}</td>${cells}</tr>`;
    }).join('');

    let m = document.getElementById('kmuh-modal');
    if (!m) { m = document.createElement('div'); m.id = 'kmuh-modal'; document.body.appendChild(m); }

    m.innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <h2 style="margin:0;">📊 ${title}</h2>
            <div style="display:flex; gap:10px;">
                <button id="saveM"  style="padding:10px 35px; background:#2ecc71; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold; font-size:14px;">💾 寫入班表</button>
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
                    <tr style="background:#f1f2f6;"><th rowspan="2" class="sticky-col">職編</th><th rowspan="2" class="sticky-name">姓名</th>${thW}</tr>
                    <tr style="background:#f1f2f6;">${thD}</tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>
        </div>`;

    setupModalEvents(m, title);
}

function setupModalEvents(m, title) {
    const tip = document.getElementById('kmuh-tip') || document.createElement('div');
    if (!tip.id) { tip.id = 'kmuh-tip'; document.body.appendChild(tip); }

    const showTip = e => {
        const td = e.target.closest('[data-kmuh-tip]');
        if (!td) return;
        tip.innerHTML = td.getAttribute('data-kmuh-tip').split('\n')
            .map(l => `<div style="white-space:nowrap; line-height:1.8;">${l}</div>`).join('');
        tip.style.display = 'block';
    };
    const moveTip = e => {
        if (tip.style.display === 'none') return;
        const x = e.clientX + 14, y = e.clientY - tip.offsetHeight - 10;
        tip.style.left = Math.min(x, window.innerWidth - tip.offsetWidth - 10) + 'px';
        tip.style.top  = Math.max(y, 10) + 'px';
    };
    const hideTip = () => { tip.style.display = 'none'; };

    m.addEventListener('mouseover',  showTip);
    m.addEventListener('mousemove',  moveTip);
    m.addEventListener('mouseleave', hideTip);
    m.addEventListener('mouseout', e => { if (!e.target.closest('[data-kmuh-tip]')) hideTip(); });

    m.querySelectorAll('.editable-cell').forEach(cell => {
        cell.addEventListener('blur', e => {
            const pIdx   = parseInt(e.target.dataset.pIdx);
            const sIdx   = parseInt(e.target.dataset.sIdx);
            const newVal = e.target.innerText.trim().toUpperCase();
            if (modalState.dataset.data[pIdx].shifts[sIdx] !== newVal) {
                modalState.dataset.data[pIdx].shifts[sIdx] = newVal;
                revalidateAndRefresh(title);
            }
        });
        cell.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); e.target.blur(); } });
    });

    document.getElementById('closeM').onclick = () => {
        m.remove(); tip.remove();
        const style = document.getElementById('kmuh-modal-style'); if (style) style.remove();
        chrome.runtime.sendMessage({ action: "modalClosed" });
    };

    document.getElementById('saveM').onclick = async () => {
        if (!confirm("確定要將目前修改後的班表寫入網頁嗎？")) return;
        const excelMap = {};
        modalState.dataset.data.forEach(p => { excelMap[p.empId] = { name: p.name, shifts: p.shifts }; });
        const res = await executeInjectionFlowFromMap(excelMap);
        if (res.success) { alert("班表寫入完成！"); document.getElementById('closeM').click(); }
        else alert("寫入失敗：" + (res.message || "未知錯誤"));
    };
}

function revalidateAndRefresh(title) {
    const { dataset, storage, hrTimeMap, cycleRanges, ffRanges, oldMonthDays, newMonthDays, oldYymm, targetYymm, nhRequired } = modalState;
    const excelMap = {};
    dataset.data.forEach(p => { excelMap[p.empId] = { name: p.name, shifts: p.shifts }; });
    const check = runDetailedCheck(storage.lastMonthData, excelMap, storage.shiftDict || [], hrTimeMap, cycleRanges, ffRanges, oldMonthDays, newMonthDays, oldYymm, targetYymm, nhRequired || 0);
    modalState.dataset.errors = check.errors;
    renderModalContent(title);
}

// ─────────────────────────────────────────────────────────────────
// 網頁班表擷取
// ─────────────────────────────────────────────────────────────────
function captureWebSchedule() {
    const h    = getHeaders();
    const d    = h.dates.filter(x => x !== "").length;
    const yymm = document.getElementById("ctl00_ContentPlaceHolder1_FIELD_yymm")?.value || "";
    const res  = [];
    document.querySelectorAll("input[id^='Hidden_empno_']").forEach(f => {
        const sfx   = f.id.split('_').pop();
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
    if (/^\d{1,2}$/.test(s)) { const n = parseInt(s); if (n >= 1 && n <= 31) return { month: null, day: n }; }
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
            if (empIdColIdx === -1 && EMP_KEYWORDS.some(k => val.includes(k)))  empIdColIdx = ci;
            if (nameColIdx  === -1 && NAME_KEYWORDS.some(k => val.includes(k))) nameColIdx  = ci;
            if (day1ColIdx  === -1) {
                const cd = parseCellDate(row[ci]), cd2 = parseCellDate(row[ci + 1]);
                if (cd?.day === 1 && cd2?.day === 2) day1ColIdx = ci;
            }
        }
        if (empIdColIdx !== -1 && nameColIdx !== -1 && day1ColIdx !== -1) break;
    }
    if (empIdColIdx === -1) {
        const colHits = {};
        for (let ri = 0; ri < data.length; ri++) {
            const row = data[ri]; if (!row) continue;
            for (let ci = 0; ci < (day1ColIdx !== -1 ? day1ColIdx : row.length); ci++) {
                const val = String(row[ci] || "").trim();
                if (/^\d{6,7}$/.test(val)) colHits[ci] = (colHits[ci] || 0) + 1;
            }
        }
        let bestCol = -1, bestHits = 1;
        for (const [ci, hits] of Object.entries(colHits)) {
            if (hits > bestHits) { bestHits = hits; bestCol = parseInt(ci); }
        }
        if (bestCol !== -1) empIdColIdx = bestCol;
    }
    if (nameColIdx === -1 && empIdColIdx !== -1) nameColIdx = empIdColIdx + 1;
    return { empIdColIdx, nameColIdx, day1ColIdx, monthDays, isFormatValid: empIdColIdx !== -1 && day1ColIdx !== -1 };
}

function parseExcel(data, targetYymm) {
    const layout = detectExcelLayout(data, targetYymm);
    if (!layout.isFormatValid) {
        return { error: "INVALID_FORMAT", message: "❌ 無法辨識 Excel 格式。\n請確認檔案中是否包含「職編」關鍵字，以及「1號」日期欄位。" };
    }
    const m = {};
    data.forEach(r => {
        const rawId = String(r[layout.empIdColIdx] || "").trim();
        if (!/^\d{6,7}$/.test(rawId)) return;
        const empId  = formatEmpId(rawId);
        const name   = String(r[layout.nameColIdx] || "").trim();
        const shifts = [];
        for (let i = 0; i < layout.monthDays; i++) {
            let val = r[layout.day1ColIdx + i];
            val = (val !== undefined && val !== null) ? String(val).replace(/[\r\n]/g, '').trim() : "";
            shifts.push(val);
        }
        m[empId] = { name, shifts };
    });
    if (Object.keys(m).length === 0) {
        return { error: "NO_DATA", message: "❌ 格式辨識成功，但未找到任何有效的員工資料列。\n請確認職編是否為 6~7 位數字。" };
    }
    return m;
}

async function executeInjectionFlow(excelData) {
    const storage = await chrome.storage.local.get(['lastMonthData', 'shiftDict', 'blankFillMode', 'blankFillCode']);
    const { oldYymm } = deriveMonthContext(storage.lastMonthData);
    const excelMap    = parseExcel(excelData, oldYymm ? getNextYM(oldYymm) : "");
    return executeInjectionFlowFromMap(excelMap);
}

async function executeInjectionFlowFromMap(excelMap) {
    const storage    = await chrome.storage.local.get(['shiftDict', 'blankFillMode', 'blankFillCode']);
    const customDict = storage.shiftDict || [];
    const isFill     = (storage.blankFillMode || 'keep') === 'fill' && storage.blankFillCode;
    const fillCode   = storage.blankFillCode || '';

    const webMap = {};
    document.querySelectorAll("input[id^='Hidden_empno_']").forEach(f => {
        const empId = formatEmpId(f.value.split('-')[0]);
        if (empId) webMap[empId] = f.id.split('_').pop();
    });

    for (let id in excelMap) {
        const sfx = webMap[formatEmpId(id)];
        if (!sfx) continue;
        excelMap[id].shifts.forEach((code, i) => {
            const dd = String(i + 1).padStart(2, '0');
            const el = document.getElementById(`Field_day${dd}_${sfx}`);
            if (!el) return;
            let finalCode = code;
            if (!finalCode && !isFill) return;
            if (!finalCode && isFill)  finalCode = fillCode;

            const dictEntry = customDict.find(x => String(x.excel).trim() === String(finalCode).trim());
            let overCode = '', amCode = '', pmCode = '', nightCode = '';
            if (dictEntry && dictEntry.sys) {
                finalCode = dictEntry.sys;
                overCode  = String(dictEntry.over  || '').trim();
                amCode    = String(dictEntry.am    || '').trim();
                pmCode    = String(dictEntry.pm    || '').trim();
                nightCode = String(dictEntry.night || '').trim();
            }
            if (el.value !== finalCode) { el.value = finalCode; el.style.backgroundColor = "#fff3cd"; }

            // 逾時欄位：只要 Excel 有填班別就寫入（dictEntry 有 overCode 則填入，否則一律清空）
            const overEl = document.getElementById(`Field_whr${dd}_${sfx}`);
            if (overEl && overEl.value !== overCode) {
                overEl.value = overCode;
                overEl.style.backgroundColor = overCode ? "#fff3cd" : "";
            }

            [
                { id: `Field_wareaa${dd}_${sfx}`,  val: amCode    },
                { id: `Field_wareab${dd}_${sfx}`,  val: pmCode    },
                { id: `Field_wareac${dd}_${sfx}`,  val: nightCode },
            ].forEach(({ id, val }) => {
                if (!val) return;
                const f = document.getElementById(id);
                if (f && f.value !== val) { f.value = val; f.style.backgroundColor = "#fff3cd"; }
            });
        });
    }
    return { success: true };
}