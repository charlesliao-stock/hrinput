// content.js
// 以 world: "ISOLATED"（預設）注入，可正常使用 chrome.* API。
// IE 相容 polyfill 已移至 content_main.js（world: "MAIN"）處理。
// 頁面驗證統一由 popup.js 的 sendMessage 負責，content.js 不重複處理。
// HR 預設班別統一由 background.js 初始化至 storage，content.js 不再 hardcode。

console.log("🚀 [KMUH Helper] 核心啟動 (ISOLATED World)");

// 擴充功能被重新載入/更新後，舊分頁仍殘留的 content script 實例，
// 其 chrome.runtime context 會失效(invalidated)，此時呼叫 chrome.runtime.sendMessage
// 會直接丟出 Uncaught Error，導致 Modal 的關閉/按鈕功能整個卡住。
// 這種情況下訊息本來就送不出去(背景已是新的 service worker)，改成安靜失敗即可，
// 不影響當下 DOM 上該做的收尾動作(例如 Modal 已經 remove 掉)。
function safeSendMessage(msg) {
    try {
        chrome.runtime.sendMessage(msg);
    } catch (err) {
        console.warn("[KMUH Helper] chrome.runtime context 已失效，訊息未送出，請重新整理頁面：", err);
    }
}

// ─────────────────────────────────────────────────────────────────
// 純函式 / 工具
// ─────────────────────────────────────────────────────────────────
function formatEmpId(id) {
    if (!id) return "";
    const s = String(id).trim();
    if (!/^\d+$/.test(s)) return "";
    return s.padStart(7, '0');
}

// 職編有效性檢核：純數字、若第一碼為 0 則去掉該碼，
// 剩餘長度需為 6 或 7 碼，否則視為無效職編。
function isValidEmpId(id) {
    const s = String(id || "").trim();
    if (!/^\d+$/.test(s)) return false;
    const stripped = s[0] === '0' ? s.slice(1) : s;
    return stripped.length === 6 || stripped.length === 7;
}

function getNextYM(yymm) {
    if (!yymm || yymm.length !== 6) return "";
    let y = parseInt(yymm.substring(0, 4)), m = parseInt(yymm.substring(4, 6)) + 1;
    if (m > 12) { m = 1; y++; }
    return String(y) + String(m).padStart(2, '0');
}

// ── 合併後的週期解析（原 parseCyclePeriods / parseFFPeriods） ─────
// 傳入不同括號即可區分四週變形（【】）與 FF 雙週（《》）。
// text 由呼叫端傳入共用，避免兩次呼叫各自讀一次 document.body.innerText
// （innerText 會觸發版面配置(layout)計算，是較昂貴的操作）。
function parsePeriods(bracketOpen, bracketClose, text) {
    const esc = (c) => c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(
        `${esc(bracketOpen)}(\\d+)${esc(bracketClose)}\\s*(\\d{1,2}\\/\\d{1,2})\\s*[~～]\\s*(\\d{1,2}\\/\\d{1,2})`,
        'g'
    );
    const periods = [];
    let m;
    while ((m = re.exec(text)) !== null) {
        periods.push({ label: m[1], start: m[2], end: m[3] });
    }
    return periods;
}
const parseCyclePeriods = (text) => parsePeriods('【', '】', text);
const parseFFPeriods    = (text) => parsePeriods('《', '》', text);

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

// ── 完整四週週期外的日曆週／雙週估算區塊 ───────────────────────────
// 週一至週日為一個日曆週；週期外只以連續1或2個日曆週為一個計算單位。
// 剩餘不足1週按1週估算，不足2週按2週估算；最後不完整週的跨月週六估WW、
// 週日估FF。預估只參與計數與提示，不改寫尚未匯入的日期。
function dateToGlobalIdx(date, oldYymm) {
    if (!date || !oldYymm) return null;
    const base = ymmBaseDate(oldYymm);
    return Math.round((date - base) / 86400000);
}

function mondayOfDate(date) {
    const monday = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const dayOfWeek = monday.getDay(); // 日=0，週一=1
    monday.setDate(monday.getDate() - ((dayOfWeek + 6) % 7));
    return monday;
}

// 取得一段週期中「完整落在員工可用資料範圍」的日曆週。
// 新進／次月調入人員在上月沒有可用班表，因此月初跨月的那一週不算完整週；
// 只有完整週才納入本次 WW／FF 目標。例如可用範圍只覆蓋四週週期後半的兩個完整週，
// 回傳兩週，WW 與 FF 的硬目標即各為2。一般既有人員不使用此縮減規則。
function getAvailableFullCalendarWeeks({ startIdx, endIdx, availableStartIdx, availableEndIdx,
    oldYymm, targetYymm, oldMonthDays }) {
    if (![startIdx, endIdx, availableStartIdx, availableEndIdx].every(Number.isFinite)) return [];
    if (startIdx > endIdx || availableStartIdx > availableEndIdx) return [];
    const firstDate = giToDate(startIdx, oldYymm, targetYymm, oldMonthDays);
    const lastDate = giToDate(endIdx, oldYymm, targetYymm, oldMonthDays);
    const weeks = [];
    let weekStart = mondayOfDate(firstDate);
    while (weekStart <= lastDate) {
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekEnd.getDate() + 6);
        const calendarStartIdx = dateToGlobalIdx(weekStart, oldYymm);
        const calendarEndIdx = dateToGlobalIdx(weekEnd, oldYymm);
        if (calendarStartIdx >= startIdx && calendarEndIdx <= endIdx &&
            calendarStartIdx >= availableStartIdx && calendarEndIdx <= availableEndIdx) {
            weeks.push({
                calendarStartIdx,
                calendarEndIdx,
                actualStartIdx: calendarStartIdx,
                actualEndIdx: calendarEndIdx,
                full: true,
                estimatedFFIdxs: [],
                estimatedWWIdxs: [],
            });
        }
        weekStart.setDate(weekStart.getDate() + 7);
    }
    return weeks;
}

// 新進／次月調入人員的週期檢核：月初跨月的不完整週不列入 WW／FF 硬目標，
// 只檢查完整落在目前可用資料範圍內的日曆週，每週各需1個指定類別。
function checkAvailableFullWeekTarget({ r, combined, matchFn, errType, label, typeLabel,
    oldYymm, targetYymm, oldMonthDays, validEnd, empId, err }) {
    const weeks = getAvailableFullCalendarWeeks({
        startIdx: r.startIdx,
        endIdx: r.endIdx,
        availableStartIdx: oldMonthDays,
        availableEndIdx: validEnd,
        oldYymm,
        targetYymm,
        oldMonthDays,
    });
    if (weeks.length === 0) return true;
    const count = weeks.reduce((total, week) => {
        const lo = Math.max(week.actualStartIdx, oldMonthDays);
        const hi = Math.min(week.actualEndIdx, validEnd);
        return total + (lo <= hi ? combined.slice(lo, hi + 1).filter(matchFn).length : 0);
    }, 0);
    const required = weeks.length;
    if (count !== required) {
        const first = weeks[0], last = weeks[weeks.length - 1];
        // 新進／調入人員缺少上月資料時，只有完整雙週 FF=2 或完整四週 WW=4
        // 才是硬限制；少於該週期的不足僅提示，不阻擋匯入／寫入。
        const hardTarget = typeLabel === 'FF' ? weeks.length >= 2 : weeks.length >= 4;
        const scopeText = hardTarget
            ? `完整${typeLabel === 'FF' ? '雙週' : '四週'}週期`
            : `僅${required}個完整可用日曆週，未形成完整${typeLabel === 'FF' ? '雙週' : '四週'}週期`;
        err.push({
            empId,
            startIdx: first.actualStartIdx,
            endIdx: last.actualEndIdx,
            type: errType,
            estimated: !hardTarget,
            blocking: hardTarget,
            suggestion: !hardTarget,
            msg: `${hardTarget ? '' : '💡 建議修改：'}${label} ${giToDateStr(first.actualStartIdx, oldYymm, targetYymm, oldMonthDays)}～${giToDateStr(last.actualEndIdx, oldYymm, targetYymm, oldMonthDays)} ${typeLabel}=${count}（應${required}；${scopeText}）`,
        });
    }
    return true;
}

function buildPostCycleCalendarBlocks(cycleRanges, oldMonthDays, validEnd, oldYymm, targetYymm) {
    if (!oldYymm || !targetYymm || !Array.isArray(cycleRanges)) return [];
    const completeCycles = cycleRanges.filter(r => r && r.endIdx <= validEnd);
    // 只有完整四週週期結束後的日期才屬於「四週週期以外」；
    // 若目前只有一個尚未結束的四週週期，該段由 cycleRanges 本身處理，不在此重複建立區塊。
    if (completeCycles.length === 0) return [];

    const postStartIdx = Math.max(...completeCycles.map(r => r.endIdx)) + 1;
    if (postStartIdx > validEnd) return [];

    const firstDate = giToDate(postStartIdx, oldYymm, targetYymm, oldMonthDays);
    const lastDate = giToDate(validEnd, oldYymm, targetYymm, oldMonthDays);
    let weekStart = mondayOfDate(firstDate);
    const weekInfos = [];

    while (weekStart <= lastDate) {
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekEnd.getDate() + 6);
        const weekStartIdx = dateToGlobalIdx(weekStart, oldYymm);
        const weekEndIdx = dateToGlobalIdx(weekEnd, oldYymm);
        const actualStartIdx = Math.max(postStartIdx, weekStartIdx);
        const actualEndIdx = Math.min(validEnd, weekEndIdx);
        if (actualStartIdx <= actualEndIdx) {
            const estimatedFFIdxs = [], estimatedWWIdxs = [];
            // 只把目前匯入範圍以後的週六／週日作為推估值；
            // 已有資料的日期一律以實際 Excel／網頁資料計算。
            for (let gi = Math.max(actualEndIdx + 1, validEnd + 1); gi <= weekEndIdx; gi++) {
                const d = giToDate(gi, oldYymm, targetYymm, oldMonthDays);
                if (d.getDay() === 6) estimatedWWIdxs.push(gi);
                if (d.getDay() === 0) estimatedFFIdxs.push(gi);
            }
            weekInfos.push({
                calendarStartIdx: weekStartIdx,
                calendarEndIdx: weekEndIdx,
                actualStartIdx,
                actualEndIdx,
                full: actualStartIdx === weekStartIdx && actualEndIdx === weekEndIdx,
                estimatedFFIdxs,
                estimatedWWIdxs,
            });
        }
        weekStart.setDate(weekStart.getDate() + 7);
    }

    const blocks = [];
    for (let i = 0; i < weekInfos.length; ) {
        const first = weekInfos[i];
        const weeks = [first];
        // 週期外以雙週為最大計算單位；不足兩個日曆週也以雙週估算，
        // 因此部分週可以和相鄰完整週合併，不能只因 full=false 就拆開。
        while (weeks.length < 2 && weekInfos[i + weeks.length]) {
            weeks.push(weekInfos[i + weeks.length]);
        }
        const actualStartIdx = Math.min(...weeks.map(w => w.actualStartIdx));
        const actualEndIdx = Math.max(...weeks.map(w => w.actualEndIdx));
        const calendarStartIdx = Math.min(...weeks.map(w => w.calendarStartIdx));
        const calendarEndIdx = Math.max(...weeks.map(w => w.calendarEndIdx));
        const estimatedFFIdxs = weeks.flatMap(w => w.estimatedFFIdxs);
        const estimatedWWIdxs = weeks.flatMap(w => w.estimatedWWIdxs);
        blocks.push({
            weeks,
            weekCount: weeks.length,
            requiredFF: weeks.length,
            requiredWW: weeks.length,
            calendarStartIdx,
            calendarEndIdx,
            actualStartIdx,
            actualEndIdx,
            estimatedFFIdxs,
            estimatedWWIdxs,
            estimated: weeks.some(w => !w.full || w.estimatedFFIdxs.length > 0 || w.estimatedWWIdxs.length > 0),
        });
        i += weeks.length;
    }
    return blocks;
}

// 建立尚未完成四週週期的估算區塊。完整週期由 cycleRanges 的原有硬規則處理；
// 未完成週期則按連續1／2個日曆週分組，實際值與跨月預估值合併計數。
function buildIncompleteCycleCalendarBlocks(cycleRanges, hasOldData, oldMonthDays, validEnd, oldYymm, targetYymm) {
    const blocks = [];
    (cycleRanges || []).filter(r => r && r.startIdx <= validEnd && r.endIdx > validEnd).forEach(range => {
        const availableStartIdx = hasOldData ? Math.max(0, range.startIdx) : Math.max(oldMonthDays, range.startIdx);
        const availableEndIdx = Math.min(validEnd, range.endIdx);
        if (availableStartIdx > availableEndIdx) return;

        const firstDate = giToDate(availableStartIdx, oldYymm, targetYymm, oldMonthDays);
        const lastDate = giToDate(availableEndIdx, oldYymm, targetYymm, oldMonthDays);
        let weekStart = mondayOfDate(firstDate);
        const weeks = [];
        while (weekStart <= lastDate) {
            const weekEnd = new Date(weekStart);
            weekEnd.setDate(weekEnd.getDate() + 6);
            const calendarStartIdx = dateToGlobalIdx(weekStart, oldYymm);
            const calendarEndIdx = dateToGlobalIdx(weekEnd, oldYymm);
            const actualStartIdx = Math.max(calendarStartIdx, availableStartIdx);
            const actualEndIdx = Math.min(calendarEndIdx, availableEndIdx);
            if (actualStartIdx <= actualEndIdx) {
                const estimatedFFIdxs = [], estimatedWWIdxs = [];
                const estimateFrom = Math.max(calendarStartIdx, range.startIdx);
                const estimateTo = Math.min(calendarEndIdx, range.endIdx);
                for (let gi = estimateFrom; gi <= estimateTo; gi++) {
                    if (gi >= actualStartIdx && gi <= actualEndIdx) continue;
                    const d = giToDate(gi, oldYymm, targetYymm, oldMonthDays);
                    if (d.getDay() === 6) estimatedWWIdxs.push(gi);
                    if (d.getDay() === 0) estimatedFFIdxs.push(gi);
                }
                weeks.push({
                    calendarStartIdx,
                    calendarEndIdx,
                    actualStartIdx,
                    actualEndIdx,
                    full: actualStartIdx === calendarStartIdx && actualEndIdx === calendarEndIdx,
                    estimatedFFIdxs,
                    estimatedWWIdxs,
                });
            }
            weekStart.setDate(weekStart.getDate() + 7);
        }

        for (let i = 0; i < weeks.length; i += 2) {
            const groupWeeks = weeks.slice(i, i + 2);
            blocks.push({
                weeks: groupWeeks,
                weekCount: groupWeeks.length,
                requiredFF: groupWeeks.length,
                requiredWW: groupWeeks.length,
                calendarStartIdx: Math.min(...groupWeeks.map(w => w.calendarStartIdx)),
                calendarEndIdx: Math.max(...groupWeeks.map(w => w.calendarEndIdx)),
                actualStartIdx: Math.min(...groupWeeks.map(w => w.actualStartIdx)),
                actualEndIdx: Math.max(...groupWeeks.map(w => w.actualEndIdx)),
                estimatedFFIdxs: groupWeeks.flatMap(w => w.estimatedFFIdxs || []),
                estimatedWWIdxs: groupWeeks.flatMap(w => w.estimatedWWIdxs || []),
                estimated: groupWeeks.some(w => !w.full || w.estimatedFFIdxs.length || w.estimatedWWIdxs.length),
                incompleteCycle: true,
                sourceRange: range,
            });
        }
    });
    return blocks;
}

// ─────────────────────────────────────────────────────────────────
// 訊息監聽入口
// ─────────────────────────────────────────────────────────────────
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {

    if (request.action === "readAndMemorize") {
        const data = captureWebSchedule();
        const now  = new Date();
        const sysYymm = String(now.getFullYear()) + String(now.getMonth() + 1).padStart(2, '0');
        // 月份不一致時，不在 content script 用 confirm()（會讓網頁分頁搶走焦點，
        // 導致 popup 視窗因失焦而被瀏覽器關閉，使用者按下確定也不會有任何反應）。
        // 改成回報 monthMismatch，交由 popup.js 在自己的視窗內跳出確認，
        // 使用者確認後再帶 forceProceed 重新呼叫一次。
        if (data.yymm && data.yymm !== sysYymm && !request.forceProceed) {
            return sendResponse({ success: false, monthMismatch: true, pageYymm: data.yymm, sysYymm });
        }
        const bodyText  = document.body.innerText;
        const periods   = parseCyclePeriods(bodyText);
        const ffPeriods = parseFFPeriods(bodyText);
        data.cyclePeriods = periods;
        data.ffPeriods    = ffPeriods;
        data.savedAt      = Date.now();
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

    if (request.action === "preflightWarnings") {
        handlePreflightWarnings(request).then(res => sendResponse(res));
        return true;
    }

    if (request.action === "injectOnly") {
        executeInjectionFlow(request.excelData).then(res => sendResponse(res));
        return true;
    }
});

// ─────────────────────────────────────────────────────────────────
// 步驟 2（前置檢查）：僅比對本月／下月人員名單差異，供 popup 整合成單一
// 匯入前確認視窗使用，不執行完整的班別規則檢測、也不開啟報告視窗。
// ─────────────────────────────────────────────────────────────────
async function handlePreflightWarnings(req) {
    const storage = await chrome.storage.local.get(['lastMonthData']);
    const { targetYymm } = deriveMonthContext(storage.lastMonthData);
    const excelMap = parseExcel(req.excelData, targetYymm);
    if (excelMap.error) return { success: false, message: excelMap.message };

    const { departedWarnings, noOldDataWarnings } = computeMembershipWarnings(storage.lastMonthData, excelMap);
    const nhRequired = parseInt(
        document.getElementById('ctl00_ContentPlaceHolder1_lbncount')?.textContent?.trim() || '0'
    , 10) || 0;
    return { success: true, departedWarnings, noOldDataWarnings, nhRequired };
}

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
            if (!hrShiftsList.map(x => String(x).toUpperCase()).includes(cStr) && !customDict.some(d => String(d.excel).trim().toUpperCase() === cStr)) {
                unknownCodes.add(cStr);
            }
        });
    }
    if (unknownCodes.size > 0) return { success: false, unknownCodes: Array.from(unknownCodes) };

    // 匯入 Excel 後立即檢查每個實際出現的自定班別是否有自身 W+／N+ 對應。
    // 缺漏時先暫存帶有 originExcel 的待補列並停止後續檢核；不等到一鍵配置才發現，
    // 也不允許以相同 HR 系統代號的其他自定班別代號代替。
    const overtimeGaps = collectOvertimeMappingGaps(excelMap, customDict);
    if (overtimeGaps.length > 0) {
        await chrome.storage.local.set({ pendingOvertimeGaps: overtimeGaps });
        return { success: false, overtimeGaps, missingDictRows: overtimeGaps };
    }

    const dataWithId  = Object.entries(excelMap).map(([id, v]) => ({ empId: id, noCheck: false, ...v }));
    // 週期文字通常同時列出目前涵蓋月份與後續週期，取最早一段作為錨點，
    // 才能讓 buildCheckRanges 往後展開出本次目標月份的全部區間。
    const lastCycle   = (lastData?.cyclePeriods || [])[0] || null;
    const lastFF      = (lastData?.ffPeriods    || [])[0] || null;
    const cycleRanges = buildCheckRanges(lastCycle, targetMonth, 28, oldYymm, oldMonthDays);
    const ffRanges    = buildCheckRanges(lastFF,    targetMonth, 14, oldYymm, oldMonthDays);
    const postCycleCalendarBlocks = buildPostCycleCalendarBlocks(
        cycleRanges, oldMonthDays, oldMonthDays + newMonthDays - 1, oldYymm, targetYymm
    );

    const allRanges = [...cycleRanges, ...ffRanges];
    const biStart   = allRanges.length > 0 ? Math.min(...allRanges.map(r => r.startIdx)) : oldMonthDays;
    const biEnd     = allRanges.length > 0 ? Math.max(...allRanges.map(r => r.endIdx))   : oldMonthDays + 27;

    const cycleLabel = cycleRanges.map((r, i) => `【${i + 1}】${r.start}～${r.end}`).join('、') || '未知';
    const ffLabel    = ffRanges.map((r, i)    => `《${i + 1}》${r.start}～${r.end}`).join('、') || '未知';
    const nhRequired = parseInt(
        document.getElementById('ctl00_ContentPlaceHolder1_lbncount')?.textContent?.trim() || '0'
    , 10) || 0;
    const nhLabel  = nhRequired > 0 ? `　／　NH/N+ 應排：${nhRequired} 天` : '';
    const postCycleLabel = postCycleCalendarBlocks.map(b => {
        const start = giToDateStr(b.calendarStartIdx, oldYymm, targetYymm, oldMonthDays);
        const end = giToDateStr(b.calendarEndIdx, oldYymm, targetYymm, oldMonthDays);
        return `週期外${b.weekCount === 4 ? '四週' : b.weekCount === 2 ? '雙週' : '日曆週'} ${start}～${end}`;
    }).join('、');
    const postLabel = postCycleLabel ? `　／　週期外配置：${postCycleLabel}` : '';
    const infoText = `四週變形：${cycleLabel}　／　FF雙週：${ffLabel}${postLabel}${nhLabel}`;

    // popup.js 匯入前確認視窗中，護理長勾選的國定假日日期（下個月「幾號」的整數陣列），
    // 轉換成本次資料集的 global day index（= oldMonthDays + 幾號 - 1），供「一鍵配置」使用。
    const nhDates = Array.isArray(req.nhDates)
        ? req.nhDates.map(d => oldMonthDays + d - 1).filter(gi => gi >= oldMonthDays)
        : [];

    const check = runDetailedCheck(lastData, excelMap, customDict, hrTimeMap, cycleRanges, ffRanges, oldMonthDays, newMonthDays, oldYymm, targetYymm, nhRequired, postCycleCalendarBlocks);
    if (req.showReport || check.errors.length > 0) {
        showModal("Excel 班表預覽與檢測報告", {
            headers: getHeaders(), data: dataWithId, errors: check.errors,
            monthDays: oldMonthDays, biStart, biEnd, cycleRanges, ffRanges, postCycleCalendarBlocks, nhRequired, nhDates,
            blankFillMode: req.blankFillMode || 'keep',
            blankFillCode: req.blankFillCode || '',
            departedWarnings: check.departedWarnings || [],
            isExcelReport: true, // 標記此為Excel匯入報告（區別於步驟1的本月預覽），「一鍵配置」等功能僅在此顯示
        }, infoText);
    }
    // 建議修改（blocking:false，例如跨月四週WW/W+、雙週FF計數不符）不列入阻擋匯入的判斷
    const blockingErrors = check.errors.filter(e => e.blocking !== false);
    return { success: blockingErrors.length === 0, noOldDataWarnings: check.noOldDataWarnings, departedWarnings: check.departedWarnings || [] };
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

// 兩個班別間的實際休息間距（分鐘）：daysBetween 為中間相隔的天數(0=緊鄰隔天)。
// 原本 runDetailedCheck / hasRestViolationAt 各自重複這條算式、且都多帶了會互相抵消的
// prevGi*1440 項，這裡合併成一份、順便化簡，兩處呼叫端負責的邊界情況完全相同。
function restGapMinutes(daysBetween, startMin, prevEndMin) {
    return (daysBetween + 1) * 1440 + startMin - prevEndMin;
}

// ── 月初日期 + 星期計算（用於新調入/跨月推算 WW週六、FF週日數） ──────
function ymmBaseDate(yymm) {
    return mmddToDate(`${yymm.substring(4, 6)}/01`, yymm);
}

function countWeekdayInRange(startIdx, endIdx, oldYymm, dow) {
    if (endIdx < startIdx || !oldYymm) return 0;
    const base = ymmBaseDate(oldYymm);
    let count = 0;
    for (let gi = startIdx; gi <= endIdx; gi++) {
        const d = new Date(base.getFullYear(), base.getMonth(), base.getDate() + gi);
        if (d.getDay() === dow) count++;
    }
    return count;
}

// ── 四週WW/雙週FF數量檢查（含新調入推算、跨月未來推算） ───────────
// dow: 6=週六(WW用)、0=週日(FF用)
function checkPeriodRange({ r, combined, matchFn, required, errType, label, typeLabel,
    hasOldData, oldMonthDays, validEnd, oldYymm, empId, err, dow, skipCrossMonthEstimate }) {

    // 範圍整段都在「新調入前」→ 員工尚未調入，整段跳過不檢查
    if (!hasOldData && r.endIdx < oldMonthDays) return;

    // 新調入人員：範圍跨過「下月1號」(調入日) → 前段(調入前)用週六/週日推算，後段用Excel實際數
    if (!hasOldData && r.startIdx < oldMonthDays && r.endIdx >= oldMonthDays) {
        const estOld    = countWeekdayInRange(r.startIdx, oldMonthDays - 1, oldYymm, dow);
        const actualNew = combined.slice(oldMonthDays, Math.min(r.endIdx, validEnd) + 1).filter(matchFn).length;
        let estFuture = 0;
        if (r.endIdx > validEnd) estFuture = countWeekdayInRange(validEnd + 1, r.endIdx, oldYymm, dow);
        const count = estOld + actualNew + estFuture;
        if (count !== required) {
            const detail = estFuture > 0
                ? `新調入人員：前段推算${estOld}天＋本月實際${actualNew}天＋後段推算${estFuture}天`
                : `新調入人員：前段推算${estOld}天＋後段實際${actualNew}天`;
            err.push({ empId, startIdx: r.startIdx, endIdx: r.endIdx, type: errType, estimated: true,
                blocking: false, suggestion: true,
                msg: `💡 建議修改：${label} ${r.start}～${r.end} ${typeLabel}=${count}（應${required}，${detail}，不強制鎖定）` });
        }
        return;
    }

    // 範圍整段都在目前資料範圍之後(尚無Excel資料可推算起點) → 跳過
    if (r.startIdx > validEnd) return;

    // 範圍跨過「下月最後一天」→ 後段(下下月)用週六/週日推算，前段用Excel實際數
    // 四週WW/W+：完全不檢查（skipCrossMonthEstimate=true 時直接略過）
    if (r.endIdx > validEnd) {
        if (skipCrossMonthEstimate) return;
        const knownStart  = hasOldData ? r.startIdx : Math.max(r.startIdx, oldMonthDays);
        const actualKnown = combined.slice(knownStart, validEnd + 1).filter(matchFn).length;
        const estFuture   = countWeekdayInRange(validEnd + 1, r.endIdx, oldYymm, dow);
        const count = actualKnown + estFuture;
        if (count !== required) {
            err.push({ empId, startIdx: r.startIdx, endIdx: r.endIdx, type: errType, estimated: true,
                blocking: false, suggestion: true,
                msg: `💡 建議修改：${label} ${r.start}～${r.end} ${typeLabel}=${count}（應${required}，跨下下月推算值：本月實際${actualKnown}天＋下月推算${estFuture}天，不強制鎖定）` });
        }
        return;
    }

    // 一般情況：整段範圍都落在已知資料中（皆為實際值，非推算）→ 嚴格檢核，仍列入阻擋匯入
    const startIdx = hasOldData ? r.startIdx : Math.max(r.startIdx, oldMonthDays);
    const count = combined.slice(startIdx, r.endIdx + 1).filter(matchFn).length;
    if (count !== required) {
        err.push({ empId, startIdx: r.startIdx, endIdx: r.endIdx, type: errType,
            msg: `${label} ${r.start}～${r.end} ${typeLabel}=${count}（應${required}）` });
    }
}

// 週期外的 FF／WW 檢核：完整雙週要求 2 FF + 2 WW，單一完整週要求 1 FF + 1 WW。
// 月底不完整週的跨月週六／週日以推估值計入，但錯誤只列為建議，不可強制改寫。
function checkPostCycleCalendarBlock({ block, combined, hasOldData, oldMonthDays, validEnd, oldYymm, targetYymm, empId, err }) {
    const weekStats = (block.weeks || []).map(week => {
        const countFrom = hasOldData ? week.actualStartIdx : Math.max(week.actualStartIdx, oldMonthDays);
        const actualTo = Math.min(week.actualEndIdx, validEnd);
        const actual = countFrom <= actualTo ? combined.slice(countFrom, actualTo + 1) : [];
        const actualFF = actual.filter(s => s === 'FF').length;
        const actualWW = actual.filter(s => s === 'WW' || s === 'W+').length;
        // B 規則：推估值只補缺口；已有實際同類別時，不重複計入推估值。
        const estimatedFF = actualFF === 0 ? (week.estimatedFFIdxs || []).length : 0;
        const estimatedWW = actualWW === 0 ? (week.estimatedWWIdxs || []).length : 0;
        return {
            week,
            countFrom,
            actualTo,
            ff: actualFF + estimatedFF,
            ww: actualWW + estimatedWW,
            estimated: !week.full || !!(week.estimatedFFIdxs?.length || week.estimatedWWIdxs?.length),
            estimatedFF,
            estimatedWW,
        };
    });
    const reportRange = stats => ({
        startIdx: Math.min(...stats.map(s => s.countFrom)),
        endIdx: Math.max(...stats.map(s => s.actualTo)),
    });
    const pushError = (stats, type, label, ffCount, ffTarget, wwCount, wwTarget, estimated, note = '', blockingOverride = null) => {
        if (!stats.length) return;
        const range = reportRange(stats);
        const first = stats[0].week;
        const last = stats[stats.length - 1].week;
        const startText = giToDateStr(first.calendarStartIdx, oldYymm, targetYymm, oldMonthDays);
        const endText = giToDateStr(last.calendarEndIdx, oldYymm, targetYymm, oldMonthDays);
        const estimateText = estimated ? '，含跨月推估值；不足僅提示、不強制鎖定' : '';
        const blocking = blockingOverride === null ? !estimated : blockingOverride;
        err.push({
            empId, startIdx: range.startIdx, endIdx: range.endIdx, type,
            estimated, blocking, suggestion: !blocking,
            msg: `${!blocking ? '💡 建議修改：' : ''}${label} ${startText}～${endText} FF=${ffCount}（應${ffTarget}）、WW=${wwCount}（應${wwTarget}）${note}${estimateText}`,
        });
    };

    // 週期外固定以1／2個連續日曆週為一個計算單位：雙週目標FF=2、WW/W+=2，
    // 單週目標各1。預估只用於補足計數；實際已有同類別時不重複估算。
    // 不足（含預估後仍不足）只提示；超額則必須修剪，無法安全修剪時阻擋並交人工處理。
    for (let i = 0; i < weekStats.length; i += 2) {
        const pair = weekStats.slice(i, i + 2);
        const ffCount = pair.reduce((n, s) => n + s.ff, 0);
        const wwCount = pair.reduce((n, s) => n + s.ww, 0);
        const estimated = pair.some(s => s.estimated);
        const target = pair.length;
        const isCompleteBiweek = pair.length === 2 && pair.every(s => s.week.full) && !estimated;
        const label = pair.length === 2 ? `週期外雙週第${Math.floor(i / 2) + 1}組` : '週期外單一日曆週';
        const note = pair.length === 2
            ? '；每週1個為優先，雙週目標FF與WW各2個'
            : '；單一日曆週目標FF與WW各1個';
        if (ffCount !== target) {
            // 只有完整雙週的FF=2不足／超額才是硬限制；單週或含預估的不足只提示，
            // 但任何實際超額都必須先修剪。
            pushError(pair, pair.length === 2 ? 'POST_CALENDAR_BIWEEK' : 'POST_CALENDAR_WEEK',
                label, ffCount, target, wwCount, target, estimated, note,
                ffCount > target || isCompleteBiweek);
        }
        if (wwCount !== target) {
            // WW不足永遠只提示；WW超額即使區段含預估，超額部分也來自實際資料，必須修剪。
            pushError(pair, pair.length === 2 ? 'POST_CALENDAR_BIWEEK_WW' : 'POST_CALENDAR_WEEK_WW',
                label, ffCount, target, wwCount, target, estimated, note,
                wwCount > target);
        }
    }
}

function giToDateStr(gi, oldYymm, targetYymm, oldMonthDays) {
    if (!oldYymm || !targetYymm) return `第${gi + 1}天`;
    const d = giToDate(gi, oldYymm, targetYymm, oldMonthDays);
    return `${d.getMonth() + 1}月${d.getDate()}日`;
}

// ── 共用：比對「本月網頁資料」與「下月Excel資料」的人員名單差異 ──────
// departedWarnings：本月網頁有、下月Excel無（可能離職/調離單位）
// noOldDataWarnings：本月網頁無、下月Excel有（可能新調入/找不到舊資料）
function computeMembershipWarnings(old, exc) {
    const departedWarnings = [], noOldDataWarnings = [];
    const excIdSet = new Set(Object.keys(exc || {}).map(formatEmpId));
    (old?.data || []).forEach(p => {
        const fid = formatEmpId(p.empId);
        if (fid && !excIdSet.has(fid)) departedWarnings.push({ empId: p.empId, name: p.name || '' });
    });
    for (let id in (exc || {})) {
        const hasOldData = !!old?.data?.find(p => formatEmpId(p.empId) === formatEmpId(id));
        if (!hasOldData) noOldDataWarnings.push({ empId: id, name: exc[id].name || '' });
    }
    return { departedWarnings, noOldDataWarnings };
}

function runDetailedCheck(old, exc, dict, hrTimeMap, cycleRanges, ffRanges, oldMonthDays, newMonthDays, oldYymm, targetYymm, nhRequired = 0, postCycleCalendarBlocks = [], manualIssues = []) {
    if (!old?.data && !exc) return { errors: [], noOldDataWarnings: [], departedWarnings: [] };
    const err = [];
    const toDate = (gi) => giToDateStr(gi, oldYymm, targetYymm, oldMonthDays);

    // 本月網頁有、下月Excel無 → 可能離職/調離單位；本月網頁無、下月Excel有 → 可能新調入
    const { departedWarnings, noOldDataWarnings } = computeMembershipWarnings(old, exc);

    for (let id in exc) {
        const oStf       = old?.data?.find(p => formatEmpId(p.empId) === formatEmpId(id));
        const hasOldData = !!oStf;
        const validStart = hasOldData ? 0 : oldMonthDays;
        const validEnd   = oldMonthDays + newMonthDays - 1;
        // 未完成的四週週期也按連續1／2個日曆週建立估算區塊；預估只計數，
        // 不會把不存在於目前匯入資料的日期寫回班表。
        const incompleteCycleBlocks = buildIncompleteCycleCalendarBlocks(
            cycleRanges, hasOldData, oldMonthDays, validEnd, oldYymm, targetYymm
        );
        const checkCalendarBlocks = [...postCycleCalendarBlocks, ...incompleteCycleBlocks];
        // 「不檢查」勾選：完全跳過四週WW/W+、雙週FF（含間隔）、NH/N+ 檢查
        const noCheck    = !!exc[id]?.noCheck;

        const oldShifts      = hasOldData ? oStf.shifts : Array(oldMonthDays).fill('');
        const rawExcelShifts = exc[id].shifts;
        const combined = [...oldShifts, ...rawExcelShifts].map(s => {
            const d = dict.find(x => String(x.excel).trim().toUpperCase() === String(s).trim().toUpperCase());
            return d ? d.sys : s;
        });

        // 勾選「不檢查」時，整條資料完全跳過所有規則檢核與建議提示。
        if (noCheck) continue;

        // W+ / N+ 建議更換提醒
        for (let i = 0; i < rawExcelShifts.length; i++) {
            const rawCode = String(rawExcelShifts[i] || "").trim();
            const gi = oldMonthDays + i;
            if (rawCode === 'W+' || rawCode === 'N+') {
                err.push({ empId: id, startIdx: gi, endIdx: gi, type: 'REPLACE_REQUIRED',
                    msg: `⚠️ 建議更換：${toDate(gi)} Excel 原始代號為 ${rawCode}，請更換為正確的加班代號（如加班小時別）。` });
            }
        }

        // FF 雙週檢查（含新調入前段推算 + 跨下下月後段推算）／勾選「不檢查」則完全跳過。
        // 週期外改用完整日曆週／雙週區塊檢核，避免與 ffRanges 重複計數。
        if (!noCheck) ffRanges.forEach((r, i) => {
            if (checkCalendarBlocks.some(b => r.endIdx >= b.actualStartIdx && r.startIdx <= b.calendarEndIdx)) return;
            if (!hasOldData) {
                if (r.endIdx > validEnd) return;
                checkAvailableFullWeekTarget({
                    r, combined,
                    matchFn: s => s === 'FF',
                    errType: `FF_${i + 1}`,
                    label: `FF雙週《${i + 1}》`,
                    typeLabel: 'FF',
                    oldYymm, targetYymm, oldMonthDays, validEnd,
                    empId: id, err,
                });
                return;
            }
            checkPeriodRange({
                r, combined,
                matchFn: s => s === 'FF',
                required: 2,
                errType: `FF_${i + 1}`,
                label: `FF雙週《${i + 1}》`,
                typeLabel: 'FF',
                hasOldData, oldMonthDays, validEnd, oldYymm,
                empId: id, err, dow: 0
            });
        });

        // FF 間隔檢查 (不可超過 12 天)。週期外最後不完整週的週日 FF
        // 以推估索引加入檢查，但不會將尚未匯入的日期當成可修改儲存格。
        if (!noCheck) {
            const ffCheckStart = hasOldData ? 0 : oldMonthDays;
            const ffIndices = [];
            for (let gi = ffCheckStart; gi <= validEnd; gi++) {
                if (combined[gi] === 'FF') ffIndices.push(gi);
            }
            checkCalendarBlocks.forEach(b => {
                b.weeks.forEach(week => {
                    const weekFrom = hasOldData ? week.actualStartIdx : Math.max(week.actualStartIdx, oldMonthDays);
                    const weekTo = Math.min(week.actualEndIdx, validEnd);
                    const actualWeekFF = weekFrom <= weekTo
                        ? combined.slice(weekFrom, weekTo + 1).filter(s => s === 'FF').length
                        : 0;
                    // B 規則：只有本週沒有實際 FF 時，才以推估的週日 FF 補缺口。
                    if (actualWeekFF === 0) {
                        (week.estimatedFFIdxs || []).forEach(gi => {
                            if (!ffIndices.includes(gi)) ffIndices.push(gi);
                        });
                    }
                });
            });
            ffIndices.sort((a, b) => a - b);
            for (let fi = 0; fi < ffIndices.length - 1; fi++) {
                const gap = ffIndices[fi + 1] - ffIndices[fi] - 1;
                if (gap > 12) err.push({ empId: id, startIdx: ffIndices[fi], endIdx: ffIndices[fi + 1], type: 'FF_GAP',
                    msg: `FF間隔過長：${toDate(ffIndices[fi])}(FF) 與 ${toDate(ffIndices[fi + 1])}(FF) 之間間隔 ${gap} 天（最多12天）` });
            }
        }

        // 四週變形檢查（含新調入前段推算；跨下下月後段完全不檢查）／勾選「不檢查」則完全跳過
        if (!noCheck) cycleRanges.forEach((r, i) => {
            if (!hasOldData) {
                if (r.endIdx > validEnd) return;
                checkAvailableFullWeekTarget({
                    r, combined,
                    matchFn: s => s === 'WW' || s === 'W+',
                    errType: `WW_${i + 1}`,
                    label: `四週變形【${i + 1}】`,
                    typeLabel: 'WW',
                    oldYymm, targetYymm, oldMonthDays, validEnd,
                    empId: id, err,
                });
                return;
            }
            checkPeriodRange({
                r, combined,
                matchFn: s => s === 'WW' || s === 'W+',
                required: 4,
                errType: `WW_${i + 1}`,
                label: `四週變形【${i + 1}】`,
                typeLabel: 'WW',
                hasOldData, oldMonthDays, validEnd, oldYymm,
                empId: id, err, dow: 6,
                skipCrossMonthEstimate: true // 跨月（下下月）四週WW/W+不用檢查
            });
        });

        // 週期外以1／2個連續日曆週為計算單位：雙週實際＋預估目標為2 FF、2 WW；
        // 單週目標各1。完整雙週的FF=2才硬限制；未滿雙週或不足的WW只列建議，
        // 跨月預估只計數、不寫回，且實際已有同類別時不重複估算。
        if (!noCheck) checkCalendarBlocks.forEach(block => {
            checkPostCycleCalendarBlock({
                block, combined, hasOldData, oldMonthDays, validEnd,
                oldYymm, targetYymm, empId: id, err,
            });
        });

        // NH / N+ 整月天數檢查（新調入人員照算，NH範圍本就在下個月內）／勾選「不檢查」則完全跳過
        if (!noCheck && nhRequired > 0) {
            const nhCount = combined.slice(oldMonthDays, oldMonthDays + newMonthDays)
                .filter(s => s === 'NH' || s === 'N+').length;
            if (nhCount !== nhRequired) err.push({ empId: id, startIdx: oldMonthDays, endIdx: validEnd, type: 'NH_COUNT',
                msg: `NH/N+ 天數不符：實際排 ${nhCount} 天（本月應排 ${nhRequired} 天）` });
        }

        // 接班間隔檢查 (應達 11 小時，從已知範圍起)
        const restStart = hasOldData ? 0 : Math.max(0, oldMonthDays - 1);
        let prevCode = null, prevEndMin = null, prevGi = -1;
        for (let gi = restStart; gi <= validEnd; gi++) {
            const code = combined[gi] || '';
            if (!code) continue;
            const timeInfo = getShiftTime(code, hrTimeMap);
            if (!timeInfo) { prevCode = null; prevEndMin = null; prevGi = -1; continue; }
            const { startMin, endMin } = timeInfo;
            if (prevEndMin !== null) {
                const daysBetween  = gi - prevGi - 1;
                const gap = restGapMinutes(daysBetween, startMin, prevEndMin);
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
    appendManualIssueErrors(err, manualIssues, oldYymm, targetYymm, oldMonthDays);
    return { errors: err, noOldDataWarnings, departedWarnings };
}

// 將一鍵配置找不到安全解的結果帶回同一份檢核報告；此類提示不取代原本的
// FF／WW 數量或接班間隔阻擋錯誤，僅用來明確告知使用者「已保留現況、請人工處理」。
function appendManualIssueErrors(err, manualIssues, oldYymm, targetYymm, oldMonthDays) {
    const seen = new Set();
    (manualIssues || []).forEach(issue => {
        const range = issue.range || issue.block || {};
        const startIdx = Number.isFinite(range.startIdx) ? range.startIdx
            : (Number.isFinite(range.actualStartIdx) ? range.actualStartIdx : range.calendarStartIdx);
        const endIdx = Number.isFinite(range.endIdx) ? range.endIdx
            : (Number.isFinite(range.actualEndIdx) ? range.actualEndIdx : range.calendarEndIdx);
        if (!Number.isFinite(startIdx) || !Number.isFinite(endIdx)) return;
        const key = `${formatEmpId(issue.empId)}|${issue.type || 'MANUAL_REQUIRED'}|${startIdx}|${endIdx}`;
        if (seen.has(key)) return;
        seen.add(key);
        const startText = giToDateStr(startIdx, oldYymm, targetYymm, oldMonthDays);
        const endText = giToDateStr(endIdx, oldYymm, targetYymm, oldMonthDays);
        err.push({
            empId: issue.empId,
            startIdx,
            endIdx,
            type: 'MANUAL_REQUIRED',
            estimated: !!(issue.block?.estimated),
            blocking: false,
            suggestion: true,
            manualRequired: true,
            msg: `⚠️ 需人工處理：${startText}～${endText} ${issue.message || '找不到完全安全的自動調整方案，已保留現況。'}`,
        });
    });
}

// ─────────────────────────────────────────────────────────────────
// 一鍵完成 WW/FF 配置：核心演算法（純函式，不動 DOM，方便測試/除錯）
// ─────────────────────────────────────────────────────────────────

// 整理「代表放假」的代碼集合：有填系統代號比系統代號，沒填比 Excel 原始代號
function buildLeaveCodeSet(shiftDict, hrShifts) {
    const set = new Set();
    // 1. 掃描自定義班別
    (shiftDict || []).forEach(d => {
        const sys = String(d.sys || '').trim().toUpperCase();
        const excel = String(d.excel || '').trim().toUpperCase();
        // FF／WW 是 HR 固定的放假符號，兼容舊版字典未保存 isLeave 的資料。
        if (!d.isLeave && !['FF', 'WW'].includes(sys) && !['FF', 'WW'].includes(excel)) return;
        if (sys) set.add(sys);
        if (excel) set.add(excel);
    });
    // 2. 掃描 HR 內建班別
    (hrShifts || []).forEach(h => {
        const code = String(typeof h === 'string' ? h : (h.code || '')).trim().toUpperCase();
        // background.js 舊版預設清單沒有 isLeave 欄位，FF／WW 仍應視為放假。
        const isBuiltInLeave = ['FF', 'WW'].includes(code);
        if (typeof h !== 'string' && !h.isLeave && !isBuiltInLeave) return;
        if (code) set.add(code);
    });
    return set;
}

// 將 Excel 原始代號依字典轉換成系統代號；若字典找不到、或找到但系統代號留空（放假符號常見），
// 直接回傳原始代號本身 —— 這樣不論有無填系統代號，都能透過 buildLeaveCodeSet 正確辨識放假格。
function convertCell(raw, dict) {
    const r = String(raw || '').trim().toUpperCase();
    if (!r) return '';
    const d = (dict || []).find(x => String(x.excel).trim().toUpperCase() === r);
    return (d && d.sys) ? String(d.sys).trim().toUpperCase() : r;
}

// 由「原始Excel代碼」取得「真正的原始班別」（用於退回時，避免誤退回成W+/N+自己）。
// 若該Excel代碼在字典中對應的 sys 本身就是 W+ 或 N+（代表原始資料進來時就已經是加班班別），
// 改取字典裡的「逾時(over)」欄位，那才是真正沒有加班標記的原始班別；
// 否則（sys 不是 W+/N+）就跟 convertCell 結果相同，直接沿用。
function getBaseShiftCode(raw, dict) {
    const r = String(raw || '').trim().toUpperCase();
    if (!r) return '';
    const d = (dict || []).find(x => String(x.excel).trim().toUpperCase() === r);
    if (!d) return r;
    const sys = String(d.sys || '').trim().toUpperCase();
    if ((sys === 'W+' || sys === 'N+') && d.over) return String(d.over).trim().toUpperCase();
    return sys || r;
}

// global day index → 實際日期物件（gi=0 為 oldYymm 1號；gi=oldMonthDays 為 targetYymm 1號）
function giToDate(gi, oldYymm, targetYymm, oldMonthDays) {
    let year, month, day;
    if (gi < oldMonthDays) {
        year = parseInt(oldYymm.substring(0, 4));    month = parseInt(oldYymm.substring(4, 6));    day = gi + 1;
    } else {
        year = parseInt(targetYymm.substring(0, 4));  month = parseInt(targetYymm.substring(4, 6));  day = gi - oldMonthDays + 1;
    }
    return new Date(year, month - 1, day);
}

// 該日與最近的星期日相差幾天（星期日本身為 0）
function sundayDistance(gi, oldYymm, targetYymm, oldMonthDays) {
    const dow = giToDate(gi, oldYymm, targetYymm, oldMonthDays).getDay();
    return Math.min(dow, 7 - dow);
}

// 規格書多處候選格挑選優先序統一為「週日優先，其次週六，再其次是離週日最近」
// （純用 sundayDistance 排序會讓週六與週一同為距離1而無法區分優先順序，故另立此函式）。
// 數值越小代表優先序越高：週日 = -1（最優先）、週六 = 0（次優先）、其餘依離週日最近遞增。
function sundayThenSaturdayRank(gi, oldYymm, targetYymm, oldMonthDays) {
    const dow = giToDate(gi, oldYymm, targetYymm, oldMonthDays).getDay();
    if (dow === 0) return -1;
    if (dow === 6) return 0;
    return sundayDistance(gi, oldYymm, targetYymm, oldMonthDays); // 恆 >=1，自然排在週六之後
}

// 往前尋找最近一個已經是 FF 的 global index（找不到回傳 null）
function lastFFBefore(combined, gi) {
    for (let k = gi - 1; k >= 0; k--) if (combined[k] === 'FF') return k;
    return null;
}

// 往後尋找最近一個已經是 FF 的 global index（找不到回傳 null）
function nextFFAfter(combined, gi) {
    for (let k = gi + 1; k < combined.length; k++) if (combined[k] === 'FF') return k;
    return null;
}

// ── 階段一：國假戰區處理 (NH/N+ 判定與絕對保護) ──
function lockNhDatesV3(combined, nhDates, lockedIdx, leaveCodeSet, ffRanges, cycleRanges, oldMonthDays, validEnd) {
    let debt = 0;
    const protectedIdx = new Set();

    nhDates.forEach(gi => {
        if (gi < oldMonthDays || gi > validEnd) return;
        // SS 一律不予異動（規格4.5）：這天若已鎖定（含已是SS），完全略過，不做NH/N+判定也不產生欠債
        if (lockedIdx.has(gi)) return;

        const isLeaveCode = leaveCodeSet.has(combined[gi]) || ['FF', 'WW'].includes(combined[gi]);

        // NH 優先處理：若此國假日不在任何「完整」的四週區間內
        // （不屬於已知 cycleRanges，或屬於但該區間跨月尾未滿四週，例如月底孤兒天數），
        // 略過後續 FF 資源匱乏判斷，直接依當天原始狀態鎖定：
        // 上班日 → N+；假別日 → NH。
        if (!cycleRanges.some(r => gi >= r.startIdx && gi <= r.endIdx && r.endIdx <= validEnd)) {
            combined[gi] = isLeaveCode ? 'NH' : 'N+';
            lockedIdx.add(gi);
            return;
        }

        if (!isLeaveCode) {
            // 上班日直接轉 N+
            combined[gi] = 'N+';
            lockedIdx.add(gi);
        } else {
            // 找出該國假落在哪個雙週區間
            const r = ffRanges.find(fr => gi >= fr.startIdx && gi <= fr.endIdx);
            if (r) {
                const fullLo = Math.max(r.startIdx, 0);
                const hi = Math.min(r.endIdx, validEnd);
                const leaveCandidates = [];
                for (let k = fullLo; k <= hi; k++) {
                    if (lockedIdx.has(k)) continue; // SS（或其他已鎖定格）不可被徵用為FF保護候選
                    if (leaveCodeSet.has(combined[k]) || ['FF', 'WW'].includes(combined[k])) {
                        leaveCandidates.push(k);
                    }
                }

                // 資源匱乏預判：這雙週假太少，強制保護
                if (leaveCandidates.length <= 2) {
                    leaveCandidates.forEach(k => {
                        if (k >= oldMonthDays && k <= hi) {
                            combined[k] = 'FF';
                            protectedIdx.add(k);
                        }
                    });
                    // 若這天被保護成 FF 了，NH 無法排入，產生欠債
                    if (combined[gi] === 'FF') debt++;
                } else {
                    combined[gi] = 'NH';
                    lockedIdx.add(gi);
                }
            } else {
                combined[gi] = 'NH';
                lockedIdx.add(gi);
            }
        }
    });
    return { debt, protectedIdx };
}

// ── 階段二：雙週 FF 精準補齊 (不足 2 個才補；規格未要求修剪超過 2 個的狀況) ──
// 規格書 5.階段二 明確規定：此階段候選格挑選「不考慮違規組／安全組」（是否安全與此階段無關，
// 交由後續階段即時判斷），優先順序統一為「週日優先，其次週六，再其次是離週日最近」。
function assignFFRangeV3(combined, lo, hi, lockedIdx, leaveCodeSet, oldYymm, targetYymm, oldMonthDays, fullLo, protectedIdx, preferredFFIdx = new Set()) {
    const countFrom = (typeof fullLo === 'number') ? fullLo : lo;

    // 找出目前區間內所有的 FF
    const currentFFs = [];
    for (let gi = countFrom; gi <= hi; gi++) {
        if (combined[gi] === 'FF') currentFFs.push(gi);
    }

    const priorityCompare = (a, b) =>
        sundayThenSaturdayRank(a, oldYymm, targetYymm, oldMonthDays) - sundayThenSaturdayRank(b, oldYymm, targetYymm, oldMonthDays);
    // 原始 Excel 已標示 FF 的候選優先保留，讓原本的 FF 不會無故被 WW 池洗掉。
    const ffPreferenceCompare = (a, b) =>
        (Number(!preferredFFIdx.has(a)) - Number(!preferredFFIdx.has(b))) || priorityCompare(a, b);

    // 把雙週切成前後兩個「半週」，讓 2 個 FF 盡量各落在不同週，避免同一週出現多個 FF、
    // 也避免兩個FF都往同一邊靠而讓另一邊跟鄰近雙週的FF距離過遠（>12天）
    const span = hi - countFrom + 1;
    const halfLen = Math.ceil(span / 2);
    const halves = [
        { lo: countFrom, hi: Math.min(countFrom + halfLen - 1, hi) },
        { lo: Math.min(countFrom + halfLen, hi + 1), hi: hi }
    ].filter(h => h.lo <= h.hi);

    // FF 超額：不再只依固定半週／原始FF順序修剪，改由全域候選組合搜尋。
    // 只要移除或挪動後會使包含上月FF在內的間隔超過12天，就不強制改動，交由人工處理。
    if (currentFFs.length > 2) {
        const mutableIndices = [];
        for (let gi = lo; gi <= hi; gi++) {
            if (lockedIdx.has(gi) || protectedIdx.has(gi)) continue;
            const v = combined[gi];
            if (v === 'FF' || v === 'WW' || leaveCodeSet.has(v)) mutableIndices.push(gi);
        }
        return normalizeFFSelectionForRange({
            combined,
            currentFFIndices: currentFFs,
            mutableIndices,
            targetFF: 2,
            weekInfos: halves.map(h => ({ actualStartIdx: h.lo, actualEndIdx: h.hi, estimatedFFIdxs: [] })),
            lockedIdx,
            protectedIdx,
            leaveCodeSet,
            hrTimeMap: null,
            oldYymm,
            targetYymm,
            oldMonthDays,
            validEnd: hi,
        });
    }

    // 精準補齊：FF < 2 才補，不足才動作
    let totalFF = currentFFs.length;
    const isFFCandidate = (gi) => {
        if (lockedIdx.has(gi) || protectedIdx.has(gi) || combined[gi] === 'FF') return false;
        return leaveCodeSet.has(combined[gi]) || combined[gi] === 'WW';
    };

    const tryFillHalf = (half, skipIfHalfAlreadyHasFF) => {
        if (totalFF >= 2) return;
        if (skipIfHalfAlreadyHasFF) {
            let hasFF = false;
            for (let gi = half.lo; gi <= half.hi; gi++) if (combined[gi] === 'FF') { hasFF = true; break; }
            if (hasFF) return; // 這半週已經有FF了，優先讓另一半週補，不重複塞同一半週
        }

        const candidates = [];
        for (let gi = Math.max(half.lo, lo); gi <= half.hi; gi++) {
            if (isFFCandidate(gi)) candidates.push(gi);
        }
        if (candidates.length === 0) return;

        // 依「週日優先，其次週六，再其次離週日最近」排序
        candidates.sort(ffPreferenceCompare);

        let chosen = null;
        for (const gi of candidates) {
            const prevFF = lastFFBefore(combined, gi);
            const nextFF = nextFFAfter(combined, gi);
            const prevOk = prevFF === null || (gi - prevFF - 1) <= 12;
            const nextOk = nextFF === null || (nextFF - gi - 1) <= 12;
            if (prevOk && nextOk) { chosen = gi; break; }
        }
        if (chosen === null) chosen = candidates[0]; // 找不到不破間隔的，只好硬湊

        combined[chosen] = 'FF';
        totalFF++;
    };

    // 第一輪：每半週最多補 1 個，優先讓前後兩半週都各有 1 個
    halves.forEach(half => tryFillHalf(half, true));
    // 第二輪：若某半週完全沒有候選導致仍不足 2 個，才允許同一半週補第 2 個
    halves.forEach(half => tryFillHalf(half, false));
}


// Fisher–Yates 洗牌：安全候選格數量 > 1 時，隨機決定修剪順序（沒有優先順序限制）
function shuffleArr(arr) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
}

// 模擬「若把 targetGi 這天改成 SS」是否會造成接班間隔不足11小時（跟 runDetailedCheck 的 REST_SHORT 判斷邏輯一致）。
// SS 有實際上下班時間（跟 FF/WW/NH/N+/W+ 是純flag、無時間不同），所以WW→SS可能會讓原本不會被檢查到的間隔冒出來。
function hasRestViolationAt(combined, hrTimeMap, targetGi) {
    let prevEndMin = null, prevGi = -1;
    for (let k = 0; k < combined.length; k++) {
        const code = combined[k] || '';
        if (!code) continue;
        const timeInfo = getShiftTime(code, hrTimeMap);
        if (!timeInfo) { prevEndMin = null; prevGi = -1; continue; }
        const { startMin, endMin } = timeInfo;
        if (prevEndMin !== null) {
            const daysBetween = k - prevGi - 1;
            const gap = restGapMinutes(daysBetween, startMin, prevEndMin);
            if (gap < 660 && (prevGi === targetGi || k === targetGi)) return true;
        }
        prevEndMin = endMin; prevGi = k;
    }
    return false;
}

// 若 WW 後面緊接 FF，直接把兩天交換前先確認 FF 位置不會造成 11 小時違規。
// 這可避免後續把原本的 WW 轉成 SS 時，形成例如 4N－SS 的短休息間隔。
function swapUnsafeWWWithFollowingFF(combined, hrTimeMap, lo, hi, lockedIdx, protectedIdx) {
    for (let wwGi = lo; wwGi < hi; wwGi++) {
        const ffGi = wwGi + 1;
        if (lockedIdx.has(wwGi) || lockedIdx.has(ffGi)) continue;
        if (protectedIdx?.has(wwGi) || protectedIdx?.has(ffGi)) continue;
        if (combined[wwGi] !== 'WW' || combined[ffGi] !== 'FF') continue;

        const originalClone = combined.slice();
        originalClone[wwGi] = 'SS';
        if (!hasRestViolationAt(originalClone, hrTimeMap, wwGi)) continue;

        const clone = combined.slice();
        clone[wwGi] = 'FF';
        clone[ffGi] = 'WW';
        if (hasRestViolationAt(clone, hrTimeMap, ffGi)) continue;

        combined[wwGi] = 'FF';
        combined[ffGi] = 'WW';
        return { wwGi, ffGi };
    }
    return null;
}

// ── 階段三：四週 WW/W+ 徵用 (漏斗模型) ──
// 只處理「完整」的四週區間 (呼叫端已過濾 r.endIdx > validEnd 的不完整區間)。
// 前提：WW 總池已於呼叫前，由 runAutoConfig 對「整月」一次性建立完成。
// 模式A：先統計總數，不足4時，逐週選一天週六(優先)/週日轉W+，週末用完仍不足則選平常上班日，達4立即停止。
// 模式B：先逐週強制選一天週六(優先)/週日轉W+ (不論該週是否已有假別)，再統計總數，仍不足才隨機選平常上班日補足，達4立即停止。
function assignWWRangeV3(combined, lo, hi, lockedIdx, leaveCodeSet, wwMode, oldYymm, targetYymm, oldMonthDays, fullLo, neverOvertimeIdx, targetWW = 4) {
    const needed = Math.max(0, Number.isFinite(targetWW) ? targetWW : 4);
    const countFrom = (typeof fullLo === 'number') ? fullLo : lo;

    const countTotal = () => {
        let c = 0;
        for (let gi = countFrom; gi <= hi; gi++) if (combined[gi] === 'WW' || combined[gi] === 'W+') c++;
        return c;
    };

    const isWorkday = (gi) => {
        if (lockedIdx.has(gi)) return false;
        if (neverOvertimeIdx && neverOvertimeIdx.has(gi)) return false; // 原始代號為「非上班」，永遠不可轉W+/N+
        const v = combined[gi];
        if (!v) return false;
        if (v === 'WW' || v === 'W+' || v === 'FF' || v === 'NH' || v === 'N+') return false;
        if (leaveCodeSet.has(v)) return false;
        return true;
    };

    // 依本次目標週數切分區塊 (僅用於挑選「每週一天」週末候選；週別分配為原則，不強制)
    const span = hi - countFrom + 1;
    const weekCount = Math.max(1, Math.min(4, needed || 1));
    const chunkLen = Math.ceil(span / weekCount);
    const weeks = [];
    for (let c = 0; c < weekCount; c++) {
        const cLo = Math.max(countFrom + c * chunkLen, lo);
        const cHi = Math.min(countFrom + (c + 1) * chunkLen - 1, hi);
        if (cLo <= cHi) weeks.push({ cLo, cHi });
    }

    const getWeekWorkdays = (week, targetDow = null) => {
        const arr = [];
        for (let gi = week.cLo; gi <= week.cHi; gi++) {
            if (!isWorkday(gi)) continue;
            if (targetDow === null || giToDate(gi, oldYymm, targetYymm, oldMonthDays).getDay() === targetDow) arr.push(gi);
        }
        return arr;
    };

    const fillRandomWorkdays = () => {
        let total = countTotal();
        if (total >= needed) return;
        const allWorkdays = [];
        for (let gi = lo; gi <= hi; gi++) if (isWorkday(gi)) allWorkdays.push(gi);
        const shuffled = shuffleArr(allWorkdays);
        for (let k = 0; k < shuffled.length && total < needed; k++) {
            combined[shuffled[k]] = 'W+';
            total++;
        }
    };

    if (wwMode === 'B') {
        // 模式B：不論該週是否已有假別，每週強制選一天週六(優先)或週日轉W+
        weeks.forEach(week => {
            const sat = getWeekWorkdays(week, 6);
            let chosen = sat.length > 0 ? sat[0] : null;
            if (chosen === null) {
                const sun = getWeekWorkdays(week, 0);
                if (sun.length > 0) chosen = sun[0];
            }
            if (chosen !== null) combined[chosen] = 'W+';
        });
        // 再統計總數，仍不足才隨機選平常上班日補足
        fillRandomWorkdays();
    } else {
        // 模式A：先統計總數，不足4時，逐週檢查該週是否已有WW/W+，沒有才選一天週六(優先)或週日轉W+，達4立即停止
        let total = countTotal();
        for (const week of weeks) {
            if (total >= needed) break;

            let alreadyHas = false;
            for (let gi = week.cLo; gi <= week.cHi; gi++) {
                if (combined[gi] === 'WW' || combined[gi] === 'W+') { alreadyHas = true; break; }
            }
            if (alreadyHas) continue; // 該週已有假別，不重複補，改找下一個還沒有的週

            const sat = getWeekWorkdays(week, 6);
            let chosen = sat.length > 0 ? sat[0] : null;
            if (chosen === null) {
                const sun = getWeekWorkdays(week, 0);
                if (sun.length > 0) chosen = sun[0];
            }
            if (chosen !== null) {
                combined[chosen] = 'W+';
                total++;
            }
        }
        // 週末用完仍不足4，選平常上班日 (隨機)
        fillRandomWorkdays();
    }
}
// ── 加班代號反查：嚴格依「原始 Excel 代號」找自身的 W+／N+ 對應。
// 同一個 HR 系統代號可能有白1、白2、白5等多個系列；因此不能只用
// 「系統=targetFlag 且 逾時=原始系統代號」找第一筆，否則白5會被錯配成白1+。
//
// 對應優先序：
// 1. 字典列若有 overtimeFor／originExcel／baseExcel 等明確來源欄位，必須與原始 Excel 完全相同。
// 2. 相容既有字典資料，以精確系列命名尋找：W+ 支援「原代號+」或「W原代號」；
//    N+ 支援「N原代號」或「原代號N+」。這些都是完整代號比對，不接受模糊前綴或第一筆替代。
// 3. 找不到自身對應即回傳 null，交由呼叫端保留原班別並建立該原始 Excel 的缺漏提示。
function findOvertimeSubCode(originCode, targetFlag, dict, originExcelCode = '') {
    const normalizedOriginCode = String(originCode || '').trim().toUpperCase();
    const normalizedOriginExcel = String(originExcelCode || '').trim().toUpperCase();
    const rows = (dict || []).filter(d =>
        String(d.sys || '').trim().toUpperCase() === String(targetFlag || '').trim().toUpperCase() &&
        String(d.over || '').trim().toUpperCase() === normalizedOriginCode
    );
    if (!normalizedOriginExcel || !normalizedOriginCode || rows.length === 0) return null;

    const explicit = rows.find(d => {
        const source = String(d.overtimeFor ?? d.originExcel ?? d.baseExcel ?? d.sourceExcel ?? '').trim().toUpperCase();
        return source && source === normalizedOriginExcel;
    });
    if (explicit) return String(explicit.excel || '').trim().toUpperCase() || null;

    const exactNames = targetFlag === 'W+'
        ? new Set([`${normalizedOriginExcel}+`, `W${normalizedOriginExcel}`])
        : new Set([`N${normalizedOriginExcel}`, `${normalizedOriginExcel}N+`]);
    const named = rows.filter(d => exactNames.has(String(d.excel || '').trim().toUpperCase()));
    return named.length === 1 ? String(named[0].excel || '').trim().toUpperCase() || null : null;
}

// 匯入 Excel 完成後立即檢查目前班表實際出現的自定班別，
// 每一個原始 Excel 代號都必須有「自身」的 W+ 與 N+ 對應。
// 只檢查有明確系統代號、且不是放假列或既有 W+/N+ 列的自定班別；
// 不以同 HR 系統代號的其他 Excel 代號替代。
function collectOvertimeMappingGaps(excelMap, customDict) {
    const sourceRows = new Map();
    (customDict || []).forEach(row => {
        const excel = String(row?.excel || '').trim().toUpperCase();
        const sys = String(row?.sys || '').trim().toUpperCase();
        if (!excel || !sys || sys === 'W+' || sys === 'N+' || row?.isLeave) return;
        sourceRows.set(excel, { excel, sys });
    });

    const importedCodes = new Set();
    Object.values(excelMap || {}).forEach(person => {
        (person?.shifts || []).forEach(raw => {
            const excel = String(raw || '').trim().toUpperCase();
            if (excel && sourceRows.has(excel)) importedCodes.add(excel);
        });
    });

    const gaps = [];
    importedCodes.forEach(originExcel => {
        const source = sourceRows.get(originExcel);
        ['W+', 'N+'].forEach(targetFlag => {
            if (findOvertimeSubCode(source.sys, targetFlag, customDict, originExcel)) return;
            gaps.push({ sys: targetFlag, over: source.sys, originExcel });
        });
    });
    return gaps;
}

// ── 週期外完整日曆週／雙週一鍵配置 ───────────────────────────────
function configurePostCycleCalendarBlock({ block, combined, baseParts, lockedIdx, protectedIdx,
    leaveCodeSet, hrTimeMap, oldMonthDays, validEnd, oldYymm, targetYymm, neverOvertimeIdx }) {
    const weekResults = [];
    const priorityCompare = (a, b) =>
        sundayThenSaturdayRank(a, oldYymm, targetYymm, oldMonthDays) - sundayThenSaturdayRank(b, oldYymm, targetYymm, oldMonthDays) || a - b;

    const isEditableLeave = (gi) => {
        if (gi < oldMonthDays || gi > validEnd || lockedIdx.has(gi) || protectedIdx?.has(gi)) return false;
        return leaveCodeSet.has(combined[gi]) || combined[gi] === 'WW';
    };
    const isEditableWorkday = (gi) => {
        if (gi < oldMonthDays || gi > validEnd || lockedIdx.has(gi) || protectedIdx?.has(gi)) return false;
        if (neverOvertimeIdx?.has(gi)) return false;
        const v = combined[gi];
        if (!v || leaveCodeSet.has(v) || ['WW', 'W+', 'FF', 'NH', 'N+'].includes(v)) return false;
        return true;
    };
    const currentFFIndices = () => {
        const out = [];
        for (let gi = 0; gi < combined.length; gi++) if (combined[gi] === 'FF') out.push(gi);
        return out;
    };
    const canPlaceFF = (gi, estimatedFFIdxs, removedFFIdxs = []) => {
        const removed = new Set(removedFFIdxs);
        const ffIndices = currentFFIndices().filter(x => x !== gi && !removed.has(x));
        (estimatedFFIdxs || []).forEach(x => { if (!ffIndices.includes(x)) ffIndices.push(x); });
        ffIndices.push(gi);
        ffIndices.sort((a, b) => a - b);
        const pos = ffIndices.indexOf(gi);
        const prev = pos > 0 ? ffIndices[pos - 1] : null;
        const next = pos + 1 < ffIndices.length ? ffIndices[pos + 1] : null;
        return (prev === null || gi - prev - 1 <= 12) && (next === null || next - gi - 1 <= 12);
    };

    // 嘗試將「不安全的 WW－FF」交換成「FF－WW」，但交換後的 FF 位置
    // 仍須與所有既有／推估 FF 維持最多 12 天間隔；並且交換後的 WW 必須能安全轉 SS。
    const trySmartSwapUnsafeWW = (week, weekLo, weekHi) => {
        for (let wwGi = weekLo; wwGi < weekHi; wwGi++) {
            const ffGi = wwGi + 1;
            if (lockedIdx.has(wwGi) || lockedIdx.has(ffGi)) continue;
            if (protectedIdx?.has(wwGi) || protectedIdx?.has(ffGi)) continue;
            if (combined[wwGi] !== 'WW' || combined[ffGi] !== 'FF') continue;

            const directClone = combined.slice();
            directClone[wwGi] = 'SS';
            const directUnsafe = hasRestViolationAt(directClone, hrTimeMap, wwGi);
            if (!directUnsafe) continue;

            // 移除原 FF，將 FF 放到前一天的 WW 位置，再檢查全域 FF 間隔。
            const swapClone = combined.slice();
            swapClone[wwGi] = 'FF';
            swapClone[ffGi] = 'WW';
            const ffSpacingOk = canPlaceFF(wwGi, week.estimatedFFIdxs || [], [ffGi]);
            if (!ffSpacingOk) continue;

            // 交換後原 FF 位置的 WW 若仍不能安全轉 SS，就不做半套交換。
            const finalClone = swapClone.slice();
            finalClone[ffGi] = 'SS';
            if (hasRestViolationAt(finalClone, hrTimeMap, ffGi)) continue;

            combined[wwGi] = 'FF';
            combined[ffGi] = 'SS';
            return { wwGi, ffGi };
        }
        return null;
    };

    block.weeks.forEach(week => {
        const lo = Math.max(week.actualStartIdx, oldMonthDays);
        const hi = Math.min(week.actualEndIdx, validEnd);
        if (lo > hi) return;

        const actualIdxs = [];
        for (let gi = Math.max(week.actualStartIdx, 0); gi <= Math.min(week.actualEndIdx, validEnd); gi++) actualIdxs.push(gi);
        let ffs = actualIdxs.filter(gi => combined[gi] === 'FF');
        const editableFFs = ffs.filter(gi => !lockedIdx.has(gi) && !protectedIdx?.has(gi));

        // 先將同一日曆週多出的 FF 退回 WW；鎖定／保護的格子不強制改動。
        if (ffs.length > 1) {
            const keep = ffs.slice().sort((a, b) => {
                const aFixed = Number(lockedIdx.has(a) || protectedIdx?.has(a));
                const bFixed = Number(lockedIdx.has(b) || protectedIdx?.has(b));
                return aFixed - bFixed || priorityCompare(a, b);
            })[0];
            ffs.slice().sort(priorityCompare).forEach(gi => {
                if (gi !== keep && !lockedIdx.has(gi) && !protectedIdx?.has(gi)) combined[gi] = 'WW';
            });
            ffs = actualIdxs.filter(gi => combined[gi] === 'FF');
        }

        const estimatedFF = ffs.length === 0 ? (week.estimatedFFIdxs || []) : [];
        if (ffs.length === 0 && estimatedFF.length === 0) {
            const candidates = actualIdxs.filter(isEditableLeave).sort(priorityCompare);
            const chosen = candidates.find(gi => canPlaceFF(gi, week.estimatedFFIdxs || []));
            if (chosen !== undefined) {
                combined[chosen] = 'FF';
                ffs = [chosen];
            }
        }

        // 若本週沒有實際 WW（或 WW 已由 FF 佔用），以可轉換工作日補一個 W+。
        let wwCount = actualIdxs.filter(gi => combined[gi] === 'WW' || combined[gi] === 'W+').length;
        const estimatedWW = wwCount === 0 ? (week.estimatedWWIdxs || []) : [];
        if (wwCount + estimatedWW.length < 1) {
            const workdays = actualIdxs.filter(isEditableWorkday).sort((a, b) => {
                const ad = giToDate(a, oldYymm, targetYymm, oldMonthDays).getDay();
                const bd = giToDate(b, oldYymm, targetYymm, oldMonthDays).getDay();
                const rank = d => d === 6 ? 0 : d === 0 ? 1 : 2;
                return rank(ad) - rank(bd) || priorityCompare(a, b);
            });
            if (workdays.length > 0) {
                combined[workdays[0]] = 'W+';
                wwCount++;
            }
        }

        // 只將可安全轉 SS 的超額 WW 轉換；若會造成 11 小時違規，先嘗試
        // 4N－WW－FF → 4N－FF－WW 的智能交換，再將交換後 WW 轉 SS。
        const expectedWW = 1;
        let totalWW = actualIdxs.filter(gi => combined[gi] === 'WW' || combined[gi] === 'W+').length + (estimatedWW.length ? 1 : 0);
        while (totalWW > expectedWW) {
            const swapped = trySmartSwapUnsafeWW(week, lo, hi);
            if (!swapped) break;
            totalWW = actualIdxs.filter(gi => combined[gi] === 'WW' || combined[gi] === 'W+').length + (estimatedWW.length ? 1 : 0);
        }

        const wwCandidates = actualIdxs.filter(gi => !lockedIdx.has(gi) && !protectedIdx?.has(gi) && combined[gi] === 'WW').sort(priorityCompare);
        for (const gi of wwCandidates) {
            if (totalWW <= expectedWW) break;
            const clone = combined.slice();
            clone[gi] = 'SS';
            if (!hasRestViolationAt(clone, hrTimeMap, gi)) {
                combined[gi] = 'SS';
                totalWW--;
            }
        }

        // 多出的 W+ 不屬於 WW 休假池，退回該日原始班別；不可安全處理的格子不強制變更。
        let excessW = actualIdxs.filter(gi => combined[gi] === 'W+').length + (estimatedWW.length ? 1 : 0) - expectedWW;
        if (excessW > 0) {
            for (const gi of actualIdxs) {
                if (excessW <= 0 || combined[gi] !== 'W+' || lockedIdx.has(gi) || protectedIdx?.has(gi)) continue;
                const idx = gi - oldMonthDays;
                const origin = baseParts[idx] || '';
                if (origin) {
                    combined[gi] = origin;
                    excessW--;
                }
            }
        }

        weekResults.push({
            week,
            actualFF: actualIdxs.filter(gi => combined[gi] === 'FF').length,
            actualWW: actualIdxs.filter(gi => combined[gi] === 'WW' || combined[gi] === 'W+').length,
            estimatedFF: (week.estimatedFFIdxs || []).length,
            estimatedWW: (week.estimatedWWIdxs || []).length,
        });
    });
    return weekResults;
}

// ── 總整合 ──
function runAutoConfig(modalState, wwMode) {
    const { dataset, storage, hrTimeMap, oldMonthDays, newMonthDays, cycleRanges, ffRanges, postCycleCalendarBlocks = [], oldYymm, targetYymm } = modalState;
    const dict = storage.shiftDict || [];
    const leaveCodeSet = buildLeaveCodeSet(dict, storage.hrShifts);
    const nhDates = dataset.nhDates || [];
    const validEnd = oldMonthDays + newMonthDays - 1;
    const unresolved = [];
    const manualIssues = [];

    dataset.data.forEach(p => {
        if (p.noCheck) return;

        const oStf = (storage.lastMonthData?.data || []).find(x => formatEmpId(x.empId) === formatEmpId(p.empId));
        const hasOldData = !!oStf;
        const oldPart = Array(oldMonthDays).fill('');
        if (oStf) {
            for (let i = 0; i < Math.min(oldMonthDays, oStf.shifts.length); i++) oldPart[i] = oStf.shifts[i];
        }
        const newPart = p.shifts.map(s => convertCell(s, dict));
        const baseParts = p.shifts.map(s => getBaseShiftCode(s, dict)); // 字典轉換後、套用W+/N+前的sys代號，供階段五還原使用
        const preferredFFIdx = new Set();
        p.shifts.forEach((raw, i) => {
            if (convertCell(raw, dict) === 'FF') preferredFFIdx.add(oldMonthDays + i);
        });
        const combined = [...oldPart, ...newPart];
        // 新進／次月調入人員只把完整落在本月可用資料內的日曆週納入目標；
        // 既有人員維持原本完整週期的固定目標。
        const availableFullWeeksFor = range => hasOldData ? [] : getAvailableFullCalendarWeeks({
            startIdx: range.startIdx,
            endIdx: range.endIdx,
            availableStartIdx: oldMonthDays,
            availableEndIdx: validEnd,
            oldYymm,
            targetYymm,
            oldMonthDays,
        });
        // 完整四週週期以外或尚未完成的四週週期，統一用1／2個日曆週估算。
        // 這些區塊的跨月日期只參與計數，實際可用日期才允許寫回。
        const incompleteCycleBlocks = buildIncompleteCycleCalendarBlocks(
            cycleRanges, hasOldData, oldMonthDays, validEnd, oldYymm, targetYymm
        );
        const configCalendarBlocks = [...postCycleCalendarBlocks, ...incompleteCycleBlocks];
        const lockedIdx = new Set();

        // 這天「原始」系統代號若屬於放假/非上班符號，永遠不可被轉為 W+/N+ 加班標記
        // (以原始代號為準，不受後續 combined[gi] 被覆寫成 WW/FF 等影響)
        const neverOvertimeIdx = new Set();
        newPart.forEach((code, i) => {
            if (leaveCodeSet.has(code)) neverOvertimeIdx.add(oldMonthDays + i);
        });

        // 規格 4.5：SS 一律不予異動。先鎖定所有原始已經是 SS 的格子，
        // 確保後續（含階段一國假處理）任何階段都不會再覆寫這些格子。
        for (let gi = oldMonthDays; gi <= validEnd; gi++) {
            if (combined[gi] === 'SS') lockedIdx.add(gi);
        }

        // 階段一：國假戰區處理
        let { debt: totalDebt, protectedIdx } = lockNhDatesV3(combined, nhDates, lockedIdx, leaveCodeSet, ffRanges, cycleRanges, oldMonthDays, validEnd);

        // 建立整月 WW 總池：保留匯入時原本的 FF，其餘放假符號才轉為 WW。
        // 保留原始 FF 可讓 4N－WW－FF 在 WW 過多時先交換為 4N－FF－WW，再安全轉 SS。
        for (let gi = oldMonthDays; gi <= validEnd; gi++) {
            if (!lockedIdx.has(gi) && leaveCodeSet.has(combined[gi])) {
                combined[gi] = preferredFFIdx.has(gi) ? 'FF' : 'WW';
            }
        }

        // 新進／次月調入人員月初不完整週不納入 WW／FF 硬目標，
        // 但其中的 O 等自定義放假符號仍須完成放假池轉換，不能恢復成原始符號。
        // 後續 assignFF／assignWW 只使用可用完整週的索引，因此這些部分週只會被轉成
        // WW（或保留既有 FF），不會被算入完整週的目標數，也不會被拿去補足硬目標。

        // 階段二：雙週 FF 精準補齊，只處理完整的 FF 雙週區間 (不完整的交由階段五(3)孤兒掃尾統一處理)
        ffRanges.forEach(r => {
            if (r.endIdx > validEnd) return; // 跳過跨月未滿的不完整雙週
            let fullLo = Math.max(r.startIdx, 0);
            let lo = Math.max(r.startIdx, oldMonthDays), hi = Math.min(r.endIdx, validEnd);
            let targetFF = 2;
            if (!hasOldData) {
                const availableWeeks = availableFullWeeksFor(r);
                // 月初跨月的部分週不列入目標；每個完整可用週配置1個FF，
                // 因此一個雙週只剩1個完整週時 targetFF=1，兩個完整週時 targetFF=2。
                targetFF = Math.min(2, availableWeeks.length);
                if (targetFF === 0) return;
                lo = Math.min(...availableWeeks.map(w => w.actualStartIdx));
                hi = Math.max(...availableWeeks.map(w => w.actualEndIdx));
                fullLo = lo;
            }
            if (lo > hi) return;
            const ffResult = assignFFRangeV3(combined, lo, hi, lockedIdx, leaveCodeSet, oldYymm, targetYymm, oldMonthDays, fullLo, protectedIdx, preferredFFIdx, targetFF);
            if (ffResult?.unresolved) {
                manualIssues.push({ empId: p.empId, name: p.name, type: 'FF_EXCESS_MANUAL', range: r,
                    message: 'FF 超額無法在不違反 FF 間隔的前提下安全調整，請人工處理。' });
            }
        });

        // 階段二 (3)：次月四週變形週期以外（月底孤兒天數）不預先挑FF，
        // 留給階段五之(3)統一做 FF/WW/SS 的最終收斂（該階段會自行重新計算孤兒範圍）。

        // 階段三：四週 WW/W+ 徵用 (漏斗模型)，只處理完整的四週區間
        cycleRanges.forEach(r => {
            if (r.endIdx > validEnd) return;
            let fullLo = Math.max(r.startIdx, 0);
            let lo = Math.max(r.startIdx, oldMonthDays), hi = Math.min(r.endIdx, validEnd);
            let targetWW = 4;
            if (!hasOldData) {
                const availableWeeks = availableFullWeeksFor(r);
                // 月初跨月的不完整週不計入；每個完整可用週只配置1個WW。
                targetWW = Math.min(4, availableWeeks.length);
                if (targetWW === 0) return;
                lo = Math.min(...availableWeeks.map(w => w.actualStartIdx));
                hi = Math.max(...availableWeeks.map(w => w.actualEndIdx));
                fullLo = lo;
            }
            if (lo > hi) return;

            assignWWRangeV3(combined, lo, hi, lockedIdx, leaveCodeSet, wwMode, oldYymm, targetYymm, oldMonthDays, fullLo, neverOvertimeIdx, targetWW);
        });

        // 階段四：全域 NH 債務清償 (整月範圍，含月底孤兒天數，均可用來還債)
        if (totalDebt > 0) {
            const wwCandidates = [];
            for (let gi = oldMonthDays; gi <= validEnd; gi++) {
                if (lockedIdx.has(gi)) continue;
                if (combined[gi] === 'WW') wwCandidates.push(gi);
            }
            // 分安全組 / 違規組：模擬轉 SS 是否違反 11 小時接班間隔
            const violatingGroup = [];
            wwCandidates.forEach(gi => {
                const clone = combined.slice();
                clone[gi] = 'SS';
                if (hasRestViolationAt(clone, hrTimeMap, gi)) violatingGroup.push(gi);
            });

            // 違規組優先轉 NH 還債
            const violatingOrder = shuffleArr(violatingGroup);
            for (let k = 0; k < violatingOrder.length && totalDebt > 0; k++) {
                combined[violatingOrder[k]] = 'NH';
                lockedIdx.add(violatingOrder[k]);
                totalDebt--;
            }

            // 違規組用盡仍有欠債，隨機挑選工作日轉 N+ 徹底還清
            if (totalDebt > 0) {
                const globalWorkIdxs = [];
                for (let gi = oldMonthDays; gi <= validEnd; gi++) {
                    if (lockedIdx.has(gi)) continue;
                    if (neverOvertimeIdx.has(gi)) continue; // 原始代號為「非上班」，不可轉N+
                    const v = combined[gi];
                    if (!v || v === 'WW' || v === 'W+' || v === 'FF' || v === 'NH' || v === 'N+') continue;
                    if (leaveCodeSet.has(v)) continue;
                    globalWorkIdxs.push(gi);
                }
                const workOrder = shuffleArr(globalWorkIdxs);
                for (let k = 0; k < workOrder.length && totalDebt > 0; k++) {
                    combined[workOrder[k]] = 'N+';
                    totalDebt--;
                }
            }
        }

        // 階段五 (1)(2)：四週內修剪 —— 任何轉換前一律先計算該四週目前 WW/W+ 總數：
        // 總數 <=4 完全不動；總數 >4 才安全組WW優先轉SS，安全組用盡仍超額則W+退回原始sys代號，直到總數=4。
        cycleRanges.forEach(r => {
            if (r.endIdx > validEnd) return; // 不完整四週不修剪
            let fullLo = Math.max(r.startIdx, 0);
            let lo = Math.max(r.startIdx, oldMonthDays), hi = Math.min(r.endIdx, validEnd);
            let targetWW = 4;
            let cycleWeeks = [];
            if (!hasOldData) {
                cycleWeeks = availableFullWeeksFor(r);
                targetWW = Math.min(4, cycleWeeks.length);
                if (targetWW === 0) return;
                lo = Math.min(...cycleWeeks.map(w => w.actualStartIdx));
                hi = Math.max(...cycleWeeks.map(w => w.actualEndIdx));
                fullLo = lo;
            } else {
                // 既有人員維持完整四週固定目標，但週別切分使用真正的7日曆日。
                for (let weekLo = fullLo; weekLo <= hi; weekLo += 7) {
                    cycleWeeks.push({
                        actualStartIdx: weekLo,
                        actualEndIdx: Math.min(weekLo + 6, hi),
                        estimatedFFIdxs: [],
                        estimatedWWIdxs: [],
                    });
                }
            }
            if (lo > hi) return;

            const cycleFFGroups = ffRanges
                .filter(fr => fr.startIdx >= r.startIdx && fr.endIdx <= r.endIdx)
                .flatMap(fr => {
                    const weeks = hasOldData ? [] : availableFullWeeksFor(fr);
                    if (hasOldData) return [{ lo: fr.startIdx, hi: fr.endIdx, target: 2, estimatedFFIdxs: [] }];
                    if (weeks.length === 0) return [];
                    return [{
                        lo: Math.min(...weeks.map(w => w.actualStartIdx)),
                        hi: Math.max(...weeks.map(w => w.actualEndIdx)),
                        target: Math.min(2, weeks.length),
                        estimatedFFIdxs: [],
                    }];
                });
            const result = normalizeWWExcessForRange({
                combined,
                rangeIndices: getRangeIndices(lo, hi, oldMonthDays, validEnd),
                weeks: cycleWeeks,
                targetWW,
                hrTimeMap,
                lockedIdx,
                protectedIdx,
                baseParts,
                oldMonthDays,
                oldYymm,
                targetYymm,
                ffGroups: cycleFFGroups,
                weekInfos: cycleWeeks,
            });
            if (result.remaining > 0) {
                manualIssues.push({
                    empId: p.empId,
                    name: p.name,
                    type: 'WW_EXCESS_MANUAL',
                    range: r,
                    message: '完整四週內仍有無法安全轉換的超額WW，請人工處理。',
                });
            }
        });

        // 階段五 (3)：完整四週週期後，依完整日曆週／雙週配置。
        // 連續兩個完整日曆週為一個雙週：每週 1 FF + 1 WW（合計 2 FF + 2 WW）；
        // 若只剩一個完整週，則以該週 1 FF + 1 WW 計算。月末不完整週由 helper
        // 納入跨月週六 WW／週日 FF 推估，但不強制改寫尚未匯入的日期。
        configCalendarBlocks.forEach(block => {
                configurePostCycleCalendarBlockV2({
                block, combined, baseParts, lockedIdx, protectedIdx,
                leaveCodeSet, hrTimeMap, oldMonthDays, validEnd,
                oldYymm, targetYymm, neverOvertimeIdx,
                manualIssues, empId: p.empId, empName: p.name,
            });
        });

        // 寫回與字典防呆
        for (let i = 0; i < newPart.length; i++) {
            const gi = oldMonthDays + i;
            const finalVal = combined[gi];
            if (finalVal === newPart[i]) continue;

            if (finalVal === 'N+' || finalVal === 'W+') {
                // 優先取轉換前的原始 HR 系統代號；若 Excel 原值本身是既有 N+/W+，
                // getBaseShiftCode 會回傳其「逾時」欄位，避免把 N+ 自己當成 over。
                const originExcelCode = String(p.shifts[i] || '').trim().toUpperCase();
                const originCode = getBaseShiftCode(p.shifts[i], dict) || newPart[i];
                const subCode = findOvertimeSubCode(originCode, finalVal, dict, originExcelCode);
                if (subCode) {
                    p.shifts[i] = subCode;
                } else {
                    unresolved.push({ empId: p.empId, name: p.name, gi, originCode, originExcel: originExcelCode, targetFlag: finalVal });
                }
            } else {
                p.shifts[i] = finalVal;
            }
        }
    });

    // 若 N+ 是因國假／推估配置落在原本空白的日期，該格沒有可直接讀取的 over。
    // 對有原始 Excel 代號的格子，前面的嚴格一對一查找已經完成；這裡只處理真正沒有原始代號的空白格。
    // 優先從同一員工的 W+ 缺漏組合推導；若全體只有唯一一個 W+ 原始代號，也可安全共用，
    // 讓同一組 85→W+、85→N+ 的兩列都自動填入 85，使用者仍只需填各自 Excel 代號。
    const knownWOverByEmp = new Map();
    const allKnownWOrigins = new Set();
    unresolved.forEach(u => {
        const over = String(u.originCode || '').trim().toUpperCase();
        if (u.targetFlag !== 'W+' || !over || over === 'W+' || over === 'N+') return;
        const empKey = formatEmpId(u.empId);
        if (!knownWOverByEmp.has(empKey)) knownWOverByEmp.set(empKey, new Set());
        knownWOverByEmp.get(empKey).add(over);
        allKnownWOrigins.add(over);
    });
    unresolved.forEach(u => {
        const currentOver = String(u.originCode || '').trim().toUpperCase();
        if (u.targetFlag !== 'N+' || currentOver) return;
        const empOrigins = knownWOverByEmp.get(formatEmpId(u.empId));
        if (empOrigins && empOrigins.size === 1) {
            u.originCode = [...empOrigins][0];
        } else if (allKnownWOrigins.size === 1) {
            u.originCode = [...allKnownWOrigins][0];
        }
    });

    const seenPairs = new Set();
    const missingDictRows = [];
    unresolved.forEach(u => {
        const key = `${u.targetFlag}|${u.originCode}|${u.originExcel || ''}`;
        if (seenPairs.has(key)) return;
        seenPairs.add(key);
        missingDictRows.push({ sys: u.targetFlag, over: u.originCode, originExcel: u.originExcel || '' });
    });

    return { unresolved, missingDictRows, manualIssues };
}

// ─────────────────────────────────────────────────────────────────
// UI：Modal 報告視窗
// ─────────────────────────────────────────────────────────────────
let modalState = {
    dataset: null, info: '', storage: null, hrTimeMap: {},
    oldYymm: '', targetYymm: '', oldMonthDays: 0, newMonthDays: 0,
    cycleRanges: [], ffRanges: [], postCycleCalendarBlocks: [], nhRequired: 0,
    autoConfigSnapshot: null, // 「一鍵配置」執行前的整表快照，供「清除配置結果」還原
    manualIssues: [], // 最近一次一鍵配置仍無安全解的區段；手動修正該員工後即清除
};

async function showModal(title, dataset, info) {
    const oldModal = document.getElementById('kmuh-modal'); if (oldModal) oldModal.remove();
    const oldStyle = document.getElementById('kmuh-modal-style'); if (oldStyle) oldStyle.remove();

    const storage   = await chrome.storage.local.get(['shiftDict', 'hrShifts', 'lastMonthData']);
    const hrTimeMap = buildHrTimeMap(storage.hrShifts);
    const { oldYymm, targetYymm, targetMonth, oldMonthDays, newMonthDays } =
        deriveMonthContext(storage.lastMonthData);

    // 與匯入前預檢一致：使用最早一段週期作為錨點，保留跨月的完整區間。
    const lastCycle   = (storage.lastMonthData?.cyclePeriods || [])[0] || null;
    const lastFF      = (storage.lastMonthData?.ffPeriods    || [])[0] || null;
    const cycleRanges = buildCheckRanges(lastCycle, targetMonth, 28, oldYymm, oldMonthDays);
    const ffRanges    = buildCheckRanges(lastFF,    targetMonth, 14, oldYymm, oldMonthDays);
    const postCycleCalendarBlocks = dataset.postCycleCalendarBlocks || buildPostCycleCalendarBlocks(
        cycleRanges, oldMonthDays, oldMonthDays + newMonthDays - 1, oldYymm, targetYymm
    );

    modalState = { dataset, info, storage, hrTimeMap, oldYymm, targetYymm, oldMonthDays, newMonthDays, cycleRanges, ffRanges, postCycleCalendarBlocks, nhRequired: dataset.nhRequired || 0, autoConfigSnapshot: null, manualIssues: [] };
    renderModalContent(title);
}

// ── 錯誤顏色對應表（移出迴圈，僅定義一次） ────────────────────────
const ERR_COLOR_MAP = {
    WW:               { border: '#e74c3c', bg: '#fff2f2' }, // 嚴格檢核未過（範圍完全落於已匯入資料內）
    FF:               { border: '#e74c3c', bg: '#fff2f2' }, // 嚴格檢核未過（範圍完全落於已匯入資料內）
    SUGGEST:          { border: '#3498db', bg: '#eaf4fb' }, // 建議修改（推算值，不強制鎖定）
    GAP:              { border: '#e67e22', bg: '#fff8f0' },
    REST:             { border: '#8e44ad', bg: '#fdf2ff' },
    REPLACE_REQUIRED: { border: '#f39c12', bg: '#fef5e7' },
    NH:               { border: '#0f6e56', bg: '#e1f5ee' },
    MANUAL:           { border: '#c0392b', bg: '#fff4e5' },
};

function getErrColor(type, estimated, blocking) {
    if (type === 'MANUAL_REQUIRED')   return ERR_COLOR_MAP.MANUAL;
    if (blocking === false)          return ERR_COLOR_MAP.SUGGEST;
    if (!type)                       return ERR_COLOR_MAP.WW;
    if (type === 'REPLACE_REQUIRED') return ERR_COLOR_MAP.REPLACE_REQUIRED;
    if (type === 'FF_GAP')           return ERR_COLOR_MAP.GAP;
    if (type === 'REST_SHORT')       return ERR_COLOR_MAP.REST;
    if (type === 'NH_COUNT')         return ERR_COLOR_MAP.NH;
    if (type.startsWith('FF_'))      return ERR_COLOR_MAP.FF;
    return ERR_COLOR_MAP.WW;
}

function renderModalContent(title) {
    const { dataset, info, oldMonthDays, cycleRanges, ffRanges, postCycleCalendarBlocks = [] } = modalState;
    const h        = dataset.headers;
    const mDays    = oldMonthDays;
    const total    = dataset.data.length;
    const blockingErrs  = dataset.errors?.filter(e => e.blocking !== false) || [];
    const suggestErrs   = dataset.errors?.filter(e => e.blocking === false) || [];
    const errorIds      = new Set(blockingErrs.map(e => formatEmpId(e.empId)));
    const suggestIds    = new Set(suggestErrs.map(e => formatEmpId(e.empId)));
    const errCount      = errorIds.size;
    const suggestCount  = suggestIds.size;

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
        ...postCycleCalendarBlocks.map((b, i) => {
            const start = giToDateStr(b.calendarStartIdx, modalState.oldYymm, modalState.targetYymm, modalState.oldMonthDays);
            const end = giToDateStr(b.calendarEndIdx, modalState.oldYymm, modalState.targetYymm, modalState.oldMonthDays);
            return `<span style="display:inline-flex;align-items:center;gap:3px;margin-right:8px;"><span style="display:inline-block;width:12px;height:12px;background:#e2e8f0;border:1px solid #94a3b8;border-radius:2px;"></span>週期外${b.weekCount === 2 ? '雙週' : '日曆週'} ${start}～${end}</span>`;
        }),
    ].join('');

    const errLegend = [
        { color: '#e74c3c', bg: '#fff2f2', label: '四週WW/W+、FF雙週／週期外日曆週數量錯誤（實際資料，須修正）' },
        { color: '#3498db', bg: '#eaf4fb', label: '💡 建議修改（不強制鎖定）：跨月最後不完整日曆週的推估值' },
        { color: '#e67e22', bg: '#fff8f0', label: 'FF間隔超過12天' },
        { color: '#8e44ad', bg: '#fdf2ff', label: '接班間距不足11小時' },
        { color: '#f39c12', bg: '#fef5e7', label: '建議更換 W+/N+' },
        { color: '#0f6e56', bg: '#e1f5ee', label: 'NH/N+ 天數不符' },
        { color: '#c0392b', bg: '#fff4e5', label: '需人工處理：無完全安全的自動調整方案（已保留現況）' },
    ].map(x =>
        `<span style="display:inline-flex;align-items:center;gap:3px;margin-right:10px;"><span style="display:inline-block;width:24px;height:14px;background:${x.bg};border:2px solid ${x.color};border-radius:2px;"></span>${x.label}</span>`
    ).join('');

    const oldStyle = document.getElementById('kmuh-modal-style'); if (oldStyle) oldStyle.remove();
    const style = document.createElement('style');
    style.id = 'kmuh-modal-style';
    style.innerHTML = `
        #kmuh-modal { position:fixed; top:2%; left:2%; width:96%; height:94%; background:#fdfdfe; z-index:10000; padding:25px; box-shadow:0 15px 60px rgba(0,0,0,0.4); overflow:auto; border-radius:15px; font-family:sans-serif; }
        .summary-row { display:flex; gap:15px; margin-bottom:15px; }
        .card { flex:1; padding:12px 15px; border-radius:10px; color:white; display:flex; align-items:baseline; justify-content:center; gap:8px; white-space:nowrap; }
        .card span { font-size:13px; }
        .card-val { font-size:1.8em; font-weight:bold; }
        .card-blue { background:#3498db; } .card-green { background:#2ecc71; } .card-red { background:#e74c3c; }
        .table-container { overflow-x:auto; border:1px solid #dfe6e9; border-radius:8px; }
        .report-table { width:100%; border-collapse:separate; border-spacing:0; background:white; }
        .report-table th, .report-table td { border:1px solid #ecf0f1; padding:8px; text-align:center; font-size:13px; min-width:32px; }
        .sticky-check{ position:sticky; left:0;     background:#f8f9fa !important; z-index:6; font-weight:bold; border-right:1px solid #dfe6e9 !important; min-width:44px; }
        .sticky-col  { position:sticky; left:45px;  background:#f8f9fa !important; z-index:5; font-weight:bold; border-right:2px solid #bdc3c7 !important; min-width:70px; }
        .sticky-name { position:sticky; left:116px; background:#f8f9fa !important; z-index:5; font-weight:bold; border-right:2px solid #bdc3c7 !important; min-width:60px; }
        .no-check-cb { width:16px; height:16px; cursor:pointer; }
        .cell-err { background:#fff2f2 !important; border:2px solid #ff7675 !important; }
        .tooltip { position:relative; cursor:help; }
        #kmuh-tip { position:fixed; background:#2d3436; color:white; padding:8px 14px; border-radius:6px; font-size:12px; z-index:99999; pointer-events:none; display:none; box-shadow:0 4px 12px rgba(0,0,0,0.4); max-width:360px; }
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
                // 阻擋性錯誤優先於「建議修改」顯示；同等級再比範圍大小
                const bigErr  = cellErrs.reduce((a, b) => {
                    const aBlocking = a.blocking !== false, bBlocking = b.blocking !== false;
                    if (aBlocking !== bBlocking) return aBlocking ? a : b;
                    return (b.endIdx - b.startIdx) > (a.endIdx - a.startIdx) ? b : a;
                });
                const { border, bg } = getErrColor(bigErr.type, bigErr.estimated, bigErr.blocking);
                const isFirst = gi === bigErr.startIdx, isLast = gi === bigErr.endIdx;
                borderStyle = `border-top:2px solid ${border} !important; border-bottom:2px solid ${border} !important;`
                    + (isFirst ? `border-left:2px solid ${border} !important;`  : 'border-left:none !important;')
                    + (isLast  ? `border-right:2px solid ${border} !important;` : 'border-right:none !important;');
                bgStyle  = `background:${bg} !important;`;
                tipText  = cellErrs.map(e => e.msg).join('\n');
            } else if (isBlank && isFill) {
                tipText = `將填入 ${dataset.blankFillCode}`;
            }
            const wkBg   = h.weekdays[i] === '日' || h.weekdays[i] === '六' ? '#fef9f9' : 'white';
            const cellBg = cellErrs.length > 0 ? '' : `background:${wkBg};`;
            const tipAttr = tipText ? `data-kmuh-tip="${tipText.replace(/"/g, '&quot;')}"` : '';
            const cls     = (tipText ? 'tooltip ' : '') + 'editable-cell';
            return `<td class="${cls}" ${tipAttr} contenteditable="true" data-p-idx="${pIdx}" data-s-idx="${i}" style="${cellBg}${bgStyle}${borderStyle}">${displayVal}</td>`;
        }).join('');
        const checkAttr = p.noCheck ? 'checked' : '';
        return `<tr><td class="sticky-check"><input type="checkbox" class="no-check-cb" data-p-idx="${pIdx}" ${checkAttr} title="勾選後完全不檢查此人的四週WW/W+、雙週FF及NH/N+"></td><td class="sticky-col">${p.empId || ''}</td><td class="sticky-name">${p.name || ''}</td>${cells}</tr>`;
    }).join('');

    let m = document.getElementById('kmuh-modal');
    if (!m) { m = document.createElement('div'); m.id = 'kmuh-modal'; document.body.appendChild(m); }

    m.innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <h2 style="margin:0;">📊 ${title}</h2>
            <div style="display:flex; gap:10px;">
                ${dataset.isExcelReport ? `
                <button id="autoConfigBtn" style="padding:10px 20px; background:#9b59b6; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold; font-size:14px;">⚡ 一鍵完成WW/FF配置</button>
                <button id="clearConfigBtn" ${modalState.autoConfigSnapshot ? '' : 'disabled'} style="padding:10px 20px; background:#95a5a6; color:white; border:none; border-radius:6px; font-weight:bold; font-size:14px; ${modalState.autoConfigSnapshot ? 'cursor:pointer;' : 'opacity:.55;cursor:not-allowed;'}">🧹 清除配置結果</button>
                ` : ''}
                <button id="saveM"  style="padding:10px 35px; background:#2ecc71; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold; font-size:14px;">💾 寫入班表</button>
                <button id="closeM" style="padding:10px 35px; background:#3498db; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold; font-size:14px;">✖ 關閉</button>
            </div>
        </div>
        ${info ? `<div style="margin-bottom:8px; padding:8px 12px; background:#eaf4fb; border-radius:6px; font-size:13px; color:#2c3e50;">ℹ️ ${info}</div>` : ''}
        ${dataset.departedWarnings?.length ? `<div style="margin-bottom:8px; padding:10px 14px; background:#fdf0e0; border-radius:6px; font-size:13px; color:#7d4500; border:1px solid #f5c88a;"><b>⚠️ 本月有、下月班表無（可能離職或調離單位）：</b><span style="margin-left:8px;">${dataset.departedWarnings.map(w => `${w.empId}${w.name ? '（' + w.name + '）' : ''}`).join('、')}</span></div>` : ''}
        <div style="margin-bottom:8px; padding:8px 12px; background:#fff3cd; border-radius:6px; font-size:13px; color:#856404; border:1px solid #ffeeba;">💡 提示：您可以直接點擊表格中的班別進行修改，系統會自動重新驗證。完整四週週期後，依完整日曆週計算：每週維持 1 個 WW／W+ 與 1 個 FF，連續兩週合計 2 個 WW／W+ 與 2 個 FF；WW 不足時以 W+ 補足，FF 間隔不得超過 12 天。最後不完整日曆週的跨月週六／週日採推估，且僅補足實際資料缺口；超額轉 SS 前會先檢查 11 小時接班間隔，違規時不強制轉換。勾選「不檢查」可完全跳過該員的四週WW/W+、雙週FF、NH/N+檢查。</div>
        ${legendItems ? `<div style="margin-bottom:6px; padding:6px 12px; background:#f8f9fa; border-radius:6px; font-size:12px; color:#555; display:flex; flex-wrap:wrap; gap:4px; align-items:center;"><b style="margin-right:6px;">檢查區間：</b>${legendItems}</div>` : ''}
        <div style="margin-bottom:12px; padding:6px 12px; background:#f8f9fa; border-radius:6px; font-size:12px; color:#555; display:flex; flex-wrap:wrap; gap:4px; align-items:center;"><b style="margin-right:6px;">錯誤類型：</b>${errLegend}</div>
        <div class="summary-row">
            <div class="card card-blue"><span>檢測總人數</span><div class="card-val">${total}</div></div>
            <div class="card card-green"><span>通過檢核</span><div class="card-val">${total - errCount}</div></div>
            <div class="card card-red"><span>違反規範</span><div class="card-val">${errCount}</div></div>
            <div class="card" style="background:#3498db;"><span>建議修改（不鎖定）</span><div class="card-val">${suggestCount}</div></div>
        </div>
        <div class="table-container">
            <table class="report-table">
                <thead>
                    <tr style="background:#f1f2f6;"><th rowspan="2" class="sticky-check" title="勾選後完全不檢查此人的四週WW/W+、雙週FF及NH/N+">不檢查</th><th rowspan="2" class="sticky-col">職編</th><th rowspan="2" class="sticky-name">姓名</th>${thW}</tr>
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
            .map((l, idx, arr) => `<div style="white-space:normal; line-height:1.6;${idx > 0 ? 'margin-top:6px; padding-top:6px; border-top:1px solid rgba(255,255,255,0.25);' : ''}">${l}</div>`).join('');
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

    m.querySelectorAll('.no-check-cb').forEach(cb => {
        cb.addEventListener('change', e => {
            const pIdx = parseInt(e.target.dataset.pIdx);
            modalState.dataset.data[pIdx].noCheck = e.target.checked;
            const empId = modalState.dataset.data[pIdx].empId;
            modalState.manualIssues = (modalState.manualIssues || []).filter(u => formatEmpId(u.empId) !== formatEmpId(empId));
            revalidateAndRefresh(title);
        });
    });

    const commitEditableCell = cell => {
        const pIdx   = parseInt(cell.dataset.pIdx);
        const sIdx   = parseInt(cell.dataset.sIdx);
        const newVal = cell.innerText.trim().toUpperCase();
        if (modalState.dataset.data[pIdx].shifts[sIdx] !== newVal) {
            modalState.dataset.data[pIdx].shifts[sIdx] = newVal;
            const empId = modalState.dataset.data[pIdx].empId;
            modalState.manualIssues = (modalState.manualIssues || []).filter(u => formatEmpId(u.empId) !== formatEmpId(empId));
            revalidateAndRefresh(title);
            return true;
        }
        return false;
    };

    const selectCellText = cell => {
        if (!cell || !cell.isConnected) return;
        const selection = window.getSelection();
        const range = document.createRange();
        range.selectNodeContents(cell);
        selection.removeAllRanges();
        selection.addRange(range);
    };

    m.querySelectorAll('.editable-cell').forEach(cell => {
        cell.addEventListener('blur', e => {
            commitEditableCell(e.target);
        });
        cell.addEventListener('keydown', e => {
            if (e.key === 'Enter') {
                e.preventDefault();
                e.target.blur();
                return;
            }
            if (e.key !== 'Tab') return;

            // 先記住目前格座標；若內容有變更，重新驗證會重繪 modal，
            // 因此必須在重繪後以座標重新取得下一格。
            e.preventDefault();
            const currentRow = parseInt(e.target.dataset.pIdx);
            const currentCol = parseInt(e.target.dataset.sIdx);
            const direction = e.shiftKey ? -1 : 1;
            const targetCol = currentCol + direction;

            commitEditableCell(e.target);

            const focusNextCell = () => {
                // commitEditableCell 可能同步重繪整個 modal；此時依 row／column 座標
                // 從目前連接中的 modal 重新取得目標格，避免舊 NodeList 索引失效。
                const liveModal = document.getElementById('kmuh-modal') || m;
                const nextCell = liveModal.querySelector(`.editable-cell[data-p-idx="${currentRow}"][data-s-idx="${targetCol}"]`);
                if (!nextCell) return;

                const focusAndSelect = () => {
                    if (!nextCell.isConnected) return;
                    nextCell.focus();
                    selectCellText(nextCell);
                };
                // 先立即處理，確保 Shift+Tab 不會停留在原格；再補兩次非同步處理，
                // 防止 Chromium 在 keydown 結束時重設 contenteditable 的 selection。
                focusAndSelect();
                requestAnimationFrame(focusAndSelect);
                setTimeout(focusAndSelect, 0);
            };
            focusNextCell();
        });
    });

    document.getElementById('closeM').onclick = () => {
        m.remove(); tip.remove();
        const style = document.getElementById('kmuh-modal-style'); if (style) style.remove();
        safeSendMessage({ action: "modalClosed" });
    };

const autoConfigBtn = document.getElementById('autoConfigBtn');
    if (autoConfigBtn) autoConfigBtn.onclick = async () => {
        // 修正：除了 wwMode，一併重抓 shiftDict 與 hrShifts，確保字典為最新狀態
        const set = await chrome.storage.local.get(['wwMode', 'shiftDict', 'hrShifts']);
        
        // 未設定或資料值不合法時，預設採用模式 A。
        // 模式 B 只有在使用者明確選取並儲存 B 時才啟用。
        const selectedWwMode = set.wwMode === 'B' ? 'B' : 'A';
        if (modalState.dataset.nhRequired > 0 && (!modalState.dataset.nhDates || modalState.dataset.nhDates.length === 0)) {
            alert('本月應排國定假日天數大於 0，但尚未取得您勾選的國定假日日期，請重新從步驟2匯入 Excel 一次。');
            return;
        }
        
        // 將最新字典寫回 modalState，供後續 runAutoConfig 使用
        modalState.storage.shiftDict = set.shiftDict;
        modalState.storage.hrShifts = set.hrShifts;
        modalState.hrTimeMap = buildHrTimeMap(set.hrShifts);

        if (!confirm('即將依「NH鎖定 → FF分配 → WW/W+分配 → 剩餘轉SS」的順序自動配置表格中所有員工（不檢查者除外），確定要執行嗎？')) return;

        const snapshot = JSON.parse(JSON.stringify(modalState.dataset.data));
        modalState.autoConfigSnapshot = snapshot;
        
        // 這裡的 modalState 已經包含最新的字典快照了
        const { unresolved, missingDictRows, manualIssues } = runAutoConfig(modalState, selectedWwMode);
        modalState.manualIssues = manualIssues || [];
        
        // 後續重驗證會把 manualIssues 合併回同一份報告，讓保留現況的區段仍清楚可見。

        let changedCount = 0;
        modalState.dataset.data.forEach((p, idx) => {
            const before = snapshot[idx];
            p.shifts.forEach((s, i) => { if (s !== before.shifts[i]) changedCount++; });
        });

        revalidateAndRefresh(title);
        let msg = changedCount > 0
            ? `已完成一鍵配置，共修改了 ${changedCount} 格，詳情請自行核對表格中已標色的變動。`
            : '已執行一鍵配置，但這次沒有任何格子需要調整。';
        if (manualIssues && manualIssues.length > 0) {
            const seen = new Set();
            const lines = manualIssues.filter(u => {
                const key = `${u.empId}|${u.type}|${u.range?.startIdx ?? u.block?.calendarStartIdx ?? ''}`;
                if (seen.has(key)) return false;
                seen.add(key); return true;
            }).map(u => `・${u.empId}${u.name ? '(' + u.name + ')' : ''}：${u.message}`);
            msg += `\n\n⚠️ 有 ${seen.size} 個超額 FF／WW 區段找不到完全安全的自動調整方案，已保留現況，請人工處理：\n${lines.join('\n')}`;
        }
        if (unresolved && unresolved.length > 0) {
            const { oldYymm, targetYymm, oldMonthDays } = modalState;
            const lines = unresolved.map(u => {
                const d = giToDate(u.gi, oldYymm, targetYymm, oldMonthDays);
                return `・${u.empId}${u.name ? '(' + u.name + ')' : ''} ${d.getMonth() + 1}/${d.getDate()}：原班別「${u.originCode}」找不到可轉為${u.targetFlag}的加班字典對應，已保留原班別`;
            });
            msg += `\n\n⚠️ 另有 ${unresolved.length} 格因字典中缺少對應設定而無法自動代換（這也是四週WW/W+或雙週FF數量仍不足的原因），即將為您開啟「班別字典管理」，並自動在下方新增缺漏組合的空列，補上 Excel 代號並儲存後，重新執行一鍵配置即可：\n${lines.join('\n')}`;
        }
        alert(msg);

        if (missingDictRows && missingDictRows.length > 0) {
            await chrome.storage.local.set({ pendingOvertimeGaps: missingDictRows });
            safeSendMessage({ action: 'openDictManager' });
        }
    };

    const clearConfigBtn = document.getElementById('clearConfigBtn');
    if (clearConfigBtn) clearConfigBtn.onclick = () => {
        if (!modalState.autoConfigSnapshot) return;
        if (!confirm('確定要回復到上次「一鍵配置」之前的狀態嗎？之後若有手動修改過這些格子，也會一併復原，此動作無法復原。')) return;
        modalState.dataset.data = modalState.autoConfigSnapshot;
        modalState.autoConfigSnapshot = null;
        modalState.manualIssues = [];
        revalidateAndRefresh(title);
    };

    document.getElementById('saveM').onclick = async () => {
        // 寫入前總防呆：若表格中仍有「違反規範」的阻擋性錯誤(紅框標示)尚未修正，一律擋下，
        // 不能讓使用者無視畫面上顯示的紅色錯誤直接寫入 HR 系統。
        // (「建議修改」type 為 blocking:false 的跨月推算值，不在此限，維持可寫入)
        const blockingErrs = (modalState.dataset.errors || []).filter(e => e.blocking !== false);
        if (blockingErrs.length > 0) {
            const blockingEmpIds = new Set(blockingErrs.map(e => formatEmpId(e.empId)));
            alert(`尚有 ${blockingEmpIds.size} 位人員的班表存在「違反規範」的錯誤(表格中紅框標示)，請先修正這些格子（或執行「⚡ 一鍵完成WW/FF配置」）後再寫入。`);
            return;
        }

// 新增：寫入前也重抓一次字典，確保防呆檢查不過期
        const { shiftDict } = await chrome.storage.local.get(['shiftDict']);
        modalState.storage.shiftDict = shiftDict;        
// 寫入前防呆：若表格中仍有「代表放假」的原始代號尚未被轉換（未跑一鍵配置或手動填完），先擋下，
        // 避免把 O/OFF 等原始 Excel 代號直接誤寫進 HR 正式系統。
const dict = modalState.storage.shiftDict || [];
        const leaveCodeSet = buildLeaveCodeSet(dict);
        const finalLeaveFlags = new Set(['FF', 'WW', 'NH']);
        let leftoverCount = 0;
        modalState.dataset.data.forEach(p => {
            if (p.noCheck) return;
            p.shifts.forEach(s => {
                const normalized = convertCell(s, dict);
                // FF／WW／NH 是一鍵配置後的合法系統結果，不應再被當作未轉換的原始放假符號。
                if (leaveCodeSet.has(normalized) && !finalLeaveFlags.has(normalized)) leftoverCount++;
            });
        });
        if (leftoverCount > 0) {
            alert(`尚有 ${leftoverCount} 格為「代表放假」的原始代號尚未轉換，請先執行「⚡ 一鍵完成WW/FF配置」，或手動修改這些格子後再寫入。`);
            return;
        }
        if (!confirm("確定要將目前修改後的班表寫入網頁嗎？")) return;
        const excelMap = {};
        modalState.dataset.data.forEach(p => { excelMap[p.empId] = { name: p.name, shifts: p.shifts }; });
        const res = await executeInjectionFlowFromMap(excelMap);
        if (res.success) { alert("班表寫入完成！"); document.getElementById('closeM').click(); }
        else alert("寫入失敗：" + (res.message || "未知錯誤"));
    };
}

function revalidateAndRefresh(title) {
    const { dataset, storage, hrTimeMap, cycleRanges, ffRanges, postCycleCalendarBlocks = [], oldMonthDays, newMonthDays, oldYymm, targetYymm, nhRequired } = modalState;
    const excelMap = {};
    dataset.data.forEach(p => { excelMap[p.empId] = { name: p.name, shifts: p.shifts, noCheck: !!p.noCheck }; });
    const check = runDetailedCheck(storage.lastMonthData, excelMap, storage.shiftDict || [], hrTimeMap, cycleRanges, ffRanges, oldMonthDays, newMonthDays, oldYymm, targetYymm, nhRequired || 0, postCycleCalendarBlocks, modalState.manualIssues || []);
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

// ── 日期行判定：需連續 8 欄皆為遞增連續數字（1,2,3...8）才視為日期行起點 ──
// （原本只檢查相鄰 2 欄為 1、2，容易誤判班表資料中恰好出現「1、2」的班別代碼。
//   與 popup.js 的 isDateRowStart 保持相同邏輯。）
const DATE_RUN_LENGTH = 8;
function isDateRowStart(row, ci) {
    for (let k = 0; k < DATE_RUN_LENGTH; k++) {
        const cd = parseCellDate(row[ci + k]);
        if (!cd || cd.day !== k + 1) return false;
    }
    return true;
}

function detectExcelLayout(data, targetYymm) {
    const targetMonth = parseInt(targetYymm.substring(4, 6));
    const targetYear  = parseInt(targetYymm.substring(0, 4));
    const monthDays   = new Date(targetYear, targetMonth, 0).getDate();
    let empIdColIdx = -1, nameColIdx = -1, day1ColIdx = -1;
    const EMP_KEYWORDS  = ["職編", "員工編號", "工號", "員編", "職員編號"];
    const NAME_KEYWORDS = ["姓名", "員工姓名", "名字"];
    // ── 水平掃描：只看前 20 欄（A~T），不限列數 ─────────────────────
    const SCAN_COL_LIMIT = 20;
    for (let ri = 0; ri < data.length; ri++) {
        const row = data[ri];
        if (!row) continue;
        const colLimit = Math.min(SCAN_COL_LIMIT, row.length);
        for (let ci = 0; ci < colLimit; ci++) {
            const val = String(row[ci] || "").trim();
            if (empIdColIdx === -1 && EMP_KEYWORDS.some(k => val.includes(k)))  empIdColIdx = ci;
            if (nameColIdx  === -1 && NAME_KEYWORDS.some(k => val.includes(k))) nameColIdx  = ci;
            if (day1ColIdx  === -1 && isDateRowStart(row, ci)) day1ColIdx = ci;
        }
        if (empIdColIdx !== -1 && nameColIdx !== -1 && day1ColIdx !== -1) break;
    }
    if (empIdColIdx === -1) {
        const colHits = {};
        const fallbackLimit = Math.min(SCAN_COL_LIMIT, day1ColIdx !== -1 ? day1ColIdx : SCAN_COL_LIMIT);
        for (let ri = 0; ri < data.length; ri++) {
            const row = data[ri]; if (!row) continue;
            for (let ci = 0; ci < fallbackLimit; ci++) {
                const val = String(row[ci] || "").trim();
                if (isValidEmpId(val)) colHits[ci] = (colHits[ci] || 0) + 1;
            }
        }
        let bestCol = -1, bestHits = 1;
        for (const [ci, hits] of Object.entries(colHits)) {
            if (hits > bestHits) { bestHits = hits; bestCol = parseInt(ci); }
        }
        if (bestCol !== -1) empIdColIdx = bestCol;
    }
    // 姓名欄關鍵字掃描不到時，改用內容特徵偵測：
    // 若某欄「多數」儲存格內容皆為 2 個(含)以上的純中文字，視為姓名欄
    // （掃描範圍限制在 1號日期欄之前的表頭資料區，避免誤判到班表班別欄）
    if (nameColIdx === -1) {
        const nameScanColLimit = day1ColIdx !== -1 ? day1ColIdx : SCAN_COL_LIMIT;
        const chineseNameRe = /^[\u4e00-\u9fa5]{2,}/;  // 不再要求整格都是中文，只要求開頭是姓名
        const colStats = {};
        for (let ri = 0; ri < data.length; ri++) {
            const row = data[ri]; if (!row) continue;
            for (let ci = 0; ci < Math.min(nameScanColLimit, row.length); ci++) {
                if (ci === empIdColIdx) continue;
                const val = String(row[ci] || "").trim();
                if (!val) continue;
                if (!colStats[ci]) colStats[ci] = { hit: 0, total: 0 };
                colStats[ci].total++;
                if (chineseNameRe.test(val)) colStats[ci].hit++;
            }
        }
        let bestCol = -1, bestHits = 0;
        for (const [ci, s] of Object.entries(colStats)) {
            if (s.hit >= 2 && s.hit / s.total >= 0.7 && s.hit > bestHits) {
                bestHits = s.hit;
                bestCol  = parseInt(ci);
            }
        }
        if (bestCol !== -1) nameColIdx = bestCol;
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
        if (!isValidEmpId(rawId)) return;
        const empId  = formatEmpId(rawId);
        const name   = String(r[layout.nameColIdx] || "").trim();
        const shifts = [];
        for (let i = 0; i < layout.monthDays; i++) {
            let val = r[layout.day1ColIdx + i];
            val = (val !== undefined && val !== null) ? String(val).replace(/[\r\n]/g, '').trim().toUpperCase() : "";
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

            const dictEntry = customDict.find(x => String(x.excel).trim().toUpperCase() === String(finalCode).trim().toUpperCase());
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


// ─────────────────────────────────────────────────────────────────
// 循環式 FF／WW 超額處理（定稿規則）
// 只處理超額；既有不足 FF／WW 的補足流程不由這組 helper 改寫。
// ─────────────────────────────────────────────────────────────────
function collectEffectiveFFIndices(combined, weekInfos = [], extraFFIdxs = []) {
    const out = [];
    for (let gi = 0; gi < combined.length; gi++) if (combined[gi] === 'FF') out.push(gi);
    (extraFFIdxs || []).forEach(gi => { if (!out.includes(gi)) out.push(gi); });
    // B 規則：某週已有實際FF時，不再把該週推估FF重複計入。
    (weekInfos || []).forEach(w => {
        const lo = Math.max(0, w.actualStartIdx ?? w.calendarStartIdx ?? 0);
        const hi = Math.min(combined.length - 1, w.actualEndIdx ?? w.calendarEndIdx ?? -1);
        const hasActual = lo <= hi && combined.slice(lo, hi + 1).some(v => v === 'FF');
        if (!hasActual) (w.estimatedFFIdxs || []).forEach(gi => { if (!out.includes(gi)) out.push(gi); });
    });
    return out.sort((a, b) => a - b);
}

function isFFSpacingValid(combined, weekInfos = [], extraFFIdxs = []) {
    const ff = collectEffectiveFFIndices(combined, weekInfos, extraFFIdxs);
    for (let i = 1; i < ff.length; i++) {
        if (ff[i] - ff[i - 1] - 1 > 12) return false;
    }
    return true;
}

function ffGroupCountValid(combined, groupRanges = [], weekInfos = []) {
    for (const group of groupRanges || []) {
        let count = 0;
        if (Array.isArray(group.weeks) && group.weeks.length > 0) {
            // B 規則逐週計算：某週已有實際 FF 就不再加入該週預估 FF，
            // 但其他沒有實際 FF 的週仍可使用自己的預估值。
            group.weeks.forEach(w => {
                const lo = Math.max(0, w.actualStartIdx ?? w.calendarStartIdx ?? 0);
                const hi = Math.min(combined.length - 1, w.actualEndIdx ?? w.calendarEndIdx ?? -1);
                const actual = lo <= hi ? combined.slice(lo, hi + 1).filter(v => v === 'FF').length : 0;
                count += actual || (w.estimatedFFIdxs || []).length;
            });
        } else {
            const lo = Math.max(0, group.lo);
            const hi = Math.min(combined.length - 1, group.hi);
            count = lo <= hi ? combined.slice(lo, hi + 1).filter(v => v === 'FF').length : 0;
            if (count === 0 && group.estimatedFFIdxs?.length) count += group.estimatedFFIdxs.length;
        }
        if (count !== group.target) return false;
    }
    return true;
}

// 在本月可移動放假池中搜尋「保留正好 targetFF 個FF」的最佳安全組合。
// currentFFIndices 可包含上月FF；mutableIndices 僅允許本月 FF／WW／其他放假符號。
function normalizeFFSelectionForRange({ combined, currentFFIndices, mutableIndices, targetFF,
    weekInfos = [], groupRanges = [], lockedIdx = new Set(), protectedIdx = new Set(), leaveCodeSet = new Set(),
    oldYymm = '', targetYymm = '', oldMonthDays = 0, validEnd = combined.length - 1 }) {
    const mutable = [...new Set((mutableIndices || []).filter(gi => {
        if (gi < oldMonthDays || gi > validEnd || lockedIdx.has(gi) || protectedIdx.has(gi)) return false;
        const v = combined[gi];
        return v === 'FF' || v === 'WW' || leaveCodeSet.has(v);
    }))].sort((a, b) => a - b);
    const current = [...new Set(currentFFIndices || [])].sort((a, b) => a - b);
    const mutableSet = new Set(mutable);
    const fixedFF = current.filter(gi => !mutableSet.has(gi));
    const need = targetFF - fixedFF.length;
    if (need < 0) return { applied: false, unresolved: true, reason: 'fixed-ff-over-target', remaining: current.length - targetFF };
    if (need === 0) {
        // 固定FF已足夠時，只能把本月多餘FF降為WW；但移除後仍須重驗證
        // 上月延伸FF間隔與雙週總量，否則保留現況交由人工處理。
        const trial = combined.slice();
        mutable.forEach(gi => {
            if (current.includes(gi)) trial[gi] = 'WW';
        });
        if (!isFFSpacingValid(trial, weekInfos) || !ffGroupCountValid(trial, groupRanges, weekInfos)) {
            return { applied: false, unresolved: true, reason: 'removing-ff-breaks-safety', remaining: current.length - targetFF };
        }
        mutable.forEach(gi => { combined[gi] = trial[gi]; });
        return { applied: true, unresolved: false, selected: fixedFF, removed: Math.max(0, current.length - targetFF) };
    }

    const candidates = mutable.filter(gi => ['FF', 'WW'].includes(combined[gi]) || combined[gi]);
    const currentSet = new Set(current);
    // 預估 FF 也算入目標；因此實際候選數量可介於「補足預估後的最少數量」
    // 與「完全不採用預估的最多數量」之間，不能固定只選 target-fixedFF 個。
    const estimatedCount = (weekInfos || []).reduce((total, w) => {
        const lo = Math.max(0, w.actualStartIdx ?? w.calendarStartIdx ?? 0);
        const hi = Math.min(combined.length - 1, w.actualEndIdx ?? w.calendarEndIdx ?? -1);
        const actual = lo <= hi ? combined.slice(lo, hi + 1).filter(v => v === 'FF').length : 0;
        return total + (actual === 0 ? (w.estimatedFFIdxs || []).length : 0);
    }, 0);
    const maxNeed = Math.max(0, need);
    const minNeed = Math.max(0, maxNeed - estimatedCount);
    if (candidates.length < minNeed) return { applied: false, unresolved: true, reason: 'not-enough-leave-pool', remaining: current.length - targetFF };

    const plans = [];
    const evaluate = selected => {
        const selectedSet = new Set(selected);
        const trial = combined.slice();
        mutable.forEach(gi => {
            if (selectedSet.has(gi)) trial[gi] = 'FF';
            else if (currentSet.has(gi)) trial[gi] = 'WW';
        });
        if (!isFFSpacingValid(trial, weekInfos)) return;
        if (!ffGroupCountValid(trial, groupRanges, weekInfos)) return;

        let weeklyPenalty = 0;
        (weekInfos || []).forEach(w => {
            const lo = Math.max(0, w.actualStartIdx ?? w.calendarStartIdx ?? 0);
            const hi = Math.min(trial.length - 1, w.actualEndIdx ?? w.calendarEndIdx ?? -1);
            const actualCount = lo <= hi ? trial.slice(lo, hi + 1).filter(v => v === 'FF').length : 0;
            const count = actualCount || (w.estimatedFFIdxs || []).length;
            // 每週1個是優先目標，但不凌駕雙週正好2個及安全條件。
            weeklyPenalty += Math.abs(count - 1);
        });
        const moved = mutable.reduce((n, gi) => n + (trial[gi] === combined[gi] ? 0 : 1), 0);
        const ff = collectEffectiveFFIndices(trial, weekInfos);
        const spacingSlack = ff.slice(1).reduce((n, gi, i) => n + (12 - (gi - ff[i] - 1)), 0);
        plans.push({ trial, selected: selected.slice(), weeklyPenalty, moved, spacingSlack });
    };

    const choose = (start, chosen, desired) => {
        if (chosen.length === desired) { evaluate(chosen); return; }
        const left = desired - chosen.length;
        for (let i = start; i <= candidates.length - left; i++) {
            chosen.push(candidates[i]);
            choose(i + 1, chosen, desired);
            chosen.pop();
        }
    };
    for (let desired = minNeed; desired <= maxNeed; desired++) choose(0, [], desired);
    if (plans.length === 0) return { applied: false, unresolved: true, reason: 'no-safe-ff-combination', remaining: current.length - targetFF };

    plans.sort((a, b) =>
        a.weeklyPenalty - b.weeklyPenalty ||
        b.spacingSlack - a.spacingSlack ||
        a.moved - b.moved ||
        a.selected.join(',').localeCompare(b.selected.join(','))
    );
    const best = plans[0];
    mutable.forEach(gi => { combined[gi] = best.trial[gi]; });
    return { applied: true, unresolved: false, selected: best.selected, removed: Math.max(0, current.length - targetFF), plan: best };
}

function getRangeIndices(lo, hi, oldMonthDays, validEnd) {
    const out = [];
    for (let gi = Math.max(lo, oldMonthDays); gi <= Math.min(hi, validEnd); gi++) out.push(gi);
    return out;
}

function effectiveWWCountForWeeks(combined, weeks) {
    let count = 0;
    (weeks || []).forEach(w => {
        const lo = Math.max(0, w.actualStartIdx);
        const hi = Math.min(combined.length - 1, w.actualEndIdx);
        const actual = lo <= hi ? combined.slice(lo, hi + 1).filter(v => v === 'WW' || v === 'W+').length : 0;
        count += actual || (w.estimatedWWIdxs || []).length;
    });
    return count;
}

function isFFLayoutValidForGroups(combined, ffGroups = [], weekInfos = []) {
    return isFFSpacingValid(combined, weekInfos) && ffGroupCountValid(combined, ffGroups, weekInfos);
}

// 不限相鄰日期搜尋 WW／FF 交換。每次只回傳一組完整安全交換，呼叫端套用後會重新計算。
function findNonAdjacentSmartSwap({ combined, rangeIndices, hrTimeMap, lockedIdx = new Set(), protectedIdx = new Set(),
    ffGroups = [], weekInfos = [] }) {
    const wwIndices = rangeIndices.filter(gi => combined[gi] === 'WW' && !lockedIdx.has(gi) && !protectedIdx.has(gi));
    const ffIndices = rangeIndices.filter(gi => combined[gi] === 'FF' && !lockedIdx.has(gi) && !protectedIdx.has(gi));
    for (const wwGi of wwIndices) {
        const direct = combined.slice();
        direct[wwGi] = 'SS';
        if (!hasRestViolationAt(direct, hrTimeMap, wwGi)) continue; // 安全組應先由外層處理

        for (const ffGi of ffIndices) {
            if (ffGi === wwGi) continue;
            const swapped = combined.slice();
            swapped[wwGi] = 'FF';
            swapped[ffGi] = 'WW';
            if (!isFFLayoutValidForGroups(swapped, ffGroups, weekInfos)) continue;

            const finalTrial = swapped.slice();
            finalTrial[ffGi] = 'SS';
            if (hasRestViolationAt(finalTrial, hrTimeMap, ffGi)) continue;
            return { wwGi, ffGi, trial: finalTrial };
        }
    }
    return null;
}

// 以「安全WW先轉SS，再循環找非相鄰交換，最後還原過多W+」收斂WW。
function normalizeWWExcessForRange({ combined, rangeIndices, weeks = [], targetWW, hrTimeMap,
    lockedIdx = new Set(), protectedIdx = new Set(), baseParts = [], oldMonthDays = 0,
    oldYymm = '', targetYymm = '', ffGroups = [], weekInfos = [] }) {
    const actualIndices = rangeIndices.filter(gi => gi >= 0 && gi < combined.length);
    const totalWW = () => effectiveWWCountForWeeks(combined, weeks);
    const weekIndexFor = gi => (weeks || []).findIndex(w => {
        const lo = Math.max(0, w.actualStartIdx);
        const hi = Math.min(combined.length - 1, w.actualEndIdx);
        return gi >= lo && gi <= hi;
    });
    const safeCandidate = () => {
        const counts = (weeks || []).map(w => {
            const lo = Math.max(0, w.actualStartIdx);
            const hi = Math.min(combined.length - 1, w.actualEndIdx);
            return lo <= hi ? combined.slice(lo, hi + 1).filter(v => v === 'WW' || v === 'W+').length : 0;
        });
        const candidates = actualIndices.filter(gi =>
            combined[gi] === 'WW' && !lockedIdx.has(gi) && !protectedIdx.has(gi)
        ).filter(gi => {
            const trial = combined.slice();
            trial[gi] = 'SS';
            return !hasRestViolationAt(trial, hrTimeMap, gi);
        });
        // 同樣安全時，先從 WW 較多的週轉換；並列時輪流偏向較早週，
        // 讓最後的安全分布盡量接近每週1個，而不是把WW全修剪在同一週。
        candidates.sort((a, b) => {
            const wa = weekIndexFor(a), wb = weekIndexFor(b);
            return (Number.isFinite(counts[wb]) ? counts[wb] : -1) - (Number.isFinite(counts[wa]) ? counts[wa] : -1) || a - b;
        });
        return candidates[0];
    };

    let iterations = 0;
    const seenStates = new Set();
    const maxIterations = Math.max(20, actualIndices.length * 3);
    while (totalWW() > targetWW && iterations++ < maxIterations) {
        const stateKey = actualIndices.map(gi => `${gi}:${combined[gi] || ''}`).join('|');
        if (seenStates.has(stateKey)) break;
        seenStates.add(stateKey);
        const safe = safeCandidate();
        if (safe !== undefined) {
            combined[safe] = 'SS';
            continue;
        }
        const swap = findNonAdjacentSmartSwap({ combined, rangeIndices: actualIndices, hrTimeMap, lockedIdx, protectedIdx, ffGroups, weekInfos });
        if (!swap) break;
        combined[swap.wwGi] = 'FF';
        combined[swap.ffGi] = 'SS';
    }

    // W+ 只在總量仍超額時退回原始班別；不足時的W+補足完全維持既有流程。
    while (totalWW() > targetWW) {
        const wPlus = actualIndices.find(gi => combined[gi] === 'W+' && !lockedIdx.has(gi) && !protectedIdx.has(gi));
        if (wPlus === undefined) break;
        const idx = wPlus - oldMonthDays;
        const origin = baseParts[idx] || '';
        if (!origin) break;
        combined[wPlus] = origin;
    }
    return { remaining: Math.max(0, totalWW() - targetWW), iterations };
}

// 週期外區塊的新流程：以1／2個日曆週為計算單位；FF在完整雙週才硬性2個，
// WW不足仍沿用每週轉W+，不足（含預估後仍不足）只提示，WW超額則循環安全收斂。
function configurePostCycleCalendarBlockV2({ block, combined, baseParts, lockedIdx, protectedIdx,
    leaveCodeSet, hrTimeMap, oldMonthDays, validEnd, oldYymm, targetYymm, neverOvertimeIdx,
    manualIssues = [], empId = '', empName = '' }) {
    const weeks = block.weeks || [];
    const priority = (a, b) => sundayThenSaturdayRank(a, oldYymm, targetYymm, oldMonthDays) - sundayThenSaturdayRank(b, oldYymm, targetYymm, oldMonthDays) || a - b;
    const actualRange = weeks.flatMap(w => getRangeIndices(w.actualStartIdx, w.actualEndIdx, oldMonthDays, validEnd));
    const mutableLeave = gi => {
        if (gi < oldMonthDays || gi > validEnd || lockedIdx.has(gi) || protectedIdx.has(gi)) return false;
        const v = combined[gi];
        return v === 'FF' || v === 'WW' || leaveCodeSet.has(v);
    };
    const editableWorkday = gi => {
        if (gi < oldMonthDays || gi > validEnd || lockedIdx.has(gi) || protectedIdx.has(gi) || neverOvertimeIdx?.has(gi)) return false;
        const v = combined[gi];
        return !!v && !leaveCodeSet.has(v) && !['FF', 'WW', 'W+', 'NH', 'N+'].includes(v);
    };

    const ffGroups = [];
    for (let i = 0; i < weeks.length; i += 2) {
        const groupWeeks = weeks.slice(i, i + 2);
        const lo = Math.min(...groupWeeks.map(w => w.actualStartIdx));
        const hi = Math.max(...groupWeeks.map(w => w.actualEndIdx));
        ffGroups.push({
            lo,
            hi,
            target: groupWeeks.length,
            weeks: groupWeeks,
            estimatedFFIdxs: groupWeeks.flatMap(w => w.estimatedFFIdxs || []),
        });
    }

    // 第一階段：只處理超額FF；單週1個為偏好，雙週／四週按每個雙週正好2個控制。
    for (let g = 0; g < ffGroups.length; g++) {
        const group = ffGroups[g];
        const groupWeeks = weeks.slice(g * 2, g * 2 + 2);
        const groupIndices = [];
        for (let gi = Math.max(0, group.lo); gi <= Math.min(combined.length - 1, group.hi); gi++) groupIndices.push(gi);
        const editableGroupIndices = actualRange.filter(gi => gi >= group.lo && gi <= group.hi);
        // 上月 FF 可參與雙週總量與分布計算，但只允許修改本月可用日期。
        const currentFF = groupIndices.filter(gi => combined[gi] === 'FF');
        const weeklyPenalty = groupWeeks.reduce((sum, w) => {
            const wLo = Math.max(0, w.actualStartIdx);
            const wHi = Math.min(combined.length - 1, w.actualEndIdx);
            const count = groupIndices.filter(gi => gi >= wLo && gi <= wHi && combined[gi] === 'FF').length;
            const estimatedCount = count === 0 ? (w.estimatedFFIdxs || []).length : 0;
            return sum + Math.abs((count || estimatedCount) - 1);
        }, 0);
        // 超額一定處理；雙週剛好2個但集中同一週時，也嘗試重新分布。
        if (currentFF.length < group.target || (currentFF.length === group.target && weeklyPenalty === 0)) continue;
        const mutable = groupIndices.filter(mutableLeave);
        const result = normalizeFFSelectionForRange({
            combined, currentFFIndices: currentFF, mutableIndices: mutable, targetFF: group.target,
            weekInfos: groupWeeks, groupRanges: [group], lockedIdx, protectedIdx, leaveCodeSet,
            oldYymm, targetYymm, oldMonthDays, validEnd,
        });
        if (result.unresolved) {
            manualIssues.push({ empId, name: empName, type: 'FF_EXCESS_MANUAL', block, message: '完整雙週內的超額FF無法在不違反FF間隔的前提下安全調整，請人工處理。' });
        }
    }

    // 保留既有FF不足處理：有放假池就補，沒有就不把一般上班日改成FF。
    weeks.forEach((week, weekIndex) => {
        const lo = Math.max(week.actualStartIdx, oldMonthDays);
        const hi = Math.min(week.actualEndIdx, validEnd);
        if (lo > hi) return;
        const actualIdxs = [];
        for (let gi = lo; gi <= hi; gi++) actualIdxs.push(gi);
        const actualFF = actualIdxs.filter(gi => combined[gi] === 'FF').length;
        const ffGroup = ffGroups[Math.floor(weekIndex / 2)];
        const groupAllIndices = [];
        if (ffGroup) {
            for (let x = Math.max(0, ffGroup.lo); x <= Math.min(combined.length - 1, ffGroup.hi); x++) groupAllIndices.push(x);
        }
        const groupFFCount = ffGroup
            ? groupAllIndices.filter(x => combined[x] === 'FF').length
            : actualFF;
        // 每週1個是偏好；若雙週／單週總量已達目標，即使本月週為0也不得再補成超額。
        if (actualFF === 0 && groupFFCount < (ffGroup?.target ?? 1) && !(week.estimatedFFIdxs || []).length) {
            const candidates = actualIdxs.filter(mutableLeave).sort(priority);
            const chosen = candidates.find(gi => {
                const trial = combined.slice();
                trial[gi] = 'FF';
                const trialGroupCount = ffGroup
                    ? groupAllIndices.filter(x => trial[x] === 'FF').length
                    : trial.filter(x => x === 'FF').length;
                return trialGroupCount <= (ffGroup?.target ?? 1) && isFFSpacingValid(trial, weeks);
            });
            if (chosen !== undefined) combined[chosen] = 'FF';
        }

        const actualWW = actualIdxs.filter(gi => combined[gi] === 'WW' || combined[gi] === 'W+').length;
        const estimatedWW = actualWW === 0 ? (week.estimatedWWIdxs || []).length : 0;
        if (actualWW + estimatedWW < 1) {
            const workdays = actualIdxs.filter(editableWorkday).sort((a, b) => {
                const ad = giToDate(a, oldYymm, targetYymm, oldMonthDays).getDay();
                const bd = giToDate(b, oldYymm, targetYymm, oldMonthDays).getDay();
                const rank = d => d === 6 ? 0 : d === 0 ? 1 : 2;
                return rank(ad) - rank(bd) || priority(a, b);
            });
            if (workdays.length > 0) combined[workdays[0]] = 'W+';
        }
    });

    // 第二階段：WW超額以整個週期外區塊處理；完整四週可重新分布，總量控制在區塊目標。
    const wwTarget = weeks.length;
    const wwResult = normalizeWWExcessForRange({
        combined, rangeIndices: actualRange, weeks, targetWW: wwTarget, hrTimeMap,
        lockedIdx, protectedIdx, baseParts, oldMonthDays, oldYymm, targetYymm,
        ffGroups, weekInfos: weeks,
    });
    if (wwResult.remaining > 0) {
        manualIssues.push({ empId, name: empName, type: 'WW_EXCESS_MANUAL', block, message: '完整週期外區塊仍有無法安全轉換的超額WW，請人工處理。' });
    }

    return weeks.map(week => {
        const lo = Math.max(week.actualStartIdx, oldMonthDays);
        const hi = Math.min(week.actualEndIdx, validEnd);
        const vals = lo <= hi ? combined.slice(lo, hi + 1) : [];
        return {
            week,
            actualFF: vals.filter(v => v === 'FF').length,
            actualWW: vals.filter(v => v === 'WW' || v === 'W+').length,
            estimatedFF: (week.estimatedFFIdxs || []).length,
            estimatedWW: (week.estimatedWWIdxs || []).length,
        };
    });
}
