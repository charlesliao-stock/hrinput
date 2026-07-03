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
        // 月份不一致時，不在 content script 用 confirm()（會讓網頁分頁搶走焦點，
        // 導致 popup 視窗因失焦而被瀏覽器關閉，使用者按下確定也不會有任何反應）。
        // 改成回報 monthMismatch，交由 popup.js 在自己的視窗內跳出確認，
        // 使用者確認後再帶 forceProceed 重新呼叫一次。
        if (data.yymm && data.yymm !== sysYymm && !request.forceProceed) {
            return sendResponse({ success: false, monthMismatch: true, pageYymm: data.yymm, sysYymm });
        }
        const periods   = parseCyclePeriods();
        const ffPeriods = parseFFPeriods();
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
    return { success: true, departedWarnings, noOldDataWarnings };
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

    const dataWithId  = Object.entries(excelMap).map(([id, v]) => ({ empId: id, noCheck: false, ...v }));
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
            departedWarnings: check.departedWarnings || [],
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

function runDetailedCheck(old, exc, dict, hrTimeMap, cycleRanges, ffRanges, oldMonthDays, newMonthDays, oldYymm, targetYymm, nhRequired = 0) {
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
        // 「不檢查」勾選：完全跳過四週WW/W+、雙週FF（含間隔）、NH/N+ 檢查
        const noCheck    = !!exc[id]?.noCheck;

        const oldShifts      = hasOldData ? oStf.shifts : Array(oldMonthDays).fill('');
        const rawExcelShifts = exc[id].shifts;
        const combined = [...oldShifts, ...rawExcelShifts].map(s => {
            const d = dict.find(x => String(x.excel).trim().toUpperCase() === String(s).trim().toUpperCase());
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

        // FF 雙週檢查（含新調入前段推算 + 跨下下月後段推算）／勾選「不檢查」則完全跳過
        if (!noCheck) ffRanges.forEach((r, i) => {
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

        // FF 間隔檢查 (不可超過 12 天，只在已知範圍內)／勾選「不檢查」則完全跳過
        if (!noCheck) {
            const ffCheckStart = hasOldData ? 0 : oldMonthDays;
            const ffIndices = [];
            for (let gi = ffCheckStart; gi <= validEnd; gi++) {
                if (combined[gi] === 'FF') ffIndices.push(gi);
            }
            for (let fi = 0; fi < ffIndices.length - 1; fi++) {
                const gap = ffIndices[fi + 1] - ffIndices[fi] - 1;
                if (gap > 12) err.push({ empId: id, startIdx: ffIndices[fi], endIdx: ffIndices[fi + 1], type: 'FF_GAP',
                    msg: `FF間隔過長：${toDate(ffIndices[fi])}(FF) 與 ${toDate(ffIndices[fi + 1])}(FF) 之間間隔 ${gap} 天（最多12天）` });
            }
        }

        // 四週變形檢查（含新調入前段推算；跨下下月後段完全不檢查）／勾選「不檢查」則完全跳過
        if (!noCheck) cycleRanges.forEach((r, i) => {
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
    return { errors: err, noOldDataWarnings, departedWarnings };
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
    WW:               { border: '#e74c3c', bg: '#fff2f2' }, // 嚴格檢核未過（範圍完全落於已匯入資料內）
    FF:               { border: '#e74c3c', bg: '#fff2f2' }, // 嚴格檢核未過（範圍完全落於已匯入資料內）
    SUGGEST:          { border: '#3498db', bg: '#eaf4fb' }, // 建議修改（推算值，不強制鎖定）
    GAP:              { border: '#e67e22', bg: '#fff8f0' },
    REST:             { border: '#8e44ad', bg: '#fdf2ff' },
    REPLACE_REQUIRED: { border: '#f39c12', bg: '#fef5e7' },
    NH:               { border: '#0f6e56', bg: '#e1f5ee' },
};

function getErrColor(type, estimated, blocking) {
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
    const { dataset, info, oldMonthDays, cycleRanges, ffRanges } = modalState;
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
    ].join('');

    const errLegend = [
        { color: '#e74c3c', bg: '#fff2f2', label: '四週WW/W+、FF雙週數量錯誤（落於已匯入資料內，須修正）' },
        { color: '#3498db', bg: '#eaf4fb', label: '💡 建議修改（不強制鎖定）：跨月推算值的雙週FF數量' },
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
                <button id="saveM"  style="padding:10px 35px; background:#2ecc71; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold; font-size:14px;">💾 寫入班表</button>
                <button id="closeM" style="padding:10px 35px; background:#3498db; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold; font-size:14px;">✖ 關閉</button>
            </div>
        </div>
        ${info ? `<div style="margin-bottom:8px; padding:8px 12px; background:#eaf4fb; border-radius:6px; font-size:13px; color:#2c3e50;">ℹ️ ${info}</div>` : ''}
        ${dataset.departedWarnings?.length ? `<div style="margin-bottom:8px; padding:10px 14px; background:#fdf0e0; border-radius:6px; font-size:13px; color:#7d4500; border:1px solid #f5c88a;"><b>⚠️ 本月有、下月班表無（可能離職或調離單位）：</b><span style="margin-left:8px;">${dataset.departedWarnings.map(w => `${w.empId}${w.name ? '（' + w.name + '）' : ''}`).join('、')}</span></div>` : ''}
        <div style="margin-bottom:8px; padding:8px 12px; background:#fff3cd; border-radius:6px; font-size:13px; color:#856404; border:1px solid #ffeeba;">💡 提示：您可以直接點擊表格中的班別進行修改，系統會自動重新驗證。四週WW/W+、雙週FF數量檢查中，若區間完全落在本次已匯入的資料內（實際值），不符規範仍會鎖定匯入；若區間跨到下個月尚未匯入的範圍：雙週FF會用「六/日」推算並列為建議修改、不鎖定匯入，四週WW/W+則完全不檢查。勾選「不檢查」可完全跳過該員的四週WW/W+、雙週FF、NH/N+檢查。</div>
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
            revalidateAndRefresh(title);
        });
    });

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
    dataset.data.forEach(p => { excelMap[p.empId] = { name: p.name, shifts: p.shifts, noCheck: !!p.noCheck }; });
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