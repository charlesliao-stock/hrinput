// background.js - Service Worker
// 1. 首次安裝時初始化 HR 預設班別到 storage（唯一定義來源）
// 2. 持久監聽 content.js 的 modalClosed 訊息並執行分頁跳轉

// HR 系統內建班別預設清單（含上下班時間）
// start / end 為 "HH:MM" 格式；休假類與加班類不設時間（null），接班檢測時跳過
const DEFAULT_HR_SHIFTS = [
  { code: "84",  start: "08:00", end: "16:30" },
  { code: "85",  start: "08:00", end: "17:30" },
  { code: "4N",  start: "16:00", end: "00:30" },  // 跨日班，00:30 表示隔天凌晨0:30
  { code: "5G",  start: "17:30", end: "21:30" },
  { code: "PH",  start: "00:00", end: "08:30" },
  { code: "SS",  start: "08:00", end: "17:30" },  // 放假，但時間比照85
  { code: "VV",  start: "08:00", end: "17:30" },  // 放假，但時間比照85
  { code: "DL",  start: "13:30", end: "22:00" },
  { code: "FF",  start: null,    end: null    },   // 例休，跳過接班檢測
  { code: "WW",  start: null,    end: null    },   // 週休，跳過接班檢測
  { code: "W+",  start: null,    end: null    },   // 休息日加班，跳過接班檢測
  { code: "NH",  start: null,    end: null    },   // 國定假，跳過接班檢測
  { code: "N+",  start: null,    end: null    },   // 假日加班，跳過接班檢測
];

chrome.runtime.onInstalled.addListener(() => {
  // 初始化 autoMode 預設為開啟（僅在尚未設定時寫入）
  chrome.storage.local.get('autoMode', (d) => {
    if (d.autoMode === undefined) {
      chrome.storage.local.set({ autoMode: true });
    }
  });

  chrome.storage.local.get('hrShifts', (data) => {
    if (!data.hrShifts || data.hrShifts.length === 0) {
      chrome.storage.local.set({ hrShifts: DEFAULT_HR_SHIFTS });
    } else {
      // 已有舊資料：若為舊格式（純字串陣列），自動遷移為物件格式
      const needsMigration = data.hrShifts.length > 0 && typeof data.hrShifts[0] === 'string';
      if (needsMigration) {
        const migrated = DEFAULT_HR_SHIFTS.filter(d =>
          data.hrShifts.includes(d.code)
        );
        // 補上舊資料中有但 DEFAULT 沒有的代號（時間設為 null）
        data.hrShifts.forEach(code => {
          if (!migrated.find(d => d.code === code)) {
            migrated.push({ code, start: null, end: null });
          }
        });
        chrome.storage.local.set({ hrShifts: migrated });
      }
    }
  });
});

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {

  if (request.action === "modalClosed") {
    chrome.storage.local.get(['pendingNextUrl', 'autoMode'], (data) => {
      if (data.autoMode && data.pendingNextUrl) {
        const url = data.pendingNextUrl;
        chrome.storage.local.remove('pendingNextUrl');
        if (sender && sender.tab && sender.tab.id) {
          chrome.tabs.update(sender.tab.id, { url });
        } else {
          chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
            if (tabs && tabs[0]) chrome.tabs.update(tabs[0].id, { url });
          });
        }
      } else {
        chrome.storage.local.remove('pendingNextUrl');
      }
    });
    sendResponse({ received: true });
    return true;
  }

  if (request.action === "setPendingUrl") {
    chrome.storage.local.set({ pendingNextUrl: request.url }, () => {
      sendResponse({ ok: true });
    });
    return true;
  }

  if (request.action === "clearPendingUrl") {
    chrome.storage.local.remove('pendingNextUrl');
    sendResponse({ ok: true });
    return true;
  }

});
