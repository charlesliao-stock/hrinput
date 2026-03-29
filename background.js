// background.js - Service Worker
// 1. 首次安裝時初始化 HR 預設班別到 storage（唯一定義來源）
// 2. 持久監聽 content.js 的 modalClosed 訊息並執行分頁跳轉
// 3. 定期檢查 GitHub 更新

const VERSION_CHECK_URL = "https://raw.githubusercontent.com/charlesliao-stock/hrinput/main/version.json";

// HR 系統內建班別預設清單（含上下班時間）
const DEFAULT_HR_SHIFTS = [
  { code: "84",  start: "08:00", end: "16:30" },
  { code: "85",  start: "08:00", end: "17:30" },
  { code: "4N",  start: "16:00", end: "00:30" },
  { code: "5G",  start: "17:30", end: "21:30" },
  { code: "PH",  start: "00:00", end: "08:30" },
  { code: "SS",  start: "08:00", end: "17:30" },
  { code: "VV",  start: "08:00", end: "17:30" },
  { code: "DL",  start: "13:30", end: "22:00" },
  { code: "FF",  start: null,    end: null    },
  { code: "WW",  start: null,    end: null    },
  { code: "W+",  start: null,    end: null    },
  { code: "NH",  start: null,    end: null    },
  { code: "N+",  start: null,    end: null    },
];

// 檢查更新函式
async function checkForUpdates() {
  try {
    const response = await fetch(VERSION_CHECK_URL);
    const data = await response.json();
    const currentVersion = chrome.runtime.getManifest().version;

    if (data.version !== currentVersion) {
      console.log(`[更新偵測] 發現新版本: ${data.version} (目前: ${currentVersion})`);
      chrome.storage.local.set({ 
        updateAvailable: true, 
        latestVersion: data.version,
        downloadUrl: data.downloadUrl,
        changelog: data.changelog
      });
      // 在圖示上顯示 "New" 標籤
      chrome.action.setBadgeText({ text: "New" });
      chrome.action.setBadgeBackgroundColor({ color: "#e74c3c" });
    } else {
      chrome.storage.local.set({ updateAvailable: false });
      chrome.action.setBadgeText({ text: "" });
    }
  } catch (error) {
    console.error("[更新偵測] 檢查失敗:", error);
  }
}

chrome.runtime.onInstalled.addListener(() => {
  // 初始化 autoMode
  chrome.storage.local.get('autoMode', (d) => {
    if (d.autoMode === undefined) {
      chrome.storage.local.set({ autoMode: true });
    }
  });

  // 初始化 HR 班別
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

  // 設定定時檢查更新 (每 6 小時檢查一次)
  chrome.alarms.create("checkUpdate", { periodInMinutes: 360 });
  checkForUpdates();
});

// 監聽定時任務
chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === "checkUpdate") {
    checkForUpdates();
  }
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

  if (request.action === "manualCheckUpdate") {
    checkForUpdates().then(() => sendResponse({ ok: true }));
    return true;
  }
});
