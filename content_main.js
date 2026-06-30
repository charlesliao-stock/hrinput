// content_main.js
// ⚠️ 此檔案以 world: "MAIN" 注入，專門修補 HR 網頁的 IE 相容 API。
// 不可在此使用 chrome.* API（MAIN world 中 chrome 物件不存在）。

(function () {
    if (!window.attachEvent && window.addEventListener) {
        window.attachEvent = function (event, handler) {
            const eventName = event.startsWith('on') ? event.substring(2) : event;
            window.addEventListener(eventName, handler);
        };
        console.log("🔧 [KMUH Helper] 已成功注入 window.attachEvent 相容性修補");
    }
    if (!window.showModalDialog) {
        window.showModalDialog = function (url, arg, feat) {
            console.log("🔧 [KMUH Helper] 攔截到 showModalDialog 呼叫");
            const msg = (typeof arg === 'string') ? arg.replace(/<[^>]+>/g, '') : "網頁發生錯誤，請檢查輸入內容。";
            alert("⚠️ 網頁錯誤訊息：\n\n" + msg);
        };
        console.log("🔧 [KMUH Helper] 已成功注入 window.showModalDialog 相容性修補");
    }
})();
