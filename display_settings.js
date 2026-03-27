document.addEventListener('DOMContentLoaded', async () => {
    const data = await chrome.storage.local.get([
        'showWebPreview', 'showExcelReport'
    ]);

    document.getElementById('showWebPreview').checked  = data.showWebPreview  !== false;
    document.getElementById('showExcelReport').checked = data.showExcelReport !== false;

    const updateWindowHeight = () => {
        const height = document.body.offsetHeight + 40;
        chrome.windows.getCurrent((win) => {
            chrome.windows.update(win.id, { height });
        });
    };
    setTimeout(updateWindowHeight, 100);

    document.getElementById('saveBtn').onclick = async () => {
        await chrome.storage.local.set({
            showWebPreview:  document.getElementById('showWebPreview').checked,
            showExcelReport: document.getElementById('showExcelReport').checked
        });
        window.close();
    };
});
