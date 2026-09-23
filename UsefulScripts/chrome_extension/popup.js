document.getElementById("openBtn").addEventListener("click", () => {
  chrome.tabs.create({ url: chrome.runtime.getURL("downloader.html") });
});
